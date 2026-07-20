"""Create management-fee extracts from an Excel control workbook."""

from __future__ import annotations

import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Iterator, Optional, Tuple

import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.copier import WorksheetCopy
from openpyxl.worksheet.worksheet import Worksheet

CLASS_NUMBERS: range = range(1, 7)
RAW_COLUMNS: list[str] = [
    'Fund Code', 'Fund Name', 'Date', 'Class Number', 'Class Name',
    'NAV', 'Mgmt Fee', 'Source File',
]
MISSING_COLUMNS: list[str] = ['Fund Code', 'Fund Name', 'Date', 'Missing File']
ERROR_COLUMNS: list[str] = ['Fund Code', 'Fund Name', 'Date', 'File', 'Error']
REQUIRED_SETTINGS: frozenset[str] = frozenset({
    'Source Folder',
    'Source Sheet Name',
    'Start Date',
    'End Date',
    'Use Previous Quarter If Dates Blank',
    'Output Folder',
    'Output Workbook Name',
    'Include Buffer Dates In Output',
    'Calc Lookback Days',
    'Calc Lookahead Days',
    'Max Workers',
    'Fund Mapping Sheet',
    'Mapping Output Sheet Name',
    'Class Template Sheet',
    'Template First Data Row',
    'Template Last Data Row',
})
_DATE_FORMATS: Tuple[str, ...] = ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y')
_HEADER_FILL = PatternFill('solid', fgColor='D9EAF7')


@dataclass(frozen=True)
class Fund:
    """A single fund selected for extraction.

    Attributes:
        code: The fund's short code, e.g. "ABC123".
        name: The fund's full name, used in default source file names.
        prefix: An optional custom source-file naming pattern.
    """

    code: str
    name: str
    prefix: str


@dataclass(frozen=True)
class RunConfig:
    """Validated operational settings for one extract run.

    Attributes:
        source_folder: Default folder to look for source .xls files in.
        source_sheet: Name of the worksheet holding class data in each
            source file.
        output_folder: Folder to save the output workbook to.
        output_name: File name (including .xlsx extension) for the output
            workbook.
        start: First date of interest to the business (report window).
        end: Last date of interest to the business (report window).
        read_start: First date to actually read source files for, after
            applying the lookback buffer.
        read_end: Last date to actually read source files for, after
            applying the lookahead buffer.
        show_buffers: Whether buffer dates (outside start/end) are shown
            unhidden in the output.
        workers: Maximum number of threads used to read source files.
        mapping_sheet: Name of the fund mapping sheet in the control
            workbook.
        output_mapping: Name of the sheet to copy the fund mapping into in
            the output workbook.
        template_sheet: Name of the per-class template sheet in the
            control workbook.
        first_row: First data row in the class template.
        last_row: Last data row in the class template.
    """

    source_folder: Path
    source_sheet: str
    output_folder: Path
    output_name: str
    start: date
    end: date
    read_start: date
    read_end: date
    show_buffers: bool
    workers: int
    mapping_sheet: str
    output_mapping: str
    template_sheet: str
    first_row: int
    last_row: int


@dataclass
class ExtractResult:
    """Container for the outputs of one management-fee extract run.

    Attributes:
        raw: Raw NAV/Mgmt Fee rows read from source files.
        missing: Records of expected weekday source files that were not
            found on disk.
        errors: Records of source files that existed but failed to read.
    """

    raw: list[dict[str, Any]] = field(default_factory=list)
    missing: list[dict[str, Any]] = field(default_factory=list)
    errors: list[dict[str, Any]] = field(default_factory=list)


def is_blank(value: Any) -> bool:
    """Determine whether a cell value should be treated as empty.

    Args:
        value: The raw value read from an Excel cell.

    Returns:
        True if the value is None or a string containing only whitespace.
    """
    return value is None or (isinstance(value, str) and not value.strip())


def parse_bool(value: Any, name: str) -> bool:
    """Parse a Y/N-style setting into a boolean.

    Args:
        value: The raw setting value (e.g. "Y", "No", True, 1).
        name: The setting name, used in the error message.

    Returns:
        True for Y/Yes/True/1, False for N/No/False/0.

    Raises:
        ValueError: If the value cannot be interpreted as a boolean.
    """
    text = str(value or '').strip().upper()
    if text in {'Y', 'YES', 'TRUE', '1'}:
        return True
    if text in {'N', 'NO', 'FALSE', '0'}:
        return False
    raise ValueError(f'{name} must be Y or N.')


def parse_int(value: Any, name: str, minimum: int = 0) -> int:
    """Parse and validate an integer setting.

    Args:
        value: The raw setting value.
        name: The setting name, used in error messages.
        minimum: The smallest value that is accepted.

    Returns:
        The parsed integer.

    Raises:
        ValueError: If the value is not an integer or is below `minimum`.
    """
    try:
        number = int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f'{name} must be an integer.') from exc
    if number < minimum:
        raise ValueError(f'{name} must be at least {minimum}.')
    return number


def parse_date(value: Any, name: str) -> Optional[date]:
    """Parse a setting value into a date.

    Args:
        value: The raw setting value. May already be a date/datetime.
        name: The setting name, used in the error message.

    Returns:
        The parsed date, or None if `value` is blank.

    Raises:
        ValueError: If the value is non-blank and matches none of the
            supported date formats.
    """
    if is_blank(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for pattern in _DATE_FORMATS:
        try:
            return datetime.strptime(str(value).strip(), pattern).date()
        except ValueError:
            continue
    raise ValueError(f'{name} must be DD/MM/YYYY.')


def previous_quarter() -> Tuple[date, date]:
    """Compute the start and end dates of the previous calendar quarter.

    Returns:
        A (start_date, end_date) tuple for the quarter before the one
        containing today's date.
    """
    today = date.today()
    quarter = (today.month - 1) // 3 + 1
    year, prior_quarter = (
        (today.year - 1, 4) if quarter == 1 else (today.year, quarter - 1)
    )
    month = (prior_quarter - 1) * 3 + 1
    start = date(year, month, 1)
    next_start = (
        date(year + 1, 1, 1) if month == 10 else date(year, month + 3, 1)
    )
    return start, next_start - timedelta(days=1)


def date_range(start: date, end: date) -> Iterator[date]:
    """Yield each date from `start` to `end`, inclusive.

    Args:
        start: The first date to yield.
        end: The last date to yield.

    Yields:
        Each date in the inclusive range, in order.
    """
    for offset in range((end - start).days + 1):
        yield start + timedelta(days=offset)


def select_file(config: Optional[RunConfig] = None) -> Path:
    """Prompt the user to choose the control workbook or output location.

    Uses a Tkinter file dialog when a display is available, and falls back
    to a console prompt (or the configured output path) when Tkinter
    cannot be imported.

    Args:
        config: When provided, selects a save location for the output
            workbook. When None, selects the control workbook to open.

    Returns:
        The path chosen by the user.

    Raises:
        KeyboardInterrupt: If the user cancels the file dialog.
    """
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        if config:
            return config.output_folder / config.output_name
        return Path(input('Control workbook path: ').strip().strip('"'))

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    try:
        if config:
            selected = filedialog.asksaveasfilename(
                title='Confirm output workbook save location',
                initialdir=str(config.output_folder),
                initialfile=config.output_name,
                defaultextension='.xlsx',
                filetypes=[('Excel workbook', '*.xlsx')],
            )
        else:
            selected = filedialog.askopenfilename(
                title='Select management fee control workbook',
                filetypes=[('Excel workbooks', '*.xlsx *.xlsm')],
            )
    finally:
        root.destroy()
    if not selected:
        raise KeyboardInterrupt
    return Path(selected)


def _read_settings_sheet(control_file: Path) -> dict[str, Any]:
    """Read the Settings sheet of the control workbook into a dict.

    Args:
        control_file: Path to the control workbook.

    Returns:
        A mapping of setting name to its raw value.

    Raises:
        ValueError: If any setting in `REQUIRED_SETTINGS` is missing.
    """
    workbook = load_workbook(control_file, read_only=True, data_only=True)
    try:
        settings = {
            str(row[0]).strip(): row[1]
            for row in workbook['Settings'].iter_rows(
                min_row=2, values_only=True
            )
            if row and not is_blank(row[0])
        }
    finally:
        workbook.close()
    missing = sorted(REQUIRED_SETTINGS - settings.keys())
    if missing:
        raise ValueError('Missing Settings row(s): ' + ', '.join(missing))
    return settings


def _resolve_date_range(settings: dict[str, Any]) -> Tuple[date, date]:
    """Determine the run's start and end dates from raw settings.

    Falls back to the previous calendar quarter when both dates are blank
    and that fallback is enabled.

    Args:
        settings: The raw settings mapping from the Settings sheet.

    Returns:
        A (start_date, end_date) tuple.

    Raises:
        ValueError: If only one of Start/End Date is given, or if End Date
            is earlier than Start Date.
    """
    start = parse_date(settings['Start Date'], 'Start Date')
    end = parse_date(settings['End Date'], 'End Date')
    use_previous_quarter = parse_bool(
        settings['Use Previous Quarter If Dates Blank'],
        'Use Previous Quarter If Dates Blank',
    )
    if start is None and end is None and use_previous_quarter:
        start, end = previous_quarter()
    elif start is None or end is None:
        raise ValueError('Provide both Start Date and End Date.')
    if end < start:
        raise ValueError('End Date cannot be earlier than Start Date.')
    return start, end


def build_config(control_file: Path) -> RunConfig:
    """Read and validate all operational settings from Excel.

    Args:
        control_file: Path to the control workbook.

    Returns:
        The validated run configuration.

    Raises:
        ValueError: If a required setting is missing or fails validation.
    """
    settings = _read_settings_sheet(control_file)
    start, end = _resolve_date_range(settings)
    lookback = parse_int(settings['Calc Lookback Days'], 'Calc Lookback Days')
    lookahead = parse_int(
        settings['Calc Lookahead Days'], 'Calc Lookahead Days'
    )
    first_row = parse_int(
        settings['Template First Data Row'], 'Template First Data Row', 1
    )
    last_row = parse_int(
        settings['Template Last Data Row'],
        'Template Last Data Row',
        first_row,
    )
    output_name = str(settings['Output Workbook Name']).strip()
    if not output_name.lower().endswith('.xlsx'):
        output_name += '.xlsx'
    return RunConfig(
        source_folder=Path(str(settings['Source Folder']).strip()),
        source_sheet=str(settings['Source Sheet Name']).strip(),
        output_folder=Path(str(settings['Output Folder']).strip()),
        output_name=output_name,
        start=start,
        end=end,
        read_start=start - timedelta(days=lookback),
        read_end=end + timedelta(days=lookahead),
        show_buffers=parse_bool(
            settings['Include Buffer Dates In Output'],
            'Include Buffer Dates In Output',
        ),
        workers=parse_int(settings['Max Workers'], 'Max Workers', 1),
        mapping_sheet=str(settings['Fund Mapping Sheet']).strip(),
        output_mapping=str(settings['Mapping Output Sheet Name']).strip(),
        template_sheet=str(settings['Class Template Sheet']).strip(),
        first_row=first_row,
        last_row=last_row,
    )


def load_funds(control_file: Path, sheet_name: str) -> list[Fund]:
    """Read funds marked "Y" for inclusion from the fund mapping sheet.

    Args:
        control_file: Path to the control workbook.
        sheet_name: Name of the sheet listing funds in columns A:D
            (Fund Code, Fund Name, Include Y/N, Source Prefix).

    Returns:
        The funds marked for inclusion, in sheet order.

    Raises:
        ValueError: If an included row is missing its code or name, or if
            no funds are marked for inclusion.
    """
    workbook = load_workbook(control_file, read_only=True, data_only=False)
    try:
        funds: list[Fund] = []
        rows = workbook[sheet_name].iter_rows(min_row=2, values_only=True)
        for row_number, row in enumerate(rows, start=2):
            values = list(row[:4]) + [None] * (4 - len(row[:4]))
            code, name, include, prefix = values
            if str(include or '').strip().upper() != 'Y':
                continue
            if is_blank(code) or is_blank(name):
                raise ValueError(f'Mapping row {row_number} is incomplete.')
            funds.append(
                Fund(
                    code=str(code).strip(),
                    name=str(name).strip(),
                    prefix=str(prefix or '').strip(),
                )
            )
    finally:
        workbook.close()
    if not funds:
        raise ValueError('No funds are marked Y in Fund_Mapping.')
    return funds


def build_source_path(config: RunConfig, fund: Fund, day: date) -> Path:
    """Determine the expected source file path for a fund and date.

    Applies the fund's custom prefix pattern when one is set (supporting
    `{date}`/`YYYYMMDD` placeholders and literal `.xls` filenames), and
    otherwise falls back to the standard naming convention.

    Args:
        config: The run configuration (used for the default source folder).
        fund: The fund whose source file path is being resolved.
        day: The date the source file should correspond to.

    Returns:
        The expected path to the fund's source file for `day`.
    """
    stamp = day.strftime('%Y%m%d')
    prefix = fund.prefix.strip().strip('"')
    if prefix and not prefix.startswith('='):
        if '{date}' in prefix:
            return Path(prefix.replace('{date}', stamp))
        if 'YYYYMMDD' in prefix:
            return Path(prefix.replace('YYYYMMDD', stamp))
        if prefix.lower().endswith('.xls'):
            return Path(prefix)
        suffix = (
            f'{stamp}.xls'
            if prefix.endswith('_Unit_Price_')
            else f'_Unit_Price_{stamp}.xls'
        )
        return Path(prefix + suffix)
    return config.source_folder / (
        f'PINNACLE03_{fund.name}_Unit_Price_{stamp}.xls'
    )


def read_source(
    fund: Fund, day: date, file_path: Path, sheet_name: str
) -> list[dict[str, Any]]:
    """Read up to six unit-price classes from one binary .xls source file.

    Args:
        fund: The fund the source file belongs to.
        day: The date the source file corresponds to.
        file_path: Path to the source .xls file.
        sheet_name: Name of the worksheet holding the class data.

    Returns:
        One record per populated class (matching `RAW_COLUMNS`), for
        classes 1-6. Class 1 is always included; classes 2-6 are skipped
        once their name cell is blank.
    """
    workbook = xlrd.open_workbook(str(file_path), on_demand=True)
    try:
        worksheet = workbook.sheet_by_name(sheet_name)

        def read_cell(row: int, column: int) -> Any:
            """Read a 1-indexed cell, or None if blank or out of range."""
            row_index, column_index = row - 1, column - 1
            out_of_range = (
                row_index >= worksheet.nrows
                or column_index >= worksheet.ncols
            )
            if out_of_range:
                return None
            value = worksheet.cell_value(row_index, column_index)
            return None if value == '' else value

        rows: list[dict[str, Any]] = []
        for class_number in CLASS_NUMBERS:
            source_row = class_number + 1
            if class_number != 1 and is_blank(read_cell(source_row, 1)):
                continue
            rows.append(
                {
                    'Fund Code': fund.code,
                    'Fund Name': fund.name,
                    'Date': day,
                    'Class Number': class_number,
                    'Class Name': read_cell(source_row, 2),
                    'NAV': read_cell(source_row, 8),
                    'Mgmt Fee': read_cell(source_row, 9),
                    'Source File': file_path.name,
                }
            )
        return rows
    finally:
        workbook.release_resources()


_Task = Tuple[Fund, date, Path]


def _build_tasks(config: RunConfig, funds: list[Fund]) -> list[_Task]:
    """Build the (fund, date, file_path) work items for one run.

    Args:
        config: The run configuration (defines the date window to read).
        funds: The funds to process.

    Returns:
        One task per fund/date combination in the read window.
    """
    return [
        (fund, day, build_source_path(config, fund, day))
        for day in date_range(config.read_start, config.read_end)
        for fund in funds
    ]


def _process_task(task: _Task, sheet_name: str) -> ExtractResult:
    """Read one source file, or record why it could not be read.

    Args:
        task: A (fund, date, file_path) work item.
        sheet_name: Name of the worksheet holding the class data.

    Returns:
        An `ExtractResult` holding at most one of: raw rows, a
        missing-file record (only recorded for weekdays), or an error
        record.
    """
    fund, day, file_path = task
    result = ExtractResult()
    if not file_path.exists():
        if day.weekday() < 5:  # Monday-Friday; weekend gaps are expected.
            result.missing.append(
                {
                    'Fund Code': fund.code,
                    'Fund Name': fund.name,
                    'Date': day,
                    'Missing File': str(file_path),
                }
            )
        return result
    try:
        result.raw = read_source(fund, day, file_path, sheet_name)
    except Exception as exc:  # noqa: BLE001 - any read failure is logged.
        result.errors.append(
            {
                'Fund Code': fund.code,
                'Fund Name': fund.name,
                'Date': day,
                'File': str(file_path),
                'Error': str(exc),
            }
        )
    return result


def process_files(config: RunConfig, funds: list[Fund]) -> ExtractResult:
    """Read all source files concurrently and collect audit results.

    Args:
        config: The run configuration.
        funds: The funds to process.

    Returns:
        The combined raw rows, missing-file records, and error records
        across all funds and dates in the read window.
    """
    tasks = _build_tasks(config, funds)
    combined = ExtractResult()
    with ThreadPoolExecutor(
        max_workers=min(config.workers, len(tasks))
    ) as pool:
        futures = [
            pool.submit(_process_task, task, config.source_sheet)
            for task in tasks
        ]
        for future in as_completed(futures):
            result = future.result()
            combined.raw.extend(result.raw)
            combined.missing.extend(result.missing)
            combined.errors.extend(result.errors)
    return combined


def write_log(
    workbook: Workbook,
    name: str,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    """Write a list of records to a new worksheet as a formatted table.

    Replaces any existing sheet with the same name.

    Args:
        workbook: The workbook to add the sheet to.
        name: Name of the sheet to create.
        columns: Column headings, in order; also used as record keys.
        rows: The records to write, one per output row.
    """
    if name in workbook.sheetnames:
        del workbook[name]
    worksheet = workbook.create_sheet(name)
    for column, heading in enumerate(columns, start=1):
        cell = worksheet.cell(row=1, column=column, value=heading)
        cell.font = Font(bold=True)
        cell.fill = _HEADER_FILL
        cell.alignment = Alignment(horizontal='center')
    for row_number, record in enumerate(rows, start=2):
        for column, heading in enumerate(columns, start=1):
            cell = worksheet.cell(
                row=row_number, column=column, value=record.get(heading)
            )
            if heading == 'Date':
                cell.number_format = 'dd/mm/yyyy'
    worksheet.freeze_panes = 'A2'


def copy_mapping(
    workbook: Workbook, source_name: str, output_name: str
) -> Worksheet:
    """Copy the fund mapping sheet to a visible output sheet.

    Reuses an existing sheet named `output_name` (case-insensitively) if
    one is already present, instead of creating a duplicate.

    Args:
        workbook: The workbook containing `source_name`.
        source_name: Name of the sheet to copy from.
        output_name: Name of the sheet to copy into.

    Returns:
        The populated output worksheet.
    """
    source = workbook[source_name]
    target = next(
        (
            sheet
            for sheet in workbook.worksheets
            if sheet.title.lower() == output_name.lower()
        ),
        None,
    )
    if target is None:
        target = workbook.create_sheet(output_name)
    WorksheetCopy(source, target).copy_worksheet()
    target.sheet_state = 'visible'
    target.freeze_panes = 'A2'
    return target


def _validate_template_capacity(config: RunConfig, required_rows: int) -> None:
    """Ensure the class template has enough rows for the read window.

    Args:
        config: The run configuration.
        required_rows: Number of dates that must fit in the template.

    Raises:
        ValueError: If the template does not have enough rows.
    """
    capacity = config.last_row - config.first_row + 1
    if required_rows > capacity:
        raise ValueError(
            f'Template supports {capacity} dates; run requires '
            f'{required_rows}.'
        )


def _group_records_by_fund_class(
    records: list[dict[str, Any]],
) -> Tuple[
    dict[Tuple[str, int], dict[date, dict[str, Any]]],
    dict[Tuple[str, int], str],
]:
    """Group raw records by fund/class and collect each class's name.

    Args:
        records: Raw records as produced by `read_source`.

    Returns:
        A tuple of:
          - Records keyed by (fund_code, class_number), each mapping date
            to its record.
          - The first non-blank class name seen for each
            (fund_code, class_number).
    """
    grouped: dict[Tuple[str, int], dict[date, dict[str, Any]]] = {}
    class_names: dict[Tuple[str, int], str] = {}
    for record in records:
        key = (record['Fund Code'], record['Class Number'])
        grouped.setdefault(key, {})[record['Date']] = record
        if not is_blank(record['Class Name']):
            class_names.setdefault(key, record['Class Name'])
    return grouped, class_names


def _has_reportable_data(records: dict[date, dict[str, Any]]) -> bool:
    """Check whether any record in a class has a NAV or management fee.

    Args:
        records: Date-keyed records for a single fund/class.

    Returns:
        True if at least one record has a non-blank NAV or Mgmt Fee.
    """
    return any(
        not is_blank(row.get('NAV')) or not is_blank(row.get('Mgmt Fee'))
        for row in records.values()
    )


def _populate_class_worksheet(
    worksheet: Worksheet,
    config: RunConfig,
    fund: Fund,
    class_number: int,
    class_name: Optional[str],
    records: dict[date, dict[str, Any]],
    required_rows: int,
) -> None:
    """Fill in one class worksheet's header cells and date rows.

    Args:
        worksheet: The freshly copied class worksheet to populate.
        config: The run configuration.
        fund: The fund the worksheet belongs to.
        class_number: The share class number (1-6).
        class_name: The class's display name, if known.
        records: Date-keyed source records for this fund/class.
        required_rows: Number of date rows the run needs.
    """
    worksheet.title = f'{fund.code}_Class_{class_number}'[:31]
    worksheet.sheet_state = 'visible'
    worksheet['A1'] = f'Class {class_number}'
    worksheet['I2'] = fund.code
    worksheet['I3'] = class_name

    # Clear prior inputs only; template formulas remain in columns D:F.
    for row in range(config.first_row, config.last_row + 1):
        for column in range(1, 4):
            worksheet.cell(row=row, column=column).value = None

    for offset, day in enumerate(
        date_range(config.read_start, config.read_end)
    ):
        row = config.first_row + offset
        record = records.get(day, {})
        worksheet.cell(row=row, column=1, value=day)
        worksheet.cell(row=row, column=2, value=record.get('NAV'))
        worksheet.cell(row=row, column=3, value=record.get('Mgmt Fee'))
        worksheet.row_dimensions[row].hidden = (
            not config.show_buffers and not config.start <= day <= config.end
        )

    # Blank out any unused template rows beyond the dates actually written.
    for row in range(config.first_row + required_rows, config.last_row + 1):
        for column in range(1, 7):
            worksheet.cell(row=row, column=column).value = None


def _build_class_worksheets(
    workbook: Workbook,
    template: Worksheet,
    config: RunConfig,
    funds: list[Fund],
    records: list[dict[str, Any]],
    required_rows: int,
) -> None:
    """Create and populate one worksheet per fund/class with reportable data.

    Funds/classes with no non-blank NAV or Mgmt Fee in the read window are
    skipped entirely.

    Args:
        workbook: The output workbook being assembled.
        template: The class template worksheet to copy.
        config: The run configuration.
        funds: The funds to build worksheets for.
        records: Raw records as produced by `read_source`.
        required_rows: Number of date rows the run needs.
    """
    grouped, class_names = _group_records_by_fund_class(records)
    for fund in funds:
        for class_number in CLASS_NUMBERS:
            key = (fund.code, class_number)
            class_records = grouped.get(key, {})
            if not _has_reportable_data(class_records):
                continue
            worksheet = workbook.copy_worksheet(template)
            _populate_class_worksheet(
                worksheet,
                config,
                fund,
                class_number,
                class_names.get(key),
                class_records,
                required_rows,
            )


def _remove_working_sheets(
    workbook: Workbook, config: RunConfig, keep: Worksheet
) -> None:
    """Remove the settings, mapping, and template sheets from the output.

    Args:
        workbook: The output workbook.
        config: The run configuration (names the sheets to remove).
        keep: A worksheet to never delete, even if its name matches.
    """
    for name in ('Settings', config.mapping_sheet, config.template_sheet):
        if name in workbook.sheetnames and workbook[name] is not keep:
            del workbook[name]


def _sort_results(result: ExtractResult) -> None:
    """Sort raw, missing, and error records into a stable, readable order.

    Args:
        result: The extract result to sort in place.
    """
    result.raw.sort(
        key=lambda row: (row['Date'], row['Fund Code'], row['Class Number'])
    )
    result.missing.sort(key=lambda row: (row['Date'], row['Fund Code']))
    result.errors.sort(key=lambda row: (row['Date'], row['Fund Code']))


def create_output(
    control_file: Path,
    output_file: Path,
    config: RunConfig,
    funds: list[Fund],
    result: ExtractResult,
) -> None:
    """Copy the class template per fund/class, write logs, and save.

    Args:
        control_file: Path to the control workbook (used as the template
            source).
        output_file: Path to write the populated output workbook to.
        config: The run configuration.
        funds: The funds included in this run.
        result: The raw rows, missing-file records, and error records
            collected by `process_files`.

    Raises:
        ValueError: If the class template does not have enough rows for
            the configured date range.
    """
    workbook = load_workbook(control_file, data_only=False)
    try:
        template = workbook[config.template_sheet]
        required_rows = (config.read_end - config.read_start).days + 1
        _validate_template_capacity(config, required_rows)

        _build_class_worksheets(
            workbook, template, config, funds, result.raw, required_rows
        )
        mapping = copy_mapping(
            workbook, config.mapping_sheet, config.output_mapping
        )
        _remove_working_sheets(workbook, config, keep=mapping)

        _sort_results(result)
        write_log(workbook, 'Raw_Data', RAW_COLUMNS, result.raw)
        write_log(workbook, 'Missing_Files', MISSING_COLUMNS, result.missing)
        write_log(workbook, 'Errors', ERROR_COLUMNS, result.errors)

        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        output_file.parent.mkdir(parents=True, exist_ok=True)
        workbook.save(output_file)
    finally:
        workbook.close()


def _print_summary(
    output_file: Path, funds: list[Fund], result: ExtractResult
) -> None:
    """Print a run summary to stdout.

    Args:
        output_file: Path the output workbook was saved to.
        funds: The funds included in the run.
        result: The collected raw rows, missing-file records, and errors.
    """
    print(f'Created: {output_file}')
    print(f'Funds: {len(funds)} | Raw rows: {len(result.raw)}')
    print(
        f'Missing files: {len(result.missing)} | Errors: {len(result.errors)}'
    )


def main() -> int:
    """Run the management-fee extract end to end.

    Prompts for the control workbook and output location, reads and
    validates settings, reads source files concurrently, and writes the
    populated output workbook.

    Returns:
        A process exit code: 0 on success, 130 if cancelled by the user,
        or 1 on any other error.
    """
    try:
        control_file = select_file()
        config = build_config(control_file)
        funds = load_funds(control_file, config.mapping_sheet)
        output_file = select_file(config)
        if output_file.resolve() == control_file.resolve():
            raise ValueError('Output cannot overwrite the control workbook.')
        result = process_files(config, funds)
        create_output(control_file, output_file, config, funds, result)
        _print_summary(output_file, funds, result)
        return 0
    except KeyboardInterrupt:
        print('Cancelled.', file=sys.stderr)
        return 130
    except Exception as exc:  # noqa: BLE001 - top-level error boundary.
        print(f'Fatal error: {exc}', file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
