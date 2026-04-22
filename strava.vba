' Strava API Integration for Excel VBA
'
' SETUP INSTRUCTIONS:
' 1. Go to https://www.strava.com/settings/api and create an application
' 2. Set "Authorization Callback Domain" to: localhost
' 3. Note your Client ID and Client Secret
' 4. Run the macro StravaSetup to begin

Option Explicit

' Public URL constants
  Public Const URL_Strava_CreateApp  As String = "https://www.strava.com/settings/api"
  Public Const URL_Strava_Dashboard  As String = "https://www.strava.com/dashboard"

' Private API constants
  Private Const STRAVA_AUTH_URL      As String = "https://www.strava.com/oauth/authorize"
  Private Const STRAVA_TOKEN_URL     As String = "https://www.strava.com/oauth/token"
  Private Const STRAVA_API_BASE      As String = "https://www.strava.com/api/v3"
  Private Const STRAVA_REDIRECT_URI  As String = "http://localhost"
  Private Const STRAVA_SCOPE         As String = "read,activity:read_all,profile:read_all"
  Private Const HTTPRequest_Client   As String = "WinHttp.WinHttpRequest.5.1"

' Registry paths
  Private Const REG_APP   As String = "StravaAPI"
  Private Const REG_CREDS As String = "Credentials"


' =============================================================================
' PUBLIC MACROS — run these from Excel
' =============================================================================

Sub StravaSetup()
' Guides the user through first-time setup: save Client ID/Secret, then
' open the browser for OAuth authorisation and exchange the code for tokens.

      Dim ClientID As String
      Dim ClientSecret As String

      ' Prompt for Client ID
      ClientID = InputBox( _
            "Step 1 of 2 — Enter your Strava Client ID." & vbNewLine & vbNewLine & _
            "You can find it at: " & URL_Strava_CreateApp, _
            " Strava Setup", _
            StravaGetSetting("ClientID"))
      If ClientID = "" Then Exit Sub
      SaveSetting REG_APP, REG_CREDS, "ClientID", ClientID

      ' Prompt for Client Secret
      ClientSecret = InputBox( _
            "Step 2 of 2 — Enter your Strava Client Secret." & vbNewLine & vbNewLine & _
            "(Found on the same page as your Client ID.)", _
            " Strava Setup", _
            StravaGetSetting("ClientSecret"))
      If ClientSecret = "" Then Exit Sub
      SaveSetting REG_APP, REG_CREDS, "ClientSecret", ClientSecret

      ' Kick off the OAuth flow
      Call StravaAuthenticate

End Sub

Sub StravaAuthenticate()
' Opens the Strava authorisation page in the user's browser, then asks the
' user to paste back the redirect URL (or just the code= value) so that
' tokens can be exchanged and saved.

      Dim ClientID As String
      ClientID = StravaGetSetting("ClientID")
      If ClientID = "" Then
            MsgBox "Client ID not set. Please run StravaSetup first.", vbCritical, " Strava"
            Exit Sub
      End If

      ' Build the authorisation URL and open it in the browser
      Dim AuthURL As String
      AuthURL = STRAVA_AUTH_URL & _
                "?client_id=" & ClientID & _
                "&redirect_uri=" & STRAVA_REDIRECT_URI & _
                "&response_type=code" & _
                "&approval_prompt=force" & _
                "&scope=" & STRAVA_SCOPE
      ActiveWorkbook.FollowHyperlink Address:=AuthURL, NewWindow:=True

      ' Ask the user to paste back whatever appeared in the browser address bar
      Dim RedirectResponse As String
      RedirectResponse = InputBox( _
            "After clicking 'Authorize' in the browser, your browser will show an" & vbNewLine & _
            "error page with a URL like:" & vbNewLine & vbNewLine & _
            "  http://localhost?state=&code=XXXXXXXX&scope=..." & vbNewLine & vbNewLine & _
            "Paste the full URL (or just the code value) here:", _
            " Strava — Paste Redirect URL")
      If RedirectResponse = "" Then Exit Sub

      ' Extract the code from the URL (or accept a bare code)
      Dim AuthCode As String
      AuthCode = StravaExtractParam(RedirectResponse, "code")
      If AuthCode = "" Then AuthCode = RedirectResponse   ' User pasted the raw code

      If AuthCode = "" Then
            MsgBox "Could not find an authorisation code. Please try again.", vbCritical, " Strava"
            Exit Sub
      End If

      ' Exchange the code for tokens
      If StravaExchangeCode(AuthCode) Then
            MsgBox "Successfully authenticated with Strava!" & vbNewLine & vbNewLine & _
                   "You can now run StravaGetAthlete or StravaGetActivities.", _
                   vbInformation, " Strava"
      Else
            MsgBox "Token exchange failed. Check your Client ID and Secret and try again.", _
                   vbCritical, " Strava"
      End If

End Sub

Sub StravaGetAthlete()
' Fetches the authenticated athlete profile and writes it to a worksheet
' called "Strava_Athlete" (created if it does not exist).

      Dim Token As String
      Token = StravaAccessToken()
      If Token = "" Then Exit Sub

      Dim JSON As String
      JSON = StravaApiGet("/athlete", Token)
      If JSON = "" Then Exit Sub

      Dim WS As Worksheet
      Set WS = StravaGetOrCreateSheet("Strava_Athlete")
      WS.Cells.ClearContents

      ' Write headers and values for the most useful fields
      Dim Fields(0 To 11, 0 To 1) As String
      Fields(0, 0) = "Field":          Fields(0, 1) = "Value"
      Fields(1, 0) = "ID":             Fields(1, 1) = StravaJsonValue(JSON, "id")
      Fields(2, 0) = "First Name":     Fields(2, 1) = StravaJsonValue(JSON, "firstname")
      Fields(3, 0) = "Last Name":      Fields(3, 1) = StravaJsonValue(JSON, "lastname")
      Fields(4, 0) = "Username":       Fields(4, 1) = StravaJsonValue(JSON, "username")
      Fields(5, 0) = "City":           Fields(5, 1) = StravaJsonValue(JSON, "city")
      Fields(6, 0) = "Country":        Fields(6, 1) = StravaJsonValue(JSON, "country")
      Fields(7, 0) = "Sex":            Fields(7, 1) = StravaJsonValue(JSON, "sex")
      Fields(8, 0) = "Premium":        Fields(8, 1) = StravaJsonValue(JSON, "premium")
      Fields(9, 0) = "Follower Count": Fields(9, 1) = StravaJsonValue(JSON, "follower_count")
      Fields(10, 0) = "Friend Count":  Fields(10, 1) = StravaJsonValue(JSON, "friend_count")
      Fields(11, 0) = "Profile URL":   Fields(11, 1) = StravaJsonValue(JSON, "profile")

      Dim i As Long
      For i = 0 To 11
            WS.Cells(i + 1, 1).Value = Fields(i, 0)
            WS.Cells(i + 1, 2).Value = Fields(i, 1)
      Next i

      WS.Columns("A:B").AutoFit
      WS.Activate
      MsgBox "Athlete profile written to sheet '" & WS.Name & "'.", vbInformation, " Strava"

End Sub

Sub StravaGetActivities(Optional ByVal PageCount As Long = 1, Optional ByVal PerPage As Long = 30)
' Fetches recent activities and writes them to a worksheet called
' "Strava_Activities" (created if it does not exist).
'
' PageCount  — number of pages to fetch (default: 1)
' PerPage    — activities per page, max 200 (default: 30)

      If PerPage > 200 Then PerPage = 200

      Dim Token As String
      Token = StravaAccessToken()
      If Token = "" Then Exit Sub

      Dim WS As Worksheet
      Set WS = StravaGetOrCreateSheet("Strava_Activities")
      WS.Cells.ClearContents

      ' Write column headers
      Dim Headers As Variant
      Headers = Array("ID", "Name", "Type", "Distance (m)", "Moving Time (s)", _
                      "Elapsed Time (s)", "Total Elevation Gain (m)", "Start Date", _
                      "Start Date Local", "Timezone", "Average Speed (m/s)", _
                      "Max Speed (m/s)", "Average Heartrate", "Max Heartrate", _
                      "Suffer Score", "Kudos Count", "Comment Count", "Achievement Count")
      Dim Col As Long
      For Col = 0 To UBound(Headers)
            WS.Cells(1, Col + 1).Value = Headers(Col)
      Next Col

      Dim Row As Long
      Row = 2

      Dim Page As Long
      For Page = 1 To PageCount
            Dim JSON As String
            JSON = StravaApiGet("/athlete/activities?per_page=" & PerPage & "&page=" & Page, Token)
            If JSON = "" Then Exit For

            ' Each activity is a JSON object inside a top-level array.
            ' Split on "},{"  to isolate individual activity objects.
            JSON = Mid$(JSON, 2, Len(JSON) - 2)   ' Strip outer [ ]
            Dim Activities() As String
            Activities = Split(JSON, "},{")

            Dim i As Long
            For i = 0 To UBound(Activities)
                  Dim Act As String
                  Act = Activities(i)
                  If Left$(Act, 1) <> "{" Then Act = "{" & Act
                  If Right$(Act, 1) <> "}" Then Act = Act & "}"

                  WS.Cells(Row, 1).Value  = StravaJsonValue(Act, "id")
                  WS.Cells(Row, 2).Value  = StravaJsonValue(Act, "name")
                  WS.Cells(Row, 3).Value  = StravaJsonValue(Act, "type")
                  WS.Cells(Row, 4).Value  = CDblSafe(StravaJsonValue(Act, "distance"))
                  WS.Cells(Row, 5).Value  = CLngSafe(StravaJsonValue(Act, "moving_time"))
                  WS.Cells(Row, 6).Value  = CLngSafe(StravaJsonValue(Act, "elapsed_time"))
                  WS.Cells(Row, 7).Value  = CDblSafe(StravaJsonValue(Act, "total_elevation_gain"))
                  WS.Cells(Row, 8).Value  = StravaJsonValue(Act, "start_date")
                  WS.Cells(Row, 9).Value  = StravaJsonValue(Act, "start_date_local")
                  WS.Cells(Row, 10).Value = StravaJsonValue(Act, "timezone")
                  WS.Cells(Row, 11).Value = CDblSafe(StravaJsonValue(Act, "average_speed"))
                  WS.Cells(Row, 12).Value = CDblSafe(StravaJsonValue(Act, "max_speed"))
                  WS.Cells(Row, 13).Value = CDblSafe(StravaJsonValue(Act, "average_heartrate"))
                  WS.Cells(Row, 14).Value = CDblSafe(StravaJsonValue(Act, "max_heartrate"))
                  WS.Cells(Row, 15).Value = CDblSafe(StravaJsonValue(Act, "suffer_score"))
                  WS.Cells(Row, 16).Value = CLngSafe(StravaJsonValue(Act, "kudos_count"))
                  WS.Cells(Row, 17).Value = CLngSafe(StravaJsonValue(Act, "comment_count"))
                  WS.Cells(Row, 18).Value = CLngSafe(StravaJsonValue(Act, "achievement_count"))
                  Row = Row + 1
            Next i

            ' Stop if we got fewer activities than requested (last page)
            If UBound(Activities) + 1 < PerPage Then Exit For
      Next Page

      WS.Rows(1).Font.Bold = True
      WS.Columns("A:R").AutoFit
      WS.Activate
      MsgBox (Row - 2) & " activities written to sheet '" & WS.Name & "'.", vbInformation, " Strava"

End Sub

Sub StravaGetActivityDetails()
' Prompts the user for a single activity ID and writes the full detail
' to a worksheet called "Strava_ActivityDetail".

      Dim Token As String
      Token = StravaAccessToken()
      If Token = "" Then Exit Sub

      Dim ActivityID As String
      ActivityID = InputBox("Enter the Strava Activity ID:", " Strava — Activity Detail")
      If ActivityID = "" Then Exit Sub

      Dim JSON As String
      JSON = StravaApiGet("/activities/" & ActivityID, Token)
      If JSON = "" Then Exit Sub

      Dim WS As Worksheet
      Set WS = StravaGetOrCreateSheet("Strava_ActivityDetail")
      WS.Cells.ClearContents

      Dim Fields(0 To 24, 0 To 1) As String
      Fields(0, 0)  = "Field":                    Fields(0, 1)  = "Value"
      Fields(1, 0)  = "ID":                       Fields(1, 1)  = StravaJsonValue(JSON, "id")
      Fields(2, 0)  = "Name":                     Fields(2, 1)  = StravaJsonValue(JSON, "name")
      Fields(3, 0)  = "Description":              Fields(3, 1)  = StravaJsonValue(JSON, "description")
      Fields(4, 0)  = "Type":                     Fields(4, 1)  = StravaJsonValue(JSON, "type")
      Fields(5, 0)  = "Distance (m)":             Fields(5, 1)  = StravaJsonValue(JSON, "distance")
      Fields(6, 0)  = "Moving Time (s)":          Fields(6, 1)  = StravaJsonValue(JSON, "moving_time")
      Fields(7, 0)  = "Elapsed Time (s)":         Fields(7, 1)  = StravaJsonValue(JSON, "elapsed_time")
      Fields(8, 0)  = "Elevation Gain (m)":       Fields(8, 1)  = StravaJsonValue(JSON, "total_elevation_gain")
      Fields(9, 0)  = "Elevation High (m)":       Fields(9, 1)  = StravaJsonValue(JSON, "elev_high")
      Fields(10, 0) = "Elevation Low (m)":        Fields(10, 1) = StravaJsonValue(JSON, "elev_low")
      Fields(11, 0) = "Start Date":               Fields(11, 1) = StravaJsonValue(JSON, "start_date")
      Fields(12, 0) = "Start Date Local":         Fields(12, 1) = StravaJsonValue(JSON, "start_date_local")
      Fields(13, 0) = "Timezone":                 Fields(13, 1) = StravaJsonValue(JSON, "timezone")
      Fields(14, 0) = "Average Speed (m/s)":      Fields(14, 1) = StravaJsonValue(JSON, "average_speed")
      Fields(15, 0) = "Max Speed (m/s)":          Fields(15, 1) = StravaJsonValue(JSON, "max_speed")
      Fields(16, 0) = "Average Cadence":          Fields(16, 1) = StravaJsonValue(JSON, "average_cadence")
      Fields(17, 0) = "Average Watts":            Fields(17, 1) = StravaJsonValue(JSON, "average_watts")
      Fields(18, 0) = "Max Watts":                Fields(18, 1) = StravaJsonValue(JSON, "max_watts")
      Fields(19, 0) = "Average Heartrate":        Fields(19, 1) = StravaJsonValue(JSON, "average_heartrate")
      Fields(20, 0) = "Max Heartrate":            Fields(20, 1) = StravaJsonValue(JSON, "max_heartrate")
      Fields(21, 0) = "Calories":                 Fields(21, 1) = StravaJsonValue(JSON, "calories")
      Fields(22, 0) = "Kudos Count":              Fields(22, 1) = StravaJsonValue(JSON, "kudos_count")
      Fields(23, 0) = "Achievement Count":        Fields(23, 1) = StravaJsonValue(JSON, "achievement_count")
      Fields(24, 0) = "Strava URL":               Fields(24, 1) = "https://www.strava.com/activities/" & StravaJsonValue(JSON, "id")

      Dim i As Long
      For i = 0 To 24
            WS.Cells(i + 1, 1).Value = Fields(i, 0)
            WS.Cells(i + 1, 2).Value = Fields(i, 1)
      Next i

      WS.Columns("A:B").AutoFit
      WS.Activate
      MsgBox "Activity details written to sheet '" & WS.Name & "'.", vbInformation, " Strava"

End Sub

Sub StravaGetAthleteStats()
' Fetches the authenticated athlete's stats (totals for all activities)
' and writes them to a worksheet called "Strava_Stats".

      Dim Token As String
      Token = StravaAccessToken()
      If Token = "" Then Exit Sub

      ' Need athlete ID first
      Dim AthleteJSON As String
      AthleteJSON = StravaApiGet("/athlete", Token)
      If AthleteJSON = "" Then Exit Sub

      Dim AthleteID As String
      AthleteID = StravaJsonValue(AthleteJSON, "id")
      If AthleteID = "" Then
            MsgBox "Could not retrieve athlete ID.", vbCritical, " Strava"
            Exit Sub
      End If

      Dim JSON As String
      JSON = StravaApiGet("/athletes/" & AthleteID & "/stats", Token)
      If JSON = "" Then Exit Sub

      Dim WS As Worksheet
      Set WS = StravaGetOrCreateSheet("Strava_Stats")
      WS.Cells.ClearContents

      Dim Row As Long
      Row = 1
      WS.Cells(Row, 1).Value = "Category":     WS.Cells(Row, 2).Value = "Metric": WS.Cells(Row, 3).Value = "Value"
      WS.Rows(Row).Font.Bold = True
      Row = Row + 1

      ' Recent totals (last 4 weeks)
      Dim RecentRun As String
      RecentRun = StravaJsonBlock(JSON, "recent_run_totals")
      Call StravaWriteStatRow(WS, Row, "Recent Runs", "Count",          StravaJsonValue(RecentRun, "count")):           Row = Row + 1
      Call StravaWriteStatRow(WS, Row, "Recent Runs", "Distance (m)",   StravaJsonValue(RecentRun, "distance")):        Row = Row + 1
      Call StravaWriteStatRow(WS, Row, "Recent Runs", "Moving Time (s)", StravaJsonValue(RecentRun, "moving_time")):    Row = Row + 1
      Call StravaWriteStatRow(WS, Row, "Recent Runs", "Elevation (m)",  StravaJsonValue(RecentRun, "elevation_gain")): Row = Row + 1

      Dim RecentRide As String
      RecentRide = StravaJsonBlock(JSON, "recent_ride_totals")
      Call StravaWriteStatRow(WS, Row, "Recent Rides", "Count",          StravaJsonValue(RecentRide, "count")):          Row = Row + 1
      Call StravaWriteStatRow(WS, Row, "Recent Rides", "Distance (m)",   StravaJsonValue(RecentRide, "distance")):       Row = Row + 1
      Call StravaWriteStatRow(WS, Row, "Recent Rides", "Moving Time (s)", StravaJsonValue(RecentRide, "moving_time")):   Row = Row + 1
      Call StravaWriteStatRow(WS, Row, "Recent Rides", "Elevation (m)",  StravaJsonValue(RecentRide, "elevation_gain")): Row = Row + 1

      ' All-time totals
      Dim AllRun As String
      AllRun = StravaJsonBlock(JSON, "all_run_totals")
      Call StravaWriteStatRow(WS, Row, "All-time Runs", "Count",          StravaJsonValue(AllRun, "count")):           Row = Row + 1
      Call StravaWriteStatRow(WS, Row, "All-time Runs", "Distance (m)",   StravaJsonValue(AllRun, "distance")):        Row = Row + 1
      Call StravaWriteStatRow(WS, Row, "All-time Runs", "Moving Time (s)", StravaJsonValue(AllRun, "moving_time")):    Row = Row + 1
      Call StravaWriteStatRow(WS, Row, "All-time Runs", "Elevation (m)",  StravaJsonValue(AllRun, "elevation_gain")): Row = Row + 1

      Dim AllRide As String
      AllRide = StravaJsonBlock(JSON, "all_ride_totals")
      Call StravaWriteStatRow(WS, Row, "All-time Rides", "Count",          StravaJsonValue(AllRide, "count")):           Row = Row + 1
      Call StravaWriteStatRow(WS, Row, "All-time Rides", "Distance (m)",   StravaJsonValue(AllRide, "distance")):        Row = Row + 1
      Call StravaWriteStatRow(WS, Row, "All-time Rides", "Moving Time (s)", StravaJsonValue(AllRide, "moving_time")):    Row = Row + 1
      Call StravaWriteStatRow(WS, Row, "All-time Rides", "Elevation (m)",  StravaJsonValue(AllRide, "elevation_gain")): Row = Row + 1

      WS.Columns("A:C").AutoFit
      WS.Activate
      MsgBox "Stats written to sheet '" & WS.Name & "'.", vbInformation, " Strava"

End Sub

Sub StravaRefreshAccessToken()
' Manually refreshes the access token using the stored refresh token.
' Called automatically by StravaAccessToken when the token has expired.

      Dim ClientID     As String
      Dim ClientSecret As String
      Dim RefreshToken As String

      ClientID     = StravaGetSetting("ClientID")
      ClientSecret = StravaGetSetting("ClientSecret")
      RefreshToken = StravaGetSetting("RefreshToken")

      If ClientID = "" Or ClientSecret = "" Or RefreshToken = "" Then
            MsgBox "Missing credentials. Please run StravaSetup.", vbCritical, " Strava"
            Exit Sub
      End If

      Dim Body As String
      Body = "client_id=" & ClientID & _
             "&client_secret=" & ClientSecret & _
             "&grant_type=refresh_token" & _
             "&refresh_token=" & RefreshToken

      Dim JSON As String
      JSON = StravaHttpPost(STRAVA_TOKEN_URL, Body)

      If StravaJsonValue(JSON, "access_token") <> "" Then
            SaveSetting REG_APP, REG_CREDS, "AccessToken",  StravaJsonValue(JSON, "access_token")
            SaveSetting REG_APP, REG_CREDS, "RefreshToken", StravaJsonValue(JSON, "refresh_token")
            SaveSetting REG_APP, REG_CREDS, "ExpiresAt",    StravaJsonValue(JSON, "expires_at")
            MsgBox "Access token refreshed successfully.", vbInformation, " Strava"
      Else
            MsgBox "Token refresh failed: " & JSON, vbCritical, " Strava"
      End If

End Sub

Sub StravaClearCredentials()
' Deletes all saved Strava credentials from the Windows Registry.

      If MsgBox("This will delete all saved Strava credentials. Continue?", _
                vbYesNo + vbQuestion, " Strava") = vbNo Then Exit Sub
      DeleteSetting REG_APP, REG_CREDS
      MsgBox "Credentials cleared. Run StravaSetup to reconnect.", vbInformation, " Strava"

End Sub


' =============================================================================
' PRIVATE HELPERS
' =============================================================================

Private Function StravaAccessToken() As String
' Returns a valid access token, refreshing it automatically if expired.

      Dim AccessToken  As String
      Dim ExpiresAtStr As String
      Dim ExpiresAt    As Long

      AccessToken  = StravaGetSetting("AccessToken")
      ExpiresAtStr = StravaGetSetting("ExpiresAt")

      If AccessToken = "" Then
            MsgBox "Not authenticated. Please run StravaSetup.", vbCritical, " Strava"
            Exit Function
      End If

      ' Refresh if the token expires within 5 minutes
      If ExpiresAtStr <> "" Then
            ExpiresAt = CLng(ExpiresAtStr)
            If Now() >= (ExpiresAt - 300) / 86400 + #1/1/1970# Then
                  Call StravaRefreshAccessToken
                  AccessToken = StravaGetSetting("AccessToken")
            End If
      End If

      StravaAccessToken = AccessToken

End Function

Private Function StravaExchangeCode(ByVal AuthCode As String) As Boolean
' POSTs the authorisation code to Strava and saves the returned tokens.

      Dim ClientID     As String
      Dim ClientSecret As String
      ClientID     = StravaGetSetting("ClientID")
      ClientSecret = StravaGetSetting("ClientSecret")

      Dim Body As String
      Body = "client_id=" & ClientID & _
             "&client_secret=" & ClientSecret & _
             "&code=" & AuthCode & _
             "&grant_type=authorization_code"

      Dim JSON As String
      JSON = StravaHttpPost(STRAVA_TOKEN_URL, Body)

      If StravaJsonValue(JSON, "access_token") <> "" Then
            SaveSetting REG_APP, REG_CREDS, "AccessToken",  StravaJsonValue(JSON, "access_token")
            SaveSetting REG_APP, REG_CREDS, "RefreshToken", StravaJsonValue(JSON, "refresh_token")
            SaveSetting REG_APP, REG_CREDS, "ExpiresAt",    StravaJsonValue(JSON, "expires_at")
            StravaExchangeCode = True
      End If

End Function

Private Function StravaApiGet(ByVal Endpoint As String, ByVal Token As String) As String
' Makes an authenticated GET request to the Strava API.

      On Error Resume Next

      With CreateObject(HTTPRequest_Client)
            .Open "GET", STRAVA_API_BASE & Endpoint, False
            .SetRequestHeader "Authorization", "Bearer " & Token
            .SetRequestHeader "Accept", "application/json"
            .Send
            If .Status = 200 Then
                  StravaApiGet = .ResponseText
            ElseIf .Status = 401 Then
                  MsgBox "Authorisation error (401). Try running StravaRefreshAccessToken.", vbCritical, " Strava"
            ElseIf .Status = 429 Then
                  MsgBox "Rate limit exceeded (429). Please wait 15 minutes and try again.", vbExclamation, " Strava"
            Else
                  MsgBox "API error " & .Status & ": " & .ResponseText, vbCritical, " Strava"
            End If
      End With

End Function

Private Function StravaHttpPost(ByVal URL As String, ByVal Body As String) As String
' Makes an unauthenticated POST request (used for token exchange/refresh).

      On Error Resume Next

      With CreateObject(HTTPRequest_Client)
            .Open "POST", URL, False
            .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            .Send Body
            StravaHttpPost = .ResponseText
      End With

End Function

Private Function StravaGetSetting(ByVal Key As String) As String
      On Error Resume Next
      StravaGetSetting = GetSetting(REG_APP, REG_CREDS, Key, "")
End Function

Private Function StravaExtractParam(ByVal URL As String, ByVal ParamName As String) As String
' Extracts a query-string parameter value from a URL string.

      Dim SearchFor As String
      SearchFor = ParamName & "="

      Dim StartPos As Long
      StartPos = InStr(URL, SearchFor)
      If StartPos = 0 Then Exit Function

      StartPos = StartPos + Len(SearchFor)
      Dim EndPos As Long
      EndPos = InStr(StartPos, URL, "&")
      If EndPos = 0 Then EndPos = Len(URL) + 1

      StravaExtractParam = Mid$(URL, StartPos, EndPos - StartPos)

End Function

Private Function StravaJsonValue(ByVal JSON As String, ByVal Key As String) As String
' Extracts a simple scalar value from a flat JSON string.
' Handles both quoted strings and unquoted numbers/booleans.

      Dim SearchFor As String
      SearchFor = """" & Key & """"

      Dim KeyPos As Long
      KeyPos = InStr(JSON, SearchFor)
      If KeyPos = 0 Then Exit Function

      Dim ColonPos As Long
      ColonPos = InStr(KeyPos + Len(SearchFor), JSON, ":")
      If ColonPos = 0 Then Exit Function

      Dim ValueStart As Long
      ValueStart = ColonPos + 1

      ' Skip whitespace
      Do While Mid$(JSON, ValueStart, 1) = " "
            ValueStart = ValueStart + 1
      Loop

      Dim FirstChar As String
      FirstChar = Mid$(JSON, ValueStart, 1)

      If FirstChar = """" Then
            ' Quoted string value
            ValueStart = ValueStart + 1
            Dim EndQuote As Long
            EndQuote = InStr(ValueStart, JSON, """")
            If EndQuote > 0 Then
                  StravaJsonValue = Mid$(JSON, ValueStart, EndQuote - ValueStart)
            End If
      ElseIf FirstChar = "{" Or FirstChar = "[" Or FirstChar = "n" Then
            ' Object, array, or null — return empty
            StravaJsonValue = ""
      Else
            ' Unquoted value (number, boolean)
            Dim ValueEnd As Long
            ValueEnd = ValueStart
            Do While ValueEnd <= Len(JSON)
                  Dim Ch As String
                  Ch = Mid$(JSON, ValueEnd, 1)
                  If Ch = "," Or Ch = "}" Or Ch = "]" Or Ch = " " Then Exit Do
                  ValueEnd = ValueEnd + 1
            Loop
            StravaJsonValue = Mid$(JSON, ValueStart, ValueEnd - ValueStart)
      End If

End Function

Private Function StravaJsonBlock(ByVal JSON As String, ByVal Key As String) As String
' Extracts a nested JSON object block (between matching braces) for a given key.

      Dim SearchFor As String
      SearchFor = """" & Key & """:"

      Dim KeyPos As Long
      KeyPos = InStr(JSON, SearchFor)
      If KeyPos = 0 Then Exit Function

      Dim OpenPos As Long
      OpenPos = InStr(KeyPos, JSON, "{")
      If OpenPos = 0 Then Exit Function

      Dim Depth As Long
      Depth = 0

      Dim i As Long
      For i = OpenPos To Len(JSON)
            Dim Ch As String
            Ch = Mid$(JSON, i, 1)
            If Ch = "{" Then Depth = Depth + 1
            If Ch = "}" Then Depth = Depth - 1
            If Depth = 0 Then
                  StravaJsonBlock = Mid$(JSON, OpenPos, i - OpenPos + 1)
                  Exit Function
            End If
      Next i

End Function

Private Sub StravaWriteStatRow(ByVal WS As Worksheet, ByVal Row As Long, _
                               ByVal Category As String, ByVal Metric As String, _
                               ByVal Value As String)
      WS.Cells(Row, 1).Value = Category
      WS.Cells(Row, 2).Value = Metric
      WS.Cells(Row, 3).Value = Value
End Sub

Private Function StravaGetOrCreateSheet(ByVal SheetName As String) As Worksheet
' Returns the named worksheet, creating it if it does not exist.

      On Error Resume Next
      Dim WS As Worksheet
      Set WS = ThisWorkbook.Sheets(SheetName)
      If WS Is Nothing Then
            Set WS = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            WS.Name = SheetName
      End If
      Set StravaGetOrCreateSheet = WS

End Function

Private Function CDblSafe(ByVal Value As String) As Double
      On Error Resume Next
      If Value = "" Or Value = "null" Then
            CDblSafe = 0
      Else
            CDblSafe = CDbl(Value)
      End If
End Function

Private Function CLngSafe(ByVal Value As String) As Long
      On Error Resume Next
      If Value = "" Or Value = "null" Then
            CLngSafe = 0
      Else
            CLngSafe = CLng(Value)
      End If
End Function
