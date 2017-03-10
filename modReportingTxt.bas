Attribute VB_Name = "modReportingTxt"
Option Explicit

Public Function GenerateTxtReport(ByRef iPreamble As Integer, ByRef iContents As Integer, ByRef iSummary As Integer, ByRef iMatches As Integer, ByRef iResponses As Integer, ByRef iDetails As Integer, ByRef iHitlistSize As Integer) As String
    Dim txtScheme As String
    Dim cReport As Concat
    Dim i As Integer
    
    Set cReport = New Concat

    Call ChangeStatusBar("Generate TXT Report...")
    
    i = 1

    If (scan_targetsecure = 1) Then
        txtScheme = "https"
    Else
        txtScheme = "http"
    End If

    With cReport
        .Concat APP_NAME & " Report" & vbCrLf
        If (iPreamble = 1) Then
            .Concat "Target: " & txtScheme & "://" & scan_targethost & ":" & scan_targetport & vbCrLf
            .Concat "Tests: " & tests_count & " test cases" & vbCrLf
            .Concat "Auditor: " & GetLocalUsername & vbCrLf
            .Concat "Scan: " & scan_date & " - " & scan_time & vbCrLf
            .Concat "Export: " & Date & " - " & Time & vbCrLf & vbCrLf
        End If
        
        If (iContents = 1) Then
            .Concat i & ". CONTENTS" & vbCrLf & vbCrLf
            If (iSummary = 1) Then
                .Concat "* Summary" & vbCrLf
            End If
            If (iMatches = 1) Then
                .Concat "* Matches" & vbCrLf
            End If
            If (iResponses = 1) Then
                .Concat "* Responses" & vbCrLf
            End If
            If (iDetails = 1) Then
                .Concat "* Details" & vbCrLf
            End If
            .Concat vbCrLf
            i = i + 1
        End If
        
        If (iSummary = 1) Then
            .Concat i & ". SUMMARY" & vbCrLf & vbCrLf
            .Concat "An advanced web server fingerprinting for the host " & scan_targethost & " and port tcp/" & scan_targetport & " was done with " & tests_count & " test cases at " & scan_date & " " & scan_time & "." & vbCrLf & vbCrLf
            .Concat "This analysis was able to determine the target httpd service as " & scan_besthitname & " with " & scan_besthitcount & " fingerprint hits in the database." & vbCrLf & vbCrLf
            i = i + 1
        End If
        
        If (iMatches = 1) Then
            .Concat i & ". LIST OF MATCHES" & vbCrLf & vbCrLf
            .Concat GenerateHitListTxt(frmMain.lsvResults, iHitlistSize) & vbCrLf
            i = i + 1
        End If
        
        If (iResponses = 1) Then
            .Concat i & ". HTTP RESPONSE HEADER" & vbCrLf & vbCrLf
            .Concat "Timing Minimum: " & Round(MinimumTiming, timing_decimals) & " seconds" & vbCrLf
            .Concat "Timing Maximum: " & Round(MaximumTiming, timing_decimals) & " seconds" & vbCrLf
            .Concat "Timing Average: " & Round(AverageTiming, timing_decimals) & " seconds" & vbCrLf
            .Concat vbCrLf
            .Concat ShowTestCaseTxt(APP_TESTNAME_GETEXISTING, response_getexist, scan_test_getexisting) & vbCrLf
            .Concat ShowTestCaseTxt(APP_TESTNAME_GETLONG, response_getlongrequest, scan_test_getlong) & vbCrLf
            .Concat ShowTestCaseTxt(APP_TESTNAME_GETNONEXISTING, response_get_nonexistent, scan_test_getnonexisting) & vbCrLf
            .Concat ShowTestCaseTxt(APP_TESTNAME_HEADEXISTING, response_head, scan_test_head) & vbCrLf
            .Concat ShowTestCaseTxt(APP_TESTNAME_OPTIONS, response_options, scan_test_options) & vbCrLf
            .Concat ShowTestCaseTxt(APP_TESTNAME_DELETEEXISTING, response_delete, scan_test_wrongmethod) & vbCrLf
            .Concat ShowTestCaseTxt(APP_TESTNAME_WRONGMETHOD, response_testmethod, scan_test_nonexistingmethod) & vbCrLf
            .Concat ShowTestCaseTxt(APP_TESTNAME_WRONGVERSION, response_protocolversion, scan_test_wrongprotocol) & vbCrLf
            .Concat ShowTestCaseTxt(APP_TESTNAME_ATTACKREQUEST, response_attackrequest, scan_test_attack) & vbCrLf
            i = i + 1
        End If
    
        If (iDetails = 1) Then
            .Concat i & ". FINGERPRINT DETAILS" & vbCrLf & vbCrLf
            .Concat ShowTestCaseTxt(APP_TESTNAME_GETEXISTING, GenerateFingerprintDetails(response_getexist), 0) & vbCrLf
            .Concat ShowTestCaseTxt(APP_TESTNAME_GETLONG, GenerateFingerprintDetails(response_getlongrequest), 0) & vbCrLf
            .Concat ShowTestCaseTxt(APP_TESTNAME_GETNONEXISTING, GenerateFingerprintDetails(response_get_nonexistent), 0) & vbCrLf
            .Concat ShowTestCaseTxt(APP_TESTNAME_HEADEXISTING, GenerateFingerprintDetails(response_head), 0) & vbCrLf
            .Concat ShowTestCaseTxt(APP_TESTNAME_OPTIONS, GenerateFingerprintDetails(response_options), 0) & vbCrLf
            .Concat ShowTestCaseTxt(APP_TESTNAME_DELETEEXISTING, GenerateFingerprintDetails(response_delete), 0) & vbCrLf
            .Concat ShowTestCaseTxt(APP_TESTNAME_WRONGMETHOD, GenerateFingerprintDetails(response_testmethod), 0) & vbCrLf
            .Concat ShowTestCaseTxt(APP_TESTNAME_WRONGVERSION, GenerateFingerprintDetails(response_protocolversion), 0) & vbCrLf
            .Concat ShowTestCaseTxt(APP_TESTNAME_ATTACKREQUEST, GenerateFingerprintDetails(response_attackrequest), 0) & vbCrLf
            i = i + 1
        End If
    
        .Concat "(c) 2007-" & Year(Now) & " by " & APP_COPYRIGHT_OWNER & vbCrLf & APP_WEBSITE_URL & vbCrLf
    
        GenerateTxtReport = .Value
    End With
    
    Call ChangeStatusBarDone
End Function

Public Function ShowTestCaseTxt(ByRef sName As String, ByRef sResponse As String, ByRef iTestEnabled As Integer) As String
    Dim cTestcase As Concat
    Dim iLength As Integer
    
    Set cTestcase = New Concat
    
    iLength = Len(sResponse)
    
    With cTestcase
        .Concat "--- " & sName & " ---" & vbCrLf & vbCrLf
        If (iLength) Then
            .Concat sResponse
        ElseIf (iTestEnabled = 0) Then
            .Concat "[test not enabled]"
        Else
            .Concat "[no response available]"
        End If
        .Concat vbCrLf & vbCrLf
    
        ShowTestCaseTxt = .Value
    End With
End Function

Public Function GenerateHitListTxt(ByRef lSource As ListView, ByRef iCount As Integer) As String
    Dim cResults As Concat
    Dim iListItemCount As Integer
    Dim i As Integer
    
    Set cResults = New Concat
    
    iListItemCount = lSource.ListItems.Count
    
    If (iListItemCount > iCount) Then
        iListItemCount = iCount
    End If
    
    With cResults
        .Concat BlockAsciiTable(4, " ")
        .Concat BlockAsciiTable(50, "Name")
        .Concat BlockAsciiTable(6, "Hits")
        .Concat "Match" & vbCrLf
        For i = 1 To iListItemCount
             .Concat BlockAsciiTable(4, i & ". ")
             .Concat BlockAsciiTable(50, lSource.ListItems(i).ListSubItems(1).Text)
             .Concat BlockAsciiTable(6, lSource.ListItems(i).ListSubItems(2).Text)
             .Concat Round(lSource.ListItems(i).ListSubItems(3).Text, 2) & "%" & vbCrLf
        Next i
        .Concat vbCrLf
    
        GenerateHitListTxt = .Value
    End With
End Function

Public Function BlockAsciiTable(ByRef iWidth As Integer, ByRef sString As String) As String
    Dim iStringLength As Integer
    
    iStringLength = Len(sString)
    
    If (iStringLength < iWidth) Then
        BlockAsciiTable = sString & String$(iWidth - iStringLength, " ")
    ElseIf (iStringLength > iWidth) Then
        BlockAsciiTable = Mid$(sString, 1, iWidth - 4) & "... "
    Else
        BlockAsciiTable = sString
    End If
End Function
