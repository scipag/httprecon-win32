Attribute VB_Name = "modReportingHtml"
Option Explicit

Public Function GenerateHtmlReport(ByRef iPreamble As Integer, ByRef iContents As Integer, ByRef iSummary As Integer, ByRef iMatches As Integer, ByRef iResponses As Integer, ByRef iDetails As Integer, ByRef iHitlistSize As Integer) As String
    Dim txtScheme As String
    Dim cReport As Concat

    Set cReport = New Concat

    Call ChangeStatusBar("Generate HTML Report...")

    If (scan_targetsecure = 1) Then
        txtScheme = "https"
    Else
        txtScheme = "http"
    End If

    With cReport
        .Concat "<?xml version='1.0' encoding='iso-8859-1' ?><!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.1//EN"" ""http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd""> " & vbCrLf
        .Concat "<html xmlns=""http://www.w3.org/1999/xhtml"">" & vbCrLf
        .Concat "<head>" & vbCrLf
        .Concat "<title>" & APP_NAME & " Report (" & txtScheme & "://" & HtmlEncode(scan_targethost) & ":" & scan_targetport & "/)</title>" & vbCrLf
        .Concat GetEmbeddedCSS()
        .Concat "<meta name='keywords' content='httprecon, Webserver, web server, HTTPD, http, Fingerprinting, Report' />" & vbCrLf
        .Concat "<meta name='description' content='The httprecon project is doing some research in the field of web server fingerprinting, also known as http fingerprinting. The goal is the highly accurate identification of given httpd implementations.' />" & vbCrLf
        .Concat "</head>" & vbCrLf
        .Concat "<body>" & vbCrLf
        
        .Concat "<h3>" & APP_NAME & " Report</h3>" & vbCrLf
        If (iPreamble = 1) Then
            .Concat "Target: <a href='" & txtScheme & "://" & HtmlEncode(scan_targethost) & ":" & scan_targetport & "'>" & txtScheme & "://" & HtmlEncode(scan_targethost) & ":" & scan_targetport & "</a><br />" & vbCrLf
            .Concat "Tests: " & tests_count & " test cases<br />" & vbCrLf
            .Concat "Auditor: " & GetLocalUsername & "<br />" & vbCrLf
            .Concat "Scan: " & scan_date & " - " & scan_time & "<br />" & vbCrLf
            .Concat "Export: " & Date & " - " & Time & vbCrLf
        End If
        
        If (iContents = 1) Then
            .Concat "<h4 id='contents'>Contents</h4>" & vbCrLf
            .Concat "<ol style='list-style-type:decimal'>" & vbCrLf
            If (iSummary = 1) Then
                .Concat "<li><a href='#summary'>Summary</a></li>" & vbCrLf
            End If
            If (iMatches = 1) Then
                .Concat "<li><a href='#matches'>Matches</a></li>" & vbCrLf
            End If
            If (iResponses = 1) Then
                .Concat "<li><a href='#responses'>Responses</a></li>" & vbCrLf
            End If
            If (iDetails = 1) Then
                .Concat "<li><a href='#details'>Details</a></li>" & vbCrLf
            End If
            .Concat "</ol>" & vbCrLf
        End If
        
        If (iSummary = 1) Then
            .Concat "<h4 id='summary'>Summary <a href='#'>&uarr;</a></h4>" & vbCrLf
            .Concat "An advanced web server fingerprinting for the host " & HtmlEncode(scan_targethost) & " and port tcp/" & scan_targetport & " was done with " & tests_count & " test cases at " & scan_date & " " & scan_time & ".<br /><br />" & vbCrLf
            .Concat "This analysis was able to determine the target httpd service as " & HtmlEncode(scan_besthitname) & " with " & scan_besthitcount & " fingerprint hits in the database." & vbCrLf
        End If
        
        If (iMatches = 1) Then
            .Concat "<h4 id='matches'>List of Matches <a href='#'>&uarr;</a></h4>" & vbCrLf
            .Concat GenerateHitListHtml(iHitlistSize)
        End If
        
        If (iResponses = 1) Then
            .Concat "<h4 id='responses'>HTTP Response Header <a href='#'>&uarr;</a></h4>" & vbCrLf
            .Concat "Timing Minimum: " & Round(MinimumTiming, timing_decimals) & " seconds<br />" & vbCrLf
            .Concat "Timing Maximum: " & Round(MaximumTiming, timing_decimals) & " seconds<br />" & vbCrLf
            .Concat "Timing Average: " & Round(AverageTiming, timing_decimals) & " seconds<br /><br />" & vbCrLf
            .Concat ShowTestCaseHtml(APP_TESTNAME_GETEXISTING, response_getexist, timing_getexist, scan_test_getexisting) & vbCrLf
            .Concat ShowTestCaseHtml(APP_TESTNAME_GETLONG, response_getlongrequest, timing_getlongrequest, scan_test_getlong) & vbCrLf
            .Concat ShowTestCaseHtml(APP_TESTNAME_GETNONEXISTING, response_get_nonexistent, timing_get_nonexistent, scan_test_getnonexisting) & vbCrLf
            .Concat ShowTestCaseHtml(APP_TESTNAME_HEADEXISTING, response_head, timing_head, scan_test_head) & vbCrLf
            .Concat ShowTestCaseHtml(APP_TESTNAME_OPTIONS, response_options, timing_options, scan_test_options) & vbCrLf
            .Concat ShowTestCaseHtml(APP_TESTNAME_DELETEEXISTING, response_delete, timing_delete, scan_test_wrongmethod) & vbCrLf
            .Concat ShowTestCaseHtml(APP_TESTNAME_WRONGMETHOD, response_testmethod, timing_testmethod, scan_test_nonexistingmethod) & vbCrLf
            .Concat ShowTestCaseHtml(APP_TESTNAME_WRONGVERSION, response_protocolversion, timing_protocolversion, scan_test_wrongprotocol) & vbCrLf
            .Concat ShowTestCaseHtml(APP_TESTNAME_ATTACKREQUEST, response_attackrequest, timing_attackrequest, scan_test_attack) & vbCrLf
        End If
    
        If (iDetails = 1) Then
            .Concat "<h4 id='details'>Fingerprint Details <a href='#'>&uarr;</a></h4>" & vbCrLf
            .Concat ShowTestCaseHtml(APP_TESTNAME_GETEXISTING, GenerateFingerprintDetails(response_getexist), timing_getexist, 0) & vbCrLf
            .Concat ShowTestCaseHtml(APP_TESTNAME_GETLONG, GenerateFingerprintDetails(response_getlongrequest), timing_getlongrequest, 0) & vbCrLf
            .Concat ShowTestCaseHtml(APP_TESTNAME_GETNONEXISTING, GenerateFingerprintDetails(response_get_nonexistent), timing_get_nonexistent, 0) & vbCrLf
            .Concat ShowTestCaseHtml(APP_TESTNAME_HEADEXISTING, GenerateFingerprintDetails(response_head), timing_head, 0) & vbCrLf
            .Concat ShowTestCaseHtml(APP_TESTNAME_OPTIONS, GenerateFingerprintDetails(response_options), timing_options, 0) & vbCrLf
            .Concat ShowTestCaseHtml(APP_TESTNAME_DELETEEXISTING, GenerateFingerprintDetails(response_delete), timing_delete, 0) & vbCrLf
            .Concat ShowTestCaseHtml(APP_TESTNAME_WRONGMETHOD, GenerateFingerprintDetails(response_testmethod), timing_testmethod, 0) & vbCrLf
            .Concat ShowTestCaseHtml(APP_TESTNAME_WRONGVERSION, GenerateFingerprintDetails(response_protocolversion), timing_protocolversion, 0) & vbCrLf
            .Concat ShowTestCaseHtml(APP_TESTNAME_ATTACKREQUEST, GenerateFingerprintDetails(response_attackrequest), timing_attackrequest, 0) & vbCrLf
        End If
    
        .Concat "<br /><div id='bottom' class='copyright'>&copy; 2007-" & Year(Now) & " by <a href='" & APP_WEBSITE_URL & "'>" & APP_COPYRIGHT_OWNER & "</a></div>" & vbCrLf
    
        .Concat "</body>" & vbCrLf
        .Concat "</html>" & vbCrLf
    
        GenerateHtmlReport = .Value
    End With

    Call ChangeStatusBarDone
End Function

Public Function ShowTestCaseHtml(ByRef sName As String, ByRef sResponse As String, ByRef sTiming As Single, ByRef iTestEnabled As Integer) As String
    Dim cTestcase As Concat
    Dim iLength As Integer
    
    Set cTestcase = New Concat
    
    iLength = Len(sResponse)
    
    With cTestcase
        .Concat "<table class='table' id='" & sName & "'>" & vbCrLf
        .Concat "<tr class='databaseheader'><td>" & HtmlEncode(sName) & "</td><tr>" & vbCrLf
        If (iLength) Then
            .Concat "<tr><td class='response' title='Length: " & iLength & " bytes / Timing: " & NormalizeTiming(sTiming) & "'>" & HtmlEncode(sResponse) & "</td><tr>" & vbCrLf
        ElseIf (iTestEnabled = 0) Then
            .Concat "<tr class='databaseline'><td class='databaseline'>test not enabled</td><tr>" & vbCrLf
        Else
            .Concat "<tr class='databaseline'><td class='databaseline'>no response available</td><tr>" & vbCrLf
        End If
        .Concat "</table><br />" & vbCrLf
        
        ShowTestCaseHtml = .Value
    End With
End Function

Public Function GenerateHitListHtml(ByRef iCount As Integer) As String
    Dim cResults As Concat
    Dim iListItemCount As Integer
    Dim i As Integer
    
    Set cResults = New Concat
    
    iListItemCount = frmMain.lsvResults.ListItems.Count
    
    If (iListItemCount > iCount) Then
        iListItemCount = iCount
    End If
    
    With cResults
        .Concat "<table class='table'><tr class='databaseheader'><td style='width:20px'>&nbsp;</td><td>Name</td><td style='width:60px'>Hits</td><td style='width:60px'>Match</td></tr>" & vbCrLf
        For i = 1 To iListItemCount
             .Concat "<tr class='databaseline'><td style='text-align:right' class='databaseline'>" & i & ".</td><td class='databaseline'>" & HtmlEncode(frmMain.lsvResults.ListItems(i).ListSubItems(1).Text) & "</td><td class='databaseline'>" & HtmlEncode(frmMain.lsvResults.ListItems(i).ListSubItems(2).Text) & "</td><td class='databaseline'>" & Round(frmMain.lsvResults.ListItems(i).ListSubItems(3).Text, 2) & "% </td></tr>" & vbCrLf
        Next i
        .Concat "</table>" & vbCrLf
    
        GenerateHitListHtml = .Value
    End With
End Function

Public Function HtmlEncode(ByRef sInput As String) As String
    Dim sOutput As String
    
    sOutput = Replace$(sOutput, "&", "&amp;", 1, , vbBinaryCompare)
    
    'HTML encoding (prevents Cross Site Scripting)
    sOutput = Replace$(sInput, "<", "&gt;", 1, , vbBinaryCompare)
    sOutput = Replace$(sOutput, ">", "&lt;", 1, , vbBinaryCompare)
    sOutput = Replace$(sOutput, ChrW$(34), "&quot;", 1, , vbBinaryCompare)
    
    'UTF-7 encoding (fixes flaw found by Stefan Friedli)
    sOutput = Replace$(sInput, "+ADw-", "&gt;", 1, , vbBinaryCompare)
    sOutput = Replace$(sOutput, "+AD4-", "&lt;", 1, , vbBinaryCompare)
    sOutput = Replace$(sOutput, "+ACc-", "&quot;", 1, , vbBinaryCompare)
    
    sOutput = Replace$(sOutput, vbTab, " ", 1, , vbBinaryCompare)
    sOutput = Replace$(sOutput, vbCrLf, "<br />" & vbCrLf, 1, , vbBinaryCompare)
    
    HtmlEncode = sOutput
End Function

Public Function GetEmbeddedCSS() As String
    Dim cCSS As Concat
    
    Set cCSS = New Concat

    With cCSS
        .Concat "<style type=""text/css"">" & vbCrLf
        .Concat "<!-- " & vbCrLf
        
        .Concat "body{" & vbCrLf
        .Concat "font-family:verdana;" & vbCrLf
        .Concat "font-size:11px;" & vbCrLf
        .Concat "color:black;" & vbCrLf
        .Concat "}" & vbCrLf
        
        .Concat "a{" & vbCrLf
        .Concat "color:darkred;" & vbCrLf
        .Concat "text-decoration:none;" & vbCrLf
        .Concat "}" & vbCrLf
    
        .Concat "a:hover{" & vbCrLf
        .Concat "color:red;" & vbCrLf
        .Concat "}" & vbCrLf
    
        .Concat "table.table{" & vbCrLf
        .Concat "border:1px solid darkred;" & vbCrLf
        .Concat "width:640px;" & vbCrLf
        .Concat "}" & vbCrLf
    
        .Concat "tr.databaseheader{" & vbCrLf
        .Concat "font-weight:bold;" & vbCrLf
        .Concat "background-color:darkred;" & vbCrLf
        .Concat "color:white;" & vbCrLf
        .Concat "}" & vbCrLf
            
        .Concat "tr.databaseline:hover{" & vbCrLf
        .Concat "background-color:lightgrey;" & vbCrLf
        .Concat "}" & vbCrLf
    
        .Concat "td.databaseline{" & vbCrLf
        .Concat "border:1px solid lightgrey;" & vbCrLf
        .Concat "}" & vbCrLf
    
        .Concat "td.response{" & vbCrLf
        .Concat "font-family:'courier new';" & vbCrLf
        .Concat "color:lightgreen;" & vbCrLf
        .Concat "background:black;" & vbCrLf
        .Concat "}" & vbCrLf
    
        .Concat ".copyright{" & vbCrLf
        .Concat "font-size:10px;" & vbCrLf
        .Concat "}" & vbCrLf
    
        .Concat " -->" & vbCrLf
        .Concat "</style>" & vbCrLf
    
        GetEmbeddedCSS = .Value
    End With
End Function
