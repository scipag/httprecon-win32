Attribute VB_Name = "modReportingXml"
Option Explicit

Public Function GenerateXmlReport(ByRef iMatches As Integer, ByRef iResponses As Integer, ByRef iHitlistSize As Integer) As String
    Dim cReport As Concat

    Set cReport = New Concat

    Call ChangeStatusBar("Generate XML Report...")

    With cReport
        .Concat "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>" & vbCrLf
        
        .Concat "<scan " & _
            "application='" & APP_NAME & "' " & _
            "agent='" & XmlEncode(req_agent_name) & "' " & _
            "noredirects='" & req_agent_noredirect & "' " & _
            "auditor='" & GetLocalUsername & "' " & _
            "scandate='" & scan_date & "' " & _
            "scantime='" & scan_time & "' " & _
            "tests='" & tests_count & "' " & _
            "testgetexisting='" & scan_test_getexisting & "' " & _
            "testgetnonexisting='" & scan_test_getnonexisting & "' " & _
            "testgetlong='" & scan_test_getlong & "' " & _
            "testheadexisting='" & scan_test_head & "' " & _
            "testoptions='" & scan_test_options & "' " & _
            "testwrongmethod='" & scan_test_wrongmethod & "' " & _
            "testnonexistingmethod='" & scan_test_nonexistingmethod & "' " & _
            "testwrongprotocol='" & scan_test_wrongprotocol & "' " & _
            "testattack='" & scan_test_attack & "' " & _
            "reportdate='" & Date & "' " & _
            "reporttime='" & Time & "'>" & vbCrLf
        
        .Concat vbTab & "<server " & _
            "host='" & scan_targethost & "' " & _
            "port='" & scan_targetport & "' " & _
            "ssl='" & scan_targetsecure & "'>" & vbCrLf
        
        If (iMatches = 1) Then
            .Concat GenerateHitListXml(frmMain.lsvResults, iHitlistSize)
        End If

        If (iResponses = 1) Then
            .Concat ShowTestCaseXml(APP_TESTNAME_GETEXISTING, response_getexist, scan_test_getexisting, timing_getexist, "GET", req_resource_available, req_protocol_legitimate) & vbCrLf
            .Concat ShowTestCaseXml(APP_TESTNAME_GETLONG, response_getlongrequest, scan_test_getlong, timing_getlongrequest, "GET", String$(req_longrequest_length, req_longrequest_char), req_protocol_legitimate) & vbCrLf
            .Concat ShowTestCaseXml(APP_TESTNAME_GETNONEXISTING, response_get_nonexistent, scan_test_getnonexisting, timing_get_nonexistent, "GET", req_resource_notavailable, req_protocol_legitimate) & vbCrLf
            .Concat ShowTestCaseXml(APP_TESTNAME_HEADEXISTING, response_head, scan_test_head, timing_head, "HEAD", req_resource_available, req_protocol_legitimate) & vbCrLf
            .Concat ShowTestCaseXml(APP_TESTNAME_OPTIONS, response_options, scan_test_options, timing_options, "OPTIONS", "/", req_protocol_legitimate) & vbCrLf
            .Concat ShowTestCaseXml(APP_TESTNAME_DELETEEXISTING, response_delete, scan_test_wrongmethod, timing_delete, req_method_notallowed, req_resource_available, req_protocol_legitimate) & vbCrLf
            .Concat ShowTestCaseXml(APP_TESTNAME_WRONGMETHOD, response_testmethod, scan_test_nonexistingmethod, timing_testmethod, req_method_notexisting, req_resource_available, req_protocol_legitimate) & vbCrLf
            .Concat ShowTestCaseXml(APP_TESTNAME_WRONGVERSION, response_protocolversion, scan_test_wrongprotocol, timing_protocolversion, "GET", req_resource_available, req_protocol_wrong) & vbCrLf
            .Concat ShowTestCaseXml(APP_TESTNAME_ATTACKREQUEST, response_attackrequest, scan_test_attack, timing_attackrequest, "GET", req_resource_attack, req_protocol_legitimate) & vbCrLf
        End If
    
        .Concat vbTab & "</server>" & vbCrLf
        .Concat "</scan>" & vbCrLf
    
        GenerateXmlReport = .Value
    End With

    Call ChangeStatusBarDone
End Function

Public Function ShowTestCaseXml(ByRef sName As String, ByRef sResponse As String, ByRef iEnabled As Integer, ByRef sTiming As Single, ByRef sRequest As String, ByRef sResource As String, ByRef sProtocol As String) As String
    Dim cTestcase As Concat
    Dim iLength As Integer
    
    Set cTestcase = New Concat
    
    iLength = Len(sResponse)
    
    With cTestcase
        .Concat vbTab & vbTab & "<response name='" & sName & "' " & _
            "enabled='" & iEnabled & "' " & _
            "length='" & iLength & "' " & _
            "timing='" & NormalizeTiming(sTiming) & "' " & _
            "request='" & XmlEncode(sRequest) & "' " & _
            "resource='" & XmlEncode(sResource) & "' " & _
            "protocol='" & XmlEncode(sProtocol) & "'>"
        .Concat "<![CDATA[" & sResponse & "]]>"
        .Concat "</response>"
    
        ShowTestCaseXml = .Value
    End With
End Function

Public Function GenerateHitListXml(ByRef lListView As ListView, ByRef iCount As Integer) As String
    Dim cResults As Concat
    Dim iListItemCount As Integer
    Dim i As Integer
    
    Set cResults = New Concat
    
    iListItemCount = lListView.ListItems.Count
    
    If (iListItemCount > iCount) Then
        iListItemCount = iCount
    End If
    
    For i = 1 To iListItemCount
        cResults.Concat vbTab & vbTab & "<match " & _
            "name='" & XmlEncode(lListView.ListItems(i).ListSubItems(1).Text) & "' " & _
            "hits='" & lListView.ListItems(i).ListSubItems(2).Text & "' " & _
            "confidence='" & Round(lListView.ListItems(i).ListSubItems(3).Text, 2) & "' " & _
            "position='" & i & "' " & _
            "/>" & vbCrLf
    Next i
    
    GenerateHitListXml = cResults.Value
End Function

Public Function XmlEncode(ByRef sInput As String) As String
    Dim sOutput As String
    
    sOutput = Replace$(sInput, "&", "&amp;", 1, , vbBinaryCompare)
    sOutput = Replace$(sOutput, "<", "&gt;", 1, , vbBinaryCompare)
    sOutput = Replace$(sOutput, ">", "&lt;", 1, , vbBinaryCompare)
    sOutput = Replace$(sOutput, "'", "&apos;", 1, , vbBinaryCompare)
    sOutput = Replace$(sOutput, ChrW$(34), "&quot;", 1, , vbBinaryCompare)
    
    XmlEncode = sOutput
End Function

