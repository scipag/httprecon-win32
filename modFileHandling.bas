Attribute VB_Name = "modFileHandling"
Option Explicit

Public Function GenerateFingerprintXML(ByRef bIncludeScanDetails As Boolean) As String
    Dim sFullFingerprint As Concat
    
    Set sFullFingerprint = New Concat
        
    Call ChangeStatusBar("Generate XML Fingerprint...")

    With sFullFingerprint
        If (bIncludeScanDetails = True) Then
            .Concat "<scan_targethost>" & vbCrLf & scan_targethost & vbCrLf & "</scan_targethost>" & vbCrLf
            .Concat "<scan_targetport>" & vbCrLf & scan_targetport & vbCrLf & "</scan_targetport>" & vbCrLf
            .Concat "<scan_targetsecure>" & vbCrLf & scan_targetsecure & vbCrLf & "</scan_targetsecure>" & vbCrLf
            .Concat "<scan_date>" & vbCrLf & scan_date & vbCrLf & "</scan_date>" & vbCrLf
            .Concat "<scan_time>" & vbCrLf & scan_time & vbCrLf & "</scan_time>" & vbCrLf & vbCrLf
        End If
        
        .Concat "<" & APP_TESTNAME_GETEXISTING & ">" & vbCrLf & response_getexist & "</" & APP_TESTNAME_GETEXISTING & ">" & vbCrLf
        .Concat "<" & APP_TESTNAME_GETLONG & ">" & vbCrLf & response_getlongrequest & "</" & APP_TESTNAME_GETLONG & ">" & vbCrLf
        .Concat "<" & APP_TESTNAME_GETNONEXISTING & ">" & vbCrLf & response_get_nonexistent & "</" & APP_TESTNAME_GETNONEXISTING & ">" & vbCrLf
        .Concat "<" & APP_TESTNAME_WRONGVERSION & ">" & vbCrLf & response_protocolversion & "</" & APP_TESTNAME_WRONGVERSION & ">" & vbCrLf
        .Concat "<" & APP_TESTNAME_HEADEXISTING & ">" & vbCrLf & response_head & "</" & APP_TESTNAME_HEADEXISTING & ">" & vbCrLf
        .Concat "<" & APP_TESTNAME_OPTIONS & ">" & vbCrLf & response_options & "</" & APP_TESTNAME_OPTIONS & ">" & vbCrLf
        .Concat "<" & APP_TESTNAME_DELETEEXISTING & ">" & vbCrLf & response_delete & "</" & APP_TESTNAME_DELETEEXISTING & ">" & vbCrLf
        .Concat "<" & APP_TESTNAME_WRONGMETHOD & ">" & vbCrLf & response_testmethod & "</" & APP_TESTNAME_WRONGMETHOD & ">" & vbCrLf
        .Concat "<" & APP_TESTNAME_ATTACKREQUEST & ">" & vbCrLf & response_attackrequest & "</" & APP_TESTNAME_ATTACKREQUEST & ">" & vbCrLf
    End With
    
    Call ChangeStatusBarDone
    
    GenerateFingerprintXML = sFullFingerprint.Value
End Function

Public Sub ReadFingerprintXML(ByRef sFingerprints As String)
    Call ChangeStatusBar("Reading XML Fingerprint...")

    scan_targethost = ExtractFingerprintXML(sFingerprints, "scan_targethost", True, "127.0.0.1")
    scan_targetport = Val(ExtractFingerprintXML(sFingerprints, "scan_targetport", True, "80"))
    scan_date = ExtractFingerprintXML(sFingerprints, "scan_date", True, Date)
    scan_time = ExtractFingerprintXML(sFingerprints, "scan_time", True, Time)
    scan_targetsecure = Val(ExtractFingerprintXML(sFingerprints, "scan_targetsecure", True, "0"))
    
    response_attackrequest = ExtractFingerprintXML(sFingerprints, APP_TESTNAME_ATTACKREQUEST, False, vbNullString)
    response_delete = ExtractFingerprintXML(sFingerprints, APP_TESTNAME_DELETEEXISTING, False, vbNullString)
    response_getexist = ExtractFingerprintXML(sFingerprints, APP_TESTNAME_GETEXISTING, False, vbNullString)
    response_getlongrequest = ExtractFingerprintXML(sFingerprints, APP_TESTNAME_GETLONG, False, vbNullString)
    response_get_nonexistent = ExtractFingerprintXML(sFingerprints, APP_TESTNAME_GETNONEXISTING, False, vbNullString)
    response_head = ExtractFingerprintXML(sFingerprints, APP_TESTNAME_HEADEXISTING, False, vbNullString)
    response_options = ExtractFingerprintXML(sFingerprints, APP_TESTNAME_OPTIONS, False, vbNullString)
    response_testmethod = ExtractFingerprintXML(sFingerprints, APP_TESTNAME_WRONGMETHOD, False, vbNullString)
    response_protocolversion = ExtractFingerprintXML(sFingerprints, APP_TESTNAME_WRONGVERSION, False, vbNullString)
End Sub

Public Function ExtractFingerprintXML(ByRef sFingerprints As String, ByRef sTag As String, ByRef bOneLiner As Boolean, ByRef sDefaultValue As String) As String
    Dim iStart As Integer
    Dim iEnd As Integer
    Dim iContentStart As Integer
    Dim iContentLength As Integer
    Dim iOneLinePosition As Integer
    Dim sResult As String
    
    iStart = InStr(1, sFingerprints, "<" & sTag & ">", vbBinaryCompare)
    
    If (iStart > 0) Then
        iContentStart = iStart + (Len(sTag) + 4)
        If (iContentStart > 0) Then
            iEnd = InStr(1, sFingerprints, "</" & sTag & ">", vbBinaryCompare)
            If (iEnd > iStart) Then
                iContentLength = iEnd - iContentStart
                If (iContentLength > 0) Then
                    If (bOneLiner = True) Then
                        iOneLinePosition = InStr(iContentStart, sFingerprints, vbCrLf, vbBinaryCompare)
                        If (iOneLinePosition) Then
                            sResult = Mid$(sFingerprints, iContentStart, iOneLinePosition - iContentStart)
                        Else
                            sResult = Mid$(sFingerprints, iContentStart, iContentLength)
                        End If
                    Else
                        sResult = Mid$(sFingerprints, iContentStart, iContentLength)
                    End If
                End If
            End If
        End If
    End If
    
    If (LenB(sResult)) Then
        ExtractFingerprintXML = sResult
    Else
        ExtractFingerprintXML = sDefaultValue
    End If
End Function

Public Function StringToFileName(ByRef sString As String) As String
    Dim sOutput As String

    sOutput = Replace(sString, ".", "_", 1, , vbBinaryCompare)
    sOutput = Replace(sOutput, ":", "-", 1, , vbBinaryCompare)
    
    StringToFileName = sOutput
End Function
