Attribute VB_Name = "modIdentificationDetails"
Option Explicit

Public Sub IdentifyServerFingerprint(ByRef sFingerprintDirectory As String, ByRef sResponse As String, ByRef sImplementation As String)
    Call ResetResponseHighlight
    
    If (LenB(sResponse)) Then
        Call FindDetailInDatabase(sFingerprintDirectory & app_file_banner, sResponse, sImplementation, "Server: ")
        Call FindDetailInDatabase(sFingerprintDirectory & app_file_xpoweredby, sResponse, sImplementation, "X-Powered-By: ")
        Call FindDetailInDatabase(sFingerprintDirectory & app_file_protocolname, sResponse, sImplementation, vbNullString, "/")
        Call FindDetailInDatabase(sFingerprintDirectory & app_file_protocolversion, sResponse, sImplementation, "/", " ")
        Call FindDetailInDatabase(sFingerprintDirectory & app_file_statuscode, sResponse, sImplementation, " ", " ")
        Call FindDetailInDatabase(sFingerprintDirectory & app_file_statustext, sResponse, sImplementation, " ")
        Call FindDetailInDatabase(sFingerprintDirectory & app_file_optionsallowed, sResponse, sImplementation, "Allow: ")
        Call FindDetailInDatabase(sFingerprintDirectory & app_file_optionspublic, sResponse, sImplementation, "Public: ")
        Call FindDetailInDatabase(sFingerprintDirectory & app_file_contenttype, sResponse, sImplementation, "Content-Type: ")
        Call FindDetailInDatabase(sFingerprintDirectory & app_file_acceptrange, sResponse, sImplementation, "Accept-Ranges: ")
        Call FindDetailInDatabase(sFingerprintDirectory & app_file_connection, sResponse, sImplementation, "Connection: ")
        Call FindDetailInDatabase(sFingerprintDirectory & app_file_cachecontrol, sResponse, sImplementation, "Cache-Control: ")
        Call FindDetailInDatabase(sFingerprintDirectory & app_file_pragma, sResponse, sImplementation, "Pragma: ")
        Call FindDetailInDatabase(sFingerprintDirectory & app_file_htaccessrealm, sResponse, sImplementation, "realm=""", """")
    End If
End Sub

Public Sub FindDetailInDatabase(ByRef sDatabase As String, ByRef sResponse As String, ByRef sImplementation As String, Optional ByRef sBefore As String, Optional ByRef sAfter As String)
    Dim sDatabaseContent() As String
    Dim iDatabaseEntries As Integer
    Dim iDelimiterPosition As Integer
    Dim i As Integer
    
    sDatabaseContent = Split(ReadFile(sDatabase), vbCrLf, , vbBinaryCompare)
    iDatabaseEntries = UBound(sDatabaseContent)
    
    For i = 0 To iDatabaseEntries
        If LenB(sImplementation) Then
        
            If LenB(sDatabaseContent(i)) Then
                iDelimiterPosition = InStr(1, sDatabaseContent(i), APP_DATABASE_DELIMITER, vbBinaryCompare)
                
                If ((iDelimiterPosition - 1) = Len(sImplementation)) Then

                    If (sImplementation = Mid$(sDatabaseContent(i), 1, Len(sImplementation))) Then
                        Call ColorMatch(sResponse, sBefore & Mid$(sDatabaseContent(i), iDelimiterPosition + 1) & sAfter)
                    End If
                End If
            End If
        End If
    Next i

    Call SelectTopOfRichTextBox
End Sub

Public Sub ResetResponseHighlight()
    With frmMain.rtbResponses
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelColor = &HFF00&
    End With
    
    Call SelectTopOfRichTextBox
End Sub

Public Sub SelectTopOfRichTextBox()
    frmMain.rtbResponses.SelStart = 0
    frmMain.rtbResponses.SelLength = 0
End Sub

Private Sub ColorMatch(ByRef sResponse As String, ByRef sString As String)
    Dim iStart As Integer
    
    iStart = InStr(1, sResponse, sString, vbBinaryCompare) - 1
    
    If (iStart >= 0) Then
        With frmMain.rtbResponses
            .SelStart = InStr(1, LCase$(sResponse), LCase$(sString), vbBinaryCompare) - 1
            .SelLength = Len(sString)
            .SelColor = &HC0C0&
        End With
    End If
End Sub

