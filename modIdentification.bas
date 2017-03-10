Attribute VB_Name = "modIdentification"
Option Explicit

Public Const APP_DATABASE_DELIMITER As String = ";"
Public Const APP_HITPOINTS_DELIMITER As String = ":"
Public Const APP_MATCHLIST_DELIMITER As String = ";"

Public app_hitpoints_minimum As Integer
Public app_hitpoints_maximum As Integer

Public scan_besthitcount As Integer
Public scan_besthitname As String

Public scan_time As String
Public scan_date As String
Public scan_targethost As String
Public scan_targetport As Long
Public scan_targetsecure As Integer

Public scan_test_folder As String

Public scan_test_getexisting As Integer
Public scan_test_getnonexisting As Integer
Public scan_test_getlong As Integer
Public scan_test_head As Integer
Public scan_test_options As Integer
Public scan_test_wrongmethod As Integer
Public scan_test_nonexistingmethod As Integer
Public scan_test_wrongprotocol As Integer
Public scan_test_attack As Integer

Public Function FindMatchInDatabase(ByRef sDatabase As String, ByRef sFingerprint As String) As String
    Dim sDatabaseContent() As String
    Dim sFingerprintInDatabase As String
    Dim iDatabaseEntries As Integer
    Dim iDelimiterPosition As Integer
    Dim i As Integer
    Dim cMatches As Concat
    
    Set cMatches = New Concat
    
    sDatabaseContent = Split(ReadFile(sDatabase), vbCrLf, , vbBinaryCompare)
    iDatabaseEntries = UBound(sDatabaseContent)
    
    For i = 0 To iDatabaseEntries
        If LenB(sFingerprint) Then
            If LenB(sDatabaseContent(i)) Then
                iDelimiterPosition = InStr(1, sDatabaseContent(i), APP_DATABASE_DELIMITER, vbBinaryCompare)
                sFingerprintInDatabase = Mid$(sDatabaseContent(i), iDelimiterPosition + 1, Len(sDatabaseContent(i)) - iDelimiterPosition)
                
                If (sFingerprintInDatabase = sFingerprint) Then
                    cMatches.Concat Mid$(sDatabaseContent(i), 1, InStr(1, sDatabaseContent(i), APP_DATABASE_DELIMITER, vbBinaryCompare) - 1)
                    
                    If (i < iDatabaseEntries) Then
                        cMatches.Concat APP_MATCHLIST_DELIMITER
                    End If
                End If
            End If
        End If
    Next i

    FindMatchInDatabase = cMatches.Value
End Function

Public Function GenerateMatchStatistics(ByRef sMatchList As String) As String
    Dim sMatches() As String
    Dim iMatchCount As Integer
    Dim i As Integer
    Dim cMatchStatistic As Concat
    
    Set cMatchStatistic = New Concat
    
    sMatches = Split(sMatchList, APP_MATCHLIST_DELIMITER, , vbBinaryCompare)
    Call RemoveDuplicatesFromArray(sMatches)
    iMatchCount = UBound(sMatches)
    
    For i = 0 To iMatchCount
        If (LenB(sMatches(i))) Then
            cMatchStatistic.Concat sMatches(i) & APP_HITPOINTS_DELIMITER & ArrayCountIf(sMatchList, sMatches(i)) & vbCrLf
            DoEvents
        End If
    Next i
    
    GenerateMatchStatistics = cMatchStatistic.Value
End Function

Public Sub RemoveDuplicatesFromArray(ByRef sArray() As String)
    Dim lLowBound As Long
    Dim lUpBound As Long
    Dim sTempArray() As String
    Dim lCurrent As Long
    Dim i As Long
    Dim j As Long
    
    lUpBound = UBound(sArray)
    
    If (lUpBound > 0) Then
        lLowBound = LBound(sArray)
        
        ReDim sTempArray(lLowBound To lUpBound)
        
        lCurrent = lLowBound
        sTempArray(lCurrent) = sArray(lLowBound)
        
        For i = lLowBound + 1 To lUpBound
            For j = lLowBound To lCurrent
                If LenB(sTempArray(j)) = LenB(sArray(i)) Then
                    If InStrB(1, sArray(i), sTempArray(j), vbBinaryCompare) = 1 Then
                        Exit For
                    End If
                End If
            Next j
            
            If j > lCurrent Then
                lCurrent = j
                sTempArray(lCurrent) = sArray(i)
            End If
        Next i
        
        ReDim Preserve sTempArray(lLowBound To lCurrent)
        sArray = sTempArray
    End If
End Sub

Public Function ArrayCountIf(ByRef sInput As String, ByRef sSearch As String) As Integer
    Dim sArray() As String
    Dim iArrayCount As Integer
    Dim i As Integer
    Dim iSum As Integer
    
    sArray = Split(sInput, APP_MATCHLIST_DELIMITER, , vbBinaryCompare)
    iArrayCount = UBound(sArray)
    
    For i = 0 To iArrayCount
        If (sArray(i) = sSearch) Then
            iSum = iSum + 1
        End If
    Next i
    
    ArrayCountIf = iSum
End Function

Public Sub AnnounceFingerprintMatches(ByRef lListView As ListView, ByRef sFullMatchList As String, ByRef iTestsCount As Integer)
    Dim sResultList As String
    Dim sResultArray() As String
    Dim i As Integer
    Dim iResultCount As Integer
    Dim lList As ListItem
    Dim sEntry() As String
    Dim iBestHitter As Integer
    Dim dBestMatch As Double
    
    Call ChangeStatusBar("Preparing Results... (this might take a few seconds)")

    sResultList = GenerateMatchStatistics(sFullMatchList)
    sResultArray = Split(sResultList, vbCrLf, , vbBinaryCompare)
    iResultCount = UBound(sResultArray)
    
    For i = 0 To iResultCount
        If (LenB(sResultArray(i))) Then
            sEntry = Split(sResultArray(i), APP_HITPOINTS_DELIMITER, , vbBinaryCompare)
            
            If (scan_besthitcount < sEntry(1)) Then
                scan_besthitname = sEntry(0)
                scan_besthitcount = sEntry(1)
            End If
        End If
    Next i
    If (scan_besthitcount < (app_hitpoints_minimum * iTestsCount)) Then
        iBestHitter = (app_hitpoints_minimum * iTestsCount)
    ElseIf (scan_besthitcount > (app_hitpoints_maximum * iTestsCount)) Then
        iBestHitter = (app_hitpoints_maximum * iTestsCount)
    Else
        iBestHitter = scan_besthitcount
    End If
    
    If (frmMain.tbsResults.SelectedItem.Index = 1) Then
        lListView.Visible = False
        lListView.ListItems.Clear
    End If
    For i = 0 To iResultCount
        If (LenB(sResultArray(i))) Then
            sEntry = Split(sResultArray(i), APP_HITPOINTS_DELIMITER, , vbBinaryCompare)
            
            dBestMatch = (100 / iBestHitter * sEntry(1))
            If (dBestMatch > 100) Then
                dBestMatch = 100
            End If
            
            Set lList = lListView.ListItems.Add(, , vbNullString, , GenerateHttpdIcon(sEntry(0)))
                lList.SubItems(1) = sEntry(0)
                lList.SubItems(2) = sEntry(1)
                lList.SubItems(3) = dBestMatch
        End If
    Next i
    Call ListViewSort(lListView, lListView.ColumnHeaders(3), 1)
    If (lListView.Name = frmMain.lsvResults.Name) Then
        If (frmMain.tbsResults.SelectedItem.Index = 1) Then
            lListView.Visible = True
        End If
        
        If (frmMain.tbsResults.SelectedItem.Index = 3) Then
            Call frmMain.UpdateReportPreview
        End If
    End If
    
    Call ChangeStatusBarReady
End Sub

Public Function GenerateHttpdIcon(ByVal sImplementation As String) As Integer
    sImplementation = LCase$(sImplementation)

    If (InStrB(1, sImplementation, "aol", vbBinaryCompare)) Then
        GenerateHttpdIcon = 1
    ElseIf (InStrB(1, sImplementation, "abyss", vbBinaryCompare)) Then
        GenerateHttpdIcon = 40
    ElseIf (InStrB(1, sImplementation, "allegro", vbBinaryCompare)) Then
        GenerateHttpdIcon = 91
    ElseIf (InStrB(1, sImplementation, "and-http", vbBinaryCompare)) Then
        GenerateHttpdIcon = 41
    ElseIf (InStrB(1, sImplementation, "anti-web", vbBinaryCompare)) Then
        GenerateHttpdIcon = 51
    ElseIf (InStrB(1, sImplementation, "apache", vbBinaryCompare)) Then
        GenerateHttpdIcon = 2
    ElseIf (InStrB(1, sImplementation, "araneida", vbBinaryCompare)) Then
        GenerateHttpdIcon = 92
    ElseIf (InStrB(1, sImplementation, "axis", vbBinaryCompare)) Then
        GenerateHttpdIcon = 59
    ElseIf (InStrB(1, sImplementation, "badblue", vbBinaryCompare)) Then
        GenerateHttpdIcon = 62
    ElseIf (InStrB(1, sImplementation, "barracuda", vbBinaryCompare)) Then
        GenerateHttpdIcon = 80
    ElseIf (InStrB(1, sImplementation, "basehttp", vbBinaryCompare)) Then
        GenerateHttpdIcon = 92
    ElseIf (InStrB(1, sImplementation, "boa", vbBinaryCompare)) Then
        GenerateHttpdIcon = 82
    ElseIf (InStrB(1, sImplementation, "bea", vbBinaryCompare)) Then
        GenerateHttpdIcon = 3
    ElseIf (InStrB(1, sImplementation, "belkin", vbBinaryCompare)) Then
        GenerateHttpdIcon = 81
    ElseIf (InStrB(1, sImplementation, "bozo", vbBinaryCompare)) Then
        GenerateHttpdIcon = 90
    ElseIf (InStrB(1, sImplementation, "caudium", vbBinaryCompare)) Then
        GenerateHttpdIcon = 31
    ElseIf (InStrB(1, sImplementation, "cherokee", vbBinaryCompare)) Then
        GenerateHttpdIcon = 33
    ElseIf (InStrB(1, sImplementation, "cisco", vbBinaryCompare)) Then
        GenerateHttpdIcon = 4
    ElseIf (InStrB(1, sImplementation, "cl-http", vbBinaryCompare)) Then
        GenerateHttpdIcon = 93
    ElseIf (InStrB(1, sImplementation, "compaq", vbBinaryCompare)) Then
        GenerateHttpdIcon = 5
    ElseIf (InStrB(1, sImplementation, "cougar", vbBinaryCompare)) Then
        GenerateHttpdIcon = 10
    ElseIf (InStrB(1, sImplementation, "dell", vbBinaryCompare)) Then
        GenerateHttpdIcon = 77
    ElseIf (InStrB(1, sImplementation, "divar", vbBinaryCompare)) Then
        GenerateHttpdIcon = 76
    ElseIf (InStrB(1, sImplementation, "dwhttpd", vbBinaryCompare)) Then
        GenerateHttpdIcon = 75
    ElseIf (InStrB(1, sImplementation, "emule", vbBinaryCompare)) Then
        GenerateHttpdIcon = 27
    ElseIf (InStrB(1, sImplementation, "firecat", vbBinaryCompare)) Then
        GenerateHttpdIcon = 42
    ElseIf (InStrB(1, sImplementation, "flexwatch", vbBinaryCompare)) Then
        GenerateHttpdIcon = 82
    ElseIf (InStrB(1, sImplementation, "fnord", vbBinaryCompare)) Then
        GenerateHttpdIcon = 84
    ElseIf (InStrB(1, sImplementation, "gatling", vbBinaryCompare)) Then
        GenerateHttpdIcon = 43
    ElseIf (InStrB(1, sImplementation, "globalscape", vbBinaryCompare)) Then
        GenerateHttpdIcon = 100
    ElseIf (InStrB(1, sImplementation, "google", vbBinaryCompare)) Then
        GenerateHttpdIcon = 34
    ElseIf (InStrB(1, sImplementation, "hp", vbBinaryCompare)) Then
        GenerateHttpdIcon = 7
    ElseIf (InStrB(1, sImplementation, "hiawatha", vbBinaryCompare)) Then
        GenerateHttpdIcon = 44
    ElseIf (InStrB(1, sImplementation, "httpi", vbBinaryCompare)) Then
        GenerateHttpdIcon = 94
    ElseIf (InStrB(1, sImplementation, "ibm", vbBinaryCompare)) Then
        GenerateHttpdIcon = 8
    ElseIf (InStrB(1, sImplementation, "icewarp", vbBinaryCompare)) Then
        GenerateHttpdIcon = 50
    ElseIf (InStrB(1, sImplementation, "indy", vbBinaryCompare)) Then
        GenerateHttpdIcon = 85
    ElseIf (InStrB(1, sImplementation, "iis 4", vbBinaryCompare)) Then
        GenerateHttpdIcon = 9
    ElseIf (InStrB(1, sImplementation, "iis 5", vbBinaryCompare)) Then
        GenerateHttpdIcon = 9
    ElseIf (InStrB(1, sImplementation, "iis ", vbBinaryCompare)) Then
        GenerateHttpdIcon = 10
    ElseIf (InStrB(1, sImplementation, "jana", vbBinaryCompare)) Then
        GenerateHttpdIcon = 11
    ElseIf (InStrB(1, sImplementation, "jetty", vbBinaryCompare)) Then
        GenerateHttpdIcon = 37
    ElseIf (InStrB(1, sImplementation, "jigsaw", vbBinaryCompare)) Then
        GenerateHttpdIcon = 55
    ElseIf (InStrB(1, sImplementation, "lancom", vbBinaryCompare)) Then
        GenerateHttpdIcon = 65
    ElseIf (InStrB(1, sImplementation, "konica", vbBinaryCompare)) Then
        GenerateHttpdIcon = 66
    ElseIf (InStrB(1, sImplementation, "lexmark", vbBinaryCompare)) Then
        GenerateHttpdIcon = 79
    ElseIf (InStrB(1, sImplementation, "lighttpd", vbBinaryCompare)) Then
        GenerateHttpdIcon = 29
    ElseIf (InStrB(1, sImplementation, "linksys", vbBinaryCompare)) Then
        GenerateHttpdIcon = 12
    ElseIf (InStrB(1, sImplementation, "listmanager", vbBinaryCompare)) Then
        GenerateHttpdIcon = 58
    ElseIf (InStrB(1, sImplementation, "litespeed", vbBinaryCompare)) Then
        GenerateHttpdIcon = 49
    ElseIf (InStrB(1, sImplementation, "lotus", vbBinaryCompare)) Then
        GenerateHttpdIcon = 6
    ElseIf (InStrB(1, sImplementation, "mikrotik", vbBinaryCompare)) Then
        GenerateHttpdIcon = 13
    ElseIf (InStrB(1, sImplementation, "mongrel", vbBinaryCompare)) Then
        GenerateHttpdIcon = 86
    ElseIf (InStrB(1, sImplementation, "net2phone", vbBinaryCompare)) Then
        GenerateHttpdIcon = 64
    ElseIf (InStrB(1, sImplementation, "netgear", vbBinaryCompare)) Then
        GenerateHttpdIcon = 35
    ElseIf (InStrB(1, sImplementation, "netopia", vbBinaryCompare)) Then
        GenerateHttpdIcon = 63
    ElseIf (InStrB(1, sImplementation, "netscape", vbBinaryCompare)) Then
        GenerateHttpdIcon = 14
    ElseIf (InStrB(1, sImplementation, "nginx", vbBinaryCompare)) Then
        GenerateHttpdIcon = 38
    ElseIf (InStrB(1, sImplementation, "novell", vbBinaryCompare)) Then
        GenerateHttpdIcon = 15
    ElseIf (InStrB(1, sImplementation, "omnihttpd", vbBinaryCompare)) Then
        GenerateHttpdIcon = 73
    ElseIf (InStrB(1, sImplementation, "oracle", vbBinaryCompare)) Then
        GenerateHttpdIcon = 39
    ElseIf (InStrB(1, sImplementation, "orion", vbBinaryCompare)) Then
        GenerateHttpdIcon = 96
    ElseIf (InStrB(1, sImplementation, "osu", vbBinaryCompare)) Then
        GenerateHttpdIcon = 95
    ElseIf (InStrB(1, sImplementation, "packetshaper", vbBinaryCompare)) Then
        GenerateHttpdIcon = 87
    ElseIf (InStrB(1, sImplementation, "philips", vbBinaryCompare)) Then
        GenerateHttpdIcon = 78
    ElseIf (InStrB(1, sImplementation, "publicfile", vbBinaryCompare)) Then
        GenerateHttpdIcon = 88
    ElseIf (InStrB(1, sImplementation, "qnap", vbBinaryCompare)) Then
        GenerateHttpdIcon = 71
    ElseIf (InStrB(1, sImplementation, "resin", vbBinaryCompare)) Then
        GenerateHttpdIcon = 56
    ElseIf (InStrB(1, sImplementation, "ricoh", vbBinaryCompare)) Then
        GenerateHttpdIcon = 72
    ElseIf (InStrB(1, sImplementation, "roxen", vbBinaryCompare)) Then
        GenerateHttpdIcon = 45
    ElseIf (InStrB(1, sImplementation, "smc", vbBinaryCompare)) Then
        GenerateHttpdIcon = 16
    ElseIf (InStrB(1, sImplementation, "snap", vbBinaryCompare)) Then
        GenerateHttpdIcon = 17
    ElseIf (InStrB(1, sImplementation, "sonicwall", vbBinaryCompare)) Then
        GenerateHttpdIcon = 52
    ElseIf (InStrB(1, sImplementation, "sony", vbBinaryCompare)) Then
        GenerateHttpdIcon = 61
    ElseIf (InStrB(1, sImplementation, "squid", vbBinaryCompare)) Then
        GenerateHttpdIcon = 30
    ElseIf (InStrB(1, sImplementation, "stweb", vbBinaryCompare)) Then
        GenerateHttpdIcon = 97
    ElseIf (InStrB(1, sImplementation, "sun", vbBinaryCompare)) Then
        GenerateHttpdIcon = 18
    ElseIf (InStrB(1, sImplementation, "swat", vbBinaryCompare)) Then
        GenerateHttpdIcon = 28
    ElseIf (InStrB(1, sImplementation, "symantec", vbBinaryCompare)) Then
        GenerateHttpdIcon = 74
    ElseIf (InStrB(1, sImplementation, "tandberg", vbBinaryCompare)) Then
        GenerateHttpdIcon = 89
    ElseIf (InStrB(1, sImplementation, "tcl", vbBinaryCompare)) Then
        GenerateHttpdIcon = 58
    ElseIf (InStrB(1, sImplementation, "thttpd", vbBinaryCompare)) Then
        GenerateHttpdIcon = 19
    ElseIf (InStrB(1, sImplementation, "tomcat", vbBinaryCompare)) Then
        GenerateHttpdIcon = 20
    ElseIf (InStrB(1, sImplementation, "ubicom", vbBinaryCompare)) Then
        GenerateHttpdIcon = 22
    ElseIf (InStrB(1, sImplementation, "userland", vbBinaryCompare)) Then
        GenerateHttpdIcon = 54
    ElseIf (InStrB(1, sImplementation, "virtuoso", vbBinaryCompare)) Then
        GenerateHttpdIcon = 46
    ElseIf (InStrB(1, sImplementation, "vnc", vbBinaryCompare)) Then
        GenerateHttpdIcon = 23
    ElseIf (InStrB(1, sImplementation, "vqserver", vbBinaryCompare)) Then
        GenerateHttpdIcon = 98
    ElseIf (InStrB(1, sImplementation, "vs", vbBinaryCompare)) Then
        GenerateHttpdIcon = 99
    ElseIf (InStrB(1, sImplementation, "wdaemon", vbBinaryCompare)) Then
        GenerateHttpdIcon = 53
    ElseIf (InStrB(1, sImplementation, "webcamxp", vbBinaryCompare)) Then
        GenerateHttpdIcon = 69
    ElseIf (InStrB(1, sImplementation, "wn", vbBinaryCompare)) Then
        GenerateHttpdIcon = 57
    ElseIf (InStrB(1, sImplementation, "webrick", vbBinaryCompare)) Then
        GenerateHttpdIcon = 47
    ElseIf (InStrB(1, sImplementation, "xitami", vbBinaryCompare)) Then
        GenerateHttpdIcon = 32
    ElseIf (InStrB(1, sImplementation, "xserver", vbBinaryCompare)) Then
        GenerateHttpdIcon = 24
    ElseIf (InStrB(1, sImplementation, "yaws", vbBinaryCompare)) Then
        GenerateHttpdIcon = 48
    ElseIf (InStrB(1, sImplementation, "zeus", vbBinaryCompare)) Then
        GenerateHttpdIcon = 25
    ElseIf (InStrB(1, sImplementation, "zope", vbBinaryCompare)) Then
        GenerateHttpdIcon = 26
    ElseIf (InStrB(1, sImplementation, "zyxel", vbBinaryCompare)) Then
        GenerateHttpdIcon = 60
    ElseIf (InStrB(1, sImplementation, "4d", vbBinaryCompare)) Then
        GenerateHttpdIcon = 36
    
' Operating systems collector
    ElseIf (InStrB(1, sImplementation, "bsd", vbBinaryCompare)) Then
        GenerateHttpdIcon = 67
    ElseIf (InStrB(1, sImplementation, "debian", vbBinaryCompare)) Then
        GenerateHttpdIcon = 68
    ElseIf (InStrB(1, sImplementation, "suse", vbBinaryCompare)) Then
        GenerateHttpdIcon = 70
    ElseIf (InStrB(1, sImplementation, "linux", vbBinaryCompare)) Then
        GenerateHttpdIcon = 21
    ElseIf (InStrB(1, sImplementation, "windows", vbBinaryCompare)) Then
        GenerateHttpdIcon = 10
    ElseIf (InStrB(1, sImplementation, "microsoft", vbBinaryCompare)) Then
        GenerateHttpdIcon = 10
    Else
        GenerateHttpdIcon = 101
    End If
End Function

Public Function IdentifyGlobalFingerprint(ByRef sFingerprintDirectory As String, ByRef sOriginalResponse As String) As String
    If (LenB(sOriginalResponse)) Then
        Dim cFullMatchList As Concat
    
        Set cFullMatchList = New Concat
        
        Call AddTestCount(sOriginalResponse)
        
        With cFullMatchList
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_banner, GetBanner(sOriginalResponse))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_xpoweredby, GetXPoweredBy(sOriginalResponse))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_protocolname, GetProtocolName(sOriginalResponse))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_protocolversion, GetProtocolVersion(sOriginalResponse))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_statuscode, GetStatusCode(sOriginalResponse))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_statustext, GetStatusText(sOriginalResponse))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_headerspace, GetHeaderSpace(sOriginalResponse))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_headercapitalafterdash, GetHeaderCapitalAfterDash(sOriginalResponse))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_headerorder, GetHeaderOrder(sOriginalResponse, vbNullString))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_headerorder, GetHeaderOrder(sOriginalResponse, "X-|Set-Cookie"))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_optionsallowed, GetOptionsAllowed(sOriginalResponse))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_optionspublic, GetOptionsPublic(sOriginalResponse))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_optionsdelimiter, GetOptionsDelimiter(sOriginalResponse))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_etaglength, GetEtagLength(sOriginalResponse))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_etagquotes, GetEtagQuotes(sOriginalResponse))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_contenttype, GetContentType(sOriginalResponse))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_acceptrange, GetAcceptRange(sOriginalResponse))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_connection, GetConnection(sOriginalResponse))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_cachecontrol, GetCacheControl(sOriginalResponse))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_pragma, GetPragma(sOriginalResponse))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_varyorder, GetVaryOrder(sOriginalResponse))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_varycapitalize, GetVaryCapitalized(sOriginalResponse))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_varydelimiter, GetVaryDelimiter(sOriginalResponse))
            .Concat FindMatchInDatabase(sFingerprintDirectory & app_file_htaccessrealm, GetHtaccessRealm(sOriginalResponse))
        End With

        IdentifyGlobalFingerprint = cFullMatchList.Value
    End If
End Function

Public Sub ServerAnalysis()
    Call frmMain.DisableElements
    Call ChangeStatusBar("Starting Analysis...")

    scan_time = Time
    scan_date = Date
    scan_besthitcount = 0
    scan_besthitname = vbNullString
    
    tests_warning = False
    
    With frmMain
        scan_targethost = .txtTargetHost.Text
        scan_targetport = .cboTargetPort.Text
        .Caption = .UpdateCaption
        .fraTarget.Caption = "Target (unknown)"
    End With
    
    Call WriteConfigurationToFile(app_configuration_filename)
    DoEvents

    If (RunTestRequests(scan_targethost, scan_targetport, scan_targetsecure) = True) Then
        Call AnalyzeFingerprintsAndShowResult
    Else
        Call ChangeStatusBar("Target " & scan_targethost & ":" & scan_targetport & " is not a web server. Aborting.")
'        MsgBox "Target " & scan_targethost & ":" & scan_targetport & " is not a web server." & vbCrLf & _
'            "Please check your settings.", vbExclamation + vbOKOnly, "No web server found"
    End If

    Call frmMain.EnableElements
End Sub

Public Sub AnalyzeFingerprintsAndShowResult()
    Dim cFullIdentifyList As Concat

    Set cFullIdentifyList = New Concat
    
    tests_count = 0
    
    With cFullIdentifyList
        .Concat IdentifyGlobalFingerprint(app_dir_attackrequest, response_attackrequest)
        .Concat IdentifyGlobalFingerprint(app_dir_deleteexisting, response_delete)
        .Concat IdentifyGlobalFingerprint(app_dir_getexisting, response_getexist)
        .Concat IdentifyGlobalFingerprint(app_dir_getlong, response_getlongrequest)
        .Concat IdentifyGlobalFingerprint(app_dir_getnonexisting, response_get_nonexistent)
        .Concat IdentifyGlobalFingerprint(app_dir_headexisting, response_head)
        .Concat IdentifyGlobalFingerprint(app_dir_options, response_options)
        .Concat IdentifyGlobalFingerprint(app_dir_wrongmethod, response_testmethod)
        .Concat IdentifyGlobalFingerprint(app_dir_wrongversion, response_protocolversion)
    End With
    
    Call FillResponses
    Call AnnounceFingerprintMatches(frmMain.lsvResults, cFullIdentifyList.Value, tests_count)
    
    With frmMain
        .fraTarget.Caption = "Target (" & scan_besthitname & ")"
        .tbsResults.Tabs(1).Caption = "Matchlist (" & .lsvResults.ListItems.Count & " Implementations)"
        .mnuFileSaveAsScanItem.Enabled = True
        .mnuFingerprintingSaveFingerprintItem.Enabled = True
        .mnuFingerprintingReanalyzeItem.Enabled = True
        .mnuReportingGenerateReportItem.Enabled = True
        
        Call .EnableElements
    End With
End Sub

Public Sub FillResponses()
    Dim iIndex As Integer
    
    iIndex = frmMain.tbsViews.SelectedItem.Index
    
    If (iIndex = 1) Then
        Call ShowResponseData(response_getexist, timing_getexist)
        scan_test_folder = app_dir_getexisting
    ElseIf (iIndex = 2) Then
        Call ShowResponseData(response_getlongrequest, timing_getlongrequest)
        scan_test_folder = app_dir_getlong
    ElseIf (iIndex = 3) Then
        Call ShowResponseData(response_get_nonexistent, timing_get_nonexistent)
        scan_test_folder = app_dir_getnonexisting
    ElseIf (iIndex = 4) Then
        Call ShowResponseData(response_protocolversion, timing_protocolversion)
        scan_test_folder = app_dir_wrongversion
    ElseIf (iIndex = 5) Then
        Call ShowResponseData(response_head, timing_head)
        scan_test_folder = app_dir_headexisting
    ElseIf (iIndex = 6) Then
        Call ShowResponseData(response_options, timing_options)
        scan_test_folder = app_dir_options
    ElseIf (iIndex = 7) Then
        Call ShowResponseData(response_delete, timing_delete)
        scan_test_folder = app_dir_deleteexisting
    ElseIf (iIndex = 8) Then
        Call ShowResponseData(response_testmethod, timing_testmethod)
        scan_test_folder = app_dir_wrongmethod
    ElseIf (iIndex = 9) Then
        Call ShowResponseData(response_attackrequest, timing_attackrequest)
        scan_test_folder = app_dir_attackrequest
    End If
    
    Call ResetResponseHighlight
End Sub

Public Sub ShowResponseData(ByRef sResponse As String, ByRef sTiming As Single)
    With frmMain
        .rtbResponses.Text = sResponse
        .rtbResponses.ToolTipText = "Length: " & Len(sResponse) & " bytes / " & _
            "Timing: " & NormalizeTiming(sTiming)
    
        .txtFingerprint.Text = GenerateFingerprintDetails(sResponse)
    End With
End Sub

Public Sub AnalyzeTestFingerprintsAndShowResult()
    With frmMain
        .lsvResultsForTest.ListItems.Clear
        Call AnnounceFingerprintMatches(.lsvResultsForTest, IdentifyGlobalFingerprint(scan_test_folder, .rtbResponses.Text), 1)
    End With
End Sub

Public Function GenerateFingerprintDetails(ByRef sOriginalResponse As String) As String
    Dim cFingerprintDetails As Concat

    If (LenB(sOriginalResponse)) Then
        Set cFingerprintDetails = New Concat
        
        With cFingerprintDetails
            .Concat "Protocol Name" & vbTab & vbTab & GetProtocolName(sOriginalResponse) & vbCrLf
            .Concat "Protocol Version" & vbTab & GetProtocolVersion(sOriginalResponse) & vbCrLf
            .Concat "Statuscode" & vbTab & vbTab & GetStatusCode(sOriginalResponse) & vbCrLf
            .Concat "Statustext" & vbTab & vbTab & GetStatusText(sOriginalResponse) & vbCrLf
            .Concat "Banner" & vbTab & vbTab & vbTab & GetBanner(sOriginalResponse) & vbCrLf
            .Concat "X-Powered-By" & vbTab & vbTab & GetXPoweredBy(sOriginalResponse) & vbCrLf
            .Concat "Header Spaces" & vbTab & vbTab & GetHeaderSpace(sOriginalResponse) & vbCrLf
            .Concat "Capital after Dash" & vbTab & GetHeaderCapitalAfterDash(sOriginalResponse) & vbCrLf
            .Concat "Header-Order Full" & vbTab & GetHeaderOrder(sOriginalResponse, "") & vbCrLf
            .Concat "Header-Order Limit" & vbTab & GetHeaderOrder(sOriginalResponse, "X-|Set-Cookie") & vbCrLf
            .Concat "Options-Allowed" & vbTab & vbTab & GetOptionsAllowed(sOriginalResponse) & vbCrLf
            .Concat "Options-Public" & vbTab & vbTab & GetOptionsPublic(sOriginalResponse) & vbCrLf
            .Concat "Options-Delimiter" & vbTab & GetOptionsDelimiter(sOriginalResponse) & vbCrLf
            .Concat "ETag" & vbTab & vbTab & vbTab & GetEtag(sOriginalResponse) & vbCrLf
            .Concat "ETag-Length" & vbTab & vbTab & GetEtagLength(sOriginalResponse) & vbCrLf
            .Concat "ETag-Quotes" & vbTab & vbTab & GetEtagQuotes(sOriginalResponse) & vbCrLf
            .Concat "Content-Type" & vbTab & vbTab & GetContentType(sOriginalResponse) & vbCrLf
            .Concat "Accept-Range" & vbTab & vbTab & GetAcceptRange(sOriginalResponse) & vbCrLf
            .Concat "Connection" & vbTab & vbTab & GetConnection(sOriginalResponse) & vbCrLf
            .Concat "Cache-Control" & vbTab & vbTab & GetCacheControl(sOriginalResponse) & vbCrLf
            .Concat "Pragma" & vbTab & vbTab & vbTab & GetPragma(sOriginalResponse) & vbCrLf
            .Concat "Vary-Order" & vbTab & vbTab & GetVaryOrder(sOriginalResponse) & vbCrLf
            .Concat "Vary-Capitalized" & vbTab & GetVaryCapitalized(sOriginalResponse) & vbCrLf
            .Concat "Vary-Delimiter" & vbTab & vbTab & GetVaryDelimiter(sOriginalResponse) & vbCrLf
            .Concat "htaccess-Realm" & vbTab & vbTab & GetHtaccessRealm(sOriginalResponse) & vbCrLf
        End With
       
        GenerateFingerprintDetails = cFingerprintDetails.Value
    End If
End Function
