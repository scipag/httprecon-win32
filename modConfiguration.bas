Attribute VB_Name = "modConfiguration"
Option Explicit

Public Const APP_NAME As String = "httprecon 7.3"
Public Const APP_WEBSITE_URL As String = "http://www.computec.ch/projekte/httprecon/"
Public Const APP_COPYRIGHT_OWNER As String = "httprecon project"

Public Const PROJECT_WEBSERVER As String = "www.computec.ch"
Public Const PROJECT_WEBPORT As Long = 80
Public Const PROJECT_WEBUPLOAD_FILE As String = "/projekte/httprecon/?s=upload"

Public Const PROJECT_WEBDB As String = "http://www.computec.ch/projekte/httprecon/?s=database"

Public app_configuration_filename As String

Public Sub LoadConfigFromFile(Optional ByRef sConfigurationFileName As String)
    On Error Resume Next
    
    Dim iFreeFile As Integer
    Dim sTempString As String
    
    Dim scan_targethostV As Boolean
    Dim scan_targetportV As Boolean
    Dim scan_targetsecureV As Boolean
    Dim app_hitpoints_minimumV As Boolean
    Dim app_hitpoints_maximumV As Boolean
    Dim scan_test_getexistingV As Boolean
    Dim scan_test_getnonexistingV As Boolean
    Dim scan_test_getlongV As Boolean
    Dim scan_test_headV As Boolean
    Dim scan_test_optionsV As Boolean
    Dim scan_test_wrongmethodV As Boolean
    Dim scan_test_nonexistingmethodV As Boolean
    Dim scan_test_wrongprotocolV As Boolean
    Dim scan_test_attackV As Boolean
    Dim req_timeout_connectV As Boolean
    Dim req_timeout_sendV As Boolean
    Dim req_timeout_receiveV As Boolean
    Dim req_protocol_legitimateV As Boolean
    Dim req_protocol_wrongV As Boolean
    Dim req_resource_availableV As Boolean
    Dim req_resource_notavailableV As Boolean
    Dim req_resource_attackV As Boolean
    Dim req_longrequest_lengthV As Boolean
    Dim req_longrequest_charV As Boolean
    Dim req_method_notallowedV As Boolean
    Dim req_method_notexistingV As Boolean
    Dim req_agent_nameV As Boolean
    Dim req_agent_noredirectV As Boolean
    Dim time_decimalsV As Boolean
        
    If (LenB(sConfigurationFileName)) Then
        app_configuration_filename = sConfigurationFileName
    Else
        app_configuration_filename = App.Path & "\config\default.cfg"
    End If
    
    If (Dir$(app_configuration_filename, 16) <> "") Then
        iFreeFile = FreeFile
        Open app_configuration_filename For Input As #iFreeFile
            Do While Not EOF(iFreeFile)
                Line Input #iFreeFile, sTempString
                
                If (Left$(sTempString, 1) <> "#") Then
                    If (InStrB(1, sTempString, "=", vbBinaryCompare)) Then
                        If (Left$(sTempString, 16) = "scan_targethost=") Then
                            scan_targethost = Mid$(sTempString, 17, Len(sTempString))
                            If (LenB(scan_targethost)) Then
                                scan_targethostV = True
                            End If
                        ElseIf (Left$(sTempString, 16) = "scan_targetport=") Then
                            scan_targetport = Val(Mid$(sTempString, 17, Len(sTempString)))
                            If (LenB(scan_targetport)) Then
                                scan_targetportV = True
                            End If
                        ElseIf (Left$(sTempString, 18) = "scan_targetsecure=") Then
                            If (Mid$(sTempString, 19, Len(sTempString)) = 1) Then
                                scan_targetsecure = 1
                            Else
                                scan_targetsecure = 0
                            End If
                            If (LenB(scan_targetsecure)) Then
                                scan_targetsecureV = True
                            End If
                        ElseIf (Left$(sTempString, 22) = "app_hitpoints_minimum=") Then
                            app_hitpoints_minimum = Val(Mid$(sTempString, 23, Len(sTempString)))
                            If (LenB(app_hitpoints_minimum)) Then
                                app_hitpoints_minimumV = True
                            End If
                        ElseIf (Left$(sTempString, 22) = "app_hitpoints_maximum=") Then
                            app_hitpoints_maximum = Val(Mid$(sTempString, 23, Len(sTempString)))
                            If (LenB(app_hitpoints_maximum)) Then
                                app_hitpoints_maximumV = True
                            End If
                        ElseIf (Left$(sTempString, 22) = "scan_test_getexisting=") Then
                            scan_test_getexisting = Val(Mid$(sTempString, 23, Len(sTempString)))
                            If (LenB(scan_test_getexisting)) Then
                                scan_test_getexistingV = True
                            End If
                        ElseIf (Left$(sTempString, 25) = "scan_test_getnonexisting=") Then
                            scan_test_getnonexisting = Val(Mid$(sTempString, 26, Len(sTempString)))
                            If (LenB(scan_test_getnonexisting)) Then
                                scan_test_getnonexistingV = True
                            End If
                        ElseIf (Left$(sTempString, 18) = "scan_test_getlong=") Then
                            scan_test_getlong = Val(Mid$(sTempString, 19, Len(sTempString)))
                            If (LenB(scan_test_getlong)) Then
                                scan_test_getlongV = True
                            End If
                        ElseIf (Left$(sTempString, 15) = "scan_test_head=") Then
                            scan_test_head = Val(Mid$(sTempString, 16, Len(sTempString)))
                            If (LenB(scan_test_head)) Then
                                scan_test_headV = True
                            End If
                        ElseIf (Left$(sTempString, 18) = "scan_test_options=") Then
                            scan_test_options = Val(Mid$(sTempString, 19, Len(sTempString)))
                            If (LenB(scan_test_options)) Then
                                scan_test_optionsV = True
                            End If
                        ElseIf (Left$(sTempString, 22) = "scan_test_wrongmethod=") Then
                            scan_test_wrongmethod = Val(Mid$(sTempString, 23, Len(sTempString)))
                            If (LenB(scan_test_wrongmethod)) Then
                                scan_test_wrongmethodV = True
                            End If
                        ElseIf (Left$(sTempString, 28) = "scan_test_nonexistingmethod=") Then
                            scan_test_nonexistingmethod = Val(Mid$(sTempString, 29, Len(sTempString)))
                            If (LenB(scan_test_nonexistingmethod)) Then
                                scan_test_nonexistingmethodV = True
                            End If
                        ElseIf (Left$(sTempString, 24) = "scan_test_wrongprotocol=") Then
                            scan_test_wrongprotocol = Val(Mid$(sTempString, 25, Len(sTempString)))
                            If (LenB(scan_test_wrongprotocol)) Then
                                scan_test_wrongprotocolV = True
                            End If
                        ElseIf (Left$(sTempString, 17) = "scan_test_attack=") Then
                            scan_test_attack = Val(Mid$(sTempString, 18, Len(sTempString)))
                            If (LenB(scan_test_attack)) Then
                                scan_test_attackV = True
                            End If
                        ElseIf (Left$(sTempString, 20) = "req_timeout_connect=") Then
                            req_timeout_connect = Val(Mid$(sTempString, 21, Len(sTempString)))
                            If (LenB(req_timeout_connect)) Then
                                req_timeout_connectV = True
                            End If
                        ElseIf (Left$(sTempString, 17) = "req_timeout_send=") Then
                            req_timeout_send = Val(Mid$(sTempString, 18, Len(sTempString)))
                            If (LenB(req_timeout_send)) Then
                                req_timeout_sendV = True
                            End If
                        ElseIf (Left$(sTempString, 20) = "req_timeout_receive=") Then
                            req_timeout_receive = Val(Mid$(sTempString, 21, Len(sTempString)))
                            If (LenB(req_timeout_receive)) Then
                                req_timeout_receiveV = True
                            End If
                        ElseIf (Left$(sTempString, 24) = "req_protocol_legitimate=") Then
                            req_protocol_legitimate = Mid$(sTempString, 25, Len(sTempString))
                            If (LenB(req_protocol_legitimate)) Then
                                req_protocol_legitimateV = True
                            End If
                        ElseIf (Left$(sTempString, 23) = "req_resource_available=") Then
                            req_resource_available = Mid$(sTempString, 24, Len(sTempString))
                            If (LenB(req_resource_available)) Then
                                req_resource_availableV = True
                            End If
                        ElseIf (Left$(sTempString, 26) = "req_resource_notavailable=") Then
                            req_resource_notavailable = Mid$(sTempString, 27, Len(sTempString))
                            If (LenB(req_resource_notavailable)) Then
                                req_resource_notavailableV = True
                            End If
                        ElseIf (Left$(sTempString, 19) = "req_resource_attack=") Then
                            req_resource_attack = Mid$(sTempString, 20, Len(sTempString))
                            If (LenB(req_resource_attack)) Then
                                req_resource_attackV = True
                            End If
                        ElseIf (Left$(sTempString, 23) = "req_longrequest_length=") Then
                            req_longrequest_length = Mid$(sTempString, 24, Len(sTempString))
                            If (LenB(req_longrequest_length)) Then
                                req_longrequest_lengthV = True
                            End If
                        ElseIf (Left$(sTempString, 21) = "req_longrequest_char=") Then
                            req_longrequest_char = Mid$(sTempString, 22, Len(sTempString))
                            If (LenB(req_longrequest_char)) Then
                                req_longrequest_charV = True
                            End If
                        ElseIf (Left$(sTempString, 22) = "req_method_notallowed=") Then
                            req_method_notallowed = Mid$(sTempString, 23, Len(sTempString))
                            If (LenB(req_method_notallowed)) Then
                                req_method_notallowedV = True
                            End If
                        ElseIf (Left$(sTempString, 23) = "req_method_notexisting=") Then
                            req_method_notexisting = Mid$(sTempString, 24, Len(sTempString))
                            If (LenB(req_method_notexisting)) Then
                                req_method_notexistingV = True
                            End If
                        ElseIf (Left$(sTempString, 19) = "req_protocol_wrong=") Then
                            req_protocol_wrong = Mid$(sTempString, 20, Len(sTempString))
                            If (LenB(req_protocol_wrong)) Then
                                req_protocol_wrongV = True
                            End If
                        ElseIf (Left$(sTempString, 15) = "req_agent_name=") Then
                            req_agent_name = Mid$(sTempString, 16, Len(sTempString))
                            If (LenB(req_agent_name)) Then
                                req_agent_nameV = True
                            End If
                        ElseIf (Left$(sTempString, 21) = "req_agent_noredirect=") Then
                            req_agent_noredirect = Val(Mid$(sTempString, 22, Len(sTempString)))
                            If (LenB(req_agent_noredirect)) Then
                                req_agent_noredirectV = True
                            End If
                        ElseIf (Left$(sTempString, 14) = "time_decimals=") Then
                            timing_decimals = Mid$(sTempString, 15, Len(sTempString))
                            If (LenB(timing_decimals)) Then
                                time_decimalsV = True
                            End If
                        End If
                    End If
                End If
            Loop
        Close
    End If

    If scan_targethostV = False Then
        scan_targethost = "127.0.0.1"
    End If
    
    If scan_targetportV = False Or scan_targetport < 1 Or scan_targetport > 65535 Then
        scan_targetport = 80
    End If

    If scan_targetsecureV = False Then
        scan_targetsecure = 0
    End If
    
    If app_hitpoints_minimumV = False Then
        app_hitpoints_minimum = 7
    End If
    
    If app_hitpoints_maximumV = False Then
        app_hitpoints_maximum = 14
    End If

    If scan_test_getexistingV = False Then
        scan_test_getexisting = 1
    End If

    If scan_test_getnonexistingV = False Then
        scan_test_getnonexisting = 1
    End If

    If scan_test_getlongV = False Then
        scan_test_getlong = 1
    End If

    If scan_test_headV = False Then
        scan_test_head = 1
    End If

    If scan_test_optionsV = False Then
        scan_test_options = 1
    End If

    If scan_test_wrongmethodV = False Then
        scan_test_wrongmethod = 1
    End If

    If scan_test_nonexistingmethodV = False Then
        scan_test_nonexistingmethod = 1
    End If

    If scan_test_wrongprotocolV = False Then
        scan_test_wrongprotocol = 1
    End If

    If scan_test_attackV = False Then
        scan_test_attack = 1
    End If

    If req_protocol_legitimateV = False Then
        req_protocol_legitimate = "HTTP/1.1"
    End If

    If req_timeout_connectV = False Or req_timeout_connect < 500 Then
        req_timeout_connect = Rand(4950, 5010)
    End If

    If req_timeout_sendV = False Or req_timeout_send < 500 Then
        req_timeout_send = Rand(4950, 5010)
    End If

    If req_timeout_receiveV = False Or req_timeout_receive < 500 Then
        req_timeout_receive = Rand(4950, 5010)
    End If

    If req_resource_availableV = False Then
        req_resource_available = "/"
    End If

    If req_resource_notavailableV = False Then
        req_resource_notavailable = "/" & ChrW$(Rand(97, 122)) & ChrW$(Rand(65, 90)) & Rand(0, 9) & ChrW$(Rand(97, 122)) & ChrW$(Rand(65, 90)) & Rand(0, 9) & ".html"
    End If

    If req_resource_attackV = False Then
        req_resource_attack = "/etc/passwd?format=" & String$(Rand(2, 4), "%") & "&xss=""><script>alert('xss');</script>&traversal=../../&sql=' OR " & Rand(0, 1) & ";"
    End If
    
    If req_longrequest_lengthV = False Then
        req_longrequest_length = 1024
    End If
    
    If req_longrequest_charV = False Then
        req_longrequest_char = ChrW$(Rand(97, 122))
    End If

    If req_method_notallowedV = False Then
        req_method_notallowed = "DELETE"
    End If

    If req_method_notexistingV = False Then
        req_method_notexisting = "TEST"
    End If

    If req_protocol_wrongV = False Then
        req_protocol_wrong = "HTTP/9.8"
        'req_protocol_wrong = "HTTP/9." & Rand(7, 9)
    End If

    If req_agent_nameV = False Then
        req_agent_name = "Mozilla/5.0 (Windows; U; Windows NT 5.2; en-US; rv:1.8.1.11) Gecko/20071127 Firefox/2.0.0." & Rand(1, 12)
        'req_agent_name = APP_NAME
    End If

    If req_agent_noredirectV = False Then
        req_agent_noredirect = 0
    End If

    If time_decimalsV = False Then
        timing_decimals = 3
    End If
End Sub

Public Sub WriteConfigurationToFile(ByRef sConfigurationFileName As String)
    Dim iFreeFile As Integer
    Dim cConfigContent As Concat
    
    Set cConfigContent = New Concat
    
    app_configuration_filename = sConfigurationFileName
    Call ChangeStatusBar("Save configuration to " & app_configuration_filename & ".")
    
    With cConfigContent
        .Concat _
        "# " & APP_NAME & " configuration file" & vbNewLine & _
        "# " & vbNewLine & _
        "#   Date   " & Date & vbNewLine & _
        "#   Time   " & Time & vbNewLine & _
        "#   File   " & app_configuration_filename & vbNewLine & _
        "#   User   " & GetLocalUsername & vbNewLine & _
        "#" & vbNewLine

        .Concat _
        "# Disclaimer: This config file is generated automatically by the software" & vbNewLine & _
        "# itself during runtime. Please do not manually edit these values unless you" & vbNewLine & _
        "# do know what you are doing. Manual changes work only if the application is" & vbNewLine & _
        "# not running." & vbNewLine & _
        "#" & vbNewLine & _
        "# See the online help, documentation and the official project web site at" & vbNewLine & _
        "# " & APP_WEBSITE_URL & " for more details." & vbNewLine & vbNewLine
    
        .Concat "# Scan" & vbNewLine
        .Concat "scan_targethost=" & scan_targethost & vbNewLine
        .Concat "scan_targetport=" & scan_targetport & vbNewLine
        .Concat "scan_targetsecure=" & scan_targetsecure & vbNewLine
        .Concat vbNewLine
        .Concat "# Statistics" & vbNewLine
        .Concat "app_hitpoints_minimum=" & app_hitpoints_minimum & vbNewLine
        .Concat "app_hitpoints_maximum=" & app_hitpoints_maximum & vbNewLine
        .Concat vbNewLine
        .Concat "# Tests" & vbNewLine
        .Concat "scan_test_getexisting=" & scan_test_getexisting & vbNewLine
        .Concat "scan_test_getnonexisting=" & scan_test_getnonexisting & vbNewLine
        .Concat "scan_test_getlong=" & scan_test_getlong & vbNewLine
        .Concat "scan_test_head=" & scan_test_head & vbNewLine
        .Concat "scan_test_options=" & scan_test_options & vbNewLine
        .Concat "scan_test_wrongmethod=" & scan_test_wrongmethod & vbNewLine
        .Concat "scan_test_nonexistingmethod=" & scan_test_nonexistingmethod & vbNewLine
        .Concat "scan_test_wrongprotocol=" & scan_test_wrongprotocol & vbNewLine
        .Concat "scan_test_attack=" & scan_test_attack & vbNewLine
        .Concat vbNewLine
        .Concat "# Requests" & vbNewLine
        .Concat "req_timeout_connect=" & req_timeout_connect & vbNewLine
        .Concat "req_timeout_send=" & req_timeout_send & vbNewLine
        .Concat "req_timeout_receive=" & req_timeout_receive & vbNewLine
        .Concat "req_protocol_legitimate=" & req_protocol_legitimate & vbNewLine
        .Concat "req_protocol_wrong=" & req_protocol_wrong & vbNewLine
        .Concat "req_resource_available=" & req_resource_available & vbNewLine
        .Concat "req_resource_notavailable=" & req_resource_notavailable & vbNewLine
        .Concat "req_resource_attack=" & req_resource_attack & vbNewLine
        .Concat "req_longrequest_length=" & req_longrequest_length & vbNewLine
        .Concat "req_longrequest_char=" & req_longrequest_char & vbNewLine
        .Concat "req_method_notallowed=" & req_method_notallowed & vbNewLine
        .Concat "req_method_notexisting=" & req_method_notexisting & vbNewLine
        .Concat "req_agent_name=" & req_agent_name & vbNewLine
        .Concat "req_agent_noredirect=" & req_agent_noredirect & vbNewLine
        .Concat vbNewLine
        .Concat "# Timing" & vbNewLine
        .Concat "time_decimals=" & timing_decimals & vbNewLine
    End With
    
    On Error Resume Next
    iFreeFile = FreeFile
    Open sConfigurationFileName For Output As #iFreeFile
        Print #iFreeFile, cConfigContent.Value
    Close
    
    Call ChangeStatusBarDone
End Sub

Public Function Rand(ByRef lLow As Long, ByRef lHigh As Long) As Long
    Rand = Int((lHigh - lLow + 1) * Rnd) + lLow
End Function
