Attribute VB_Name = "modDatabaseHandling"
Option Explicit

Public Const APP_TESTNAME_ATTACKREQUEST As String = "attack_request"
Public Const APP_TESTNAME_DELETEEXISTING As String = "delete_existing"
Public Const APP_TESTNAME_GETEXISTING As String = "get_existing"
Public Const APP_TESTNAME_GETLONG As String = "get_long"
Public Const APP_TESTNAME_GETNONEXISTING As String = "get_nonexisting"
Public Const APP_TESTNAME_HEADEXISTING As String = "head_existing"
Public Const APP_TESTNAME_OPTIONS As String = "options"
Public Const APP_TESTNAME_WRONGMETHOD As String = "wrong_method"
Public Const APP_TESTNAME_WRONGVERSION As String = "wrong_version"

Public app_dir_databases As String
Public app_dir_attackrequest As String
Public app_dir_deleteexisting As String
Public app_dir_getexisting As String
Public app_dir_getlong As String
Public app_dir_getnonexisting As String
Public app_dir_headexisting As String
Public app_dir_options As String
Public app_dir_wrongmethod As String
Public app_dir_wrongversion As String

Public app_file_banner As String
Public app_file_xpoweredby As String
Public app_file_protocolname As String
Public app_file_protocolversion As String
Public app_file_statuscode As String
Public app_file_statustext As String
Public app_file_headerspace As String
Public app_file_headercapitalafterdash As String
Public app_file_headerorder As String
Public app_file_optionsallowed As String
Public app_file_optionspublic As String
Public app_file_optionsdelimiter As String
Public app_file_etaglength As String
Public app_file_etagquotes As String
Public app_file_contenttype As String
Public app_file_acceptrange As String
Public app_file_connection As String
Public app_file_cachecontrol As String
Public app_file_pragma As String
Public app_file_varyorder As String
Public app_file_varycapitalize As String
Public app_file_varydelimiter As String
Public app_file_htaccessrealm As String

Public Sub InitializeDirectories()
    Call ChangeStatusBar("Initialize Directories...")

    app_dir_databases = App.Path & "\database\"
    
    app_dir_attackrequest = app_dir_databases & APP_TESTNAME_ATTACKREQUEST & "\"
    app_dir_deleteexisting = app_dir_databases & APP_TESTNAME_DELETEEXISTING & "\"
    app_dir_getexisting = app_dir_databases & APP_TESTNAME_GETEXISTING & "\"
    app_dir_getlong = app_dir_databases & APP_TESTNAME_GETLONG & "\"
    app_dir_getnonexisting = app_dir_databases & APP_TESTNAME_GETNONEXISTING & "\"
    app_dir_headexisting = app_dir_databases & APP_TESTNAME_HEADEXISTING & "\"
    app_dir_options = app_dir_databases & APP_TESTNAME_OPTIONS & "\"
    app_dir_wrongmethod = app_dir_databases & APP_TESTNAME_WRONGMETHOD & "\"
    app_dir_wrongversion = app_dir_databases & APP_TESTNAME_WRONGVERSION & "\"
    
    Call ChangeStatusBarDone
End Sub

Public Sub InitializeFiles()
    Call ChangeStatusBar("Initialize Files...")
    
    Const sExtension As String = ".fdb"

    app_file_banner = "banner" & sExtension
    app_file_protocolname = "protocol-name" & sExtension
    app_file_protocolversion = "protocol-version" & sExtension
    app_file_statuscode = "statuscode" & sExtension
    app_file_statustext = "statustext" & sExtension
    app_file_headerspace = "header-space" & sExtension
    app_file_headercapitalafterdash = "header-capitalafterdash" & sExtension
    app_file_headerorder = "header-order" & sExtension
    app_file_optionsallowed = "options-allowed" & sExtension
    app_file_optionspublic = "options-public" & sExtension
    app_file_optionsdelimiter = "options-delimited" & sExtension
    app_file_etaglength = "etag-legth" & sExtension
    app_file_etagquotes = "etag-quotes" & sExtension
    app_file_contenttype = "content-type" & sExtension
    app_file_acceptrange = "accept-range" & sExtension
    app_file_connection = "connection" & sExtension
    app_file_cachecontrol = "cache-control" & sExtension
    app_file_pragma = "pragma" & sExtension
    app_file_varyorder = "vary-order" & sExtension
    app_file_varycapitalize = "vary-capitalize" & sExtension
    app_file_varydelimiter = "vary-delimiter" & sExtension
    app_file_xpoweredby = "x-powered-by" & sExtension
    app_file_htaccessrealm = "htaccess-realm" & sExtension
    
    Call ChangeStatusBarDone
End Sub

Public Sub SaveAllFingerprintsToDatabase(ByRef sImplementationName As String, ByRef sDatabasePath As String, ByRef sOriginalResponse As String)
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_banner, sImplementationName, GetBanner(sOriginalResponse))
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_xpoweredby, sImplementationName, GetXPoweredBy(sOriginalResponse))
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_protocolname, sImplementationName, GetProtocolName(sOriginalResponse))
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_protocolversion, sImplementationName, GetProtocolVersion(sOriginalResponse))
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_statuscode, sImplementationName, GetStatusCode(sOriginalResponse))
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_statustext, sImplementationName, GetStatusText(sOriginalResponse))
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_headerspace, sImplementationName, GetHeaderSpace(sOriginalResponse))
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_headercapitalafterdash, sImplementationName, GetHeaderCapitalAfterDash(sOriginalResponse))
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_headerorder, sImplementationName, GetHeaderOrder(sOriginalResponse, vbNullString))
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_optionsallowed, sImplementationName, GetOptionsAllowed(sOriginalResponse))
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_optionspublic, sImplementationName, GetOptionsPublic(sOriginalResponse))
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_optionsdelimiter, sImplementationName, GetOptionsDelimiter(sOriginalResponse))
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_etaglength, sImplementationName, GetEtagLength(sOriginalResponse))
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_etagquotes, sImplementationName, GetEtagQuotes(sOriginalResponse))
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_contenttype, sImplementationName, GetContentType(sOriginalResponse))
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_acceptrange, sImplementationName, GetAcceptRange(sOriginalResponse))
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_connection, sImplementationName, GetConnection(sOriginalResponse))
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_cachecontrol, sImplementationName, GetCacheControl(sOriginalResponse))
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_pragma, sImplementationName, GetPragma(sOriginalResponse))
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_varyorder, sImplementationName, GetVaryOrder(sOriginalResponse))
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_varycapitalize, sImplementationName, GetVaryCapitalized(sOriginalResponse))
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_varydelimiter, sImplementationName, GetVaryDelimiter(sOriginalResponse))
    Call SaveNewFingerprintToDatabase(sDatabasePath & app_file_htaccessrealm, sImplementationName, GetHtaccessRealm(sOriginalResponse))
End Sub

Public Sub SaveAllFingerprintsToAllDatabases(ByRef sImplementationName As String)
    Call ChangeStatusBar("Save All Fingerprints to Database...")

    Call SaveAllFingerprintsToDatabase(sImplementationName, app_dir_attackrequest, response_attackrequest)
    Call SaveAllFingerprintsToDatabase(sImplementationName, app_dir_deleteexisting, response_delete)
    Call SaveAllFingerprintsToDatabase(sImplementationName, app_dir_getexisting, response_getexist)
    Call SaveAllFingerprintsToDatabase(sImplementationName, app_dir_getlong, response_getlongrequest)
    Call SaveAllFingerprintsToDatabase(sImplementationName, app_dir_getnonexisting, response_get_nonexistent)
    Call SaveAllFingerprintsToDatabase(sImplementationName, app_dir_headexisting, response_head)
    Call SaveAllFingerprintsToDatabase(sImplementationName, app_dir_options, response_options)
    Call SaveAllFingerprintsToDatabase(sImplementationName, app_dir_wrongmethod, response_testmethod)
    Call SaveAllFingerprintsToDatabase(sImplementationName, app_dir_wrongversion, response_protocolversion)
    
    Call ChangeStatusBarDone
End Sub

Public Sub SaveNewFingerprintToDatabase(ByRef sFileName As String, ByRef sImplementationName As String, ByRef sFingerprintValue As String)
    Dim sNewEntry As String
    
    If (Dir$(sFileName, 16) <> "") Then
        If (LenB(sFingerprintValue)) Then
            If (LenB(sImplementationName)) Then
                sNewEntry = sImplementationName & ";" & sFingerprintValue
                If (IsAlreadyInDatabase(sFileName, sNewEntry) = False) Then
                    Open sFileName For Append As #1
                        Print #1, sNewEntry
                    Close
                End If
            End If
        End If
    End If
End Sub

Public Function IsAlreadyInDatabase(ByRef sDatabase As String, ByRef sNewEntry As String) As Boolean
    Dim sDatabaseContent As String
    
    sDatabaseContent = ReadFile(sDatabase)
    
    If (InStrB(1, sDatabaseContent, sNewEntry, vbBinaryCompare)) Then
        IsAlreadyInDatabase = True
    Else
        IsAlreadyInDatabase = False
    End If
End Function

Public Function ReadFile(ByRef sFileName As String) As String
    Dim sFileContent As String
    
    If (Dir$(sFileName, 16) <> "") Then
        Open sFileName For Input As #1
            sFileContent = Input(LOF(1), #1)
        Close
    End If
    
    ReadFile = sFileContent
End Function

Public Sub SaveFingerprints(ByRef sImplementationName As String)
    If (LenB(sImplementationName)) Then
        Call SaveAllFingerprintsToAllDatabases(sImplementationName)
        Call AnalyzeFingerprintsAndShowResult
    End If
End Sub
