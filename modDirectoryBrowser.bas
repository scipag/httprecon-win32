Attribute VB_Name = "modDirectoryBrowser"
Option Explicit

Private Declare Function SHBrowseForFolder Lib "Shell32" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)

Private Const MAX_PATH As Integer = 260
Private Const BIF_RETURNONLYFSDIRS As String = &H1&

Private Type BROWSEINFO
    hwndOwner       As Long
    pIDLRoot        As Long
    pszDisplayName  As Long
    lpszTitle       As String
    ulFlags         As Long
    lpfnCallback    As Long
    lParam          As Long
    iImage          As Long
End Type
    
Public Function BrowseForFolder(Optional vParent As Variant, Optional ByRef sTitle As String) As String
    Dim tBI As BROWSEINFO
    Dim lhWndParent As Long
    Dim lPIDL As Long
    Dim sPath As String
    
    If IsMissing(sTitle) Then sTitle = "Please choose a directory"
    If IsMissing(vParent) = False Then lhWndParent = vParent.hwnd
    
    tBI.hwndOwner = lhWndParent
    tBI.lpszTitle = sTitle
    tBI.ulFlags = BIF_RETURNONLYFSDIRS
    
    lPIDL = SHBrowseForFolder(tBI)
    
    If (lPIDL <> 0) Then
        sPath = Space$(MAX_PATH)
        SHGetPathFromIDList lPIDL, sPath
        
        sPath = Left$(sPath, InStr(sPath, ChrW$(0)) - 1)
        
        CoTaskMemFree lPIDL
    Else
        sPath = vbNullString
    End If
    
    BrowseForFolder = sPath
End Function

