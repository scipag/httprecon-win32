Attribute VB_Name = "modReporting"
Option Explicit

Declare Function GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function GetLocalUsername() As String
    Dim sTemp As String
    
    sTemp = String(255, 0)
    GetUserName sTemp, 255
    GetLocalUsername = Left$(sTemp, InStr(sTemp, ChrW$(0)) - 1)
End Function


