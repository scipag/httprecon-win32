Attribute VB_Name = "modAnalysisAuthenticate"
Option Explicit

Public Function GetHtaccessRealm(ByRef sInput As String) As String
    Dim sAuthenticateLine As String
    Dim iRealmStart As Integer
    Dim iRealmLength As Integer
    
    Const sRealmString As String = "realm="""
    
    sAuthenticateLine = GetHeaderValue(sInput, "WWW-Authenticate")
    
    iRealmStart = InStr(1, sAuthenticateLine, sRealmString, vbBinaryCompare)
   
    If iRealmStart Then
        iRealmLength = iRealmStart + Len(sRealmString)
    
        GetHtaccessRealm = Mid$(sAuthenticateLine, iRealmLength, (InStr(iRealmStart + Len(sRealmString), sAuthenticateLine, ChrW$(34), vbBinaryCompare) - iRealmLength))
    End If
End Function

