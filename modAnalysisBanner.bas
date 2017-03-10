Attribute VB_Name = "modAnalysisBanner"
Option Explicit

Public Function GetBanner(ByRef sInput As String) As String
    GetBanner = GetHeaderValue(sInput, "Server")
End Function

Public Function GetXPoweredBy(ByRef sInput As String) As String
    GetXPoweredBy = GetHeaderValue(sInput, "X-Powered-By")
End Function

Public Function PreFetchBanner(ByRef sRequest As String) As String
    Dim sBanner As String
    
    sBanner = GetHeaderValue(sRequest, "Server", True)
    
    If (LenB(sBanner)) Then
        PreFetchBanner = sBanner
    Else
        PreFetchBanner = "no banner available"
    End If
End Function
