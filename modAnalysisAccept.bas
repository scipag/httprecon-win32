Attribute VB_Name = "modAnalysisAccept"
Option Explicit

Public Function GetContentType(ByRef sInput As String) As String
    GetContentType = GetHeaderValue(sInput, "Content-Type")
End Function

Public Function GetAcceptRange(ByRef sInput As String) As String
    GetAcceptRange = GetHeaderValue(sInput, "Accept-Ranges")
End Function


