Attribute VB_Name = "modAnalysisEtag"
Option Explicit

Public Function GetEtag(ByRef sInput As String) As String
    GetEtag = GetHeaderValue(LCase$(sInput), LCase$("ETag"))
End Function

Public Function GetEtagLength(ByRef sInput As String) As Integer
    GetEtagLength = Len(GetEtag(sInput))
End Function

Public Function GetEtagQuotes(ByRef sInput As String) As String
    If (InStrB(1, GetEtag(sInput), ChrW$(34), vbBinaryCompare)) Then
        GetEtagQuotes = ChrW$(34)
    End If
End Function
