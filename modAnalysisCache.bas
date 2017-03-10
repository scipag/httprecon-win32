Attribute VB_Name = "modAnalysisCache"
Option Explicit

Public Function GetCacheControl(ByRef sInput As String) As String
    GetCacheControl = GetHeaderValue(sInput, "Cache-Control")
End Function

Public Function GetPragma(ByRef sInput As String) As String
    GetPragma = GetHeaderValue(sInput, "Pragma")
End Function

Public Function GetVaryOrder(ByRef sInput As String) As String
    GetVaryOrder = Replace(GetHeaderValue(sInput, "Vary"), ", ", ",", , , vbBinaryCompare)
End Function

Public Function GetVaryCapitalized(ByRef sInput As String) As String
    Dim sVaryElements As String
    
    sVaryElements = GetVaryOrder(sInput)

    If (LenB(sVaryElements)) Then
        If (sVaryElements = LCase$(sVaryElements)) Then
            GetVaryCapitalized = 0
        Else
            GetVaryCapitalized = 1
        End If
    End If
End Function

Public Function GetVaryDelimiter(ByRef sInput As String) As String
    Dim sVaryEntries As String
    
    sVaryEntries = GetVaryOrder(sInput)
    
    If (InStrB(1, sVaryEntries, ", ", vbBinaryCompare)) Then
        GetVaryDelimiter = ", "
    ElseIf (InStrB(1, sVaryEntries, ",", vbBinaryCompare)) Then
        GetVaryDelimiter = ","
    End If
End Function
