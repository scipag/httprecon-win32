Attribute VB_Name = "modAnalysisOptions"
Option Explicit

Public Function GetOptionsAllowed(ByRef sInput As String) As String
    GetOptionsAllowed = Replace(GetHeaderValue(sInput, "Allow"), ", ", ",", , , vbBinaryCompare)
End Function

Public Function GetOptionsPublic(ByRef sInput As String) As String
    GetOptionsPublic = Replace(GetHeaderValue(sInput, "Public"), ", ", ",", , , vbBinaryCompare)
End Function

Public Function GetOptionsDelimiter(ByRef sInput As String) As String
    Dim sAllowedOptions As String
    
    sAllowedOptions = GetOptionsAllowed(sInput)
    
    If (InStrB(1, sAllowedOptions, ", ", vbBinaryCompare)) Then
        GetOptionsDelimiter = ", "
    ElseIf (InStrB(1, sAllowedOptions, ",", vbBinaryCompare)) Then
        GetOptionsDelimiter = ","
    End If
End Function
