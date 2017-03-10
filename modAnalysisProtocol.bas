Attribute VB_Name = "modAnalysisProtocol"
Option Explicit

Public Function GetProtocolName(ByRef sInput As String) As String
    Dim sProtocolName As String
    
    If (Len(sInput) > 14) Then
        sProtocolName = Mid$(sInput, 1, 4)
    End If
    
    GetProtocolName = sProtocolName
End Function

Public Function GetProtocolVersion(ByRef sInput As String) As String
    Dim sProtocolVersion As String
    
    If (Len(sInput) > 14) Then
        sProtocolVersion = Mid$(sInput, 6, 3)
    End If
    
    GetProtocolVersion = sProtocolVersion
End Function

