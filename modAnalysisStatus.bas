Attribute VB_Name = "modAnalysisStatus"
Option Explicit

Public Function GetStatusCode(ByRef sInput As String) As String
    Dim sStatusCode As String

    If (Len(sInput) > 14) Then
        sStatusCode = Mid$(sInput, 10, 3)
    End If
    
    If (IsNumeric(sStatusCode)) Then
        GetStatusCode = sStatusCode
    Else
        GetStatusCode = 0
    End If
End Function

Public Function GetStatusText(ByRef sInput As String) As String
    On Error Resume Next 'Workaround for some crash bug
    
    Dim sStatusText As String
    Dim iLineEnd As Integer

    Const iLineStart As Integer = "14"
    
    If (Len(sInput) > iLineStart) Then
        iLineEnd = InStr(iLineStart, sInput, vbCrLf, vbBinaryCompare) - iLineStart

        If (iLineEnd > iLineStart) Then
            sStatusText = Mid$(sInput, iLineStart, iLineEnd)
        End If
    End If
    
    GetStatusText = sStatusText
End Function
