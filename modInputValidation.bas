Attribute VB_Name = "modInputValidation"
Option Explicit

Public Function AllowIntegersOnly(ByRef lInput As Long, ByRef lMinimum As Long, ByRef lMaximum As Long, ByRef lDefault As Long)
    If LenB(lInput) = 0 Or lInput = 0 Then
        AllowIntegersOnly = lDefault
    Else
        If lInput < lMinimum Then
            AllowIntegersOnly = lMinimum
        ElseIf lInput > lMaximum Then
            AllowIntegersOnly = lMaximum
        Else
            AllowIntegersOnly = lInput
        End If
    End If
End Function

Public Function PreventEmptyInput(ByRef sInput As String, ByRef sDefault As String) As String
    If (LenB(sInput)) Then
        PreventEmptyInput = sInput
    Else
        PreventEmptyInput = sDefault
    End If
End Function
