Attribute VB_Name = "modTiming"
Option Explicit

Public timing_decimals As Integer

Public timing_start As Long

Public timing_attackrequest As Single
Public timing_delete As Single
Public timing_getexist As Single
Public timing_getlongrequest As Single
Public timing_get_nonexistent As Single
Public timing_head As Single
Public timing_options As Single
Public timing_testmethod As Single
Public timing_protocolversion As Single

Public Function CurrentTime() As Single
    CurrentTime = Timer
End Function

Public Function NormalizeTiming(ByRef sTiming As Single) As String
    NormalizeTiming = sTiming & " sec"
End Function

Public Function AverageTiming() As Single
    Dim sTotalTime As Single
    
    If (timing_attackrequest > 0) Then
        sTotalTime = sTotalTime + timing_attackrequest
    End If

    If (timing_delete > 0) Then
        sTotalTime = sTotalTime + timing_delete
    End If

    If (timing_getexist > 0) Then
        sTotalTime = sTotalTime + timing_getexist
    End If

    If (timing_getlongrequest > 0) Then
        sTotalTime = sTotalTime + timing_getlongrequest
    End If

    If (timing_get_nonexistent > 0) Then
        sTotalTime = sTotalTime + timing_get_nonexistent
    End If

    If (timing_options > 0) Then
        sTotalTime = sTotalTime + timing_options
    End If

    If (timing_testmethod > 0) Then
        sTotalTime = sTotalTime + timing_testmethod
    End If

    If (timing_protocolversion > 0) Then
        sTotalTime = sTotalTime + timing_protocolversion
    End If

    If (tests_count > 0) Then
        AverageTiming = (sTotalTime / tests_count)
    End If
End Function

Public Function MinimumTiming() As Single
    Dim sMinimumTime As Single
    
    If (timing_attackrequest > 0) Then
        sMinimumTime = timing_attackrequest
    End If

    If (timing_delete > 0 And timing_delete < sMinimumTime) Then
        sMinimumTime = timing_delete
    End If

    If (timing_getexist > 0 And timing_getexist < sMinimumTime) Then
        sMinimumTime = timing_getexist
    End If

    If (timing_getlongrequest > 0 And timing_getlongrequest < sMinimumTime) Then
        sMinimumTime = timing_getlongrequest
    End If

    If (timing_get_nonexistent > 0 And timing_get_nonexistent < sMinimumTime) Then
        sMinimumTime = timing_get_nonexistent
    End If

    If (timing_options > 0 And timing_options < sMinimumTime) Then
        sMinimumTime = timing_options
    End If

    If (timing_testmethod > 0 And timing_testmethod < sMinimumTime) Then
        sMinimumTime = timing_testmethod
    End If

    If (timing_protocolversion > 0 And timing_protocolversion < sMinimumTime) Then
        sMinimumTime = timing_protocolversion
    End If

    If (tests_count > 0) Then
        MinimumTiming = sMinimumTime
    End If
End Function

Public Function MaximumTiming() As Single
    Dim sMaximumTime As Single
    
    If (timing_attackrequest > 0) Then
        sMaximumTime = timing_attackrequest
    End If

    If (timing_delete > 0 And timing_delete > sMaximumTime) Then
        sMaximumTime = timing_delete
    End If

    If (timing_getexist > 0 And timing_getexist > sMaximumTime) Then
        sMaximumTime = timing_getexist
    End If

    If (timing_getlongrequest > 0 And timing_getlongrequest > sMaximumTime) Then
        sMaximumTime = timing_getlongrequest
    End If

    If (timing_get_nonexistent > 0 And timing_get_nonexistent > sMaximumTime) Then
        sMaximumTime = timing_get_nonexistent
    End If

    If (timing_options > 0 And timing_options > sMaximumTime) Then
        sMaximumTime = timing_options
    End If

    If (timing_testmethod > 0 And timing_testmethod > sMaximumTime) Then
        sMaximumTime = timing_testmethod
    End If

    If (timing_protocolversion > 0 And timing_protocolversion > sMaximumTime) Then
        sMaximumTime = timing_protocolversion
    End If

    If (tests_count > 0) Then
        MaximumTiming = sMaximumTime
    End If
End Function

