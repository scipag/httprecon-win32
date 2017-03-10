Attribute VB_Name = "modLogging"
Option Explicit

Public Sub ChangeStatusBar(ByRef sMessage As String)
    frmMain.stbStatus.SimpleText = sMessage
End Sub

Public Sub ChangeStatusBarDone()
    frmMain.stbStatus.SimpleText = frmMain.stbStatus.SimpleText & " Done."
End Sub

Public Sub ChangeStatusBarReady()
    frmMain.stbStatus.SimpleText = "Ready."
End Sub
