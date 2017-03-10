VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Generation"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4470
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cdgReportSaveAs 
      Left            =   120
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "HTML Report (*.html)|*.html"
      DialogTitle     =   "Save Report As"
      FileName        =   "127.0.0.1-80.html"
      Filter          =   "HTML Report (*.html)|*.html|XML Report (*.xml)|*.xml|TXT Report (*.txt)|*.txt|CSV Report (*.csv)|*.csv"
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Hitlist Size"
      Height          =   735
      Index           =   1
      Left            =   2280
      TabIndex        =   15
      Top             =   120
      Width           =   2055
      Begin VB.TextBox txtHitlistSize 
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "20"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Items"
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Format"
      Height          =   1695
      Index           =   2
      Left            =   2280
      TabIndex        =   14
      Top             =   960
      Width           =   2055
      Begin VB.OptionButton optCsv 
         Caption         =   "CS&V"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton optXML 
         Caption         =   "&XML"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton optTXT 
         Caption         =   "T&XT"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton optHTML 
         Caption         =   "&HTML"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Details"
      Height          =   2535
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   2055
      Begin VB.CheckBox chkResponses 
         Caption         =   "&Responses"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "All gathered responses"
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkPreamble 
         Caption         =   "&Preamble"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Additional information about the scan"
         Top             =   360
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkContents 
         Caption         =   "Con&tents"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Table of Contents"
         Top             =   720
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkDetails 
         Caption         =   "&Details"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Fingerprint details"
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkMatches 
         Caption         =   "&Matches"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "A list of the best matches"
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkSummary 
         Caption         =   "Summar&y"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Text summarizing the results"
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   2280
      Picture         =   "frmReport.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cancel Report Generation"
      Top             =   2760
      Width           =   1212
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   615
      Left            =   960
      Picture         =   "frmReport.frx":0772
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Save Report"
      Top             =   2760
      Width           =   1212
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim sFileName As String
    Dim sOverride As String
    
    Call DisableElements
    
    If (Dir$(App.Path & "\reports", 16) <> "") Then
        cdgReportSaveAs.InitDir = App.Path & "\reports"
    Else
        cdgReportSaveAs.InitDir = App.Path
    End If

    If (optHTML.Value = True) Then
        cdgReportSaveAs.FileName = StringToFileName(scan_targethost & ":" & scan_targetport) & ".html"
        cdgReportSaveAs.FilterIndex = 1
    ElseIf (optXML.Value = True) Then
        cdgReportSaveAs.FileName = StringToFileName(scan_targethost & ":" & scan_targetport) & ".xml"
        cdgReportSaveAs.FilterIndex = 2
    ElseIf (optTXT.Value = True) Then
        cdgReportSaveAs.FileName = StringToFileName(scan_targethost & ":" & scan_targetport) & ".txt"
        cdgReportSaveAs.FilterIndex = 3
    Else
        cdgReportSaveAs.FileName = StringToFileName(scan_targethost & ":" & scan_targetport) & ".csv"
        cdgReportSaveAs.FilterIndex = 4
    End If
    
    On Error GoTo Cancel
    cdgReportSaveAs.ShowSave
    sFileName = cdgReportSaveAs.FileName
    
    If (LenB(sFileName)) Then
        If (Dir$(sFileName, 16) <> "") Then
            sOverride = MsgBox(sFileName & " already exists." & vbCrLf & "Do you want to replace it?", _
                vbExclamation + vbYesNo, "Report Save As")
        Else
            sOverride = 6
        End If

        If (sOverride = 6) Then
            Open sFileName For Output As #1
                If (Me.optHTML.Value = True) Then
                    Print #1, GenerateHtmlReport(chkPreamble.Value, chkContents.Value, chkSummary.Value, _
                        chkMatches.Value, chkResponses.Value, chkDetails.Value, txtHitlistSize.Text)
                ElseIf (optTXT.Value = True) Then
                    Print #1, GenerateTxtReport(chkPreamble.Value, chkContents.Value, chkSummary.Value, _
                        chkMatches.Value, chkResponses.Value, chkDetails.Value, txtHitlistSize.Text)
                ElseIf (optXML.Value = True) Then
                    Print #1, GenerateXmlReport(chkMatches.Value, chkResponses.Value, txtHitlistSize.Text)
                Else
                    Print #1, GenerateCsvReport(txtHitlistSize.Text)
                End If
            Close
        End If
    
        Call ShellExecute(frmMain.hwnd, "Open", sFileName, "", App.Path, 1)
        Unload Me
    End If

Cancel:
    Call EnableElements
End Sub

Private Sub optCsv_GotFocus()
    Call ActivateDetails
End Sub

Private Sub optCsv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ActivateDetails
End Sub

Private Sub optHTML_GotFocus()
    Call ActivateDetails
End Sub

Private Sub optHTML_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ActivateDetails
End Sub

Private Sub optTXT_GotFocus()
    Call ActivateDetails
End Sub

Private Sub optTXT_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ActivateDetails
End Sub

Private Sub optXML_GotFocus()
    Call ActivateDetails
End Sub

Private Sub optXML_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ActivateDetails
End Sub

Private Sub txtHitlistSize_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtHitlistSize_LostFocus()
    txtHitlistSize.Text = AllowIntegersOnly(CInt(Val(txtHitlistSize.Text)), 1, 999, 20)
End Sub

Private Sub ActivateDetails()
    If optCsv.Value = True Then
        chkPreamble.Enabled = False
        chkContents.Enabled = False
        chkSummary.Enabled = False
        chkMatches.Enabled = False
        chkResponses.Enabled = False
        chkDetails.Enabled = False
    ElseIf optXML.Value = True Then
        chkPreamble.Enabled = False
        chkContents.Enabled = False
        chkSummary.Enabled = False
        chkMatches.Enabled = True
        chkResponses.Enabled = True
        chkDetails.Enabled = False
    Else
        chkPreamble.Enabled = True
        chkContents.Enabled = True
        chkSummary.Enabled = True
        chkMatches.Enabled = True
        chkResponses.Enabled = True
        chkDetails.Enabled = True
    End If
End Sub

Private Sub DisableElements()
    chkPreamble.Enabled = False
    chkContents.Enabled = False
    chkSummary.Enabled = False
    chkMatches.Enabled = False
    chkResponses.Enabled = False
    chkContents.Enabled = False
    txtHitlistSize.Enabled = False
    optHTML.Enabled = False
    optXML.Enabled = False
    optTXT.Enabled = False
    optCsv.Enabled = False
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    
    Screen.MousePointer = vbHourglass
End Sub

Private Sub EnableElements()
    chkPreamble.Enabled = True
    chkContents.Enabled = True
    chkSummary.Enabled = True
    chkMatches.Enabled = True
    chkResponses.Enabled = True
    chkContents.Enabled = True
    txtHitlistSize.Enabled = True
    optHTML.Enabled = True
    optXML.Enabled = True
    optTXT.Enabled = True
    optCsv.Enabled = True
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    
    Screen.MousePointer = vbNormal
End Sub

