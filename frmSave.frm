VERSION 5.00
Begin VB.Form frmSave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save Fingerprints"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   3975
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraFrame 
      Height          =   1692
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   3735
      Begin VB.CheckBox chkUpload 
         Caption         =   "&Submit fingerprint to project online database"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Note: Your fingerprints will be submitted to the project web site"
         Top             =   240
         Value           =   1  'Checked
         Width           =   3492
      End
      Begin VB.TextBox txtRemarks 
         Height          =   732
         Left            =   240
         MaxLength       =   127
         MultiLine       =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "Example: ""Internal and behind a Squid proxy."""
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label lblLabel 
         Caption         =   "Optional remarks for fingerprint maintainer:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   2040
      Picture         =   "frmSave.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancel Database Update"
      Top             =   3600
      Width           =   1212
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   615
      Left            =   720
      Picture         =   "frmSave.frx":0786
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Save Fingerprints"
      Top             =   3600
      Width           =   1212
   End
   Begin VB.Frame fraFrame 
      Height          =   1692
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.ComboBox cboImplementationName 
         Height          =   288
         Left            =   240
         TabIndex        =   1
         Text            =   "Apache 2.0.63"
         ToolTipText     =   "Example: Apache 2.0.63"
         Top             =   1200
         Width           =   3252
      End
      Begin VB.Label lblLabel 
         Caption         =   "Apache 2.0.63"
         Height          =   252
         Index           =   4
         Left            =   1200
         TabIndex        =   12
         Top             =   720
         Width           =   2412
      End
      Begin VB.Label lblLabel 
         Caption         =   "<name> <version> [details]"
         Height          =   252
         Index           =   2
         Left            =   1200
         TabIndex        =   11
         Top             =   480
         Width           =   2412
      End
      Begin VB.Label lblLabel 
         Caption         =   "Example:"
         Height          =   252
         Index           =   3
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   732
      End
      Begin VB.Label lblLabel 
         Caption         =   "Syntax:"
         Height          =   252
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   732
      End
      Begin VB.Label lblLabel 
         Caption         =   "Name of the httpd implementation you suspect."
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3492
      End
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboImplementationName_Click()
    Call DisableButtons
End Sub

Private Sub cboImplementationName_GotFocus()
    cboImplementationName.SelStart = 0
    cboImplementationName.SelLength = Len(cboImplementationName.Text)
End Sub

Private Sub cboImplementationName_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboImplementationName, KeyAscii, iLeftOff
End Sub

Private Sub cboImplementationName_KeyUp(KeyCode As Integer, Shift As Integer)
    Call DisableButtons
End Sub

Private Sub chkUpload_Click()
    If (chkUpload.Value = 0) Then
        txtRemarks.Enabled = False
        MsgBox "It is very sad that you do not want to participate to the project." & vbCrLf & _
            "Please submit new fingerprints, those will be added to the public" & vbCrLf & _
            "repository, to improve the quality of the enumeration.", vbInformation, "Help to improve"
    Else
        txtRemarks.Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim sImplementationName As String
    Dim sFullFingerprint As String
    Dim sRemarks As String
    
    Call DisableElements
    
    sImplementationName = cboImplementationName.Text
    sRemarks = txtRemarks.Text

    Call SaveFingerprints(sImplementationName)
        
    If (chkUpload.Value = 1) Then
        sFullFingerprint = GenerateFingerprintXML(False)
        Call PostFingerprinToWebsite(sImplementationName, sRemarks, sFullFingerprint)
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    Call DisableElements
    
    Call GenerateImplementationList
    cboImplementationName.Text = scan_besthitname
    
    cboImplementationName.SelStart = 0
    cboImplementationName.SelLength = Len(cboImplementationName.Text)
    
    txtRemarks.Text = "Target was " & scan_targethost & ":" & scan_targetport & vbCrLf & _
        "Scan time was " & scan_date & " " & scan_time & vbCrLf & _
        APP_NAME & " generated " & scan_besthitcount & " hits"
        
    Call EnableElements
End Sub

Private Sub DisableButtons()
    If (LenB(Trim$(Me.cboImplementationName.Text)) = 0) Then
        cmdSave.Enabled = False
    Else
        cmdSave.Enabled = True
    End If
End Sub

Private Sub GenerateImplementationList()
    Dim iListItemCount As Integer
    Dim i As Integer
    
    Call ListViewSort(frmMain.lsvResults, frmMain.lsvResults.ColumnHeaders(2), 0)
    iListItemCount = frmMain.lsvResults.ListItems.Count
    
    For i = 1 To iListItemCount
         cboImplementationName.AddItem frmMain.lsvResults.ListItems(i).ListSubItems(1).Text
    Next i

    Call ListViewSort(frmMain.lsvResults, frmMain.lsvResults.ColumnHeaders(3), 1)
End Sub

Private Sub DisableElements()
    cboImplementationName.Enabled = False
    chkUpload.Enabled = False
    txtRemarks.Enabled = False
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    
    Screen.MousePointer = vbHourglass
End Sub

Private Sub EnableElements()
    cboImplementationName.Enabled = True
    chkUpload.Enabled = True
    If (chkUpload.Value = 1) Then
        txtRemarks.Enabled = True
    End If
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    
    Screen.MousePointer = vbNormal
End Sub
