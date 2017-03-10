VERSION 5.00
Begin VB.Form frmUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Software Update"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   6750
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   615
      Left            =   3480
      Picture         =   "frmUpdate.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   615
      Left            =   2160
      Picture         =   "frmUpdate.frx":07B5
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Frame frmUpdateAailable 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton cmdCheck 
         Caption         =   "Check &now"
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Top             =   5040
         Width           =   2535
      End
      Begin VB.Label lblFeatures 
         Height          =   1335
         Left            =   240
         TabIndex        =   10
         Top             =   3480
         Width           =   6135
      End
      Begin VB.Label lblBugfixes 
         Height          =   1335
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   6135
      End
      Begin VB.Label lblLabel 
         Caption         =   "Features:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label lblLabel 
         Caption         =   "Bugfixes:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblImportance 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   4680
         TabIndex        =   6
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblDate 
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Caption         =   "loading..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   6135
      End
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheck_Click()
    cmdCheck.Enabled = False
    cmdClose.Enabled = False
    cmdUpdate.Enabled = False
    Call UpdateCheck
    cmdCheck.Enabled = True
    cmdClose.Enabled = True
    cmdUpdate.Enabled = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub UpdateCheck()
    cmdCheck.Enabled = False
    Call ParseUpdateInformation(SendGetRequest(APP_WEBSITE_URL & "?s=updatecheck&v=" & APP_NAME))
    cmdCheck.Enabled = True
End Sub

Private Sub ParseUpdateInformation(ByRef sResponse As String)
    Dim sElements() As String
    Dim sBugfixes() As String
    Dim sFeatures() As String
    Dim sBugfixesListing As String
    Dim sFeaturesListing As String
    Dim iBugfixesCount As Integer
    Dim iFeaturesCount As Integer
    Dim sDownloadVersion As String
    Dim sDownloadUrl As String
    Dim i As Integer
    
    sElements = Split(sResponse, ";", , vbBinaryCompare)
    sBugfixes = Split(sElements(3), "*", , vbBinaryCompare)
    sFeatures = Split(sElements(4), "*", , vbBinaryCompare)
    iBugfixesCount = UBound(sBugfixes)
    iFeaturesCount = UBound(sFeatures)
    
    If (Len(sElements(0)) > 2) And (Len(sElements(0)) < 5) Then
        sDownloadVersion = "httprecon " & sElements(0)
        lblTitle = sDownloadVersion
    Else
        lblTitle = "download failed"
    End If

    If (Len(sElements(2)) > 5) Then
        lblImportance.Caption = sElements(1)
    Else
        lblImportance.Caption = vbNullString
    End If

    If (Len(sElements(2)) = 8) Then
        lblDate.Caption = Mid$(sElements(2), 5, 2) & "/" & Mid$(sElements(2), 7, 2) & "/" & Mid$(sElements(2), 1, 4)
        cmdUpdate.Enabled = True
    Else
        lblDate.Caption = vbNullString
        cmdUpdate.Enabled = False
    End If

    For i = 0 To iBugfixesCount
        If (i = 0) Then
            sBugfixesListing = "* " & sBugfixes(0) & vbCrLf
        Else
            sBugfixesListing = sBugfixesListing & "* " & sBugfixes(i) & vbCrLf
        End If
    Next i
    lblBugfixes.Caption = sBugfixesListing

    For i = 0 To iFeaturesCount
        If (i = 0) Then
            sFeaturesListing = "* " & sFeatures(0) & vbCrLf
        Else
            sFeaturesListing = sFeaturesListing & "* " & sFeatures(i) & vbCrLf
        End If
    Next i
    lblFeatures.Caption = sFeaturesListing

    sDownloadUrl = RTrim(sElements(5))
    cmdUpdate.ToolTipText = sDownloadUrl
    
    If (sDownloadVersion <> APP_NAME) Then
        Call ShellExecute(frmMain.hwnd, "Open", sDownloadUrl, "", App.Path, 1)
    End If
End Sub

Private Sub cmdUpdate_Click()
    If (Mid$(cmdUpdate.ToolTipText, 1, 4) = "http") Then
        Call ShellExecute(frmMain.hwnd, "Open", cmdUpdate.ToolTipText, "", App.Path, 1)
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Call UpdateCheck
End Sub
