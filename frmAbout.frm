VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About httprecon"
   ClientHeight    =   2190
   ClientLeft      =   30
   ClientTop       =   285
   ClientWidth     =   4950
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer timAnimation 
      Interval        =   55
      Left            =   4440
      Top             =   1680
   End
   Begin VB.CommandButton cmdOkay 
      Cancel          =   -1  'True
      Caption         =   "&Okay"
      Height          =   372
      Left            =   1920
      TabIndex        =   1
      Top             =   1680
      Width           =   1092
   End
   Begin VB.Frame Frame1 
      Height          =   1332
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.Line linLine1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         X1              =   840
         X2              =   840
         Y1              =   360
         Y2              =   1080
      End
      Begin VB.Line linLine2 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         X1              =   240
         X2              =   960
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "httprecon"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   252
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   612
      End
      Begin VB.Label lblProjectWebsite 
         AutoSize        =   -1  'True
         Caption         =   "http://www.computec.ch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1080
         MouseIcon       =   "frmAbout.frx":0CCA
         MousePointer    =   99  'Custom
         TabIndex        =   4
         ToolTipText     =   "Visit web site"
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label lblDeveloperName 
         Caption         =   "(c) 2007-2009 by Marc Ruef"
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lblApplicationName 
         Caption         =   "httprecon"
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
         Left            =   1080
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
      Begin VB.Image imgLogo 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   732
         Left            =   240
         MouseIcon       =   "frmAbout.frx":0FD4
         MousePointer    =   99  'Custom
         Picture         =   "frmAbout.frx":12DE
         Stretch         =   -1  'True
         ToolTipText     =   "Web Server Fingerprinting"
         Top             =   360
         Width           =   732
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOkay_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & APP_NAME
    lblApplicationName.Caption = APP_NAME
    lblProjectWebsite.Caption = APP_WEBSITE_URL
    Call ResetAnimation
End Sub

Private Sub imgLogo_Click()
    Call OpenProjectWebsite
End Sub

Private Sub lblProjectWebsite_Click()
    Call OpenProjectWebsite
End Sub

Private Sub timAnimation_Timer()
    If (linLine1.Y1 > 360) Then
        linLine1.Y1 = linLine1.Y1 - 20
    End If

    If (linLine2.X1 > 240) Then
        linLine2.X1 = linLine2.X1 - 20
    Else
        timAnimation.Enabled = False
    End If
    
    If (lblTitle.Width < 612) Then
        lblTitle.Width = lblTitle.Width + 20
    End If
End Sub

Private Sub ResetAnimation()
    linLine1.Y1 = 1080
    linLine2.X1 = 960
    lblTitle.Width = 12
End Sub
