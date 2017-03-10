VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "httprecon"
   ClientHeight    =   7260
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8055
   HasDC           =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lsvResultsForTest 
      CausesValidation=   0   'False
      Height          =   2295
      Left            =   5520
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4048
      SortKey         =   3
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      SmallIcons      =   "imlHttpdIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "string"
         Object.Width           =   512
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "string"
         Text            =   "Name"
         Object.Width           =   7410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "number"
         Text            =   "Hits"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "number"
         Text            =   "Match %"
         Object.Width           =   1411
      EndProperty
   End
   Begin MSComctlLib.ListView lsvResults 
      CausesValidation=   0   'False
      Height          =   2175
      Left            =   240
      TabIndex        =   6
      Top             =   4560
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3836
      SortKey         =   3
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      SmallIcons      =   "imlHttpdIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "string"
         Object.Width           =   512
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "string"
         Text            =   "Name"
         Object.Width           =   7410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "number"
         Text            =   "Hits"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "number"
         Text            =   "Match %"
         Object.Width           =   1411
      EndProperty
   End
   Begin MSComctlLib.ImageList imlHttpdIcons 
      Left            =   7440
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   101
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E43
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0ECC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1143
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":137D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1777
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2130
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2217
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":243E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2931
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A53
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F49
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":325A
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":33B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":361E
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3703
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B07
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3ED9
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":42B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":46FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4ADF
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5305
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":56FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5B07
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5EA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":62D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6655
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":69CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6DD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":71DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":75E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A04
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7DDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":81CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8623
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8A81
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8E70
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9227
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":95E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9A22
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9DB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A1F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A594
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A97B
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AD0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B0E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B46A
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B829
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BBDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BF7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C398
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C791
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CBCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CF66
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D346
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D73E
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DB38
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DF48
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E352
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E776
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EB59
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EF7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F3AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F786
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FB45
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FF45
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1034D
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":106EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10AA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10E51
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":111F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":115A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":119F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11DE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1220F
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1263D
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":129C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12D4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":130E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1347A
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1389D
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13CE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14064
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":144A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1487A
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14C36
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15026
            Key             =   ""
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15484
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1584E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtReportPreview 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   4560
      Visible         =   0   'False
      Width           =   7575
   End
   Begin RichTextLib.RichTextBox rtbResponses 
      Height          =   2535
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4471
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":15BDB
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdgOpenScanlist 
      Left            =   3960
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "Scanlist Files (*.scl)|*.scl"
      DialogTitle     =   "Open Scanlist Files"
      Filter          =   "Scanlist Files (*.scl)|*.scl|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
   End
   Begin MSComDlg.CommonDialog cdgOpenScan 
      Left            =   4440
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "Fingerprint Scan Files (*.fps)|*.fps"
      DialogTitle     =   "Open Fingerprint Scan Files"
      Filter          =   "Fingerprint Scan Files (*.fps)|*.fps|All Files (*.*)|*.*"
   End
   Begin MSComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   12
      Top             =   6960
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   529
      Style           =   1
      SimpleText      =   "Ready."
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdgSaveAsScan 
      Left            =   4920
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "Fingerprint Scan Files (*.fps)|*.fps"
      DialogTitle     =   "Save Fingerprint Scan Files"
      FileName        =   "127-0-0-1-80.fps"
      Filter          =   "Fingerprints Scan Files (*.fps)|*.fps|All Files (*.*)|*.*"
   End
   Begin MSComctlLib.TabStrip tbsViews 
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5318
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   9
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "GET existing"
            Object.ToolTipText     =   "GET / HTTP/1.1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "GET long request"
            Object.ToolTipText     =   "GET /aaa(...) HTTP/1.1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "GET non-existing"
            Object.ToolTipText     =   "GET /404test_.html HTTP/1.1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "GET wrong protocol"
            Object.ToolTipText     =   "GET / HTTP/9.8"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "HEAD existing"
            Object.ToolTipText     =   "HEAD / HTTP/1.1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "OPTIONS common"
            Object.ToolTipText     =   "OPTIONS / HTTP/1.1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "DELETE existing"
            Object.ToolTipText     =   "DELETE / HTTP/1.1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "TEST method"
            Object.ToolTipText     =   "TEST / HTTP/1.1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Attack Request"
            Object.ToolTipText     =   "GET <attack_request> HTTP/1.1"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraTarget 
      Caption         =   "Target"
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   7812
      Begin VB.ComboBox cboScheme 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Usually: http (non-encrypted)"
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cboTargetPort 
         Height          =   315
         Left            =   4080
         TabIndex        =   1
         Text            =   "80"
         ToolTipText     =   "Usually: 80 (http)"
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdAnalyze 
         Caption         =   "&Analyze"
         Height          =   375
         Left            =   6480
         TabIndex        =   2
         ToolTipText     =   "Analyze Web Server"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtTargetHost 
         Height          =   285
         Left            =   1200
         MaxLength       =   255
         TabIndex        =   0
         Text            =   "127.0.0.1"
         ToolTipText     =   "Example: www.computec.ch"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   ":"
         Height          =   255
         Left            =   3840
         TabIndex        =   11
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.TextBox txtFingerprint 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   4560
      Visible         =   0   'False
      Width           =   7575
   End
   Begin MSComctlLib.TabStrip tbsResults 
      Height          =   2655
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4683
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Full Matchlist"
            Object.ToolTipText     =   "Full List of Matches"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fingerprint Details"
            Object.ToolTipText     =   "Full Fingerprint Details"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Report Preview"
            Object.ToolTipText     =   "Text Report Preview"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNewItem 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpenScanlistItem 
         Caption         =   "Open Scan&list..."
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpenScanItem 
         Caption         =   "&Open Scan..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSaveAsScanItem 
         Caption         =   "&Save As Scan..."
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExitItem 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuEditCopyItem 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAllItem 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuConfiguration 
      Caption         =   "&Configuration"
      Begin VB.Menu mnuConfigurationEditItem 
         Caption         =   "&Edit Settings..."
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuConfigurationSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfigurationReloadItem 
         Caption         =   "&Reload Configuration"
         Shortcut        =   +{F5}
      End
   End
   Begin VB.Menu mnuFingerprinting 
      Caption         =   "Finger&printing"
      Begin VB.Menu mnuFingerprintingAnalyzeItem 
         Caption         =   "&Analyze (network access)"
      End
      Begin VB.Menu mnuFingerprintingReanalyzeItem 
         Caption         =   "&Re-Analyze (without network)"
         Enabled         =   0   'False
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFingerprintingSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFingerprintingOpenSiteInBrowserItem 
         Caption         =   "Open Web Site in &Browser..."
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuFingerprintingSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFingerprintingOnlineDBItem 
         Caption         =   "&Online Fingerprint Database..."
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuFingerprintingSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFingerprintingSaveFingerprintItem 
         Caption         =   "&Save Fingerprint..."
         Enabled         =   0   'False
         Shortcut        =   ^{F3}
      End
   End
   Begin VB.Menu mnuReporting 
      Caption         =   "&Reporting"
      Begin VB.Menu mnuReportingGenerateReportItem 
         Caption         =   "&Generate Report..."
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpOnlineDocumentationItem 
         Caption         =   "Online &Documentation..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpOnlineFAQItem 
         Caption         =   "Online &FAQ..."
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpCheckForUpdatesItem 
         Caption         =   "Check for &Updates..."
      End
      Begin VB.Menu mnuHelpHomepageItem 
         Caption         =   "httprecon &Home Page..."
         Shortcut        =   +^{F1}
      End
      Begin VB.Menu mnuHelpSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAboutItem 
         Caption         =   "&About"
         Shortcut        =   +{F1}
      End
   End
   Begin VB.Menu mnuPopUpListing 
      Caption         =   "PopUpListing"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUpListingTopMatchesOfTestItem 
         Caption         =   "&Top Matches of Selected Test..."
      End
      Begin VB.Menu mnuPopUpListingDetailAnalysisItem 
         Caption         =   "&Detail Analysis"
      End
      Begin VB.Menu mnuPopUpListingSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopUpListingInvestigateFindingItem 
         Caption         =   "&Investigate Finding with Google..."
      End
      Begin VB.Menu mnuPopUpListingLookupVulnerabilitiesItem 
         Caption         =   "&Lookup Vulnerabilities on OSVDB..."
      End
      Begin VB.Menu mnuPopUpListingSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopUpListingCopyThisHitItem 
         Caption         =   "Copy &This Hit to Clipboard"
      End
      Begin VB.Menu mnuPopUpListingCopyBestHitsItem 
         Caption         =   "Copy &Best Hits to Clipboard..."
      End
      Begin VB.Menu mnuPopUpListingCopyAllHitsItem 
         Caption         =   "Copy &All Hits to Clipboard"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' httprecon
' (c) 2007-2009 by Marc Ruef <marc.ruef-at-computec.ch>
' http://www.computec.ch/projekte/httprecon/
'
' This file is part of httprecon.
'
' httprecon is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' httprecon is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with httprecon. If not, see <http://www.gnu.org/licenses/>.

Private Sub cboScheme_Click()
    If (cboScheme.Text = "http://") Then
        Call ChangeSSLMode(False)
    Else
        Call ChangeSSLMode(True)
    End If
End Sub

Private Sub cboScheme_KeyUp(KeyCode As Integer, Shift As Integer)
    If (cboScheme.Text = "http://") Then
        Call ChangeSSLMode(False)
    Else
        Call ChangeSSLMode(True)
    End If
End Sub

Private Sub cboTargetPort_Change()
    Dim sInput As String
    
    sInput = cboTargetPort.Text
    
    Call ChangeSSLMode(False)
    If (LenB(sInput) = 0) Then
        cboTargetPort.Text = 80
    ElseIf (sInput > 65535) Then
        cboTargetPort.Text = 65535
    ElseIf (sInput = 443) Then
        Call ChangeSSLMode(True)
    ElseIf (sInput = 8443) Then
        Call ChangeSSLMode(True)
    End If
End Sub

Private Sub cboTargetPort_Click()
    Call cboTargetPort_Change
End Sub

Private Sub cboTargetPort_GotFocus()
    cboTargetPort.SelStart = 0
    cboTargetPort.SelLength = Len(cboTargetPort.Text)
End Sub

Private Sub cboTargetPort_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select

    Static iLeftOff As Long
    ComboAutoComplete cboTargetPort, KeyAscii, iLeftOff
End Sub

Private Sub cboTargetPort_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13) Then
        Call ServerAnalysis
    End If
End Sub

Private Sub cmdAnalyze_Click()
    Call ServerAnalysis
End Sub

Private Sub Form_Load()
    frmMain.Caption = APP_NAME
    
    cboScheme.AddItem ("http://")
    cboScheme.AddItem ("https://")
    
    With cboTargetPort
        .AddItem ("80")
        .AddItem ("81")
        .AddItem ("82")
        .AddItem ("443")
        .AddItem ("800")
        .AddItem ("888")
        .AddItem ("2301")
        .AddItem ("8000")
        .AddItem ("8001")
        .AddItem ("8080")
        .AddItem ("8081")
        .AddItem ("8443")
        .AddItem ("8888")
    End With
    
    Randomize
    Call LoadConfigFromFile
    
    txtTargetHost.Text = scan_targethost
    cboTargetPort.Text = scan_targetport
    If (scan_targetsecure = 1) Then
        Call ChangeSSLMode(True)
    Else
        Call ChangeSSLMode(False)
    End If
    
    Call InitializeDirectories
    Call InitializeFiles

    rtbResponses.RightMargin = rtbResponses.Width + 1
    
    Call ChangeStatusBarReady
End Sub

Private Sub Form_Resize()
    If frmMain.WindowState <> vbMinimized Then
        If WindowState <> vbMinimized Then
            If Height < 6000 Then
                Height = 6000
            End If
            
            If Width < 7000 Then
                Width = 7000
            End If
        End If
    
        fraTarget.Width = frmMain.Width - 360
        cmdAnalyze.Left = fraTarget.Width - cmdAnalyze.Width - 120
        
        tbsViews.Width = fraTarget.Width
        rtbResponses.Width = fraTarget.Width - 240
        tbsViews.Height = (frmMain.Height - fraTarget.Height - stbStatus.Height) / 2 - 480
        rtbResponses.Height = tbsViews.Height - 480
        
        tbsResults.Top = tbsViews.Top + tbsViews.Height + 120
        tbsResults.Width = fraTarget.Width
        
        lsvResults.Width = rtbResponses.Width
        lsvResults.Top = tbsResults.Top + 360
        txtFingerprint.Width = lsvResults.Width
        txtFingerprint.Top = lsvResults.Top
        txtReportPreview.Width = lsvResults.Width
        txtReportPreview.Top = lsvResults.Top
        
        tbsResults.Height = tbsViews.Height - 360
        lsvResults.Height = tbsResults.Height - 480
        txtFingerprint.Height = lsvResults.Height
        txtReportPreview.Height = lsvResults.Height
    End If
End Sub

Private Sub lsvResults_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call ListViewSort(lsvResults, ColumnHeader, (lsvResults.SortOrder + 1) Mod 2)
End Sub

Private Sub lsvResults_DblClick()
    If lsvResults.ListItems.Count Then
        Call ShowBestHitsForTest
    End If
End Sub

Private Sub lsvResults_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lsvResults.ListItems.Count Then
        Call IdentifyServerFingerprint(scan_test_folder, rtbResponses.Text, lsvResults.SelectedItem.SubItems(1))
    End If
End Sub

Private Sub lsvResults_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ListVwItem As MSComctlLib.ListItem
    Dim sValue As String
    
    If KeyCode = 13 Then
        If lsvResults.ListItems.Count Then
            Call ShowBestHitsForTest
        End If
    Else
        sValue = ChrW$(KeyCode)
        For Each ListVwItem In lsvResults.ListItems
            If Mid$(LCase(ListVwItem.SubItems(1)), 1, 1) = LCase$(sValue) Then
                ListVwItem.Selected = True
                ListVwItem.EnsureVisible
                Exit For
            End If
        Next
        lsvResults.SetFocus
    End If
End Sub

Private Sub lsvResults_KeyUp(KeyCode As Integer, Shift As Integer)
    If lsvResults.ListItems.Count Then
        Call IdentifyServerFingerprint(scan_test_folder, rtbResponses.Text, lsvResults.SelectedItem.SubItems(1))
    End If
End Sub

Private Sub lsvResults_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If lsvResults.ListItems.Count Then
            PopupMenu mnuPopUpListing
        End If
    End If
End Sub

Private Sub mnuConfigurationEditItem_Click()
    frmConfiguration.Show vbModal, frmMain
End Sub

Private Sub mnuConfigurationReloadItem_Click()
    Call LoadConfigFromFile(app_configuration_filename)
End Sub

Private Sub mnuEditCopyItem_Click()
    Clipboard.Clear
    Clipboard.SetText ActiveControl.SelText, vbCFText
End Sub

Private Sub mnuEditSelectAllItem_Click()
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
End Sub

Private Sub mnuFileExitItem_Click()
    Unload Me
End Sub

Private Sub mnuFileNewItem_Click()
    Call ResetAll
End Sub

Private Sub mnuFileOpenScanItem_Click()
    Dim sFileName As String
    Dim sFileContent As String
    
    If (Dir$(App.Path & "\scans", 16) <> "") Then
        cdgOpenScan.InitDir = App.Path & "\scans"
    Else
        cdgOpenScan.InitDir = App.Path
    End If
    
    On Error GoTo Cancel
    cdgOpenScan.ShowOpen
    sFileName = cdgOpenScan.FileName
    
    If LenB(sFileName) Then
        If (Dir$(sFileName, 16) <> "") Then
            Call ResetAll
            
            sFileContent = ReadFile(sFileName)
            Call ReadFingerprintXML(sFileContent)
            frmMain.Caption = UpdateCaption & " (" & Mid$(sFileName, InStrRev(sFileName, "\", , vbBinaryCompare) + 1) & ")"
            
            txtTargetHost = scan_targethost
            cboTargetPort = scan_targetport
            If (scan_targetsecure) Then
                Call ChangeSSLMode(True)
            Else
                Call ChangeSSLMode(False)
            End If
        
            Call AnalyzeFingerprintsAndShowResult
        End If
    End If

Cancel:
End Sub

Private Sub mnuFileOpenScanlistItem_Click()
    Dim sFileName As String
    Dim sFileContent As String
    Dim sScanListItems() As String
    Dim iScanListItemsCount As Integer
    Dim i As Integer
    Dim sReportPath As String
    Dim sReportFileName As String
    
    cdgOpenScanlist.InitDir = App.Path
    
    On Error GoTo Cancel
    cdgOpenScanlist.ShowOpen
    sFileName = cdgOpenScanlist.FileName
    
    If LenB(sFileName) Then
        If (Dir$(sFileName, 16) <> "") Then
            sFileContent = ReadFile(sFileName)
            sScanListItems = Split(sFileContent, vbCrLf, , vbBinaryCompare)
            iScanListItemsCount = UBound(sScanListItems)
            
            sReportPath = BrowseForFolder(Me, "Choose the destination directory for report files (html export and scan fingerprint).")
            
            If (LenB(sReportPath)) Then
                For i = 0 To iScanListItemsCount
                    If (LenB(sScanListItems(i))) Then
                        Call ResetAll
                        
                        If (Left$(sScanListItems(i), 8) = "https://") Then
                            scan_targetsecure = 1
                            Call ChangeSSLMode(True)
                        Else
                            scan_targetsecure = 0
                            Call ChangeSSLMode(False)
                        End If
                        
                        scan_targetport = ExtractTargetPort(sScanListItems(i))
                        cboTargetPort = scan_targetport
                        
                        scan_targethost = SanitizeHostInput(sScanListItems(i))
                        txtTargetHost = scan_targethost
                        
                        Call ServerAnalysis
                        
                        'This is for training mode only
                        'frmSave.Show vbModal, frmMain
                        
                        sReportFileName = sReportPath & "\" & StringToFileName(scan_targethost & ":" & scan_targetport) & ".html"
                        On Error Resume Next
                        Open sReportFileName For Output As #1
                            Print #1, GenerateHtmlReport(1, 1, 1, 1, 1, 1, 20)
                        Close
                        
                        sReportFileName = sReportPath & "\" & StringToFileName(scan_targethost & ":" & scan_targetport) & ".fps"
                        Open sReportFileName For Output As #1
                            Print #1, GenerateFingerprintXML(True)
                        Close
                        
                        Call ChangeStatusBar("Scanlist with " & iScanListItemsCount & " items finished. Ready.")
                    End If
                Next i
            End If
        End If
    End If

Cancel:
End Sub

Private Sub mnuFileSaveAsScanItem_Click()
    Dim sFileName As String
    Dim sOverride As String
    
    If (Dir$(App.Path & "\scans", 16) <> "") Then
        cdgSaveAsScan.InitDir = App.Path & "\scans"
    Else
        cdgSaveAsScan.InitDir = App.Path
    End If
    
    cdgSaveAsScan.FileName = StringToFileName(scan_targethost & ":" & scan_targetport) & ".fps"
    
    On Error GoTo Cancel
    cdgSaveAsScan.ShowSave
    sFileName = cdgSaveAsScan.FileName
    
    If (LenB(sFileName)) Then
        If (Dir$(sFileName, 16) <> "") Then
            sOverride = MsgBox(sFileName & " already exists." & vbCrLf & "Do you want to replace it?", vbExclamation + vbYesNo, "Scan Save As")
        Else
            sOverride = 6
        End If
        
        If (sOverride = 6) Then
            Open sFileName For Output As #1
                Print #1, GenerateFingerprintXML(True)
            Close
            
            frmMain.Caption = UpdateCaption & " (" & Mid$(sFileName, InStrRev(sFileName, "\", , vbBinaryCompare) + 1) & ")"
        End If
    End If

Cancel:
End Sub

Private Sub mnuFingerprintingAnalyzeItem_Click()
    Call ServerAnalysis
End Sub

Private Sub mnuFingerprintingOnlineDBItem_Click()
    Call ShellExecute(frmMain.hwnd, "Open", PROJECT_WEBDB, "", App.Path, 1)
End Sub

Private Sub mnuFingerprintingOpenSiteInBrowserItem_Click()
    Dim sScheme As String
    
    Call ChangeStatusBar("Open web site in browser...")
    If (cboScheme.Text = "https://") Then
        sScheme = "https://"
    Else
        sScheme = "http://"
    End If
    
    Call ShellExecute(frmMain.hwnd, "Open", sScheme & txtTargetHost.Text & ":" & CInt(cboTargetPort.Text), "", App.Path, 1)
    Call ChangeStatusBarDone
End Sub

Private Sub mnuFingerprintingReanalyzeItem_Click()
    Call DisableElements
    Call AnalyzeFingerprintsAndShowResult
    Call EnableElements
End Sub

Private Sub mnuFingerprintingSaveFingerprintItem_Click()
    frmSave.Show vbModal, frmMain
End Sub

Private Sub mnuHelpAboutItem_Click()
    frmAbout.Show vbModal, frmMain
End Sub

Private Sub mnuHelpCheckForUpdatesItem_Click()
    frmUpdate.Show vbModal, frmMain
End Sub

Private Sub mnuHelpHomepageItem_Click()
    Call OpenProjectWebsite
End Sub

Private Sub mnuHelpOnlineDocumentationItem_Click()
    Call ChangeStatusBar("Open browserrecon project online documentation...")
    Call ShellExecute(frmMain.hwnd, "Open", "http://www.computec.ch/projekte/httprecon/?s=documentation", "", App.Path, 1)
    Call ChangeStatusBarDone
End Sub

Private Sub mnuHelpOnlineFAQItem_Click()
    Call ChangeStatusBar("Open browserrecon project online FAQ...")
    Call ShellExecute(frmMain.hwnd, "Open", "http://www.computec.ch/projekte/httprecon/?s=faq", "", App.Path, 1)
    Call ChangeStatusBarDone
End Sub

Private Sub mnuPopUpListingCopyAllHitsItem_Click()
    Clipboard.Clear
    Clipboard.SetText GenerateHitListCsv(frmMain.lsvResults, lsvResults.ListItems.Count, ";")
End Sub

Private Sub mnuPopUpListingCopyBestHitsItem_Click()
    Dim iCount As Integer
    
    iCount = CInt(Val(InputBox("How many items shall your top list contain?", "Create Top List", 20)))
    iCount = AllowIntegersOnly(CLng(iCount), 1, 999, 20)
    
    Clipboard.Clear
    Clipboard.SetText GenerateHitListCsv(frmMain.lsvResults, iCount, ";")
End Sub

Private Sub mnuPopUpListingCopyThisHitItem_Click()
    Dim sHitData As String
    
    sHitData = "Implementation: " & lsvResults.SelectedItem.SubItems(1) & vbCrLf & _
        "Hits: " & lsvResults.SelectedItem.SubItems(2) & vbCrLf & _
        "Match: " & lsvResults.SelectedItem.SubItems(3) & "%" & vbCrLf
    
    Clipboard.Clear
    Clipboard.SetText sHitData
End Sub

Private Sub mnuPopUpListingDetailAnalysisItem_Click()
    If lsvResults.ListItems.Count Then
        Call IdentifyServerFingerprint(scan_test_folder, rtbResponses.Text, lsvResults.SelectedItem.SubItems(1))
        rtbResponses.SetFocus
    End If
End Sub

Private Sub mnuPopUpListingInvestigateFindingItem_Click()
    Call InvestigateFinding
End Sub

Private Sub mnuPopUpListingLookupVulnerabilitiesItem_Click()
    Call LookupVulnerabilities
End Sub

Private Sub mnuPopUpListingTopMatchesOfTestItem_Click()
    Call ShowBestHitsForTest
End Sub

Private Sub mnuReportingGenerateReportItem_Click()
    frmReport.Show vbModal, frmMain
End Sub

Private Sub rtbResponses_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If rtbResponses.SelLength Then
            mnuEditCopyItem.Enabled = True
        Else
            mnuEditCopyItem.Enabled = False
        End If
        
        If (LenB(rtbResponses.Text)) Then
            mnuEditSelectAllItem.Enabled = True
        Else
            mnuEditSelectAllItem.Enabled = False
        End If
        
        PopupMenu mnuEdit
    End If
End Sub

Private Sub tbsResults_Click()
    Dim iIndex As Integer
    
    iIndex = tbsResults.SelectedItem.Index
    
    If (iIndex = 1) Then
        lsvResults.Visible = True
        txtFingerprint.Visible = False
        txtReportPreview.Visible = False
    ElseIf (iIndex = 2) Then
        txtFingerprint.Visible = True
        lsvResults.Visible = False
        txtReportPreview.Visible = False
    ElseIf (iIndex = 3) Then
        txtReportPreview.Visible = True
        lsvResults.Visible = False
        txtFingerprint.Visible = False
        Call UpdateReportPreview
    End If
End Sub

Private Sub tbsViews_Click()
    Call FillResponses
End Sub

Private Sub txtTargetHost_GotFocus()
    txtTargetHost.SelStart = 0
    txtTargetHost.SelLength = Len(txtTargetHost.Text)
End Sub

Private Sub txtTargetHost_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13) Then
        Call ServerAnalysis
    End If
End Sub

Private Sub txtTargetHost_LostFocus()
    Dim sNewTarget As String
    
    sNewTarget = txtTargetHost.Text
    If (LenB(sNewTarget)) Then
        sNewTarget = SanitizeHostInput(sNewTarget)
        txtTargetHost.Text = sNewTarget
        
        cmdAnalyze.Enabled = True
        mnuFingerprintingAnalyzeItem.Enabled = True
        mnuFingerprintingOpenSiteInBrowserItem.Enabled = True
    Else
        cmdAnalyze.Enabled = False
        mnuFingerprintingAnalyzeItem.Enabled = False
        mnuFingerprintingOpenSiteInBrowserItem.Enabled = False
    End If
End Sub

Public Function UpdateCaption() As String
    Dim sScheme As String

    If (scan_targetsecure = 1) Then
        sScheme = "https://"
    Else
        sScheme = "http://"
    End If
    
    UpdateCaption = APP_NAME & " - " & sScheme & scan_targethost & ":" & scan_targetport & "/"
End Function

Public Sub UpdateReportPreview()
    If (lsvResults.ListItems.Count) Then
        txtReportPreview.Text = GenerateTxtReport(1, 1, 1, 1, 1, 1, 20)
    Else
        txtReportPreview.Text = vbNullString
    End If
End Sub

Public Sub ResetAll()
    frmMain.Caption = APP_NAME
    
    scan_besthitcount = 0
    scan_besthitname = vbNullString
    
    scan_time = vbNullString
    scan_date = vbNullString
    scan_targethost = "127.0.0.1"
    scan_targetport = 80
    Call ChangeSSLMode(False)
    
    frmMain.fraTarget.Caption = "Target"
    
    Call ResetResponseHighlight
    
    With frmMain
        .txtTargetHost = scan_targethost
        .cboTargetPort = scan_targetport
        .cboScheme.ListIndex = 0
    End With
    
    response_attackrequest = vbNullString
    response_delete = vbNullString
    response_getexist = vbNullString
    response_getlongrequest = vbNullString
    response_get_nonexistent = vbNullString
    response_head = vbNullString
    response_options = vbNullString
    response_testmethod = vbNullString
    response_protocolversion = vbNullString
    
    timing_attackrequest = Empty
    timing_delete = Empty
    timing_getexist = Empty
    timing_getlongrequest = Empty
    timing_get_nonexistent = Empty
    timing_head = Empty
    timing_options = Empty
    timing_testmethod = Empty
    timing_protocolversion = Empty
    
    With frmMain
        .lsvResults.ListItems.Clear
        .rtbResponses.Text = vbNullString
        .rtbResponses.ToolTipText = vbNullString
        .txtFingerprint.Text = vbNullString
        .txtReportPreview.Text = vbNullString
    
        .mnuFileSaveAsScanItem.Enabled = False
        .mnuFingerprintingReanalyzeItem.Enabled = False
        .mnuFingerprintingSaveFingerprintItem.Enabled = False
        .mnuReportingGenerateReportItem.Enabled = False

        .txtTargetHost.SelStart = 0
        .txtTargetHost.SelLength = Len(.txtTargetHost.Text)
        .txtTargetHost.SetFocus
    End With
End Sub

Public Sub ChangeSSLMode(ByRef bSecure As Boolean)
    If (bSecure = False) Then
        frmMain.cboScheme.ListIndex = 0
        scan_targetsecure = 0
    Else
        frmMain.cboScheme.ListIndex = 1
        scan_targetsecure = 1
    End If
End Sub

Public Sub DisableElements()
    rtbResponses.SetFocus
    rtbResponses.Text = vbNullString
    txtFingerprint.Text = vbNullString
    txtReportPreview.Text = vbNullString
    lsvResults.ListItems.Clear
    cmdAnalyze.Enabled = False
    mnuFileNewItem.Enabled = False
    mnuFileOpenScanlistItem.Enabled = False
    mnuFileOpenScanItem.Enabled = False
    mnuFileSaveAsScanItem.Enabled = False
    mnuConfigurationEditItem.Enabled = False
    mnuConfigurationReloadItem.Enabled = False
    mnuFingerprintingAnalyzeItem.Enabled = False
    mnuFingerprintingReanalyzeItem.Enabled = False
    mnuFingerprintingSaveFingerprintItem.Enabled = False
    mnuReportingGenerateReportItem.Enabled = False
    txtTargetHost.Enabled = False
    cboTargetPort.Enabled = False
    cboScheme.Enabled = False
    
    Screen.MousePointer = vbArrowHourglass
End Sub

Public Sub EnableElements()
    cmdAnalyze.Enabled = True
    mnuFileNewItem.Enabled = True
    mnuFileOpenScanlistItem.Enabled = True
    mnuFileOpenScanItem.Enabled = True
    mnuFileSaveAsScanItem.Enabled = True
    mnuConfigurationEditItem.Enabled = True
    mnuConfigurationReloadItem.Enabled = True
    mnuFingerprintingAnalyzeItem.Enabled = True
    mnuFingerprintingReanalyzeItem.Enabled = True
    mnuFingerprintingSaveFingerprintItem.Enabled = True
    txtTargetHost.Enabled = True
    cboTargetPort.Enabled = True
    cboScheme.Enabled = True
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub InvestigateFinding()
    Dim sSearch As String
    
    If lsvResults.ListItems.Count Then
        sSearch = lsvResults.SelectedItem.SubItems(1)
    
        Call ChangeStatusBar("Open web browser to google to search for" & sSearch & "...")
        Call ShellExecute(frmMain.hwnd, "Open", "http://www.google.com/search?q=" & sSearch, "", App.Path, 1)
        Call ChangeStatusBarDone
    End If
End Sub

Private Sub LookupVulnerabilities()
    Dim sSearch As String
    
    If lsvResults.ListItems.Count Then
        sSearch = lsvResults.SelectedItem.SubItems(1)
    
        Call ChangeStatusBar("Open web browser to OSVDB to lookup vulnerabilities for " & sSearch & "...")
        Call ShellExecute(frmMain.hwnd, "Open", "http://osvdb.org/search?request=" & sSearch, "", App.Path, 1)
        Call ChangeStatusBarDone
    End If
End Sub

Public Sub ShowBestHitsForTest()
    Call AnalyzeTestFingerprintsAndShowResult
    MsgBox "The top results for the selected response only are:" & vbCrLf & vbCrLf & _
        GenerateHitListCsv(lsvResultsForTest, 10, " - "), vbInformation + vbOKOnly, "Matchlist for Test"
    lsvResultsForTest.ListItems.Clear
End Sub

