VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfiguration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuration Settings"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   Icon            =   "frmConfiguration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTiming 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   240
      TabIndex        =   23
      Top             =   600
      Width           =   6375
      Begin VB.TextBox txtTimingReceive 
         Height          =   285
         Left            =   120
         MaxLength       =   5
         TabIndex        =   4
         Text            =   "5000"
         ToolTipText     =   "5000"
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox txtTimingSend 
         Height          =   285
         Left            =   120
         MaxLength       =   5
         TabIndex        =   3
         Text            =   "5000"
         ToolTipText     =   "5000"
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtTimingConnect 
         Height          =   285
         Left            =   120
         MaxLength       =   5
         TabIndex        =   2
         Text            =   "5000"
         ToolTipText     =   "5000"
         Top             =   840
         Width           =   615
      End
      Begin VB.Image imgImage 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   120
         Picture         =   "frmConfiguration.frx":058A
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgImage 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   120
         Picture         =   "frmConfiguration.frx":095D
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgImage 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   120
         Picture         =   "frmConfiguration.frx":0D2C
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblLabel 
         Caption         =   "Amount of time in milliseconds which shall be waited for a provoked response before aborting with a timeout. Suggested value: 5000"
         Height          =   495
         Index           =   9
         Left            =   120
         TabIndex        =   46
         Top             =   2760
         Width           =   6135
      End
      Begin VB.Label lblLabel 
         Caption         =   "Amount of time in milliseconds which shall be waited to send a full request before aborting with a timeout. Suggested value: 5000"
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   45
         Top             =   1560
         Width           =   6135
      End
      Begin VB.Label lblLabel 
         Caption         =   $"frmConfiguration.frx":10C3
         Height          =   495
         Index           =   5
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   6135
      End
      Begin VB.Label lblLabel 
         Caption         =   "req_timeout_receive"
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
         Index           =   8
         Left            =   480
         TabIndex        =   26
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label lblLabel 
         Caption         =   "req_timeout_send"
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
         Index           =   6
         Left            =   480
         TabIndex        =   25
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblLabel 
         Caption         =   "req_timeout_connect"
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
         Index           =   4
         Left            =   480
         TabIndex        =   24
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Frame fraTests 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CheckBox chkTestAttack 
         Caption         =   "scan_test_&attack*"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "This request may cause harm to the target service."
         Top             =   3120
         Value           =   1  'Checked
         Width           =   3000
      End
      Begin VB.CheckBox chkTestWrongprotocol 
         Caption         =   "scan_test_wrong&protocol"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Value           =   1  'Checked
         Width           =   3000
      End
      Begin VB.CheckBox chkTestNonexistingmethod 
         Caption         =   "scan_test_nonexisting&method"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Value           =   1  'Checked
         Width           =   3000
      End
      Begin VB.CheckBox chkTestWrontmethod 
         Caption         =   "scan_test_&wrongmethod"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Value           =   1  'Checked
         Width           =   3000
      End
      Begin VB.CheckBox chkTestOptions 
         Caption         =   "scan_test_&options"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Value           =   1  'Checked
         Width           =   3000
      End
      Begin VB.CheckBox chkTestHead 
         Caption         =   "scan_test_&head"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Value           =   1  'Checked
         Width           =   3000
      End
      Begin VB.CheckBox chkTestGetlong 
         Caption         =   "scan_test_get&long*"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "This request may cause harm to the target service."
         Top             =   960
         Value           =   1  'Checked
         Width           =   3000
      End
      Begin VB.CheckBox chkTestGetnonexisting 
         Caption         =   "scan_test_get&nonexisting"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Value           =   1  'Checked
         Width           =   3000
      End
      Begin VB.CheckBox chkTestGetexisting 
         Caption         =   "scan_test_&getexisting"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   3000
      End
      Begin VB.Image imgImage 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   4080
         Picture         =   "frmConfiguration.frx":114D
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblLabel 
         Caption         =   "The test for getting an existing file is required."
         Height          =   495
         Index           =   14
         Left            =   4440
         TabIndex        =   51
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblLabel 
         Caption         =   "As more test cases you activate, as higher the accuracy of the enumeration will be."
         Height          =   855
         Index           =   16
         Left            =   4440
         TabIndex        =   50
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label lblLabel 
         Caption         =   "The test cases flagged with an * might cause harm to the target service and might be detected easily by security systems."
         Height          =   1215
         Index           =   15
         Left            =   4440
         TabIndex        =   49
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.Frame fraAgent 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   240
      TabIndex        =   42
      Top             =   600
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CheckBox chkPreventRedirects 
         Caption         =   "req_agent_no&redirect"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   2535
      End
      Begin VB.ComboBox cboAgentName 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Text            =   "Mozilla/5.0 (Windows; U; Windows NT 5.2; en-GB; rv:1.9.0.5) Gecko/2008120122 Firefox/3.0.5"
         ToolTipText     =   "Mozilla/5.0 (Windows; U; Windows NT 5.2; en-GB; rv:1.9.0.5) Gecko/2008120122 Firefox/3.0.5"
         Top             =   960
         Width           =   6135
      End
      Begin VB.Label lblLabel 
         Caption         =   "&Prevent redirects (e.g. http response code 3xx)"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   70
         Top             =   2280
         Width           =   3735
      End
      Begin VB.Label lblLabel 
         Caption         =   $"frmConfiguration.frx":14FC
         Height          =   1455
         Index           =   3
         Left            =   4560
         TabIndex        =   68
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Image imgImage 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   16
         Left            =   4200
         Picture         =   "frmConfiguration.frx":1584
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label lblBrowserreconUserAgents 
         Alignment       =   1  'Right Justify
         Caption         =   "Lookup known user-agents at the browserrecon project"
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
         Left            =   120
         MouseIcon       =   "frmConfiguration.frx":198E
         MousePointer    =   99  'Custom
         TabIndex        =   66
         ToolTipText     =   "Show possible user-agents"
         Top             =   1440
         Width           =   6075
      End
      Begin VB.Image imgImage 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   15
         Left            =   120
         Picture         =   "frmConfiguration.frx":1C98
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblLabel 
         Caption         =   $"frmConfiguration.frx":2089
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   61
         Top             =   480
         Width           =   6135
      End
      Begin VB.Label lblLabel 
         Caption         =   "req_agent_name"
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
         Index           =   0
         Left            =   480
         TabIndex        =   43
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame fraLongrequest 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   240
      TabIndex        =   22
      Top             =   600
      Visible         =   0   'False
      Width           =   6375
      Begin VB.ComboBox cboLongrequestLength 
         Height          =   315
         Left            =   120
         TabIndex        =   65
         Text            =   "1024"
         ToolTipText     =   "1024"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtLongrequestChar 
         Height          =   285
         Left            =   120
         MaxLength       =   1
         TabIndex        =   67
         Text            =   "a"
         ToolTipText     =   "a"
         Top             =   2280
         Width           =   495
      End
      Begin VB.Image imgImage 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   120
         Picture         =   "frmConfiguration.frx":2134
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgImage 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   120
         Picture         =   "frmConfiguration.frx":24F0
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblLabel 
         Caption         =   "The definition of the length of the long request in bytes which is used in the according test case. Suggested value: 1024"
         Height          =   495
         Index           =   32
         Left            =   120
         TabIndex        =   60
         Top             =   480
         Width           =   6135
      End
      Begin VB.Label lblLabel 
         Caption         =   $"frmConfiguration.frx":28CE
         Height          =   495
         Index           =   34
         Left            =   120
         TabIndex        =   59
         Top             =   1800
         Width           =   6135
      End
      Begin VB.Label lblLabel 
         Caption         =   "req_longrequest_length"
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
         Index           =   31
         Left            =   480
         TabIndex        =   41
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblLabel 
         Caption         =   "req_longrequest_char"
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
         Index           =   33
         Left            =   480
         TabIndex        =   40
         Top             =   1560
         Width           =   2055
      End
   End
   Begin VB.Frame fraResources 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   240
      TabIndex        =   36
      Top             =   600
      Visible         =   0   'False
      Width           =   6375
      Begin VB.ComboBox cboResourcesAttack 
         Height          =   315
         Left            =   120
         TabIndex        =   64
         Text            =   "/etc/passwd?format=%%%%&xss=""><script>alert('xss');</script>&traversal=../../&sql=' OR 1;"
         ToolTipText     =   "/etc/passwd?format=%%%%&xss=""><script>alert('xss');</script>&traversal=../../&sql=' OR 1;"
         Top             =   3240
         Width           =   6135
      End
      Begin VB.ComboBox cboResourcesNotavailable 
         Height          =   315
         Left            =   120
         TabIndex        =   63
         Text            =   "/404test_.html"
         ToolTipText     =   "/404test_.html"
         Top             =   2040
         Width           =   6135
      End
      Begin VB.ComboBox cboResourcesAvailable 
         Height          =   315
         Left            =   120
         TabIndex        =   62
         Text            =   "/"
         ToolTipText     =   "/"
         Top             =   840
         Width           =   6135
      End
      Begin VB.Image imgImage 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   14
         Left            =   120
         Picture         =   "frmConfiguration.frx":2973
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgImage 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   13
         Left            =   120
         Picture         =   "frmConfiguration.frx":2D88
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgImage 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   12
         Left            =   120
         Picture         =   "frmConfiguration.frx":318E
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblLabel 
         Caption         =   "The definition of the resource which shall be used within all requests fetching an existing object. Suggested value: /"
         Height          =   495
         Index           =   26
         Left            =   120
         TabIndex        =   58
         Top             =   360
         Width           =   6135
      End
      Begin VB.Label lblLabel 
         Caption         =   $"frmConfiguration.frx":3593
         Height          =   495
         Index           =   28
         Left            =   120
         TabIndex        =   57
         Top             =   1560
         Width           =   6135
      End
      Begin VB.Label lblLabel 
         Caption         =   $"frmConfiguration.frx":3624
         Height          =   495
         Index           =   30
         Left            =   120
         TabIndex        =   56
         Top             =   2760
         Width           =   6135
      End
      Begin VB.Label lblLabel 
         Caption         =   "req_resource_attack"
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
         Index           =   29
         Left            =   480
         TabIndex        =   39
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label lblLabel 
         Caption         =   "req_resource_notavailable"
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
         Index           =   27
         Left            =   480
         TabIndex        =   38
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "req_resource_available"
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
         Index           =   25
         Left            =   480
         TabIndex        =   37
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.Frame fraProtocols 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   240
      TabIndex        =   33
      Top             =   600
      Visible         =   0   'False
      Width           =   6375
      Begin VB.ComboBox cboProtocolsWrong 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Text            =   "HTTP/9.8"
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox txtProtocolsLegitimate 
         Height          =   285
         Left            =   120
         MaxLength       =   128
         TabIndex        =   18
         Text            =   "HTTP/1.1"
         ToolTipText     =   "HTTP/1.1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.Image imgImage 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   11
         Left            =   120
         Picture         =   "frmConfiguration.frx":36C8
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgImage 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   10
         Left            =   120
         Picture         =   "frmConfiguration.frx":3ACA
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblLabel 
         Caption         =   $"frmConfiguration.frx":3EC4
         Height          =   495
         Index           =   22
         Left            =   120
         TabIndex        =   55
         Top             =   480
         Width           =   6135
      End
      Begin VB.Label lblLabel 
         Caption         =   "The http protocol version which shall be used within the test case for wrong protocol definitions. Suggested value: HTTP/9.8"
         Height          =   495
         Index           =   24
         Left            =   120
         TabIndex        =   54
         Top             =   1800
         Width           =   6135
      End
      Begin VB.Label lblLabel 
         Caption         =   "req_protocol_legitimate"
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
         Index           =   21
         Left            =   480
         TabIndex        =   35
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblLabel 
         Caption         =   "req_protocol_wrong"
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
         Index           =   23
         Left            =   480
         TabIndex        =   34
         Top             =   1560
         Width           =   2175
      End
   End
   Begin VB.Frame fraMethods 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   240
      TabIndex        =   30
      Top             =   600
      Visible         =   0   'False
      Width           =   6375
      Begin VB.ComboBox cboMethodsNotexisting 
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Text            =   "TEST"
         ToolTipText     =   "TEST"
         Top             =   2280
         Width           =   1695
      End
      Begin VB.ComboBox cboMethodsNotallowed 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Text            =   "DELETE"
         ToolTipText     =   "DELETE"
         Top             =   960
         Width           =   1695
      End
      Begin VB.Image imgImage 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   9
         Left            =   120
         Picture         =   "frmConfiguration.frx":3F55
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgImage 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   8
         Left            =   120
         Picture         =   "frmConfiguration.frx":435F
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblLabel 
         Caption         =   $"frmConfiguration.frx":4773
         Height          =   495
         Index           =   18
         Left            =   120
         TabIndex        =   53
         Top             =   480
         Width           =   6135
      End
      Begin VB.Label lblLabel 
         Caption         =   $"frmConfiguration.frx":4818
         Height          =   495
         Index           =   20
         Left            =   120
         TabIndex        =   52
         Top             =   1800
         Width           =   6135
      End
      Begin VB.Label lblLabel 
         Caption         =   "req_method_notexisting"
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
         Index           =   19
         Left            =   480
         TabIndex        =   32
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label lblLabel 
         Caption         =   "req_method_notallowed"
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
         Index           =   17
         Left            =   480
         TabIndex        =   31
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame fraStatistics 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   240
      TabIndex        =   27
      Top             =   600
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox txtStatisticsHitpointsmin 
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "7"
         ToolTipText     =   "7"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtStatisticsHitpointsmax 
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "14"
         ToolTipText     =   "14"
         Top             =   2280
         Width           =   615
      End
      Begin VB.Image imgImage 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   120
         Picture         =   "frmConfiguration.frx":48BC
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgImage 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   120
         Picture         =   "frmConfiguration.frx":4CC1
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblLabel 
         Caption         =   "Amount of minimum required hitpoints per test case to reach 100 % in the matches. Suggested value: 7"
         Height          =   495
         Index           =   11
         Left            =   120
         TabIndex        =   48
         Top             =   480
         Width           =   6135
      End
      Begin VB.Label lblLabel 
         Caption         =   "Amount of maximum possible hitpoints per test case to set the level of 100 % in the matches. Suggested value: 14"
         Height          =   495
         Index           =   13
         Left            =   120
         TabIndex        =   47
         Top             =   1800
         Width           =   6135
      End
      Begin VB.Label lblLabel 
         Caption         =   "app_hitpoints_minimum"
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
         Index           =   10
         Left            =   480
         TabIndex        =   29
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblLabel 
         Caption         =   "app_hitpoints_maximum"
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
         Index           =   12
         Left            =   480
         TabIndex        =   28
         Top             =   1560
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   3480
      Picture         =   "frmConfiguration.frx":50C8
      Style           =   1  'Graphical
      TabIndex        =   71
      ToolTipText     =   "Cancel Changes"
      Top             =   4440
      Width           =   1212
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   615
      Left            =   2160
      Picture         =   "frmConfiguration.frx":5460
      Style           =   1  'Graphical
      TabIndex        =   69
      ToolTipText     =   "Save Settings"
      Top             =   4440
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip tbsSettings 
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7435
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   8
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Timing"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Statistics"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tests"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Methods"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Protocols"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Resources"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Longrequest"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Agent"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboAgentName_Change()
    cboAgentName.Text = PreventEmptyInput(cboAgentName.Text, "Mozilla/5.0 (Windows; U; Windows NT 5.2; en-US; rv:1.8.1.11) Gecko/20071127 Firefox/2.0.0.11")
End Sub

Private Sub cboLongrequestLength_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub cboLongrequestLength_LostFocus()
    cboLongrequestLength.Text = AllowIntegersOnly(CLng(Val(cboLongrequestLength.Text)), 1, 65535, 1024)
End Sub

Private Sub cboMethodsNotallowed_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboMethodsNotallowed, KeyAscii, iLeftOff
End Sub

Private Sub cboMethodsNotallowed_LostFocus()
    cboMethodsNotallowed.Text = PreventEmptyInput(cboMethodsNotallowed.Text, "DELETE")
End Sub

Private Sub cboMethodsNotexisting_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboMethodsNotexisting, KeyAscii, iLeftOff
End Sub

Private Sub cboMethodsNotexisting_LostFocus()
    cboMethodsNotexisting.Text = PreventEmptyInput(cboMethodsNotexisting.Text, "TEST")
End Sub

Private Sub cboProtocolsWrong_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboProtocolsWrong, KeyAscii, iLeftOff
End Sub

Private Sub cboProtocolsWrong_LostFocus()
    cboProtocolsWrong.Text = PreventEmptyInput(cboProtocolsWrong.Text, "HTTP/9.8")
End Sub

Private Sub cboResourcesAttack_LostFocus()
    cboResourcesAttack.Text = PreventEmptyInput(cboResourcesAttack.Text, "/etc/passwd?format=%%%%&xss=""><script>alert('xss');</script>&traversal=../../&sql=' OR 1;")
End Sub

Private Sub cboResourcesAvailable_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboResourcesAvailable, KeyAscii, iLeftOff
End Sub

Private Sub cboResourcesAvailable_LostFocus()
    cboResourcesAvailable.Text = PreventEmptyInput(cboResourcesAvailable.Text, "/")
End Sub

Private Sub cboResourcesNotavailable_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboResourcesNotavailable, KeyAscii, iLeftOff
End Sub

Private Sub cboResourcesNotavailable_LostFocus()
    cboResourcesNotavailable.Text = PreventEmptyInput(cboResourcesNotavailable.Text, "/404test_.html")
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Call SaveConfiguration
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "Configuration - " & Replace(app_configuration_filename, App.Path, vbNullString, , , 1)
    
    Call ReadConfigTemplate(App.Path & "\config_templates\methods_not_allowed.contem", cboMethodsNotallowed)
    Call ReadConfigTemplate(App.Path & "\config_templates\methods_not_existing.contem", cboMethodsNotexisting)
    cboMethodsNotexisting.AddItem ChrW$(Rand(65, 90)) & ChrW$(Rand(65, 90)) & ChrW$(Rand(65, 90)) & ChrW$(Rand(65, 90)) & ChrW$(Rand(65, 90))
    Call ReadConfigTemplate(App.Path & "\config_templates\protocols_wrong.contem", cboProtocolsWrong)
    Call ReadConfigTemplate(App.Path & "\config_templates\resources_available.contem", cboResourcesAvailable)
    Call ReadConfigTemplate(App.Path & "\config_templates\resources_not_available.contem", cboResourcesNotavailable)
    Call ReadConfigTemplate(App.Path & "\config_templates\resources_attack_item.contem", cboResourcesAttack)
    Call ReadConfigTemplate(App.Path & "\config_templates\long_request_length.contem", cboLongrequestLength)
    Call ReadConfigTemplate(App.Path & "\config_templates\agent_names.contem", cboAgentName)
    
    Call FillConfiguration
End Sub

Private Sub lblBrowserreconUserAgents_Click()
    Call ChangeStatusBar("Open browserrecon project user-agents web site...")
    Call ShellExecute(frmMain.hwnd, "Open", "http://www.computec.ch/projekte/browserrecon/?s=database&t=&f=user-agent", "", App.Path, 1)
    Call ChangeStatusBarDone
End Sub

Private Sub tbsSettings_Click()
    Dim iIndex As Integer
    
    iIndex = tbsSettings.SelectedItem.Index

    fraTiming.Visible = False
    fraStatistics.Visible = False
    fraTests.Visible = False
    fraMethods.Visible = False
    fraProtocols.Visible = False
    fraResources.Visible = False
    fraLongrequest.Visible = False
    fraAgent.Visible = False

    If (iIndex = 1) Then
        fraTiming.Visible = True
    ElseIf (iIndex = 2) Then
        fraStatistics.Visible = True
    ElseIf (iIndex = 3) Then
        fraTests.Visible = True
    ElseIf (iIndex = 4) Then
        fraMethods.Visible = True
    ElseIf (iIndex = 5) Then
        fraProtocols.Visible = True
    ElseIf (iIndex = 6) Then
        fraResources.Visible = True
    ElseIf (iIndex = 7) Then
        fraLongrequest.Visible = True
    ElseIf (iIndex = 8) Then
        fraAgent.Visible = True
    End If
End Sub

Private Sub FillConfiguration()
    Call DisableElements

    txtTimingConnect.Text = req_timeout_connect
    txtTimingSend.Text = req_timeout_send
    txtTimingReceive.Text = req_timeout_receive
    
    txtStatisticsHitpointsmin.Text = app_hitpoints_minimum
    txtStatisticsHitpointsmax.Text = app_hitpoints_maximum
    
    'Call SetCheckboxesTest(scan_test_getexisting, chkTestGetexisting)
    Call SetCheckboxesTest(scan_test_getnonexisting, chkTestGetnonexisting)
    Call SetCheckboxesTest(scan_test_getlong, chkTestGetlong)
    Call SetCheckboxesTest(scan_test_head, chkTestHead)
    Call SetCheckboxesTest(scan_test_options, chkTestOptions)
    Call SetCheckboxesTest(scan_test_wrongmethod, chkTestWrontmethod)
    Call SetCheckboxesTest(scan_test_nonexistingmethod, chkTestNonexistingmethod)
    Call SetCheckboxesTest(scan_test_wrongprotocol, chkTestWrongprotocol)
    Call SetCheckboxesTest(scan_test_attack, chkTestAttack)
    
    cboMethodsNotallowed.Text = req_method_notallowed
    cboMethodsNotexisting.Text = req_method_notexisting
    
    txtProtocolsLegitimate.Text = req_protocol_legitimate
    cboProtocolsWrong.Text = req_protocol_wrong
    
    cboResourcesAvailable.Text = req_resource_available
    cboResourcesNotavailable.Text = req_resource_notavailable
    cboResourcesAttack.Text = req_resource_attack
    
    cboLongrequestLength.Text = req_longrequest_length
    txtLongrequestChar.Text = req_longrequest_char
    
    cboAgentName.Text = req_agent_name
    Call SetCheckboxesTest(req_agent_noredirect, chkPreventRedirects)
    
    Call EnableElements
End Sub

Private Sub SetCheckboxesTest(ByRef iValue As Integer, ByRef cCheckbox As CheckBox)
    If (iValue = 0) Then
        cCheckbox.Value = 0
    Else
        cCheckbox.Value = 1
    End If
End Sub

Private Function GetCheckboxesTest(ByRef cCheckbox As CheckBox) As Integer
    If (cCheckbox.Value = 0) Then
        GetCheckboxesTest = 0
    Else
        GetCheckboxesTest = 1
    End If
End Function

Private Sub SaveConfiguration()
    Call DisableElements

    req_timeout_connect = CInt(txtTimingConnect.Text)
    req_timeout_send = CInt(txtTimingSend.Text)
    req_timeout_receive = CInt(txtTimingReceive.Text)
    
    app_hitpoints_minimum = CInt(txtStatisticsHitpointsmin.Text)
    app_hitpoints_maximum = CInt(txtStatisticsHitpointsmax.Text)
    
    scan_test_getexisting = GetCheckboxesTest(chkTestGetexisting)
    scan_test_getnonexisting = GetCheckboxesTest(chkTestGetnonexisting)
    scan_test_getlong = GetCheckboxesTest(chkTestGetlong)
    scan_test_head = GetCheckboxesTest(chkTestHead)
    scan_test_options = GetCheckboxesTest(chkTestOptions)
    scan_test_wrongmethod = GetCheckboxesTest(chkTestWrontmethod)
    scan_test_nonexistingmethod = GetCheckboxesTest(chkTestNonexistingmethod)
    scan_test_wrongprotocol = GetCheckboxesTest(chkTestWrongprotocol)
    scan_test_attack = GetCheckboxesTest(chkTestAttack)
    
    req_method_notallowed = Trim(cboMethodsNotallowed.Text)
    req_method_notexisting = Trim(cboMethodsNotexisting.Text)
    
    req_protocol_legitimate = Trim(txtProtocolsLegitimate.Text)
    req_protocol_wrong = Trim(cboProtocolsWrong.Text)
    
    req_resource_available = cboResourcesAvailable.Text
    req_resource_notavailable = cboResourcesNotavailable.Text
    req_resource_attack = cboResourcesAttack.Text
    
    req_longrequest_length = CInt(cboLongrequestLength.Text)
    req_longrequest_char = txtLongrequestChar.Text
    
    req_agent_name = Trim(cboAgentName.Text)
    req_agent_noredirect = GetCheckboxesTest(chkPreventRedirects)
    
    Call WriteConfigurationToFile(app_configuration_filename)
    
    Call EnableElements
End Sub

Private Sub txtLongrequestChar_DblClick()
    txtLongrequestChar.Text = ChrW$(Rand(97, 122))
End Sub

Private Sub txtLongrequestChar_LostFocus()
    txtLongrequestChar.Text = PreventEmptyInput(txtLongrequestChar.Text, "a")
End Sub

Private Sub txtProtocolsLegitimate_LostFocus()
    txtProtocolsLegitimate.Text = PreventEmptyInput(txtProtocolsLegitimate.Text, "HTTP/1.1")
End Sub

Private Sub txtStatisticsHitpointsmax_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtStatisticsHitpointsmax_LostFocus()
    txtStatisticsHitpointsmax.Text = AllowIntegersOnly(CLng(Val(txtStatisticsHitpointsmax.Text)), 1, 99, 14)
End Sub

Private Sub txtStatisticsHitpointsmin_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtStatisticsHitpointsmin_LostFocus()
    txtStatisticsHitpointsmin.Text = AllowIntegersOnly(CLng(Val(txtStatisticsHitpointsmin.Text)), 1, 99, 7)
End Sub

Private Sub txtTimingConnect_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtTimingConnect_LostFocus()
    txtTimingConnect.Text = AllowIntegersOnly(CLng(Val(txtTimingConnect.Text)), 50, 30000, 5000)
End Sub

Private Sub txtTimingReceive_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtTimingReceive_LostFocus()
    txtTimingReceive.Text = AllowIntegersOnly(CLng(Val(txtTimingReceive.Text)), 50, 30000, 5000)
End Sub

Private Sub txtTimingSend_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtTimingSend_LostFocus()
    txtTimingSend.Text = AllowIntegersOnly(CLng(Val(txtTimingSend.Text)), 50, 30000, 5000)
End Sub

Private Sub ReadConfigTemplate(ByRef sConfigTemplate As String, ByRef cComboBox As ComboBox)
    Dim sFileContent As String
    Dim sTemplateArray() As String
    Dim iTemplateArrayCount As Integer
    Dim i As Long
    
    Call DisableElements
    
    sFileContent = ReadFile(sConfigTemplate)
    sTemplateArray = Split(sFileContent, vbCrLf, , vbBinaryCompare)
    iTemplateArrayCount = UBound(sTemplateArray)

    For i = 0 To iTemplateArrayCount
        cComboBox.AddItem sTemplateArray(i)
    Next i
    
    Call EnableElements
End Sub

Private Sub DisableElements()
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    
    Screen.MousePointer = vbHourglass
End Sub

Private Sub EnableElements()
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    
    Screen.MousePointer = vbNormal
End Sub
