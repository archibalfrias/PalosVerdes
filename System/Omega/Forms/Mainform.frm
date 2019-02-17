VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MainForm 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   Caption         =   "xxx"
   ClientHeight    =   8280
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   11280
   Icon            =   "Mainform.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMain 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6525
      Left            =   0
      ScaleHeight     =   6525
      ScaleWidth      =   3405
      TabIndex        =   5
      Top             =   1455
      Width           =   3400
      Begin MSComctlLib.ImageList ImageListMother 
         Left            =   1440
         Top             =   3960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mainform.frx":0CCA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mainform.frx":19A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mainform.frx":267E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mainform.frx":3358
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mainform.frx":4032
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mainform.frx":4D0C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   375
         Left            =   360
         TabIndex        =   23
         Top             =   4320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Load Purc"
         Height          =   375
         Left            =   1800
         TabIndex        =   22
         Top             =   3840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Load Inv"
         Height          =   375
         Left            =   1800
         TabIndex        =   21
         Top             =   3480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   495
         Left            =   1800
         TabIndex        =   20
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   4920
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   3600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox imgSplitter 
         BorderStyle     =   0  'None
         Height          =   6495
         Left            =   3310
         MousePointer    =   9  'Size W E
         ScaleHeight     =   6495
         ScaleWidth      =   105
         TabIndex        =   16
         Top             =   0
         Width           =   100
         Begin RPVGCC.b8LineVertical b8LineVertical1 
            Height          =   2175
            Left            =   10
            TabIndex        =   17
            Top             =   0
            Width           =   60
            _ExtentX        =   106
            _ExtentY        =   3836
         End
      End
      Begin VB.PictureBox picDayTime 
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   1200
         Left            =   0
         ScaleHeight     =   1200
         ScaleWidth      =   3315
         TabIndex        =   6
         Top             =   5280
         Width           =   3315
         Begin VB.PictureBox picDayTimeInside 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   1140
            Left            =   30
            ScaleHeight     =   1140
            ScaleWidth      =   3255
            TabIndex        =   7
            Top             =   30
            Width           =   3255
            Begin VB.Timer TimerDateTime 
               Interval        =   1000
               Left            =   2280
               Top             =   720
            End
            Begin VB.Timer TimerSeparator 
               Interval        =   500
               Left            =   2760
               Top             =   720
            End
            Begin VB.Label lblWeekDay 
               BackStyle       =   0  'Transparent
               Caption         =   "Thu"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   495
               Left            =   360
               TabIndex        =   14
               Top             =   120
               Width           =   615
            End
            Begin VB.Label lblDate 
               BackStyle       =   0  'Transparent
               Caption         =   "Sep 24, 2009"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   495
               Left            =   1200
               TabIndex        =   13
               Top             =   120
               Width           =   3375
            End
            Begin VB.Label lblHour 
               BackStyle       =   0  'Transparent
               Caption         =   "00"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   615
               Left            =   240
               TabIndex        =   12
               Top             =   435
               Width           =   855
            End
            Begin VB.Label lblMinute 
               BackStyle       =   0  'Transparent
               Caption         =   "00"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   615
               Left            =   1245
               TabIndex        =   11
               Top             =   435
               Width           =   735
            End
            Begin VB.Label lblAMPM 
               BackStyle       =   0  'Transparent
               Caption         =   "PM"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   615
               Left            =   2055
               TabIndex        =   10
               Top             =   435
               Width           =   855
            End
            Begin VB.Label lblSeparator 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   615
               Left            =   960
               TabIndex        =   9
               Top             =   435
               Width           =   255
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   ","
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   495
               Left            =   945
               TabIndex        =   8
               Top             =   120
               Width           =   135
            End
         End
      End
      Begin MSComctlLib.TreeView trView 
         Height          =   2775
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4895
         _Version        =   393217
         Indentation     =   294
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "ImageListMother2"
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Mainform.frx":59E6
      End
      Begin MSComctlLib.ImageList ImageListMother1 
         Left            =   120
         Top             =   3000
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   13
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mainform.frx":5D00
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mainform.frx":65DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mainform.frx":72B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mainform.frx":7F8E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mainform.frx":80E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mainform.frx":89C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mainform.frx":929C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mainform.frx":9B76
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mainform.frx":A450
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mainform.frx":AD2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mainform.frx":C6BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mainform.frx":E3C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mainform.frx":E818
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageListMother2 
         Left            =   960
         Top             =   3120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mainform.frx":10FCA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mainform.frx":118A4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FF8080&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   11220
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   11280
      Begin VB.Timer TimerLogIn 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   840
         Top             =   480
      End
      Begin VB.TextBox txtActiveForm 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   2775
      End
      Begin VB.Timer Timer_Text_Check 
         Interval        =   200
         Left            =   9240
         Top             =   120
      End
      Begin VB.Timer Timer_Text_Blink 
         Enabled         =   0   'False
         Interval        =   400
         Left            =   8760
         Top             =   120
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   8640
         ScaleHeight     =   255
         ScaleWidth      =   1695
         TabIndex        =   3
         Top             =   600
         Width           =   1695
         Begin VB.Image Image3 
            Height          =   195
            Left            =   960
            Picture         =   "Mainform.frx":135AE
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Image Image4 
            Height          =   195
            Left            =   1320
            Picture         =   "Mainform.frx":139F0
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Image Image1 
            Height          =   195
            Left            =   0
            Picture         =   "Mainform.frx":146BA
            Stretch         =   -1  'True
            Top             =   0
            Width           =   255
         End
         Begin VB.Image Image2 
            Height          =   195
            Left            =   0
            Picture         =   "Mainform.frx":14A44
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5880
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin VB.Timer Timer_CheckIdle 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   3480
         Top             =   120
      End
      Begin VB.Timer Timer_when_Idle 
         Interval        =   1000
         Left            =   3960
         Top             =   120
      End
      Begin VB.PictureBox picProgressBar 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   1095
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1440
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Timer TimerSplash 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   240
         Top             =   480
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   7980
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   441
            MinWidth        =   441
            Picture         =   "Mainform.frx":14DCE
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   1940
            MinWidth        =   1940
            Text            =   "LOGIN NAME:"
            TextSave        =   "LOGIN NAME:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   2469
            MinWidth        =   2469
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Text            =   "Server"
            TextSave        =   "Server"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Main"
      Begin VB.Menu mnuSBar1 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Main|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient}"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMainLogInOut 
         Caption         =   ""
      End
      Begin VB.Menu mnuMainBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "Change &Password"
      End
      Begin VB.Menu mnuMainConfigureSystem 
         Caption         =   "&Configure System"
      End
      Begin VB.Menu mnuBackupDatabase 
         Caption         =   "Backup Database"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUpdateMenu 
         Caption         =   "Update Menu"
      End
      Begin VB.Menu mnuUpdateEarnDedOT 
         Caption         =   "Update Earning / Deduction / Earning and Overtime Multiplier"
      End
      Begin VB.Menu mnuUpdateGovtTables 
         Caption         =   "Update Govt Tables"
         Begin VB.Menu mnuUpdateGovtTablesSSS1 
            Caption         =   "SSS Tables"
         End
         Begin VB.Menu mnuUpdateGovtTablesSSS 
            Caption         =   "SSS Tables"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuUpdateGovtTablesPHIC1 
            Caption         =   "PhilHealth Tables"
         End
         Begin VB.Menu mnuUpdateGovtTablesPHIC 
            Caption         =   "PhilHealth Tables"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuUpdateGovtTablesPagIbig 
            Caption         =   "PagIbig Tables"
         End
         Begin VB.Menu mnuUpdateGovtTablesTax1 
            Caption         =   "Tax Tables"
         End
         Begin VB.Menu mnuUpdateGovtTablesTax 
            Caption         =   "Tax Tables"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuMainBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLockedSystem 
         Caption         =   "&Locked System"
      End
      Begin VB.Menu mnuMainExitSystem 
         Caption         =   "&Exit System"
      End
   End
   Begin VB.Menu mnuPersonnelCompensation 
      Caption         =   "&Personnel / Compensation"
      Begin VB.Menu mnuSBar2 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Personnel / Compensation|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient}"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPersonnelCompensationInformation 
         Caption         =   "Information"
      End
      Begin VB.Menu mnuPersonnelCompensationAssignID 
         Caption         =   "Assign ID Number"
      End
      Begin VB.Menu mnuPersonnelCompensationBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPersonnelCompensationActionMemo 
         Caption         =   "Action Memo"
      End
      Begin VB.Menu mnuPersonnelCompensationDeactivationMemo 
         Caption         =   "Deactivation Memo"
      End
      Begin VB.Menu mnuPersonnelCompensationBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPersonnelCompensationPosition 
         Caption         =   "Position"
      End
      Begin VB.Menu mnuPersonnelCompensationEmploymentStatus 
         Caption         =   "Employment Status"
      End
      Begin VB.Menu mnuPersonnelCompensationBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPersonnelCompensationSSSTable 
         Caption         =   "SSS Table"
      End
      Begin VB.Menu mnuPersonnelCompensationPHICTable 
         Caption         =   "Phil Health Table"
      End
      Begin VB.Menu mnuPersonnelCompensationPagIbigTable 
         Caption         =   "Pag Ibig Table"
      End
      Begin VB.Menu mnuPersonnelCompensationTaxTable 
         Caption         =   "Tax Table"
      End
      Begin VB.Menu mnuPersonnelCompensationExemption 
         Caption         =   "Exemption"
      End
      Begin VB.Menu mnuPersonnelCompensationBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPersonnelCompensationLoans 
         Caption         =   "Loans"
      End
      Begin VB.Menu mnuPersonnelCompensationCompensation 
         Caption         =   "Compensation"
      End
      Begin VB.Menu mnuPersonnelCompensationBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPersonnelCompensationAbsentUndertime 
         Caption         =   "Absent / Undertime / Late Employee"
      End
      Begin VB.Menu mnuPersonnelCompensationServiceCharge 
         Caption         =   "Service Charge"
      End
      Begin VB.Menu mnuServiceChargeSummary 
         Caption         =   "Service Charge Summary"
      End
   End
   Begin VB.Menu mnuInventory 
      Caption         =   "&Inventory"
      Begin VB.Menu mnuSBar3 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Inventory|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient} "
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInventorySection 
         Caption         =   "S&ection"
      End
      Begin VB.Menu mnuInventoryClass 
         Caption         =   "Cl&assification"
      End
      Begin VB.Menu mnuInventorySupplierInfo 
         Caption         =   "S&upplier Information"
      End
      Begin VB.Menu mnuInventoryItemInfo 
         Caption         =   "Item Information"
      End
      Begin VB.Menu mnuInventoryBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInventoryFixedAsset 
         Caption         =   "&Fixed Asset"
      End
      Begin VB.Menu mnuInventoryBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInventoryPurchaseOrder 
         Caption         =   "&Purchase Order"
      End
      Begin VB.Menu mnuInventoryReceivingReport 
         Caption         =   "&Receiving Report"
      End
      Begin VB.Menu mnuInventoryBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInventoryStockTransfer 
         Caption         =   "&Stocks Transfer"
      End
      Begin VB.Menu mnuInventoryStockAdjustment 
         Caption         =   "Stocks &Adjustment"
      End
      Begin VB.Menu mnuInventoryBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInventoryStocksIssuance 
         Caption         =   "Stocks &Issuance"
      End
      Begin VB.Menu mnuInventoryBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInventoryMenuManagement 
         Caption         =   "Menu Management"
      End
   End
   Begin VB.Menu mnuGolfOperation 
      Caption         =   "&Golf Operation"
      Begin VB.Menu mnuSBar4 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Golf Operation|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient}"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGolfOperationMemberInfo 
         Caption         =   "&Members Information"
      End
      Begin VB.Menu mnuGolfOperationMemberIDNumber 
         Caption         =   "M&embers ID Number"
      End
      Begin VB.Menu mnuGolfOperationMemberActionMemo 
         Caption         =   "Member'&s Memo"
      End
      Begin VB.Menu mnuGolfOperationBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGolfOperationCorporateAccounts 
         Caption         =   "C&orporate Accounts"
      End
      Begin VB.Menu mnuCompanyAccount 
         Caption         =   "Com&pany Accounts"
      End
      Begin VB.Menu mnuGolfOperationBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGolfCartInfo 
         Caption         =   "GolfCart Information"
      End
      Begin VB.Menu mnuCaddyInformation 
         Caption         =   "Caddy Information"
      End
      Begin VB.Menu mnuGolfOperationBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGolfOperationBagTagAssign 
         Caption         =   "&Bag Tag Assignment"
      End
      Begin VB.Menu mnuGolfOperationPassport 
         Caption         =   "&Registration"
      End
      Begin VB.Menu mnuGolfOperationLockerRoom 
         Caption         =   "Locker Room"
      End
      Begin VB.Menu mnuGolfOperationProShop 
         Caption         =   "Pro Shop"
      End
      Begin VB.Menu mnuGolfOperationGolferLoungeTeeHouse 
         Caption         =   "Golfer's Lounge / Tee Houses"
      End
   End
   Begin VB.Menu mnuCashier 
      Caption         =   "&Cashier"
      Begin VB.Menu mnuCashierChargeInvoice 
         Caption         =   "&Charge Invoice"
      End
      Begin VB.Menu mnuCashierAR 
         Caption         =   "&Acknowledgement Receipt"
      End
      Begin VB.Menu mnuCashierOR 
         Caption         =   "&Official Receipt"
      End
   End
   Begin VB.Menu mnuAccounting 
      Caption         =   "&Accounting"
      Begin VB.Menu mnuAccountingDMCM 
         Caption         =   "Debit / Credit Memo"
      End
      Begin VB.Menu mnuAccountingPettyCash 
         Caption         =   "Petty Cash"
      End
      Begin VB.Menu mnuAccountingBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccountingCV 
         Caption         =   "Check Voucher"
      End
      Begin VB.Menu mnuAccountingJV 
         Caption         =   "Journal Voucher"
      End
      Begin VB.Menu mnuAccountingBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccountingChartofAccount 
         Caption         =   "Chart of Accounts"
      End
   End
   Begin VB.Menu mnuScoring 
      Caption         =   "&Scoring"
      Begin VB.Menu mnuSBar5 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Scoring|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient}"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHoleYardageParHandicap 
         Caption         =   "Hole / Yardage / Par / Handicap"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTournamentSetup 
         Caption         =   "Tournament Setup"
      End
      Begin VB.Menu mnuPlayerSetup 
         Caption         =   "Player Setup"
      End
      Begin VB.Menu mnuTeamSetup 
         Caption         =   "Team Setup"
      End
      Begin VB.Menu mnuPairing 
         Caption         =   "Pairing"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuScoringBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnterScores 
         Caption         =   "Enter Scores"
      End
      Begin VB.Menu mnuScoringReport 
         Caption         =   "Reports"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuScoringBar2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEvaluation 
         Caption         =   "Evaluation"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "&Utility"
      Begin VB.Menu mnuSBar6 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Utility|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient}"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUtilityAccessRights 
         Caption         =   "&User Rights"
      End
      Begin VB.Menu mnuMainCompanySetup 
         Caption         =   "C&ompany Setup"
      End
      Begin VB.Menu mnuUtilityLocation 
         Caption         =   "Location"
      End
      Begin VB.Menu mnuUtilityBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPersonnelCompensationDepartment 
         Caption         =   "Department"
      End
      Begin VB.Menu mnuUtilityBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPersonnelCompensationOvertimeRestdayRate 
         Caption         =   "Overtime / Restday Rate"
      End
      Begin VB.Menu mnuPersonnelCompensationGenPayrollPeriod 
         Caption         =   "Generate Payroll Period"
      End
      Begin VB.Menu mnuPersonnelCompensationServiceChargeSetup 
         Caption         =   "Service Charge Setup"
      End
      Begin VB.Menu mnuUtilityBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPersonnelAllowanceEntry 
         Caption         =   "Allowance Entry"
      End
      Begin VB.Menu mnuPersonnelAllowanceGeneration 
         Caption         =   "Allowance Browse"
      End
      Begin VB.Menu mnuUtilityBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShareIDNumber 
         Caption         =   "Share ID Number"
      End
      Begin VB.Menu mnuGolfOperationGreenFees 
         Caption         =   "Green &Fees"
      End
      Begin VB.Menu mnuGolfOperationMonthlyDues 
         Caption         =   "Monthly &Dues"
      End
      Begin VB.Menu mnuUtilityBar5 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuSBar7 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Utility|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient}"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSystem 
         Caption         =   "System Information"
      End
      Begin VB.Menu mnuToolsBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIM 
         Caption         =   "Instant Messaging"
      End
      Begin VB.Menu mnuQuotes 
         Caption         =   "Quotes"
      End
      Begin VB.Menu mnuMSTools 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalculator 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mnuNotepad 
         Caption         =   "Notepad"
      End
      Begin VB.Menu mnuWindowsExplorer 
         Caption         =   "Windows Explorer"
      End
      Begin VB.Menu mnuOnScreenKeyboard 
         Caption         =   "On-Screen Keyboard"
      End
      Begin VB.Menu mnuOfficeTools 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMSWord 
         Caption         =   "Microsoft Word"
      End
      Begin VB.Menu mnuMSExcel 
         Caption         =   "Microsoft Excel"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuSBar8 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient}"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuHelpTopics 
         Caption         =   "Help Topics"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cn As New ADODB.Connection

Private iPanelDrag As Integer
Private ResizeGate As Boolean
Dim client As RECT
Dim upperleft As POINTAPI
Dim tRC As RECT
    
Dim Form As Form
Dim Loaded As Boolean
    
Dim strTime, Array1, Array2, strNames, strUser, oNet, oFSO, sPath, sFileName, i, iPK, _
strPath, dDebit, dCredit, dBalance, sForms, strExcelPath, strWordPath

Dim iWorkSheet, strRange, sValue, TotRow, RowCnt

Private Sub BoxInCursor(bSet As Boolean)
'======================================================================
' This routine forces mouse in a set rectangle when resizing objects
'======================================================================
If bSet Then
    
    'Get information about our wndow
    GetClientRect hwnd, client
    upperleft.x = client.Left
    upperleft.y = client.Top
    'Convert window coördinates to screen coördinates
    ClientToScreen hwnd, upperleft
    'move our rectangle
    OffsetRect client, upperleft.x, upperleft.y
    InflateRect client, -2, -2
    Select Case iPanelDrag
    Case 1 ''''''
    
    Case 2 ' right side of main listing is being resized
        client.Left = client.Left + 120
        client.Right = client.Right - 320
    Case 3  ' updown  sidebar being resized
        client.Top = client.Top + 108
        client.Bottom = client.Bottom - 105
    End Select
    'limit the cursor movement
    ClipCursor client
Else
    ClipCursor ByVal 0&
End If
End Sub

Private Sub ShowObjectInStatusBar(ByVal bShowObject As Boolean)

    SendMessageAny Statusbar1.hwnd, SB_GETRECT, 3, tRC
    With tRC
        .Top = (.Top * Screen.TwipsPerPixelY)
        .Left = (.Left * Screen.TwipsPerPixelX)
        .Bottom = (.Bottom * Screen.TwipsPerPixelY) - .Top
        .Right = (.Right * Screen.TwipsPerPixelX) - .Left
    End With
    With Picture2

        SetParent .hwnd, Statusbar1.hwnd
        .Move tRC.Left + 40, tRC.Top + 30, tRC.Right - 80, tRC.Bottom - 80
        .Visible = True
    End With
End Sub

Private Sub Command1_Click()

'CommonDialog1.DialogTitle = "OPEN FILE"
'CommonDialog1.Filename = ""
'CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
'CommonDialog1.FilterIndex = 1
'CommonDialog1.ShowOpen
'strPath = CommonDialog1.Filename
'If Trim(strPath) = "" Then Exit Sub
'txtPath.Text = strPath
    
Screen.MousePointer = vbHourglass

ConnOmega.Execute "DELETE FROM tbl_GL_Accounts WHERE (Dept >=26) AND (Dept <=35)"

'Set cn = New ADODB.Connection
'cn.Provider = "Microsoft.Jet.OLEDB.4.0"
'cn.ConnectionString = _
'    "Data Source= " & Trim(txtPath.Text) & ";" & _
'    "Extended Properties=Excel 8.0;"
'cn.CursorLocation = adUseClient
'If cn.State = adStateOpen Then cn.Close
'cn.Open
'i = 0
'Set rs = New ADODB.Recordset
'If rs.State = adStateOpen Then rs.Close
'rs.Open "SELECT * FROM [GLBegBalance$] ", cn, adOpenDynamic, adLockOptimistic

s = "SELECT GLBalance.* " & _
    " FROM GLBalance"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    'If IsNull(rs!AccountName) = False Then
        i = i + 1
        ConnOmega.Execute "INSERT INTO tbl_GL_Accounts " & _
                          " (AccountCode, AccountName, Dept, withSL) " & _
                          " VALUES ('" & rs!AccountCode & "', " & _
                          " '" & FORMATSQL(rs!AccountName) & "', " & _
                          " " & rs!Dept & ", " & rs!withSL & ")"
        dBalance = CDbl(Format(IIf(IsNull(rs!Balance), 0, rs!Balance), "#,##0.00"))
        dDebit = 0: dCredit = 0
        If CDbl(dBalance) > 0 Then
            dDebit = CDbl(dBalance)
            dCredit = 0
        Else
            dDebit = 0
            dCredit = CDbl(dBalance) * -1
        End If
        If CDbl(IIf(IsNull(rs!Balance), 0, rs!Balance)) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_GL_Transaction " & _
                              " (GLCode, DocDate, DocNumber, Debit, Credit) " & _
                              " VALUES ('" & rs!AccountCode & "', '09/30/2012', " & _
                              " 'END_BAL'," & CDbl(dDebit) & ", " & _
                              " " & CDbl(dCredit) & ")"
        End If
    'End If
    rs.MoveNext
Wend
rs.Close

'If cn.State = adStateOpen Then cn.Close
Screen.MousePointer = vbDefault
MsgBox i
End Sub

Private Sub Command2_Click()
Screen.MousePointer = vbHourglass
's = "SELECT a_MAINTENANCE.* " & _
    " FROM a_MAINTENANCE"
s = "SELECT a_DRYGOODS.* " & _
    " FROM a_DRYGOODS"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    t = "SELECT tbl_Inv_Items.* " & _
        " FROM tbl_Inv_Items " & _
        " WHERE (ItemCode = '" & rs!ItemCode & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        ConnOmega.Execute "UPDATE tbl_Inv_Items " & _
                          " SET Unit = '" & rs!UOM_Inv & "'" & _
                          " WHERE (PK = " & rt!PK & ")"
        ConnOmega.Execute "INSERT INTO tbl_Inv_Items_Transaction " & _
                          " (ItemKey, Cleared, InOut, DocType, DocNumber, DocDate, Location, QuantityIn, Cost, NetCost) " & _
                          " VALUES (" & rt!PK & ", 0, 'I', 1, 'Beg_Inv_093012', '09/30/2012', " & rs!Location & ", " & _
                          " " & CDbl(Format(rs!Qty, "#,##0.00")) & ", " & CDbl(Format(rs!Cost, "#,##0.00")) & ", " & _
                          " " & CDbl(Format(rs!Cost, "#,##0.00")) & ")"
    End If
    rt.Close
    rs.MoveNext
Wend
rs.Close
Screen.MousePointer = vbDefault
End Sub

Private Sub Command4_Click()

CommonDialog1.DialogTitle = "OPEN FILE"
CommonDialog1.Filename = ""
CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowOpen
strPath = CommonDialog1.Filename
If Trim(strPath) = "" Then Exit Sub
txtPath.Text = strPath
    
Screen.MousePointer = vbHourglass

Set cn = New ADODB.Connection
cn.Provider = "Microsoft.Jet.OLEDB.4.0"
cn.ConnectionString = _
    "Data Source= " & Trim(txtPath.Text) & ";" & _
    "Extended Properties=Excel 8.0;"
cn.CursorLocation = adUseClient
If cn.State = adStateOpen Then cn.Close
cn.Open
Set rs = New ADODB.Recordset
If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT * FROM [December$] ", cn, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    If IsNull(rs!ItemCode) = False Then
        t = "SELECT PK " & _
            " FROM tbl_Inv_Items " & _
            " WHERE (ItemCode = '" & rs!ItemCode & "')"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Inv_Items_Transaction " & _
                              " (ItemKey, Cleared, InOut, DocType, DocNumber, Location, DocDate, QuantityIn, Cost, " & _
                              " NetCost, NetVAT) " & _
                              " VALUES (" & rt!PK & ", 0, 'I', 2, '" & rs!RRNumber & "', 2, " & _
                              " '" & FormatDateTime(DateSerial(2012, Month(rs!DocDate), Day(rs!DocDate)), vbShortDate) & "', " & _
                              " " & CDbl(Format(rs!Qty, "#0.00")) & ", " & CDbl(Format(rs!U_COST, "#0.00")) & ", " & _
                              " " & CDbl(Format(rs!U_COST, "#0.00")) & ", " & CDbl(Format(rs!NET_VAT, "#0.00")) & ")"
        End If
        rt.Close
    End If
    rs.MoveNext
Wend
rs.Close
Screen.MousePointer = vbDefault

MsgBox "Done!                   ", vbInformation

End Sub

Private Sub Command5_Click()
CREATE_MODIFIED_STABLE_FORD "tmp_" & gbl_UserName & "_Scoring_ModStableFord"
End Sub

Private Sub Image1_Click()
If Timer_Text_Blink.Enabled = True Then
    Image1.Visible = True
    Image2.Visible = False
    Image1.ZOrder 0
    
    
    
    t = "SELECT PK, Date_Time, Message, From_User, MsgType, " & _
        " Convert(datetime, Convert(char(6), Date_Time, 12), 102) as ActDate " & _
        " From tbl_InstantMessaging " & _
        " WHERE (Opened = 0) " & _
        " AND (To_User = '" & gbl_UserName & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        While Not rt.EOF
            If CDbl(rt!MsgType) = 0 Then
                
                Loaded = False
                strUser = rs!From_User
                
                sForms = ""
                For Each Form In Forms
                    If Trim(strUser) = Form.Caption Then
                        Loaded = True
                        Exit For
                    End If
                Next Form
                
                If Loaded = True Then
                    Form.ZOrder 0
                Else
                    If iMsgLoaded = 0 Then
                        Load frmInstantMessagingPM
                        frmInstantMessagingPM.Caption = strUser
                        frmInstantMessagingPM.lblTitle.Caption = strUser
                        frmInstantMessagingPM.Show
                    Else
                        Dim objForm As New frmInstantMessagingPM
                        objForm.Caption = strUser
                        objForm.lblTitle.Caption = strUser
                        objForm.Show
                    End If
                End If
                
            ElseIf CDbl(rt!MsgType) = 1 Then
                If DateValue(FormatDateTime(rt!ActDate, vbShortDate)) = DateValue(FormatDateTime(Date, vbShortDate)) Then
                    If IsLoaded(frmInstantMessaging) Then
                        frmInstantMessaging.ZOrder 0
                    Else
                        frmInstantMessaging.Show
                    End If
                End If
            End If
            rt.MoveNext
        Wend
    End If
    rt.Close
    
End If
End Sub

Private Sub Image2_Click()
If Timer_Text_Blink.Enabled = True Then
    Image1.Visible = True
    Image2.Visible = False
    Image1.ZOrder 0
    
    t = "SELECT PK, Date_Time, Message, From_User, MsgType, " & _
        " Convert(datetime, Convert(char(6), Date_Time, 12), 102) as ActDate " & _
        " From tbl_InstantMessaging " & _
        " WHERE (Opened = 0) " & _
        " AND (To_User = '" & gbl_UserName & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        While Not rt.EOF
            If CDbl(rt!MsgType) = 0 Then
                
                Loaded = False
                strUser = rs!From_User
                
                sForms = ""
                For Each Form In Forms
                    If Trim(strUser) = Form.Caption Then
                        Loaded = True
                        Exit For
                    End If
                Next Form
                
                If Loaded = True Then
                    Form.ZOrder 0
                Else
                    If iMsgLoaded = 0 Then
                        Load frmInstantMessagingPM
                        frmInstantMessagingPM.Caption = strUser
                        frmInstantMessagingPM.lblTitle.Caption = strUser
                        frmInstantMessagingPM.Show
                    Else
                        Dim objForm As New frmInstantMessagingPM
                        objForm.Caption = strUser
                        objForm.lblTitle.Caption = strUser
                        objForm.Show
                    End If
                End If
                
            ElseIf CDbl(rt!MsgType) = 1 Then
                If DateValue(FormatDateTime(rt!ActDate, vbShortDate)) = DateValue(FormatDateTime(Date, vbShortDate)) Then
                    If IsLoaded(frmInstantMessaging) Then
                        frmInstantMessaging.ZOrder 0
                    Else
                        frmInstantMessaging.Show
                    End If
                End If
            End If
            rt.MoveNext
        Wend
    End If
    rt.Close
    
End If
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> vbRightButton Then
    iPanelDrag = 2
    BoxInCursor True
    SetCapture imgSplitter.hwnd
End If
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If iPanelDrag = 2 Then
    picMain.Width = picMain.Width + x
    trView.Width = trView.Width + x
    picDayTime.Width = picDayTime.Width + x
    picDayTimeInside.Width = picDayTimeInside.Width + x
    imgSplitter.Left = picMain.Width - imgSplitter.Width
    frmBackground.Width = frmBackground.Width - x
End If
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
BoxInCursor False
ReleaseCapture
If iPanelDrag Then
    If iPanelDrag = 2 Then trView.SetFocus
    If iPanelDrag = 4 Then Me.SetFocus
    If iPanelDrag < 3 Then Me.SetFocus   'DoGradient picMain, 1
    iPanelDrag = 0
End If
ResizeGate = False
End Sub

Private Sub MDIForm_Load()
Me.Caption = App.ProductName

mnuPersonnelCompensation.Visible = False
mnuInventory.Visible = False
mnuGolfOperation.Visible = False
mnuCashier.Visible = False
mnuAccounting.Visible = False
mnuScoring.Visible = False
mnuUtility.Visible = False

Select Case Weekday(Now, vbMonday)
    Case 1
        lblWeekDay.ForeColor = &HFF00&
        lblWeekDay.Caption = "Mon" '"Monday"
    Case 2
        lblWeekDay.ForeColor = &HFF00&
        lblWeekDay.Caption = "Tue" '"Tuesday"
    Case 3
        lblWeekDay.ForeColor = &HFF00&
        lblWeekDay.Caption = "Wed" '"Wednesday"
    Case 4
        lblWeekDay.ForeColor = &HFF00&
        lblWeekDay.Caption = "Thu" '"Thursday"
    Case 5
        lblWeekDay.ForeColor = &HFF00&
        lblWeekDay.Caption = "Fri" '"Friday"
    Case 6
        lblWeekDay.ForeColor = &HFF00&
        lblWeekDay.Caption = "Sat" '"Saturday"
    Case 7
        lblWeekDay.ForeColor = &HFF&
        lblWeekDay.Caption = "Sun" '"Sunday"
End Select
lblDate.Caption = Format(Now, "mmm dd, yyyy") 'Format(Now, "mmmm dd, yyyy")
strTime = Format(Time, "hh:mm:ss AM/PM")
Array1 = Split(strTime, ":", -1, 1)
lblHour.Caption = Array1(0)
lblMinute.Caption = Array1(1)
Array2 = Split(Array1(2), " ", -1, 1)
lblAMPM.Caption = Array2(1)

LOAD_HIDE_MENU False

LOAD_TreeView

'mnuSBar1.Visible = True
'mnuSBar2.Visible = True
'mnuSBar3.Visible = True
'mnuSBar4.Visible = True
'mnuSBar5.Visible = True
'mnuSBar6.Visible = True
'mnuSBar7.Visible = True
'mnuSBar8.Visible = True

'If mnuSBar1.Visible = True Then
'    mnuMSTools.Caption = "-Microsoft Tools"
'    mnuOfficeTools.Caption = "-MS Office Tools"
'    SetMenus hWnd, ImageList1
'Else
    mnuMSTools.Caption = "-"
    mnuOfficeTools.Caption = "-"
'End If

frmBackground.Show
frmBackground.Top = 0
frmBackground.Left = 0

If PassStartWizard = 1 Then
    Unload frmConnectionWizard
Else
    Unload frmSplash
End If

'Timer_Text_Check.Enabled = True

mnuMainLogInOut.Caption = "&Log In"
mnuLockedSystem.Enabled = False
ShowProgressInStatusBar True
ShowObjectInStatusBar True
iMsgLoaded = 0

'MsgBox gbl_Server
'StatusBar1.Panels(6).Text = gbl_Server

MainForm.Statusbar1.Panels(7).Alignment = sbrCenter
MainForm.Statusbar1.Panels(7).Text = UCase(gbl_Server)

TimerLogIn.Enabled = True

End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
With Statusbar1
    .Panels(1).Width = 250.01
    .Panels(2).Width = 1099.84
    .Panels(3).Width = 2000.12
    .Panels(4).Width = 380
    .Panels(6).Width = 3000
    .Panels(7).Width = 2000
    .Panels(5).Width = Me.Width - (.Panels(1).Width + .Panels(2).Width + _
                                  .Panels(3).Width + .Panels(4).Width + _
                                  .Panels(6).Width + .Panels(7).Width + 250)
    '.Panels(6).Text = gbl_Server
End With

'picMain.Width = 3430
'imgSplitter.Height = picMain.ScaleHeight
'b8LineVertical1.Height = imgSplitter.Height
'imgSplitter.Left = picMain.Width - 100
'trView.Top = 80
'trView.Left = 80
'trView.Height = picMain.ScaleHeight - picDayTime.Height - 130
'trView.Width = picMain.ScaleWidth - imgSplitter.Width - 80
'
'picDayTime.Top = picMain.ScaleHeight - picDayTime.Height
'picDayTime.Width = picMain.Width - 80
'picDayTimeInside.Width = picDayTime.Width - 60


frmBackground.Top = 0
frmBackground.Left = 0
frmBackground.Height = Me.ScaleHeight
frmBackground.Width = Me.ScaleWidth
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If Trim(Check_Open_Forms) <> "" Then MsgBox "Please Close All Opened Forms!            ", vbInformation, "Close Forms": frmBackground.ZOrder 1: Cancel = -1: Exit Sub
ConnOmega.Execute "UPDATE tbl_Users_Account SET Online = 0 WHERE (UserName = '" & FORMATSQL(CStr(gbl_UserName)) & "')"
'If mnuSBar1.Visible = True Then
'    ReleaseMenus hWnd
'End If
End
End Sub

Private Sub mnuAccountingCV_Click()
'LOAD_FORM "Check Voucher", "Open", frmCheckVoucher, 0
End Sub

Private Sub mnuBackupDatabase_Click()
'If StatusBar1.Panels(3).Text = "" Then Exit Sub
'If AccessRights("Allow Backup", "Backup") = False Then
'    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
'           "ACCESS DENIED!                                      ", vbCritical, "Alert"
'    Exit Sub
'End If
'
'If MsgBox("Continue SQL Backup?                         ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
'
'StatusBar1.Panels(4).Text = "SQL Backup on progress . . . . "
'
'Create_Backup
'
'StatusBar1.Panels(4).Text = ""
'Exit Sub
'PG:
'Exit Sub
End Sub

Private Sub mnuCaddyInformation_Click()
LOAD_FORM "Caddy Information", "Open", frmCaddyInformation, 0
End Sub

Private Sub mnuCalculator_Click()
If Statusbar1.Panels(3).Text = "" Then Exit Sub
On Error GoTo PG:
Shell "calc.exe", vbNormalFocus
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub mnuChangePassword_Click()
If gbl_UserName = "" Then Exit Sub
LOAD_FORM "", "", bChangePassword, 0
End Sub

Private Sub mnuCompanyAccount_Click()
LOAD_FORM "Corporate Account", "Open", frmMembershipCompany, 0
End Sub

Private Sub mnuEnterScores_Click()
Select Case ScoringType
    Case 1:  If WithIndividualPlay = 0 And WithTeamPlay = 1 Then LOAD_FORM "Scoring Score Card", "Open", frmScoreCardTeamOnly, 0 Else LOAD_FORM "Scoring Score Card", "Open", frmScoreCard, 0
    Case 2:  LOAD_FORM "Scoring Score Card", "Open", frmScoreCardModifiedStableFord, 0
    Case 3:  LOAD_FORM "Scoring Score Card", "Open", frmScoreCardsSystem36, 0
    Case 4:  MsgBox "Not Activated!                     ", vbCritical, "Sorry"
    Case 5:  LOAD_FORM "Scoring Score Card", "Open", frmScoreCardModifiedMolave, 0
End Select
If TournamentKey = 0 Then MsgBox "Please Set One Tournament!               ", vbCritical, "Error...": Exit Sub

'If ScoringType = 3 Then
'    LOAD_FORM "Scoring Score Card", "Open", frmScoreCardsSystem36, 0
'ElseIf ScoringType = 2 Then
'    LOAD_FORM "Scoring Score Card", "Open", frmScoreCardModifiedStableFord, 0
'Else
'    If WithIndividualPlay = 0 And WithTeamPlay = 1 Then
'        LOAD_FORM "Scoring Score Card", "Open", frmScoreCardTeamOnly, 0
'    Else
'        LOAD_FORM "Scoring Score Card", "Open", frmScoreCard, 0
'    End If
'End If
'If TournamentKey = 0 Then MsgBox "Please Set One Tournament!               ", vbCritical, "Error...": Exit Sub
End Sub

Private Sub mnuGolfCartInfo_Click()
'"Golf Cart Information"
'If IsLoaded(frmGolfCartInformation) Then frmGolfCartInformation.ZOrder 0 Else frmGolfCartInformation.Show
LOAD_FORM "Golf Cart Information", "Open", frmGolfCartInformation, 0
End Sub

Private Sub mnuGolfOperationCorporateAccounts_Click()
LOAD_FORM "Corporate Account", "Open", frmMembershipCorporate, 0
End Sub

Private Sub mnuGolfOperationMemberActionMemo_Click()
LOAD_FORM "Membership Action", "Open", frmMembershipAction, 0
End Sub

Private Sub mnuGolfOperationMemberIDNumber_Click()
LOAD_FORM "Membership ID Number", "Open", frmMembershipIDNumber, 0
End Sub

Private Sub mnuGolfOperationMemberInfo_Click()
LOAD_FORM "Membership Information", "Open", frmMembershipInformation, 0
'If IsLoaded(frmMembershipInformation) Then frmMembershipInformation.ZOrder 0 Else frmMembershipInformation.Show
End Sub

Private Sub mnuHoleYardageParHandicap_Click()
If IsLoaded(frmHoleYardageParHandicap) Then frmHoleYardageParHandicap.ZOrder 0 Else frmHoleYardageParHandicap.Show
End Sub

Private Sub mnuIM_Click()
LOAD_FORM "", "", frmInstantMessaging, 0
End Sub

Private Sub mnuInventoryClass_Click()
LOAD_FORM "Inventory Classification", "Open", frmInvClass, 0
End Sub

Private Sub mnuInventoryFixedAsset_Click()
LOAD_FORM "Fixed Assets", "Open", frmFAItems, 0
End Sub

Private Sub mnuInventoryItemInfo_Click()
LOAD_FORM "Inventory Items", "Open", frmInvItems, 0
End Sub

Private Sub mnuInventoryMenuManagement_Click()
LOAD_FORM "Menu Management", "Open", frmMenuMngt, 0
End Sub

Private Sub mnuInventoryPurchaseOrder_Click()
LOAD_FORM "Purchase Order", "Open", frmInvPO, 0
End Sub

Private Sub mnuInventoryReceivingReport_Click()
LOAD_FORM "Receiving Report", "Open", frmInvRR, 0
End Sub

Private Sub mnuInventorySection_Click()
LOAD_FORM "Inventory Section", "Open", frmInvSection, 0
End Sub

Private Sub mnuInventoryStockAdjustment_Click()
LOAD_FORM "Stock Adjustment", "Open", frmInvStockAdjustment, 0
End Sub

Private Sub mnuInventoryStocksIssuance_Click()
LOAD_FORM "Stock Issuance", "Open", frmInvStockIssuance, 0
End Sub

Private Sub mnuInventoryStockTransfer_Click()
'frmInvStockTransfer.Show
LOAD_FORM "Stock Transfer", "Open", frmInvStockTransfer, 0
End Sub

Private Sub mnuInventorySupplierInfo_Click()
LOAD_FORM "Inventory Supplier", "Open", frmInvSupplier, 0
End Sub

Private Sub mnuLockedSystem_Click()
frmSystemLocked.Show 1
End Sub

Private Sub mnuMainCompanySetup_Click()
LOAD_FORM "Company Information", "Open", frmCompany, 0
End Sub

Private Sub mnuMainConfigureSystem_Click()
LOAD_FORM "", "", frmSystemConfig, 1
End Sub

Private Sub mnuMainExitSystem_Click()
If Trim(Check_Open_Forms) <> "" Then MsgBox "Please Close All Opened Forms!            ", vbInformation, "Close Forms": frmBackground.ZOrder 1: Exit Sub
If MsgBox("Terminate Program!          ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbYes Then End
End Sub

Private Sub mnuMainLogInOut_Click()
If Trim(mnuMainLogInOut.Caption) = "Log &Out" Then
    If Trim(Check_Open_Forms) <> "" Then MsgBox "Please Close All Opened Forms!            ", vbInformation, "Close Forms": frmBackground.ZOrder 1: Exit Sub
    ConnOmega.Execute "UPDATE tbl_Users_Account SET Online = 0 WHERE (UserName = '" & FORMATSQL(CStr(gbl_UserName)) & "')"
    gbl_UserName = ""
    gbl_Password = ""
    gbl_CompleteName = ""
    gbl_LockWhenIdle = 0
    gbl_Idle_Time = 0
    gbl_Slides_Background = 0
    gbl_Slides_Time = 0
    gbl_Quotes_Time = 360
    Statusbar1.Panels(3).Text = ""
    mnuLockedSystem.Enabled = False
    mnuMainLogInOut.Caption = "Log &In"
    LOAD_HIDE_MENU False
    Set gbl_FORM = Nothing
Else
    LogInWithOutLoading = 1
    gbl_MODULE = ""
    aLogIn.Show 1
End If
End Sub

Private Sub mnuMSExcel_Click()
If Statusbar1.Panels(3).Text = "" Then Exit Sub
strExcelPath = GetSetting(App.EXEName, "ExcelPath", "ExcelP", "")
If Trim(strExcelPath) <> "" Then
    On Error GoTo AG:
    Shell strExcelPath, vbNormalFocus
Else
AG:
    CommonDialog1.DialogTitle = "Select MS Excel Application"
    CommonDialog1.InitDir = "C:\Program Files\"
    CommonDialog1.Filter = "Application(*.exe)|*.exe"
    CommonDialog1.ShowOpen
    strExcelPath = Trim(CommonDialog1.Filename)
    If Trim(strExcelPath) = "" Then Exit Sub
    Shell strExcelPath, vbNormalFocus
    SaveSetting App.EXEName, "ExcelPath", "ExcelP", strExcelPath
End If
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub mnuMSWord_Click()
If Statusbar1.Panels(3).Text = "" Then Exit Sub
strWordPath = GetSetting(App.EXEName, "WordPath", "WordP", "")
If Trim(strWordPath) <> "" Then
    On Error GoTo AG:
    Shell strWordPath, vbNormalFocus
Else
AG:
    CommonDialog1.DialogTitle = "Select MS Word Application"
    CommonDialog1.InitDir = "C:\Program Files\"
    CommonDialog1.Filter = "Application(*.exe)|*.exe"
    CommonDialog1.ShowOpen
    strWordPath = Trim(CommonDialog1.Filename)
    If Trim(strWordPath) = "" Then Exit Sub
    Shell strWordPath, vbNormalFocus
    SaveSetting App.EXEName, "WordPath", "WordP", strWordPath
End If
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub mnuNotepad_Click()
If Statusbar1.Panels(3).Text = "" Then Exit Sub
On Error GoTo PG:
Shell "notepad.exe", vbNormalFocus
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub mnuOnScreenKeyboard_Click()
If Statusbar1.Panels(3).Text = "" Then Exit Sub
'On Error GoTo PG:
'Shell "osk", vbNormalFocus
'Exit Sub
'PG:
'MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
'Exit Sub
End Sub

Private Sub mnuPersonnelAllowanceEntry_Click()
LOAD_FORM "Allowance", "Open", frmPersonnelAllowance, 0
End Sub

Private Sub mnuPersonnelAllowanceGeneration_Click()
LOAD_FORM "Allowance", "Generate", frmPersonnelAllowanceBrowse, 0
End Sub

Private Sub mnuPersonnelCompensationAbsentUndertime_Click()
LOAD_FORM "Absent/Late/Undertime Employee", "Open", frmAbsentUndertimeEmployee, 0
End Sub

Private Sub mnuPersonnelCompensationActionMemo_Click()
LOAD_FORM "Personnel Action Memo", "Open", frmPersonnelAction, 0
End Sub

Private Sub mnuPersonnelCompensationAssignID_Click()
LOAD_FORM "Personnel ID Number", "Open", frmPersonnelIDNumber, 0
End Sub

Private Sub mnuPersonnelCompensationCompensation_Click()
LOAD_FORM "Personnel Compensation", "Open", frmPersonnelCompensation, 0
End Sub

Private Sub mnuPersonnelCompensationDeactivationMemo_Click()
LOAD_FORM "Personnel Action Memo", "Open", frmPersonnelDeactivation, 0
End Sub

Private Sub mnuPersonnelCompensationDepartment_Click()
LOAD_FORM "Personnel Department", "Open", frmPersonnelDept, 0
End Sub

Private Sub mnuPersonnelCompensationEmploymentStatus_Click()
LOAD_FORM "Personnel Employment Status", "Open", frmPersonnelEmploymentStatus, 0
End Sub

Private Sub mnuPersonnelCompensationExemption_Click()
LOAD_FORM "Personnel Gov't Table", "PERSONAL_EXEMP", frmTaxExemption, 0
End Sub

Private Sub mnuPersonnelCompensationInformation_Click()
LOAD_FORM "Personnel Information", "Open", frmPersonnelInformation, 0
'LOAD_FORM "Personnel Information", "Open", frmPersonnelInfo, 0
'If IsLoaded(frmPersonnelInformation) Then frmPersonnelInformation.ZOrder 0 Else frmPersonnelInformation.Show
End Sub

Private Sub mnuPersonnelCompensationLoans_Click()
LOAD_FORM "Personnel Loans", "Open", frmPersonnelLoans, 0
End Sub

Private Sub mnuPersonnelCompensationOvertimeRestdayRate_Click()
LOAD_FORM "Personnel Overtime/Restday Rate", "Open", frmPersonnelOTRDRate, 0
End Sub

Private Sub mnuPersonnelCompensationPagIbigTable_Click()
LOAD_FORM "Personnel Gov't Table", "PAGIBIG", frmPagIbigTable, 0
End Sub

Private Sub mnuPersonnelCompensationPHICTable_Click()
LOAD_FORM "Personnel Gov't Table", "PHIC", frmPhilHealthTable, 0
End Sub

Private Sub mnuPersonnelCompensationPosition_Click()
'LOAD_FORM "Personnel Position", "Open", frmPersonnelPost, 0
End Sub

Private Sub mnuPersonnelCompensationServiceCharge_Click()
'If IsLoaded(frmServiceCharge) Then frmServiceCharge.ZOrder 0 Else frmServiceCharge.Show
LOAD_FORM "Service Charge", "Open", frmServiceCharge, 0
End Sub

Private Sub mnuPersonnelCompensationServiceChargeSetup_Click()
'If IsLoaded(frmServiceChargeSetup) Then frmServiceChargeSetup.ZOrder 0 Else frmServiceChargeSetup.Show
LOAD_FORM "Service Charge Setup", "Open", frmServiceChargeSetup, 0
End Sub

Private Sub mnuPersonnelCompensationSSSTable_Click()
LOAD_FORM "Personnel Gov't Table", "SSS", frmSSSTable, 0
End Sub

Private Sub mnuPersonnelCompensationTaxTable_Click()
LOAD_FORM "Personnel Gov't Table", "TAX", frmTaxTable, 0
End Sub

Private Sub mnuPlayerSetup_Click()
LOAD_FORM "Scoring Player Information", "Open", frmPlayerSetup, 0
If AccessRights("Scoring Player Information", "Open") = False Then Exit Sub

If TournamentKey = 0 Then MsgBox "Please Set One Tournament!               ", vbCritical, "Error...": Exit Sub


End Sub

Private Sub mnuQuotes_Click()
LOAD_FORM "", "", frmQuotes, 0
End Sub

Private Sub mnuServiceChargeSummary_Click()
'If IsLoaded(frmServiceChargeSummary) Then frmServiceChargeSummary.ZOrder 0 Else frmServiceChargeSummary.Show
LOAD_FORM "Service Charge Summary", "Open", frmServiceChargeSummary, 0
End Sub

Private Sub mnuShareIDNumber_Click()
LOAD_FORM "Share ID Number", "Open", frmMembershipShareID, 0
End Sub

Private Sub mnuSystem_Click()
If Statusbar1.Panels(3).Text = "" Then Exit Sub
Call StartSysInfo
End Sub

Private Sub mnuTeamSetup_Click()
LOAD_FORM "Scoring Team Information", "Open", frmTeamSetup, 0
If AccessRights("Scoring Team Information", "Open") = False Then Exit Sub

If TournamentKey = 0 Then MsgBox "Please Set One Tournament!               ", vbCritical, "Error...": Exit Sub

End Sub

Private Sub mnuTournamentSetup_Click()
LOAD_FORM "Scoring Tournament Information", "Open", frmTournamentSetup, 0
End Sub

Private Sub mnuUpdateEarnDedOT_Click()

If MsgBox("CONTINUE UPDATE EARNINGS / DEDUCTIONS / OVERTIME ACCOUNTS?                             ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub

MainForm.CommonDialog1.CancelError = True
On Error GoTo Err
MainForm.CommonDialog1.DialogTitle = "Open"
MainForm.CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
MainForm.CommonDialog1.ShowOpen
sPath = MainForm.CommonDialog1.Filename

Screen.MousePointer = vbHourglass

On Error GoTo PG:
Set xlsApp = CreateObject("Excel.Application")
With xlsApp
    .Workbooks.Open (sPath)
    .Visible = False
    .DisplayAlerts = False
    iWorkSheet = 1  'Earnings
    .Workbooks(1).Sheets(iWorkSheet).Activate
    For i = 2 To .Workbooks(1).Sheets(iWorkSheet).UsedRange.Rows.Count
        strRange = EXCEL_RANGE(1, i)
        sValue = .Range(strRange).Value
        If CDbl(sValue) = 0 Then Exit For
        s = "SELECT tbl_Personnel_Payroll_Earnings_Table.* " & _
            " FROM tbl_Personnel_Payroll_Earnings_Table " & _
            " WHERE (PK = " & .Range(EXCEL_RANGE(1, i)).Value & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Earnings_Table " & _
                              " (PK, Description, Abbvt, Sorting, RefRate, Tax, SSS, PHIC, PagIbig, " & _
                              " ViewInActionModule, HoursDesc, ViewInHours, DedtoRef, Month13, LastModified) " & _
                              " VALUES (" & .Range(EXCEL_RANGE(1, i)).Value & ", " & _
                              " '" & FORMATSQL(CStr(.Range(EXCEL_RANGE(2, i)).Value)) & "', " & _
                              " '" & FORMATSQL(CStr(.Range(EXCEL_RANGE(3, i)).Value)) & "', " & _
                              " " & .Range(EXCEL_RANGE(4, i)).Value & ", " & _
                              " " & .Range(EXCEL_RANGE(5, i)).Value & ",  " & _
                              " " & .Range(EXCEL_RANGE(6, i)).Value & ",  " & _
                              " " & .Range(EXCEL_RANGE(7, i)).Value & ",  " & _
                              " " & .Range(EXCEL_RANGE(8, i)).Value & ",  " & _
                              " " & .Range(EXCEL_RANGE(9, i)).Value & ",  " & _
                              " '" & FORMATSQL(CStr(.Range(EXCEL_RANGE(10, i)).Value)) & "',  " & _
                              " " & .Range(EXCEL_RANGE(11, i)).Value & ", " & _
                              " " & .Range(EXCEL_RANGE(12, i)).Value & ", " & _
                              " " & .Range(EXCEL_RANGE(13, i)).Value & ", " & _
                              " " & .Range(EXCEL_RANGE(14, i)).Value & ", " & _
                              " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
        Else
            ConnOmega.Execute "UPDATE tbl_Personnel_Payroll_Earnings_Table " & _
                              " SET Description = '" & FORMATSQL(CStr(.Range(EXCEL_RANGE(2, i)).Value)) & "', " & _
                              " Abbvt = '" & FORMATSQL(CStr(.Range(EXCEL_RANGE(3, i)).Value)) & "', " & _
                              " Sorting = " & .Range(EXCEL_RANGE(4, i)).Value & ", " & _
                              " RefRate = " & .Range(EXCEL_RANGE(5, i)).Value & ", " & _
                              " Tax = " & .Range(EXCEL_RANGE(6, i)).Value & ", " & _
                              " SSS = " & .Range(EXCEL_RANGE(7, i)).Value & ", " & _
                              " PHIC = " & .Range(EXCEL_RANGE(8, i)).Value & ", " & _
                              " PagIbig = " & .Range(EXCEL_RANGE(9, i)).Value & ", " & _
                              " ViewInActionModule = " & .Range(EXCEL_RANGE(10, i)).Value & ", " & _
                              " HoursDesc = '" & FORMATSQL(CStr(.Range(EXCEL_RANGE(11, i)).Value)) & "', " & _
                              " ViewInHours = " & .Range(EXCEL_RANGE(12, i)).Value & ", " & _
                              " DedtoRef = " & .Range(EXCEL_RANGE(13, i)).Value & ", " & _
                              " Month13 = " & .Range(EXCEL_RANGE(14, i)).Value & ", " & _
                              " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                              " WHERE (PK = " & .Range(EXCEL_RANGE(1, i)).Value & ")"
        End If
        rs.Close
    Next i
    
    iWorkSheet = 2  'Deductions
    .Workbooks(1).Sheets(iWorkSheet).Activate
    For i = 2 To .Workbooks(1).Sheets(iWorkSheet).UsedRange.Rows.Count
        strRange = EXCEL_RANGE(1, i)
        sValue = .Range(strRange).Value
        If CDbl(sValue) = 0 Then Exit For
        s = "SELECT tbl_Personnel_Payroll_Deductions_Table.* " & _
            " FROM tbl_Personnel_Payroll_Deductions_Table " & _
            " WHERE (PK = " & .Range(EXCEL_RANGE(1, i)).Value & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Deductions_Table " & _
                              " (PK, Description, Abbvt, Sorting, ViewInDeductionModule, DedSched, EmployerShare, " & _
                              " GovtDed, GovtDedEmpr, GovtDedMain, RefAccnt, WithSL, FixDed, LastModified) " & _
                              " VALUES (" & .Range(EXCEL_RANGE(1, i)).Value & ", " & _
                              " '" & FORMATSQL(.Range(EXCEL_RANGE(2, i)).Value) & "', " & _
                              " '" & FORMATSQL(.Range(EXCEL_RANGE(3, i)).Value) & "', " & _
                              " " & .Range(EXCEL_RANGE(4, i)).Value & ", " & _
                              " " & .Range(EXCEL_RANGE(5, i)).Value & ", " & _
                              " " & .Range(EXCEL_RANGE(6, i)).Value & ",  " & _
                              " " & .Range(EXCEL_RANGE(7, i)).Value & ",  " & _
                              " " & .Range(EXCEL_RANGE(8, i)).Value & ",  " & _
                              " " & .Range(EXCEL_RANGE(9, i)).Value & ",  " & _
                              " " & .Range(EXCEL_RANGE(10, i)).Value & ",  " & _
                              " " & .Range(EXCEL_RANGE(11, i)).Value & ",  " & _
                              " " & .Range(EXCEL_RANGE(12, i)).Value & ",  " & _
                              " " & .Range(EXCEL_RANGE(13, i)).Value & ",  " & _
                              " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
        Else
            ConnOmega.Execute "UPDATE tbl_Personnel_Payroll_Deductions_Table " & _
                              " SET Description = '" & FORMATSQL(.Range(EXCEL_RANGE(2, i)).Value) & "', " & _
                              " Abbvt = '" & FORMATSQL(.Range(EXCEL_RANGE(3, i)).Value) & "', " & _
                              " Sorting = " & .Range(EXCEL_RANGE(4, i)).Value & ", " & _
                              " ViewInDeductionModule = " & .Range(EXCEL_RANGE(5, i)).Value & ", " & _
                              " DedSched = " & .Range(EXCEL_RANGE(6, i)).Value & ", " & _
                              " EmployerShare = " & .Range(EXCEL_RANGE(7, i)).Value & ", " & _
                              " GovtDed = " & .Range(EXCEL_RANGE(8, i)).Value & ", " & _
                              " GovtDedEmpr = " & .Range(EXCEL_RANGE(9, i)).Value & ", " & _
                              " GovtDedMain = " & .Range(EXCEL_RANGE(10, i)).Value & ", " & _
                              " RefAccnt = " & .Range(EXCEL_RANGE(11, i)).Value & ", " & _
                              " WithSL = " & .Range(EXCEL_RANGE(12, i)).Value & ", " & _
                              " FixDed = " & .Range(EXCEL_RANGE(13, i)).Value & ", " & _
                              " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                              " WHERE (PK = " & .Range(EXCEL_RANGE(1, i)).Value & ")"
        End If
        rs.Close
    Next i
    
    iWorkSheet = 3  'Earning Multiplier
    .Workbooks(1).Sheets(iWorkSheet).Activate
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Earning_Multiplier"
    For i = 2 To .Workbooks(1).Sheets(iWorkSheet).UsedRange.Rows.Count
        strRange = EXCEL_RANGE(1, i)
        sValue = .Range(strRange).Value
        If CDbl(sValue) = 0 Then Exit For
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Earning_Multiplier " & _
                          " (EarningKey, CompKey, EffectDate, Multiplier) " & _
                          " VALUES (" & .Range(EXCEL_RANGE(1, i)).Value & ", " & _
                          " " & .Range(EXCEL_RANGE(2, i)).Value & ", " & _
                          " '" & FormatDateTime(.Range(EXCEL_RANGE(3, i)).Value, vbShortDate) & "', " & _
                          " " & CDbl(.Range(EXCEL_RANGE(4, i)).Value) & ")"
    Next i
    
    iWorkSheet = 4  'Overtime Multiplier
    .Workbooks(1).Sheets(iWorkSheet).Activate
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Overtime_Multiplier"
    For i = 2 To .Workbooks(1).Sheets(iWorkSheet).UsedRange.Rows.Count
        strRange = EXCEL_RANGE(1, i)
        sValue = .Range(strRange).Value
        If CDbl(sValue) = 0 Then Exit For
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Overtime_Multiplier " & _
                          " (EarningKey, CompKey, EffectDate, Multiplier) " & _
                          " VALUES (" & .Range(EXCEL_RANGE(1, i)).Value & ", " & _
                          " " & .Range(EXCEL_RANGE(2, i)).Value & ", " & _
                          " '" & FormatDateTime(.Range(EXCEL_RANGE(3, i)).Value, vbShortDate) & "', " & _
                          " " & CDbl(.Range(EXCEL_RANGE(4, i)).Value) & ")"
    Next i
    
End With
xlsApp.Application.Quit
Set xlsApp = Nothing

Screen.MousePointer = vbDefault

MsgBox "Update Successfully!                        ", vbInformation, "Success"

Exit Sub
PG:
Screen.MousePointer = vbDefault
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub

Exit Sub
Err:
Screen.MousePointer = vbDefault
Exit Sub

End Sub

Private Sub mnuUpdateGovtTablesPagIbig_Click()

If MsgBox("CONTINUE UPDATE PAGIBIG TABLE?                             ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub

MainForm.CommonDialog1.CancelError = True
On Error GoTo Err
MainForm.CommonDialog1.DialogTitle = "Open"
MainForm.CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
MainForm.CommonDialog1.ShowOpen
sPath = MainForm.CommonDialog1.Filename

Screen.MousePointer = vbHourglass

On Error GoTo PG:
Set xlsApp = CreateObject("Excel.Application")
With xlsApp
    .Workbooks.Open (sPath)
    .Visible = False
    .DisplayAlerts = False
    iWorkSheet = 1  'PagIbig Main
    .Workbooks(1).Sheets(iWorkSheet).Activate
    For i = 2 To .Workbooks(1).Sheets(iWorkSheet).UsedRange.Rows.Count
        strRange = EXCEL_RANGE(1, i)
        sValue = .Range(strRange).Value
        If CDbl(sValue) = 0 Then Exit For
        s = "SELECT tbl_Govt_PagIbigTable.* " & _
            " FROM tbl_Govt_PagIbigTable " & _
            " WHERE (PK = " & .Range(EXCEL_RANGE(1, i)).Value & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Govt_PagIbigTable " & _
                              " (PK, EffectDate,LastModified) " & _
                              " VALUES (" & .Range(EXCEL_RANGE(1, i)).Value & ", " & _
                              " '" & FormatDateTime(.Range(EXCEL_RANGE(2, i)).Value, vbShortDate) & "', " & _
                              " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
        Else
            ConnOmega.Execute "UPDATE tbl_Govt_PagIbigTable " & _
                              " SET EffectDate = '" & FormatDateTime(.Range(EXCEL_RANGE(2, i)).Value, vbShortDate) & "', " & _
                              " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                              " WHERE (PK = " & .Range(EXCEL_RANGE(1, i)).Value & ")"
        End If
        rs.Close
    Next i
    
    iWorkSheet = 2  'PagIbig Details
    .Workbooks(1).Sheets(iWorkSheet).Activate
    ConnOmega.Execute "DELETE FROM tbl_Govt_PagIbigTable_Details"
    For i = 2 To .Workbooks(1).Sheets(iWorkSheet).UsedRange.Rows.Count
        strRange = EXCEL_RANGE(1, i)
        sValue = .Range(strRange).Value
        If CDbl(sValue) = 0 Then Exit For
        ConnOmega.Execute "INSERT INTO tbl_Govt_PagIbigTable_Details " & _
                          " (MasterKey, Line, RangeFrom, RangeTo, Employee, Employer) " & _
                          " VALUES (" & .Range(EXCEL_RANGE(1, i)).Value & ", " & _
                          " " & .Range(EXCEL_RANGE(2, i)).Value & ", " & _
                          " " & CDbl(.Range(EXCEL_RANGE(3, i)).Value) & ", " & _
                          " " & CDbl(.Range(EXCEL_RANGE(4, i)).Value) & ", " & _
                          " " & CDbl(.Range(EXCEL_RANGE(5, i)).Value) & ", " & _
                          " " & CDbl(.Range(EXCEL_RANGE(6, i)).Value) & ")"
    Next i
    
    End With
xlsApp.Application.Quit
Set xlsApp = Nothing

Screen.MousePointer = vbDefault

MsgBox "Update Successfully!                        ", vbInformation, "Success"

Exit Sub
PG:
Screen.MousePointer = vbDefault
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub

Exit Sub
Err:
Screen.MousePointer = vbDefault
Exit Sub

End Sub

Private Sub mnuUpdateGovtTablesPHIC_Click()
If MsgBox("CONTINUE UPDATE PHIL HEALTH TABLE?                             ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub

MainForm.CommonDialog1.CancelError = True
On Error GoTo Err
MainForm.CommonDialog1.DialogTitle = "Open"
MainForm.CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
MainForm.CommonDialog1.ShowOpen
sPath = MainForm.CommonDialog1.Filename

Screen.MousePointer = vbHourglass

On Error GoTo PG:

Set cn = New ADODB.Connection
cn.Provider = "Microsoft.Jet.OLEDB.4.0"
cn.ConnectionString = _
    "Data Source= " & sPath & ";" & _
    "Extended Properties=Excel 8.0;"
cn.CursorLocation = adUseClient
If cn.State = adStateOpen Then cn.Close
cn.Open

Set rs = New ADODB.Recordset
If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT * FROM [PHICTable$]", cn, adOpenDynamic, adLockOptimistic
If rs.RecordCount > 0 Then
    While Not rs.EOF
        iPK = rs!PK
        a = "SELECT tbl_Govt_PhilHealthTable.* " & _
            " FROM tbl_Govt_PhilHealthTable " & _
            " WHERE (PK = " & iPK & ")"
        If ra.State = adStateOpen Then ra.Close
        ra.Open a, ConnOmega
        If ra.RecordCount = 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Govt_PhilHealthTable " & _
                              " (PK, EffectDate, MinAmt, MaxAmt) " & _
                              " VALUES (" & iPK & ", " & _
                              " '" & FormatDateTime(rs!Effectdate, vbShortDate) & "', " & _
                              " " & CDbl(rs!MinAmt) & ", " & CDbl(rs!MaxAmt) & ")"
        Else
            ConnOmega.Execute "UPDATE tbl_Govt_PhilHealthTable " & _
                              " SET EffectDate = '" & FormatDateTime(rs!Effectdate, vbShortDate) & "', " & _
                              " MinAmt = " & CDbl(rs!MinAmt) & ", " & _
                              " MaxAmt = " & CDbl(rs!MaxAmt) & " " & _
                              " WHERE (PK = " & iPK & ")"
        End If
        ra.Close
        
        ConnOmega.Execute "DELETE FROM tbl_Govt_PhilHealthTable_Details WHERE (MasterKey = " & iPK & ")"
        
        Set rt = New ADODB.Recordset
        If rt.State = adStateOpen Then rt.Close
        rt.Open "SELECT * FROM [PHICTableDetails$] WHERE ([MasterKey] = " & iPK & ")", cn, adOpenDynamic, adLockOptimistic
        If rt.RecordCount > 0 Then
            While Not rt.EOF
                ConnOmega.Execute "INSERT INTO tbl_Govt_PhilHealthTable_Details " & _
                                  " (MasterKey, Line, RangeFrom, RangeTo, Employee, Employer) " & _
                                  " VALUES (" & rt!MasterKey & ", " & rt!line & ", " & CDbl(rt!RangeFrom) & ", " & _
                                  " " & CDbl(rt!RangeTo) & ", " & CDbl(rt!Employee) & ", " & CDbl(rt!Employer) & ")"
                rt.MoveNext
            Wend
        End If
        rt.Close
        
        rs.MoveNext
    Wend
End If
rs.Close

If cn.State = adStateOpen Then cn.Close

Screen.MousePointer = vbDefault

MsgBox "Update Successfully!                        ", vbInformation, "Success"

Exit Sub
PG:
Screen.MousePointer = vbDefault
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub

Exit Sub
Err:
Exit Sub
End Sub

Private Sub mnuUpdateGovtTablesPHIC1_Click()

If MsgBox("CONTINUE UPDATE PHIC TABLE?                             ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub

MainForm.CommonDialog1.CancelError = True
On Error GoTo Err
MainForm.CommonDialog1.DialogTitle = "Open"
MainForm.CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
MainForm.CommonDialog1.ShowOpen
sPath = MainForm.CommonDialog1.Filename

Screen.MousePointer = vbHourglass

On Error GoTo PG:
Set xlsApp = CreateObject("Excel.Application")
With xlsApp
    .Workbooks.Open (sPath)
    .Visible = False
    .DisplayAlerts = False
    iWorkSheet = 1  'PHIC Main
    .Workbooks(1).Sheets(iWorkSheet).Activate
    For i = 2 To .Workbooks(1).Sheets(iWorkSheet).UsedRange.Rows.Count
        strRange = EXCEL_RANGE(1, i)
        sValue = .Range(strRange).Value
        If CDbl(sValue) = 0 Then Exit For
        s = "SELECT tbl_Govt_PhilHealthTable.* " & _
            " FROM tbl_Govt_PhilHealthTable " & _
            " WHERE (PK = " & .Range(EXCEL_RANGE(1, i)).Value & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Govt_PhilHealthTable " & _
                              " (PK, EffectDate, MinAmt, MaxAmt, LastModified) " & _
                              " VALUES (" & .Range(EXCEL_RANGE(1, i)).Value & ", " & _
                              " '" & FormatDateTime(.Range(EXCEL_RANGE(2, i)).Value, vbShortDate) & "', " & _
                              " " & CDbl(.Range(EXCEL_RANGE(3, i)).Value) & ", " & _
                              " " & CDbl(.Range(EXCEL_RANGE(4, i)).Value) & ", " & _
                              " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
        Else
            ConnOmega.Execute "UPDATE tbl_Govt_PhilHealthTable " & _
                              " SET EffectDate = '" & FormatDateTime(.Range(EXCEL_RANGE(2, i)).Value, vbShortDate) & "', " & _
                              " MinAmt = " & CDbl(.Range(EXCEL_RANGE(3, i)).Value) & ", " & _
                              " MaxAmt = " & CDbl(.Range(EXCEL_RANGE(4, i)).Value) & ", " & _
                              " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                              " WHERE (PK = " & .Range(EXCEL_RANGE(1, i)).Value & ")"
        End If
        rs.Close
    Next i
    
    iWorkSheet = 2  'PHIC Details
    .Workbooks(1).Sheets(iWorkSheet).Activate
    ConnOmega.Execute "DELETE FROM tbl_Govt_PhilHealthTable_Details"
    For i = 2 To .Workbooks(1).Sheets(iWorkSheet).UsedRange.Rows.Count
        strRange = EXCEL_RANGE(1, i)
        sValue = .Range(strRange).Value
        If CDbl(sValue) = 0 Then Exit For
        ConnOmega.Execute "INSERT INTO tbl_Govt_PhilHealthTable_Details " & _
                          " (MasterKey, Line, RangeFrom, RangeTo, wPercent, Percentage, Employee, Employer) " & _
                          " VALUES (" & .Range(EXCEL_RANGE(1, i)).Value & ", " & _
                          " " & .Range(EXCEL_RANGE(2, i)).Value & ", " & _
                          " " & CDbl(.Range(EXCEL_RANGE(3, i)).Value) & ", " & _
                          " " & CDbl(.Range(EXCEL_RANGE(4, i)).Value) & ", " & _
                          " " & CDbl(.Range(EXCEL_RANGE(5, i)).Value) & ", " & _
                          " " & CDbl(.Range(EXCEL_RANGE(6, i)).Value) & ", " & _
                          " " & CDbl(.Range(EXCEL_RANGE(7, i)).Value) & ", " & _
                          " " & CDbl(.Range(EXCEL_RANGE(8, i)).Value) & ")"
    Next i
    
    End With
xlsApp.Application.Quit
Set xlsApp = Nothing

Screen.MousePointer = vbDefault

MsgBox "Update Successfully!                        ", vbInformation, "Success"

Exit Sub
PG:
Screen.MousePointer = vbDefault
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub

Exit Sub
Err:
Screen.MousePointer = vbDefault
Exit Sub
End Sub

Private Sub mnuUpdateGovtTablesSSS_Click()

If MsgBox("CONTINUE UPDATE SSS TABLE?                             ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub

MainForm.CommonDialog1.CancelError = True
On Error GoTo Err
MainForm.CommonDialog1.DialogTitle = "Open"
MainForm.CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
MainForm.CommonDialog1.ShowOpen
sPath = MainForm.CommonDialog1.Filename

Screen.MousePointer = vbHourglass

On Error GoTo PG:

Set cn = New ADODB.Connection
cn.Provider = "Microsoft.Jet.OLEDB.4.0"
cn.ConnectionString = _
    "Data Source= " & sPath & ";" & _
    "Extended Properties=Excel 8.0;"
cn.CursorLocation = adUseClient
If cn.State = adStateOpen Then cn.Close
cn.Open

Set rs = New ADODB.Recordset
If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT * FROM [SSSTable$]", cn, adOpenDynamic, adLockOptimistic
If rs.RecordCount > 0 Then
    While Not rs.EOF
        iPK = rs!PK
        a = "SELECT tbl_Govt_SSSTable.* " & _
            " FROM tbl_Govt_SSSTable " & _
            " WHERE (PK = " & iPK & ")"
        If ra.State = adStateOpen Then ra.Close
        ra.Open a, ConnOmega
        If ra.RecordCount = 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Govt_SSSTable " & _
                              " (PK, EffectDate, MinAmt, MaxAmt) " & _
                              " VALUES (" & iPK & ", " & _
                              " '" & FormatDateTime(rs!Effectdate, vbShortDate) & "', " & _
                              " " & CDbl(rs!MinAmt) & ", " & CDbl(rs!MaxAmt) & ")"
        Else
            ConnOmega.Execute "UPDATE tbl_Govt_SSSTable " & _
                              " SET EffectDate = '" & FormatDateTime(rs!Effectdate, vbShortDate) & "', " & _
                              " MinAmt = " & CDbl(rs!MinAmt) & ", " & _
                              " MaxAmt = " & CDbl(rs!MaxAmt) & " " & _
                              " WHERE (PK = " & iPK & ")"
        End If
        ra.Close
        
        ConnOmega.Execute "DELETE FROM tbl_Govt_SSSTable_Details WHERE (MasterKey = " & iPK & ")"
        
        Set rt = New ADODB.Recordset
        If rt.State = adStateOpen Then rt.Close
        rt.Open "SELECT * FROM [SSSTableDetails$] WHERE ([MasterKey] = " & iPK & ")", cn, adOpenDynamic, adLockOptimistic
        If rt.RecordCount > 0 Then
            While Not rt.EOF
                ConnOmega.Execute "INSERT INTO tbl_Govt_SSSTable_Details " & _
                                  " (MasterKey, Line, RangeFrom, RangeTo, Employee, Employer, EC) " & _
                                  " VALUES (" & rt!MasterKey & ", " & rt!line & ", " & CDbl(rt!RangeFrom) & ", " & _
                                  " " & CDbl(rt!RangeTo) & ", " & CDbl(rt!Employee) & ", " & CDbl(rt!Employer) & ", " & _
                                  " " & CDbl(rt!EC) & ")"
                rt.MoveNext
            Wend
        End If
        rt.Close
        
        rs.MoveNext
    Wend
End If
rs.Close

If cn.State = adStateOpen Then cn.Close

Screen.MousePointer = vbDefault

MsgBox "Update Successfully!                        ", vbInformation, "Success"

Exit Sub
PG:
Screen.MousePointer = vbDefault
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub

Exit Sub
Err:
Screen.MousePointer = vbDefault
Exit Sub
End Sub

Private Sub mnuUpdateGovtTablesSSS1_Click()

If MsgBox("CONTINUE UPDATE SSS TABLE?                             ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub

MainForm.CommonDialog1.CancelError = True
On Error GoTo Err
MainForm.CommonDialog1.DialogTitle = "Open"
MainForm.CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
MainForm.CommonDialog1.ShowOpen
sPath = MainForm.CommonDialog1.Filename

Screen.MousePointer = vbHourglass

On Error GoTo PG:
Set xlsApp = CreateObject("Excel.Application")
With xlsApp
    .Workbooks.Open (sPath)
    .Visible = False
    .DisplayAlerts = False
    iWorkSheet = 1  'SSS Main
    .Workbooks(1).Sheets(iWorkSheet).Activate
    For i = 2 To .Workbooks(1).Sheets(iWorkSheet).UsedRange.Rows.Count
        strRange = EXCEL_RANGE(1, i)
        sValue = .Range(strRange).Value
        If CDbl(sValue) = 0 Then Exit For
        s = "SELECT tbl_Govt_SSSTable.* " & _
            " FROM tbl_Govt_SSSTable " & _
            " WHERE (PK = " & .Range(EXCEL_RANGE(1, i)).Value & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Govt_SSSTable " & _
                              " (PK, EffectDate, MinAmt, MaxAmt, LastModified) " & _
                              " VALUES (" & .Range(EXCEL_RANGE(1, i)).Value & ", " & _
                              " '" & FormatDateTime(.Range(EXCEL_RANGE(2, i)).Value, vbShortDate) & "', " & _
                              " " & CDbl(.Range(EXCEL_RANGE(3, i)).Value) & ", " & _
                              " " & CDbl(.Range(EXCEL_RANGE(4, i)).Value) & ", " & _
                              " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
        Else
            ConnOmega.Execute "UPDATE tbl_Govt_SSSTable " & _
                              " SET EffectDate = '" & FormatDateTime(.Range(EXCEL_RANGE(2, i)).Value, vbShortDate) & "', " & _
                              " MinAmt = " & CDbl(.Range(EXCEL_RANGE(3, i)).Value) & ", " & _
                              " MaxAmt = " & CDbl(.Range(EXCEL_RANGE(4, i)).Value) & ", " & _
                              " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                              " WHERE (PK = " & .Range(EXCEL_RANGE(1, i)).Value & ")"
        End If
        rs.Close
    Next i
    
    iWorkSheet = 2  'SSS Details
    .Workbooks(1).Sheets(iWorkSheet).Activate
    ConnOmega.Execute "DELETE FROM tbl_Govt_SSSTable_Details"
    For i = 2 To .Workbooks(1).Sheets(iWorkSheet).UsedRange.Rows.Count
        strRange = EXCEL_RANGE(1, i)
        sValue = .Range(strRange).Value
        If CDbl(sValue) = 0 Then Exit For
        ConnOmega.Execute "INSERT INTO tbl_Govt_SSSTable_Details " & _
                          " (MasterKey, Line, RangeFrom, RangeTo, Employee, Employer, EC) " & _
                          " VALUES (" & .Range(EXCEL_RANGE(1, i)).Value & ", " & _
                          " " & .Range(EXCEL_RANGE(2, i)).Value & ", " & _
                          " " & CDbl(.Range(EXCEL_RANGE(3, i)).Value) & ", " & _
                          " " & CDbl(.Range(EXCEL_RANGE(4, i)).Value) & ", " & _
                          " " & CDbl(.Range(EXCEL_RANGE(5, i)).Value) & ", " & _
                          " " & CDbl(.Range(EXCEL_RANGE(6, i)).Value) & ", " & _
                          " " & CDbl(.Range(EXCEL_RANGE(7, i)).Value) & ")"
    Next i
    
    End With
xlsApp.Application.Quit
Set xlsApp = Nothing

Screen.MousePointer = vbDefault

MsgBox "Update Successfully!                        ", vbInformation, "Success"

Exit Sub
PG:
Screen.MousePointer = vbDefault
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub

Exit Sub
Err:
Screen.MousePointer = vbDefault
Exit Sub
End Sub

Private Sub mnuUpdateGovtTablesTax_Click()

If MsgBox("CONTINUE UPDATE TAX TABLE?                             ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub

MainForm.CommonDialog1.CancelError = True
'On Error GoTo Err
MainForm.CommonDialog1.DialogTitle = "Open"
MainForm.CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
MainForm.CommonDialog1.ShowOpen
sPath = MainForm.CommonDialog1.Filename

Screen.MousePointer = vbHourglass

'On Error GoTo PG:

Set cn = New ADODB.Connection
cn.Provider = "Microsoft.Jet.OLEDB.4.0"
cn.ConnectionString = _
    "Data Source= " & sPath & ";" & _
    "Extended Properties=Excel 8.0;"
cn.CursorLocation = adUseClient
If cn.State = adStateOpen Then cn.Close
cn.Open

'Tax Category
Set rs = New ADODB.Recordset
If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT * FROM [TaxCategory$]", cn, adOpenDynamic, adLockOptimistic
If rs.RecordCount > 0 Then
    While Not rs.EOF
        If rs!PK <> Null Then
            ConnOmega.Execute "DELETE FROM tbl_Govt_TaxCategory WHERE (PK = " & rs!PK & ")"
            ConnOmega.Execute "INSERT INTO tbl_Govt_TaxCategory " & _
                              " (PK, TaxCategory, LastModified) " & _
                              " VALUES (" & rs!PK & ", '" & FORMATSQL(rs!TaxCategory) & "', " & _
                              " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
        End If
        rs.MoveNext
    Wend
End If
rs.Close

'Tax Status
Set rs = New ADODB.Recordset
If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT * FROM [TaxStatus$]", cn, adOpenDynamic, adLockOptimistic
If rs.RecordCount > 0 Then
    While Not rs.EOF
        If rs!PK <> Null Then
            ConnOmega.Execute "DELETE FROM tbl_Govt_TaxStatus WHERE (PK = " & rs!PK & ")"
            ConnOmega.Execute "INSERT INTO tbl_Govt_TaxStatus " & _
                              " (PK, TaxStatus, LastModified) " & _
                              " VALUES (" & rs!PK & ", '" & FORMATSQL(rs!TaxStatus) & "', " & _
                              " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
        End If
        rs.MoveNext
    Wend
End If
rs.Close

'Tax Status Exemption
Set rs = New ADODB.Recordset
If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT * FROM [TaxStatus_Exemption$]", cn, adOpenDynamic, adLockOptimistic
If rs.RecordCount > 0 Then
    While Not rs.EOF
        If rs!MasterKey <> Null Then
            ConnOmega.Execute "INSERT INTO tbl_Govt_TaxStatus_Exemption " & _
                              " (MasterKey, EffectDate, Exemption) " & _
                              " VALUES (" & rs!MasterKey & ", '" & FormatDateTime(rs!Effectdate, vbShortDate) & "', " & _
                              " '" & CDbl(rs!Exemption) & ")"
        End If
        rs.MoveNext
    Wend
End If
rs.Close

'Tax Table
Set rs = New ADODB.Recordset
If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT * FROM [TaxTable$]", cn, adOpenDynamic, adLockOptimistic
If rs.RecordCount > 0 Then
    While Not rs.EOF
        If rs!PK <> Null Then
            ConnOmega.Execute "DELETE FROM tbl_Govt_TaxTable WHERE (PK = " & rs!PK & ")"
            ConnOmega.Execute "INSERT INTO tbl_Govt_TaxTable " & _
                              " (PK, EffectDate, LastModified) " & _
                              " VALUES (" & rs!PK & ", '" & FormatDateTime(rs!Effectdate, vbShortDate) & "', " & _
                              " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
        End If
        rs.MoveNext
    Wend
End If
rs.Close

'Tax Table Det
Set rs = New ADODB.Recordset
If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT * FROM [TaxTable_Det$]", cn, adOpenDynamic, adLockOptimistic
If rs.RecordCount > 0 Then
    While Not rs.EOF
        If rs!MasterKey <> Null Then
            ConnOmega.Execute "INSERT INTO tbl_Govt_TaxTable_Det " & _
                              " (MasterKey, TaxCategoryKey) " & _
                              " VALUES (" & rs!MasterKey & ", " & rs!TaxCategoryKey & ")"
        End If
        rs.MoveNext
    Wend
End If
rs.Close

'Tax Table Det Det
Set rs = New ADODB.Recordset
If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT * FROM [TaxTable_Det_Det$]", cn, adOpenDynamic, adLockOptimistic
If rs.RecordCount > 0 Then
    While Not rs.EOF
        If rs!MasterKey <> Null Then
            ConnOmega.Execute "INSERT INTO tbl_Govt_TaxTable_Det_Det " & _
                              " (MasterKey, TaxCategoryKey, TaxStatusKey) " & _
                              " VALUES (" & rs!MasterKey & ", " & rs!TaxCategoryKey & ", " & _
                              " " & rs!TaxStatusKey & ")"
        End If
        rs.MoveNext
    Wend
End If
rs.Close

'Tax Table Det Det Det
Set rs = New ADODB.Recordset
If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT * FROM [TaxTable_Det_Det_Det$]", cn, adOpenDynamic, adLockOptimistic
If rs.RecordCount > 0 Then
    While Not rs.EOF
        If rs!MasterKey <> Null Then
            ConnOmega.Execute "INSERT INTO tbl_Govt_TaxTable_Det_Det_Det " & _
                              " (MasterKey, TaxCategoryKey, TaxStatusKey, " & _
                              " Line, CompLevel, Constant, Percentage) " & _
                              " VALUES (" & rs!MasterKey & ", " & rs!TaxCategoryKey & ", " & _
                              " " & rs!TaxStatusKey & ", " & rs!line & ", " & _
                              " " & CDbl(rs!CompLevel) & ", " & CDbl(rs!Constant) & ", " & _
                              " " & CDbl(rs!Percentage) & ")"
        End If
        rs.MoveNext
    Wend
End If
rs.Close
If cn.State = adStateOpen Then cn.Close

Screen.MousePointer = vbDefault

MsgBox "Update Successfully!                        ", vbInformation, "Success"

Exit Sub
PG:
Screen.MousePointer = vbDefault
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub

Exit Sub
Err:
Screen.MousePointer = vbDefault
Exit Sub
End Sub

Private Sub mnuUpdateGovtTablesTax1_Click()

If MsgBox("CONTINUE UPDATE TAX TABLE?                             ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub

MainForm.CommonDialog1.CancelError = True
On Error GoTo Err
MainForm.CommonDialog1.DialogTitle = "Open"
MainForm.CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
MainForm.CommonDialog1.ShowOpen
sPath = MainForm.CommonDialog1.Filename

Screen.MousePointer = vbHourglass

On Error GoTo PG:

Set xlsApp = CreateObject("Excel.Application")
With xlsApp
    .Workbooks.Open (sPath)
    .Visible = False
    .DisplayAlerts = False
    iWorkSheet = 1  'Tax Category
    .Workbooks(1).Sheets(iWorkSheet).Activate
    
    TotRow = .Workbooks(1).Sheets(iWorkSheet).UsedRange.Rows.Count
    For i = 2 To CDbl(TotRow)
        strRange = EXCEL_RANGE(1, i)
        sValue = .Range(strRange).Value
        If CDbl(sValue) = 0 Then Exit For
        s = "SELECT tbl_Govt_TaxCategory.* " & _
            " FROM tbl_Govt_TaxCategory " & _
            " WHERE (PK = " & .Range(EXCEL_RANGE(1, i)).Value & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Govt_TaxCategory " & _
                              " (PK, TaxCategory, LastModified) " & _
                              " VALUES (" & .Range(EXCEL_RANGE(1, i)).Value & ", " & _
                              " '" & FORMATSQL(.Range(EXCEL_RANGE(2, i)).Value) & "', " & _
                              " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
        Else
            ConnOmega.Execute "UPDATE tbl_Govt_TaxCategory " & _
                              " SET TaxCategory = '" & FORMATSQL(.Range(EXCEL_RANGE(2, i)).Value) & "', " & _
                              " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                              " WHERE (PK = " & .Range(EXCEL_RANGE(1, i)).Value & ")"
        End If
        rs.Close
    Next i
    
    iWorkSheet = 2  'Tax Status
    .Workbooks(1).Sheets(iWorkSheet).Activate
    TotRow = .Workbooks(1).Sheets(iWorkSheet).UsedRange.Rows.Count
    For i = 2 To CDbl(TotRow)
        strRange = EXCEL_RANGE(1, i)
        sValue = .Range(strRange).Value
        If CDbl(sValue) = 0 Then Exit For
        s = "SELECT tbl_Govt_TaxStatus.* " & _
            " FROM tbl_Govt_TaxStatus " & _
            " WHERE (PK = " & .Range(EXCEL_RANGE(1, i)).Value & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Govt_TaxStatus " & _
                              " (PK, TaxStatus, LastModified) " & _
                              " VALUES (" & .Range(EXCEL_RANGE(1, i)).Value & ", " & _
                              " '" & FORMATSQL(.Range(EXCEL_RANGE(2, i)).Value) & "', " & _
                              " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
        Else
            ConnOmega.Execute "UPDATE tbl_Govt_TaxStatus " & _
                              " SET TaxStatus = '" & FORMATSQL(.Range(EXCEL_RANGE(2, i)).Value) & "', " & _
                              " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                              " WHERE (PK = " & .Range(EXCEL_RANGE(1, i)).Value & ")"
        End If
        rs.Close
    Next i
    
    iWorkSheet = 3  'Tax Status Exemption
    .Workbooks(1).Sheets(iWorkSheet).Activate
    ConnOmega.Execute "DELETE FROM tbl_Govt_TaxStatus_Exemption"
    TotRow = .Workbooks(1).Sheets(iWorkSheet).UsedRange.Rows.Count
    For i = 2 To CDbl(TotRow)
        strRange = EXCEL_RANGE(1, i)
        sValue = .Range(strRange).Value
        If CDbl(sValue) = 0 Then Exit For
        ConnOmega.Execute "INSERT INTO tbl_Govt_TaxStatus_Exemption " & _
                          " (MasterKey, EffectDate, Exemption) " & _
                          " VALUES (" & .Range(EXCEL_RANGE(1, i)).Value & ", " & _
                          " '" & FormatDateTime(.Range(EXCEL_RANGE(2, i)).Value, vbShortDate) & "', " & _
                          " " & CDbl(.Range(EXCEL_RANGE(3, i)).Value) & ")"
    Next i
    
    iWorkSheet = 4  'Tax Table
    .Workbooks(1).Sheets(iWorkSheet).Activate
    TotRow = .Workbooks(1).Sheets(iWorkSheet).UsedRange.Rows.Count
    For i = 2 To CDbl(TotRow)
        strRange = EXCEL_RANGE(1, i)
        sValue = .Range(strRange).Value
        If CDbl(sValue) = 0 Then Exit For
        s = "SELECT tbl_Govt_TaxTable.* " & _
            " FROM tbl_Govt_TaxTable " & _
            " WHERE (PK = " & .Range(EXCEL_RANGE(1, i)).Value & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Govt_TaxTable " & _
                              " (PK, EffectDate, LastModified) " & _
                              " VALUES (" & .Range(EXCEL_RANGE(1, i)).Value & ", " & _
                              " '" & FormatDateTime(.Range(EXCEL_RANGE(2, i)).Value, vbShortDate) & "', " & _
                              " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
        Else
            ConnOmega.Execute "UPDATE tbl_Govt_TaxTable " & _
                              " SET EffectDate = '" & FormatDateTime(.Range(EXCEL_RANGE(2, i)).Value, vbShortDate) & "', " & _
                              " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                              " WHERE (PK = " & .Range(EXCEL_RANGE(1, i)).Value & ")"
        End If
        rs.Close
    Next i
    
    iWorkSheet = 5  'Tax Table Det
    .Workbooks(1).Sheets(iWorkSheet).Activate
    ConnOmega.Execute "DELETE FROM tbl_Govt_TaxTable_Det"
    TotRow = .Workbooks(1).Sheets(iWorkSheet).UsedRange.Rows.Count
    For i = 2 To CDbl(TotRow)
        strRange = EXCEL_RANGE(1, i)
        sValue = .Range(strRange).Value
        If CDbl(sValue) = 0 Then Exit For
        ConnOmega.Execute "INSERT INTO tbl_Govt_TaxTable_Det " & _
                          " (MasterKey, TaxCategoryKey) " & _
                          " VALUES (" & .Range(EXCEL_RANGE(1, i)).Value & ", " & _
                          " " & .Range(EXCEL_RANGE(2, i)).Value & ")"
    Next i
    
    iWorkSheet = 6  'Tax Table Det Det
    .Workbooks(1).Sheets(iWorkSheet).Activate
    ConnOmega.Execute "DELETE FROM tbl_Govt_TaxTable_Det_Det"
    TotRow = .Workbooks(1).Sheets(iWorkSheet).UsedRange.Rows.Count
    For i = 2 To CDbl(TotRow)
        strRange = EXCEL_RANGE(1, i)
        sValue = .Range(strRange).Value
        If CDbl(sValue) = 0 Then Exit For
        ConnOmega.Execute "INSERT INTO tbl_Govt_TaxTable_Det_Det " & _
                          " (MasterKey, TaxCategoryKey, TaxStatusKey) " & _
                          " VALUES (" & .Range(EXCEL_RANGE(1, i)).Value & ", " & _
                          " " & .Range(EXCEL_RANGE(2, i)).Value & ", " & _
                          " " & .Range(EXCEL_RANGE(3, i)).Value & ")"
    Next i
    
    iWorkSheet = 7  'Tax Table Det Det Det
    .Workbooks(1).Sheets(iWorkSheet).Activate
    ConnOmega.Execute "DELETE FROM tbl_Govt_TaxTable_Det_Det_Det"
    TotRow = .Workbooks(1).Sheets(iWorkSheet).UsedRange.Rows.Count
    For i = 2 To CDbl(TotRow)
        strRange = EXCEL_RANGE(1, i)
        sValue = .Range(strRange).Value
        If CDbl(sValue) = 0 Then Exit For
        ConnOmega.Execute "INSERT INTO tbl_Govt_TaxTable_Det_Det_Det " & _
                          " (MasterKey, TaxCategoryKey, TaxStatusKey, " & _
                          " Line, CompLevel, Constant, Percentage) " & _
                          " VALUES (" & .Range(EXCEL_RANGE(1, i)).Value & ", " & _
                          " " & .Range(EXCEL_RANGE(2, i)).Value & ", " & _
                          " " & .Range(EXCEL_RANGE(3, i)).Value & ", " & _
                          " " & .Range(EXCEL_RANGE(4, i)).Value & ", " & _
                          " " & CDbl(.Range(EXCEL_RANGE(5, i)).Value) & ", " & _
                          " " & CDbl(.Range(EXCEL_RANGE(6, i)).Value) & ", " & _
                          " " & CDbl(.Range(EXCEL_RANGE(7, i)).Value) & ")"
    Next i
    
End With
xlsApp.Application.Quit
Set xlsApp = Nothing

Screen.MousePointer = vbDefault

MsgBox "Update Successfully!                        ", vbInformation, "Success"

Exit Sub
PG:
Screen.MousePointer = vbDefault
xlsApp.Application.Quit
Set xlsApp = Nothing
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub

Exit Sub
Err:
Exit Sub

End Sub

Private Sub mnuUpdateMenu_Click()

If MsgBox("CONTINUE UPDATE SYSTEM MENU?                             ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub

MainForm.CommonDialog1.CancelError = True
On Error GoTo Err
MainForm.CommonDialog1.DialogTitle = "Open"
MainForm.CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
MainForm.CommonDialog1.ShowOpen
sPath = MainForm.CommonDialog1.Filename

On Error GoTo PG:

Set cn = New ADODB.Connection
cn.Provider = "Microsoft.Jet.OLEDB.4.0"
cn.ConnectionString = _
    "Data Source= " & sPath & ";" & _
    "Extended Properties=Excel 8.0;"
cn.CursorLocation = adUseClient
If cn.State = adStateOpen Then cn.Close
cn.Open

Set rs = New ADODB.Recordset
If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT * FROM [Menu$] ", cn, adOpenDynamic, adLockOptimistic
If rs.RecordCount > 0 Then
    ConnOmega.Execute "DELETE FROM tbl_System_Menu"
    While Not rs.EOF
        If IsNull(rs![Root]) = False Then
            ConnOmega.Execute "INSERT INTO tbl_System_Menu " & _
                              " (Root, MenuRelative, MenuKey, MenuName, ImageKey, Sorting, Viewed, " & _
                              " FormModule, FormAction, FormName, FormModal, DirectAccess, Non_DA_Sub, OpenImageKey) " & _
                              " VALUES (" & rs![Root] & ", '" & FORMATSQL(IIf(IsNull(rs![MenuRelative]), "", rs![MenuRelative])) & "', " & _
                              " '" & FORMATSQL(IIf(IsNull(rs![MenuKey]), "", rs![MenuKey])) & "', " & _
                              " '" & FORMATSQL(IIf(IsNull(rs![MenuName]), "", rs![MenuName])) & "', " & _
                              " " & rs![ImageKey] & ", " & rs![Sorting] & ", " & rs![Viewed] & ", " & _
                              " '" & FORMATSQL(IIf(IsNull(rs![FormModule]), "", rs![FormModule])) & "', " & _
                              " '" & FORMATSQL(IIf(IsNull(rs![FormAction]), "", rs![FormAction])) & "', " & _
                              " '" & FORMATSQL(IIf(IsNull(rs![FormName]), "", rs![FormName])) & "', " & _
                              " " & IIf(IsNull(rs![FormModal]), 0, rs![FormModal]) & ", " & _
                              " " & IIf(IsNull(rs![DirectAccess]), 0, rs![DirectAccess]) & ", " & _
                              " '" & FORMATSQL(IIf(IsNull(rs![Non_DA_Sub]), "", rs![Non_DA_Sub])) & "', " & rs!OpenImageKey & ")"
        End If
        rs.MoveNext
    Wend
End If
rs.Close

'Set rs = New ADODB.Recordset
'If rs.State = adStateOpen Then rs.Close
'rs.Open "SELECT * FROM [SC$] ", cn, adOpenDynamic, adLockOptimistic
'If rs.RecordCount > 0 Then
'    While Not rs.EOF
'        If IsNull(rs![PK]) = False Then
'            ConnOmega.Execute "INSERT INTO tbl_Service_Charge " & _
'                              " (PK, Ctrl, sMonth, sYear, MonthYear, LastModified) " & _
'                              " VALUES (" & rs![PK] & ", '" & FORMATSQL(IIf(IsNull(rs![Ctrl]), "", rs![Ctrl])) & "', " & _
'                              " " & IIf(IsNull(rs![sMonth]), "", rs![sMonth]) & ", " & _
'                              " " & IIf(IsNull(rs![sYear]), "", rs![sYear]) & ", " & _
'                              " '" & FORMATSQL(IIf(IsNull(rs![MonthYear]), "", rs![MonthYear])) & "', " & _
'                              " '" & FORMATSQL(IIf(IsNull(rs![LastModified]), "", rs![LastModified])) & "')"
'        End If
'        rs.MoveNext
'    Wend
'End If
'rs.Close
'
'Set rs = New ADODB.Recordset
'If rs.State = adStateOpen Then rs.Close
'rs.Open "SELECT * FROM [SCDetail$] ", cn, adOpenDynamic, adLockOptimistic
'If rs.RecordCount > 0 Then
'    While Not rs.EOF
'        If IsNull(rs![MasterKey]) = False Then
'            ConnOmega.Execute "INSERT INTO tbl_Service_Charge_Detail " & _
'                              " (MasterKey, sDate, ServiceCharge, Locked) " & _
'                              " VALUES (" & rs![MasterKey] & ", '" & FormatDateTime(rs![sDate], vbShortDate) & "', " & _
'                              " " & CDbl(IIf(IsNull(rs![ServiceCharge]), 0, rs![ServiceCharge])) & ", " & _
'                              " " & IIf(IsNull(rs![Locked]), "", rs![Locked]) & ")"
'        End If
'        rs.MoveNext
'    Wend
'End If
'rs.Close

MsgBox "Update Successfully!                        ", vbInformation, "Success"

LOAD_TreeView

Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub

Exit Sub
Err:
Exit Sub
End Sub

Private Sub mnuUtilityAccessRights_Click()
'If IsLoaded(aUserAccount) Then aUserAccount.ZOrder 0 Else aUserAccount.Show
LOAD_FORM "User's Account", "Open", aUserAccount, 0
End Sub

Private Sub mnuWindowsExplorer_Click()
If Statusbar1.Panels(3).Text = "" Then Exit Sub
On Error GoTo PG:
Shell ("explorer"), vbNormalFocus
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub picMain_Resize()
On Error Resume Next
imgSplitter.Height = picMain.ScaleHeight
b8LineVertical1.Height = imgSplitter.Height
imgSplitter.Left = picMain.Width - 100
trView.Top = 80
trView.Left = 80
trView.Height = picMain.ScaleHeight - picDayTime.Height - 130
trView.Width = picMain.ScaleWidth - imgSplitter.Width - 80

picDayTime.Top = picMain.ScaleHeight - picDayTime.Height
picDayTime.Width = picMain.Width - 80
picDayTimeInside.Width = picDayTime.Width - 60
End Sub

Private Sub Timer_CheckIdle_Timer()
Timer_CheckIdle.Enabled = False
If Statusbar1.Panels(3).Text <> "" Then
    If gbl_LockWhenIdle = 0 Then Timer_CheckIdle.Enabled = True: Exit Sub
    If sysidle Then
        blnIsIdle = True
    Else
        SystemIdleTime = 0
        blnIsIdle = False
    End If
End If
Timer_CheckIdle.Enabled = True
End Sub

Private Sub Timer_Text_Blink_Timer()
If Image2.Visible = True Then
    Image2.Visible = False
Else
    Image2.Visible = True
End If
End Sub

Private Sub Timer_Text_Check_Timer()
If Trim(gbl_UserName) = "" Then Timer_Text_Blink.Enabled = False: Image1.Visible = True: Image2.Visible = False: Exit Sub



ConnOmega.Execute "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED"
AA = "SELECT PK, Date_Time, Message, From_User, MsgType, " & _
    " Convert(datetime, Convert(char(6), Date_Time,12), 102) as ActDate " & _
    " From tbl_InstantMessaging " & _
    " WHERE (Opened = 0) " & _
    " AND (To_User = '" & gbl_UserName & "')"
If raa.State = adStateOpen Then raa.Close
raa.Open AA, ConnOmega
ConnOmega.Execute "SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
If raa.RecordCount > 0 Then
     If CDbl(raa!MsgType) = 0 Then
        If Trim(txtActiveForm.Text) <> "frmInstantMessagingPM" Then
            If Image2.Visible = False Then
                Image1.Visible = False
                Image2.Visible = True
                Image2.ZOrder 0
            End If
            If Timer_Text_Blink.Enabled = False Then
                Timer_Text_Blink.Enabled = True
            End If
        Else
            If frmInstantMessagingPM.Caption <> rs!From_User Then
                
                Loaded = False
                
                For Each Form In Forms
                    If Trim(raa!From_User) = Form.Caption Then
                        Loaded = True
                        Exit For
                    End If
                Next Form
                
                If Loaded = True Then
                    Form.ZOrder 0
                Else
                    If Image2.Visible = False Then
                        Image1.Visible = False
                        Image2.Visible = True
                        Image2.ZOrder 0
                    End If
                    If Timer_Text_Blink.Enabled = False Then
                        Timer_Text_Blink.Enabled = True
                    End If
                End If
                
            Else
                frmInstantMessagingPM.Timer_Msg.Enabled = True
                If Image1.Visible = False Then
                    Image1.Visible = True
                    Image2.Visible = False
                    Image1.ZOrder 0
                End If
            End If
        End If
    ElseIf CDbl(raa!MsgType) = 1 Then
        If DateValue(FormatDateTime(raa!ActDate, vbShortDate)) = DateValue(FormatDateTime(Date, vbShortDate)) Then
            If Trim(txtActiveForm.Text) <> "frmInstantMessaging" Then
                If Image2.Visible = False Then
                    Image1.Visible = False
                    Image2.Visible = True
                    Image2.ZOrder 0
                End If
            End If
            If Timer_Text_Blink.Enabled = False Then
                Timer_Text_Blink.Enabled = True
            End If
        Else
            If Image1.Visible = False Then
                Image1.Visible = True
                Image2.Visible = False
                Image1.ZOrder 0
            End If
        End If
    End If
Else
    Timer_Text_Blink.Enabled = False
    If Image1.Visible = False Then
        Image1.Visible = True
        Image2.Visible = False
        Image1.ZOrder 0
    End If
End If
raa.Close
End Sub

Private Sub Timer_when_Idle_Timer()
Timer_when_Idle.Enabled = False
If Statusbar1.Panels(3).Text = "" Then Timer_when_Idle.Enabled = True: Exit Sub
If blnIsIdle = True Then
    SystemIdleTime = SystemIdleTime + 1
    If CDbl(SystemIdleTime) = CDbl(gbl_Idle_Time) Then
        SystemIdleTime = 0
        If IsLoaded(frmSystemLocked) = False Then
            frmSystemLocked.Show 1
            Exit Sub
        End If
    End If
End If
Timer_when_Idle.Enabled = True
End Sub

Private Sub TimerDateTime_Timer()
Select Case Weekday(Now, vbMonday)
    Case 1
        lblWeekDay.ForeColor = &HFF00&
        lblWeekDay.Caption = "Mon" '"Monday"
    Case 2
        lblWeekDay.ForeColor = &HFF00&
        lblWeekDay.Caption = "Tue" '"Tuesday"
    Case 3
        lblWeekDay.ForeColor = &HFF00&
        lblWeekDay.Caption = "Wed" '"Wednesday"
    Case 4
        lblWeekDay.ForeColor = &HFF00&
        lblWeekDay.Caption = "Thu" '"Thursday"
    Case 5
        lblWeekDay.ForeColor = &HFF00&
        lblWeekDay.Caption = "Fri" '"Friday"
    Case 6
        lblWeekDay.ForeColor = &HFF00&
        lblWeekDay.Caption = "Sat" '"Saturday"
    Case 7
        lblWeekDay.ForeColor = &HFF&
        lblWeekDay.Caption = "Sun" '"Sunday"
End Select
lblDate.Caption = Format(Now, "mmm dd, yyyy") 'Format(Now, "mmmm dd, yyyy")
strTime = Format(Time, "hh:mm:ss AM/PM")
Array1 = Split(strTime, ":", -1, 1)
lblHour.Caption = Array1(0)
lblMinute.Caption = Array1(1)
Array2 = Split(Array1(2), " ", -1, 1)
lblAMPM.Caption = Array2(1)
End Sub

Private Sub TimerLogIn_Timer()
TimerLogIn.Enabled = False
LogInWithOutLoading = 1
gbl_MODULE = ""
aLogIn.Show 1
End Sub

Private Sub TimerSeparator_Timer()
If lblSeparator.Visible = True Then
    lblSeparator.Visible = False
Else
    lblSeparator.Visible = True
End If
End Sub

Private Sub TimerSplash_Timer()
TimerSplash.Enabled = False
'frmSplash.Show 1
End Sub

Private Sub LOAD_TreeView()
trView.ImageList = ImageListMother
s = "SELECT tbl_System_Menu.* " & _
    " FROM tbl_System_Menu " & _
    " WHERE (Viewed = 1) " & _
    " ORDER BY Root DESC, Sorting"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
With trView.Nodes
    .Clear
    While Not rs.EOF
        If rs!Root = 1 Then
            .Add , , CStr(rs!MenuKey), CStr(rs!MenuName), CInt(rs!ImageKey)
        Else
            .Add CStr(rs!MenuRelative), tvwChild, CStr(rs!MenuKey), CStr(rs!MenuName), CInt(rs!ImageKey)
        End If
        rs.MoveNext
    Wend
End With
rs.Close
End Sub

Private Sub trView_Collapse(ByVal Node As MSComctlLib.Node)
trView.Nodes.Item(Node.Index).Image = 1
End Sub

Private Sub trView_Expand(ByVal Node As MSComctlLib.Node)
trView.Nodes.Item(Node.Index).Image = 2
End Sub

Private Sub trView_NodeClick(ByVal Node As MSComctlLib.Node)

gbl_Form_Caption = Node.Text
iTreeViewIndex = Node.Index
'"Perfect Days (Daily)"
Select Case Node.Key
    Case "HumanResourceSetupPagIbigAdd":                    LOAD_FORM "PagIbig Additional Contribution", "Open", frmPersonnelPagIbigAddContri, 0
    Case "HumanResourceSetupDailyPerfectHours":             LOAD_FORM "Perfect Days (Daily)", "Open", frmPersonnelSetupDailyPerfectHours, 0
    Case "HumanResourceCompensationHours":                  LOAD_FORM "Personnel - Hours", "Open", frmPersonnelHours, 0
    Case "HumanResourceProfle":                             LOAD_FORM "Personnel Information", "Open", frmPersonnelInformation, 0
    Case "HumanResourceAssignID":                           LOAD_FORM "Personnel ID Number", "Open", frmPersonnelIDNumber, 0
    Case "HumanResourceActionMemo":                         LOAD_FORM "Personnel Action Memo", "Open", frmPersonnelActionV2, 0 'frmPersonnelAction, 0
    Case "HumanResourceDeactivationMemo":                   LOAD_FORM "Personnel Action Memo", "Open", frmPersonnelDeactivation, 0
    Case "HumanResourceMovementReportActive":
                                                            If AccessRights("Personnel Action Memo", "Open") = False Then
                                                                MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
                                                                       "ACCESS DENIED!                                      ", vbCritical, "Alert"
                                                                Exit Sub
                                                            End If
                                                            frmProgressBar.picAlphalist.Visible = False
                                                            frmProgressBar.TimerActive.Enabled = True
                                                            frmProgressBar.Width = 6260
                                                            frmProgressBar.Height = 970
                                                            frmProgressBar.Show 1
                                        
    Case "HumanResourceMovementReportInactive":
                                                            If AccessRights("Personnel Action Memo", "Open") = False Then
                                                                MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
                                                                       "ACCESS DENIED!                                      ", vbCritical, "Alert"
                                                                Exit Sub
                                                            End If
                                                            frmProgressBar.picAlphalist.Visible = False
                                                            frmProgressBar.TimerInactive.Enabled = True
                                                            frmProgressBar.Width = 6260
                                                            frmProgressBar.Height = 970
                                                            frmProgressBar.Show 1
    Case "HumanResourceMovementReportHeadCnt":
                                                            If AccessRights("Personnel Action Memo", "Open") = False Then
                                                                MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
                                                                       "ACCESS DENIED!                                      ", vbCritical, "Alert"
                                                                Exit Sub
                                                            End If
                                                            frmProgressBar.picAlphalist.Visible = False
                                                            frmProgressBar.TimerHeadCount.Enabled = True
                                                            frmProgressBar.Width = 6260
                                                            frmProgressBar.Height = 970
                                                            frmProgressBar.Show 1
    Case "HumanResourceMovementReportAlphalist":
                                                            If AccessRights("Personnel Action Memo", "Open") = False Then
                                                                MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
                                                                       "ACCESS DENIED!                                      ", vbCritical, "Alert"
                                                                Exit Sub
                                                            End If
                                                            frmProgressBar.picAlphalist.Visible = True
                                                            frmProgressBar.Width = 3730
                                                            frmProgressBar.Height = 1820
                                                            frmProgressBar.txtAsOf.Text = Format(Date, "mm/dd/yyyy")
                                                            frmProgressBar.txtAsOf.TabIndex = 0
                                                            frmProgressBar.Show 1
                                        
    Case "HumanResourceGovtTableSSS":                       LOAD_FORM "Personnel Gov't Table", "SSS", frmSSSTable, 0
    Case "HumanResourceGovtTablePHIC":                      LOAD_FORM "Personnel Gov't Table", "PHIC", frmPhilHealthTable, 0
    Case "HumanResourceGovtTablePagIbig":                   LOAD_FORM "Personnel Gov't Table", "PAGIBIG", frmPagIbigTable, 0
    Case "HumanResourceGovtTableTax":                       LOAD_FORM "Personnel Gov't Table", "TAX", frmTaxTable, 0
    Case "HumanResourceGovtTableExemption":                 LOAD_FORM "Personnel Gov't Table", "PERSONAL_EXEMP", frmTaxExemption, 0
    Case "HumanResourceCompensationLoans":                  LOAD_FORM "Personnel Loans", "Open", frmPersonnelLoans, 0
    Case "HumanResourceCompensationMortuary":               LOAD_FORM "Mortuary", "Open", frmPersonnelMortuaryV2, 0 'frmPersonnelCompensationMortuary, 0
    Case "HumanResourceCompensationDeductions":             LOAD_FORM "Personnel Deduction", "Open", frmPersonnelDeductions, 0
    Case "HumanResourceCompensationforDed":                 LOAD_FORM "Personnel - For Deduction", "Open", frmPersonnelDeductionsForPayroll, 0
    Case "HumanResourceCompensationCompensation":           LOAD_FORM "Personnel Compensation", "Open", frmPersonnelPayroll, 0 'frmPersonnelCompensation, 0
    Case "HumanResourceCompensationReport":                 LOAD_FORM "Personnel Compensation", "Open", frmPersonnelPayrollReport, 0 ' frmPersonnelCompensationReport, 0
    Case "HumanResourceCompensationLockPayroll":            LOAD_FORM "Personnel Compensation", "Locked Payroll", frmPersonnelPayrollLocked, 0 ' frmPersonnelCompensationLocked, 0
    Case "HumanResourceServiceChargeAbsLateUnder":          LOAD_FORM "Absent/Late/Undertime Employee", "Open", frmPersonnelAbsentLateUndertime, 0 'frmAbsentUndertimeEmployee, 0
    Case "HumanResourceServiceChargeServCharge":            LOAD_FORM "Service Charge", "Open", frmServiceCharge, 0
    Case "HumanResourceServiceChargeSummary":               LOAD_FORM "Service Charge Summary", "Open", frmServiceChargeSummary, 0
    Case "HumanResourceAllowanceEntry":                     LOAD_FORM "Allowance", "Open", frmPersonnelAllowance, 0
    Case "HumanResourceAllowanceBrowse":                    LOAD_FORM "Allowance", "Generate", frmPersonnelAllowanceBrowse, 0
    Case "HumanResourceSetupPosition":                      LOAD_FORM "Personnel Position", "Open", frmPersonnelPosition, 0
    Case "HumanResourceSetupEmploymentStatus":              LOAD_FORM "Personnel Employment Status", "Open", frmPersonnelEmploymentStatus, 0
    Case "HumanResourceSetupOTRDRate":                      LOAD_FORM "Personnel Overtime/Restday Rate", "Open", frmPersonnelOTRDRate, 0
    Case "HumanResourceSetupPayPeriod"
    Case "HumanResourceSetupServiceCharge":                 LOAD_FORM "Service Charge Setup", "Open", frmServiceChargeSetup, 0
    
    Case "FinanceAndControllerInventoryGroupingSection":    LOAD_FORM "Inventory Section", "Open", frmInvSection, 0
    Case "FinanceAndControllerInventoryGroupingClass":      LOAD_FORM "Inventory Classification", "Open", frmInvClass, 0
    Case "FinanceAndControllerSupplier":                    LOAD_FORM "Inventory Supplier", "Open", frmInvSupplier, 0
    Case "FinanceAndControllerInventoryItem":               LOAD_FORM "Inventory Items", "Open", frmInvItems, 0
    Case "FinanceAndControllerFixedAssetItems":             LOAD_FORM "Fixed Assets", "Open", frmFAItems, 0
    Case "FinanceAndControllerFixedAssetLapsing"
    Case "FinanceAndControllerPurchasesPO":                 LOAD_FORM "Purchase Order", "Open", frmInvPO, 0
    Case "FinanceAndControllerPurchasesRR":                 LOAD_FORM "Receiving Report", "Open", frmInvRR, 0 'MsgBox "Under Construction!                 ", vbInformation, "Omega" 'LOAD_FORM "Receiving Report", "Open", frmInvRR, 0
    Case "FinanceAndControllerPurchasesPI":                 LOAD_FORM "Purchase Invoice", "Open", frmInvPI, 0
    Case "FinanceAndControllerStocksMovementST":            LOAD_FORM "Stock Transfer", "Open", frmInvStockTransfer, 0
    Case "FinanceAndControllerStocksMovementSA":            LOAD_FORM "Stock Adjustment", "Open", frmInvStockAdjustment, 0
    Case "FinanceAndControllerStocksMovementSI":            LOAD_FORM "Stock Issuance", "Open", frmInvStockIssuance, 0
    
    Case "MembershipInformation":                           LOAD_FORM "Membership Information", "Open", frmMembershipInformation, 0
    Case "MembershipIDNumber":                              LOAD_FORM "Membership ID Number", "Open", frmMembershipIDNumber, 0
    Case "MembershipMovement":                              LOAD_FORM "Membership Action", "Open", frmMembershipAction, 0 '  frmMembershipShareHolder, 0 '
    Case "MembershipCompanyAccount":                        LOAD_FORM "Corporate Account", "Open", frmMembershipCompany, 0
    Case "MembershipCorporateAccount":                      LOAD_FORM "Corporate Account", "Open", frmMembershipCorporate, 0
    Case "MembershipSharesIDNumber":                        LOAD_FORM "Share ID Number", "Open", frmMembershipShareID, 0
    Case "MembershipGreenFees"
    Case "MembershipMonthlyDues"
    
    Case "FinanceAndControllerPostingRange":
    Case "FinanceAndControllerAccountingDMCMMemo":
    Case "FinanceAndControllerAccountingPettyCash":
    Case "FinanceAndControllerAccountingCV":                LOAD_FORM "Check Voucher", "Open", frmAcctgCheckVoucher, 0
    Case "FinanceAndControllerAccountingJV":                LOAD_FORM "Journal Voucher", "Open", frmAcctgGeneralJournal, 0
    Case "FinanceAndControllerAccountingChartOfAccounts":   LOAD_FORM "Chart Of Accounts", "Open", frmAcctgChartOfAccounts, 0
                                        
    
    Case "GolfOperationGolfCartInfo":                       LOAD_FORM "Golf Cart Information", "Open", frmGolfCartInformation, 0
    Case "GolfOperationCaddyInfo":                          LOAD_FORM "Caddy Information", "Open", frmCaddyInformation, 0
    Case "GolfOperationBagDrop":                            LOAD_FORM "Bag Drop", "Open", frmOperationBagDrop, 0
    Case "GolfOperationRegistration":                       LOAD_FORM "Registration", "Open", frmOperationRegistration, 0
    Case "GolfOperationProShop":                            LOAD_FORM "Pro Shop", "Open", frmOperationProShop, 0
    Case "GolfOperationProShopItem":                        LOAD_FORM "Pro Shop Items", "Open", frmOperationProShopItems, 0
    Case "GolfOperationProShopSetupBrand":                  LOAD_FORM "Pro Shop Items (Brand)", "Open", frmOperationProShopItemsBrand, 0
    Case "GolfOperationProShopSetupModel":                  LOAD_FORM "Pro Shop Items (Model)", "Open", frmOperationProShopItemsModel, 0
    Case "GolfOperationProShopSetupSizes":                  LOAD_FORM "Pro Shop Items (Sizes)", "Open", frmOperationProShopItemsSizes, 0
    Case "GolfOperationProShopSetupColor":                  LOAD_FORM "Pro Shop Items (Color)", "Open", frmOperationProShopItemsColor, 0
    Case "GolfOperationProShopSetupItemType":               LOAD_FORM "Pro Shop Items (Item Type)", "Open", frmOperationProShopItemsItemType, 0
    
    Case "GolfOperationDrivingRange":                       'LOAD_FORM "Driving Range", "Open", frmBagTag, 0 'frmBagTag
    Case "GolfOperationLocker":                             LOAD_FORM "Locker Room", "Open", frmOperationLockerRoom, 0 'frmBagTag
    Case "GolfOperationGolfCart":                           'LOAD_FORM "Golf Cart Operation", "Open", frmBagTag, 0 'frmBagTag
    Case "FoodNBeverageSetupLocation":                      LOAD_FORM "FnB Location", "Open", frmFnBLocations, 0 'frmBagTag
    
    Case "GolfOperationScoringTournamentSetup":             LOAD_FORM "Scoring Tournament Information", "Open", frmTournamentSetup, 0
    Case "GolfOperationScoringPlayerSetup":                 LOAD_FORM "Scoring Player Information", "Open", frmPlayerSetup, 0
                                                            If AccessRights("Scoring Player Information", "Open") = False Then Exit Sub
                                                            If TournamentKey = 0 Then MsgBox "Please Set One Tournament!               ", vbCritical, "Error...": Exit Sub
    Case "GolfOperationScoringTeamSetup":                   LOAD_FORM "Scoring Team Information", "Open", frmTeamSetup, 0
                                                            If AccessRights("Scoring Team Information", "Open") = False Then Exit Sub
                                                            If TournamentKey = 0 Then MsgBox "Please Set One Tournament!               ", vbCritical, "Error...": Exit Sub

    Case "GolfOperationScoringScoreCard":
                                                            If TournamentKey = 0 Then MsgBox "Please Set One Tournament!               ", vbCritical, "Error...": Exit Sub
                                                            t = "SELECT dbo.tbl_Scoring_TournamentInfo_Location.* " & _
                                                                " From dbo.tbl_Scoring_TournamentInfo_Location " & _
                                                                " WHERE (MasterKey = " & TournamentKey & ")" & _
                                                                " AND (HomeCourt = 1)"
                                                            If rt.State = adStateOpen Then rt.Close
                                                            rt.Open t, ConnOmega
                                                            If rt.RecordCount > 0 Then
                                                                LocationKey = rt!LocationKey
                                                            End If
                                                            rt.Close
                                                            Select Case ScoringType
                                                                Case 1
                                                                    If WithIndividualPlay = 0 And WithTeamPlay = 1 Then
                                                                        LOAD_FORM "Scoring Score Card", "Open", frmScoreCardTeamOnly, 0
                                                                        'If WithIndividualPlay = 0 And WithTeamPlay = 1 Then LOAD_FORM "Scoring Score Card", "Open", frmScoreCardTeamOnly, 0 Else LOAD_FORM "Scoring Score Card", "Open", frmScoreCard, 0
                                                                    Else
                                                                        gbl_Form_Caption = gbl_Form_Caption & " (Stableford)"
                                                                        LOAD_FORM "Scoring Score Card", "Open", frmScoreCardAll, 0
                                                                    End If
                                                                Case 2:  gbl_Form_Caption = gbl_Form_Caption & " (Modified Stableford)": LOAD_FORM "Scoring Score Card", "Open", frmScoreCardAll, 0 'LOAD_FORM "Scoring Score Card", "Open", frmScoreCardModifiedStableFord, 0
                                                                Case 3:  LOAD_FORM "Scoring Score Card", "Open", frmScoreCardsSystem36, 0
                                                                Case 4:  MsgBox "Not Activated!                     ", vbCritical, "Sorry"
                                                                Case 5:  gbl_Form_Caption = gbl_Form_Caption & " (Modified Molave)": LOAD_FORM "Scoring Score Card", "Open", frmScoreCardAll, 0 'LOAD_FORM "Scoring Score Card", "Open", frmScoreCardModifiedMolave, 0
                                                            End Select
    
    Case "UtilityUserRights":                               LOAD_FORM "User's Account", "Open", aUserAccount, 0
    Case "UtilityCompany":                                  LOAD_FORM "Company Information", "Open", frmCompany, 0
    Case "UtilityLocation"
    Case "UtilityDepartment":                               LOAD_FORM "Personnel Department", "Open", frmPersonnelDept, 0
    Case "UtilityBackupDatabase":                           If Statusbar1.Panels(3).Text = "" Then Exit Sub
                                                            If AccessRights("Allow Backup", "Backup") = False Then
                                                                MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
                                                                       "ACCESS DENIED!                                      ", vbCritical, "Alert"
                                                                Exit Sub
                                                            End If
                                                            
                                                            If MsgBox("Continue SQL Backup?                         ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
                                                            
                                                            Statusbar1.Panels(5).Text = "SQL Backup on progress . . . . "
                                                            DoEvents
                                                            sPath = "D:\Backup\"
                                                            sFileName = Format(Date, "mmddyy") & "F"
                                                            Create_Backup CStr(sPath), CStr(sFileName)
                                                            DoEvents
                                                            If UCase(gbl_UserName) <> "ARCHIE" Then Exit Sub

                                                            If MsgBox("COPY BACK-UP DATABASE?                           ", vbInformation + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
                                                            
                                                            On Error GoTo Err
                                                            MainForm.CommonDialog1.DialogTitle = "Save"
                                                            MainForm.CommonDialog1.InitDir = App.Path
                                                            MainForm.CommonDialog1.Filename = sFileName
                                                            MainForm.CommonDialog1.ShowSave
                                                            sPath = MainForm.CommonDialog1.Filename
                                                            
                                                            Statusbar1.Panels(5).Text = "Copying Back-up File . . . . "
                                                            DoEvents
                                                            On Error GoTo PH:
                                                            FileCopy "\\" & gbl_Server & "\backup\" & sFileName, sPath
                                                            DoEvents
                                                            Statusbar1.Panels(5).Text = ""
                                                            Exit Sub
                                                            
                                                            
COPY_LogIn:
                                                            Set oNet = CreateObject("WScript.Network")
                                                            Set oFSO = CreateObject("Scripting.FileSystemObject")
                                                            oNet.MapNetworkDrive "Z:", "\\" & gbl_Server & "\backup", False, "Administrator", "albert"
                                                            FileCopy "z:\" & sFileName, sPath
                                                            oNet.RemoveNetworkDrive "z:", True, False
                                                            DoEvents
                                                            Statusbar1.Panels(5).Text = ""
                                        
End Select

Exit Sub
PG:
Exit Sub

Exit Sub
PH:
If Err.Number = 52 Then GoTo COPY_LogIn:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub

Exit Sub
Err:
Exit Sub
End Sub
