VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{76880EFA-2CCC-4791-B35E-F6A7359CAFDD}#1.0#0"; "prjXTab.ocx"
Begin VB.Form frmMembershipInformation 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMembershipInformation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   120
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   121
         Top             =   105
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   1429
         ButtonWidth     =   1217
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   22
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Add"
               Key             =   "Add"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Edit"
               Key             =   "Edit"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Delete"
               Key             =   "Delete"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "First"
               Key             =   "First"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Back"
               Key             =   "Back"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Next"
               Key             =   "Next"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Last"
               Key             =   "Last"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Find"
               Key             =   "Find"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Print"
               Key             =   "Print"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               Key             =   "Refresh"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         MousePointer    =   99
         MouseIcon       =   "frmMembershipInformation.frx":0CCA
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   15000
         Y1              =   1005
         Y2              =   1005
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   15000
         Y1              =   90
         Y2              =   90
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   15000
         Y1              =   910
         Y2              =   910
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   33
      Top             =   6240
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
            MinWidth        =   2469
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   26458
            MinWidth        =   26458
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
   Begin RPVGCC.b8Container picSLCreditCard 
      Height          =   855
      Left            =   600
      TabIndex        =   80
      Top             =   4080
      Visible         =   0   'False
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox txtTypeCredit1 
         Height          =   315
         Left            =   8985
         TabIndex        =   84
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtCreditCard1 
         Height          =   315
         Left            =   5145
         TabIndex        =   83
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtTypeCredit 
         Height          =   315
         Left            =   5400
         TabIndex        =   82
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox txtCreditCard 
         Height          =   315
         Left            =   120
         TabIndex        =   81
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "TYPE OF CREDIT CARD"
         Height          =   255
         Left            =   5400
         TabIndex        =   86
         Top             =   120
         Width           =   3735
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "CREDIT CARD AND NOs."
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   120
         Width           =   5175
      End
   End
   Begin RPVGCC.b8Container picSLGolf 
      Height          =   855
      Left            =   600
      TabIndex        =   87
      Top             =   2760
      Visible         =   0   'False
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox txtMemberSince 
         Height          =   315
         Left            =   5400
         TabIndex        =   91
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox txtGolf 
         Height          =   315
         Left            =   120
         TabIndex        =   90
         Top             =   360
         Width           =   5175
      End
      Begin VB.TextBox txtGolf1 
         Height          =   315
         Left            =   5145
         TabIndex        =   89
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtMemberSince1 
         Height          =   315
         Left            =   8985
         TabIndex        =   88
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "MEMBER SINCE"
         Height          =   255
         Left            =   5400
         TabIndex        =   93
         Top             =   120
         Width           =   3735
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "GOLF/SPORTS CLUB"
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   120
         Width           =   5175
      End
   End
   Begin RPVGCC.b8Container picSLChild 
      Height          =   2175
      Left            =   720
      TabIndex        =   94
      Top             =   2640
      Visible         =   0   'False
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   3836
      BackColor       =   8438015
      Begin VB.TextBox txtChildStatusKey1 
         Height          =   315
         Left            =   4800
         TabIndex        =   119
         Top             =   1320
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtChildStatus1 
         Height          =   315
         Left            =   4560
         TabIndex        =   118
         Top             =   1320
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.ComboBox cmbChildStatus 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   117
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtChildPicturePath1 
         Height          =   315
         Left            =   4320
         TabIndex        =   109
         Top             =   1320
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtChildPicturePath 
         Height          =   315
         Left            =   5640
         TabIndex        =   108
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   7200
         ScaleHeight     =   1665
         ScaleWidth      =   1665
         TabIndex        =   107
         Top             =   240
         Width           =   1695
         Begin VB.Image imgImageChild 
            Height          =   1665
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1665
         End
         Begin VB.Image imgPicture2 
            Height          =   1665
            Left            =   0
            Picture         =   "frmMembershipInformation.frx":0FE4
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1665
         End
      End
      Begin VB.TextBox txtChildFName1 
         Height          =   315
         Left            =   3360
         TabIndex        =   106
         Top             =   1320
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtChildGName1 
         Height          =   315
         Left            =   3600
         TabIndex        =   105
         Top             =   1320
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtChildMName1 
         Height          =   315
         Left            =   3840
         TabIndex        =   104
         Top             =   1320
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtChildBDate1 
         Height          =   315
         Left            =   4080
         TabIndex        =   103
         Top             =   1320
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtChildBDate 
         Height          =   315
         Left            =   1560
         TabIndex        =   98
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtChildMName 
         Height          =   315
         Left            =   1560
         TabIndex        =   97
         Top             =   960
         Width           =   5535
      End
      Begin VB.TextBox txtChildGName 
         Height          =   315
         Left            =   1560
         TabIndex        =   96
         Top             =   600
         Width           =   5535
      End
      Begin VB.TextBox txtChildFName 
         Height          =   315
         Left            =   1560
         TabIndex        =   95
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "CIVIL STATUS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   116
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "BIRTH DATE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   102
         Top             =   1290
         Width           =   1215
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "FAMILY NAME"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   101
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "GIVEN NAME"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   100
         Top             =   645
         Width           =   1215
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "MIDDLE NAME"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   99
         Top             =   975
         Width           =   1215
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   0
      ScaleHeight     =   5055
      ScaleWidth      =   10455
      TabIndex        =   34
      Top             =   1080
      Width           =   10455
      Begin prjXTab.XTab XTab1 
         Height          =   5055
         Left            =   0
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   0
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   8916
         TabCaption(0)   =   "Personal Information"
         TabContCtrlCnt(0)=   1
         Tab(0)ContCtrlCap(1)=   "picPersonal"
         TabCaption(1)   =   "Family Information"
         TabContCtrlCnt(1)=   1
         Tab(1)ContCtrlCap(1)=   "picFamily"
         TabCaption(2)   =   "Employment/Business/Miscellaneous"
         TabContCtrlCnt(2)=   1
         Tab(2)ContCtrlCap(1)=   "picEmployment"
         ActiveTab       =   1
         TabTheme        =   1
         ShowFocusRect   =   0   'False
         ActiveTabBackStartColor=   16514555
         ActiveTabBackEndColor=   13023396
         InActiveTabBackStartColor=   16777215
         InActiveTabBackEndColor=   15397104
         BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OuterBorderColor=   13023396
         BottomRightInnerBorderColor=   13023396
         DisabledTabBackColor=   13023396
         DisabledTabForeColor=   13023396
         PictureMaskColor=   13023396
         Begin VB.PictureBox picEmployment 
            BackColor       =   &H00C6B8A4&
            BorderStyle     =   0  'None
            Height          =   4575
            Left            =   -74760
            ScaleHeight     =   4575
            ScaleWidth      =   9975
            TabIndex        =   69
            Top             =   480
            Width           =   9975
            Begin VB.TextBox txtBusinessFax 
               Height          =   315
               Left            =   7920
               MaxLength       =   100
               TabIndex        =   31
               Top             =   960
               Width           =   2055
            End
            Begin VB.TextBox txtBusinessTel 
               Height          =   315
               Left            =   7920
               MaxLength       =   100
               TabIndex        =   29
               Top             =   600
               Width           =   2055
            End
            Begin VB.TextBox txtBusinessNature 
               Height          =   315
               Left            =   2160
               MaxLength       =   100
               TabIndex        =   32
               Top             =   1320
               Width           =   7815
            End
            Begin VB.TextBox txtBusinessAddress 
               Height          =   315
               Left            =   2160
               MaxLength       =   100
               TabIndex        =   30
               Top             =   960
               Width           =   4935
            End
            Begin VB.TextBox txtPosition 
               Height          =   315
               Left            =   2160
               MaxLength       =   100
               TabIndex        =   28
               Top             =   600
               Width           =   4935
            End
            Begin VB.TextBox txtNameBusiness 
               Height          =   315
               Left            =   2160
               MaxLength       =   100
               TabIndex        =   27
               Top             =   240
               Width           =   7815
            End
            Begin MSComctlLib.ListView lstOtherGolf 
               Height          =   1215
               Left            =   0
               TabIndex        =   78
               Top             =   2040
               Width           =   9975
               _ExtentX        =   17595
               _ExtentY        =   2143
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   0   'False
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   4
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "#"
                  Object.Width           =   706
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Membership in Other Golf/Sports Clubs"
                  Object.Width           =   10936
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Member Since"
                  Object.Width           =   5292
               EndProperty
            End
            Begin MSComctlLib.ListView lstCards 
               Height          =   1215
               Left            =   0
               TabIndex        =   79
               Top             =   3360
               Width           =   9975
               _ExtentX        =   17595
               _ExtentY        =   2143
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   0   'False
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   4
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "#"
                  Object.Width           =   706
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Credit Card and Number"
                  Object.Width           =   9172
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Type of Credit Card"
                  Object.Width           =   7056
               EndProperty
            End
            Begin VB.Label Label36 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Miscellaneous Data"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   77
               Top             =   1800
               Width           =   9975
            End
            Begin VB.Label Label35 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Employment / Business"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   76
               Top             =   0
               Width           =   9975
            End
            Begin VB.Label Label34 
               BackStyle       =   0  'Transparent
               Caption         =   "Fax No."
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   7200
               TabIndex        =   75
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "Tel No."
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   7200
               TabIndex        =   74
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label32 
               BackStyle       =   0  'Transparent
               Caption         =   "Nature of Business"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   73
               Top             =   1320
               Width           =   2055
            End
            Begin VB.Label Label31 
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   72
               Top             =   960
               Width           =   2055
            End
            Begin VB.Label Label23 
               BackStyle       =   0  'Transparent
               Caption         =   "Position/Designation"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   71
               Top             =   600
               Width           =   2055
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "Name of Business/Employer"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   70
               Top             =   270
               Width           =   2175
            End
         End
         Begin VB.PictureBox picFamily 
            BackColor       =   &H00C6B8A4&
            BorderStyle     =   0  'None
            Height          =   4575
            Left            =   240
            ScaleHeight     =   4575
            ScaleWidth      =   9975
            TabIndex        =   57
            Top             =   480
            Width           =   9975
            Begin VB.TextBox txtSpouseMName 
               Height          =   315
               Left            =   1680
               MaxLength       =   100
               TabIndex        =   21
               Top             =   720
               Width           =   5775
            End
            Begin VB.TextBox txtSpouseGName 
               Height          =   315
               Left            =   1680
               MaxLength       =   100
               TabIndex        =   20
               Top             =   360
               Width           =   5775
            End
            Begin VB.TextBox txtSpouseLName 
               Height          =   315
               Left            =   1680
               MaxLength       =   100
               TabIndex        =   19
               Top             =   0
               Width           =   5775
            End
            Begin VB.TextBox txtSpouseContact 
               Height          =   315
               Left            =   1680
               MaxLength       =   100
               TabIndex        =   22
               Top             =   1080
               Width           =   5775
            End
            Begin VB.TextBox txtSpouseOccupation 
               Height          =   315
               Left            =   1680
               MaxLength       =   100
               TabIndex        =   23
               Top             =   1440
               Width           =   5775
            End
            Begin VB.TextBox txtSpouseCompany 
               Height          =   315
               Left            =   1680
               MaxLength       =   100
               TabIndex        =   24
               Top             =   1800
               Width           =   5775
            End
            Begin VB.TextBox txtSpouseCollege 
               Height          =   315
               Left            =   1680
               MaxLength       =   100
               TabIndex        =   25
               Top             =   2160
               Width           =   5775
            End
            Begin VB.TextBox txtSpouseDegreeObtained 
               Height          =   315
               Left            =   2160
               MaxLength       =   100
               TabIndex        =   26
               Top             =   2520
               Width           =   7815
            End
            Begin VB.PictureBox picSpouse 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   2415
               Left            =   7560
               ScaleHeight     =   2385
               ScaleWidth      =   2385
               TabIndex        =   58
               Top             =   0
               Width           =   2415
               Begin VB.Image imgSpouse 
                  Height          =   2385
                  Left            =   0
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   2385
               End
               Begin VB.Image imgSpouseLogo 
                  Height          =   2385
                  Left            =   0
                  Picture         =   "frmMembershipInformation.frx":37AE
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   2385
               End
            End
            Begin MSComctlLib.ListView lstChildren 
               Height          =   1335
               Left            =   0
               TabIndex        =   59
               Top             =   3240
               Width           =   9975
               _ExtentX        =   17595
               _ExtentY        =   2355
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   0   'False
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   11
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "#"
                  Object.Width           =   706
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Name"
                  Object.Width           =   9702
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   3
                  Text            =   "Date of Birth"
                  Object.Width           =   2646
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   4
                  Text            =   "Age"
                  Object.Width           =   1235
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "ImagePath"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "LName"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   7
                  Text            =   "FName"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   8
                  Text            =   "MName"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   9
                  Text            =   "Civil Status"
                  Object.Width           =   2646
               EndProperty
               BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   10
                  Text            =   "CivilStatusKey"
                  Object.Width           =   0
               EndProperty
            End
            Begin VB.Label Label24 
               BackStyle       =   0  'Transparent
               Caption         =   "Spouse Family Name"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   68
               Top             =   30
               Width           =   1575
            End
            Begin VB.Label Label25 
               BackStyle       =   0  'Transparent
               Caption         =   "Occupation"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   67
               Top             =   1470
               Width           =   1095
            End
            Begin VB.Label Label26 
               BackStyle       =   0  'Transparent
               Caption         =   "Contact No."
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   66
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label27 
               BackStyle       =   0  'Transparent
               Caption         =   "Company and Position"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   65
               Top             =   1830
               Width           =   1815
            End
            Begin VB.Label Label28 
               BackStyle       =   0  'Transparent
               Caption         =   "College/University"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   64
               Top             =   2190
               Width           =   1575
            End
            Begin VB.Label Label29 
               BackStyle       =   0  'Transparent
               Caption         =   "Degree Obtained and When"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   63
               Top             =   2550
               Width           =   2175
            End
            Begin VB.Label Label30 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Dependents other than spouse: Unmarried children eligible for privileges under this membership."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   62
               Top             =   3000
               Width           =   9975
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Spouse Given Name"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   61
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label Label21 
               BackStyle       =   0  'Transparent
               Caption         =   "Spouse Middle Name"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   60
               Top             =   720
               Width           =   1575
            End
         End
         Begin VB.PictureBox picPersonal 
            BackColor       =   &H00C6B8A4&
            BorderStyle     =   0  'None
            Height          =   4575
            Left            =   -74760
            ScaleHeight     =   4575
            ScaleWidth      =   9975
            TabIndex        =   36
            Top             =   480
            Width           =   9975
            Begin VB.PictureBox picMember 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   2415
               Left            =   7560
               ScaleHeight     =   2385
               ScaleWidth      =   2385
               TabIndex        =   56
               Top             =   0
               Width           =   2415
               Begin VB.Image imgMember 
                  Height          =   2385
                  Left            =   0
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   2385
               End
               Begin VB.Image imgMemberLogo 
                  Height          =   2385
                  Left            =   0
                  Picture         =   "frmMembershipInformation.frx":8164
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   2385
               End
            End
            Begin VB.TextBox txtAffiliation 
               Height          =   315
               Left            =   0
               MaxLength       =   100
               TabIndex        =   18
               Top             =   4200
               Width           =   9975
            End
            Begin VB.TextBox txtEmail 
               Height          =   315
               Left            =   7560
               MaxLength       =   100
               TabIndex        =   12
               Top             =   2520
               Width           =   2415
            End
            Begin VB.TextBox txtDegreeObtained 
               Height          =   315
               Left            =   2280
               MaxLength       =   100
               TabIndex        =   17
               Top             =   3600
               Width           =   7695
            End
            Begin VB.TextBox txtCollegeUniversity 
               Height          =   315
               Left            =   2280
               MaxLength       =   100
               TabIndex        =   16
               Top             =   3240
               Width           =   7695
            End
            Begin VB.TextBox txtTIN 
               Height          =   315
               Left            =   5520
               MaxLength       =   100
               TabIndex        =   9
               Top             =   2160
               Width           =   1935
            End
            Begin VB.TextBox txtResCertNo2 
               Height          =   315
               Left            =   5520
               MaxLength       =   100
               TabIndex        =   15
               Top             =   2880
               Width           =   4455
            End
            Begin VB.TextBox txtResCertNo1 
               Height          =   315
               Left            =   4080
               MaxLength       =   100
               TabIndex        =   14
               Top             =   2880
               Width           =   1095
            End
            Begin VB.TextBox txtResCertNo 
               Height          =   315
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   13
               Top             =   2880
               Width           =   1935
            End
            Begin VB.TextBox txtCitizen1 
               Height          =   315
               Left            =   5520
               MaxLength       =   100
               TabIndex        =   11
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtCitizen 
               Height          =   315
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   10
               Top             =   2520
               Width           =   1935
            End
            Begin VB.ComboBox cmbCivilStatus 
               Height          =   315
               Left            =   3240
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   2160
               Width           =   1815
            End
            Begin VB.ComboBox cmbSex 
               Height          =   315
               Left            =   1080
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   2160
               Width           =   1095
            End
            Begin VB.TextBox txtPlaceBirth 
               Height          =   315
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   4
               Top             =   1440
               Width           =   6375
            End
            Begin VB.TextBox txtDateBirth 
               Height          =   315
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   5
               Text            =   "11/24/1979"
               Top             =   1800
               Width           =   1095
            End
            Begin VB.TextBox txtResidence 
               Height          =   315
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   3
               Top             =   1080
               Width           =   6375
            End
            Begin VB.TextBox txtContact 
               Height          =   315
               Left            =   3240
               MaxLength       =   100
               TabIndex        =   6
               Top             =   1800
               Width           =   4215
            End
            Begin VB.TextBox txtLastName 
               Height          =   315
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   0
               Top             =   0
               Width           =   6375
            End
            Begin VB.TextBox txtFirstName 
               Height          =   315
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   1
               Top             =   360
               Width           =   6375
            End
            Begin VB.TextBox txtMiddleName 
               Height          =   315
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   2
               Top             =   720
               Width           =   6375
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Middle Name"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   55
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Given Name"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   54
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "Affiliations (Social/Fraternal/Professional/Civil Clubs)"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   53
               Top             =   3960
               Width           =   6375
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "Email Add"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   6720
               TabIndex        =   52
               Top             =   2520
               Width           =   975
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "Degree Obtained and When"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   51
               Top             =   3630
               Width           =   2175
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "College/Universities Attended"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   50
               Top             =   3270
               Width           =   2175
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "TIN"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5160
               TabIndex        =   49
               Top             =   2160
               Width           =   375
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "At"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5280
               TabIndex        =   48
               Top             =   2880
               Width           =   375
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Issued On"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   3120
               TabIndex        =   47
               Top             =   2880
               Width           =   855
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Res. Cert. No."
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   46
               Top             =   2880
               Width           =   1215
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "(if naturalized, state date when)"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   3120
               TabIndex        =   45
               Top             =   2520
               Width           =   2415
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Citizenship"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   44
               Top             =   2520
               Width           =   975
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Civil Status"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   2280
               TabIndex        =   43
               Top             =   2160
               Width           =   1095
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Sex"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   42
               Top             =   2160
               Width           =   1095
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Date of Birth"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   41
               Top             =   1800
               Width           =   975
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Place of Birth"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   40
               Top             =   1470
               Width           =   1095
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Contact No."
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   2280
               TabIndex        =   39
               Top             =   1800
               Width           =   975
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Residence"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   38
               Top             =   1110
               Width           =   1095
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Family Name"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   37
               Top             =   30
               Width           =   1095
            End
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10440
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipInformation.frx":CB1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipInformation.frx":D7F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipInformation.frx":E4CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipInformation.frx":F1A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipInformation.frx":FE82
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipInformation.frx":10B5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipInformation.frx":11836
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipInformation.frx":12510
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipInformation.frx":131EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipInformation.frx":13AC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipInformation.frx":1479E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipInformation.frx":15478
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipInformation.frx":16152
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipInformation.frx":16E2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipInformation.frx":17B06
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RPVGCC.b8Container picSearch 
      Height          =   4095
      Left            =   3120
      TabIndex        =   110
      Top             =   1080
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7223
      BackColor       =   15396057
      Begin VB.ListBox lstResult 
         Height          =   2595
         Left            =   120
         TabIndex        =   114
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   113
         Top             =   480
         Width           =   4215
      End
      Begin VB.CommandButton cmdCancel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2280
         Picture         =   "frmMembershipInformation.frx":187E0
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   3480
         Width           =   1560
      End
      Begin VB.CommandButton cmdOK 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   600
         Picture         =   "frmMembershipInformation.frx":18F3C
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   3480
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   115
         Top             =   45
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   609
         Caption         =   "Search"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   8.25
         AutoFunction    =   0   'False
         Icon            =   "frmMembershipInformation.frx":195AE
         ShadowVisible   =   0   'False
      End
   End
End
Attribute VB_Name = "frmMembershipInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Public isFind As Long

Public ChildRow As Long
Public MemberRow As Long
Public CardRow As Long

Public FocusDetail As Long

Public TRANSDetail As Long
Const isDetRefresh = 0
Const isDetAdding = 1
Const isDetEditting = 2

Dim FromAdding  As Long
Dim tmp         As Long

Dim MemberPicturePath, SpousePicturePath, ChildPicturePath, Filename, x, i, j, iPK, MemberName, SpouseName, ChildName



Private Sub BROWSER(sMemberName, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If sMemberName <> "" Then
            s = "SELECT TOP 1 tbl_Member_Information.* " & _
                " FROM tbl_Member_Information " & _
                " WHERE (LastName + ',  ' + FirstName + '  ' + MiddleName = '" & FORMATSQL(CStr(sMemberName)) & "') " & _
                " AND (ViewNot = 0) " & _
                " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName "
        Else
            s = "SELECT TOP 1 tbl_Member_Information.* " & _
                " FROM tbl_Member_Information " & _
                " WHERE (ViewNot = 0) " & _
                " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName "
                
        End If
    Case "is_HOME"
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Member_Information.* " & _
            " FROM tbl_Member_Information " & _
            " WHERE (ViewNot = 0) " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName "
    Case "is_PAGEUP"
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Member_Information.* " & _
            " FROM tbl_Member_Information " & _
            " WHERE (LastName + ',  ' + FirstName + '  ' + MiddleName < '" & FORMATSQL(CStr(sMemberName)) & "') " & _
            " AND (ViewNot = 0) " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName DESC"
    Case "is_PAGEDOWN"
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Member_Information.* " & _
            " FROM tbl_Member_Information " & _
            " WHERE (LastName + ',  ' + FirstName + '  ' + MiddleName > '" & FORMATSQL(CStr(sMemberName)) & "') " & _
            " AND (ViewNot = 0) " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName "
    Case "is_END"
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Member_Information.* " & _
            " FROM tbl_Member_Information " & _
            " WHERE (ViewNot = 0) " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName DESC"
    Case "is_FIND"
        s = "SELECT TOP 1 tbl_Member_Information.* " & _
            " FROM tbl_Member_Information " & _
            " WHERE (PK = " & sMemberName & ") " & _
            " AND (ViewNot = 0) " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName "
    Case Else: Exit Sub
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    cmbSex.ListIndex = rs!Gender
    cmbCivilStatus.ListIndex = rs!CivilStatus
    txtLastName.Text = rs!LastName
    txtFirstName.Text = rs!FirstName
    txtMiddleName.Text = rs!MiddleName
    txtResidence.Text = rs!Residence
    txtPlaceBirth.Text = rs!BirthPlace
    txtDateBirth.Text = Format(rs!BirthDate, "mm/dd/yyyy")
    txtContact.Text = rs!ContactNo
    txtEmail.Text = rs!EmailAdd
    txtTIN.Text = rs!TIN
    txtCitizen.Text = rs!Citizenship
    If IsNull(rs!Citizenship1) = False Then
        txtCitizen1.Text = Format(rs!Citizenship1, "mm/dd/yyyy")
    Else
        txtCitizen1.Text = ""
    End If
    txtResCertNo.Text = rs!ResCertNo
    If IsNull(rs!ResCertNo1) = False Then
        txtResCertNo1.Text = Format(rs!ResCertNo1, "mm/dd/yyyy")
    Else
        txtResCertNo1.Text = ""
    End If
    txtResCertNo2.Text = rs!ResCertNo2
    txtCollegeUniversity.Text = rs!CollegeUniversity
    txtDegreeObtained.Text = rs!DegreeObtained
    txtAffiliation.Text = rs!Affiliation
    txtSpouseLName.Text = rs!SpouseLName
    txtSpouseGName.Text = rs!SpouseGName
    txtSpouseMName.Text = rs!SpouseMName
    txtSpouseContact.Text = rs!SpouseContact
    txtSpouseOccupation.Text = rs!SpouseOccupation
    txtSpouseCompany.Text = rs!SpouseCompany
    txtSpouseCollege.Text = rs!SpouseCollege
    txtSpouseDegreeObtained.Text = rs!SpouseDegreeObtained
    txtNameBusiness.Text = rs!BusinessName
    txtPosition.Text = rs!BusinessPosition
    txtBusinessTel.Text = rs!BusinessTel
    txtBusinessAddress.Text = rs!BusinessAddress
    txtBusinessFax.Text = rs!BusinessFax
    txtBusinessNature.Text = rs!BusinessNature
    
    If IsNull(rs!MemberPicture) = False Then
        imgMember.Picture = LoadPicture(SHOW_IMAGES(rs!PK, 0, "Member"))
        imgMember.Visible = True
        imgMemberLogo.Visible = False
    Else
        imgMember.Visible = False
        imgMemberLogo.Visible = True
    End If
    
    If IsNull(rs!SpousePicture) = False Then
        imgSpouse.Picture = LoadPicture(SHOW_IMAGES(rs!PK, 0, "Member Spouse"))
        imgSpouse.Visible = True
        imgSpouseLogo.Visible = False
    Else
        imgSpouse.Visible = False
        imgSpouseLogo.Visible = True
    End If
    
    t = "SELECT tbl_Member_Dependent.* " & _
        " FROM tbl_Member_Dependent " & _
        " WHERE (MemberKey = " & rs!PK & ") " & _
        " ORDER BY Line"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstChildren.ListItems.Clear
        j = 0
        While Not rt.EOF
            j = j + 1
            Set x = lstChildren.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = Format(j, "0#")
            x.SubItems(2) = rt!ChildLName & ",  " & rt!ChildGName & "  " & rt!ChildMName
            x.SubItems(3) = Format(rt!ChildBirthDate, "mm/dd/yyyy")
            x.SubItems(4) = IIf(IsDate(rt!ChildBirthDate) = True, Get_Age(rt!ChildBirthDate, Date), " ")
            x.SubItems(5) = SHOW_IMAGES(rs!PK, j, "Member Child")
            x.SubItems(6) = rt!ChildLName
            x.SubItems(7) = rt!ChildGName
            x.SubItems(8) = rt!ChildMName
            x.SubItems(9) = IIf(rt!ChildStatus = 1, "SINGLE", IIf(rt!ChildStatus = 2, "MARRIED", ""))
            x.SubItems(10) = rt!ChildStatus
            rt.MoveNext
        Wend
    Else
        lstChildren.ListItems.Clear
        Set x = lstChildren.ListItems.Add()
        x.Text = ""
        x.SubItems(1) = " "
        x.SubItems(2) = " "
        x.SubItems(3) = " "
        x.SubItems(4) = " "
        x.SubItems(5) = " "
        x.SubItems(6) = " "
        x.SubItems(7) = " "
        x.SubItems(8) = " "
        x.SubItems(9) = " "
        x.SubItems(10) = "0"
    End If
    rt.Close
    
    t = "SELECT tbl_Member_OtherGolf.* " & _
        " FROM tbl_Member_OtherGolf " & _
        " WHERE (MemberKey = " & rs!PK & ") " & _
        " ORDER BY Line"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstOtherGolf.ListItems.Clear
        j = 0
        While Not rt.EOF
            j = j + 1
            Set x = lstOtherGolf.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = Format(j, "0#")
            x.SubItems(2) = rt!OtherGolfClubs
            x.SubItems(3) = rt!MemberSince
            rt.MoveNext
        Wend
    Else
        lstOtherGolf.ListItems.Clear
        Set x = lstOtherGolf.ListItems.Add()
        x.Text = ""
        x.SubItems(1) = " "
        x.SubItems(2) = " "
        x.SubItems(3) = " "
    End If
    rt.Close
    
    t = "SELECT tbl_Member_CardInfo.* " & _
        " FROM tbl_Member_CardInfo " & _
        " WHERE (MemberKey = " & rs!PK & ") " & _
        " ORDER BY Line"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstCards.ListItems.Clear
        j = 0
        While Not rt.EOF
            j = j + 1
            Set x = lstCards.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = Format(j, "0#")
            x.SubItems(2) = rt!CardAccount
            x.SubItems(3) = rt!CardType
            rt.MoveNext
        Wend
    Else
        lstCards.ListItems.Clear
        Set x = lstCards.ListItems.Add()
        x.Text = ""
        x.SubItems(1) = " "
        x.SubItems(2) = " "
        x.SubItems(3) = " "
    End If
    rt.Close
    
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", "Last Modified : " & rs!LastModified)
    
    SaveSetting App.EXEName, "MemberName", "NameMember", rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If AccessRights("Membership Information", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
XTab1.ActiveTab = 0
DoEvents
txtLastName.SetFocus
End Sub

Private Sub PRESS_F2()
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("Membership Information", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_EDITTING

End Sub

Private Sub PRESS_DELETE()
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("Membership Information", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If MsgBox("ARE YOU SURE IN DELETING THIS RECORD?                    ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_Member_Information WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "MemberName", "NameMember", ""), "is_PAGEDOWN"
If Trim(txtLastName.Text) = "" Then BROWSER GetSetting(App.EXEName, "MemberName", "NameMember", ""), "is_END"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()


If Trim(txtLastName.Text) = "" Then XTab1.ActiveTab = 0: DoEvents: MsgBox "Please Supply Family Name!                    ", vbCritical, "Error...": txtLastName.SetFocus: Exit Sub
If Trim(txtFirstName.Text) = "" Then XTab1.ActiveTab = 0: DoEvents: MsgBox "Please Supply Given Name!                    ", vbCritical, "Error...":  txtFirstName.SetFocus: Exit Sub
If Trim(txtMiddleName.Text) = "" Then XTab1.ActiveTab = 0: DoEvents: MsgBox "Please Supply Middle Name!                    ", vbCritical, "Error...":  txtMiddleName.SetFocus: Exit Sub
If IsDate(txtDateBirth.Text) = False Then XTab1.ActiveTab = 0: DoEvents: MsgBox "Please Supply a Valid Date of Birth!                    ", vbCritical, "Error...": txtDateBirth.SetFocus: Exit Sub
If cmbSex.ListIndex <= 0 Then XTab1.ActiveTab = 0: DoEvents: MsgBox "Please Select Gender!                  ", vbCritical, "Error...":  cmbSex.SetFocus: Exit Sub
If cmbCivilStatus.ListIndex <= 0 Then XTab1.ActiveTab = 0: DoEvents: MsgBox "Please Select Civil Status!                  ", vbCritical, "Error...":  cmbCivilStatus.SetFocus: Exit Sub
If Trim(txtCitizen1.Text) <> "" Then If IsDate(txtCitizen1.Text) = False Then XTab1.ActiveTab = 0: DoEvents: MsgBox "Please Supply a Valid Date!                 ", vbCritical, "Error...": txtCitizen1.SetFocus: Exit Sub
If Trim(txtResCertNo1.Text) <> "" Then If IsDate(txtResCertNo1.Text) = False Then XTab1.ActiveTab = 0: DoEvents: MsgBox "Please Supply a Valid Date!                    ", vbCritical, "Error...":  txtResCertNo1.SetFocus: Exit Sub

With lstChildren.ListItems
    For i = 1 To .Count
        If Trim(.Item(i).SubItems(6)) <> "" And _
        Trim(.Item(i).SubItems(7)) <> "" And _
        Trim(.Item(i).SubItems(8)) <> "" Then
             If IsDate(.Item(i).SubItems(3)) = False Then XTab1.ActiveTab = 1: DoEvents: MsgBox "Please Supply a valid date for Dependent Birthday!                    ", vbCritical, "Error...": lstChildren.SetFocus: Exit Sub
             If CDbl(.Item(i).SubItems(10)) = 0 Then XTab1.ActiveTab = 1: DoEvents: MsgBox "Please Select Dependent Civil Status!                        ": lstChildren.SetFocus: Exit Sub
        End If
    Next i
End With

On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    
    FromAdding = 1
    
    MemberName = Trim(txtLastName.Text) & ",  " & Trim(txtFirstName.Text) & "  " & Trim(txtMiddleName.Text)
    SpouseName = Trim(txtSpouseLName.Text) & ",  " & Trim(txtSpouseGName.Text) & "  " & Trim(txtSpouseMName.Text)
    
    ConnOmega.Execute "INSERT INTO tbl_Member_Information " & _
                      " (LastName, FirstName, MiddleName, Residence, BirthPlace, BirthDate, ContactNo, EmailAdd, Gender, CivilStatus, TIN, Citizenship, " & _
                      " ResCertNo, ResCertNo2, CollegeUniversity, DegreeObtained, Affiliation, SpouseLName, SpouseGName, SpouseMName, " & _
                      " SpouseContact, SpouseOccupation, SpouseCompany, SpouseCollege, SpouseDegreeObtained, BusinessName, BusinessPosition, BusinessTel, " & _
                      " BusinessAddress, BusinessFax, BusinessNature) " & _
                      " VALUES ('" & FORMATSQL(Trim(txtLastName.Text)) & "', '" & FORMATSQL(Trim(txtFirstName.Text)) & "', '" & FORMATSQL(Trim(txtMiddleName.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtResidence.Text)) & "', '" & FORMATSQL(Trim(txtPlaceBirth.Text)) & "', '" & FormatDateTime(txtDateBirth.Text, vbShortDate) & "', " & _
                      " '" & FORMATSQL(Trim(txtContact.Text)) & "', '" & FORMATSQL(Trim(txtEmail.Text)) & "', " & cmbSex.ListIndex & ", " & cmbCivilStatus.ListIndex & ", " & _
                      " '" & FORMATSQL(Trim(txtTIN.Text)) & "', '" & FORMATSQL(Trim(txtCitizen.Text)) & "', '" & FORMATSQL(Trim(txtResCertNo.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtResCertNo2.Text)) & "', '" & FORMATSQL(Trim(txtCollegeUniversity.Text)) & "', '" & FORMATSQL(Trim(txtDegreeObtained.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtAffiliation.Text)) & "', '" & FORMATSQL(Trim(txtSpouseLName.Text)) & "', '" & FORMATSQL(Trim(txtSpouseGName.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtSpouseMName.Text)) & "', '" & FORMATSQL(Trim(txtSpouseContact.Text)) & "', '" & FORMATSQL(Trim(txtSpouseOccupation.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtSpouseCompany.Text)) & "', '" & FORMATSQL(Trim(txtSpouseCollege.Text)) & "', '" & FORMATSQL(Trim(txtSpouseDegreeObtained.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtNameBusiness.Text)) & "', '" & FORMATSQL(Trim(txtPosition.Text)) & "', '" & FORMATSQL(Trim(txtBusinessTel.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtBusinessAddress.Text)) & "', '" & FORMATSQL(Trim(txtBusinessFax.Text)) & "', '" & FORMATSQL(Trim(txtBusinessNature.Text)) & "')"
    
    iPK = 0
    s = "SELECT PK " & _
        " FROM tbl_Member_Information " & _
        " WHERE (LastName + ',  '+ FirstName + '  ' + MiddleName = '" & FORMATSQL(CStr(MemberName)) & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        iPK = rs!PK
    End If
    rs.Close
    
End If

If TRANSACTIONTYPE = is_EDITTING Then
    
    FromAdding = 0
    
    iPK = Statusbar1.Panels(1).Text
    MemberName = Trim(txtLastName.Text) & ",  " & Trim(txtFirstName.Text) & "  " & Trim(txtMiddleName.Text)
    SpouseName = Trim(txtSpouseLName.Text) & ",  " & Trim(txtSpouseGName.Text) & "  " & Trim(txtSpouseMName.Text)
    
    ConnOmega.Execute "UPDATE tbl_Member_Information " & _
                      " SET LastName = '" & FORMATSQL(Trim(txtLastName.Text)) & "', FirstName = '" & FORMATSQL(Trim(txtFirstName.Text)) & "', " & _
                      " MiddleName = '" & FORMATSQL(Trim(txtMiddleName.Text)) & "', Residence = '" & FORMATSQL(Trim(txtResidence.Text)) & "', " & _
                      " BirthPlace = '" & FORMATSQL(Trim(txtPlaceBirth.Text)) & "', BirthDate = '" & FormatDateTime(txtDateBirth.Text, vbShortDate) & "', " & _
                      " ContactNo = '" & FORMATSQL(Trim(txtContact.Text)) & "', EmailAdd = '" & FORMATSQL(Trim(txtEmail.Text)) & "', " & _
                      " Gender = " & cmbSex.ListIndex & ", CivilStatus = " & cmbCivilStatus.ListIndex & ", TIN = '" & FORMATSQL(Trim(txtTIN.Text)) & "', " & _
                      " Citizenship = '" & FORMATSQL(Trim(txtCitizen.Text)) & "',  " & _
                      " ResCertNo = '" & FORMATSQL(Trim(txtResCertNo.Text)) & "', ResCertNo2 = '" & FORMATSQL(Trim(txtResCertNo2.Text)) & "', " & _
                      " CollegeUniversity = '" & FORMATSQL(Trim(txtCollegeUniversity.Text)) & "', DegreeObtained = '" & FORMATSQL(Trim(txtDegreeObtained.Text)) & "', " & _
                      " Affiliation = '" & FORMATSQL(Trim(txtAffiliation.Text)) & "', SpouseLName = '" & FORMATSQL(Trim(txtSpouseLName.Text)) & "', " & _
                      " SpouseGName = '" & FORMATSQL(Trim(txtSpouseGName.Text)) & "', SpouseMName = '" & FORMATSQL(Trim(txtSpouseMName.Text)) & "', " & _
                      " SpouseContact = '" & FORMATSQL(Trim(txtSpouseContact.Text)) & "', SpouseOccupation = '" & FORMATSQL(Trim(txtSpouseOccupation.Text)) & "', " & _
                      " SpouseCompany = '" & FORMATSQL(Trim(txtSpouseCompany.Text)) & "', SpouseCollege = '" & FORMATSQL(Trim(txtSpouseCollege.Text)) & "', " & _
                      " SpouseDegreeObtained = '" & FORMATSQL(Trim(txtSpouseDegreeObtained.Text)) & "', BusinessName = '" & FORMATSQL(Trim(txtNameBusiness.Text)) & "', " & _
                      " BusinessPosition = '" & FORMATSQL(Trim(txtPosition.Text)) & "', BusinessTel = '" & FORMATSQL(Trim(txtBusinessTel.Text)) & "', " & _
                      " BusinessAddress = '" & FORMATSQL(Trim(txtBusinessAddress.Text)) & "', BusinessFax = '" & FORMATSQL(Trim(txtBusinessFax.Text)) & "', " & _
                      " BusinessNature = '" & FORMATSQL(Trim(txtBusinessNature.Text)) & "' " & _
                      " WHERE (PK = " & iPK & ")"
    
End If

If CDbl(iPK) <> 0 Then
    If IsDate(txtCitizen1.Text) = True Then ConnOmega.Execute "UPDATE tbl_Member_Information SET Citizenship1 = '" & FormatDateTime(txtCitizen1.Text, vbShortDate) & "' WHERE (PK = " & iPK & ")"
    If IsDate(txtResCertNo1.Text) = True Then ConnOmega.Execute "UPDATE tbl_Member_Information SET ResCertNo1 = '" & FormatDateTime(txtResCertNo1.Text, vbShortDate) & "' WHERE (PK = " & iPK & ")"
    
    t = "SELECT PK " & _
        " FROM tbl_Member_IDNumber " & _
        " WHERE (MemberKey = " & iPK & ") " & _
        " AND (MemberType = 1)"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        ConnOmega.Execute "UPDATE tbl_Member_IDNumber " & _
                          " SET MemberName = '" & FORMATSQL(CStr(MemberName)) & "' " & _
                          " WHERE (PK = " & rt!PK & ") " & _
                          " AND (MemberType = 1)"
        rt.MoveNext
    Wend
    rt.Close
    
    t = "SELECT PK " & _
        " FROM tbl_Member_IDNumber " & _
        " WHERE (MemberKey = " & iPK & ") " & _
        " AND (MemberType = 2)"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        ConnOmega.Execute "UPDATE tbl_Member_IDNumber " & _
                          " SET MemberName = '" & FORMATSQL(CStr(SpouseName)) & "' " & _
                          " WHERE (PK = " & rt!PK & ") " & _
                          " AND (MemberType = 2)"
        rt.MoveNext
    Wend
    rt.Close
    
    
    If Trim(CStr(MemberPicturePath)) <> "" Then
        SAVE_IMAGES iPK, 0, CStr(MemberPicturePath), "Member"
        t = "SELECT PK " & _
            " FROM tbl_Member_IDNumber " & _
            " WHERE (MemberKey = " & iPK & ") " & _
            " AND (MemberType = 1)"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        While Not rt.EOF
            SAVE_IMAGES rt!PK, 0, SHOW_IMAGES(iPK, 0, "Member"), "Member ID Number"
            rt.MoveNext
        Wend
        rt.Close
    End If
    
    If Trim(CStr(SpousePicturePath)) <> "" Then
        SAVE_IMAGES iPK, 0, CStr(SpousePicturePath), "Member Spouse"
        t = "SELECT PK " & _
            " FROM tbl_Member_IDNumber " & _
            " WHERE (MemberKey = " & iPK & ") " & _
            " AND (MemberType = 2)"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        While Not rt.EOF
            SAVE_IMAGES rt!PK, 0, SHOW_IMAGES(iPK, 0, "Member Spouse"), "Member ID Number"
            rt.MoveNext
        Wend
        rt.Close
    End If
    
    ConnOmega.Execute "DELETE FROM tbl_Member_Dependent WHERE (MemberKey = " & iPK & ")"
    With lstChildren.ListItems
        j = 0
        For i = 1 To .Count
            If Trim(.Item(i).SubItems(6)) <> "" And _
            Trim(.Item(i).SubItems(7)) <> "" And _
            Trim(.Item(i).SubItems(8)) <> "" And _
            IsDate(.Item(i).SubItems(3)) = True Then
                j = j + 1
                
                ChildName = Trim(.Item(i).SubItems(6)) & ",  " & Trim(.Item(i).SubItems(7)) & "  " & Trim(.Item(i).SubItems(8))
                
                ConnOmega.Execute "INSERT INTO tbl_Member_Dependent " & _
                                  " (MemberKey, Line, ChildLName, ChildGName, ChildMName, ChildBirthDate, ChildStatus) " & _
                                  " VALUES (" & iPK & ", " & j & ", '" & FORMATSQL(Trim(.Item(i).SubItems(6))) & "', " & _
                                  " '" & FORMATSQL(Trim(.Item(i).SubItems(7))) & "', '" & FORMATSQL(Trim(.Item(i).SubItems(8))) & "', " & _
                                  " '" & FormatDateTime(.Item(i).SubItems(3), vbShortDate) & "', " & _
                                  " " & .Item(i).SubItems(10) & ")"
                
                t = "SELECT PK " & _
                    " FROM tbl_Member_IDNumber " & _
                    " WHERE (MemberKey = " & iPK & ") " & _
                    " AND (MemberType = 3) " & _
                    " AND (MemberChildLine = " & j & ")"
                If rt.State = adStateOpen Then rt.Close
                rt.Open t, ConnOmega
                While Not rt.EOF
                    ConnOmega.Execute "UPDATE tbl_Member_IDNumber " & _
                                      " SET MemberName = '" & FORMATSQL(CStr(ChildName)) & "' " & _
                                      " WHERE (PK = " & rt!PK & ")"
                    rt.MoveNext
                Wend
                rt.Close
                
                If Trim(.Item(i).SubItems(5)) <> "" Then
                    SAVE_IMAGES iPK, j, Trim(.Item(i).SubItems(5)), "Member Child"
                    t = "SELECT PK " & _
                        " FROM tbl_Member_IDNumber " & _
                        " WHERE (MemberKey = " & iPK & ") " & _
                        " AND (MemberType = 3) " & _
                        " AND (MemberChildLine = " & j & ")"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        SAVE_IMAGES rt!PK, j, SHOW_IMAGES(iPK, j, "Member Child"), "Member ID Number (Child)"
                        rt.MoveNext
                    Wend
                    rt.Close
                End If
            End If
        Next i
    End With
    
    ConnOmega.Execute "DELETE FROM tbl_Member_OtherGolf WHERE (MemberKey = " & iPK & ")"
    With lstOtherGolf.ListItems
        j = 0
        For i = 1 To .Count
            If Trim(.Item(i).SubItems(2)) <> "" Then
                j = j + 1
                ConnOmega.Execute "INSERT INTO tbl_Member_OtherGolf " & _
                                  " (MemberKey, Line, OtherGolfClubs, MemberSince) " & _
                                  " VALUES (" & iPK & ", " & j & ", " & _
                                  " '" & FORMATSQL(Trim(.Item(i).SubItems(2))) & "', " & _
                                  " '" & FORMATSQL(Trim(.Item(i).SubItems(3))) & "')"
            End If
        Next i
    End With
    
    ConnOmega.Execute "DELETE FROM tbl_Member_CardInfo WHERE (MemberKey = " & iPK & ")"
    With lstCards.ListItems
        j = 0
        For i = 1 To .Count
            If Trim(.Item(i).SubItems(2)) <> "" Then
                j = j + 1
                ConnOmega.Execute "INSERT INTO tbl_Member_CardInfo " & _
                                  " (MemberKey, Line, CardAccount, CardType) " & _
                                  " VALUES (" & iPK & ", " & j & ", " & _
                                  " '" & FORMATSQL(Trim(.Item(i).SubItems(2))) & "', " & _
                                  " '" & FORMATSQL(Trim(.Item(i).SubItems(3))) & "')"
            End If
        Next i
    End With
    
    ConnOmega.Execute "UPDATE tbl_Member_Information SET LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' WHERE (PK = " & iPK & ")"
    
End If

CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER MemberName, "is_LOAD"

Select Case FocusDetail
    Case 1: txtSpouseDegreeObtained.SetFocus
    Case 2: txtBusinessNature.SetFocus
    Case 3: txtBusinessNature.SetFocus
End Select

If FromAdding = 1 Then
    If MsgBox("Add Another Information?                         ", vbInformation + vbYesNo + vbDefaultButton1, "Add") = vbYes Then PRESS_INSERT Else Exit Sub
End If

Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
PopupMenu MainFormPopupF.mnuMemberFind, , Toolbar1.Buttons(15).Left, Toolbar1.Buttons(15).Top + Toolbar1.Buttons(15).Height

'picToolbar.Enabled = False
'picMain.Enabled = False
'picSearch.ZOrder 0
'txtSearch.Text = ""
'picSearch.Visible = True
'txtSearch.SetFocus
End Sub

Private Sub PRESS_F9()
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub

End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSearch.Visible = True Then cmdCancel_Click: Exit Sub
    Unload Me
Else
    If picSLChild.Visible = True Then
        With lstChildren.ListItems
            Select Case TRANSDetail
                Case isDetAdding
                    If .Count > 1 Then
                        .Remove ChildRow
                    Else
                        .Item(1).SubItems(1) = " "
                        .Item(1).SubItems(2) = " "
                        .Item(1).SubItems(3) = " "
                        .Item(1).SubItems(4) = " "
                        .Item(1).SubItems(5) = " "
                        .Item(1).SubItems(6) = " "
                        .Item(1).SubItems(7) = " "
                        .Item(1).SubItems(8) = " "
                        .Item(1).SubItems(9) = " "
                        .Item(1).SubItems(10) = "0"
                    End If
                    ChildRow = .Count
                Case isDetEditting
                    .Item(ChildRow).SubItems(2) = Trim(txtChildFName1.Text) & ",  " & Trim(txtChildGName1.Text) & "  " & Trim(txtChildMName1.Text)
                    .Item(ChildRow).SubItems(3) = txtChildBDate1.Text
                    .Item(ChildRow).SubItems(4) = IIf(IsDate(txtChildBDate1.Text), Get_Age(txtChildBDate1.Text, Date), " ")
                    .Item(ChildRow).SubItems(5) = Trim(txtChildPicturePath1.Text)
                    .Item(ChildRow).SubItems(6) = Trim(txtChildFName1.Text)
                    .Item(ChildRow).SubItems(7) = Trim(txtChildGName1.Text)
                    .Item(ChildRow).SubItems(8) = Trim(txtChildMName1.Text)
                    .Item(ChildRow).SubItems(9) = Trim(txtChildStatus1.Text)
                    .Item(ChildRow).SubItems(10) = Trim(txtChildStatusKey1.Text)
            End Select
        End With
        picToolbar.Enabled = True
        picMain.Enabled = True
        picSLChild.Visible = False
        lstChildren.ListItems(ChildRow).EnsureVisible
        lstChildren.ListItems(ChildRow).Selected = True
        lstChildren.SetFocus
        Exit Sub
    End If
    If picSLGolf.Visible = True Then
        With lstOtherGolf.ListItems
            Select Case TRANSDetail
                Case isDetAdding
                    If .Count > 1 Then
                        .Remove MemberRow
                    Else
                        .Item(1).SubItems(1) = " "
                        .Item(1).SubItems(2) = " "
                        .Item(1).SubItems(3) = " "
                    End If
                    MemberRow = .Count
                Case isDetEditting
                    .Item(MemberRow).SubItems(2) = Trim(txtGolf1.Text)
                    .Item(MemberRow).SubItems(3) = Trim(txtMemberSince1.Text)
            End Select
        End With
        picToolbar.Enabled = True
        picMain.Enabled = True
        picSLGolf.Visible = False
        lstOtherGolf.ListItems(MemberRow).EnsureVisible
        lstOtherGolf.ListItems(MemberRow).Selected = True
        lstOtherGolf.SetFocus
        Exit Sub
    End If
    If picSLCreditCard.Visible = True Then
        With lstCards.ListItems
            Select Case TRANSDetail
                Case isDetAdding
                    If .Count > 1 Then
                        .Remove CardRow
                    Else
                        .Item(1).SubItems(1) = " "
                        .Item(1).SubItems(2) = " "
                        .Item(1).SubItems(3) = " "
                    End If
                    CardRow = .Count
                Case isDetEditting
                    .Item(CardRow).SubItems(2) = Trim(txtCreditCard1.Text)
                    .Item(CardRow).SubItems(3) = Trim(txtTypeCredit1.Text)
            End Select
        End With
        picToolbar.Enabled = True
        picMain.Enabled = True
        picSLCreditCard.Visible = False
        lstCards.ListItems(CardRow).EnsureVisible
        lstCards.ListItems(CardRow).Selected = True
        lstCards.SetFocus
        Exit Sub
    End If
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "MemberName", "NameMember", ""), "is_LOAD"
    Select Case FocusDetail
        Case 1: txtSpouseDegreeObtained.SetFocus
        Case 2: txtBusinessNature.SetFocus
        Case 3: txtBusinessNature.SetFocus
    End Select
End If
End Sub

Private Sub CLEARTEXT()

ChildRow = 0
MemberRow = 0
CardRow = 0

cmbSex.ListIndex = 0
cmbCivilStatus.ListIndex = 0
txtLastName.Text = ""
txtFirstName.Text = ""
txtMiddleName.Text = ""
txtResidence.Text = ""
txtPlaceBirth.Text = ""
txtDateBirth.Text = ""
txtContact.Text = ""
txtEmail.Text = ""
txtTIN.Text = ""
txtCitizen.Text = ""
txtCitizen1.Text = ""
txtResCertNo.Text = ""
txtResCertNo1.Text = ""
txtResCertNo2.Text = ""
txtCollegeUniversity.Text = ""
txtDegreeObtained.Text = ""
txtAffiliation.Text = ""
txtSpouseLName.Text = ""
txtSpouseGName.Text = ""
txtSpouseMName.Text = ""
txtSpouseContact.Text = ""
txtSpouseOccupation.Text = ""
txtSpouseCompany.Text = ""
txtSpouseCollege.Text = ""
txtSpouseDegreeObtained.Text = ""
txtNameBusiness.Text = ""
txtPosition.Text = ""
txtBusinessTel.Text = ""
txtBusinessAddress.Text = ""
txtBusinessFax.Text = ""
txtBusinessNature.Text = ""

lstChildren.ListItems.Clear
Set x = lstChildren.ListItems.Add()
x.Text = ""
x.SubItems(1) = " "
x.SubItems(2) = " "
x.SubItems(3) = " "
x.SubItems(4) = " "
x.SubItems(5) = " "
x.SubItems(6) = " "
x.SubItems(7) = " "
x.SubItems(8) = " "
x.SubItems(9) = " "
x.SubItems(10) = "0"

lstOtherGolf.ListItems.Clear
Set x = lstOtherGolf.ListItems.Add()
x.Text = ""
x.SubItems(1) = " "
x.SubItems(2) = " "
x.SubItems(3) = " "

lstCards.ListItems.Clear
Set x = lstCards.ListItems.Add()
x.Text = ""
x.SubItems(1) = " "
x.SubItems(2) = " "
x.SubItems(3) = " "

imgMemberLogo.Visible = True
imgMember.Picture = LoadPicture("")
imgMember.Visible = False

imgSpouseLogo.Visible = True
imgSpouse.Picture = LoadPicture("")
imgSpouse.Visible = False

Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""

MemberPicturePath = ""
SpousePicturePath = ""
ChildPicturePath = ""

End Sub

Private Sub LOCKTEXT(bln As Boolean)
cmbSex.Locked = bln
cmbCivilStatus.Locked = bln
txtLastName.Locked = bln
txtFirstName.Locked = bln
txtMiddleName.Locked = bln
txtResidence.Locked = bln
txtPlaceBirth.Locked = bln
txtDateBirth.Locked = bln
txtContact.Locked = bln
txtEmail.Locked = bln
txtTIN.Locked = bln
txtCitizen.Locked = bln
txtCitizen1.Locked = bln
txtResCertNo.Locked = bln
txtResCertNo1.Locked = bln
txtResCertNo2.Locked = bln
txtCollegeUniversity.Locked = bln
txtDegreeObtained.Locked = bln
txtAffiliation.Locked = bln
txtSpouseLName.Locked = bln
txtSpouseGName.Locked = bln
txtSpouseMName.Locked = bln
txtSpouseContact.Locked = bln
txtSpouseOccupation.Locked = bln
txtSpouseCompany.Locked = bln
txtSpouseCollege.Locked = bln
txtSpouseDegreeObtained.Locked = bln
txtNameBusiness.Locked = bln
txtPosition.Locked = bln
txtBusinessTel.Locked = bln
txtBusinessAddress.Locked = bln
txtBusinessFax.Locked = bln
txtBusinessNature.Locked = bln
If bln Then
    imgMemberLogo.ToolTipText = ""
    imgMember.ToolTipText = ""
    imgSpouseLogo.ToolTipText = ""
    imgSpouse.ToolTipText = ""
Else
    imgMemberLogo.ToolTipText = "Double Click to Insert Picture"
    imgMember.ToolTipText = "Double Click to Insert Picture"
    imgSpouseLogo.ToolTipText = "Double Click to Insert Picture"
    imgSpouse.ToolTipText = "Double Click to Insert Picture"
End If
End Sub

Private Sub TOOLBARFUNC(intSelect As Integer)
With Toolbar1
    Select Case intSelect
        Case 1      '=== REFRESH ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 4
            .Buttons(9).Image = 5
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 12
            .Buttons(21).Image = 13
            '.Buttons(23).Image = 13
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "First"
            .Buttons(9).Caption = "Back"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            '.Buttons(19).Caption = "Post"
            .Buttons(19).Caption = "Refresh"
            .Buttons(21).Caption = "Close"
            .Buttons(1).Enabled = True
            .Buttons(3).Enabled = True
            .Buttons(5).Enabled = True
            .Buttons(7).Enabled = True
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = True
            .Buttons(13).Enabled = True
            .Buttons(15).Enabled = True
            .Buttons(17).Enabled = True
            .Buttons(19).Enabled = True
            .Buttons(21).Enabled = True
            '.Buttons(23).Enabled = True
            .Buttons(1).ToolTipText = "NEW (Ins)"
            .Buttons(3).ToolTipText = "EDIT (F2)"
            .Buttons(5).ToolTipText = "DELETE (Del)"
            .Buttons(7).ToolTipText = "FIRST (Home)"
            .Buttons(9).ToolTipText = "BACK (PgUp)"
            .Buttons(11).ToolTipText = "NEXT (PgDown)"
            .Buttons(13).ToolTipText = "LAST (End)"
            .Buttons(15).ToolTipText = "FIND (F6)"
            .Buttons(17).ToolTipText = "PRINT (F9)"
            '.Buttons(19).ToolTipText = "POST (F8)"
            .Buttons(19).ToolTipText = "REFRESH (F11)"
            .Buttons(21).ToolTipText = "CLOSE (Esc)"
        Case 2      '=== ADD/EDIT ====
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 14
            .Buttons(9).Image = 15
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 12
            .Buttons(21).Image = 13
            '.Buttons(23).Image = 13
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "Save"
            .Buttons(9).Caption = "Undo"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            '.Buttons(19).Caption = "Post"
            .Buttons(19).Caption = "Refresh"
            .Buttons(21).Caption = "Close"
            .Buttons(1).Enabled = False
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = False
            .Buttons(7).Enabled = True
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
            .Buttons(19).Enabled = False
            .Buttons(21).Enabled = False
            '.Buttons(23).Enabled = False
            .Buttons(1).ToolTipText = ""
            .Buttons(3).ToolTipText = ""
            .Buttons(5).ToolTipText = ""
            .Buttons(7).ToolTipText = "SAVE (F5)"
            .Buttons(9).ToolTipText = "UNDO (Esc)"
            .Buttons(11).ToolTipText = ""
            .Buttons(13).ToolTipText = ""
            .Buttons(15).ToolTipText = ""
            .Buttons(17).ToolTipText = ""
            .Buttons(19).ToolTipText = ""
            .Buttons(21).ToolTipText = ""
            '.Buttons(23).ToolTipText = ""
        Case 3      '=== FIND ===
           .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 4
            .Buttons(9).Image = 15
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 12
            .Buttons(21).Image = 13
            '.Buttons(23).Image = 13
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "First"
            .Buttons(9).Caption = "Undo"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            '.Buttons(19).Caption = "Post"
            .Buttons(19).Caption = "Refresh"
            .Buttons(21).Caption = "Close"
            .Buttons(1).Enabled = False
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = False
            .Buttons(7).Enabled = False
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
            .Buttons(19).Enabled = False
            .Buttons(21).Enabled = False
            '.Buttons(23).Enabled = False
            .Buttons(1).ToolTipText = ""
            .Buttons(3).ToolTipText = ""
            .Buttons(5).ToolTipText = ""
            .Buttons(7).ToolTipText = ""
            .Buttons(9).ToolTipText = "UNDO (Esc)"
            .Buttons(11).ToolTipText = ""
            .Buttons(13).ToolTipText = ""
            .Buttons(15).ToolTipText = ""
            .Buttons(17).ToolTipText = ""
            .Buttons(19).ToolTipText = ""
            .Buttons(21).ToolTipText = ""
            '.Buttons(23).ToolTipText = ""
        Case 4      '=== EMPTY DETAIL ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 14
            .Buttons(9).Image = 15
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 12
            .Buttons(21).Image = 13
            '.Buttons(23).Image = 13
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "Save"
            .Buttons(9).Caption = "Undo"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            '.Buttons(19).Caption = "Post"
            .Buttons(19).Caption = "Refresh"
            .Buttons(21).Caption = "Close"
            .Buttons(1).Enabled = True
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = False
            .Buttons(7).Enabled = True
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
            .Buttons(19).Enabled = False
            .Buttons(21).Enabled = False
            '.Buttons(23).Enabled = False
            .Buttons(1).ToolTipText = "NEW (Ins)"
            .Buttons(3).ToolTipText = ""
            .Buttons(5).ToolTipText = ""
            .Buttons(7).ToolTipText = "SAVE (F5)"
            .Buttons(9).ToolTipText = "UNDO (Esc)"
            .Buttons(11).ToolTipText = ""
            .Buttons(13).ToolTipText = ""
            .Buttons(15).ToolTipText = ""
            .Buttons(17).ToolTipText = ""
            .Buttons(19).ToolTipText = ""
            .Buttons(21).ToolTipText = ""
            '.Buttons(23).ToolTipText = ""
        Case 5      '=== NOT EMPTY DETAIL ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 14
            .Buttons(9).Image = 15
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 12
            .Buttons(21).Image = 13
            '.Buttons(23).Image = 13
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "Save"
            .Buttons(9).Caption = "Undo"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            '.Buttons(19).Caption = "Post"
            .Buttons(19).Caption = "Refresh"
            .Buttons(21).Caption = "Close"
            .Buttons(1).Enabled = True
            .Buttons(3).Enabled = True
            .Buttons(5).Enabled = True
            .Buttons(7).Enabled = True
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
            .Buttons(19).Enabled = False
            .Buttons(21).Enabled = False
            '.Buttons(23).Enabled = False
            .Buttons(1).ToolTipText = "NEW (Ins)"
            .Buttons(3).ToolTipText = "EDIT (F2)"
            .Buttons(5).ToolTipText = "DELET (Del)"
            .Buttons(7).ToolTipText = "SAVE (F5)"
            .Buttons(9).ToolTipText = "UNDO (Esc)"
            .Buttons(11).ToolTipText = ""
            .Buttons(13).ToolTipText = ""
            .Buttons(15).ToolTipText = ""
            .Buttons(17).ToolTipText = ""
            .Buttons(19).ToolTipText = ""
            .Buttons(21).ToolTipText = ""
            '.Buttons(23).ToolTipText = ""
    End Select
End With
End Sub

Private Sub b8TitleBar2_CLoseClick()
cmdCancel_Click
End Sub


Private Sub cmbChildStatus_Click()
If TRANSDetail = isDetAdding Or _
TRANSDetail = isDetEditting Then
    If cmbChildStatus.ListIndex <= 0 Then Exit Sub
    With lstChildren.ListItems
        .Item(ChildRow).SubItems(9) = Trim(cmbChildStatus.Text)
        .Item(ChildRow).SubItems(10) = cmbChildStatus.ListIndex
    End With
End If
End Sub

Private Sub cmbChildStatus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If cmbChildStatus.ListIndex = 0 Then MsgBox "Please Select Civil Status!                 ", vbCritical, "Error...": cmbChildStatus.SetFocus: Exit Sub
    picToolbar.Enabled = True
    picMain.Enabled = True
    picSLChild.Visible = False
    lstChildren.SetFocus
End If
End Sub

Private Sub cmdCancel_Click()
picToolbar.Enabled = True
picMain.Enabled = True
picSearch.Visible = False
End Sub

Private Sub cmdOK_Click()
If lstResult.ListIndex = -1 Then Exit Sub
BROWSER lstResult.ItemData(lstResult.ListIndex), "is_FIND"
cmdCancel_Click
End Sub

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert:   PRESS_INSERT
    Case vbKeyF2:       PRESS_F2
    Case vbKeyDelete:   PRESS_DELETE
    Case vbKeyF5:       PRESS_F5
    Case vbKeyF6:       PRESS_F6
    Case vbKeyF9:       PRESS_F9
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "MemberName", "NameMember", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "MemberName", "NameMember", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "MemberName", "NameMember", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "MemberName", "NameMember", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.Height - Me.Height) / 20
Me.Left = (MainForm.Width - Me.Width) / 5
XTab1.ActiveTab = 0
cmbSex.Clear
cmbSex.AddItem "--Select--"
cmbSex.AddItem "MALE"
cmbSex.AddItem "FEMALE"
cmbCivilStatus.Clear
cmbCivilStatus.AddItem "--Select--"
cmbCivilStatus.AddItem "SINGLE"
cmbCivilStatus.AddItem "MARRIED"
cmbChildStatus.Clear
cmbChildStatus.AddItem "--Select--"
cmbChildStatus.AddItem "SINGLE"
cmbChildStatus.AddItem "MARRIED"

ChildRow = 0
MemberRow = 0
CardRow = 0

CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "MemberName", "NameMember", ""), "is_LOAD"
If Trim(txtLastName.Text) = "" Then BROWSER GetSetting(App.EXEName, "MemberName", "NameMember", ""), "is_HOME"


tmp = SetWindowLong(txtLastName.hwnd, GWL_STYLE, GetWindowLong(txtLastName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtFirstName.hwnd, GWL_STYLE, GetWindowLong(txtFirstName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtMiddleName.hwnd, GWL_STYLE, GetWindowLong(txtMiddleName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtResidence.hwnd, GWL_STYLE, GetWindowLong(txtResidence.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPlaceBirth.hwnd, GWL_STYLE, GetWindowLong(txtPlaceBirth.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtContact.hwnd, GWL_STYLE, GetWindowLong(txtContact.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtTIN.hwnd, GWL_STYLE, GetWindowLong(txtTIN.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtCitizen.hwnd, GWL_STYLE, GetWindowLong(txtCitizen.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtCitizen1.hwnd, GWL_STYLE, GetWindowLong(txtCitizen1.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtResCertNo.hwnd, GWL_STYLE, GetWindowLong(txtResCertNo.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtResCertNo1.hwnd, GWL_STYLE, GetWindowLong(txtResCertNo1.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtResCertNo2.hwnd, GWL_STYLE, GetWindowLong(txtResCertNo2.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtCollegeUniversity.hwnd, GWL_STYLE, GetWindowLong(txtCollegeUniversity.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtDegreeObtained.hwnd, GWL_STYLE, GetWindowLong(txtDegreeObtained.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtAffiliation.hwnd, GWL_STYLE, GetWindowLong(txtAffiliation.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSpouseLName.hwnd, GWL_STYLE, GetWindowLong(txtSpouseLName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSpouseGName.hwnd, GWL_STYLE, GetWindowLong(txtSpouseGName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSpouseMName.hwnd, GWL_STYLE, GetWindowLong(txtSpouseMName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSpouseContact.hwnd, GWL_STYLE, GetWindowLong(txtSpouseContact.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSpouseOccupation.hwnd, GWL_STYLE, GetWindowLong(txtSpouseOccupation.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSpouseCompany.hwnd, GWL_STYLE, GetWindowLong(txtSpouseCompany.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSpouseCollege.hwnd, GWL_STYLE, GetWindowLong(txtSpouseCollege.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSpouseDegreeObtained.hwnd, GWL_STYLE, GetWindowLong(txtSpouseDegreeObtained.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtNameBusiness.hwnd, GWL_STYLE, GetWindowLong(txtNameBusiness.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPosition.hwnd, GWL_STYLE, GetWindowLong(txtPosition.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtBusinessTel.hwnd, GWL_STYLE, GetWindowLong(txtBusinessTel.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtBusinessAddress.hwnd, GWL_STYLE, GetWindowLong(txtBusinessAddress.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtBusinessFax.hwnd, GWL_STYLE, GetWindowLong(txtBusinessFax.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtBusinessNature.hwnd, GWL_STYLE, GetWindowLong(txtBusinessNature.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtChildFName.hwnd, GWL_STYLE, GetWindowLong(txtChildFName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtChildGName.hwnd, GWL_STYLE, GetWindowLong(txtChildGName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtChildMName.hwnd, GWL_STYLE, GetWindowLong(txtChildMName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtGolf.hwnd, GWL_STYLE, GetWindowLong(txtGolf.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtMemberSince.hwnd, GWL_STYLE, GetWindowLong(txtMemberSince.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtCreditCard.hwnd, GWL_STYLE, GetWindowLong(txtCreditCard.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtTypeCredit.hwnd, GWL_STYLE, GetWindowLong(txtTypeCredit.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)

End Sub


Private Sub Form_Unload(Cancel As Integer)
If picSLChild.Visible = True Then Cancel = -1
If picSLGolf.Visible = True Then Cancel = -1
If picSLCreditCard.Visible = True Then Cancel = -1
If picSearch.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub imgImageChild_DblClick()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    MainForm.CommonDialog1.CancelError = True
    On Error GoTo ErrorHandler
    MainForm.CommonDialog1.Filter = "Image Files|*.JPG;*.JPEG;*.JPE;*.BMP;*.RLE;*.DIB;*.GIF;*.PNG;*.TIF;*.TIFF"
    MainForm.CommonDialog1.ShowOpen
    Filename = Trim(MainForm.CommonDialog1.Filename)
    If ((FileLen(Filename) \ 1024) + 1) > CDbl(IMAGEFILESIZE(Date)) Then
        MsgBox "Image is too large please reduce the size to " & IMAGEFILESIZE(Date) & "kb or below!          ", vbCritical, "Error..."
        Exit Sub
    End If
    txtChildPicturePath.Text = Filename
    imgImageChild.Picture = LoadPicture(Filename)
    imgPicture2.Visible = False
    imgImageChild.Visible = True
End If
Exit Sub
ErrorHandler:
Exit Sub
End Sub

Private Sub imgMember_DblClick()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    MainForm.CommonDialog1.CancelError = True
    On Error GoTo ErrorHandler
    MainForm.CommonDialog1.Filter = "Image Files|*.JPG;*.JPEG;*.JPE;*.BMP;*.RLE;*.DIB;*.GIF;*.PNG;*.TIF;*.TIFF"
    MainForm.CommonDialog1.ShowOpen
    Filename = Trim(MainForm.CommonDialog1.Filename)
    If ((FileLen(Filename) \ 1024) + 1) > CDbl(IMAGEFILESIZE(Date)) Then
        MsgBox "Image is too large please reduce the size to " & IMAGEFILESIZE(Date) & "kb or below!          ", vbCritical, "Error..."
        Exit Sub
    End If
    MemberPicturePath = Filename
    imgMember.Picture = LoadPicture(Filename)
    imgMemberLogo.Visible = False
    imgMember.Visible = True
End If
Exit Sub
ErrorHandler:
Exit Sub
End Sub

Private Sub imgMemberLogo_DblClick()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    MainForm.CommonDialog1.CancelError = True
    On Error GoTo ErrorHandler
    MainForm.CommonDialog1.Filter = "Image Files|*.JPG;*.JPEG;*.JPE;*.BMP;*.RLE;*.DIB;*.GIF;*.PNG;*.TIF;*.TIFF"
    MainForm.CommonDialog1.ShowOpen
    Filename = Trim(MainForm.CommonDialog1.Filename)
    If ((FileLen(Filename) \ 1024) + 1) > CDbl(IMAGEFILESIZE(Date)) Then
        MsgBox "Image is too large please reduce the size to " & IMAGEFILESIZE(Date) & "kb or below!          ", vbCritical, "Error..."
        Exit Sub
    End If
    MemberPicturePath = Filename
    imgMember.Picture = LoadPicture(Filename)
    imgMemberLogo.Visible = False
    imgMember.Visible = True
End If
Exit Sub
ErrorHandler:
Exit Sub
End Sub

Private Sub imgPicture2_DblClick()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    MainForm.CommonDialog1.CancelError = True
    On Error GoTo ErrorHandler
    MainForm.CommonDialog1.Filter = "Image Files|*.JPG;*.JPEG;*.JPE;*.BMP;*.RLE;*.DIB;*.GIF;*.PNG;*.TIF;*.TIFF"
    MainForm.CommonDialog1.ShowOpen
    Filename = Trim(MainForm.CommonDialog1.Filename)
    If ((FileLen(Filename) \ 1024) + 1) > CDbl(IMAGEFILESIZE(Date)) Then
        MsgBox "Image is too large please reduce the size to " & IMAGEFILESIZE(Date) & "kb or below!          ", vbCritical, "Error..."
        Exit Sub
    End If
    txtChildPicturePath.Text = Filename
    imgImageChild.Picture = LoadPicture(Filename)
    imgPicture2.Visible = False
    imgImageChild.Visible = True
End If
Exit Sub
ErrorHandler:
Exit Sub
End Sub

Private Sub imgSpouse_DblClick()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    MainForm.CommonDialog1.CancelError = True
    On Error GoTo ErrorHandler
    MainForm.CommonDialog1.Filter = "Image Files|*.JPG;*.JPEG;*.JPE;*.BMP;*.RLE;*.DIB;*.GIF;*.PNG;*.TIF;*.TIFF"
    MainForm.CommonDialog1.ShowOpen
    Filename = Trim(MainForm.CommonDialog1.Filename)
    If ((FileLen(Filename) \ 1024) + 1) > CDbl(IMAGEFILESIZE(Date)) Then
        MsgBox "Image is too large please reduce the size to " & IMAGEFILESIZE(Date) & "kb or below!          ", vbCritical, "Error..."
        Exit Sub
    End If
    SpousePicturePath = Filename
    imgSpouse.Picture = LoadPicture(Filename)
    imgSpouseLogo.Visible = False
    imgSpouse.Visible = True
End If
Exit Sub
ErrorHandler:
Exit Sub
End Sub

Private Sub imgSpouseLogo_DblClick()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    MainForm.CommonDialog1.CancelError = True
    On Error GoTo ErrorHandler
    MainForm.CommonDialog1.Filter = "Image Files|*.JPG;*.JPEG;*.JPE;*.BMP;*.RLE;*.DIB;*.GIF;*.PNG;*.TIF;*.TIFF"
    MainForm.CommonDialog1.ShowOpen
    Filename = Trim(MainForm.CommonDialog1.Filename)
    If ((FileLen(Filename) \ 1024) + 1) > CDbl(IMAGEFILESIZE(Date)) Then
        MsgBox "Image is too large please reduce the size to " & IMAGEFILESIZE(Date) & "kb or below!          ", vbCritical, "Error..."
        Exit Sub
    End If
    SpousePicturePath = Filename
    imgSpouse.Picture = LoadPicture(Filename)
    imgSpouseLogo.Visible = False
    imgSpouse.Visible = True
End If
Exit Sub
ErrorHandler:
Exit Sub
End Sub

Private Sub lstCards_GotFocus()
FocusDetail = 3
TRANSDetail = isDetRefresh
On Error GoTo PG:
CardRow = lstCards.SelectedItem.Index
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub lstCards_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo PG:
CardRow = lstCards.SelectedItem.Index
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub lstCards_LostFocus()
FocusDetail = 0
End Sub

Private Sub lstCards_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
If CardRow = 0 Then Exit Sub
If Trim(lstCards.ListItems.Item(CardRow).SubItems(2)) = "" Then MainFormPopupF.mnuMemberDetailsEdit.Enabled = False: MainFormPopupF.mnuMemberDetailsDelete.Enabled = False Else MainFormPopupF.mnuMemberDetailsEdit.Enabled = True: MainFormPopupF.mnuMemberDetailsDelete.Enabled = True
If Button = 2 Then PopupMenu MainFormPopupF.mnuMemberDetails
End Sub

Private Sub lstChildren_GotFocus()
FocusDetail = 1
TRANSDetail = isDetRefresh
On Error GoTo PG:
ChildRow = lstChildren.SelectedItem.Index
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub lstChildren_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo PG:
ChildRow = lstChildren.SelectedItem.Index
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub lstChildren_LostFocus()
FocusDetail = 0
End Sub

Private Sub lstChildren_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
If ChildRow = 0 Then Exit Sub
If Trim(lstChildren.ListItems.Item(ChildRow).SubItems(2)) = "" Then MainFormPopupF.mnuMemberDetailsEdit.Enabled = False: MainFormPopupF.mnuMemberDetailsDelete.Enabled = False Else MainFormPopupF.mnuMemberDetailsEdit.Enabled = True: MainFormPopupF.mnuMemberDetailsDelete.Enabled = True
If Button = 2 Then PopupMenu MainFormPopupF.mnuMemberDetails
End Sub

Private Sub lstOtherGolf_GotFocus()
FocusDetail = 2
TRANSDetail = isDetRefresh
On Error GoTo PG:
MemberRow = lstOtherGolf.SelectedItem.Index
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub lstOtherGolf_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo PG:
MemberRow = lstOtherGolf.SelectedItem.Index
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub lstOtherGolf_LostFocus()
FocusDetail = 0
End Sub

Private Sub lstOtherGolf_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
If MemberRow = 0 Then Exit Sub
If Trim(lstOtherGolf.ListItems.Item(MemberRow).SubItems(2)) = "" Then MainFormPopupF.mnuMemberDetailsEdit.Enabled = False: MainFormPopupF.mnuMemberDetailsDelete.Enabled = False Else MainFormPopupF.mnuMemberDetailsEdit.Enabled = True: MainFormPopupF.mnuMemberDetailsDelete.Enabled = True
If Button = 2 Then PopupMenu MainFormPopupF.mnuMemberDetails
End Sub

Private Sub lstResult_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "MemberName", "NameMember", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "MemberName", "NameMember", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "MemberName", "NameMember", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "MemberName", "NameMember", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":   PRESS_F9
    Case "Refresh": BROWSER Statusbar1.Panels(1).Text, "is_FIND"
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtChildBDate_Change()
If TRANSDetail = isDetAdding Or _
TRANSDetail = isDetEditting Then
    With lstChildren.ListItems
        .Item(ChildRow).SubItems(3) = Trim(txtChildBDate.Text)
        If IsDate(txtChildBDate.Text) = True Then
            .Item(ChildRow).SubItems(4) = Get_Age(txtChildBDate.Text, Date)
        Else
            .Item(ChildRow).SubItems(4) = " "
        End If
    End With
End If
End Sub

Private Sub txtChildBDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    cmbChildStatus.SetFocus
End If
If KeyCode = vbKeyUp Then txtChildMName.SetFocus
End Sub

Private Sub txtChildBDate_LostFocus()
If IsDate(txtChildBDate.Text) = True Then
    txtChildBDate.Text = Format(FormatDateTime(txtChildBDate.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub

Private Sub txtChildFName_Change()
If TRANSDetail = isDetAdding Or _
TRANSDetail = isDetEditting Then
    With lstChildren.ListItems
        .Item(ChildRow).SubItems(2) = Trim(txtChildFName.Text) & IIf(Trim(txtChildGName.Text) = "", "", ",  " & Trim(txtChildGName.Text)) & IIf(Trim(txtChildMName.Text) = "", "", "  " & Trim(txtChildMName.Text))
        .Item(ChildRow).SubItems(6) = Trim(txtChildFName.Text)
    End With
End If
End Sub

Private Sub txtChildFName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then txtChildGName.SetFocus
End Sub

Private Sub txtChildGName_Change()
If TRANSDetail = isDetAdding Or _
TRANSDetail = isDetEditting Then
    With lstChildren.ListItems
        .Item(ChildRow).SubItems(2) = IIf(Trim(txtChildFName.Text) = "", "", Trim(txtChildFName.Text) & ",  ") & IIf(Trim(txtChildGName.Text) = "", "", Trim(txtChildGName.Text)) & IIf(Trim(txtChildMName.Text) = "", "", "  " & Trim(txtChildMName.Text))
        .Item(ChildRow).SubItems(7) = Trim(txtChildGName.Text)
    End With
End If
End Sub

Private Sub txtChildGName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then txtChildMName.SetFocus
If KeyCode = vbKeyUp Then txtChildFName.SetFocus
End Sub

Private Sub txtChildMName_Change()
If TRANSDetail = isDetAdding Or _
TRANSDetail = isDetEditting Then
    With lstChildren.ListItems
        .Item(ChildRow).SubItems(2) = IIf(Trim(txtChildFName.Text) = "", "", Trim(txtChildFName.Text) & ",  ") & IIf(Trim(txtChildGName.Text) = "", "", Trim(txtChildGName.Text)) & IIf(Trim(txtChildMName.Text) = "", "", "  " & Trim(txtChildMName.Text))
        .Item(ChildRow).SubItems(8) = Trim(txtChildMName.Text)
    End With
End If
End Sub

Private Sub txtChildMName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then txtChildBDate.SetFocus
If KeyCode = vbKeyUp Then txtChildGName.SetFocus
End Sub

Private Sub txtChildPicturePath_Change()
If TRANSDetail = isDetAdding Or _
TRANSDetail = isDetEditting Then
    With lstChildren.ListItems
        .Item(ChildRow).SubItems(5) = Trim(txtChildPicturePath.Text)
    End With
End If
End Sub

Private Sub txtCreditCard_Change()
If TRANSDetail = isDetAdding Or _
TRANSDetail = isDetEditting Then
    With lstCards.ListItems
        .Item(CardRow).SubItems(2) = Trim(txtCreditCard.Text)
    End With
End If
End Sub

Private Sub txtCreditCard_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtTypeCredit.SetFocus
End Sub

Private Sub txtGolf_Change()
If TRANSDetail = isDetAdding Or _
TRANSDetail = isDetEditting Then
    With lstOtherGolf.ListItems
        .Item(MemberRow).SubItems(2) = Trim(txtGolf.Text)
    End With
End If
End Sub

Private Sub txtGolf_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtMemberSince.SetFocus
End Sub

Private Sub txtMemberSince_Change()
If TRANSDetail = isDetAdding Or _
TRANSDetail = isDetEditting Then
    With lstOtherGolf.ListItems
        .Item(MemberRow).SubItems(3) = Trim(txtMemberSince.Text)
    End With
End If
End Sub

Private Sub txtMemberSince_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    picToolbar.Enabled = True
    picMain.Enabled = True
    picSLGolf.Visible = False
    lstOtherGolf.SetFocus
End If
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstResult.Clear: Exit Sub
lstResult.Clear
Select Case isFind
    Case 1
        s = "SELECT PK, LastName + ',  ' + FirstName + '  ' + MiddleName AS MemberName " & _
            " From tbl_Member_Information " & _
            " WHERE (LastName LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
    Case 2
        s = "SELECT PK, LastName + ',  ' + FirstName + '  ' + MiddleName AS MemberName " & _
            " From tbl_Member_Information " & _
            " WHERE (FirstName LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
    Case 3
        s = "SELECT PK, LastName + ',  ' + FirstName + '  ' + MiddleName AS MemberName " & _
            " From tbl_Member_Information " & _
            " WHERE (MiddleName LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResult.AddItem rs!MemberName
    lstResult.ItemData(lstResult.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResult.ListCount Then lstResult.ListIndex = 0
End Sub

Private Sub txtSearch_GotFocus()
HTEXT txtSearch
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResult.SetFocus
If KeyCode = vbKeyDown Then lstResult.SetFocus
End Sub

Private Sub txtTypeCredit_Change()
If TRANSDetail = isDetAdding Or _
TRANSDetail = isDetEditting Then
    With lstCards.ListItems
        .Item(CardRow).SubItems(3) = Trim(txtTypeCredit.Text)
    End With
End If
End Sub

Private Sub txtTypeCredit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    picToolbar.Enabled = True
    picMain.Enabled = True
    picSLCreditCard.Visible = False
    lstCards.SetFocus
End If
End Sub

Private Sub XTab1_Click()
Select Case XTab1.ActiveTab
    Case 0: txtLastName.SetFocus
    Case 1: txtSpouseLName.SetFocus
    Case 2: txtNameBusiness.SetFocus
End Select
End Sub

