VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelCompensation 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPersonnelCompensation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   184
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   185
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
         MouseIcon       =   "frmPersonnelCompensation.frx":08CA
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
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   7140
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11994
            MinWidth        =   11994
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin RPVGCC.b8Container picProgress 
      Height          =   975
      Left            =   2760
      TabIndex        =   137
      Top             =   3360
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1720
      BackColor       =   13023396
      Begin VB.PictureBox picProgressBar 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5235
         TabIndex        =   138
         Top             =   120
         Width           =   5295
      End
   End
   Begin RPVGCC.b8Container picTaxWithHeldAlpha 
      Height          =   1815
      Left            =   3480
      TabIndex        =   144
      Top             =   2760
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3201
      BackColor       =   15396057
      Begin VB.ComboBox cmbDivisionAlpha 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   150
         Top             =   2640
         Width           =   3495
      End
      Begin VB.CommandButton cmdOKTaxAlpha 
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
         Left            =   360
         Picture         =   "frmPersonnelCompensation.frx":0BE4
         Style           =   1  'Graphical
         TabIndex        =   147
         Top             =   1080
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelTaxAlpha 
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
         Left            =   2040
         Picture         =   "frmPersonnelCompensation.frx":1256
         Style           =   1  'Graphical
         TabIndex        =   146
         Top             =   1080
         Width           =   1560
      End
      Begin VB.TextBox txtTaxAlphaYear 
         Height          =   315
         Left            =   1680
         TabIndex        =   145
         Top             =   600
         Width           =   1215
      End
      Begin RPVGCC.b8TitleBar b8TitleBar4 
         Height          =   345
         Left            =   40
         TabIndex        =   148
         Top             =   40
         Width           =   3890
         _ExtentX        =   6853
         _ExtentY        =   609
         Caption         =   "Tax WithHeld (Alpha List)"
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
         Icon            =   "frmPersonnelCompensation.frx":19B2
         ShadowVisible   =   0   'False
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "SELECT DIVISION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   151
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1080
         TabIndex        =   149
         Top             =   600
         Width           =   615
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11040
      Top             =   1080
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
            Picture         =   "frmPersonnelCompensation.frx":1F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensation.frx":2C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensation.frx":3900
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensation.frx":45DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensation.frx":52B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensation.frx":5F8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensation.frx":6C68
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensation.frx":7942
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensation.frx":861C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensation.frx":8EF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensation.frx":9BD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensation.frx":A8AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensation.frx":B584
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensation.frx":C25E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensation.frx":CF38
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBody 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   5820
      Left            =   120
      ScaleHeight     =   5820
      ScaleWidth      =   10500
      TabIndex        =   1
      Top             =   1200
      Width           =   10500
      Begin MSComctlLib.ListView lstDeduction 
         Height          =   735
         Left            =   6120
         TabIndex        =   161
         Top             =   1320
         Visible         =   0   'False
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1296
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DeductionType"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "DeductionKey"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "DeductionAmount"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DeductionSummKey"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   255
         Left            =   4200
         TabIndex        =   156
         Top             =   5520
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtAllow 
         Height          =   315
         Left            =   8160
         TabIndex        =   143
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAllowPerHour 
         Height          =   315
         Left            =   7320
         TabIndex        =   142
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtColaPerHour 
         Height          =   315
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   141
         Top             =   1725
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F1DA&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   0
         ScaleHeight     =   1185
         ScaleWidth      =   7905
         TabIndex        =   102
         Top             =   0
         Width           =   7935
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   107
            Top             =   120
            Width           =   6615
         End
         Begin VB.TextBox txtDept 
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   106
            Top             =   750
            Width           =   2895
         End
         Begin VB.TextBox txtPost 
            Height          =   315
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   105
            Top             =   750
            Width           =   2655
         End
         Begin VB.TextBox txtDivName 
            Height          =   315
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   103
            Top             =   435
            Width           =   2655
         End
         Begin VB.TextBox txtPayrollPeriod 
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   104
            Top             =   435
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "EMPLOYEE"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "DEPARTMENT"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   111
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "POSITION"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4200
            TabIndex        =   110
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "PERIOD"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   109
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "DIVISION"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4200
            TabIndex        =   108
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F1DA&
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   0
         ScaleHeight     =   2385
         ScaleWidth      =   3225
         TabIndex        =   77
         Top             =   1365
         Width           =   3255
         Begin VB.TextBox txtSL 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   81
            Top             =   1380
            Width           =   1335
         End
         Begin VB.TextBox txtTotalForAllowance 
            Height          =   285
            Left            =   120
            TabIndex        =   155
            Top             =   1920
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtColaHrs 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   152
            Top             =   440
            Width           =   1335
         End
         Begin VB.TextBox txtSLAmount 
            Height          =   285
            Left            =   1560
            TabIndex        =   86
            Top             =   1560
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtLHAmount 
            Height          =   285
            Left            =   1560
            TabIndex        =   85
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtSHAmount 
            Height          =   285
            Left            =   1560
            TabIndex        =   84
            Top             =   840
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtAmountEarned 
            Height          =   285
            Left            =   1560
            TabIndex        =   83
            Top             =   120
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtAdjustment 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   82
            Top             =   1700
            Width           =   1335
         End
         Begin VB.TextBox txtLH 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   80
            Top             =   1070
            Width           =   1335
         End
         Begin VB.TextBox txtSH 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   79
            Top             =   750
            Width           =   1335
         End
         Begin VB.TextBox txtNoHours 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   78
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label42 
            BackStyle       =   0  'Transparent
            Caption         =   "COLA HOURS"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   153
            Top             =   440
            Width           =   1215
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "ADJUSTMENTS"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   1700
            Width           =   1215
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "SICK LEAVE"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   1380
            Width           =   1215
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "LEGAL HOLIDAY"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   1070
            Width           =   1215
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "SPECIAL HOLIDAY"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   750
            Width           =   1455
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "NO OF HOURS"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   120
            Width           =   1215
         End
         Begin VB.Image Image2 
            Height          =   255
            Left            =   720
            Picture         =   "frmPersonnelCompensation.frx":DC12
            Stretch         =   -1  'True
            Top             =   2040
            Width           =   2175
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F1DA&
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   0
         ScaleHeight     =   1785
         ScaleWidth      =   3225
         TabIndex        =   64
         Top             =   3885
         Width           =   3255
         Begin VB.TextBox txtLHOTAmount 
            Height          =   285
            Left            =   1560
            TabIndex        =   72
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtSHOTAmount 
            Height          =   285
            Left            =   1560
            TabIndex        =   71
            Top             =   840
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtRDOTAmount 
            Height          =   285
            Left            =   1560
            TabIndex        =   70
            Top             =   480
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtRegOTAmount 
            Height          =   285
            Left            =   1560
            TabIndex        =   69
            Top             =   120
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtLHOT 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   1060
            Width           =   1335
         End
         Begin VB.TextBox txtSHOT 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   67
            Top             =   750
            Width           =   1335
         End
         Begin VB.TextBox txtRDOT 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   66
            Top             =   430
            Width           =   1335
         End
         Begin VB.TextBox txtRegOT 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   65
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "LEGAL HOLIDAY"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "SPECIAL HOLIDAY"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   750
            Width           =   1575
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "REST DAY"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   440
            Width           =   1215
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "REGULAR"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   120
            Width           =   1215
         End
         Begin VB.Image Image3 
            Height          =   255
            Left            =   840
            Picture         =   "frmPersonnelCompensation.frx":1080C
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F1DA&
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   6930
         ScaleHeight     =   2025
         ScaleWidth      =   3465
         TabIndex        =   47
         Top             =   2085
         Width           =   3495
         Begin VB.TextBox txtIsCont 
            Height          =   285
            Left            =   120
            TabIndex        =   58
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtEC 
            Height          =   285
            Left            =   1320
            TabIndex        =   57
            Top             =   360
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.PictureBox picCont 
            Appearance      =   0  'Flat
            BackColor       =   &H00E8F1DA&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Left            =   3060
            ScaleHeight     =   195
            ScaleWidth      =   375
            TabIndex        =   55
            Top             =   120
            Width           =   375
            Begin VB.CheckBox chkContribution 
               BackColor       =   &H00FDE9C6&
               Caption         =   "Check1"
               Height          =   195
               Left            =   0
               TabIndex        =   56
               Top             =   0
               Width           =   195
            End
         End
         Begin VB.TextBox txtPagIbigEmployer 
            Height          =   285
            Left            =   1560
            TabIndex        =   54
            Top             =   1080
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtPHICEmployer 
            Height          =   285
            Left            =   1560
            TabIndex        =   53
            Top             =   720
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtSSSEmployer 
            Height          =   285
            Left            =   1560
            TabIndex        =   52
            Top             =   360
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtWithHeld 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   1310
            Width           =   1455
         End
         Begin VB.TextBox txtSSS 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtPHIC 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   680
            Width           =   1455
         End
         Begin VB.TextBox txtPagIbig 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   990
            Width           =   1455
         End
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   0
            Top             =   0
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "CHECKED HERE>>>"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   63
            Top             =   120
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "TAX WITHHELD"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   1340
            Width           =   1215
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "SSS"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "PHIL HEALTH"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   680
            Width           =   1215
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "PAG IBIG"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   1020
            Width           =   1215
         End
         Begin VB.Image Image6 
            Height          =   255
            Left            =   120
            Picture         =   "frmPersonnelCompensation.frx":130D3
            Stretch         =   -1  'True
            Top             =   1700
            Width           =   3255
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F1DA&
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   3405
         ScaleHeight     =   2505
         ScaleWidth      =   3345
         TabIndex        =   34
         Top             =   2925
         Width           =   3375
         Begin VB.TextBox txtShortages 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   1060
            Width           =   1455
         End
         Begin VB.TextBox txtAROthers 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   440
            Width           =   1455
         End
         Begin VB.TextBox txtMortuary 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   120
            Width           =   1455
         End
         Begin VB.TextBox txtOthers 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   1690
            Width           =   1455
         End
         Begin VB.TextBox txtUniform 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   1380
            Width           =   1455
         End
         Begin VB.TextBox txtAdvances 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   750
            Width           =   1455
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "SHORTAGES"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   1090
            Width           =   1215
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "AR OTHERS"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   450
            Width           =   1215
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "MORTUARY"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "OTHERS"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   1720
            Width           =   1215
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "UNIFORM"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   1420
            Width           =   1455
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "ADVANCES"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   780
            Width           =   1215
         End
         Begin VB.Image Image5 
            Height          =   255
            Left            =   600
            Picture         =   "frmPersonnelCompensation.frx":1685F
            Stretch         =   -1  'True
            Top             =   2100
            Width           =   2175
         End
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F1DA&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   6930
         ScaleHeight     =   1425
         ScaleWidth      =   3465
         TabIndex        =   12
         Top             =   4245
         Width           =   3495
         Begin VB.Label Label44 
            BackStyle       =   0  'Transparent
            Caption         =   "ALLOWANCE"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   158
            Top             =   800
            Width           =   1695
         End
         Begin VB.Label lblAllowance 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1800
            TabIndex        =   157
            Top             =   800
            Width           =   1455
         End
         Begin VB.Label lblNetPayTmp 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   2730
            TabIndex        =   154
            Top             =   1050
            Width           =   525
         End
         Begin VB.Label lblCola 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1800
            TabIndex        =   140
            Top             =   540
            Width           =   1455
         End
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            Caption         =   "COLA"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   139
            Top             =   540
            Width           =   1695
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "NET PAY"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1050
            Width           =   975
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL DEDUCTIONS"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   280
            Width           =   1575
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL EARNINGS"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   30
            Width           =   1695
         End
         Begin VB.Label lblNetPay 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1680
            TabIndex        =   15
            Top             =   0
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label lblTotalDeductions 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1800
            TabIndex        =   14
            Top             =   280
            Width           =   1455
         End
         Begin VB.Label lblTotalEarnings 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1800
            TabIndex        =   13
            Top             =   30
            Width           =   1455
         End
         Begin VB.Line Line1 
            X1              =   2040
            X2              =   3240
            Y1              =   1080
            Y2              =   1080
         End
      End
      Begin VB.TextBox txtEmpPK 
         Height          =   315
         Left            =   6960
         TabIndex        =   11
         Top             =   1725
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtDeptKey 
         Height          =   315
         Left            =   7440
         TabIndex        =   10
         Top             =   1725
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtPostKey 
         Height          =   315
         Left            =   7680
         TabIndex        =   9
         Top             =   1725
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtStatusKey 
         Height          =   315
         Left            =   7920
         TabIndex        =   8
         Top             =   1725
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtRatePerHour 
         Height          =   315
         Left            =   8160
         TabIndex        =   7
         Top             =   1725
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtDivision 
         Height          =   315
         Left            =   7200
         TabIndex        =   6
         Top             =   1725
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtID 
         Height          =   315
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1725
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtPeriod 
         Height          =   315
         Left            =   8640
         TabIndex        =   4
         Top             =   1725
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtActionPK 
         Height          =   315
         Left            =   6960
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   8130
         ScaleHeight     =   1905
         ScaleWidth      =   2265
         TabIndex        =   2
         Top             =   0
         Width           =   2295
         Begin VB.Image imgPicture 
            Height          =   1905
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2265
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F1DA&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   3405
         ScaleHeight     =   1425
         ScaleWidth      =   3345
         TabIndex        =   19
         Top             =   1365
         Width           =   3375
         Begin VB.TextBox txtSSSTotalPaid 
            Height          =   285
            Left            =   1560
            TabIndex        =   30
            Top             =   360
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtPagIbigTotalPaid 
            Height          =   285
            Left            =   1560
            TabIndex        =   29
            Top             =   720
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtIsLoan 
            Height          =   285
            Left            =   2865
            TabIndex        =   28
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.PictureBox picLoan 
            Appearance      =   0  'Flat
            BackColor       =   &H00E8F1DA&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Left            =   3060
            ScaleHeight     =   195
            ScaleWidth      =   375
            TabIndex        =   26
            Top             =   120
            Width           =   375
            Begin VB.CheckBox chkLoan 
               BackColor       =   &H00CC8661&
               Caption         =   "Check1"
               Height          =   195
               Left            =   0
               TabIndex        =   27
               Top             =   0
               Width           =   195
            End
         End
         Begin VB.TextBox txtPagIbigLoanBalance 
            Height          =   285
            Left            =   1320
            TabIndex        =   25
            Top             =   720
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtPagIbigLoanNo 
            Height          =   285
            Left            =   1080
            TabIndex        =   24
            Top             =   720
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtSSSLoanBalance 
            Height          =   285
            Left            =   1320
            TabIndex        =   23
            Top             =   360
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtSSSLoanNo 
            Height          =   285
            Left            =   1080
            TabIndex        =   22
            Top             =   360
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtSSSLoan 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtPagIbigLoan 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   680
            Width           =   1455
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   0
            Top             =   0
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "CHECKED HERE>>>"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   33
            Top             =   105
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "SSS LOANS"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "PAG IBIG LOANS"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   700
            Width           =   1215
         End
         Begin VB.Image Image4 
            Height          =   255
            Left            =   960
            Picture         =   "frmPersonnelCompensation.frx":19966
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   2055
         End
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   1215
         Left            =   80
         Top             =   75
         Width           =   7935
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   1455
         Left            =   7005
         Top             =   4320
         Width           =   3495
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   2055
         Left            =   7005
         Top             =   2160
         Width           =   3495
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   2535
         Left            =   3480
         Top             =   3000
         Width           =   3375
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   1455
         Left            =   3480
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   1815
         Left            =   75
         Top             =   3960
         Width           =   3255
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   2415
         Left            =   75
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   1935
         Left            =   8205
         Top             =   90
         Width           =   2295
      End
   End
   Begin VB.PictureBox picPrint 
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   1800
      ScaleHeight     =   6975
      ScaleWidth      =   7455
      TabIndex        =   121
      Top             =   360
      Visible         =   0   'False
      Width           =   7455
      Begin RPVGCC.b8Container picPrint1 
         Height          =   6975
         Left            =   0
         TabIndex        =   122
         Top             =   0
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   12303
         BackColor       =   15396057
         Begin VB.ListBox lstResultPrint 
            Height          =   4545
            Left            =   3600
            TabIndex        =   132
            Top             =   1560
            Width           =   3735
         End
         Begin VB.CommandButton cmdOKPrint 
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
            Left            =   3840
            Picture         =   "frmPersonnelCompensation.frx":1B777
            Style           =   1  'Graphical
            TabIndex        =   130
            Top             =   6315
            Width           =   1560
         End
         Begin VB.CommandButton cmdCancelPrint 
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
            Left            =   5520
            Picture         =   "frmPersonnelCompensation.frx":1BDE9
            Style           =   1  'Graphical
            TabIndex        =   129
            Top             =   6315
            Width           =   1560
         End
         Begin VB.ListBox lstReportType 
            Height          =   4740
            Left            =   120
            TabIndex        =   128
            Top             =   720
            Width           =   3375
         End
         Begin VB.ComboBox cmbDivision 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   127
            Top             =   5760
            Width           =   3375
         End
         Begin VB.ComboBox cmbGroup 
            Height          =   315
            Left            =   3600
            Style           =   2  'Dropdown List
            TabIndex        =   126
            Top             =   720
            Width           =   3735
         End
         Begin VB.TextBox txtSearchPrint 
            Height          =   315
            Left            =   3600
            TabIndex        =   125
            Top             =   1155
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.ComboBox cmbPeriodPrint 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   124
            Top             =   6480
            Width           =   3375
         End
         Begin VB.TextBox txtTerms 
            Height          =   315
            Left            =   3600
            TabIndex        =   123
            Top             =   6360
            Visible         =   0   'False
            Width           =   150
         End
         Begin RPVGCC.b8TitleBar b8TitleBar3 
            Height          =   345
            Left            =   45
            TabIndex        =   131
            Top             =   45
            Width           =   7365
            _ExtentX        =   12991
            _ExtentY        =   609
            Caption         =   "Print"
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
            Icon            =   "frmPersonnelCompensation.frx":1C545
            ShadowVisible   =   0   'False
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "REPORT TYPE"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   136
            Top             =   480
            Width           =   3375
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "SELECT DIVISION"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   135
            Top             =   5520
            Width           =   3375
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "PAYROLL PERIOD"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   134
            Top             =   6240
            Width           =   1335
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "GROUP BY"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3600
            TabIndex        =   133
            Top             =   480
            Width           =   3375
         End
      End
   End
   Begin RPVGCC.b8Container picSearch 
      Height          =   4575
      Left            =   3600
      TabIndex        =   113
      Top             =   1200
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   8070
      BackColor       =   15396057
      Begin VB.ComboBox cmbPeriod 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   120
         Top             =   3360
         Width           =   2775
      End
      Begin VB.CommandButton cmdOKSearch 
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
         Left            =   360
         Picture         =   "frmPersonnelCompensation.frx":1ECF7
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   3840
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelSearch 
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
         Left            =   2040
         Picture         =   "frmPersonnelCompensation.frx":1F369
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   3840
         Width           =   1560
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   115
         Top             =   480
         Width           =   3735
      End
      Begin VB.ListBox lstResult 
         Height          =   2400
         Left            =   120
         TabIndex        =   114
         Top             =   840
         Width           =   3735
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   40
         TabIndex        =   118
         Top             =   40
         Width           =   3890
         _ExtentX        =   6853
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
         Icon            =   "frmPersonnelCompensation.frx":1FAC5
         ShadowVisible   =   0   'False
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "PERIOD"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   119
         Top             =   3360
         Width           =   615
      End
   End
   Begin RPVGCC.b8Container picAdd 
      Height          =   4935
      Left            =   1800
      TabIndex        =   92
      Top             =   1080
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8705
      BackColor       =   15396057
      Begin VB.PictureBox Picture10 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F1DA&
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   3960
         ScaleHeight     =   1785
         ScaleWidth      =   3225
         TabIndex        =   175
         Top             =   3000
         Width           =   3255
         Begin VB.TextBox txtRegOTAdd 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            TabIndex        =   179
            Top             =   120
            Width           =   1335
         End
         Begin VB.TextBox txtRDOTAdd 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            TabIndex        =   178
            Top             =   430
            Width           =   1335
         End
         Begin VB.TextBox txtSHOTAdd 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            TabIndex        =   177
            Top             =   750
            Width           =   1335
         End
         Begin VB.TextBox txtLHOTAdd 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            TabIndex        =   176
            Top             =   1060
            Width           =   1335
         End
         Begin VB.Image Image7 
            Height          =   255
            Left            =   840
            Picture         =   "frmPersonnelCompensation.frx":2005F
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label53 
            BackStyle       =   0  'Transparent
            Caption         =   "REGULAR"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   183
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label52 
            BackStyle       =   0  'Transparent
            Caption         =   "REST DAY"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   182
            Top             =   440
            Width           =   1215
         End
         Begin VB.Label Label51 
            BackStyle       =   0  'Transparent
            Caption         =   "SPECIAL HOLIDAY"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   181
            Top             =   750
            Width           =   1575
         End
         Begin VB.Label Label50 
            BackStyle       =   0  'Transparent
            Caption         =   "LEGAL HOLIDAY"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   180
            Top             =   1080
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F1DA&
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   3960
         ScaleHeight     =   2385
         ScaleWidth      =   3225
         TabIndex        =   162
         Top             =   480
         Width           =   3255
         Begin VB.TextBox txtNoHoursAdd 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            TabIndex        =   168
            Top             =   120
            Width           =   1335
         End
         Begin VB.TextBox txtSHAdd 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            TabIndex        =   167
            Top             =   750
            Width           =   1335
         End
         Begin VB.TextBox txtLHAdd 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            TabIndex        =   166
            Top             =   1070
            Width           =   1335
         End
         Begin VB.TextBox txtSLAdd 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            TabIndex        =   165
            Top             =   1380
            Width           =   1335
         End
         Begin VB.TextBox txtAdjustmentAdd 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            TabIndex        =   164
            Top             =   1700
            Width           =   1335
         End
         Begin VB.TextBox txtColaHrsAdd 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            TabIndex        =   163
            Top             =   440
            Width           =   1335
         End
         Begin VB.Image Image1 
            Height          =   255
            Left            =   720
            Picture         =   "frmPersonnelCompensation.frx":22926
            Stretch         =   -1  'True
            Top             =   2040
            Width           =   2175
         End
         Begin VB.Label Label49 
            BackStyle       =   0  'Transparent
            Caption         =   "NO OF HOURS"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   174
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label48 
            BackStyle       =   0  'Transparent
            Caption         =   "SPECIAL HOLIDAY"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   173
            Top             =   750
            Width           =   1455
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   "LEGAL HOLIDAY"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   172
            Top             =   1070
            Width           =   1215
         End
         Begin VB.Label Label46 
            BackStyle       =   0  'Transparent
            Caption         =   "SICK LEAVE"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   171
            Top             =   1380
            Width           =   1215
         End
         Begin VB.Label Label45 
            BackStyle       =   0  'Transparent
            Caption         =   "ADJUSTMENTS"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   170
            Top             =   1700
            Width           =   1215
         End
         Begin VB.Label Label43 
            BackStyle       =   0  'Transparent
            Caption         =   "COLA HOURS"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   169
            Top             =   440
            Width           =   1215
         End
      End
      Begin VB.OptionButton optInactive 
         BackColor       =   &H00EAECD9&
         Caption         =   "Inactive"
         Height          =   255
         Left            =   2880
         TabIndex        =   160
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optActive 
         BackColor       =   &H00EAECD9&
         Caption         =   "Active"
         Height          =   255
         Left            =   1800
         TabIndex        =   159
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtTo 
         Height          =   315
         Left            =   2520
         TabIndex        =   100
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox txtFrom 
         Height          =   315
         Left            =   840
         TabIndex        =   98
         Top             =   3720
         Width           =   1215
      End
      Begin VB.ListBox lstResultAdd 
         Height          =   2400
         Left            =   120
         TabIndex        =   97
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox txtSearchAdd 
         Height          =   315
         Left            =   120
         TabIndex        =   96
         Top             =   840
         Width           =   3735
      End
      Begin VB.CommandButton cmdCancelAdd 
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
         Left            =   2040
         Picture         =   "frmPersonnelCompensation.frx":25520
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   4200
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKAdd 
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
         Left            =   360
         Picture         =   "frmPersonnelCompensation.frx":25C7C
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   4200
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   93
         Top             =   45
         Width           =   7245
         _ExtentX        =   12779
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
         Icon            =   "frmPersonnelCompensation.frx":262EE
         ShadowVisible   =   0   'False
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TO"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   101
         Top             =   3720
         Width           =   255
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "PERIOD"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   99
         Top             =   3720
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmPersonnelCompensation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Public lngSSSLoanNo         As Long
Public dblSSSTotalLoan      As Double
Public dblSSSMonthly        As Double
Public dblSSSTotalPaid      As Double
Public strSSSLoanInfo '       As String
Public strPagIbigLoanInfo '   As String
Public lngPagIbigLoanNo     As Long
Public dblPagIbigTotalLoan  As Double
Public dblPagIbigMonthly    As Double
Public dblPagIbigTotalPaid  As Double

Public TRANSACTIONTYPE      As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2
Const is_FINDING = 3

Dim dSalaryWages            As Double
Dim dOvertime               As Double
Dim dColaAllowance          As Double
Dim dAREmployee             As Double
Dim Filename                As String
Dim WorkbookName            As String
Dim iWorkSheet              As Integer
Dim tmp                     As Long

Dim Arr, Arr1, Arr2, iCompKey, i, x, iMonth1, iMonth2, iMonth3, iYear1, iYear2, iYear3, dtmDateTo, sDeptName, _
RowCnt, ColCnt, strRange, iReset, strGrossTot, staTaxTot, strSSSTot, strPHICTot, strHDMFTot, _
strColaTot, strAllowTot, strValue, iCnt, strRange1, strRange2, sTaxStatus, j, k, PK, iDivision, _
iStatus, Daily, Percent, iPayee, iGLAmount, dRate, dRatePerHour, iCompensationRate, tmpAmt, tmpGLAccnt, _
iWithSL, dDebit, dCredit, dtmPayrollDate, dblSSSDif, dblPagIbigDiff, Array1, _
dtmPayrollDate1, dblPreviousGross, dblGrossForCont, dblSSSEmployer, dblSSSEmployee, dblSSSEC, dblPHICEmployer, _
dblPHICEmployee, dblPagIbigEmployer, dblPagIbigEmployee, dblBasicForTax, dblTaxExemp, dblPercent, dblConstant, _
dblBracketAmount, var1, var2, var3, dtmTo, iEmployeeKey, sInRefKey, dDedAmt, dDedTotAmt

Private Function BROWSER(strCtrl, is_Action As String)
'MsgBox IIf(IsNumeric(strCtrl) = False, 0, strCtrl)
Select Case is_Action
    Case "is_LOAD"
        If strCtrl <> "" Then
            s = "sp_Personnel_Compensation_Browse(" & IIf(IsNumeric(strCtrl) = False, 0, strCtrl) & ",0)"
        Else
            s = "sp_Personnel_Compensation_Browse(" & IIf(IsNumeric(strCtrl) = False, 0, strCtrl) & ",1)"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "sp_Personnel_Compensation_Browse(" & IIf(IsNumeric(strCtrl) = False, 0, strCtrl) & ",1)"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "sp_Personnel_Compensation_Browse(" & IIf(IsNumeric(strCtrl) = False, 0, strCtrl) & ",2)"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "sp_Personnel_Compensation_Browse(" & IIf(IsNumeric(strCtrl) = False, 0, strCtrl) & ",3)"
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "sp_Personnel_Compensation_Browse(" & IIf(IsNumeric(strCtrl) = False, 0, strCtrl) & ",4)"
    Case Else: Exit Function
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtEmpPK.Text = rs!EmpPK
    txtDivision.Text = rs!Division
    txtDeptKey.Text = rs!Dept
    txtStatusKey.Text = rs!Status
    txtPostKey.Text = rs!Positions
    txtActionPK.Text = rs!ActionMemo
    txtPeriod.Text = rs!Period
    txtRatePerHour.Text = rs!RatePerHour
    txtColaPerHour.Text = rs!ColaPerHour
    txtAllowPerHour.Text = rs!AllowPerHour
    txtName.Text = rs!IDName
    txtPayrollPeriod.Text = Format(rs!DateFrom, "mm/dd/yyyy") & " - " & Format(rs!DateTo, "mm/dd/yyyy")
    txtDept.Text = IIf(IsNull(rs!DepartmentName), "", rs!DepartmentName)
    txtPost.Text = rs!PositionName
    txtNoHours.Text = rs!NoHours 'Format(rs!NoHours, "#,##0.00")
    txtColaHrs.Text = Format(rs!ColaHours, "#,##0.00")
    txtSH.Text = Format(rs!SH_Hours, "#,##0.00")
    txtLH.Text = Format(rs!LH_Hours, "#,##0.00")
    txtSL.Text = Format(rs!SL_Hours, "#,##0.00")
    txtAdjustment.Text = Format(rs!Adjustment, "#,##0.00")
    txtRegOT.Text = Format(rs!Reg_OT_Hours, "#,##0.00")
    txtRDOT.Text = Format(rs!RD_OT_Hours, "#,##0.00")
    txtSHOT.Text = Format(rs!SH_OT_Hours, "#,##0.00")
    txtLHOT.Text = Format(rs!LH_OT_Hours, "#,##0.00")
    txtAmountEarned.Text = rs!Amount_Earned
    txtSHAmount.Text = rs!SH_Amount
    txtLHAmount.Text = rs!LH_Amount
    txtSLAmount.Text = rs!SL_Amount
    txtRegOTAmount.Text = rs!Reg_OT_Amount
    txtRDOTAmount.Text = rs!RD_OT_Amount
    txtSHOTAmount.Text = rs!SH_OT_Amount
    txtLHOTAmount.Text = rs!LH_OT_Amount
    lblTotalEarnings.Caption = Format(rs!TotalEarning, "#,##0.00")
    txtMortuary.Text = Format(rs!Mortuary, "#,##0.00")
    txtAROthers.Text = Format(rs!AR_Others, "#,##0.00")
    txtAdvances.Text = Format(rs!Advances, "#,##0.00")
    txtShortages.Text = Format(rs!Shortages, "#,##0.00")
    txtUniform.Text = Format(rs!Uniforms, "#,##0.00")
    txtOthers.Text = Format(rs!Others, "#,##0.00")
    txtIsLoan.Text = rs!Is_Have_Loan
    chkLoan.Value = rs!Is_Have_Loan
    txtSSSLoanNo.Text = rs!SSSLoan_No
    txtSSSLoan.Text = Format(rs!SSSLoan, "#,##0.00")
    txtSSSLoanBalance.Text = rs!SSSBalance
    txtPagIbigLoanNo.Text = rs!PagIbigLoan_No
    txtPagIbigLoan.Text = Format(rs!PagIbigLoan, "#,##0.00")
    txtPagIbigLoanBalance.Text = rs!PagIbigBalance
    txtIsCont.Text = rs!Is_Have_Cont
    chkContribution.Value = rs!Is_Have_Cont
    txtSSS.Text = Format(rs!SSS, "#,##0.00")
    txtSSSEmployer.Text = rs!SSS_Employer
    txtEC.Text = rs!SSS_EC
    txtPHIC.Text = Format(rs!PHIC, "#,##0.00")
    txtPHICEmployer.Text = rs!PHIC_Employer
    txtPagIbig.Text = Format(rs!PagIbig, "#,##0.00")
    txtPagIbigEmployer.Text = rs!PagIbig_Employer
    txtWithHeld.Text = Format(rs!WithHeld, "#,##0.00")
    lblTotalDeductions.Caption = Format(rs!TotalDeduction, "#,##0.00")
    lblCola.Caption = Format(rs!TotalCola, "#,##0.00")
    lblNetPay.Caption = Format(rs!NetEarning, "#,##0.00")
    txtAllow.Text = rs!TotalAllowance
    
    t = "SELECT ProfileKey " & _
        " FROM tbl_Personnel_IDNumber " & _
        " WHERE (PK = " & rs!EmpPK & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        imgPicture.Picture = LoadPicture(SHOW_IMAGES(rt!ProfileKey, 0, "Employee Profile"))
    Else
        imgPicture.Picture = LoadPicture("")
    End If
    rt.Close
    
    lstDeduction.ListItems.Clear
    
    t = "SELECT tbl_Personnel_Compensation_Deduction.* " & _
        " FROM tbl_Personnel_Compensation_Deduction " & _
        " WHERE (CompensationKey = " & rs!PrimaryKey & ") " & _
        " ORDER BY DeductionType, DeductionKey"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        Set x = lstDeduction.ListItems.Add()
        x.Text = ""
        x.SubItems(1) = rt!DeductionType
        x.SubItems(2) = rt!DeductionKey
        x.SubItems(3) = rt!Amount
        x.SubItems(4) = rt!DeductionSummKey
        rt.MoveNext
    Wend
    rt.Close
    
    StatusBar.Panels(1).Text = rs!PrimaryKey
    StatusBar.Panels(2).Text = "LAST MODIFIED BY : " & rs!PayLastMod
    StatusBar.Panels(3).Text = "" '"RATE: " & Format(GET_BASIC_RATE(rs!ActionMemo), "#,##0.00")
    StatusBar.Panels(4).Text = IIf(rs!Locked = 1, "LOCKED", "UNLOCKED")
    
    SaveSetting App.EXEName, "CompensationC", "CompC", rs!PrimaryKey
    
End If
rs.Close
End Function

Private Function GET_BASIC_RATE(intPK) As Double
s = "SELECT Basic" & _
    " From tbl_Personnel_Action " & _
    " WHERE (PK = " & intPK & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    GET_BASIC_RATE = rs!Basic
End If
rs.Close
End Function

Private Function PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If picAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
If picPrint.Visible = True Then Exit Function

If AccessRights("Personnel Compensation", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If

txtNoHoursAdd.Text = ""
txtColaHrsAdd.Text = ""
txtSHAdd.Text = ""
txtLHAdd.Text = ""
txtSLAdd.Text = ""
txtAdjustmentAdd.Text = ""
txtRegOTAdd.Text = ""
txtRDOTAdd.Text = ""
txtSHOTAdd.Text = ""
txtLHOTAdd.Text = ""

picToolbar.Enabled = False
picBody.Enabled = False
txtSearchAdd.Text = ""
txtFrom.Text = ""
txtTo.Text = ""
picAdd.ZOrder 0
picAdd.Visible = True
optActive.SetFocus
'txtSearchAdd.SetFocus

End Function

Private Function PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If StatusBar.Panels(1).Text = "" Then Exit Function
If picAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
If picPrint.Visible = True Then Exit Function
If AccessRights("Personnel Compensation", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If

If StatusBar.Panels(4).Text = "LOCKED" Then MsgBox "TRANSACTION ALREADY LOCKED!                         ", vbCritical, "Error...": Exit Function

TRANSACTIONTYPE = is_EDITTING
txtTotalForAllowance.Text = RETURNTEXTVALUE(txtNoHours) + _
                            RETURNTEXTVALUE(txtSH) + _
                            RETURNTEXTVALUE(txtLH) + _
                            RETURNTEXTVALUE(txtSL)

TOOLBARFUNC 2
LOCKTEXT False
txtNoHours.SetFocus
End Function

Private Function PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If StatusBar.Panels(1).Text = "" Then Exit Function
If picAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
If picPrint.Visible = True Then Exit Function
If AccessRights("Personnel Compensation", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If

If StatusBar.Panels(4).Text = "LOCKED" Then MsgBox "TRANSACTION ALREADY LOCKED!                         ", vbCritical, "Error...": Exit Function

If MsgBox("ARE YOU TO DELETE THIS  RECORD?      ", vbCritical + vbYesNo, "CONFIMATION") = vbNo Then Exit Function

iPayee = 0
t = "SELECT ProfileKey " & _
    " FROM tbl_Personnel_IDNumber " & _
    " WHERE (PK = " & RETURNTEXTVALUE(txtEmpPK) & ")"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then
    iPayee = rt!ProfileKey
End If
rt.Close

Arr = Split(Trim(txtPayrollPeriod.Text), " - ", -1, 1)
ConnOmega.Execute "DELETE FROM tbl_GL_Transaction " & _
                  " WHERE (PayeeType = 3) " & _
                  " AND (PayeeKey = " & iPayee & ") " & _
                  " AND (DocDate = '" & FormatDateTime(Arr(1), vbShortDate) & "') " & _
                  " AND (DocNumber = '" & "PYRL" & Format(Arr(1), "mmddyy") & "')"
                  
ConnOmega.Execute "DELETE FROM tbl_Personnel_SL " & _
                  " WHERE (iType = 3) " & _
                  " AND (EmployeeKey = " & iPayee & ") " & _
                  " AND (DocDate = '" & FormatDateTime(Arr(1), vbShortDate) & "') " & _
                  " AND (DocNumber = '" & "PYRL" & Format(Arr(1), "mmddyy") & "')"

For i = 1 To lstDeduction.ListItems.Count
    s = "SELECT tbl_Personnel_Deduction_Summary.* " & _
        " FROM tbl_Personnel_Deduction_Summary " & _
        " WHERE (OutRefKey = " & StatusBar.Panels(1).Text & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        t = "SELECT tbl_Personnel_Deduction_Summary.* " & _
            " FROM tbl_Personnel_Deduction_Summary " & _
            " WHERE (PK = " & rs!RefKey & ")"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            ConnOmega.Execute "UPDATE tbl_Personnel_Deduction_Summary " & _
                              " SET AmountUsed = AmountUsed - " & CDbl(rs!AmountOut) & ", " & _
                              " Cleared = " & IIf(CDbl(rt!AmountIn) > (CDbl(rt!AmountUsed) - CDbl(rs!AmountOut)), 0, 1) & " " & _
                              " WHERE (PK = " & rt!PK & ")"
        End If
        rt.Close
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Deduction_Summary " & _
                          " WHERE (PK = " & rs!PK & ")"
    End If
    rs.Close
Next i

ConnOmega.Execute "DELETE tbl_Personnel_Compensation_Deduction WHERE (CompensationKey = " & StatusBar.Panels(1).Text & ")"

ConnOmega.Execute "DELETE FROM tbl_Personnel_Compensation" & _
                  " WHERE (PK = " & StatusBar.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "CompensationC", "CompC", ""), "is_PAGEDOWN"
If StatusBar.Panels(1).Text = "" Then BROWSER GetSetting(App.EXEName, "CompensationC", "CompC", ""), "is_HOME"

End Function

Private Sub COMPUTE_ALLOWANCE(iEmp, iPeriod, iHour, dEffectDate, iDiv, sLastMod)
dRate = 0
iCompensationRate = 0

s = "SELECT TOP 1 CompensationRate " & _
    " From tbl_Personnel_Action " & _
    " WHERE (EmpPK = " & iEmp & ") " & _
    " AND (EffectivityDate <= '" & FormatDateTime(dEffectDate, vbShortDate) & "') " & _
    " ORDER BY EffectivityDate DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iCompensationRate = rs!CompensationRate
End If
rs.Close

's = "SELECT TOP 1 Rate, RatePerHour " & _
    " FROM tbl_Personnel_Allowance " & _
    " WHERE (EmpPK = " & iEmp & ") " & _
    " AND (EffectDate <= '" & FormatDateTime(dEffectDate, vbShortDate) & "') " & _
    " AND (Rate > 0) " & _
    " ORDER BY EffectDate DESC"
s = "SELECT TOP 1 Rate, RatePerHour " & _
    " FROM tbl_Personnel_Allowance " & _
    " WHERE (EmpPK = " & iEmp & ") " & _
    " AND (EffectDate <= '" & FormatDateTime(dEffectDate, vbShortDate) & "') " & _
    " ORDER BY EffectDate DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    If CInt(iCompensationRate) = 1 Then
        dRate = ((CDbl(rs!Rate) / 2) / 13.08333) / 8
    ElseIf CInt(iCompensationRate) = 2 Then
        dRate = CDbl(rs!Rate) / 8
    End If
End If
rs.Close

ConnOmega.Execute "DELETE FROM tbl_Personnel_Allowance_Per_Period " & _
                  " WHERE (EmpPK = " & iEmp & ") " & _
                  " AND (Period = " & iPeriod & ")"

If CDbl(dRate) <= 0 Then Exit Sub

ConnOmega.Execute "INSERT INTO tbl_Personnel_Allowance_Per_Period " & _
                  " (EmpPK, Period, Division, NoHours, RatePerHour, LastModified) " & _
                  " VALUES (" & iEmp & ", " & iPeriod & ", " & iDiv & ", " & iHour & ", " & _
                  " " & dRate & ", '" & sLastMod & "')"

End Sub

Private Function PRESS_F5()
If RETURNLABELVALUE(lblTotalEarnings) = 0 Then
    MsgBox "TOTAL EARNING MUST BE HIGHER THAN ZERO!         ", vbInformation, ""
    txtNoHours.SetFocus
    HTEXT txtNoHours
    Exit Function
End If
If CInt(txtIsLoan.Text) <> chkLoan.Value Then
    MsgBox "LOANS MUST BE CHECKED!          ", vbInformation, ""
    Label2.Visible = True
    Timer1.Enabled = True
    chkLoan.SetFocus
    Exit Function
End If
If CInt(txtIsCont.Text) <> chkContribution.Value Then
    MsgBox "CONTRIBUTIONS MUST BE CHECKED!          ", vbInformation, ""
    Label31.Visible = True
    Timer2.Enabled = True
    chkContribution.SetFocus
    Exit Function
End If
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    If CHECK_DUPLICATE_ENTRY(txtEmpPK.Text, txtPeriod.Text) = True Then MsgBox "FOUND DUPLICATE ENRTY...!       ", vbExclamation, "Alert..": Exit Function
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Compensation" & _
                      " (EmpPK, Division, Dept, Status, Positions, Period, RatePerHour, ActionMemo, NoHours, SH_Hours, LH_Hours, SL_Hours, Adjustment, Reg_OT_Hours, RD_OT_Hours, SH_OT_Hours, " & _
                      " LH_OT_Hours, Amount_Earned, SH_Amount, LH_Amount, SL_Amount, Reg_OT_Amount, RD_OT_Amount, SH_OT_Amount, LH_OT_Amount, TotalEarning, Mortuary, AR_Others, Advances, Shortages, " & _
                      " Uniforms, Others, Is_Have_Loan, SSSLoan_No, SSSLoan, SSSBalance, PagIbigLoan_No, PagIbigLoan, PagIbigBalance, Is_Have_Cont, SSS, SSS_Employer, SSS_EC, PHIC, PHIC_Employer, " & _
                      " PagIbig, PagIbig_Employer, WithHeld, TotalDeduction, NetEarning, LastModified, ColaPerHour, TotalCola, AllowPerHour, TotalAllowance, ColaHours) " & _
                      " VALUES (" & RETURNTEXTVALUE(txtEmpPK) & ", " & RETURNTEXTVALUE(txtDivision) & ", " & RETURNTEXTVALUE(txtDeptKey) & ", " & RETURNTEXTVALUE(txtStatusKey) & ", " & _
                      " " & RETURNTEXTVALUE(txtPostKey) & ", " & RETURNTEXTVALUE(txtPeriod) & ", " & RETURNTEXTVALUE(txtRatePerHour) & ", " & RETURNTEXTVALUE(txtActionPK) & ", " & _
                      " " & RETURNTEXTVALUE(txtNoHours) & ", " & RETURNTEXTVALUE(txtSH) & ", " & RETURNTEXTVALUE(txtLH) & ", " & RETURNTEXTVALUE(txtSL) & ", " & RETURNTEXTVALUE(txtAdjustment) & ", " & _
                      " " & RETURNTEXTVALUE(txtRegOT) & ", " & RETURNTEXTVALUE(txtRDOT) & ", " & RETURNTEXTVALUE(txtSHOT) & ", " & RETURNTEXTVALUE(txtLHOT) & ", " & RETURNTEXTVALUE(txtAmountEarned) & ", " & _
                      " " & RETURNTEXTVALUE(txtSHAmount) & ", " & RETURNTEXTVALUE(txtLHAmount) & ", " & RETURNTEXTVALUE(txtSLAmount) & ", " & RETURNTEXTVALUE(txtRegOTAmount) & ", " & _
                      " " & RETURNTEXTVALUE(txtRDOTAmount) & ", " & RETURNTEXTVALUE(txtSHOTAmount) & ", " & RETURNTEXTVALUE(txtLHOTAmount) & ", " & RETURNLABELVALUE(lblTotalEarnings) & ", " & _
                      " " & RETURNTEXTVALUE(txtMortuary) & ", " & RETURNTEXTVALUE(txtAROthers) & ", " & RETURNTEXTVALUE(txtAdvances) & ", " & RETURNTEXTVALUE(txtShortages) & ", " & RETURNTEXTVALUE(txtUniform) & ", " & _
                      " " & RETURNTEXTVALUE(txtOthers) & ", " & RETURNTEXTVALUE(txtIsLoan) & ", " & RETURNTEXTVALUE(txtSSSLoanNo) & ", " & RETURNTEXTVALUE(txtSSSLoan) & ", " & RETURNTEXTVALUE(txtSSSLoanBalance) & ", " & _
                      " " & RETURNTEXTVALUE(txtPagIbigLoanNo) & ", " & RETURNTEXTVALUE(txtPagIbigLoan) & ", " & RETURNTEXTVALUE(txtPagIbigLoanBalance) & ", " & RETURNTEXTVALUE(txtIsCont) & ", " & _
                      " " & RETURNTEXTVALUE(txtSSS) & ", " & RETURNTEXTVALUE(txtSSSEmployer) & ", " & RETURNTEXTVALUE(txtEC) & ", " & RETURNTEXTVALUE(txtPHIC) & ", " & RETURNTEXTVALUE(txtPHICEmployer) & ", " & _
                      " " & RETURNTEXTVALUE(txtPagIbig) & ", " & RETURNTEXTVALUE(txtPagIbigEmployer) & ", " & RETURNTEXTVALUE(txtWithHeld) & ", " & RETURNLABELVALUE(lblTotalDeductions) & ", " & _
                      " " & RETURNLABELVALUE(lblNetPay) & ", '" & CStr(Now) & " - " & gbl_CompleteName & "', " & RETURNTEXTVALUE(txtColaPerHour) & ", " & RETURNLABELVALUE(lblCola) & ", " & RETURNTEXTVALUE(txtAllowPerHour) & ", " & _
                      " " & RETURNTEXTVALUE(txtAllow) & ", " & RETURNTEXTVALUE(txtColaHrs) & ")"
    
    '=== Allowance
    Arr = Split(Trim(txtPayrollPeriod.Text), " - ", -1, 1)
    COMPUTE_ALLOWANCE RETURNTEXTVALUE(txtEmpPK), RETURNTEXTVALUE(txtPeriod), RETURNTEXTVALUE(txtNoHours) + RETURNTEXTVALUE(txtSL), Arr(1), RETURNTEXTVALUE(txtDivision), CStr(Now) & " - " & gbl_CompleteName
    
    iCompKey = 0
    s = "SELECT PK " & _
        " FROM tbl_Personnel_Compensation " & _
        " WHERE (EmpPK = " & RETURNTEXTVALUE(txtEmpPK) & ") " & _
        " AND (Period = " & RETURNTEXTVALUE(txtPeriod) & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        iCompKey = rs!PK
    End If
    rs.Close
    
    Arr = Split(Trim(txtPayrollPeriod.Text), " - ", -1, 1)
    For i = 1 To lstDeduction.ListItems.Count
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Compensation_Deduction " & _
                          " (CompensationKey, DeductionType, DeductionKey, DeductionSummKey, Amount) " & _
                          " VALUES (" & iCompKey & ", " & lstDeduction.ListItems.Item(i).SubItems(1) & ", " & _
                          " " & lstDeduction.ListItems.Item(i).SubItems(2) & ", " & _
                          " " & lstDeduction.ListItems.Item(i).SubItems(4) & ", " & _
                          " " & CDbl(lstDeduction.ListItems.Item(i).SubItems(3)) & ")"
        s = "SELECT tbl_Personnel_Deduction_Summary.* " & _
            " FROM tbl_Personnel_Deduction_Summary " & _
            " WHERE (PK = " & lstDeduction.ListItems.Item(i).SubItems(4) & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            ConnOmega.Execute "UPDATE tbl_Personnel_Deduction_Summary " & _
                              " SET AmountUsed = AmountUsed + " & CDbl(lstDeduction.ListItems.Item(i).SubItems(3)) & ", " & _
                              " Cleared = " & IIf((CDbl(lstDeduction.ListItems.Item(i).SubItems(3)) + CDbl(rs!AmountUsed)) >= CDbl(rs!AmountIn), 1, 0) & " " & _
                              " WHERE (PK = " & lstDeduction.ListItems.Item(i).SubItems(4) & ")"
        End If
        rs.Close
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Deduction_Summary " & _
                          " (EmployeeKey, DeductionType, TransDate, InOut, " & _
                          " OutRefKey, AmountOut, AmountUsed, RefKey, DocNumber) " & _
                          " VALUES (" & RETURNTEXTVALUE(txtEmpPK) & ", " & _
                          " " & lstDeduction.ListItems.Item(i).SubItems(1) & ", " & _
                          " '" & FormatDateTime(Arr(1), vbShortDate) & "', " & _
                          " 'O', " & iCompKey & ", " & CDbl(lstDeduction.ListItems.Item(i).SubItems(3)) & ", " & _
                          " " & CDbl(lstDeduction.ListItems.Item(i).SubItems(3)) & ", " & _
                          " " & lstDeduction.ListItems.Item(i).SubItems(4) & ", " & _
                          " '" & "Payroll Period " & Trim(txtPayrollPeriod.Text) & "')"
    Next i
End If

If TRANSACTIONTYPE = is_EDITTING Then
    iCompKey = StatusBar.Panels(1).Text
    ConnOmega.Execute "UPDATE tbl_Personnel_Compensation" & _
                      " SET NoHours = " & RETURNTEXTVALUE(txtNoHours) & ", " & _
                      " SH_Hours = " & RETURNTEXTVALUE(txtSH) & ", LH_Hours = " & RETURNTEXTVALUE(txtLH) & ", SL_Hours = " & RETURNTEXTVALUE(txtSL) & ", " & _
                      " Adjustment = " & RETURNTEXTVALUE(txtAdjustment) & ", Reg_OT_Hours = " & RETURNTEXTVALUE(txtRegOT) & ", RD_OT_Hours = " & RETURNTEXTVALUE(txtRDOT) & ", " & _
                      " SH_OT_Hours = " & RETURNTEXTVALUE(txtSHOT) & ", LH_OT_Hours = " & RETURNTEXTVALUE(txtLHOT) & ", Amount_Earned = " & RETURNTEXTVALUE(txtAmountEarned) & ", " & _
                      " SH_Amount = " & RETURNTEXTVALUE(txtSHAmount) & ", LH_Amount = " & RETURNTEXTVALUE(txtLHAmount) & ", SL_Amount = " & RETURNTEXTVALUE(txtSLAmount) & ", " & _
                      " Reg_OT_Amount = " & RETURNTEXTVALUE(txtRegOTAmount) & ", RD_OT_Amount = " & RETURNTEXTVALUE(txtRDOTAmount) & ", SH_OT_Amount = " & RETURNTEXTVALUE(txtSHOTAmount) & ", " & _
                      " LH_OT_Amount = " & RETURNTEXTVALUE(txtLHOTAmount) & ", TotalEarning = " & RETURNLABELVALUE(lblTotalEarnings) & ", Mortuary = " & RETURNTEXTVALUE(txtMortuary) & ", " & _
                      " AR_Others = " & RETURNTEXTVALUE(txtAROthers) & ", Advances = " & RETURNTEXTVALUE(txtAdvances) & ", Shortages = " & RETURNTEXTVALUE(txtShortages) & ", " & _
                      " Uniforms = " & RETURNTEXTVALUE(txtUniform) & ", Others = " & RETURNTEXTVALUE(txtOthers) & ", Is_Have_Loan = " & RETURNTEXTVALUE(txtIsLoan) & ", " & _
                      " SSSLoan_No = " & RETURNTEXTVALUE(txtSSSLoanNo) & ", SSSLoan = " & RETURNTEXTVALUE(txtSSSLoan) & ", SSSBalance = " & RETURNTEXTVALUE(txtSSSLoanBalance) & ", " & _
                      " PagIbigLoan_No = " & RETURNTEXTVALUE(txtPagIbigLoanNo) & ", PagIbigLoan = " & RETURNTEXTVALUE(txtPagIbigLoan) & ", PagIbigBalance = " & RETURNTEXTVALUE(txtPagIbigLoanBalance) & ", " & _
                      " Is_Have_Cont = " & RETURNTEXTVALUE(txtIsCont) & ", SSS = " & RETURNTEXTVALUE(txtSSS) & ", SSS_Employer = " & RETURNTEXTVALUE(txtSSSEmployer) & ", SSS_EC = " & RETURNTEXTVALUE(txtEC) & ", " & _
                      " PHIC = " & RETURNTEXTVALUE(txtPHIC) & ", PHIC_Employer = " & RETURNTEXTVALUE(txtPHICEmployer) & ", PagIbig = " & RETURNTEXTVALUE(txtPagIbig) & ", " & _
                      " PagIbig_Employer = " & RETURNTEXTVALUE(txtPagIbigEmployer) & ", WithHeld = " & RETURNTEXTVALUE(txtWithHeld) & ", TotalDeduction = " & RETURNLABELVALUE(lblTotalDeductions) & ", " & _
                      " NetEarning = " & RETURNLABELVALUE(lblNetPay) & ", LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "', ColaPerHour = " & RETURNTEXTVALUE(txtColaPerHour) & ", " & _
                      " TotalCola = " & RETURNLABELVALUE(lblCola) & ", AllowPerHour = " & RETURNTEXTVALUE(txtAllowPerHour) & ", TotalAllowance = " & RETURNTEXTVALUE(txtAllow) & ", " & _
                      " ColaHours = " & RETURNTEXTVALUE(txtColaHrs) & " " & _
                      " WHERE (PK = " & iCompKey & ")"
    
    Arr = Split(Trim(txtPayrollPeriod.Text), " - ", -1, 1)
    
    COMPUTE_ALLOWANCE RETURNTEXTVALUE(txtEmpPK), RETURNTEXTVALUE(txtPeriod), RETURNTEXTVALUE(txtNoHours) + RETURNTEXTVALUE(txtSL), Arr(1), RETURNTEXTVALUE(txtDivision), CStr(Now) & " - " & gbl_CompleteName
    
    For i = 1 To lstDeduction.ListItems.Count
        s = "SELECT tbl_Personnel_Deduction_Summary.* " & _
            " FROM tbl_Personnel_Deduction_Summary " & _
            " WHERE (OutRefKey = " & iCompKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            t = "SELECT tbl_Personnel_Deduction_Summary.* " & _
                " FROM tbl_Personnel_Deduction_Summary " & _
                " WHERE (PK = " & rs!RefKey & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                ConnOmega.Execute "UPDATE tbl_Personnel_Deduction_Summary " & _
                                  " SET AmountUsed = AmountUsed - " & CDbl(rs!AmountOut) & ", " & _
                                  " Cleared = " & IIf(CDbl(rt!AmountIn) > (CDbl(rt!AmountUsed) - CDbl(rs!AmountOut)), 0, 1) & " " & _
                                  " WHERE (PK = " & rt!PK & ")"
            End If
            rt.Close
        End If
        rs.Close
    Next i
    
    ConnOmega.Execute "DELETE tbl_Personnel_Compensation_Deduction WHERE (CompensationKey = " & iCompKey & ")"
    Arr = Split(Trim(txtPayrollPeriod.Text), " - ", -1, 1)
    For i = 1 To lstDeduction.ListItems.Count
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Compensation_Deduction " & _
                          " (CompensationKey, DeductionType, DeductionKey, DeductionSummKey, Amount) " & _
                          " VALUES (" & iCompKey & ", " & lstDeduction.ListItems.Item(i).SubItems(1) & ", " & _
                          " " & lstDeduction.ListItems.Item(i).SubItems(2) & ", " & _
                          " " & lstDeduction.ListItems.Item(i).SubItems(4) & ", " & _
                          " " & CDbl(lstDeduction.ListItems.Item(i).SubItems(3)) & ")"
        s = "SELECT tbl_Personnel_Deduction_Summary.* " & _
            " FROM tbl_Personnel_Deduction_Summary " & _
            " WHERE (PK = " & lstDeduction.ListItems.Item(i).SubItems(4) & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            ConnOmega.Execute "UPDATE tbl_Personnel_Deduction_Summary " & _
                              " SET AmountUsed = AmountUsed + " & CDbl(lstDeduction.ListItems.Item(i).SubItems(3)) & ", " & _
                              " Cleared = " & IIf((CDbl(lstDeduction.ListItems.Item(i).SubItems(3)) + CDbl(rs!AmountUsed)) >= CDbl(rs!AmountIn), 1, 0) & " " & _
                              " WHERE (PK = " & lstDeduction.ListItems.Item(i).SubItems(4) & ")"
        End If
        rs.Close
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Deduction_Summary " & _
                          " (EmployeeKey, DeductionType, TransDate, InOut, " & _
                          " OutRefKey, AmountOut, AmountUsed, RefKey, DocNumber) " & _
                          " VALUES (" & RETURNTEXTVALUE(txtEmpPK) & ", " & _
                          " " & lstDeduction.ListItems.Item(i).SubItems(1) & ", " & _
                          " '" & FormatDateTime(Arr(1), vbShortDate) & "', " & _
                          " 'O', " & iCompKey & ", " & CDbl(lstDeduction.ListItems.Item(i).SubItems(3)) & ", " & _
                          " " & CDbl(lstDeduction.ListItems.Item(i).SubItems(3)) & ", " & _
                          " " & lstDeduction.ListItems.Item(i).SubItems(4) & ", " & _
                          " '" & "Payroll Period " & Trim(txtPayrollPeriod.Text) & "')"
    Next i
            
End If

'GL/SL
dSalaryWages = RETURNTEXTVALUE(txtAmountEarned) + _
               RETURNTEXTVALUE(txtSHAmount) + _
               RETURNTEXTVALUE(txtLHAmount) + _
               RETURNTEXTVALUE(txtSLAmount) + _
               RETURNTEXTVALUE(txtAdjustment)
                  
dOvertime = RETURNTEXTVALUE(txtRegOTAmount) + _
            RETURNTEXTVALUE(txtRDOTAmount) + _
            RETURNTEXTVALUE(txtSHOTAmount) + _
            RETURNTEXTVALUE(txtLHOTAmount)

dColaAllowance = RETURNLABELVALUE(lblCola)

dAREmployee = RETURNTEXTVALUE(txtMortuary) + _
              RETURNTEXTVALUE(txtAROthers) + _
              RETURNTEXTVALUE(txtAdvances) + _
              RETURNTEXTVALUE(txtShortages) + _
              RETURNTEXTVALUE(txtUniform) + _
              RETURNTEXTVALUE(txtOthers)

Arr = Split(Trim(txtName.Text), " - ", -1, 1)
Arr1 = Split(Trim(txtPayrollPeriod.Text), " - ", -1, 1)


'ConnOmega.Execute "DELETE FROM tbl_GL_Transaction " & _
                  " WHERE (PayeeType = 3) " & _
                  " AND (PayeeKey = " & RETURNTEXTVALUE(txtEmpPK) & ") " & _
                  " AND (DocDate = '" & FormatDateTime(Arr1(1), vbShortDate) & "') " & _
                  " AND (DocNumber = '" & "PYRL" & Format(Arr1(1), "mmddyy") & "')"
                  
'ConnOmega.Execute "DELETE FROM tbl_Personnel_SL " & _
                  " WHERE (iType = 3) " & _
                  " AND (EmployeeKey = " & RETURNTEXTVALUE(txtEmpPK) & ") " & _
                  " AND (DocDate = '" & FormatDateTime(Arr1(1), vbShortDate) & "') " & _
                  " AND (DocNumber = '" & "PYRL" & Format(Arr1(1), "mmddyy") & "')"

'COMPUTE_GL_SL RETURNTEXTVALUE(txtEmpPK), 3, RETURNTEXTVALUE(txtDeptKey), Arr(0), Arr(1), Arr1(1), _
              "PYRL" & Format(Arr1(1), "mmddyy"), dSalaryWages, dOvertime, dColaAllowance, _
              RETURNTEXTVALUE(txtSSS), RETURNTEXTVALUE(txtPHIC), RETURNTEXTVALUE(txtPagIbig), _
              RETURNTEXTVALUE(txtWithHeld), RETURNTEXTVALUE(txtSSSLoan), RETURNTEXTVALUE(txtPagIbigLoan), _
              dAREmployee, "PYRL PER " & Trim(txtPayrollPeriod.Text), RETURNTEXTVALUE(txtEC) + _
              RETURNTEXTVALUE(txtSSSEmployer) + RETURNTEXTVALUE(txtPHICEmployer) + RETURNTEXTVALUE(txtPagIbigEmployer), _
              RETURNTEXTVALUE(txtEC) + RETURNTEXTVALUE(txtSSSEmployer), RETURNTEXTVALUE(txtPHICEmployer), _
              RETURNTEXTVALUE(txtPagIbigEmployer)
              
              
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER iCompKey, "is_LOAD"

Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Function
End Function

Private Sub INSERT_UPDATE_GL_SL(iPayeeT, iPayeeTypeT, dDocDateT, sDocNumT, _
tmpGLAccntT, sIDT, sNameT, sParticularsT, dDebitT, dCreditT, iWithSLT)
iGLAmount = CDbl(dDebitT) - CDbl(dCreditT)
If CDbl(iGLAmount) <> 0 Then
    s = "SELECT tbl_GL_Transaction.* " & _
        " FROM tbl_GL_Transaction " & _
        " WHERE (PayeeType = 3) " & _
        " AND (PayeeKey = " & iPayeeT & ") " & _
        " AND (DocDate = '" & FormatDateTime(dDocDateT, vbShortDate) & "') " & _
        " AND (DocNumber = '" & FORMATSQL(Trim(CStr(sDocNumT))) & "') " & _
        " AND (GLCode = '" & tmpGLAccntT & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount = 0 Then
        ConnOmega.Execute "INSERT INTO tbl_GL_Transaction " & _
                          " (GLCode, BookType, DocDate, DocNumber, PayeeKey, PayeeType, " & _
                          " SupplierCode, SupplierName, Particulars, Debit, Credit) " & _
                          " VALUES ('" & tmpGLAccntT & "', 5, " & _
                          " '" & FormatDateTime(dDocDateT, vbShortDate) & "', " & _
                          " '" & FORMATSQL(Trim(CStr(sDocNumT))) & "', " & iPayee & ", " & _
                          " " & iPayeeTypeT & ", '" & sIDT & "', '" & FORMATSQL(Trim(CStr(sNameT))) & "', " & _
                          " '" & FORMATSQL(Trim(CStr(sParticularsT))) & "', " & CDbl(dDebitT) & ", " & _
                          " " & CDbl(dCreditT) & ")"
    Else
        ConnOmega.Execute "UPDATE tbl_GL_Transaction " & _
                          " SET Debit = " & CDbl(dDebitT) & ", " & _
                          " Credit = " & CDbl(dCreditT) & " " & _
                          " WHERE (PK = " & rs!PK & ")"
    End If
    rs.Close
    If CDbl(iWithSLT) = 1 Then
        s = "SELECT tbl_Personnel_SL.* " & _
            " FROM tbl_Personnel_SL " & _
            " WHERE (iType = 3) " & _
            " AND (EmployeeKey = " & iPayeeT & ") " & _
            " AND (DocDate = '" & FormatDateTime(dDocDateT, vbShortDate) & "') " & _
            " AND (DocNumber = '" & FORMATSQL(Trim(CStr(sDocNumT))) & "') " & _
            " AND (GLCode = '" & tmpGLAccntT & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_SL " & _
                              " (EmployeeKey, GLCode, DocNumber, DocDate, " & _
                              " Description, Reference, iType, Debit, Credit) " & _
                              " VALUES (" & iPayeeT & ", '" & tmpGLAccntT & "', " & _
                              " '" & FORMATSQL(Trim(CStr(sDocNumT))) & "', " & _
                              " '" & FormatDateTime(dDocDateT, vbShortDate) & "', " & _
                              " '','" & FORMATSQL(Trim(CStr(sParticularsT))) & "', " & _
                              " 3, " & CDbl(dDebitT) & ", " & CDbl(dCreditT) & ")"
        Else
            ConnOmega.Execute "UPDATE tbl_Personnel_SL " & _
                              " SET Debit = " & CDbl(dDebitT) & ", " & _
                              " Credit = " & CDbl(dCreditT) & " " & _
                              " WHERE (PK = " & rs!PK & ")"
        End If
        rs.Close
    End If
End If
End Sub

Private Sub COMPUTE_GL_SL(iPayeeID, iPayeeType, iDept, sID, sName, dDocDate, sDocNum, _
dEarnings, dOTPay, dCOLAAllow, dSSS, dPHIC, dPagIbig, dITW, dSSSLoan, dPagIbigLoan, _
dDeductions, sParticulars, dSSSPHICPGBGEmployer, dSSSECEmp, dPHICEmp, dPGBGEmp)

iPayee = 0
t = "SELECT ProfileKey " & _
    " FROM tbl_Personnel_IDNumber " & _
    " WHERE (PK = " & iPayeeID & ")"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then
    iPayee = rt!ProfileKey
End If
rt.Close

ConnOmega.Execute "DELETE FROM tbl_GL_Transaction " & _
                  " WHERE (PayeeType = 3) " & _
                  " AND (PayeeKey = " & iPayee & ") " & _
                  " AND (DocDate = '" & FormatDateTime(Arr1(1), vbShortDate) & "') " & _
                  " AND (DocNumber = '" & "PYRL" & Format(Arr1(1), "mmddyy") & "')"
                  
ConnOmega.Execute "DELETE FROM tbl_Personnel_SL " & _
                  " WHERE (iType = 3) " & _
                  " AND (EmployeeKey = " & iPayee & ") " & _
                  " AND (DocDate = '" & FormatDateTime(Arr1(1), vbShortDate) & "') " & _
                  " AND (DocNumber = '" & "PYRL" & Format(Arr1(1), "mmddyy") & "')"

For i = 1 To 12
    iWithSL = 0: tmpAmt = 0: dDebit = 0: dCredit = 0
    Select Case i
        Case 1
            t = "SELECT tbl_GL_Accounts.* " & _
                " FROM tbl_GL_Accounts " & _
                " WHERE (Dept = " & iDept & ") " & _
                " AND (Link2PayrollSeries = " & i & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                tmpGLAccnt = rt!AccountCode
                iWithSL = rt!withSL
                dDebit = dEarnings
                dCredit = 0
                INSERT_UPDATE_GL_SL iPayee, iPayeeType, dDocDate, sDocNum, tmpGLAccnt, _
                                    sID, sName, sParticulars, dDebit, dCredit, iWithSL
            End If
            rt.Close
        Case 2
            t = "SELECT tbl_GL_Accounts.* " & _
                " FROM tbl_GL_Accounts " & _
                " WHERE (Dept = " & iDept & ") " & _
                " AND (Link2PayrollSeries = " & i & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                tmpGLAccnt = rt!AccountCode
                iWithSL = rt!withSL
                dDebit = dOTPay
                dCredit = 0
                INSERT_UPDATE_GL_SL iPayee, iPayeeType, dDocDate, sDocNum, tmpGLAccnt, _
                                    sID, sName, sParticulars, dDebit, dCredit, iWithSL
            End If
            rt.Close
        Case 3
            t = "SELECT tbl_GL_Accounts.* " & _
                " FROM tbl_GL_Accounts " & _
                " WHERE (Dept = " & iDept & ") " & _
                " AND (Link2PayrollSeries = " & i & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                tmpGLAccnt = rt!AccountCode
                iWithSL = rt!withSL
                dDebit = dCOLAAllow
                dCredit = 0
                INSERT_UPDATE_GL_SL iPayee, iPayeeType, dDocDate, sDocNum, tmpGLAccnt, _
                                    sID, sName, sParticulars, dDebit, dCredit, iWithSL
            End If
            rt.Close
        Case 5
            t = "SELECT tbl_GL_Accounts.* " & _
                " FROM tbl_GL_Accounts " & _
                " WHERE (Link2PayrollSeries = " & i & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                tmpGLAccnt = rt!AccountCode
                iWithSL = rt!withSL
                dDebit = 0
                dCredit = CDbl(dSSS)
                INSERT_UPDATE_GL_SL iPayee, iPayeeType, dDocDate, sDocNum, tmpGLAccnt, _
                                    sID, sName, sParticulars, dDebit, dCredit, iWithSL
            End If
            rt.Close
        Case 6
            t = "SELECT tbl_GL_Accounts.* " & _
                " FROM tbl_GL_Accounts " & _
                " WHERE (Link2PayrollSeries = " & i & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                tmpGLAccnt = rt!AccountCode
                iWithSL = rt!withSL
                dDebit = 0
                dCredit = CDbl(dPHIC)
                INSERT_UPDATE_GL_SL iPayee, iPayeeType, dDocDate, sDocNum, tmpGLAccnt, _
                                    sID, sName, sParticulars, dDebit, dCredit, iWithSL
            End If
            rt.Close
        Case 7
            t = "SELECT tbl_GL_Accounts.* " & _
                " FROM tbl_GL_Accounts " & _
                " WHERE (Link2PayrollSeries = " & i & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                tmpGLAccnt = rt!AccountCode
                iWithSL = rt!withSL
                dDebit = 0
                dCredit = CDbl(dPagIbig)
                INSERT_UPDATE_GL_SL iPayee, iPayeeType, dDocDate, sDocNum, tmpGLAccnt, _
                                    sID, sName, sParticulars, dDebit, dCredit, iWithSL
            End If
            rt.Close
        Case 8
            t = "SELECT tbl_GL_Accounts.* " & _
                " FROM tbl_GL_Accounts " & _
                " WHERE (Link2PayrollSeries = " & i & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                tmpGLAccnt = rt!AccountCode
                iWithSL = rt!withSL
                dDebit = 0
                dCredit = CDbl(dITW)
                INSERT_UPDATE_GL_SL iPayee, iPayeeType, dDocDate, sDocNum, tmpGLAccnt, _
                                    sID, sName, sParticulars, dDebit, dCredit, iWithSL
            End If
            rt.Close
        Case 9
            t = "SELECT tbl_GL_Accounts.* " & _
                " FROM tbl_GL_Accounts " & _
                " WHERE (Link2PayrollSeries = " & i & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                tmpGLAccnt = rt!AccountCode
                iWithSL = rt!withSL
                dDebit = 0
                dCredit = CDbl(dSSSLoan)
                INSERT_UPDATE_GL_SL iPayee, iPayeeType, dDocDate, sDocNum, tmpGLAccnt, _
                                    sID, sName, sParticulars, dDebit, dCredit, iWithSL
            End If
            rt.Close
        Case 10
            t = "SELECT tbl_GL_Accounts.* " & _
                " FROM tbl_GL_Accounts " & _
                " WHERE (Link2PayrollSeries = " & i & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                tmpGLAccnt = rt!AccountCode
                iWithSL = rt!withSL
                dDebit = 0
                dCredit = CDbl(dPagIbigLoan)
                INSERT_UPDATE_GL_SL iPayee, iPayeeType, dDocDate, sDocNum, tmpGLAccnt, _
                                    sID, sName, sParticulars, dDebit, dCredit, iWithSL
            End If
            rt.Close
        Case 11
            t = "SELECT tbl_GL_Accounts.* " & _
                " FROM tbl_GL_Accounts " & _
                " WHERE (Link2PayrollSeries = " & i & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                tmpGLAccnt = rt!AccountCode
                iWithSL = rt!withSL
                dDebit = 0
                dCredit = CDbl(dDeductions)
                INSERT_UPDATE_GL_SL iPayee, iPayeeType, dDocDate, sDocNum, tmpGLAccnt, _
                                    sID, sName, sParticulars, dDebit, dCredit, iWithSL
            End If
            rt.Close
        Case 12
            t = "SELECT tbl_GL_Accounts.* " & _
                " FROM tbl_GL_Accounts " & _
                " WHERE (Link2PayrollSeries = " & i & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                tmpGLAccnt = rt!AccountCode
                iWithSL = rt!withSL
                tmpAmt = (CDbl(dEarnings) + CDbl(dOTPay) + CDbl(dCOLAAllow)) - _
                         CDbl(dSSS) - CDbl(dPHIC) - CDbl(dPagIbig) - CDbl(dITW) - CDbl(dSSSLoan) - _
                         CDbl(dPagIbigLoan) - CDbl(dDeductions)
                dDebit = 0
                dCredit = CDbl(Format(tmpAmt, "#,##0.00"))
                INSERT_UPDATE_GL_SL iPayee, iPayeeType, dDocDate, sDocNum, tmpGLAccnt, _
                                    sID, sName, sParticulars, dDebit, dCredit, iWithSL
            End If
            rt.Close
    End Select
Next i

For i = 4 To 7
    iWithSL = 0: tmpAmt = 0: dDebit = 0: dCredit = 0
    Select Case i
        Case 4
            t = "SELECT tbl_GL_Accounts.* " & _
                " FROM tbl_GL_Accounts " & _
                " WHERE (Dept = " & iDept & ") " & _
                " AND (Link2PayrollSeries = " & i & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                tmpGLAccnt = rt!AccountCode
                iWithSL = rt!withSL
                dDebit = dSSSPHICPGBGEmployer
                dCredit = 0
                INSERT_UPDATE_GL_SL iPayee, iPayeeType, dDocDate, sDocNum, tmpGLAccnt, _
                                    sID, sName, sParticulars, dDebit, dCredit, iWithSL
            End If
            rt.Close
        Case 5
            t = "SELECT tbl_GL_Accounts.* " & _
                " FROM tbl_GL_Accounts " & _
                " WHERE (Link2PayrollSeries = " & i & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                tmpGLAccnt = rt!AccountCode
                iWithSL = rt!withSL
                dDebit = 0
                dCredit = CDbl(dSSSECEmp)
                INSERT_UPDATE_GL_SL iPayee, iPayeeType, dDocDate, sDocNum, tmpGLAccnt, _
                                    sID, sName, sParticulars, dDebit, dCredit, iWithSL
            End If
            rt.Close
        Case 6
            t = "SELECT tbl_GL_Accounts.* " & _
                " FROM tbl_GL_Accounts " & _
                " WHERE (Link2PayrollSeries = " & i & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                tmpGLAccnt = rt!AccountCode
                iWithSL = rt!withSL
                dDebit = 0
                dCredit = CDbl(dPHICEmp)
                INSERT_UPDATE_GL_SL iPayee, iPayeeType, dDocDate, sDocNum, tmpGLAccnt, _
                                    sID, sName, sParticulars, dDebit, dCredit, iWithSL
            End If
            rt.Close
        Case 7
            t = "SELECT tbl_GL_Accounts.* " & _
                " FROM tbl_GL_Accounts " & _
                " WHERE (Link2PayrollSeries = " & i & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                tmpGLAccnt = rt!AccountCode
                iWithSL = rt!withSL
                dDebit = 0
                dCredit = CDbl(dPGBGEmp)
                INSERT_UPDATE_GL_SL iPayee, iPayeeType, dDocDate, sDocNum, tmpGLAccnt, _
                                    sID, sName, sParticulars, dDebit, dCredit, iWithSL
            End If
            rt.Close
    End Select
    
'    s = "SELECT tbl_GL_Transaction.* " & _
'        " FROM tbl_GL_Transaction " & _
'        " WHERE (PayeeType = 3) " & _
'        " AND (PayeeKey = " & iPayee & ") " & _
'        " AND (DocDate = '" & FormatDateTime(dDocDate, vbShortDate) & "') " & _
'        " AND (DocNumber = '" & FORMATSQL(Trim(CStr(sDocNum))) & "') " & _
'        " AND (GLCode = '" & tmpGLAccnt & "')"
'    If rs.State = adStateOpen Then rs.Close
'    rs.Open s, ConnOmega
'    If rs.RecordCount = 0 Then
'        ConnOmega.Execute "INSERT INTO tbl_GL_Transaction " & _
'                          " (GLCode, BookType, DocDate, DocNumber, PayeeKey, PayeeType, " & _
'                          " SupplierCode, SupplierName, Particulars, Debit, Credit) " & _
'                          " VALUES ('" & tmpGLAccnt & "', 5, " & _
'                          " '" & FormatDateTime(dDocDate, vbShortDate) & "', " & _
'                          " '" & FORMATSQL(Trim(CStr(sDocNum))) & "', " & iPayee & ", " & _
'                          " " & iPayeeType & ", '" & sID & "', '" & FORMATSQL(Trim(CStr(sName))) & "', " & _
'                          " '" & FORMATSQL(Trim(CStr(sParticulars))) & "', " & CDbl(dDebit) & ", " & _
'                          " " & CDbl(dCredit) & ")"
'    Else
'        ConnOmega.Execute "UPDATE tbl_GL_Transaction " & _
'                          " SET Debit = " & CDbl(dDebit) & ", " & _
'                          " Credit = " & CDbl(dCredit) & " " & _
'                          " WHERE (PK = " & rs!PK & ")"
'    End If
'    rs.Close
'
'    If CDbl(iWithSL) = 1 Then
'        s = "SELECT tbl_Personnel_SL.* " & _
'            " FROM tbl_Personnel_SL " & _
'            " WHERE (iType = 3) " & _
'            " AND (EmployeeKey = " & iPayee & ") " & _
'            " AND (DocDate = '" & FormatDateTime(dDocDate, vbShortDate) & "') " & _
'            " AND (DocNumber = '" & FORMATSQL(Trim(CStr(sDocNum))) & "') " & _
'            " AND (GLCode = '" & tmpGLAccnt & "')"
'        If rs.State = adStateOpen Then rs.Close
'        rs.Open s, ConnOmega
'        If rs.RecordCount = 0 Then
'            ConnOmega.Execute "INSERT INTO tbl_Personnel_SL " & _
'                              " (EmployeeKey, GLCode, DocNumber, DocDate, " & _
'                              " Description, Reference, iType, Debit, Credit) " & _
'                              " VALUES (" & iPayee & ", '" & tmpGLAccnt & "', " & _
'                              " '" & FORMATSQL(Trim(CStr(sDocNum))) & "', " & _
'                              " '" & FormatDateTime(dDocDate, vbShortDate) & "', " & _
'                              " '','" & FORMATSQL(Trim(CStr(sParticulars))) & "', " & _
'                              " 3, " & CDbl(dDebit) & ", " & CDbl(dCredit) & ")"
'        Else
'            ConnOmega.Execute "UPDATE tbl_Personnel_SL " & _
'                              " SET Debit = " & CDbl(dDebit) & ", " & _
'                              " Credit = " & CDbl(dCredit) & " " & _
'                              " WHERE (PK = " & rs!PK & ")"
'        End If
'        rs.Close
'    End If
Next i
End Sub

Private Function PRESS_F6()
If picPrint.Visible = True Then Exit Function
If picAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
txtSearch.Text = ""
picSearch.ZOrder 0
txtSearch.Visible = True
picSearch.Visible = True
txtSearch.SetFocus
End Function

Private Function PRESS_F9()
'If picPrint.Visible = True Then Exit Function
'If picAdd.Visible = True Then Exit Function
'If picSearch.Visible = True Then Exit Function
'If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
'picToolbar.Enabled = False
'picBody.Enabled = False
'picPrint.ZOrder 0
'picPrint.Visible = True
'cmbGroup_Click
'lstReportType.SetFocus
End Function

Private Function PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSearch.Visible = True Then cmdCancelSearch_Click: Exit Function
    If picAdd.Visible = True Then cmdCancelAdd_Click: Exit Function
    If picPrint.Visible = True Then cmdCancelPrint_Click: Exit Function
    Unload Me
Else
    Label2.Visible = False
    Label31.Visible = False
    BROWSER GetSetting(App.EXEName, "CompensationC", "CompC", ""), "is_LOAD"
    TRANSACTIONTYPE = is_REFRESH
    LOCKTEXT True
    TOOLBARFUNC 1
End If
End Function

Private Function LOAD_GROUP_BY(intIndex)
With cmbGroup
    .Clear
    If intIndex = 0 Then
        .AddItem "DEPARTMENT"
        .AddItem "STATUS"
        .AddItem "POSITION"
        .AddItem "EMPLOYEE"
    Else
        .AddItem cmbDivision.Text
    End If
    .ListIndex = 0
End With
End Function

Private Function CALCULATE_EARNING()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    lblTotalEarnings.Caption = Format(RETURNTEXTVALUE(txtAmountEarned) + RETURNTEXTVALUE(txtSHAmount) + _
                               RETURNTEXTVALUE(txtLHAmount) + RETURNTEXTVALUE(txtSLAmount) + _
                               RETURNTEXTVALUE(txtAdjustment) + RETURNTEXTVALUE(txtRegOTAmount) + _
                               RETURNTEXTVALUE(txtRDOTAmount) + RETURNTEXTVALUE(txtSHOTAmount) + _
                               RETURNTEXTVALUE(txtLHOTAmount), "##,##0.00")
End If
End Function

Private Function CALCULATE_DEDUCTIONS()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    lblTotalDeductions.Caption = Format(RETURNTEXTVALUE(txtSSSLoan) + RETURNTEXTVALUE(txtPagIbigLoan) + _
                                 RETURNTEXTVALUE(txtMortuary) + RETURNTEXTVALUE(txtAROthers) + _
                                 RETURNTEXTVALUE(txtAdvances) + RETURNTEXTVALUE(txtShortages) + _
                                 RETURNTEXTVALUE(txtUniform) + RETURNTEXTVALUE(txtOthers) + _
                                 RETURNTEXTVALUE(txtSSS) + RETURNTEXTVALUE(txtPHIC) + _
                                 RETURNTEXTVALUE(txtPagIbig) + RETURNTEXTVALUE(txtWithHeld), "##,##0.00")
End If
End Function

Private Function CALCULATE_LOANS()
If (chkLoan.Value = 1 And TRANSACTIONTYPE = is_ADDING) Or _
(chkLoan.Value = 1 And TRANSACTIONTYPE = is_EDITTING) Then
    If CDbl(lblTotalEarnings.Caption) <> 0 Then
        Array1 = Split(txtPayrollPeriod.Text, " - ", -1, 1)
        dtmPayrollDate = CDate(Array1(0))
        If IS_HAVE_SSS_LOAN(txtEmpPK.Text, dtmPayrollDate) = True Then
            If CHECK_LOAN_CUTOFF("SSS_Loan", Day(CDate(dtmPayrollDate))) = True Then
                lngSSSLoanNo = GET_SSS_LOAN_NO(txtEmpPK.Text, dtmPayrollDate)
                strSSSLoanInfo = Split(GET_LOAN_INFO(lngSSSLoanNo), ";", -1, 1)
                dblSSSTotalLoan = CDbl(strSSSLoanInfo(3))
                dblSSSMonthly = CDbl(strSSSLoanInfo(0))
                dblSSSTotalPaid = GET_TOTAL_PAID_SSS(lngSSSLoanNo, txtPeriod.Text)
                dblSSSDif = CDbl(dblSSSTotalLoan) - CDbl(dblSSSTotalPaid)
                If CDbl(dblSSSDif) >= CDbl(dblSSSMonthly) Then
                    txtSSSLoan.Text = Format(CDbl(dblSSSMonthly), "##,##0.00")
                Else
                    txtSSSLoan.Text = Format(CDbl(dblSSSDif), "##,##0.00")
                End If
                txtSSSLoanNo.Text = lngSSSLoanNo
                txtSSSTotalPaid.Text = CDbl(dblSSSTotalPaid) + CDbl(txtSSSLoan.Text)
                txtSSSLoanBalance.Text = CDbl(dblSSSTotalLoan) - (CDbl(dblSSSTotalPaid) + CDbl(txtSSSLoan.Text))
            Else
                txtSSSLoanNo.Text = "0"
                txtSSSTotalPaid.Text = "0"
                txtSSSLoanBalance.Text = "0"
                txtSSSLoan.Text = "0.00"
            End If
        Else
            txtSSSLoanNo.Text = "0"
            txtSSSTotalPaid.Text = "0"
            txtSSSLoanBalance.Text = "0"
            txtSSSLoan.Text = "0.00"
        End If
        If IS_HAVE_PAGIBIG_LOAN(txtEmpPK.Text, dtmPayrollDate) = True Then
            If CHECK_LOAN_CUTOFF("PagIbig_Loan", Day(CDate(dtmPayrollDate))) = True Then
                lngPagIbigLoanNo = GET_PAGIBIG_LOAN_NO(txtEmpPK.Text, dtmPayrollDate)
                strPagIbigLoanInfo = Split(GET_LOAN_INFO(lngPagIbigLoanNo), ";", -1, 1)
                dblPagIbigTotalLoan = CDbl(strPagIbigLoanInfo(3))
                dblPagIbigMonthly = CDbl(strPagIbigLoanInfo(0))
                dblPagIbigTotalPaid = GET_TOTAL_PAID_PAGIBIG(lngPagIbigLoanNo, txtPeriod.Text)
                dblPagIbigDiff = CDbl(dblPagIbigTotalLoan) - CDbl(dblPagIbigTotalPaid)
                If dblPagIbigDiff >= dblPagIbigMonthly Then
                    txtPagIbigLoan.Text = Format(CDbl(dblPagIbigMonthly), "##,##0.00")
                Else
                    txtPagIbigLoan.Text = Format(CDbl(dblPagIbigDiff), "##,##0.00")
                End If
                txtPagIbigLoanNo.Text = lngPagIbigLoanNo
                txtPagIbigTotalPaid.Text = CDbl(dblPagIbigTotalPaid) + CDbl(txtPagIbigLoanNo.Text)
                txtPagIbigLoanBalance.Text = CDbl(dblPagIbigTotalLoan) - (CDbl(dblPagIbigTotalPaid) + CDbl(txtPagIbigLoan.Text))
            Else
                txtPagIbigLoanNo.Text = "0"
                txtPagIbigTotalPaid.Text = "0"
                txtPagIbigLoanBalance.Text = "0"
                txtPagIbigLoan.Text = "0.00"
            End If
        Else
            txtPagIbigLoanNo.Text = "0"
            txtPagIbigTotalPaid.Text = "0"
            txtPagIbigLoanBalance.Text = "0"
            txtPagIbigLoan.Text = "0.00"
        End If
    End If
Else
    txtSSSLoanNo.Text = "0"
    txtSSSLoanBalance.Text = "0"
    txtSSSTotalPaid.Text = "0"
    txtPagIbigLoanNo.Text = "0"
    txtPagIbigTotalPaid.Text = "0"
    txtPagIbigLoanBalance.Text = "0"
    txtSSSLoan.Text = "0.00"
    txtPagIbigLoan.Text = "0.00"
End If
End Function

Private Function CALCULATE_CONTRIBUTION()
Array1 = Split(txtPayrollPeriod.Text, " - ", -1, 1)
dtmPayrollDate = FormatDateTime(Array1(0), vbShortDate)
dtmPayrollDate1 = FormatDateTime(Array1(1), vbShortDate)
If (chkContribution.Value = 1 And TRANSACTIONTYPE = is_ADDING) Or _
(chkContribution.Value = 1 And TRANSACTIONTYPE = is_EDITTING) Then
    If CDbl(lblTotalEarnings.Caption) <> 0 Then
        dblPreviousGross = GET_PREVIOUS_GROSS(RETURNTEXTVALUE(txtPeriod), RETURNTEXTVALUE(txtDivision), RETURNTEXTVALUE(txtEmpPK))
        dblGrossForCont = CDbl(lblTotalEarnings.Caption) + CDbl(dblPreviousGross)
        If IS_HAVE_SSS(txtEmpPK.Text, dtmPayrollDate) = True Then
            If CHECK_CONT_CUTOFF("SSS", Day(dtmPayrollDate)) = True Then
                dblSSSEmployer = Format(GET_SSS_CONTRIBUTION_EMPLOYER(dblGrossForCont, FormatDateTime(dtmPayrollDate1, vbShortDate)), "##,##0.00")
                dblSSSEmployee = Format(GET_SSS_CONTRIBUTION_EMPLOYEE(dblGrossForCont, FormatDateTime(dtmPayrollDate1, vbShortDate)), "##,##0.00")
                dblSSSEC = Format(GET_SSS_CONTRIBUTION_EC(dblGrossForCont, FormatDateTime(dtmPayrollDate1, vbShortDate)), "##,##0.00")
            Else
                dblSSSEmployer = "0.00"
                dblSSSEmployee = "0.00"
                dblSSSEC = "0.00"
            End If
            txtSSSEmployer.Text = dblSSSEmployer
            txtSSS.Text = dblSSSEmployee
            txtEC.Text = dblSSSEC
        Else
            txtSSSEmployer.Text = "0.00"
            txtSSS.Text = "0.00"
            txtEC.Text = "0.00"
        End If
        If IS_HAVE_PHIC(txtEmpPK.Text, dtmPayrollDate) = True Then
            If CHECK_CONT_CUTOFF("PHIC", Day(dtmPayrollDate)) = True Then
                dblPHICEmployer = Format(GET_PHIC_CONTRIBUTION_EMPLOYER(dblGrossForCont, FormatDateTime(dtmPayrollDate1, vbShortDate)), "##,##0.00")
                dblPHICEmployee = Format(GET_PHIC_CONTRIBUTION_EMPLOYEE(dblGrossForCont, FormatDateTime(dtmPayrollDate1, vbShortDate)), "##,##0.00")
            Else
                dblPHICEmployer = "0.00"
                dblPHICEmployee = "0.00"
            End If
            txtPHICEmployer.Text = dblPHICEmployer
            txtPHIC.Text = dblPHICEmployee
        Else
            txtPHICEmployer.Text = "0.00"
            txtPHIC.Text = "0.00"
        End If
        If IS_HAVE_PagIbig(txtEmpPK.Text, dtmPayrollDate) = True Then
            If CHECK_CONT_CUTOFF("PagIbig", Day(dtmPayrollDate)) = True Then
                dblPagIbigEmployer = Format(GET_PAGIBIG_CONTRIBUTION_EMPLOYER(dblGrossForCont), "##,##0.00")
                dblPagIbigEmployee = Format(GET_PAGIBIG_CONTRIBUTION_EMPLOYEE(dblGrossForCont), "##,##0.00")
            Else
                dblPagIbigEmployer = "0.00"
                dblPagIbigEmployee = "0.00"
            End If
            txtPagIbigEmployer.Text = dblPagIbigEmployer
            txtPagIbig.Text = dblPagIbigEmployee
        Else
            txtPagIbigEmployer.Text = "0.00"
            txtPagIbig.Text = "0.00"
        End If
        If IS_HAVE_TIN(txtEmpPK.Text, dtmPayrollDate) = True Then
            If CHECK_CONT_CUTOFF("WithHeld", Day(dtmPayrollDate)) = True Then
                dblBasicForTax = (CDbl(lblTotalEarnings.Caption) + CDbl(dblPreviousGross)) - (CDbl(dblSSSEmployee) + CDbl(dblPHICEmployee) + CDbl(dblPagIbigEmployee))
                dblTaxExemp = COMPUTE_MONTHLY_TAX_EXEMP(GET_CURRENT_STATUS(txtEmpPK.Text, dtmPayrollDate), dblBasicForTax, dtmPayrollDate1) / 12
                dblPercent = COMPUTE_MONTHLY_TAX_PERCENT(GET_CURRENT_STATUS(txtEmpPK.Text, dtmPayrollDate), dblBasicForTax, dtmPayrollDate1)
                dblConstant = COMPUTE_MONTHLY_TAX_CONSTANT(GET_CURRENT_STATUS(txtEmpPK.Text, dtmPayrollDate), dblBasicForTax, dtmPayrollDate1)
                dblBracketAmount = COMPUTE_MONTHLY_TAX_BRACKET_AMOUNT(GET_CURRENT_STATUS(txtEmpPK.Text, dtmPayrollDate), dblBasicForTax, dtmPayrollDate1)
                var1 = CDbl(dblBasicForTax) - CDbl(dblBracketAmount)
                var2 = CDbl(var1) * (CDbl(dblPercent) / 100)
                var3 = CDbl(var2) + CDbl(dblConstant)
            Else
                txtWithHeld.Text = "0.00"
            End If
            txtWithHeld.Text = Format(IIf(var3 <= 0, 0, var3), "###,##0.00")
        Else
            txtWithHeld.Text = "0.00"
        End If
    End If
Else
    txtSSSEmployer.Text = "0.00"
    txtSSS.Text = "0.00"
    txtEC.Text = "0.00"
    txtPHICEmployer.Text = "0.00"
    txtPHIC.Text = "0.00"
    txtPagIbigEmployer.Text = "0.00"
    txtPagIbig.Text = "0.00"
    txtWithHeld.Text = "0.00"
End If
End Function

Private Function SAVE_PAYROLL(strEmpNo, intDiv, intDept, intStatus, intPositions, intPeriod, _
dblRatePerHour, intActionMemo, dblNoHours, dblSH_Hours, dblLH_Hours, dblSL_Hours, dblAdjustment, _
dblReg_OT, dblRD_OT, dblSH_OT, dblLH_OT, dblAmount_Earned, dblSH_Amount, dblLH_Amount, _
dblSL_Amount, dblReg_Amount, dblRD_Amount, dblSH_OT_Amount, dblLH_OT_Amount, _
dblTotalEarning, dblMotuary, dblAR_Others, dblAdvances, dblShortages, dblUniforms, _
dblOthers, intIsLoan, intSSSLoanNo, dblSSSLoan, dblSSSLoanBalance, intPagIbigLoanNo, _
dblPagIbigLoan, dblPagIbigLoanBalance, IsCont, dblSSS, dblSSS_Employer, dblSSS_EC, dblPHIC, _
dblPHIC_Employer, dblPagIbig, dblPagIbig_Employer, dblWithHeld, dblTotalDeduction, _
dblNetEarning, strLastModified)
ConnOmega.Execute "INSERT INTO tbl_Personnel_Compensation" & _
                " (EmpPK, Division, Dept, Status, Positions, Period, RatePerHour, ActionMemo, NoHours, " & _
                " SH_Hours, LH_Hours, SL_Hours, Adjustment, Reg_OT_Hours, RD_OT_Hours, " & _
                " SH_OT_Hours, LH_OT_Hours, Amount_Earned, SH_Amount, LH_Amount, " & _
                " SL_Amount, Reg_OT_Amount, RD_OT_Amount, SH_OT_Amount, LH_OT_Amount, " & _
                " TotalEarning, Mortuary, AR_Others, Advances, Shortages, Uniforms, " & _
                " Others, Is_Have_Loan, SSSLoan_No, SSSLoan, SSSBalance, PagIbigLoan_No, " & _
                " PagIbigLoan, PagIbigBalance, Is_Have_Cont, SSS, SSS_Employer, SSS_EC, PHIC, " & _
                " PHIC_Employer, PagIbig, PagIbig_Employer, WithHeld, TotalDeduction, " & _
                " NetEarning, LastModified) " & _
                " VALUES (" & strEmpNo & ", " & intDiv & ", " & intDept & ", " & intStatus & ", " & intPositions & ", " & intPeriod & "," & _
                " " & CDbl(dblRatePerHour) & ", " & CLng(intActionMemo) & "," & CDbl(dblNoHours) & ", " & CDbl(dblSH_Hours) & ", " & _
                " " & CDbl(dblLH_Hours) & ", " & CDbl(dblSL_Hours) & ", " & CDbl(dblAdjustment) & ", " & _
                " " & CDbl(dblReg_OT) & ", " & CDbl(dblRD_OT) & ", " & CDbl(dblSH_OT) & ", " & CDbl(dblLH_OT) & ", " & _
                " " & CDbl(dblAmount_Earned) & ", " & CDbl(dblSH_Amount) & ", " & CDbl(dblLH_Amount) & ", " & _
                " " & CDbl(dblSL_Amount) & ", " & CDbl(dblReg_Amount) & ", " & CDbl(dblRD_Amount) & ", " & CDbl(dblSH_OT_Amount) & ", " & _
                " " & CDbl(dblLH_OT_Amount) & ", " & CDbl(dblTotalEarning) & ", " & CDbl(dblMotuary) & ", " & CDbl(dblAR_Others) & ", " & _
                " " & CDbl(dblAdvances) & ", " & CDbl(dblShortages) & ", " & CDbl(dblUniforms) & ", " & CDbl(dblOthers) & ", " & _
                " " & intIsLoan & ", " & CLng(intSSSLoanNo) & ", " & CDbl(dblSSSLoan) & ", " & CDbl(dblSSSLoanBalance) & ", " & _
                " " & CLng(intPagIbigLoanNo) & ", " & CDbl(dblPagIbigLoan) & ", " & CDbl(dblPagIbigLoanBalance) & ", " & IsCont & ", " & _
                " " & CDbl(dblSSS) & ", " & CDbl(dblSSS_Employer) & ", " & CDbl(dblSSS_EC) & ", " & CDbl(dblPHIC) & ", " & CDbl(dblPHIC_Employer) & ", " & _
                " " & CDbl(dblPagIbig) & ", " & CDbl(dblPagIbig_Employer) & ", " & CDbl(dblWithHeld) & ", " & CDbl(dblTotalDeduction) & ", " & _
                " " & CDbl(dblNetEarning) & ", '" & strLastModified & "')"
End Function

Private Function UPDATE_PAYROLL(intPK, dblNoHours, dblSH_Hours, _
dblLH_Hours, dblSL_Hours, dblAdjustment, dblReg_OT, dblRD_OT, dblSH_OT, dblLH_OT, _
dblAmount_Earned, dblSH_Amount, dblLH_Amount, dblSL_Amount, dblReg_Amount, dblRD_Amount, _
dblSH_OT_Amount, dblLH_OT_Amount, dblTotalEarning, dblMotuary, dblAR_Others, dblAdvances, _
dblShortages, dblUniforms, dblOthers, intIsLoan, intSSSLoanNo, dblSSSLoan, dblSSSLoanBalance, _
intPagIbigLoanNo, dblPagIbigLoan, dblPagIbigLoanBalance, IsCont, dblSSS, dblSSS_Employer, _
dblSSS_EC, dblPHIC, dblPHIC_Employer, dblPagIbig, dblPagIbig_Employer, dblWithHeld, dblTotalDeduction, _
dblNetEarning, strLastModified)
ConnOmega.Execute "UPDATE tbl_Personnel_Compensation" & _
                " SET NoHours = " & CDbl(dblNoHours) & ", SH_Hours = " & CDbl(dblSH_Hours) & ", " & _
                " LH_Hours = " & CDbl(dblLH_Hours) & ", SL_Hours = " & CDbl(dblSL_Hours) & ", " & _
                " Adjustment = " & CDbl(dblAdjustment) & ", Reg_OT_Hours = " & CDbl(dblReg_OT) & ", " & _
                " RD_OT_Hours = " & CDbl(dblRD_OT) & ", SH_OT_Hours = " & CDbl(dblSH_OT) & ", " & _
                " LH_OT_Hours = " & CDbl(dblLH_OT) & ", Amount_Earned = " & CDbl(dblAmount_Earned) & ", " & _
                " SH_Amount = " & CDbl(dblSH_Amount) & ", LH_Amount = " & CDbl(dblLH_Amount) & ", " & _
                " SL_Amount = " & CDbl(dblSL_Amount) & ", Reg_OT_Amount = " & CDbl(dblReg_Amount) & ", " & _
                " RD_OT_Amount = " & CDbl(dblRD_Amount) & ", SH_OT_Amount = " & CDbl(dblSH_OT_Amount) & ", " & _
                " LH_OT_Amount = " & CDbl(dblLH_OT_Amount) & ", TotalEarning = " & CDbl(dblTotalEarning) & ", " & _
                " Mortuary = " & CDbl(dblMotuary) & ", AR_Others =  " & CDbl(dblAR_Others) & ", " & _
                " Advances = " & CDbl(dblAdvances) & ", Shortages = " & CDbl(dblShortages) & ", " & _
                " Uniforms = " & CDbl(dblUniforms) & ", Others = " & CDbl(dblOthers) & ", " & _
                " Is_Have_Loan = " & intIsLoan & ", SSSLoan_No = " & CLng(intSSSLoanNo) & ", " & _
                " SSSLoan = " & CDbl(dblSSSLoan) & ", SSSBalance = " & CDbl(dblSSSLoanBalance) & ", " & _
                " PagIbigLoan_No = " & CLng(intPagIbigLoanNo) & ", PagIbigLoan = " & CDbl(dblPagIbigLoan) & ", " & _
                " PagIbigBalance = " & CDbl(dblPagIbigLoanBalance) & ", Is_Have_Cont = " & IsCont & ", " & _
                " SSS = " & CDbl(dblSSS) & ", SSS_Employer = " & CDbl(dblSSS_Employer) & ", " & _
                " SSS_EC = " & CDbl(dblSSS_EC) & ", PHIC = " & CDbl(dblPHIC) & ", PHIC_Employer = " & CDbl(dblPHIC_Employer) & ", " & _
                " PagIbig = " & CDbl(dblPagIbig) & ", PagIbig_Employer = " & CDbl(dblPagIbig_Employer) & ", " & _
                " WithHeld = " & CDbl(dblWithHeld) & ", TotalDeduction = " & CDbl(dblTotalDeduction) & ", " & _
                " NetEarning = " & CDbl(dblNetEarning) & ", LastModified = '" & strLastModified & "'" & _
                " WHERE (PK = " & intPK & ")"
End Function

Private Function UPDATE_LOAN(intPK, dblTotalPaid, dblBalance)
ConnOmega.Execute "UPDATE tbl_Personnel_Loans" & _
                  " SET TotalPaid = " & CDbl(dblTotalPaid) & ", " & _
                  " Balance = " & CDbl(dblBalance) & "" & _
                  " WHERE (PK = " & intPK & ")"
End Function

Private Function CHECK_DUPLICATE_ENTRY(strEmpNo, intPeriod) As Boolean
s = "SELECT EmpPK, Period" & _
    " From tbl_Personnel_Compensation " & _
    " WHERE (EmpPK = " & strEmpNo & ") " & _
    " AND (Period = " & intPeriod & ")"
If rb.State = adStateOpen Then rb.Close
rb.Open s, ConnOmega
If Not rb.EOF Then
    CHECK_DUPLICATE_ENTRY = True
End If
rb.Close
End Function

Public Function CLEARTEXT()
Label2.Visible = False
Label31.Visible = False
txtIsLoan.Text = "0"
txtIsCont.Text = "0"
txtID.Text = ""
txtName.Text = ""
txtDept.Text = ""
txtPost.Text = ""
txtPayrollPeriod.Text = ""
txtEmpPK.Text = "0"
txtDivision.Text = "0"
txtDivName.Text = ""
txtDeptKey.Text = "0"
txtPostKey.Text = "0"
txtStatusKey.Text = "0"
txtRatePerHour.Text = "0"
txtPeriod.Text = "0"
txtActionPK.Text = "0"
txtColaPerHour.Text = "0"
txtAllowPerHour.Text = "0"
txtAllow.Text = "0"
txtNoHours.Text = "0.00"
txtColaHrs.Text = "0.00"
txtSH.Text = "0.00"
txtLH.Text = "0.00"
txtSL.Text = "0.00"
txtAdjustment.Text = "0.00"
txtAmountEarned.Text = "0.00"
txtSHAmount.Text = "0.00"
txtLHAmount.Text = "0.00"
txtSLAmount.Text = "0.00"
txtRegOT.Text = "0.00"
txtRDOT.Text = "0.00"
txtSHOT.Text = "0.00"
txtLHOT.Text = "0.00"
txtRegOTAmount.Text = "0.00"
txtRDOTAmount.Text = "0.00"
txtSHOTAmount.Text = "0.00"
txtLHOTAmount.Text = "0.00"
txtSSSLoan.Text = "0.00"
txtPagIbigLoan.Text = "0.00"
txtSSSLoanNo.Text = "0"
txtSSSLoanBalance.Text = "0"
txtSSSTotalPaid.Text = "0"
txtPagIbigTotalPaid.Text = "0"
txtPagIbigLoanNo.Text = "0"
txtPagIbigLoanBalance.Text = "0"
txtMortuary.Text = "0.00"
txtAROthers.Text = "0.00"
txtAdvances.Text = "0.00"
txtShortages.Text = "0.00"
txtUniform.Text = "0.00"
txtOthers.Text = "0.00"
txtSSS.Text = "0.00"
txtPHIC.Text = "0.00"
txtPagIbig.Text = "0.00"
txtWithHeld.Text = "0.00"
txtSSSEmployer.Text = "0"
txtEC.Text = "0"
txtPHICEmployer.Text = "0"
txtPagIbigEmployer.Text = "0"
chkLoan.Value = 0
chkContribution.Value = 0
StatusBar.Panels(1).Text = ""
StatusBar.Panels(2).Text = ""
StatusBar.Panels(3).Text = ""
StatusBar.Panels(4).Text = ""
lblTotalEarnings.Caption = "0.00"
lblTotalDeductions.Caption = "0.00"
lblCola.Caption = "0.00"
lblNetPay.Caption = "0.00"
imgPicture.Picture = LoadPicture("")
lstDeduction.ListItems.Clear
End Function

Public Function LOCKTEXT(bln As Boolean)
txtNoHours.Locked = bln
txtColaHrs.Locked = bln
txtSH.Locked = bln
txtLH.Locked = bln
txtSL.Locked = bln
txtAdjustment.Locked = bln
txtRegOT.Locked = bln
txtRDOT.Locked = bln
txtSHOT.Locked = bln
txtLHOT.Locked = bln

txtMortuary.Locked = bln
txtAROthers.Locked = bln
txtAdvances.Locked = bln
txtShortages.Locked = bln
txtUniform.Locked = bln
txtOthers.Locked = bln

'txtMortuary.Locked = True
'txtAROthers.Locked = True
'txtAdvances.Locked = True
'txtShortages.Locked = True
'txtUniform.Locked = True
'txtOthers.Locked = True

picLoan.Enabled = IIf(bln = False, True, False)
picCont.Enabled = IIf(bln = False, True, False)

End Function

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

Private Sub b8TitleBar1_CLoseClick()
cmdCancelSearch_Click
End Sub

Private Sub b8TitleBar2_CLoseClick()
cmdCancelAdd_Click
End Sub

Private Sub b8TitleBar3_CLoseClick()
cmdCancelPrint_Click
End Sub

Private Sub b8TitleBar4_CLoseClick()
cmdCancelTaxAlpha_Click
End Sub


Private Sub chkContribution_Click()
If CInt(txtIsCont.Text) = 1 Then
    If chkContribution.Value = 1 Then
        CALCULATE_CONTRIBUTION
        Timer2.Enabled = False
        Label31.Visible = False
    End If
Else
    chkContribution.Value = 0
End If
End Sub

Private Sub chkContribution_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtSSS.SetFocus
End If
End Sub

Private Sub chkLoan_Click()
If CInt(txtIsLoan.Text) = 1 Then
    If chkLoan.Value = 1 Then
        CALCULATE_LOANS
        Timer1.Enabled = False
        Label2.Visible = False
    End If
Else
    chkLoan.Value = 0
End If
End Sub

Private Sub chkLoan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtSSSLoan.SetFocus
End If
End Sub

Private Sub cmbDivision_Click()
If cmbDivision.ListIndex = -1 Then cmbPeriodPrint.Clear: txtTerms.Text = "": Exit Sub

cmbGroup_Click

LOAD_GROUP_BY lstReportType.ListIndex
Array1 = Split(FIND_PAYROLL_PERIOD(Date, cmbDivision.ListIndex + 1), ";", -1)
txtTerms.Text = CLng(Array1(3))
cmbPeriodPrint.Clear
If lstReportType.ListIndex = 19 Then
    s = "SELECT TOP 1 tbl_Personnel_Compensation_Period.DateFrom, " & _
        " tbl_Personnel_Compensation_Period.DateTo, " & _
        " tbl_Personnel_Compensation.Period " & _
        " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
        " tbl_Personnel_Compensation_Period ON tbl_Personnel_Compensation.Period = tbl_Personnel_Compensation_Period.PK " & _
        " Where (tbl_Personnel_Compensation.Division = " & cmbDivision.ListIndex + 1 & ") " & _
        " GROUP BY tbl_Personnel_Compensation_Period.DateFrom, tbl_Personnel_Compensation_Period.DateTo, tbl_Personnel_Compensation.Period " & _
        " ORDER BY tbl_Personnel_Compensation_Period.DateFrom DESC, tbl_Personnel_Compensation_Period.DateTo DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        dtmTo = rs!DateTo
    End If
    rs.Close
    
    s = "SELECT TOP 1 PK, DateFrom, DateTo " & _
        " From tbl_Personnel_Compensation_Period " & _
        " WHERE (Type = " & cmbDivision.ListIndex + 1 & ") " & _
        " AND (DateTo > '" & FormatDateTime(dtmTo, vbShortDate) & "') " & _
        " ORDER BY DateFrom, DateTo"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        cmbPeriodPrint.AddItem Format(rs!DateFrom, "mm/dd/yyyy") & " - " & Format(rs!DateTo, "mm/dd/yyyy")
        cmbPeriodPrint.ItemData(cmbPeriodPrint.NewIndex) = rs!PK
        rs.MoveNext
    Wend
    rs.Close
End If
s = "SELECT tbl_Personnel_Compensation_Period.DateFrom, " & _
    " tbl_Personnel_Compensation_Period.DateTo, " & _
    " tbl_Personnel_Compensation.Period " & _
    " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
    " tbl_Personnel_Compensation_Period ON tbl_Personnel_Compensation.Period = tbl_Personnel_Compensation_Period.PK " & _
    " Where (tbl_Personnel_Compensation.Division = " & cmbDivision.ListIndex + 1 & ") " & _
    " GROUP BY tbl_Personnel_Compensation_Period.DateFrom, tbl_Personnel_Compensation_Period.DateTo, tbl_Personnel_Compensation.Period " & _
    " ORDER BY tbl_Personnel_Compensation_Period.DateFrom DESC, tbl_Personnel_Compensation_Period.DateTo DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbPeriodPrint.AddItem Format(rs!DateFrom, "mm/dd/yyyy") & " - " & Format(rs!DateTo, "mm/dd/yyyy")
    cmbPeriodPrint.ItemData(cmbPeriodPrint.NewIndex) = rs!Period
    dtmTo = rs!DateTo
    rs.MoveNext
Wend
rs.Close

If cmbPeriodPrint.ListCount Then cmbPeriodPrint.ListIndex = 0
End Sub

Private Sub cmbDivision_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbPeriodPrint.SetFocus
End Sub

Private Sub cmbGroup_Click()
If cmbGroup.ListIndex = -1 Then Exit Sub
If lstReportType.ListIndex = -1 Then Exit Sub
If cmbPeriodPrint.ListIndex = -1 Then Exit Sub

If lstReportType.ListIndex = 0 Then
    If cmbGroup.ListIndex <> 3 Then
        Select Case cmbGroup.ListIndex
            Case 0
                s = "SELECT tbl_Personnel_Compensation.Dept as PK, " & _
                    " tbl_Personnel_Department.DepartmentName as Name " & _
                    " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
                    " tbl_Personnel_Department ON tbl_Personnel_Compensation.Dept = tbl_Personnel_Department.PK " & _
                    " Where (tbl_Personnel_Compensation.Period = " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ") " & _
                    " GROUP BY tbl_Personnel_Compensation.Dept, tbl_Personnel_Department.DepartmentName " & _
                    " ORDER BY tbl_Personnel_Department.DepartmentName"
            Case 1
                s = "SELECT tbl_Personnel_Compensation.Status AS PK, " & _
                    " tbl_Personnel_EmploymentStatus.StatusName AS Name " & _
                    " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
                    " tbl_Personnel_EmploymentStatus ON tbl_Personnel_Compensation.Status = tbl_Personnel_EmploymentStatus.PK " & _
                    " Where (tbl_Personnel_Compensation.Period = " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ") " & _
                    " GROUP BY tbl_Personnel_Compensation.Status, tbl_Personnel_EmploymentStatus.StatusName " & _
                    " ORDER BY tbl_Personnel_EmploymentStatus.StatusName"
            Case 2
                s = "SELECT tbl_Personnel_Compensation.Positions as PK, " & _
                    " tbl_Personnel_Position.PositionName as Name " & _
                    " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
                    " tbl_Personnel_Position ON tbl_Personnel_Compensation.Positions = tbl_Personnel_Position.PK " & _
                    " Where (tbl_Personnel_Compensation.Period = " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ") " & _
                    " GROUP BY tbl_Personnel_Compensation.Positions, tbl_Personnel_Position.PositionName " & _
                    " ORDER BY tbl_Personnel_Position.PositionName"
        End Select
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        txtSearchPrint.Text = ""
        txtSearch.Visible = False
        lstResultPrint.Top = 1155
        lstResultPrint.Height = 4935 '4350 '4155
        With lstResultPrint
            .Clear
            While Not rs.EOF
                .AddItem rs!Name
                .ItemData(.NewIndex) = rs!PK
                rs.MoveNext
            Wend
            .AddItem "SUPERVISORY"
            .ItemData(.NewIndex) = 0
            If .ListCount Then .ListIndex = 0
        End With
        rs.Close
        
    Else
        txtSearchPrint.Text = ""
        txtSearchPrint.Visible = True
        lstResultPrint.Clear
        lstResultPrint.Top = 1560
        lstResultPrint.Height = 4545 '3765
    End If
    
Else
    txtSearchPrint.Text = ""
    txtSearchPrint.Visible = False
    lstResultPrint.Clear
    lstResultPrint.Top = 1155
    lstResultPrint.Height = 4935 '4350 '4155
End If
End Sub

Private Sub cmbGroup_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If txtSearchPrint.Visible = True Then txtSearchPrint.SetFocus Else lstResultPrint.SetFocus
    Exit Sub
End If
End Sub

Private Sub cmbPeriod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKSearch_Click
End Sub

Private Sub cmbPeriodPrint_Click()
cmbGroup_Click
End Sub

Private Sub cmbPeriodPrint_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbGroup.SetFocus
End Sub

Private Sub cmdCancelAdd_Click()
picToolbar.Enabled = True
picBody.Enabled = True
picAdd.Visible = False
End Sub

Private Sub cmdCancelPrint_Click()
picToolbar.Enabled = True
picBody.Enabled = True
picPrint.Visible = False
End Sub

Private Sub cmdCancelSearch_Click()
picToolbar.Enabled = True
picBody.Enabled = True
picSearch.Visible = False
End Sub

Private Sub cmdCancelTaxAlpha_Click()
picTaxWithHeldAlpha.Visible = False
picPrint.Enabled = True
End Sub

Private Sub cmdOKAdd_Click()
If lstResultAdd.ListIndex = -1 Then Exit Sub
If IsDate(txtFrom.Text) = False Then Exit Sub
If IsDate(txtTo.Text) = False Then Exit Sub
If RETURNTEXTVALUE(txtNoHoursAdd) <= 0 Then MsgBox "Please Supply Number of Hours!                        ", vbCritical, "Error...": txtNoHoursAdd.SetFocus: Exit Sub
If CHECK_IF_HAVE_ACTION(lstResultAdd.ItemData(lstResultAdd.ListIndex)) = 0 Then
    MsgBox "NO PRESENT ACTION MEMO!     " & vbCrLf & _
           "" & vbCrLf & _
           "Cannot Computer Salary No Basic Rate.    ", vbInformation, ""
    txtSearchAdd.SetFocus
    HTEXT txtSearchAdd
    Exit Sub
End If
If GET_PERIOD(FormatDateTime(txtFrom.Text, vbShortDate), FormatDateTime(txtTo.Text, vbShortDate), _
GET_DIVISION(lstResultAdd.ItemData(lstResultAdd.ListIndex), Date)) = 0 Then
    MsgBox "Payroll Period Not Match to the Employee Division!      ", vbInformation, ""
    txtFrom.SetFocus
    HTEXT txtFrom
    Exit Sub
End If

If CHECK_COMPENSATION_LOCKED(GET_DIVISION(lstResultAdd.ItemData(lstResultAdd.ListIndex), Date), _
GET_PERIOD(FormatDateTime(txtFrom.Text, vbShortDate), FormatDateTime(txtTo.Text, vbShortDate), _
GET_DIVISION(lstResultAdd.ItemData(lstResultAdd.ListIndex), Date))) = True Then MsgBox "Payroll Period Already Locked!                          ", vbCritical, "Error...": Exit Sub

s = "SELECT TOP 1 tbl_Personnel_Compensation.Locked " & _
    " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
    " tbl_Personnel_Compensation_Period ON tbl_Personnel_Compensation.Period = tbl_Personnel_Compensation_Period.PK " & _
    " WHERE (tbl_Personnel_Compensation.Division = " & GET_DIVISION(lstResultAdd.ItemData(lstResultAdd.ListIndex), Date) & ") " & _
    " AND (tbl_Personnel_Compensation.Locked = 0) " & _
    " AND (tbl_Personnel_Compensation_Period.DateTo < '" & FormatDateTime(txtTo.Text, vbShortDate) & "') "
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    MsgBox "Please locked previous payroll period!                          ", vbCritical, "Error...": Exit Sub
End If
rs.Close

CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING

txtPeriod.Text = GET_PERIOD(FormatDateTime(txtFrom.Text, vbShortDate), FormatDateTime(txtTo.Text, vbShortDate), GET_DIVISION(lstResultAdd.ItemData(lstResultAdd.ListIndex), Date))
Array1 = Split(GET_EMPLOYEE_INFO(txtTo.Text, lstResultAdd.ItemData(lstResultAdd.ListIndex)), ";", -1, 1)
txtEmpPK.Text = lstResultAdd.ItemData(lstResultAdd.ListIndex)
iEmployeeKey = lstResultAdd.ItemData(lstResultAdd.ListIndex)
txtDivision.Text = CInt(Array1(0))
If CInt(Array1(0)) = 1 Then
    txtDivName.Text = "CLUB HOUSE"
ElseIf CInt(Array1(0)) = 2 Then
    txtDivName.Text = "MAINTENANCE"
Else
    txtDivName.Text = ""
End If
txtDeptKey.Text = CDbl(Array1(1))
txtDept.Text = CStr(Array1(2))
txtStatusKey.Text = CDbl(Array1(3))
txtPostKey.Text = CDbl(Array1(4))
txtPost.Text = CStr(Array1(5))
txtRatePerHour.Text = CDbl(Array1(6))

txtActionPK.Text = CDbl(Array1(9))

'txtPeriod.Text = txtPeriod.Text
txtName.Text = CStr(Array1(7)) & " - " & CStr(Array1(8))
txtPayrollPeriod.Text = Format(FormatDateTime(txtFrom.Text, vbShortDate), "mm/dd/yyyy") & " - " & Format(FormatDateTime(txtTo.Text, vbShortDate), "mm/dd/yyyy")

txtColaPerHour.Text = CDbl(Array1(11))
txtAllowPerHour.Text = CDbl(Array1(12))

imgPicture.Picture = LoadPicture(SHOW_IMAGES(Array1(13), 0, "Employee Profile"))

txtNoHours.Text = txtNoHoursAdd.Text
txtColaHrs.Text = txtColaHrsAdd.Text
txtSH.Text = txtSHAdd.Text
txtLH.Text = txtLHAdd.Text
txtSL.Text = txtSLAdd.Text
txtAdjustment.Text = txtAdjustmentAdd.Text
txtRegOT.Text = txtRegOTAdd.Text
txtRDOT.Text = txtRDOTAdd.Text
txtSHOT.Text = txtSHOTAdd.Text
txtLHOT.Text = txtLHOTAdd.Text

StatusBar.Panels(3).Text = "" '"RATE: " & Format(CDbl(Array1(10)), "##,##0.00")

If CHECK_LOAN_CUTOFF("SSS_Loan", Day(CDate(FormatDateTime(txtFrom.Text, vbShortDate)))) = True Then
    txtIsLoan.Text = "1"
Else
    txtIsLoan.Text = "0"
End If
If CHECK_CONT_CUTOFF("SSS", Day(CDate(FormatDateTime(txtFrom.Text, vbShortDate)))) = True Then
    txtIsCont.Text = "1"
Else
    txtIsCont.Text = "0"
End If

chkLoan.Value = RETURNTEXTVALUE(txtIsLoan)
chkContribution.Value = RETURNTEXTVALUE(txtIsCont)

txtMortuary.Text = "0.00"

s = "SELECT NoOfMortuary " & _
    " From tbl_Personnel_Compensation_Mortuary " & _
    " WHERE (Period = " & RETURNTEXTVALUE(txtPeriod) & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    t = "SELECT PositionLevel " & _
        " From tbl_Personnel_Position " & _
        " WHERE (PK = " & RETURNTEXTVALUE(txtPostKey) & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        u = "SELECT TOP 1 tbl_Personnel_Compensation_Mortuary_Rate.* " & _
            " FROM tbl_Personnel_Compensation_Mortuary_Rate " & _
            " WHERE (PositionLevel = " & rt!PositionLevel & ") " & _
            " AND (EffectDate <= '" & FormatDateTime(txtTo.Text, vbShortDate) & "') " & _
            " ORDER BY EffectDate DESC"
        If ru.State = adStateOpen Then ru.Close
        ru.Open u, ConnOmega
        If ru.RecordCount > 0 Then
            txtMortuary.Text = Format(CDbl(ru!Rate) * CDbl(rs!NoOfMortuary), "#,##0.00")
        End If
        ru.Close
    End If
    rt.Close
End If
rs.Close

lstDeduction.ListItems.Clear
For i = 1 To 5
    dDedAmt = 0: dDedTotAmt = 0
    s = "SELECT tbl_Personnel_Deduction_Summary.* " & _
        " FROM tbl_Personnel_Deduction_Summary " & _
        " WHERE (DeductionType = " & i & ") " & _
        " AND (EmployeeKey = " & iEmployeeKey & ") " & _
        " AND (InOut = 'I') " & _
        " AND (Cleared = 0) " & _
        " AND (TransDate <= '" & FormatDateTime(txtTo.Text, vbShortDate) & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        Set x = lstDeduction.ListItems.Add()
        x.Text = ""
        x.SubItems(1) = i
        x.SubItems(2) = rs!InRefKey
        t = "SELECT tbl_Personnel_Deduction.* " & _
            " FROM tbl_Personnel_Deduction " & _
            " WHERE (PK = " & rs!InRefKey & ")"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            dDedAmt = CDbl(rt!Amount)
            If rt!OneTimeDed = 1 Then
                If CDbl(dDedAmt) > (CDbl(rs!AmountIn) - CDbl(rs!AmountUsed)) Then
                    dDedAmt = (CDbl(rt!AmountIn) - CDbl(rt!AmountUsed))
                Else
                    dDedAmt = dDedAmt
                End If
                dDedTotAmt = dDedTotAmt + CDbl(dDedAmt)
            Else
                If CDbl(dDedAmt) > (CDbl(rs!AmountIn) - CDbl(rs!AmountUsed)) Then
                    dDedAmt = (CDbl(rt!AmountIn) - CDbl(rt!AmountUsed))
                Else
                    dDedAmt = dDedAmt
                End If
                dDedTotAmt = dDedTotAmt + CDbl(dDedAmt)
            End If
            x.SubItems(3) = dDedAmt
        Else
            x.SubItems(3) = "0"
        End If
        x.SubItems(4) = rs!PK
        rs.MoveNext
    Wend
    rs.Close
    
    Select Case i
        Case 1
            txtAROthers.Text = Format(dDedTotAmt, "#,##0.00")
        Case 2
            txtAdvances.Text = Format(dDedTotAmt, "#,##0.00")
        Case 3
            txtShortages.Text = Format(dDedTotAmt, "#,##0.00")
        Case 4
            txtUniform.Text = Format(dDedTotAmt, "#,##0.00")
        Case 5
            txtOthers.Text = Format(dDedTotAmt, "#,##0.00")
    End Select
    
Next i
        
cmdCancelAdd_Click

txtNoHours.SetFocus
'    Else
'
'    End If
''Else
'
'End If
End Sub

Private Function PAYROLL_TEMP(Division, Period)


picProgressBar.BackColor = &HFFFFFF
picPrint.Enabled = False
picProgress.ZOrder 0
picProgress.Visible = True
i = 0

ConnOmega.Execute "DELETE FROM tbl_PersonnelPayroll_Tmp WHERE (LogInName = '" & gbl_UserName & "')"

's = "SELECT qry_Payroll_Transaction.*" & _
    " From qry_Payroll_Transaction " & _
    " WHERE (Division = " & Division & ") " & _
    " AND (Period = " & Period & ")"
s = "sp_Personnel_Compensation_Print(" & Division & ", " & Period & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    DoEvents
    i = i + 1
    
    ConnOmega.Execute "INSERT INTO tbl_PersonnelPayroll_Tmp " & _
                      " (LogInName, EmpPK, Division, Dept, Status, Positions, Period, ActionMemo, NoHours, SH_Hours, LH_Hours, SL_Hours, Adjustment, Reg_OT_Hours, " & _
                      " RD_OT_Hours, SH_OT_Hours, LH_OT_Hours, Amount_Earned, SH_Amount, LH_Amount, SL_Amount, Reg_OT_Amount, RD_OT_Amount, SH_OT_Amount, LH_OT_Amount, " & _
                      " TotalEarning, Mortuary, AR_Others, Advances, Shortages, Uniforms, Others, Is_Have_Loan, SSSLoan_No, SSSLoan, SSSBalance, PagIbigLoan_No, " & _
                      " PagIbigLoan, PagIbigBalance, Is_Have_Cont, SSS, SSS_Employer, SSS_EC, PHIC, PHIC_Employer, PagIbig, PagIbig_Employer, WithHeld, TotalDeduction, " & _
                      " NetEarning, Locked, IDNumber, LName, FName, MName, BDate, DepartmentName, StatusName, PositionName, DateFrom, DateTo, Type, CompensationRate, " & _
                      " SSSNo, Is_PHIC, PHICNo, IDName, Is_TIN, TIN, Basic, RatePerHour, TotalCola, TotalAllowance) " & _
                      " VALUES ('" & gbl_UserName & "', " & rs!EmpPK & ", " & rs!Division & ", " & rs!Dept & ", " & rs!Status & ", " & rs!Positions & ", " & rs!Period & ", " & rs!ActionMemo & ", " & _
                      " " & CDbl(rs!NoHours) & ", " & CDbl(rs!SH_Hours) & ", " & CDbl(rs!LH_Hours) & ", " & CDbl(rs!SL_Hours) & ", " & CDbl(rs!Adjustment) & ", " & _
                      " " & CDbl(rs!Reg_OT_Hours) & ", " & CDbl(rs!RD_OT_Hours) & ", " & CDbl(rs!SH_OT_Hours) & ", " & CDbl(rs!LH_OT_Hours) & ", " & CDbl(rs!Amount_Earned) & ", " & _
                      " " & CDbl(rs!SH_Amount) & ", " & CDbl(rs!LH_Amount) & ", " & CDbl(rs!SL_Amount) & ", " & CDbl(rs!Reg_OT_Amount) & ", " & CDbl(rs!RD_OT_Amount) & ", " & _
                      " " & CDbl(rs!SH_OT_Amount) & ", " & CDbl(rs!LH_OT_Amount) & ", " & CDbl(rs!TotalEarning) & ", " & CDbl(rs!Mortuary) & ", " & CDbl(rs!AR_Others) & ", " & _
                      " " & CDbl(rs!Advances) & ", " & CDbl(rs!Shortages) & ", " & CDbl(rs!Uniforms) & ", " & CDbl(rs!Others) & ", " & rs!Is_Have_Loan & ", " & _
                      " " & rs!SSSLoan_No & ", " & CDbl(rs!SSSLoan) & ", " & CDbl(rs!SSSBalance) & ", " & rs!PagIbigLoan_No & ", " & CDbl(rs!PagIbigLoan) & ", " & _
                      " " & CDbl(rs!PagIbigBalance) & ", " & rs!Is_Have_Cont & ", " & CDbl(rs!SSS) & ", " & CDbl(rs!SSS_Employer) & ", " & CDbl(rs!SSS_EC) & ", " & CDbl(rs!PHIC) & ", " & _
                      " " & CDbl(rs!PHIC_Employer) & ", " & CDbl(rs!PagIbig) & ", " & CDbl(rs!PagIbig_Employer) & ", " & CDbl(rs!WithHeld) & ", " & CDbl(rs!TotalDeduction) & ", " & _
                      " " & CDbl(rs!NetEarning) & ", " & rs!Locked & ", '" & rs!IDNumber & "', '" & FORMATSQL(rs!LName) & "', '" & FORMATSQL(rs!FName) & "', '" & FORMATSQL(rs!MName) & "', " & _
                      " '" & FormatDateTime(rs!BDate, vbShortDate) & "' , '" & FORMATSQL(rs!DepartmentName) & "', '" & FORMATSQL(rs!StatusName) & "', '" & FORMATSQL(rs!PositionName) & "', " & _
                      " '" & FormatDateTime(rs!DateFrom, vbShortDate) & "', '" & FormatDateTime(rs!DateTo, vbShortDate) & "', " & rs!Type & ", " & rs!CompensationRate & ", " & _
                      " '" & rs!SSSNo & "', " & rs!Is_PHIC & ", '" & rs!PHICNo & "', '" & FORMATSQL(rs!IDName) & "', " & rs!Is_TIN & ", '" & rs!TIN & "', " & CDbl(rs!Basic) & ", " & CDbl(rs!RatePerHour) & ", " & CDbl(rs!TotalCola) & ", " & _
                      " " & CDbl(rs!TotalAllowance) & ")"
    
    UpdateProgress picProgressBar, i / rs.RecordCount
    rs.MoveNext
Wend
rs.Close

s = "SELECT tbl_PersonnelPayroll_Tmp.* " & _
    " FROM tbl_PersonnelPayroll_Tmp"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
rs.Requery
rs.Close

picProgressBar.BackColor = &HFFFFFF
picProgress.Visible = False
picPrint.Enabled = True


End Function

Private Sub PAYROLL_TEMP_RF_SV(Division, Period, PostLevel)


frmPersonnelCompensation.picProgressBar.BackColor = &HFFFFFF
frmPersonnelCompensation.picPrint.Enabled = False
frmPersonnelCompensation.picProgress.ZOrder 0
frmPersonnelCompensation.picProgress.Visible = True
i = 0

ConnOmega.Execute "DELETE FROM tbl_PersonnelPayroll_Tmp WHERE (LogInName = '" & gbl_UserName & "')"

's = "SELECT qry_Payroll_Transaction.*" & _
    " From qry_Payroll_Transaction " & _
    " WHERE (Division = " & Division & ") " & _
    " AND (Period = " & Period & ")"
s = "sp_Personnel_Compensation_Print_RF_SV(" & Division & ", " & Period & ", " & PostLevel & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    DoEvents
    i = i + 1
    
    ConnOmega.Execute "INSERT INTO tbl_PersonnelPayroll_Tmp " & _
                      " (LogInName, EmpPK, Division, Dept, Status, Positions, Period, ActionMemo, NoHours, SH_Hours, LH_Hours, SL_Hours, Adjustment, Reg_OT_Hours, " & _
                      " RD_OT_Hours, SH_OT_Hours, LH_OT_Hours, Amount_Earned, SH_Amount, LH_Amount, SL_Amount, Reg_OT_Amount, RD_OT_Amount, SH_OT_Amount, LH_OT_Amount, " & _
                      " TotalEarning, Mortuary, AR_Others, Advances, Shortages, Uniforms, Others, Is_Have_Loan, SSSLoan_No, SSSLoan, SSSBalance, PagIbigLoan_No, " & _
                      " PagIbigLoan, PagIbigBalance, Is_Have_Cont, SSS, SSS_Employer, SSS_EC, PHIC, PHIC_Employer, PagIbig, PagIbig_Employer, WithHeld, TotalDeduction, " & _
                      " NetEarning, Locked, IDNumber, LName, FName, MName, BDate, DepartmentName, StatusName, PositionName, DateFrom, DateTo, Type, CompensationRate, " & _
                      " SSSNo, Is_PHIC, PHICNo, IDName, Is_TIN, TIN, Basic, RatePerHour, TotalCola, TotalAllowance, PostLevel) " & _
                      " VALUES ('" & gbl_UserName & "', " & rs!EmpPK & ", " & rs!Division & ", " & rs!Dept & ", " & rs!Status & ", " & rs!Positions & ", " & rs!Period & ", " & rs!ActionMemo & ", " & _
                      " " & CDbl(rs!NoHours) & ", " & CDbl(rs!SH_Hours) & ", " & CDbl(rs!LH_Hours) & ", " & CDbl(rs!SL_Hours) & ", " & CDbl(rs!Adjustment) & ", " & _
                      " " & CDbl(rs!Reg_OT_Hours) & ", " & CDbl(rs!RD_OT_Hours) & ", " & CDbl(rs!SH_OT_Hours) & ", " & CDbl(rs!LH_OT_Hours) & ", " & CDbl(rs!Amount_Earned) & ", " & _
                      " " & CDbl(rs!SH_Amount) & ", " & CDbl(rs!LH_Amount) & ", " & CDbl(rs!SL_Amount) & ", " & CDbl(rs!Reg_OT_Amount) & ", " & CDbl(rs!RD_OT_Amount) & ", " & _
                      " " & CDbl(rs!SH_OT_Amount) & ", " & CDbl(rs!LH_OT_Amount) & ", " & CDbl(rs!TotalEarning) & ", " & CDbl(rs!Mortuary) & ", " & CDbl(rs!AR_Others) & ", " & _
                      " " & CDbl(rs!Advances) & ", " & CDbl(rs!Shortages) & ", " & CDbl(rs!Uniforms) & ", " & CDbl(rs!Others) & ", " & rs!Is_Have_Loan & ", " & _
                      " " & rs!SSSLoan_No & ", " & CDbl(rs!SSSLoan) & ", " & CDbl(rs!SSSBalance) & ", " & rs!PagIbigLoan_No & ", " & CDbl(rs!PagIbigLoan) & ", " & _
                      " " & CDbl(rs!PagIbigBalance) & ", " & rs!Is_Have_Cont & ", " & CDbl(rs!SSS) & ", " & CDbl(rs!SSS_Employer) & ", " & CDbl(rs!SSS_EC) & ", " & CDbl(rs!PHIC) & ", " & _
                      " " & CDbl(rs!PHIC_Employer) & ", " & CDbl(rs!PagIbig) & ", " & CDbl(rs!PagIbig_Employer) & ", " & CDbl(rs!WithHeld) & ", " & CDbl(rs!TotalDeduction) & ", " & _
                      " " & CDbl(rs!NetEarning) & ", " & rs!Locked & ", '" & rs!IDNumber & "', '" & FORMATSQL(rs!LName) & "', '" & FORMATSQL(rs!FName) & "', '" & FORMATSQL(rs!MName) & "', " & _
                      " '" & FormatDateTime(rs!BDate, vbShortDate) & "' , '" & FORMATSQL(rs!DepartmentName) & "', '" & FORMATSQL(rs!StatusName) & "', '" & FORMATSQL(rs!PositionName) & "', " & _
                      " '" & FormatDateTime(rs!DateFrom, vbShortDate) & "', '" & FormatDateTime(rs!DateTo, vbShortDate) & "', " & rs!Type & ", " & rs!CompensationRate & ", " & _
                      " '" & rs!SSSNo & "', " & rs!Is_PHIC & ", '" & rs!PHICNo & "', '" & FORMATSQL(rs!IDName) & "', " & rs!Is_TIN & ", '" & rs!TIN & "', " & CDbl(rs!Basic) & ", " & _
                      " " & CDbl(rs!RatePerHour) & ", " & CDbl(rs!TotalCola) & ", " & CDbl(rs!TotalAllowance) & ", " & rs!PositionLevel & ")"
    
    UpdateProgress frmPersonnelCompensation.picProgressBar, i / rs.RecordCount
    rs.MoveNext
Wend
rs.Close

s = "SELECT tbl_PersonnelPayroll_Tmp.* " & _
    " FROM tbl_PersonnelPayroll_Tmp"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
rs.Requery
rs.Close

frmPersonnelCompensation.picProgressBar.BackColor = &HFFFFFF
frmPersonnelCompensation.picProgress.Visible = False
frmPersonnelCompensation.picPrint.Enabled = True


End Sub


Private Sub cmdOKPrint_Click()

If cmbPeriodPrint.ListIndex = -1 Then Exit Sub
txtTerms.Text = GET_TERMS(cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex))
Select Case lstReportType.ListIndex
    Case 0  'PAYSLIP
        'PAYROLL_TEMP cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex)
        
        Select Case cmbGroup.ListIndex
            Case 0  'DEPT
                If lstResultPrint.ItemData(lstResultPrint.ListIndex) = 0 Then
                    PAYROLL_TEMP_RF_SV cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex), 2
                    frmCrystalReportViewer.PRINT_PAYSLIP_SUPERVISORY gbl_CompanyName, gbl_UserName, 2
                    If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
                Else
                    PAYROLL_TEMP_RF_SV cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex), 1
                    frmCrystalReportViewer.PRINT_PAYSLIP_DEPT gbl_CompanyName, gbl_UserName, lstResultPrint.ItemData(lstResultPrint.ListIndex)
                    If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
                End If
            Case 1  'STATUS
                If lstResultPrint.ItemData(lstResultPrint.ListIndex) = 0 Then
                    PAYROLL_TEMP_RF_SV cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex), 2
                    frmCrystalReportViewer.PRINT_PAYSLIP_SUPERVISORY gbl_CompanyName, gbl_UserName, 2
                    If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
                Else
                    PAYROLL_TEMP_RF_SV cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex), 1
                    frmCrystalReportViewer.PRINT_PAYSLIP_STATUS gbl_CompanyName, gbl_UserName, lstResultPrint.ItemData(lstResultPrint.ListIndex)
                    If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
                End If
                
            Case 2  'POSITION
                If lstResultPrint.ItemData(lstResultPrint.ListIndex) = 0 Then
                    PAYROLL_TEMP_RF_SV cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex), 2
                    frmCrystalReportViewer.PRINT_PAYSLIP_SUPERVISORY gbl_CompanyName, gbl_UserName, 2
                    If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
                Else
                    PAYROLL_TEMP_RF_SV cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex), 1
                    frmCrystalReportViewer.PRINT_PAYSLIP_POST gbl_CompanyName, gbl_UserName, lstResultPrint.ItemData(lstResultPrint.ListIndex)
                If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
                End If
                
            Case 3  'EMPLOYEE
                frmCrystalReportViewer.PRINT_PAYSLIP_EMPLOYEE gbl_CompanyName, gbl_UserName, lstResultPrint.ItemData(lstResultPrint.ListIndex)
                If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
            Case Else: Exit Sub
        End Select
    Case 1  'SIGNATURE LEDGER
        
        PopupMenu MainFormPopupF.mnuCompensationPrint, , picPrint.Left + cmdOKPrint.Left + 200, picPrint.Top + cmdOKPrint.Top + 200
        
        'PAYROLL_TEMP cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex)
        
        'frmCrystalReportViewer.PRINT_SIGNATURE_LEDGER gbl_CompanyName, gbl_UserName
        'If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
        
    Case 2  'COMPENSATION (TOP)
        
        PopupMenu MainFormPopupF.mnuCompensationPrint, , picPrint.Left + cmdOKPrint.Left + 200, picPrint.Top + cmdOKPrint.Top + 200
        
        'PAYROLL_TEMP cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex)
        
        'frmCrystalReportViewer.PRINT_COMPENSATION_SUMMARY_TOP gbl_CompanyName, gbl_UserName
        'If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
        
    Case 3  'COMPENSATION
        
        PopupMenu MainFormPopupF.mnuCompensationPrint, , picPrint.Left + cmdOKPrint.Left + 200, picPrint.Top + cmdOKPrint.Top + 200
        
        'PAYROLL_TEMP cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex)
        
'        frmCrystalReportViewer.PRINT_COMPENSATION_SUMMARY gbl_CompanyName, gbl_UserName
'        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
        
    Case 4  'DEDUCTION (TOP)
    
        PopupMenu MainFormPopupF.mnuCompensationPrint, , picPrint.Left + cmdOKPrint.Left + 200, picPrint.Top + cmdOKPrint.Top + 200
        
'        PAYROLL_TEMP cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex)
'        frmCrystalReportViewer.PRINT_DEDUCTION_SUMMARY_TOP gbl_CompanyName, gbl_UserName
'        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
        
        
    Case 5  'DEDUCTION
        
        PopupMenu MainFormPopupF.mnuCompensationPrint, , picPrint.Left + cmdOKPrint.Left + 200, picPrint.Top + cmdOKPrint.Top + 200
        
'        PAYROLL_TEMP cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex)
'        frmCrystalReportViewer.PRINT_DEDUCTION_SUMMARY gbl_CompanyName, gbl_UserName
'        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
        
    Case 6  'SSS LOANS TOP
    
    Case 7  'SSS LOANS
        
        PAYROLL_TEMP cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex)
        frmCrystalReportViewer.PRINT_SSS_LOAN gbl_CompanyName, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
        
    Case 8  'PAG IBIG LOANS TOP
            
    Case 9  'PAG IBIG LOANS
        
        PAYROLL_TEMP cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex)
        frmCrystalReportViewer.PRINT_PAGIBIG_LOAN gbl_CompanyName, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
        
    Case 10 'SSS TOP
    
    Case 11 'SSS
        
        PAYROLL_TEMP cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex)
        frmCrystalReportViewer.PRINT_SSS_COLLECTION gbl_CompanyName, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
        
    Case 12 'PHIC TOP
    
    Case 13 'PHIC
        
        PAYROLL_TEMP cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex)
        frmCrystalReportViewer.PRINT_PHIC_COLLECTION gbl_CompanyName, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
        
    Case 14 'PAG IBIG TOP
    
    Case 15 'PAG IBIG
    
        PAYROLL_TEMP cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex)
        frmCrystalReportViewer.PRINT_PAGIBIG_COLLECTION gbl_CompanyName, gbl_CompanyTelNo, gbl_CompanySSSNo, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
        
    Case 16 'WITHHELD TOP
        
        picTaxWithHeldAlpha.ZOrder 0
        picPrint.Enabled = False
        txtTaxAlphaYear.Text = Year(Now)
        picTaxWithHeldAlpha.Visible = True
        txtTaxAlphaYear.SetFocus
        
    Case 17 'WITHHELD
        
        'PAYROLL_TEMP cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex)
        
        picProgressBar.BackColor = &HFFFFFF
        picPrint.Enabled = False
        picProgress.ZOrder 0
        picProgress.Visible = True
        i = 0
        
        ConnOmega.Execute "DELETE FROM tbl_Personnel_TaxWithHeld WHERE(LogInName = '" & gbl_UserName & "')"
        
        s = "SELECT ISNULL(tbl_Personnel_Department.DepartmentName, '') AS DepartmentName, " & _
            " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
            " tbl_Personnel_Information.TIN as TinNumber, tbl_Personnel_Compensation.TotalEarning, tbl_Personnel_Compensation.WithHeld, " & _
            " tbl_Personnel_Compensation.Division, tbl_Personnel_Compensation.Period, tbl_Personnel_Compensation_Period.DateFrom, " & _
            " tbl_Personnel_Compensation_Period.DateTo , tbl_Personnel_Compensation.EmpPK " & _
            " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
            " tbl_Personnel_Action ON tbl_Personnel_Compensation.ActionMemo = tbl_Personnel_Action.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_IDNumber ON tbl_Personnel_Compensation.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_Department ON tbl_Personnel_Compensation.Dept = tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_Compensation_Period ON " & _
            " tbl_Personnel_Compensation.Period = tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
            " WHERE (tbl_Personnel_Compensation.Is_Have_Cont = 1) " & _
            " AND (tbl_Personnel_Compensation.Division = " & cmbDivision.ListIndex + 1 & ") " & _
            " AND (tbl_Personnel_Compensation.Period = " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ") " & _
            " AND (tbl_Personnel_Action.Is_TIN = 1) " & _
            " ORDER BY tbl_Personnel_Department.DepartmentName, " & _
            " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        While Not rs.EOF
            DoEvents
            i = i + 1
            
            ConnOmega.Execute "INSERT INTO tbl_Personnel_TaxWithHeld" & _
                              " (LogInName, Department, sTIN, sName, Gross1, Gross2, Tax, DateFrom, DateTo)" & _
                              " VALUES('" & gbl_UserName & "', '" & FORMATSQL(rs!DepartmentName) & "', " & _
                              " '" & rs!TinNumber & "', '" & FORMATSQL(rs!EmployeeName) & "', " & _
                              " " & CDbl(rs!TotalEarning) & ", " & GET_PREVIOUS_GROSS(rs!Period, cmbDivision.ListIndex + 1, rs!EmpPK) & ", " & _
                              " " & CDbl(rs!WithHeld) & ", '" & FormatDateTime(rs!DateFrom, vbShortDate) & "', " & _
                              " '" & FormatDateTime(rs!DateTo, vbShortDate) & "')"
            
            UpdateProgress picProgressBar, i / rs.RecordCount
            rs.MoveNext
        Wend
        rs.Close
        
        picProgressBar.BackColor = &HFFFFFF
        picProgress.Visible = False
        picPrint.Enabled = True
        
        frmCrystalReportViewer.PRINT_TAX_COLLECTION cmbDivision.ListIndex + 1, gbl_CompanyName, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
        
    Case 18 '13th MONTH (TOP SHEET)
    
    Case 19 '13th MONTH
        
        Arr = Split(cmbPeriodPrint.List(cmbPeriodPrint.ListIndex), " - ", -1, 1)
        
        Select Case Month(FormatDateTime(Arr(1), vbShortDate))
            Case 12
                iMonth1 = 9
                iMonth2 = 10
                iMonth3 = 11
                iYear1 = Year(FormatDateTime(Arr(1), vbShortDate))
                iYear2 = Year(FormatDateTime(Arr(1), vbShortDate))
                iYear3 = Year(FormatDateTime(Arr(1), vbShortDate))
            Case 3
                iMonth1 = 12
                iMonth2 = 1
                iMonth3 = 2
                iYear1 = Year(FormatDateTime(Arr(1), vbShortDate)) - 1
                iYear2 = Year(FormatDateTime(Arr(1), vbShortDate))
                iYear3 = Year(FormatDateTime(Arr(1), vbShortDate))
            Case 6
                iMonth1 = 3
                iMonth2 = 4
                iMonth3 = 5
                iYear1 = Year(FormatDateTime(Arr(1), vbShortDate))
                iYear2 = Year(FormatDateTime(Arr(1), vbShortDate))
                iYear3 = Year(FormatDateTime(Arr(1), vbShortDate))
            Case 9
                iMonth1 = 6
                iMonth2 = 7
                iMonth3 = 8
                iYear1 = Year(FormatDateTime(Arr(1), vbShortDate))
                iYear2 = Year(FormatDateTime(Arr(1), vbShortDate))
                iYear3 = Year(FormatDateTime(Arr(1), vbShortDate))
            Case Else: Exit Sub
        End Select
        
        picProgressBar.BackColor = &HFFFFFF
        picPrint.Enabled = False
        picProgress.ZOrder 0
        picProgress.Visible = True
        i = 0
        
        dtmDateTo = DateSerial(iYear3, iMonth3 + 1, 0)
        
        ConnOmega.Execute "DELETE FROM tbl_Personnel_13thMonth WHERE (LogInName = '" & gbl_UserName & "')"
        
        s = "sp_13th_Month_Report(" & iMonth1 & ", " & iYear1 & ", " & iMonth2 & ", " & iYear2 & ", " & iMonth3 & ", " & iYear3 & ", " & cmbDivision.ListIndex + 1 & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        While Not rs.EOF
            DoEvents
            i = i + 1
            sDeptName = "" 'rs!DepartmentName
            
            t = "SELECT TOP 1 tbl_Personnel_Department.DepartmentName " & _
                " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
                " tbl_Personnel_Compensation_Period ON " & _
                " tbl_Personnel_Compensation.Period = tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " tbl_Personnel_Department ON tbl_Personnel_Compensation.Dept = tbl_Personnel_Department.PK " & _
                " WHERE (tbl_Personnel_Compensation.EmpPK = " & rs!EmpPK & ") " & _
                " AND (tbl_Personnel_Compensation_Period.DateTo <= '" & DateSerial(iYear3, iMonth3 + 1, 0) & "') " & _
                " ORDER BY tbl_Personnel_Compensation_Period.DateTo DESC"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                sDeptName = rt!DepartmentName
            End If
            rt.Close
            
            ConnOmega.Execute "INSERT INTO tbl_Personnel_13thMonth " & _
                              " (LogInName, Department, IDNumber, sName, " & _
                              " Basic1, Basic2, Basic3) " & _
                              " VALUES ('" & gbl_UserName & "', " & _
                              " '" & FORMATSQL(CStr(sDeptName)) & "', '" & rs!IDNumber & "', " & _
                              " '" & FORMATSQL(rs!EmployeeName) & "', " & CDbl(rs!iMonth1) & ", " & _
                              " " & CDbl(rs!iMonth2) & ", " & CDbl(rs!iMonth3) & ")"
            
            UpdateProgress_Caption rs!EmployeeName, picProgressBar, i / rs.RecordCount
            rs.MoveNext
        Wend
        rs.Close
        
        picProgressBar.BackColor = &HFFFFFF
        picProgress.Visible = False
        picPrint.Enabled = True
        
        frmCrystalReportViewer.PRINT_13TH_MONTH cmbDivision.ListIndex + 1, Month(FormatDateTime(Arr(1), vbShortDate)), gbl_CompanyName, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
    
    Case 20     'Cola Summary
        
        picProgressBar.BackColor = &HFFFFFF
        picPrint.Enabled = False
        picProgress.ZOrder 0
        picProgress.Visible = True
        i = 0

        ConnOmega.Execute "DELETE FROM tbl_PersonnelPayroll_Cola_Tmp WHERE (LogInName = '" & gbl_UserName & "')"
        
        s = "sp_Personnel_Compensation_Print_Cola(" & cmbDivision.ListIndex + 1 & ", " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        While Not rs.EOF
            i = i + 1
            
            ConnOmega.Execute "INSERT INTO tbl_PersonnelPayroll_Cola_Tmp " & _
                              " (LogInName, EmpPK, Division, Dept, Status, Positions, " & _
                              " Period, ActionMemo, IDNumber, LName, FName, MName, " & _
                              " BDate, DepartmentName, StatusName, PositionName, " & _
                              " DateFrom, DateTo, ColaPerDay, ColaPerHour, " & _
                              " ColaHour, TotalCola) " & _
                              " VALUES ('" & gbl_UserName & "', " & rs!EmpPK & ", " & _
                              " " & rs!Division & ", " & rs!Dept & ", " & rs!Status & ", " & _
                              " " & rs!Positions & ", " & rs!Period & ", " & rs!ActionMemo & ", " & _
                              " '" & rs!IDNumber & "', '" & FORMATSQL(rs!LName) & "', " & _
                              " '" & FORMATSQL(rs!FName) & "', '" & FORMATSQL(rs!MName) & "', " & _
                              " '" & FormatDateTime(rs!BDate, vbShortDate) & "' , " & _
                              " '" & FORMATSQL(rs!DepartmentName) & "', '" & FORMATSQL(rs!StatusName) & "', " & _
                              " '" & FORMATSQL(rs!PositionName) & "', '" & FormatDateTime(rs!DateFrom, vbShortDate) & "', " & _
                              " '" & FormatDateTime(rs!DateTo, vbShortDate) & "', " & _
                              " " & CDbl(rs!ColaPerDay) & ", " & CDbl(rs!ColaPerHour) & ", " & _
                              " " & CDbl(rs!ColaHours) & ", " & CDbl(rs!TotalCola) & ")"
            
            UpdateProgress picProgressBar, i / rs.RecordCount
            rs.MoveNext
        Wend
        rs.Close
        
        picProgressBar.BackColor = &HFFFFFF
        picProgress.Visible = False
        picPrint.Enabled = True
        
        If i > 0 Then
            frmCrystalReportViewer.PRINT_COLA_SUMMARY gbl_CompanyName, gbl_UserName
            If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
        End If
        
'    Case 21     'Allowance Summary
'
'        picProgressBar.BackColor = &HFFFFFF
'        picPrint.Enabled = False
'        picProgress.ZOrder 0
'        picProgress.Visible = True
'        i = 0
'
'        ConnOmega.Execute "DELETE FROM tbl_PersonnelPayroll_Cola_Tmp WHERE (LogInName = '" & gbl_UserName & "')"
'
'        s = "sp_Personnel_Compensation_Print_Allow(" & cmbDivision.ListIndex + 1 & ", " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ")"
'        If rs.State = adStateOpen Then rs.Close
'        rs.Open s, ConnOmega
'        While Not rs.EOF
'            i = i + 1
'
'            ConnOmega.Execute "INSERT INTO tbl_PersonnelPayroll_Cola_Tmp " & _
'                              " (LogInName, EmpPK, Division, Dept, Status, Positions, " & _
'                              " Period, ActionMemo, IDNumber, LName, FName, MName, " & _
'                              " BDate, DepartmentName, StatusName, PositionName, " & _
'                              " DateFrom, DateTo, AllowPerHour, AllowHour, " & _
'                              " TotalAllow, Allow15) " & _
'                              " VALUES ('" & gbl_UserName & "', " & rs!EmpPK & ", " & _
'                              " " & rs!Division & ", " & rs!Dept & ", " & rs!Status & ", " & _
'                              " " & rs!Positions & ", " & rs!Period & ", " & rs!ActionMemo & ", " & _
'                              " '" & rs!IDNumber & "', '" & FORMATSQL(rs!LName) & "', " & _
'                              " '" & FORMATSQL(rs!FName) & "', '" & FORMATSQL(rs!MName) & "', " & _
'                              " '" & FormatDateTime(rs!BDate, vbShortDate) & "' , " & _
'                              " '" & FORMATSQL(rs!DepartmentName) & "', '" & FORMATSQL(rs!StatusName) & "', " & _
'                              " '" & FORMATSQL(rs!PositionName) & "', '" & FormatDateTime(rs!DateFrom, vbShortDate) & "', " & _
'                              " '" & FormatDateTime(rs!DateTo, vbShortDate) & "', " & _
'                              " " & CDbl(rs!AllowPerHour) & ", " & CDbl(rs!AllowHours) & ", " & _
'                              " " & CDbl(rs!TotalAllowance) & ", " & CDbl(rs!Allowance15) & ")"
'
'            UpdateProgress picProgressBar, i / rs.RecordCount
'            rs.MoveNext
'        Wend
'        rs.Close
'
'        picProgressBar.BackColor = &HFFFFFF
'        picProgress.Visible = False
'        picPrint.Enabled = True
'
'        If i > 0 Then
'            frmCrystalReportViewer.PRINT_ALLOWANCE_SUMMARY gbl_CompanyName, gbl_UserName
'            If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
'        End If
        
    Case 21     'To B P I
        
        PopupMenu MainFormPopupF.mnuCompensationPrint, , picPrint.Left + cmdOKPrint.Left + 200, picPrint.Top + cmdOKPrint.Top + 200
        

'
'        i = 0
'
'        s = "SELECT tbl_Personnel_IDNumber.AccountNumber, " & _
'            " tbl_Personnel_Information.FirstName, " & _
'            " tbl_Personnel_Information.MiddleName, " & _
'            " tbl_Personnel_Information.LastName, " & _
'            " tbl_Personnel_Compensation.TotalEarning, " & _
'            " tbl_Personnel_Compensation.TotalDeduction, " & _
'            " tbl_Personnel_Compensation.TotalCola, " & _
'            " tbl_Personnel_Compensation.TotalAllowance " & _
'            " FROM  tbl_Personnel_Compensation LEFT OUTER JOIN " & _
'            " tbl_Personnel_IDNumber ON tbl_Personnel_Compensation.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
'            " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
'            " WHERE (tbl_Personnel_IDNumber.AccountNumber <> '') " & _
'            " AND (tbl_Personnel_Compensation.Division = " & cmbDivision.ListIndex + 1 & ") " & _
'            " AND (tbl_Personnel_Compensation.Period = " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ") " & _
'            " ORDER BY tbl_Personnel_Information.LastName, tbl_Personnel_Information.FirstName, tbl_Personnel_Information.MiddleName"
'        If rs.State = adStateOpen Then rs.Close
'        rs.Open s, ConnOmega
'        If rs.RecordCount > 0 Then
'
'            Mainform.CommonDialog1.CancelError = True
'            On Error GoTo ErrorHandler
'            Mainform.CommonDialog1.DialogTitle = "Save"
'            Mainform.CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
'            Mainform.CommonDialog1.ShowSave
'            Filename = Trim(Mainform.CommonDialog1.Filename)
'
'            WorkbookName = CStr(Filename)
'
'            picProgressBar.BackColor = &HFFFFFF
'            picPrint.Enabled = False
'            picProgress.ZOrder 0
'            picProgress.Visible = True
'
'            iWorkSheet = 1
'            Set xlsApp = CreateObject("Excel.Application")
'            xlsApp.Visible = False
'            xlsApp.Workbooks.Add
'            xlsApp.DisplayAlerts = False
'            xlsApp.Workbooks(1).Sheets(2).Delete
'            xlsApp.Workbooks(1).Sheets(2).Delete
'            xlsApp.Workbooks(1).Sheets(iWorkSheet).Activate
'            xlsApp.Workbooks(1).Sheets(iWorkSheet).Name = "B P I"
'
'            RowCnt = RowCnt + 1
'            ColCnt = 0
'            ColCnt = ColCnt + 1
'            strRange = EXCEL_RANGE(ColCnt, RowCnt)
'            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = gbl_CompanyName
'            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
'            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
'            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
'
'            RowCnt = RowCnt + 1
'            ColCnt = 0
'            ColCnt = ColCnt + 1
'            strRange = EXCEL_RANGE(ColCnt, RowCnt)
'            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "For B P I ATM"
'            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
'            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
'            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
'
'            RowCnt = RowCnt + 1
'            ColCnt = 0
'            ColCnt = ColCnt + 1
'            strRange = EXCEL_RANGE(ColCnt, RowCnt)
'            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
'            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
'            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
'            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
'
'            RowCnt = RowCnt + 1
'            For k = 1 To 9
'                Select Case k
'                    Case 1: strValue = "Account Number"
'                    Case 2: strValue = "First Name"
'                    Case 3: strValue = "Middle Name"
'                    Case 4: strValue = "Last Name"
'                    Case 5: strValue = "Total Earnings"
'                    Case 6: strValue = "Total Deduction"
'                    Case 7: strValue = "C O L A"
'                    Case 8: strValue = "Allowance"
'                    Case 9: strValue = "Amount"
'                End Select
'                ColCnt = k
'                strRange = EXCEL_RANGE(ColCnt, RowCnt)
'                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = strValue
'                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
'                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
'                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
'                If k >= 5 And k <= 9 Then
'                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
'                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(ColCnt).ColumnWidth = 13
'                ElseIf k = 1 Then
'                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3
'                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(ColCnt).ColumnWidth = 20
'                Else
'                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 2
'                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(ColCnt).ColumnWidth = 20
'                End If
'            Next k
'            While Not rs.EOF
'                i = i + 1
'                strAmount = "="
'                RowCnt = RowCnt + 1
'                For k = 1 To 8
'                    strValue = rs.Fields(k - 1).Value
'                    ColCnt = k
'                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
'                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = strValue
'                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
'                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
'                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
'                    Select Case k
'                        Case 5: strAmount = strAmount & strRange
'                        Case 6: strAmount = strAmount & "-" & strRange
'                        Case 7: strAmount = strAmount & "+" & strRange
'                        Case 8: strAmount = strAmount & "+" & strRange
'                    End Select
'
'                    If k >= 5 And k <= 8 Then
'                        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
'                        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "#,##0.00"
'                    ElseIf k = 1 Then
'                        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "0000-0000-00"
'                        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3
'                    Else
'                        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 2
'                    End If
'                Next k
'                strValue = strAmount 'Mid(strAmount, 1, Len(strAmount) - 1)
'                ColCnt = ColCnt + 1
'                strRange = EXCEL_RANGE(ColCnt, RowCnt)
'                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = strValue
'                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
'                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
'                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
'                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
'                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "#,##0.00"
'
'                UpdateProgress picProgressBar, i / rs.RecordCount
'
'                rs.MoveNext
'            Wend
'        End If
'        rs.Close
'
'SAVING:
'        On Error GoTo err_saving:
'        If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
'        xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName
'
'        xlsApp.Visible = True
'
'        picProgressBar.BackColor = &HFFFFFF
'        picProgress.Visible = False
'        picPrint.Enabled = True
        
        
    Case Else: Exit Sub
End Select
Exit Sub
ErrorHandler:
'picProgressBar.BackColor = &HFFFFFF
'picProgress.Visible = False
'picPrint.Enabled = True
Exit Sub

'Exit Sub
'err_saving:
'MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
'GoTo SAVING:
End Sub


Private Sub cmdOKSearch_Click()
If lstResult.ListIndex = -1 Then Exit Sub
If cmbPeriod.ListIndex = -1 Then Exit Sub

    BROWSER cmbPeriod.ItemData(cmbPeriod.ListIndex), "is_LOAD"
    cmdCancelSearch_Click

End Sub

Private Sub cmdOKTaxAlpha_Click()
If RETURNTEXTVALUE(txtTaxAlphaYear) <= 0 Then Exit Sub


MainForm.CommonDialog1.CancelError = True
On Error GoTo ErrorHandler
MainForm.CommonDialog1.DialogTitle = "Save"
MainForm.CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
MainForm.CommonDialog1.ShowSave
Filename = Trim(MainForm.CommonDialog1.Filename)

WorkbookName = CStr(Filename)

On Error GoTo PG:

picTaxWithHeldAlpha.Visible = False
picProgressBar.BackColor = &HFFFFFF
picProgress.ZOrder 0
picProgress.Visible = True
i = 0

ConnOmega.Execute "DELETE FROM tbl_Personnel_Tax_Alphalist WHERE (LogInName = '" & gbl_UserName & "')"

s = "SELECT tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
    " tbl_Personnel_Information.TIN " & _
    " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
    " tbl_Personnel_IDNumber ON tbl_Personnel_Compensation.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK LEFT OUTER JOIN " & _
    " tbl_Personnel_Compensation_Period ON tbl_Personnel_Compensation.Period = tbl_Personnel_Compensation_Period.PK " & _
    " Where (Year(tbl_Personnel_Compensation_Period.DateTo) = " & RETURNTEXTVALUE(txtTaxAlphaYear) & ") " & _
    " GROUP BY tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName, " & _
    " tbl_Personnel_Information.TIN " & _
    " ORDER BY tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    DoEvents
    i = i + 1
    iDivision = 0
    t = "SELECT TOP 1 tbl_Personnel_Action.Division " & _
        " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
        " tbl_Personnel_IDNumber ON tbl_Personnel_Action.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
        " WHERE (YEAR(tbl_Personnel_Action.EffectivityDate) <= " & RETURNTEXTVALUE(txtTaxAlphaYear) & ") " & _
        " AND (tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName = '" & FORMATSQL(rs!EmployeeName) & "') " & _
        " ORDER BY tbl_Personnel_Action.EffectivityDate DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        iDivision = rt!Division
    End If
    rt.Close
    
    sTaxStatus = ""
    t = "SELECT TOP 1 tbl_Personnel_TaxStatus.TaxStatus " & _
        " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
        " tbl_Personnel_IDNumber ON tbl_Personnel_Action.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_TaxStatus ON tbl_Personnel_Action.TaxStatus = tbl_Personnel_TaxStatus.PK " & _
        " WHERE (YEAR(tbl_Personnel_Action.EffectivityDate) <= " & RETURNTEXTVALUE(txtTaxAlphaYear) & ") " & _
        " AND (tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName = '" & FORMATSQL(rs!EmployeeName) & "') " & _
        " ORDER BY tbl_Personnel_Action.EffectivityDate DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        sTaxStatus = rt!TaxStatus
    End If
    rt.Close
    
    For j = 1 To 12
        t = "SELECT SUM(tbl_Personnel_Compensation.TotalEarning) AS Gross, " & _
            " SUM(tbl_Personnel_Compensation.SSS) AS SSS, " & _
            " SUM(tbl_Personnel_Compensation.PHIC) AS PHIC, " & _
            " SUM(tbl_Personnel_Compensation.PagIbig) AS PagIbig, " & _
            " SUM(tbl_Personnel_Compensation.WithHeld) AS WithHeld, " & _
            " SUM(tbl_Personnel_Compensation.TotalCola) AS Cola, " & _
            " SUM(tbl_Personnel_Compensation.TotalAllowance) AS Allowance " & _
            " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
            " tbl_Personnel_Compensation_Period ON " & _
            " tbl_Personnel_Compensation.Period = tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_IDNumber ON tbl_Personnel_Compensation.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
            " WHERE (tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName = '" & FORMATSQL(rs!EmployeeName) & "') " & _
            " AND (YEAR(tbl_Personnel_Compensation_Period.DateTo) = " & RETURNTEXTVALUE(txtTaxAlphaYear) & ") " & _
            " AND (MONTH(tbl_Personnel_Compensation_Period.DateTo) = " & j & ")"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            u = "SELECT tbl_Personnel_Tax_Alphalist.* " & _
                " FROM tbl_Personnel_Tax_Alphalist " & _
                " WHERE (LogInName = '" & gbl_UserName & "') " & _
                " AND (EmployeeName = '" & FORMATSQL(rs!EmployeeName) & "')"
            If ru.State = adStateOpen Then ru.Close
            ru.Open u, ConnOmega
            If ru.RecordCount = 0 Then
                ConnOmega.Execute "INSERT INTO tbl_Personnel_Tax_Alphalist " & _
                                  " (LogInName, EmployeeName, Tin, TaxStatus, Division) " & _
                                  " VALUES ('" & gbl_UserName & "', '" & FORMATSQL(rs!EmployeeName) & "', '" & rs!TIN & "', '" & FORMATSQL(CStr(sTaxStatus)) & "', " & iDivision & ")"
            End If
            ru.Close
            
            ConnOmega.Execute "UPDATE tbl_Personnel_Tax_Alphalist " & _
                                      " SET " & "Gross" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!Gross), 0, rt!Gross)) & ", " & _
                                      " " & "Tax" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!WithHeld), 0, rt!WithHeld)) & ", " & _
                                      " " & "SSS" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!SSS), 0, rt!SSS)) & ", " & _
                                      " " & "PHIC" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!PHIC), 0, rt!PHIC)) & ", " & _
                                      " " & "HDMF" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!PagIbig), 0, rt!PagIbig)) & ", " & _
                                      " " & "Cola" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!Cola), 0, rt!Cola)) & ", " & _
                                      " " & "Allow" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!Allowance), 0, rt!Allowance)) & " " & _
                                      " WHERE (LogInName = '" & gbl_UserName & "') " & _
                                      " AND (EmployeeName = '" & FORMATSQL(rs!EmployeeName) & "')"
            
            
        End If
        rt.Close
        
    Next j
    
    UpdateProgress_Caption rs!EmployeeName, picProgressBar, i / rs.RecordCount
    rs.MoveNext
Wend
rs.Close


picProgressBar.BackColor = &HFFFFFF

i = 0: RowCnt = 0: iDivision = 0: iCnt = 0

iWorkSheet = 1
Set xlsApp = CreateObject("Excel.Application")
xlsApp.Visible = False
xlsApp.Workbooks.Add
xlsApp.DisplayAlerts = False
xlsApp.Workbooks(1).Sheets(2).Delete
xlsApp.Workbooks(1).Sheets(2).Delete
xlsApp.Workbooks(1).Sheets(iWorkSheet).Activate
xlsApp.Workbooks(1).Sheets(iWorkSheet).Name = "AlphaList"

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = gbl_CompanyName
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Alpha List for the year " & txtTaxAlphaYear.Text
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(ColCnt).ColumnWidth = 3
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4

ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(ColCnt).ColumnWidth = 1
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 2

ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(ColCnt).ColumnWidth = 30
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 2

ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3

ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3

For k = 1 To 12
    For j = 1 To 7
        ColCnt = ColCnt + 1
        Select Case j
            Case 1: strValue = "": strRange1 = EXCEL_RANGE(ColCnt, RowCnt)
            Case 2: strValue = ""
            Case 3: strValue = IIf(CDbl(k) = 1, "J A N U A R Y", _
                               IIf(CDbl(k) = 2, "F E B R U A R Y", _
                               IIf(CDbl(k) = 3, "M A R C H", _
                               IIf(CDbl(k) = 4, "A P R I L", _
                               IIf(CDbl(k) = 5, "M A Y", _
                               IIf(CDbl(k) = 6, "J U N E", _
                               IIf(CDbl(k) = 7, "J U L Y", _
                               IIf(CDbl(k) = 8, "A U G U S T", _
                               IIf(CDbl(k) = 9, "S E P T E M B E R", _
                               IIf(CDbl(k) = 10, "O C T O B E R", _
                               IIf(CDbl(k) = 11, "N O V E M B E R", _
                               IIf(CDbl(k) = 12, "D E C E M B E R", ""))))))))))))
            Case 4: strValue = ""
            Case 5: strValue = ""
            Case 6: strValue = ""
            Case 7: strValue = "": strRange2 = EXCEL_RANGE(ColCnt, RowCnt)
        End Select
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = strValue
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3
    Next j
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1, strRange2).Select
    xlsApp.Selection.Merge
    If k = 1 Or k = 3 Or k = 5 Or k = 7 Or k = 9 Or k = 11 Then
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).Interior.ColorIndex = 15
    Else
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).Interior.ColorIndex = 28
    End If
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).Font.Color = vbRed
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).HorizontalAlignment = 3
Next k

ColCnt = ColCnt + 1
strRange1 = EXCEL_RANGE(ColCnt, RowCnt)
ColCnt = ColCnt + 6
strRange2 = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1, strRange2).Select
xlsApp.Selection.Merge
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).Value = "TOTAL"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).HorizontalAlignment = 3
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).Interior.ColorIndex = 15
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).Font.Color = vbRed
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).HorizontalAlignment = 3
    
RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "#"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(ColCnt).ColumnWidth = 3
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4

ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(ColCnt).ColumnWidth = 1
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 2

ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Employee Name"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(ColCnt).ColumnWidth = 30
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 2

ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "TIN"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3

ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Tax Status"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3

For k = 1 To 12
    For j = 1 To 7
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        Select Case j
            Case 1: strValue = "Gross": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 10
            Case 2: strValue = "Tax": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 12
            Case 3: strValue = "SSS": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 14
            Case 4: strValue = "PHIC": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 16
            Case 5: strValue = "HDMF": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 17
            Case 6: strValue = "Cola": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 18
            Case 7: strValue = "Allowance": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 19
        End Select
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = strValue
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
    Next j
Next k

For j = 1 To 7
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    Select Case j
        Case 1: strValue = "Gross": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 10
        Case 2: strValue = "Tax": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 12
        Case 3: strValue = "SSS": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 14
        Case 4: strValue = "PHIC": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 16
        Case 5: strValue = "HDMF": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 17
        Case 6: strValue = "Cola": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 18
        Case 7: strValue = "Allowance": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 19
    End Select
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = strValue
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
Next j


s = "SELECT Division, EmployeeName, Tin, TaxStatus, Gross01, Tax01, SSS01, PHIC01, HDMF01, Cola01, Allow01, Gross02, Tax02, SSS02, PHIC02, HDMF02, Cola02, " & _
    " Allow02, Gross03, Tax03, SSS03, PHIC03, HDMF03, Cola03, Allow03, Gross04, Tax04, SSS04, PHIC04, HDMF04, Cola04, Allow04, Gross05, Tax05, " & _
    " SSS05, PHIC05, HDMF05, Cola05, Allow05, Gross06, Tax06, SSS06, PHIC06, HDMF06, Cola06, Allow06, Gross07, Tax07, SSS07, PHIC07, HDMF07, " & _
    " Cola07, Allow07, Gross08, Tax08, SSS08, PHIC08, HDMF08, Cola08, Allow08, Gross09, Tax09, SSS09, PHIC09, HDMF09, Cola09, Allow09, Gross10, " & _
    " Tax10, SSS10, PHIC10, HDMF10, Cola10, Allow10, Gross11, Tax11, SSS11, PHIC11, HDMF11, Cola11, Allow11, Gross12, Tax12, SSS12, PHIC12, " & _
    " HDMF12 , Cola12, Allow12 " & _
    " From tbl_Personnel_Tax_Alphalist " & _
    " WHERE (LogInName = '" & gbl_UserName & "') " & _
    " ORDER BY Division, EmployeeName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    DoEvents
    i = i + 1
    iReset = 0: strGrossTot = "=": staTaxTot = "=": strSSSTot = "=": strPHICTot = "="
    strHDMFTot = "=": strColaTot = "=": strAllowTot = "="
    
    If CDbl(iDivision) <> CDbl(rs!Division) Then
        If iDivision <> 0 Then
            RowCnt = RowCnt + 1
            ColCnt = 0
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
        End If
        iDivision = rs!Division
        iCnt = 0
    End If
    
    iCnt = iCnt + 1
    RowCnt = RowCnt + 1
    ColCnt = 0
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = iCnt
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "."
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
    
    For j = 1 To rs.Fields.Count - 1
        Select Case j
            Case 1
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs.Fields(j).Value
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(ColCnt).ColumnWidth = 30
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Color = vbBlue
            Case 2
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs.Fields(j).Value
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Color = vbBlue
            Case 3
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs.Fields(j).Value
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Color = vbRed
            Case Else
                ColCnt = ColCnt + 1
                iReset = iReset + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "#,##0.00"
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs.Fields(j).Value
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
                Select Case iReset
                    Case 1: strGrossTot = strGrossTot & strRange & "+"
                    Case 2: staTaxTot = staTaxTot & strRange & "+"
                    Case 3: strSSSTot = strSSSTot & strRange & "+"
                    Case 4: strPHICTot = strPHICTot & strRange & "+"
                    Case 5: strHDMFTot = strHDMFTot & strRange & "+"
                    Case 6: strColaTot = strColaTot & strRange & "+"
                    Case 7: strAllowTot = strAllowTot & strRange & "+"
                End Select
                If iReset = 7 Then
                    iReset = 0
                End If
        End Select
    Next j
    
    For j = 1 To 7
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        Select Case j
            Case 1: strValue = Mid(strGrossTot, 1, Len(strGrossTot) - 1)
            Case 2: strValue = Mid(staTaxTot, 1, Len(staTaxTot) - 1)
            Case 3: strValue = Mid(strSSSTot, 1, Len(strSSSTot) - 1)
            Case 4: strValue = Mid(strPHICTot, 1, Len(strPHICTot) - 1)
            Case 5: strValue = Mid(strHDMFTot, 1, Len(strHDMFTot) - 1)
            Case 6: strValue = Mid(strColaTot, 1, Len(strColaTot) - 1)
            Case 7: strValue = Mid(strAllowTot, 1, Len(strAllowTot) - 1)
        End Select
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "#,##0.00"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = strValue
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
    Next j
    
    UpdateProgress_Caption "Generating Excel Report", picProgressBar, i / rs.RecordCount
    rs.MoveNext
Wend
rs.Close

SAVING:
On Error GoTo err_saving:
If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

xlsApp.Visible = True

picProgressBar.BackColor = &HFFFFFF
picProgress.Visible = False
picPrint.Enabled = True

Exit Sub
err_saving:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING:

Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub

Exit Sub
ErrorHandler:
Exit Sub
End Sub

Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
s = "SELECT tbl_Personnel_Compensation.PK, tbl_Personnel_Compensation.EmpPK, " & _
    " tbl_Personnel_Compensation_Period.DateTo " & _
    " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
    " tbl_Personnel_Compensation_Period ON tbl_Personnel_Compensation.Period = tbl_Personnel_Compensation_Period.PK " & _
    " Where (tbl_Personnel_Compensation.Period = 359) " & _
    " ORDER BY tbl_Personnel_Compensation.Period DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    Array1 = Split(GET_EMPLOYEE_INFO(FormatDateTime(rs!DateTo, vbShortDate), rs!EmpPK), ";", -1, 1)
    ConnOmega.Execute "UPDATE tbl_Personnel_Compensation " & _
                      " SET ActionMemo = " & CDbl(Array1(9)) & " " & _
                      " WHERE (PK = " & rs!PK & ")"
    rs.MoveNext
Wend
rs.Close
Screen.MousePointer = vbDefault
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
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "CompensationC", "CompC", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "CompensationC", "CompC", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "CompensationC", "CompC", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "CompensationC", "CompC", ""), "is_END"
    Case vbKeyEscape:   PRESS_ESCAPE
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.Height - Me.Height) / 10
Me.Left = (MainForm.Width - Me.Width) / 5
'WebBrowser1.Navigate (App.Path & "\images\Save_Big.gif")
BROWSER GetSetting(App.EXEName, "CompensationC", "CompC", ""), "is_LOAD"
If Trim(txtName.Text) = "" Then BROWSER GetSetting(App.EXEName, "CompensationC", "CompC", ""), "is_HOME"
TRANSACTIONTYPE = is_REFRESH
LOCKTEXT True
TOOLBARFUNC 1

With lstReportType
    .Clear
    .AddItem "PAYSLIP"
    .AddItem "SIGNATURE LEDGER"
    .AddItem "COMPENSATION SUMMARY (TOP SHEET)"
    .AddItem "COMPENSATION SUMMARY"
    .AddItem "DEDUCTION SUMMARY (TOP SHEET)"
    .AddItem "DEDUCTION SUMMARY"
    .AddItem "SSS LOANS (TOP SHEET)"
    .AddItem "SSS LOANS"
    .AddItem "PAG-IBIG LOANS (TOP SHEET)"
    .AddItem "PAG-IBIG LOANS"
    .AddItem "SSS COLLECTIONS (TOP SHEET)"
    .AddItem "SSS COLLECTIONS"
    .AddItem "PHIC COLLECTIONS (TOP SHEET)"
    .AddItem "PHIC COLLECTIONS"
    .AddItem "PAG-IBIG COLLECTIONS (TOP SHEET)"
    .AddItem "PAG-IBIG COLLECTIONS"
    .AddItem "TAX WITHHELD (Alpha List)" '(TOP SHEET)"
    .AddItem "TAX WITHHELD"
    .AddItem "13th MONTH (TOP SHEET)"
    .AddItem "13th MONTH"
    .AddItem "COLA SUMMARY"
    '.AddItem "ALLOWANCE SUMMARY"
    .AddItem "FOR ATM"
    .ListIndex = 0
End With

With cmbDivision
    .Clear
    .AddItem "CLUB HOUSE"
    .AddItem "MAINTENANCE"
    .ListIndex = 0
End With

With cmbDivisionAlpha
    .Clear
    .AddItem "CLUB HOUSE"
    .AddItem "MAINTENANCE"
    .ListIndex = 0
End With


tmp = SetWindowLong(txtSearchAdd.hwnd, GWL_STYLE, GetWindowLong(txtSearchAdd.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearchPrint.hwnd, GWL_STYLE, GetWindowLong(txtSearchPrint.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picPrint.Visible = True Then Cancel = -1
If picSearch.Visible = True Then Cancel = -1
If picAdd.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub


Private Sub lblAllowance_Change()
lblNetPayTmp.Caption = Format(RETURNLABELVALUE(lblNetPay) + _
                       RETURNLABELVALUE(lblCola) + _
                       RETURNLABELVALUE(lblAllowance), "#,##0.00")
End Sub

Private Sub lblCola_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
'lblNetPay.Caption = Format((RETURNLABELVALUE(lblTotalEarnings) + RETURNLABELVALUE(lblCola)) - RETURNLABELVALUE(lblTotalDeductions), "#,##0.00")
'lblNetPayTmp.Caption = Format(RETURNLABELVALUE(lblNetPay) + RETURNLABELVALUE(lblCola), "#,##0.00")
lblNetPayTmp.Caption = Format(RETURNLABELVALUE(lblNetPay) + _
                       RETURNLABELVALUE(lblCola) + _
                       RETURNLABELVALUE(lblAllowance), "#,##0.00")
End Sub

Private Sub lblNetPay_Change()
lblNetPayTmp.Caption = Format(RETURNLABELVALUE(lblNetPay) + _
                       RETURNLABELVALUE(lblCola) + _
                       RETURNLABELVALUE(lblAllowance), "#,##0.00")
End Sub

Private Sub lblTotalDeductions_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
lblNetPay.Caption = Format((RETURNLABELVALUE(lblTotalEarnings)) - RETURNLABELVALUE(lblTotalDeductions), "#,##0.00")
End Sub

Private Sub lblTotalEarnings_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
lblNetPay.Caption = Format((RETURNLABELVALUE(lblTotalEarnings)) - RETURNLABELVALUE(lblTotalDeductions), "#,##0.00")
'If TRANSACTIONTYPE = is_EDITTING Then
    chkContribution.Value = 0
'End If
End Sub

Private Sub lstReportType_Click()
If lstReportType.ListIndex = -1 Then Exit Sub
LOAD_GROUP_BY lstReportType.ListIndex
cmbDivision_Click
End Sub

Private Sub lstReportType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbDivision.SetFocus
End Sub

Private Sub lstResult_Click()
If lstResult.ListIndex = -1 Then cmbPeriod.Clear: Exit Sub

cmbPeriod.Clear
s = "SELECT tbl_Personnel_Compensation.PK, " & _
    " tbl_Personnel_Compensation_Period.DateFrom, " & _
    " tbl_Personnel_Compensation_Period.DateTo " & _
    " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
    " tbl_Personnel_Compensation_Period ON tbl_Personnel_Compensation.Period = tbl_Personnel_Compensation_Period.PK " & _
    " Where (tbl_Personnel_Compensation.EmpPK = " & lstResult.ItemData(lstResult.ListIndex) & ") " & _
    " ORDER BY tbl_Personnel_Compensation_Period.DateFrom DESC, tbl_Personnel_Compensation_Period.DateTo DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbPeriod.AddItem Format(rs!DateFrom, "mm/dd/yyyy") & " - " & Format(rs!DateTo, "mm/dd/yyyy")
    cmbPeriod.ItemData(cmbPeriod.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If cmbPeriod.ListCount Then cmbPeriod.ListIndex = 0
End Sub

Private Sub lstResult_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbPeriod.SetFocus
End Sub

Private Sub lstResultAdd_Click()
If lstResultAdd.ListIndex = -1 Then Exit Sub
Array1 = Split(FIND_PAYROLL_PERIOD(Date, GET_DIVISION(lstResultAdd.ItemData(lstResultAdd.ListIndex), Date)), ";", -1, 1)
txtFrom.Text = Format(Array1(1), "mm/dd/yyyy")
txtTo.Text = Format(Array1(2), "mm/dd/yyyy")
'txtPeriod.Text = CLng(Array1(0))
End Sub

Private Sub lstResultAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtFrom.SetFocus
End Sub

Private Sub lstResultPrint_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKPrint_Click
End Sub

Private Sub optActive_Click()
If optActive.Value = True Then txtSearchAdd.Text = ""
End Sub

Private Sub optActive_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtSearchAdd.SetFocus
End Sub

Private Sub optInactive_Click()
If optInactive.Value = True Then txtSearchAdd.Text = ""
End Sub

Private Sub optInactive_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtSearchAdd.SetFocus
End Sub

Private Sub Timer1_Timer()
Static x As Integer
x = x + 1
If x = 1 Then
    Label2.Left = 960
ElseIf x = 2 Then
    Label2.Left = 1000
ElseIf x = 3 Then
    Label2.Left = 1060
ElseIf x = 4 Then
    Label2.Left = 1120
ElseIf x = 5 Then
    Label2.Left = 1180
ElseIf x = 6 Then
    Label2.Left = 1240
ElseIf x = 7 Then
    Label2.Left = 1300
ElseIf x = 8 Then
    Label2.Left = 1360
ElseIf x = 9 Then
    Label2.Left = 1420
ElseIf x = 10 Then
    Label2.Left = 1480
ElseIf x = 11 Then
    Label2.Left = 1560
    x = 0
End If
End Sub

Private Sub Timer2_Timer()
Static x As Integer
x = x + 1
If x = 1 Then
    Label31.Left = 960
ElseIf x = 2 Then
    Label31.Left = 1000
ElseIf x = 3 Then
    Label31.Left = 1060
ElseIf x = 4 Then
    Label31.Left = 1120
ElseIf x = 5 Then
    Label31.Left = 1180
ElseIf x = 6 Then
    Label31.Left = 1240
ElseIf x = 7 Then
    Label31.Left = 1300
ElseIf x = 8 Then
    Label31.Left = 1360
ElseIf x = 9 Then
    Label31.Left = 1420
ElseIf x = 10 Then
    Label31.Left = 1480
ElseIf x = 11 Then
    Label31.Left = 1560
    x = 0
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":           PRESS_INSERT
    Case "Edit":          PRESS_F2
    Case "Delete":        PRESS_DELETE
    Case "First"
        Select Case Toolbar1.Buttons(7).Caption
            Case "Save":  PRESS_F5
            Case "First": BROWSER GetSetting(App.EXEName, "CompensationC", "CompC", ""), "is_HOME"
        End Select
    Case "Back"
        Select Case Toolbar1.Buttons(9).Caption
            Case "Undo":  PRESS_ESCAPE
            Case "Back":  BROWSER GetSetting(App.EXEName, "CompensationC", "CompC", ""), "is_PAGEUP"
        End Select
    Case "Next":          BROWSER GetSetting(App.EXEName, "CompensationC", "CompC", ""), "is_PAGEDOWN"
    Case "Last":          BROWSER GetSetting(App.EXEName, "CompensationC", "CompC", ""), "is_END"
    Case "Find":          PRESS_F6
    Case "Print":         PRESS_F9
    Case "Close":         PRESS_ESCAPE
End Select
End Sub

Private Sub txtAdjustment_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
CALCULATE_EARNING
End Sub

Private Sub txtAdjustment_GotFocus()
HTEXT txtAdjustment
End Sub

Private Sub txtAdjustment_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtRegOT.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtSL.SetFocus
End If
End Sub

Private Sub txtAdjustment_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII_NEG(KeyAscii)
End Sub

Private Sub txtAdjustment_LostFocus()
'txtAdjustment.Text = Format(returntextvalue(txtAdjustment), "##,##0.00")
End Sub

Private Sub txtAdjustmentAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtRegOTAdd.SetFocus
If KeyCode = vbKeyDown Then txtRegOTAdd.SetFocus
If KeyCode = vbKeyUp Then txtSLAdd.SetFocus
End Sub

Private Sub txtAdjustmentAdd_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtAdvances_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
CALCULATE_DEDUCTIONS
End Sub

Private Sub txtAdvances_GotFocus()
HTEXT txtAdvances
End Sub

Private Sub txtAdvances_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtShortages.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtAROthers.SetFocus
End If
End Sub

Private Sub txtAdvances_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtAdvances_LostFocus()
txtAdvances.Text = Format(RETURNTEXTVALUE(txtAdvances), "##,##0.00")
End Sub

Private Sub txtAllow_Change()
lblAllowance.Caption = Format(RETURNTEXTVALUE(txtAllow), "#,##0.00")
End Sub

Private Sub txtAllowPerHour_Change()
txtAllow.Text = Format(RETURNTEXTVALUE(txtNoHours) * RETURNTEXTVALUE(txtAllowPerHour), "#,##0.00")
End Sub

Private Sub txtAmountEarned_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
CALCULATE_EARNING
End Sub

Private Sub txtAROthers_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
CALCULATE_DEDUCTIONS
End Sub

Private Sub txtAROthers_GotFocus()
HTEXT txtAROthers
End Sub

Private Sub txtAROthers_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtAdvances.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtMortuary.SetFocus
End If
End Sub

Private Sub txtAROthers_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtAROthers_LostFocus()
txtAROthers.Text = Format(RETURNTEXTVALUE(txtAROthers), "##,##0.00")
End Sub

Private Sub txtColaHrs_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
lblCola.Caption = Format(RETURNTEXTVALUE(txtColaPerHour) * RETURNTEXTVALUE(txtColaHrs), "#,##0.00")
End Sub

Private Sub txtColaHrs_GotFocus()
HTEXT txtColaHrs
End Sub

Private Sub txtColaHrs_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then txtSH.SetFocus
If KeyCode = vbKeyUp Then txtNoHours.SetFocus
End Sub

Private Sub txtColaHrsAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtSHAdd.SetFocus
If KeyCode = vbKeyDown Then txtSHAdd.SetFocus
If KeyCode = vbKeyUp Then txtNoHoursAdd.SetFocus
End Sub

Private Sub txtColaHrsAdd_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtColaPerHour_Change()
'lblCola.Caption = Format(RETURNTEXTVALUE(txtNoHours) * RETURNTEXTVALUE(txtColaPerHour), "#,##0.00")
lblCola.Caption = Format(RETURNTEXTVALUE(txtColaPerHour) * RETURNTEXTVALUE(txtColaHrs), "#,##0.00")
End Sub

Private Sub txtDept_GotFocus()
HTEXT txtDept
End Sub

Private Sub txtDept_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtPost.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtName.SetFocus
End If
End Sub

Private Sub txtDivision_Change()
If CInt(txtDivision.Text) = 1 Then
    txtDivName.Text = "CLUB HOUSE"
ElseIf CInt(txtDivision.Text) = 2 Then
    txtDivName.Text = "MAINTENANCE"
Else
    txtDivName.Text = ""
End If
End Sub

Private Sub txtDivName_GotFocus()
HTEXT txtDivName
End Sub

Private Sub txtDivName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtDept.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtPayrollPeriod.SetFocus
End If
End Sub

Private Sub txtFrom_GotFocus()
HTEXT txtFrom
End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtTo.SetFocus
If KeyCode = vbKeyUp Then lstResultAdd.SetFocus
End Sub

Private Sub txtFrom_LostFocus()
If IsDate(txtFrom.Text) Then
    txtFrom.Text = Format(FormatDateTime(Trim(txtFrom.Text), vbShortDate), "mm/dd/yyyy")
End If
End Sub

Private Sub txtID_GotFocus()
HTEXT txtID
End Sub

Private Sub txtID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtName.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtWithHeld.SetFocus
End If
End Sub

Private Sub txtLH_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
txtLHAmount.Text = Format((RETURNTEXTVALUE(txtRatePerHour) * RETURNTEXTVALUE(txtLH)), "##,##0.00")
txtTotalForAllowance.Text = RETURNTEXTVALUE(txtNoHours) + _
                            RETURNTEXTVALUE(txtSH) + _
                            RETURNTEXTVALUE(txtLH) + _
                            RETURNTEXTVALUE(txtSL)
End Sub

Private Sub txtLH_GotFocus()
HTEXT txtLH
End Sub

Private Sub txtLH_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtSL.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtSH.SetFocus
End If
End Sub

Private Sub txtLH_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtLH_LostFocus()
'txtLH.Text = Format(returntextvalue(txtLH), "##,##0.00")
End Sub

Private Sub txtLHAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtSLAdd.SetFocus
If KeyCode = vbKeyDown Then txtSLAdd.SetFocus
If KeyCode = vbKeyUp Then txtSHAdd.SetFocus
End Sub

Private Sub txtLHAdd_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtLHAmount_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
CALCULATE_EARNING
End Sub

Private Sub txtLHOT_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
txtLHOTAmount.Text = Format(RETURNTEXTVALUE(txtRatePerHour) * RETURNTEXTVALUE(txtLHOT), "##,##0.00")
End Sub

Private Sub txtLHOT_GotFocus()
HTEXT txtLHOT
End Sub

Private Sub txtLHOT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtSSSLoan.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtSHOT.SetFocus
End If
End Sub

Private Sub txtLHOT_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtLHOT_LostFocus()
'txtLHOT.Text = Format(returntextvalue(txtLHOT), "##,##0.00")
End Sub

Private Sub txtLHOTAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAdd_Click
If KeyCode = vbKeyUp Then txtSHOTAdd.SetFocus
End Sub

Private Sub txtLHOTAdd_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtLHOTAmount_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
CALCULATE_EARNING
End Sub

Private Sub txtMortuary_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
CALCULATE_DEDUCTIONS
End Sub

Private Sub txtMortuary_GotFocus()
HTEXT txtMortuary
End Sub

Private Sub txtMortuary_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtAROthers.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtPagIbigLoan.SetFocus
End If
End Sub

Private Sub txtMortuary_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtMortuary_LostFocus()
txtMortuary.Text = Format(RETURNTEXTVALUE(txtMortuary), "##,##0.00")
End Sub

Private Sub txtName_GotFocus()
HTEXT txtName
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtPayrollPeriod.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtWithHeld.SetFocus
End If
End Sub

Private Sub txtNoHours_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
txtAmountEarned.Text = Format(RETURNTEXTVALUE(txtNoHours) * RETURNTEXTVALUE(txtRatePerHour), "#,##0.00")
'lblCola.Caption = Format(RETURNTEXTVALUE(txtNoHours) * RETURNTEXTVALUE(txtColaPerHour), "#,##0.00")
'txtAllow.Text = Format(RETURNTEXTVALUE(txtNoHours) * RETURNTEXTVALUE(txtAllowPerHour), "#,##0.00")
txtTotalForAllowance.Text = RETURNTEXTVALUE(txtNoHours) + _
                            RETURNTEXTVALUE(txtSH) + _
                            RETURNTEXTVALUE(txtLH) + _
                            RETURNTEXTVALUE(txtSL)
End Sub

Private Sub txtNoHours_GotFocus()
HTEXT txtNoHours
End Sub

Private Sub txtNoHours_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
'    txtSH.SetFocus
    txtColaHrs.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtPost.SetFocus
End If
End Sub

Private Sub txtNoHours_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtNoHours_LostFocus()
'txtNoHours.Text = Format(returntextvalue(txtNoHours), "##,##0.00")
End Sub

Private Sub txtNoHoursAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtColaHrsAdd.SetFocus
If KeyCode = vbKeyDown Then txtColaHrsAdd.SetFocus
If KeyCode = vbKeyUp Then txtTo.SetFocus
End Sub

Private Sub txtNoHoursAdd_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtOthers_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
CALCULATE_DEDUCTIONS
End Sub

Private Sub txtOthers_GotFocus()
HTEXT txtOthers
End Sub

Private Sub txtOthers_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtSSS.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtUniform.SetFocus
End If
End Sub

Private Sub txtOthers_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtOthers_LostFocus()
txtOthers.Text = Format(RETURNTEXTVALUE(txtOthers), "##,##0.00")
End Sub

Private Sub txtPagIbig_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
CALCULATE_DEDUCTIONS
End Sub

Private Sub txtPagIbig_GotFocus()
HTEXT txtPagIbig
End Sub

Private Sub txtPagIbig_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtWithHeld.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtPHIC.SetFocus
End If
End Sub

Private Sub txtPagIbigLoan_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
CALCULATE_DEDUCTIONS
End Sub

Private Sub txtPagIbigLoan_GotFocus()
HTEXT txtPagIbigLoan
End Sub

Private Sub txtPagIbigLoan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtMortuary.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtSSSLoan.SetFocus
End If
End Sub

Private Sub txtPayrollPeriod_GotFocus()
HTEXT txtPayrollPeriod
End Sub

Private Sub txtPayrollPeriod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtDivName.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtName.SetFocus
End If
End Sub

Private Sub txtPHIC_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
CALCULATE_DEDUCTIONS
End Sub

Private Sub txtPHIC_GotFocus()
HTEXT txtPHIC
End Sub

Private Sub txtPHIC_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtPagIbig.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtSSS.SetFocus
End If
End Sub

Private Sub txtPost_GotFocus()
HTEXT txtPost
End Sub

Private Sub txtPost_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtNoHours.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtDept.SetFocus
End If
End Sub

Private Sub txtRDOT_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
'MsgBox RESTDAY_RATE
txtRDOTAmount.Text = Format(CDbl(RETURNTEXTVALUE(txtRatePerHour) * RETURNTEXTVALUE(txtRDOT)) * RESTDAY_RATE, "##,##0.00")
End Sub

Private Sub txtRDOT_GotFocus()
HTEXT txtRDOT
End Sub

Private Sub txtRDOT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtSHOT.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtRegOT.SetFocus
End If
End Sub

Private Sub txtRDOT_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtRDOT_LostFocus()
'txtRDOT.Text = Format(returntextvalue(txtRDOT), "##,##0.00")
End Sub

Private Sub txtRDOTAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtSHOTAdd.SetFocus
If KeyCode = vbKeyDown Then txtSHOTAdd.SetFocus
If KeyCode = vbKeyUp Then txtRegOTAdd.SetFocus
End Sub

Private Sub txtRDOTAdd_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtRDOTAmount_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
CALCULATE_EARNING
End Sub

Private Sub txtRegOT_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
txtRegOTAmount.Text = Format(CDbl(RETURNTEXTVALUE(txtRatePerHour) * RETURNTEXTVALUE(txtRegOT)) * OVERTIME_RATE, "##,##0.00")
End Sub

Private Sub txtRegOT_GotFocus()
HTEXT txtRegOT
End Sub

Private Sub txtRegOT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtRDOT.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtAdjustment.SetFocus
End If
End Sub

Private Sub txtRegOT_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtRegOT_LostFocus()
'txtRegOT.Text = Format(returntextvalue(txtRegOT), "##,##0.00")
End Sub

Private Sub txtRegOTAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtRDOTAdd.SetFocus
If KeyCode = vbKeyDown Then txtRDOTAdd.SetFocus
If KeyCode = vbKeyUp Then txtAdjustmentAdd.SetFocus
End Sub

Private Sub txtRegOTAdd_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtRegOTAmount_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
CALCULATE_EARNING
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstResult.Clear: cmbPeriod.Clear: Exit Sub

lstResult.Clear: cmbPeriod.Clear
's = "SELECT tbl_PersonnelPayroll.EmpPK, " & _
    " tbl_PersonnelProfile.IDNumber, " & _
    " tbl_PersonnelProfile.LName + ',  ' + tbl_PersonnelProfile.FName + '  ' + tbl_PersonnelProfile.MName AS EmpName " & _
    " FROM tbl_PersonnelPayroll LEFT OUTER JOIN " & _
    " tbl_PersonnelProfile ON tbl_PersonnelPayroll.EmpPK = tbl_PersonnelProfile.PK " & _
    " GROUP BY tbl_PersonnelPayroll.EmpPK, tbl_PersonnelProfile.IDNumber, " & _
    " tbl_PersonnelProfile.LName + ',  ' + tbl_PersonnelProfile.FName + '  ' + tbl_PersonnelProfile.MName, " & _
    " tbl_PersonnelProfile.LName " & _
    " HAVING (tbl_PersonnelProfile.LName LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
    " ORDER BY tbl_PersonnelProfile.LName + ',  ' + tbl_PersonnelProfile.FName + '  ' + tbl_PersonnelProfile.MName"
s = "sp_Personnel_Compensation_Search('" & FORMATSQL(Trim(txtSearch.Text)) & "%')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResult.AddItem rs!IDNumber & " - " & rs!EmployeeName
    lstResult.ItemData(lstResult.NewIndex) = rs!EmpPK
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
End Sub

Private Sub txtSearchAdd_Change()
If Trim(txtSearchAdd.Text) = "" Then lstResultAdd.Clear: Exit Sub

lstResultAdd.Clear
's = "SELECT PK, IDNumber, LName + ',  ' + " & _
    " FName + '  ' + MName AS Name" & _
    " From tbl_PersonnelProfile " & _
    " WHERE (LName LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%') " & _
    " ORDER BY LName + ',  ' + FName + '  ' + MName"
's = "sp_Personnel_Action_Search_Search('" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%')"
If optActive.Value = True Then
    iStatus = 1
End If
If optInactive.Value = True Then
    iStatus = 2
End If
If optInactive.Value = False And optActive.Value = False Then MsgBox "Please Select Status!                     ", vbCritical, "Error...": iStatus = 0: Exit Sub
s = "SELECT tbl_Personnel_IDNumber.PK, tbl_Personnel_IDNumber.IDNumber, " & _
    " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName " & _
    " FROM tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
    " WHERE (tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%') " & _
    " AND ((SELECT TOP 1 tbl_Personnel_EmploymentStatus.Active " & _
    " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
    " tbl_Personnel_EmploymentStatus ON tbl_Personnel_Action.EmpStatus = tbl_Personnel_EmploymentStatus.PK " & _
    " WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_IDNumber.PK) " & _
    " AND (tbl_Personnel_Action.EffectivityDate <= '" & FormatDateTime(Date, vbShortDate) & "') " & _
    " ORDER BY tbl_Personnel_Action.EffectivityDate DESC) = " & iStatus & ") " & _
    " ORDER BY tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
With lstResultAdd
    While Not rs.EOF
        .AddItem rs!IDNumber & " - " & rs!EmployeeName
        .ItemData(.NewIndex) = rs!PK
        rs.MoveNext
    Wend
    If .ListCount Then .ListIndex = 0
End With
rs.Close
End Sub

Private Sub txtSearchAdd_GotFocus()
HTEXT txtSearchAdd
End Sub

Private Sub txtSearchAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultAdd.SetFocus
End Sub

Private Sub txtSearchPrint_Change()
If Trim(txtSearchPrint.Text) = "" Then lstResultPrint.Clear:  Exit Sub

lstResultPrint.Clear
's = "SELECT tbl_Personnel_Compensation.EmpPK, " & _
    " tbl_PersonnelProfile.IDNumber, " & _
    " tbl_PersonnelProfile.LName + ',  ' + tbl_PersonnelProfile.FName + '  ' + tbl_PersonnelProfile.MName AS EmpName " & _
    " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
    " tbl_PersonnelProfile ON tbl_Personnel_Compensation.EmpPK = tbl_PersonnelProfile.PK " & _
    " WHERE (tbl_Personnel_Compensation.Division = " & cmbDivision.ListIndex + 1 & ") " & _
    " AND (tbl_Personnel_Compensation.Period = " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ") " & _
    " AND (tbl_PersonnelProfile.LName LIKE '" & FORMATSQL(Trim(txtSearchPrint.Text)) & "%') " & _
    " ORDER BY tbl_PersonnelProfile.LName + ',  ' + tbl_PersonnelProfile.FName + '  ' + tbl_PersonnelProfile.MName"
s = ""
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultPrint.AddItem rs!IDNumber & " - " & rs!EmpName
    lstResultPrint.ItemData(lstResultPrint.NewIndex) = rs!EmpPK
    rs.MoveNext
Wend
rs.Close
If lstResultPrint.ListCount Then lstResultPrint.ListIndex = 0
End Sub

Private Sub txtSearchPrint_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultPrint.SetFocus
End Sub

Private Sub txtSH_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub

'Daily = Format(RETURNTEXTVALUE(txtRatePerHour) * RETURNTEXTVALUE(txtSH), "##,##0.00")
'Percent = Format(CDbl(Daily) * 0.3, "##,##0.00")
'txtSHAmount.Text = CDbl(Percent) 'CDbl(Daily) + CDbl(Percent)

txtSHAmount.Text = Format(RETURNTEXTVALUE(txtSH) * RETURNTEXTVALUE(txtRatePerHour), "#,##0.00")
txtTotalForAllowance.Text = RETURNTEXTVALUE(txtNoHours) + _
                            RETURNTEXTVALUE(txtSH) + _
                            RETURNTEXTVALUE(txtLH) + _
                            RETURNTEXTVALUE(txtSL)
End Sub

Private Sub txtSH_GotFocus()
HTEXT txtSH
End Sub

Private Sub txtSH_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtLH.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtColaHrs.SetFocus
    'txtNoHours.SetFocus
End If
End Sub

Private Sub txtSH_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtSH_LostFocus()
'txtSH.Text = Format(returntextvalue(txtSH), "##,##0.00")
End Sub

Private Sub txtSHAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtLHAdd.SetFocus
If KeyCode = vbKeyDown Then txtLHAdd.SetFocus
If KeyCode = vbKeyUp Then txtColaHrsAdd.SetFocus
End Sub

Private Sub txtSHAdd_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtSHAmount_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
CALCULATE_EARNING
End Sub

Private Sub txtShortages_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
CALCULATE_DEDUCTIONS
End Sub

Private Sub txtShortages_GotFocus()
HTEXT txtShortages
End Sub

Private Sub txtShortages_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtUniform.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtAdvances.SetFocus
End If
End Sub

Private Sub txtShortages_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtShortages_LostFocus()
txtShortages.Text = Format(RETURNTEXTVALUE(txtShortages), "##,##0.00")
End Sub

Private Sub txtSHOT_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
Daily = Format(RETURNTEXTVALUE(txtRatePerHour) * RETURNTEXTVALUE(txtSHOT), "##,##0.00")
Percent = Format(CDbl(Daily) * 0.3, "##,##0.00")
'txtSHAmount.Text = CDbl(Percent) 'CDbl(Daily) + CDbl(Percent)
txtSHOTAmount.Text = CDbl(Percent) 'Format(RETURNTEXTVALUE(txtRatePerHour) * RETURNTEXTVALUE(txtSHOT), "##,##0.00")
End Sub

Private Sub txtSHOT_GotFocus()
HTEXT txtSHOT
End Sub

Private Sub txtSHOT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtLHOT.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtRDOT.SetFocus
End If
End Sub

Private Sub txtSHOT_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtSHOT_LostFocus()
'txtSHOT.Text = Format(returntextvalue(txtSHOT), "##,##0.00")
End Sub

Private Sub txtSHOTAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtLHOTAdd.SetFocus
If KeyCode = vbKeyDown Then txtLHOTAdd.SetFocus
If KeyCode = vbKeyUp Then txtRDOTAdd.SetFocus
End Sub

Private Sub txtSHOTAdd_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtSHOTAmount_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
CALCULATE_EARNING
End Sub

Private Sub txtSL_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
txtSLAmount.Text = Format(RETURNTEXTVALUE(txtRatePerHour) * RETURNTEXTVALUE(txtSL), "##,##0.00")
txtTotalForAllowance.Text = RETURNTEXTVALUE(txtNoHours) + _
                            RETURNTEXTVALUE(txtSH) + _
                            RETURNTEXTVALUE(txtLH) + _
                            RETURNTEXTVALUE(txtSL)
End Sub

Private Sub txtSL_GotFocus()
HTEXT txtSL
End Sub

Private Sub txtSL_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtAdjustment.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtLH.SetFocus
End If
End Sub

Private Sub txtSL_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtSL_LostFocus()
'txtSL.Text = Format(returntextvalue(txtSL), "##,##0.00")
End Sub

Private Sub txtSLAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtAdjustmentAdd.SetFocus
If KeyCode = vbKeyDown Then txtAdjustmentAdd.SetFocus
If KeyCode = vbKeyUp Then txtLHAdd.SetFocus
End Sub

Private Sub txtSLAdd_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtSLAmount_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
CALCULATE_EARNING
End Sub

Private Sub txtSSS_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
CALCULATE_DEDUCTIONS
End Sub

Private Sub txtSSS_GotFocus()
HTEXT txtSSS
End Sub

Private Sub txtSSS_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtPHIC.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtOthers.SetFocus
End If
End Sub

Private Sub txtSSSLoan_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
CALCULATE_DEDUCTIONS
End Sub

Private Sub txtSSSLoan_GotFocus()
HTEXT txtSSSLoan
End Sub

Private Sub txtSSSLoan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtPagIbigLoan.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtLHOT.SetFocus
End If
End Sub

Private Sub txtTaxAlphaYear_GotFocus()
HTEXT txtTaxAlphaYear
End Sub

Private Sub txtTaxAlphaYear_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKTaxAlpha_Click
End Sub

Private Sub txtTaxAlphaYear_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtTo_GotFocus()
HTEXT txtTo
End Sub

Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtNoHoursAdd.SetFocus
If KeyCode = vbKeyUp Then txtFrom.SetFocus
End Sub

Private Sub txtTo_LostFocus()
If IsDate(txtTo.Text) Then
    txtTo.Text = Format(FormatDateTime(Trim(txtTo.Text), vbShortDate), "mm/dd/yyyy")
End If
End Sub

Private Sub txtTotalForAllowance_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
txtAllow.Text = Format(RETURNTEXTVALUE(txtTotalForAllowance) * RETURNTEXTVALUE(txtAllowPerHour), "#,##0.00")
End Sub

Private Sub txtUniform_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
CALCULATE_DEDUCTIONS
End Sub

Private Sub txtUniform_GotFocus()
HTEXT txtUniform
End Sub

Private Sub txtUniform_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtOthers.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtShortages.SetFocus
End If
End Sub

Private Sub txtUniform_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtUniform_LostFocus()
txtUniform.Text = Format(RETURNTEXTVALUE(txtUniform), "##,##0.00")
End Sub

Private Sub txtWithHeld_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
CALCULATE_DEDUCTIONS
End Sub

Private Sub txtWithHeld_GotFocus()
HTEXT txtWithHeld
End Sub

Private Sub txtWithHeld_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtName.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtPagIbig.SetFocus
End If
End Sub




