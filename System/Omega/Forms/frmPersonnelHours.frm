VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelHours 
   Appearance      =   0  'Flat
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   11490
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picBatchPosting 
      Height          =   2295
      Left            =   3720
      TabIndex        =   68
      Top             =   2040
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4048
      BackColor       =   15396057
      Begin VB.TextBox txtPayrollDatePostUnpost 
         Height          =   315
         Left            =   1080
         TabIndex        =   74
         Top             =   1200
         Width           =   2790
      End
      Begin VB.ComboBox cmbPostUnpost 
         Height          =   315
         ItemData        =   "frmPersonnelHours.frx":0000
         Left            =   120
         List            =   "frmPersonnelHours.frx":0002
         TabIndex        =   73
         Top             =   480
         Width           =   3735
      End
      Begin VB.ComboBox cmbDivisionBatchPost 
         Height          =   315
         ItemData        =   "frmPersonnelHours.frx":0004
         Left            =   120
         List            =   "frmPersonnelHours.frx":0006
         TabIndex        =   72
         Top             =   840
         Width           =   3735
      End
      Begin VB.CommandButton cmdCancelPostUnpost 
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
         Picture         =   "frmPersonnelHours.frx":0008
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   1680
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKPostUnpost 
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
         Picture         =   "frmPersonnelHours.frx":0764
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   1680
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar3 
         Height          =   345
         Left            =   45
         TabIndex        =   71
         Top             =   45
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   609
         Caption         =   "Post / Unpost"
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
         Icon            =   "frmPersonnelHours.frx":0DD6
         ShadowVisible   =   0   'False
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Payroll V2 LoanKey"
      Height          =   495
      Left            =   9240
      TabIndex        =   83
      Top             =   5280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin RPVGCC.b8Container picAdd 
      Height          =   4455
      Left            =   3720
      TabIndex        =   24
      Top             =   600
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   7858
      BackColor       =   15396057
      Begin VB.TextBox txtPayrollDateAdd 
         Height          =   315
         Left            =   1080
         TabIndex        =   65
         Top             =   840
         Width           =   2790
      End
      Begin VB.ComboBox cmbDivisionAdd 
         Height          =   315
         ItemData        =   "frmPersonnelHours.frx":1370
         Left            =   120
         List            =   "frmPersonnelHours.frx":1372
         TabIndex        =   33
         Top             =   480
         Width           =   3735
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
         Picture         =   "frmPersonnelHours.frx":1374
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3840
         Width           =   1560
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
         Picture         =   "frmPersonnelHours.frx":19E6
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3840
         Width           =   1560
      End
      Begin VB.TextBox txtSearchAdd 
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Width           =   3735
      End
      Begin VB.ListBox lstResultAdd 
         Height          =   2205
         Left            =   120
         TabIndex        =   27
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox txtFrom 
         Height          =   315
         Left            =   360
         TabIndex        =   26
         Top             =   840
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtTo 
         Height          =   315
         Left            =   600
         TabIndex        =   25
         Top             =   840
         Visible         =   0   'False
         Width           =   150
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   31
         Top             =   45
         Width           =   3885
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
         Icon            =   "frmPersonnelHours.frx":2142
         ShadowVisible   =   0   'False
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Late/Undertime/Abs"
      Height          =   495
      Left            =   6000
      TabIndex        =   82
      Top             =   5280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "New Action Memo"
      Height          =   495
      Left            =   0
      TabIndex        =   81
      Top             =   5280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Loans to SL"
      Height          =   495
      Left            =   4560
      TabIndex        =   80
      Top             =   5280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Loans Ctrl/Type"
      Height          =   495
      Left            =   3120
      TabIndex        =   79
      Top             =   5280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Payroll Period"
      Height          =   495
      Left            =   1680
      TabIndex        =   78
      Top             =   5280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin RPVGCC.b8Container picSLRegularHours 
      Height          =   855
      Left            =   480
      TabIndex        =   35
      Top             =   2160
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.ComboBox cmRegularHours 
         Height          =   315
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtRegValue 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3600
         TabIndex        =   40
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtEarningRegKey 
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   39
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtEarningRegKey1 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   38
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox cmRegularHours1 
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtRegValue1 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Value"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   42
         Top             =   120
         Width           =   1095
      End
   End
   Begin RPVGCC.b8Container picProgressBar 
      Height          =   975
      Left            =   1200
      TabIndex        =   76
      Top             =   2520
      Visible         =   0   'False
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1720
      BackColor       =   15396057
      Begin VB.PictureBox picProgress 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   8835
         TabIndex        =   77
         Top             =   120
         Width           =   8895
      End
   End
   Begin RPVGCC.b8Container picSLOvertimeHours 
      Height          =   855
      Left            =   5760
      TabIndex        =   44
      Top             =   2160
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox txtOTValue1 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   50
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox cmbOvertimeHours1 
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   49
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtEarningOTKey1 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtEarningOTKey 
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtOTValue 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3600
         TabIndex        =   46
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cmbOvertimeHours 
         Height          =   315
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Value"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   52
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   120
         Width           =   1095
      End
   End
   Begin RPVGCC.b8Container picSearch 
      Height          =   4455
      Left            =   3240
      TabIndex        =   53
      Top             =   480
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   7858
      BackColor       =   15396057
      Begin VB.ComboBox cmbPayrollPeriodSearch 
         Height          =   315
         ItemData        =   "frmPersonnelHours.frx":26DC
         Left            =   1440
         List            =   "frmPersonnelHours.frx":26DE
         TabIndex        =   58
         Top             =   3480
         Width           =   3495
      End
      Begin VB.ListBox lstResultSearch 
         Height          =   2595
         Left            =   120
         TabIndex        =   57
         Top             =   840
         Width           =   4815
      End
      Begin VB.TextBox txtSearchSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   56
         Top             =   480
         Width           =   4815
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
         Left            =   2640
         Picture         =   "frmPersonnelHours.frx":26E0
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   3840
         Width           =   1560
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
         Left            =   960
         Picture         =   "frmPersonnelHours.frx":2E3C
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   3840
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   45
         TabIndex        =   59
         Top             =   45
         Width           =   4965
         _ExtentX        =   8758
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
         Icon            =   "frmPersonnelHours.frx":34AE
         ShadowVisible   =   0   'False
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Period - Division"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   3480
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Payroll V2"
      Height          =   495
      Left            =   7680
      TabIndex        =   21
      Top             =   5280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   0
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   1
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
            NumButtons      =   24
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Add"
               Key             =   "Add"
               ImageKey        =   "IMG1"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Edit"
               Key             =   "Edit"
               ImageKey        =   "IMG2"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Delete"
               Key             =   "Delete"
               ImageKey        =   "IMG3"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "First"
               Key             =   "First"
               ImageKey        =   "IMG4"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Back"
               Key             =   "Back"
               ImageKey        =   "IMG5"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Next"
               Key             =   "Next"
               ImageKey        =   "IMG6"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Last"
               Key             =   "Last"
               ImageKey        =   "IMG7"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Find"
               Key             =   "Find"
               ImageKey        =   "IMG8"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Print"
               Key             =   "Print"
               ImageKey        =   "IMG9"
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   " Post   "
               Key             =   "Post"
               ImageKey        =   "IMG10"
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               Key             =   "Refresh"
               ImageKey        =   "IMG12"
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
               ImageKey        =   "IMG13"
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         MousePointer    =   99
         MouseIcon       =   "frmPersonnelHours.frx":3A48
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   9900
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   2
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmPersonnelHours.frx":3D62
               Top             =   120
               Visible         =   0   'False
               Width           =   1395
            End
         End
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11880
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
            Picture         =   "frmPersonnelHours.frx":4475
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelHours.frx":514F
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelHours.frx":5E29
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelHours.frx":6B03
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelHours.frx":77DD
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelHours.frx":84B7
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelHours.frx":9191
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelHours.frx":9E6B
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelHours.frx":AB45
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelHours.frx":B41F
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelHours.frx":C0F9
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelHours.frx":CDD3
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelHours.frx":DAAD
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelHours.frx":E787
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelHours.frx":F461
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   5775
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   2469
            MinWidth        =   2469
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
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
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   600
      ScaleHeight     =   3855
      ScaleWidth      =   10215
      TabIndex        =   4
      Top             =   1320
      Width           =   10215
      Begin VB.TextBox txtCompType 
         Height          =   315
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox txtCutOffDate 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtControl 
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   62
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picSetFocus 
         Appearance      =   0  'Flat
         BackColor       =   &H00C6B8A4&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   8880
         ScaleHeight     =   375
         ScaleWidth      =   615
         TabIndex        =   61
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdViewPayroll 
         Caption         =   "View Payroll"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8760
         MouseIcon       =   "frmPersonnelHours.frx":1013B
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtAdjustmentRem 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtAdjustment 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "50,000.00"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C6B8A4&
         Caption         =   "Overtime"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   5160
         TabIndex        =   17
         Top             =   1800
         Width           =   5055
         Begin MSComctlLib.ListView lstOvertime 
            Height          =   1695
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   2990
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "EarningKey"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Value"
               Object.Width           =   2646
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C6B8A4&
         Caption         =   "Regular Hours"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   0
         TabIndex        =   15
         Top             =   1800
         Width           =   5055
         Begin MSComctlLib.ListView lstRegularHours 
            Height          =   1695
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   2990
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "EarningKey"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Value"
               Object.Width           =   2646
            EndProperty
         End
      End
      Begin VB.TextBox txtPosition 
         Height          =   315
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtDivision 
         Height          =   315
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox txtDepartment 
         Height          =   315
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox txtPayrollPeriod 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   0
         Width           =   7335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Comp Type"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   67
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Cut-Off Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   64
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Adjustment (Rem)"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   34
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Adjustment (Amt)"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmPersonnelHours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmp As Long

Public TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2
Const is_FINDING = 3

Private TRANS_DETAIL As Long
Const is_DET_REFRESH = 0
Const is_DET_ADDING = 1
Const is_DET_EDITTING = 2

Public locEmployeePK, locPayrollPeroid, locDivision, locActionMemoKey, locPayrollKey    As Long

Dim isFocusReg, iRowReg       As Long
Dim isFocusOT, iRowOT       As Long

Dim Array1, Array2, x, sCtrl, iPK, i, iEmpStatus, dblTotalAmt, dblRatePerHour, dDate, _
locPayrollPeroidTmp, iChkPerfHrs, dPerfectHours, dAbsentLaterUndertime, dNoHours

Dim dPerfectHoursMaster As Double

Dim cnt, locPayrollPeriodTmp

Private Sub PRESS_INSERT()
If picAdd.Visible = True Then Exit Sub
If picSLRegularHours.Visible = True Then Exit Sub
If picSLOvertimeHours.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If picBatchPosting.Visible = True Then Exit Sub
If picProgressBar.Visible = True Then Exit Sub
If CheckIfPaid(Date) = False Then MsgBox "Please contact developer!                   ", vbCritical, "Error...": Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    If AccessRights("Personnel - Hours", "Add") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    picAdd.ZOrder 0
    picMain.Enabled = False
    picToolbar.Enabled = False
    cmbDivisionAdd.Text = ""
    cmbDivisionAdd.ListIndex = -1
    txtFrom.Text = ""
    txtTo.Text = ""
    txtSearchAdd.Text = ""
    lstResultAdd.Clear
    picAdd.Visible = True
    cmbDivisionAdd.SetFocus
Else
    If isFocusReg = 1 Then
        With lstRegularHours.ListItems
            If CDbl(.Item(.Count).Text) <> 0 Then
                'MsgBox "pass over"
                Set x = .Add()
                x.Text = "0"
                x.SubItems(1) = " "
                x.SubItems(2) = " "
            Else
                'MsgBox "pass one"
                .Item(1).Text = "0"
                .Item(1).SubItems(1) = " "
                .Item(1).SubItems(2) = " "
            End If
            iRowReg = .Count
        End With
        lstRegularHours.ListItems(iRowReg).EnsureVisible
        lstRegularHours.ListItems(iRowReg).Selected = True
        picSLRegularHours.ZOrder 0
        txtEarningRegKey.Text = ""
        cmRegularHours.Text = ""
        cmRegularHours.ListIndex = -1
        txtRegValue.Text = ""
        picSLRegularHours.Visible = True
        picMain.Enabled = False
        picToolbar.Enabled = False
        TRANS_DETAIL = is_DET_ADDING
        cmRegularHours.SetFocus
        Exit Sub
    End If
    If isFocusOT = 1 Then
        With lstOvertime.ListItems
            'If .Count > 1 Then
            If CDbl(.Item(.Count).Text) <> 0 Then
                Set x = .Add()
                x.Text = "0"
                x.SubItems(1) = " "
                x.SubItems(2) = " "
            Else
                .Item(1).Text = "0"
                .Item(1).SubItems(1) = " "
                .Item(1).SubItems(2) = " "
            End If
            iRowOT = .Count
        End With
        lstOvertime.ListItems(iRowOT).EnsureVisible
        lstOvertime.ListItems(iRowOT).Selected = True
        picSLOvertimeHours.ZOrder 0
        txtEarningOTKey.Text = ""
        cmbOvertimeHours.Text = ""
        cmbOvertimeHours.ListIndex = -1
        txtOTValue.Text = ""
        picSLOvertimeHours.Visible = True
        picMain.Enabled = False
        picToolbar.Enabled = False
        TRANS_DETAIL = is_DET_ADDING
        cmbOvertimeHours.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub PRESS_F2()
If picAdd.Visible = True Then Exit Sub
If picSLRegularHours.Visible = True Then Exit Sub
If picSLOvertimeHours.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If picBatchPosting.Visible = True Then Exit Sub
If picProgressBar.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    If Trim(Statusbar1.Panels(1).Text) = "" Then Exit Sub
    If AccessRights("Personnel - Hours", "Edit") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If imgPosted.Visible = True Then MsgBox "TRANSACTION ALREADY POSTED!                         ", vbCritical, "Error...": Exit Sub
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_EDITTING
    If isFocusReg = 1 Then lstRegularHours_Click: Exit Sub
    If isFocusOT = 1 Then lstOvertime_Click: Exit Sub
Else
    If isFocusReg = 1 Then
        If Toolbar1.Buttons(3).Enabled = False Then Exit Sub
        With lstRegularHours.ListItems
            txtEarningRegKey.Text = .Item(iRowReg).Text
            cmRegularHours.Text = .Item(iRowReg).SubItems(1)
            txtRegValue.Text = .Item(iRowReg).SubItems(2)
            
            txtEarningRegKey1.Text = .Item(iRowReg).Text
            cmRegularHours1.Text = .Item(iRowReg).SubItems(1)
            txtRegValue1.Text = .Item(iRowReg).SubItems(2)
        End With
        picSLRegularHours.ZOrder 0
        picMain.Enabled = False
        picToolbar.Enabled = False
        picSLRegularHours.Visible = True
        TRANS_DETAIL = is_DET_EDITTING
        cmRegularHours.SetFocus
        Exit Sub
    End If
    If isFocusOT = 1 Then
        With lstOvertime.ListItems
            txtEarningOTKey.Text = .Item(iRowOT).Text
            cmbOvertimeHours.Text = .Item(iRowOT).SubItems(1)
            txtOTValue.Text = .Item(iRowOT).SubItems(2)
            
            txtEarningOTKey1.Text = .Item(iRowOT).Text
            cmbOvertimeHours1.Text = .Item(iRowOT).SubItems(1)
            txtOTValue1.Text = .Item(iRowOT).SubItems(2)
        End With
        picSLOvertimeHours.ZOrder 0
        picMain.Enabled = False
        picToolbar.Enabled = False
        picSLOvertimeHours.Visible = True
        TRANS_DETAIL = is_DET_EDITTING
        cmbOvertimeHours.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub PRESS_DELETE()
If picAdd.Visible = True Then Exit Sub
If picSLRegularHours.Visible = True Then Exit Sub
If picSLOvertimeHours.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If picProgressBar.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    If Trim(Statusbar1.Panels(1).Text) = "" Then Exit Sub
    If AccessRights("Personnel - Hours", "Delete") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If imgPosted.Visible = True Then MsgBox "TRANSACTION ALREADY POSTED!                         ", vbCritical, "Error...": Exit Sub
    If MsgBox("ARE YOU SURE IN DELETING THIS TRANSACTIONT?                  ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    On Error GoTo PG:
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Hours WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    CLEARTEXT
    BROWSER GetSetting(App.EXEName, "PersonnelHours", "PersonnelHours", ""), "is_PAGEDOWN"
    If Trim(txtControl.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelHours", "PersonnelHours", ""), "is_HOME"
Else
    If isFocusReg = 1 Then
        If Toolbar1.Buttons(5).Enabled = False Then Exit Sub
        With lstRegularHours.ListItems
            Dim iEarnKeyTemp, dHoursTemp
            iEarnKeyTemp = .Item(iRowReg).Text
            dHoursTemp = CDbl(.Item(iRowReg).SubItems(2))
            If .Count > 1 Then
                .Remove .Count
                If CDbl(iRowReg) > CDbl(.Count) Then iRowReg = .Count
            Else
                .Item(1).Text = "0"
                .Item(1).SubItems(1) = " "
                .Item(1).SubItems(2) = " "
                iRowReg = 1
            End If
            
            If CheckDeductToReference(iEarnKeyTemp) = 1 Then
                For i = 1 To .Count
                    If CLng(.Item(i).Text) = GetReferenceEarning(iEarnKeyTemp) Then
                        .Item(i).SubItems(2) = CDbl(.Item(i).SubItems(2)) + dHoursTemp
                    End If
                Next i
            End If
            
        End With
        lstRegularHours.ListItems(iRowReg).EnsureVisible
        lstRegularHours.ListItems(iRowReg).Selected = True
        lstRegularHours_Click
        Exit Sub
    End If
    If isFocusOT = 1 Then
        With lstOvertime.ListItems
            If .Count > 1 Then
                .Remove .Count
                If CDbl(iRowOT) > CDbl(.Count) Then iRowOT = .Count
            Else
                .Item(1).Text = "0"
                .Item(1).SubItems(1) = " "
                .Item(1).SubItems(2) = " "
                iRowOT = 1
            End If
        End With
        lstOvertime.ListItems(iRowOT).EnsureVisible
        lstOvertime.ListItems(iRowOT).Selected = True
        Exit Sub
    End If
End If
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If picAdd.Visible = True Then Exit Sub
If picSLRegularHours.Visible = True Then Exit Sub
If picSLOvertimeHours.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If picBatchPosting.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    'Array1 = Split(Trim(txtPayrollPeriod.Text), " - ", -1, 1)
    sCtrl = ""
    s = "SELECT TOP (1) dbo.tbl_Personnel_Hours.Ctrl " & _
        " FROM  dbo.tbl_Personnel_Hours LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Hours.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
        " Where (dbo.tbl_Personnel_Compensation_Period.Type = 1) " & _
        " And (Year(dbo.tbl_Personnel_Compensation_Period.DateTo) = " & Format(FormatDateTime(txtPayrollPeriod.Text, vbShortDate), "yyyy") & ") " & _
        " ORDER BY dbo.tbl_Personnel_Hours.Ctrl DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        sCtrl = Format(CDbl(rs!Ctrl) + 1, "000000000#")
    Else
        sCtrl = Format(FormatDateTime(txtPayrollPeriod.Text, vbShortDate), "yyyy") & "000000"
    End If
    rs.Close
    
    Do
        s = "SELECT tbl_Personnel_Hours.* " & _
            " FROM tbl_Personnel_Hours " & _
            " WHERE (Ctrl = '" & sCtrl & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If
        rs.Close
        sCtrl = Format(CDbl(sCtrl) + 1, "000000000#")
    Loop
    
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Hours " & _
                      " (Ctrl, EmployeeKey, PayrollPeriodKey, ActionMemoKey, Adjustment, AdjustmentRem, PerfectHours, LastModified) " & _
                      " VALUES ('" & sCtrl & "', " & locEmployeePK & ", " & locPayrollPeroid & ", " & locActionMemoKey & ", " & _
                      " " & RETURNTEXTVALUE(txtAdjustment) & ", '" & FORMATSQL(Trim(txtAdjustmentRem.Text)) & "', " & _
                      " " & CDbl(dPerfectHoursMaster) & ", '" & CStr(Now) & " - " & gbl_CompleteName & "')"
    
    s = "SELECT tbl_Personnel_Hours.* " & _
        " FROM tbl_Personnel_Hours " & _
        " WHERE (Ctrl = '" & sCtrl & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        iPK = rs!PK
    End If
    rs.Close
    
End If
If TRANSACTIONTYPE = is_EDITTING Then
    sCtrl = Trim(txtControl.Text)
    iPK = Statusbar1.Panels(1).Text
    
    ConnOmega.Execute "UPDATE tbl_Personnel_Hours " & _
                      " SET Adjustment = " & RETURNTEXTVALUE(txtAdjustment) & ", " & _
                      " AdjustmentRem = '" & FORMATSQL(Trim(txtAdjustmentRem.Text)) & "', " & _
                      " PerfectHours = " & CDbl(dPerfectHoursMaster) & ", " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & iPK & ")"
    
    'MsgBox sCtrl
End If

If CDbl(iPK) <> 0 Then
    With lstRegularHours.ListItems
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Hours_Regular WHERE (MasterKey = " & iPK & ")"
        For i = 1 To .Count
            If CDbl(.Item(i).Text) <> 0 Then
                If CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2))) <> 0 Then
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Hours_Regular " & _
                                      " (MasterKey, EarningKey, NoHours) " & _
                                      " VALUES (" & iPK & ", " & CDbl(.Item(i).Text) & ", " & _
                                      " " & CDbl(.Item(i).SubItems(2)) & ")"
                End If
            End If
        Next i
    End With
    
    With lstOvertime.ListItems
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Hours_Overtime WHERE (MasterKey = " & iPK & ")"
        For i = 1 To .Count
            If CDbl(.Item(i).Text) <> 0 Then
                If CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2))) <> 0 Then
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Hours_Overtime " & _
                                      " (MasterKey, EarningKey, NoHours) " & _
                                      " VALUES (" & iPK & ", " & CDbl(.Item(i).Text) & ", " & _
                                      " " & CDbl(.Item(i).SubItems(2)) & ")"
                End If
            End If
        Next i
    End With
    
End If
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANS_DETAIL = is_DET_REFRESH
BROWSER sCtrl, "is_LOAD"
txtName.SetFocus
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If picAdd.Visible = True Then Exit Sub
If picSLRegularHours.Visible = True Then Exit Sub
If picSLOvertimeHours.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If picBatchPosting.Visible = True Then Exit Sub
If picProgressBar.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
picSearch.ZOrder 0
txtSearchSearch.Text = ""
txtSearchSearch.Text = ""
picMain.Enabled = False
picToolbar.Enabled = False
picSearch.Visible = True
txtSearchSearch.SetFocus
End Sub

Private Sub PRESS_F8()
If picAdd.Visible = True Then Exit Sub
If picSLRegularHours.Visible = True Then Exit Sub
If picSLOvertimeHours.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If picBatchPosting.Visible = True Then Exit Sub
If picProgressBar.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Trim(Statusbar1.Panels(1).Text) = "" Then Exit Sub

PopupMenu MainFormPopupF.mnuPayrollHourPosting, , Toolbar1.Buttons(19).Left, Toolbar1.Buttons(19).Top + Toolbar1.Buttons(19).Height

'On Error GoTo PG:
'If imgPosted.Visible = True Then
'    If AccessRights("Personnel - Hours", "UnPost") = False Then
'        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
'               "ACCESS DENIED!                                      ", vbCritical, "Alert"
'        Exit Sub
'    End If
'    If MsgBox("ARE YOU SURE IN UNPOSTING THIS TRANSACTION?                        ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
'    s = "SELECT TOP (1) PayrollPeriodKey, Locked " & _
'        " From dbo.tbl_Personnel_Payroll " & _
'        " WHERE (PK = " & locPayrollKey & ") " & _
'        " AND (Locked = 1)"
'    If rs.State = adStateOpen Then rs.Close
'    rs.Open s, ConnOmega
'    If rs.RecordCount > 0 Then
'        MsgBox "This payroll was already locked!                     ", vbCritical, "Error..."
'        rs.Close
'        Exit Sub
'    End If
'    rs.Close
'    Array1 = Split(Trim(txtCutOffDate.Text), " - ", -1, 1)
'    ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll_Earnings WHERE (MasterKey = " & locPayrollKey & ")"
'    ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll_Deductions WHERE (MasterKey = " & locPayrollKey & ")"
'    ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll_EmployerShare WHERE (MasterKey = " & locPayrollKey & ")"
'    ConnOmega.Execute "DELETE FROM tbl_Personnel_Loans_SL " & _
'                      " WHERE (PayrollKey = " & locPayrollKey & ") " & _
'                      " AND (TransactionDate = '" & FormatDateTime(txtPayrollPeriod.Text, vbShortDate) & "') " & _
'                      " AND (InOut = 'O')"
'    ConnOmega.Execute "DELETE FROM tbl_Personnel_Deduction_SL " & _
'                      " WHERE (PayrollKey = " & locPayrollKey & ") " & _
'                      " AND (TransactionDate = '" & FormatDateTime(txtPayrollPeriod.Text, vbShortDate) & "') " & _
'                      " AND (InOut = 'O')"
'
'    ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll WHERE (PK = " & locPayrollKey & ")"
'
'    s = "SELECT COUNT(*) AS RecCnt " & _
'        " From dbo.tbl_Personnel_Payroll " & _
'        " WHERE (ActionMemoKey = " & locActionMemoKey & ") "
'    If rs.State = adStateOpen Then rs.Close
'    rs.Open s, ConnOmega
'    If rs.RecordCount > 0 Then
'        If CDbl(rs!RecCnt) = 0 Then
'            ConnOmega.Execute "UPDATE tbl_Personnel_ActionNew SET Locked = 0 WHERE (PK = " & locActionMemoKey & ")"
'        End If
'    End If
'    rs.Close
'
'    ConnOmega.Execute "UPDATE tbl_Personnel_Hours " & _
'                      " SET PayrollKey = Null, " & _
'                      " Posted = 0, " & _
'                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
'                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
'
'    BROWSER GetSetting(App.EXEName, "PersonnelHours", "PersonnelHours", ""), "is_LOAD"
'
'Else
'    If AccessRights("Personnel - Hours", "Post") = False Then
'        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
'               "ACCESS DENIED!                                      ", vbCritical, "Alert"
'        Exit Sub
'    End If
'    If MsgBox("ARE YOU SURE IN POSTING THIS TRANSACTION?                        ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
'    s = "SELECT TOP (1) PayrollPeriodKey, Locked " & _
'        " From dbo.tbl_Personnel_Payroll " & _
'        " WHERE (PayrollPeriodKey = " & locPayrollPeroid & ") " & _
'        " AND (Locked = 1)"
'    If rs.State = adStateOpen Then rs.Close
'    rs.Open s, ConnOmega
'    If rs.RecordCount > 0 Then
'        MsgBox "This payroll period was already locked!                     ", vbCritical, "Error..."
'        rs.Close
'        Exit Sub
'    End If
'    rs.Close
'
'    COMPUTE_COMPENSATION Statusbar1.Panels(1).Text
'    BROWSER GetSetting(App.EXEName, "PersonnelHours", "PersonnelHours", ""), "is_LOAD"
'
'    If AccessRights("Personnel Compensation", "Open") = False Then Exit Sub
'
'    If MsgBox("Successfully Posted!                     " & vbCrLf & vbCrLf & _
'              "View Compensation Module?                ", vbInformation + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
'
'    gbl_Form_Caption = "Compensation"
'    If IsLoaded(frmPersonnelPayroll) Then frmPersonnelPayroll.ZOrder 0 Else frmPersonnelPayroll.Show
'    frmPersonnelPayroll.BROWSER locPayrollKey, "is_FIND"
'
'End If
'Exit Sub
'PG:
'MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
'Exit Sub
End Sub

Private Sub PRESS_F9()
If picAdd.Visible = True Then Exit Sub
If picSLRegularHours.Visible = True Then Exit Sub
If picSLOvertimeHours.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If picBatchPosting.Visible = True Then Exit Sub
If picProgressBar.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picAdd.Visible = True Then cmdCancelAdd_Click: Exit Sub
    If picSearch.Visible = True Then cmdCancelSearch_Click: Exit Sub
    If picBatchPosting.Visible = True Then cmdCancelPostUnpost_Click: Exit Sub
    Unload Me
Else
    If picSLRegularHours.Visible = True Then
    
        If TRANS_DETAIL = is_DET_ADDING Then
            With lstRegularHours.ListItems
                If .Count > 1 Then
                    .Remove .Count
                Else
                    .Item(1).Text = "0"
                    .Item(1).SubItems(1) = " "
                    .Item(1).SubItems(2) = " "
                End If
                iRowReg = .Count
            End With
            lstRegularHours.ListItems(iRowReg).EnsureVisible
            lstRegularHours.ListItems(iRowReg).Selected = True
            picSLRegularHours.Visible = False
            picMain.Enabled = True
            picToolbar.Enabled = True
            lstRegularHours.SetFocus
            
            If CheckDeductToReference(RETURNTEXTVALUE(txtEarningRegKey)) = 1 Then
                With lstRegularHours.ListItems
                    For i = 1 To .Count
                        If CLng(.Item(i).Text) = GetReferenceEarning(RETURNTEXTVALUE(txtEarningRegKey)) Then
                            .Item(i).SubItems(2) = dPerfectHoursMaster ' - RETURNTEXTVALUE(txtRegValue)
                            Exit For
                        End If
                    Next i
                End With
            End If
            
            Exit Sub
        End If
        If TRANS_DETAIL = is_DET_EDITTING Then
            With lstRegularHours.ListItems
                .Item(iRowReg).Text = txtEarningRegKey1.Text
                .Item(iRowReg).SubItems(1) = cmRegularHours1.Text
                .Item(iRowReg).SubItems(2) = txtRegValue1.Text
            End With
            lstRegularHours.ListItems(iRowReg).EnsureVisible
            lstRegularHours.ListItems(iRowReg).Selected = True
            picSLRegularHours.Visible = False
            picMain.Enabled = True
            picToolbar.Enabled = True
            lstRegularHours.SetFocus
            
            If CheckDeductToReference(RETURNTEXTVALUE(txtEarningRegKey1)) = 1 Then
                With lstRegularHours.ListItems
                    For i = 1 To .Count
                        If CLng(.Item(i).Text) = GetReferenceEarning(RETURNTEXTVALUE(txtEarningRegKey1)) Then
                            .Item(i).SubItems(2) = dPerfectHoursMaster - RETURNTEXTVALUE(txtRegValue1)
                            Exit For
                        End If
                    Next i
                End With
            End If
            
            Exit Sub
        End If
    End If
    If picSLOvertimeHours.Visible = True Then
        If TRANS_DETAIL = is_DET_ADDING Then
            With lstOvertime.ListItems
                If .Count > 1 Then
                    .Remove .Count
                Else
                    .Item(1).Text = "0"
                    .Item(1).SubItems(1) = " "
                    .Item(1).SubItems(2) = " "
                End If
                iRowOT = .Count
            End With
            lstOvertime.ListItems(iRowOT).EnsureVisible
            lstOvertime.ListItems(iRowOT).Selected = True
            picSLOvertimeHours.Visible = False
            picMain.Enabled = True
            picToolbar.Enabled = True
            lstOvertime.SetFocus
            Exit Sub
        End If
        If TRANS_DETAIL = is_DET_EDITTING Then
            With lstOvertime.ListItems
                .Item(iRowOT).Text = txtEarningOTKey1.Text
                .Item(iRowOT).SubItems(1) = cmbOvertimeHours1.Text
                .Item(iRowOT).SubItems(2) = txtOTValue1.Text
            End With
            lstOvertime.ListItems(iRowOT).EnsureVisible
            lstOvertime.ListItems(iRowOT).Selected = True
            picSLOvertimeHours.Visible = False
            picMain.Enabled = True
            picToolbar.Enabled = True
            lstOvertime.SetFocus
            Exit Sub
        End If
    End If
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    TRANS_DETAIL = is_DET_REFRESH
    BROWSER GetSetting(App.EXEName, "PersonnelHours", "PersonnelHours", ""), "is_LOAD"
    If Trim(txtControl.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelHours", "PersonnelHours", ""), "is_HOME"
    txtName.SetFocus
End If
End Sub

Public Sub BROWSER(Ctrl, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If Ctrl <> "" Then
            s = "SELECT TOP (1) dbo.tbl_Personnel_Hours.*, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, " & _
                " dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Division.Description AS Division, " & _
                " dbo.tbl_Personnel_Department.DepartmentName, dbo.tbl_Personnel_Position.PositionName, dbo.tbl_Personnel_Compensation_Period.DateFrom, " & _
                " dbo.tbl_Personnel_Compensation_Period.DateTo, dbo.tbl_Personnel_ActionNew.DivisionKey, dbo.tbl_Personnel_Compensation_Period.PayrollDate " & _
                " FROM  dbo.tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_Hours ON dbo.tbl_Personnel_IDNumber.PK = dbo.tbl_Personnel_Hours.EmployeeKey LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Hours.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_ActionNew.DivisionKey = dbo.tbl_Personnel_Division.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Hours.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
                " WHERE (dbo.tbl_Personnel_Hours.Ctrl = '" & Ctrl & "') " & _
                " ORDER BY dbo.tbl_Personnel_Hours.Ctrl"
        Else
            s = "SELECT TOP (1) dbo.tbl_Personnel_Hours.*, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, " & _
                " dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Division.Description AS Division, " & _
                " dbo.tbl_Personnel_Department.DepartmentName, dbo.tbl_Personnel_Position.PositionName, dbo.tbl_Personnel_Compensation_Period.DateFrom, " & _
                " dbo.tbl_Personnel_Compensation_Period.DateTo, dbo.tbl_Personnel_ActionNew.DivisionKey, dbo.tbl_Personnel_Compensation_Period.PayrollDate " & _
                " FROM  dbo.tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_Hours ON dbo.tbl_Personnel_IDNumber.PK = dbo.tbl_Personnel_Hours.EmployeeKey LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Hours.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_ActionNew.DivisionKey = dbo.tbl_Personnel_Division.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Hours.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
                " ORDER BY dbo.tbl_Personnel_Hours.Ctrl"
        End If
    Case "is_HOME"
        If picBatchPosting.Visible = True Then Exit Sub
        If picAdd.Visible = True Then Exit Sub
        If picSLRegularHours.Visible = True Then Exit Sub
        If picSLOvertimeHours.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Hours.*, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, " & _
            " dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Division.Description AS Division, " & _
            " dbo.tbl_Personnel_Department.DepartmentName, dbo.tbl_Personnel_Position.PositionName, dbo.tbl_Personnel_Compensation_Period.DateFrom, " & _
            " dbo.tbl_Personnel_Compensation_Period.DateTo, dbo.tbl_Personnel_ActionNew.DivisionKey, dbo.tbl_Personnel_Compensation_Period.PayrollDate " & _
            " FROM  dbo.tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK RIGHT OUTER JOIN " & _
            " dbo.tbl_Personnel_Hours ON dbo.tbl_Personnel_IDNumber.PK = dbo.tbl_Personnel_Hours.EmployeeKey LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Hours.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_ActionNew.DivisionKey = dbo.tbl_Personnel_Division.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Hours.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
            " ORDER BY dbo.tbl_Personnel_Hours.Ctrl"
    Case "is_PAGEUP"
        If picBatchPosting.Visible = True Then Exit Sub
        If picAdd.Visible = True Then Exit Sub
        If picSLRegularHours.Visible = True Then Exit Sub
        If picSLOvertimeHours.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Hours.*, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, " & _
            " dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Division.Description AS Division, " & _
            " dbo.tbl_Personnel_Department.DepartmentName, dbo.tbl_Personnel_Position.PositionName, dbo.tbl_Personnel_Compensation_Period.DateFrom, " & _
            " dbo.tbl_Personnel_Compensation_Period.DateTo, dbo.tbl_Personnel_ActionNew.DivisionKey, dbo.tbl_Personnel_Compensation_Period.PayrollDate " & _
            " FROM  dbo.tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK RIGHT OUTER JOIN " & _
            " dbo.tbl_Personnel_Hours ON dbo.tbl_Personnel_IDNumber.PK = dbo.tbl_Personnel_Hours.EmployeeKey LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Hours.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_ActionNew.DivisionKey = dbo.tbl_Personnel_Division.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Hours.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
            " WHERE (dbo.tbl_Personnel_Hours.Ctrl < '" & Ctrl & "') " & _
            " ORDER BY dbo.tbl_Personnel_Hours.Ctrl DESC"
    Case "is_PAGEDOWN"
        If picBatchPosting.Visible = True Then Exit Sub
        If picAdd.Visible = True Then Exit Sub
        If picSLRegularHours.Visible = True Then Exit Sub
        If picSLOvertimeHours.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Hours.*, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, " & _
            " dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Division.Description AS Division, " & _
            " dbo.tbl_Personnel_Department.DepartmentName, dbo.tbl_Personnel_Position.PositionName, dbo.tbl_Personnel_Compensation_Period.DateFrom, " & _
            " dbo.tbl_Personnel_Compensation_Period.DateTo, dbo.tbl_Personnel_ActionNew.DivisionKey, dbo.tbl_Personnel_Compensation_Period.PayrollDate " & _
            " FROM  dbo.tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK RIGHT OUTER JOIN " & _
            " dbo.tbl_Personnel_Hours ON dbo.tbl_Personnel_IDNumber.PK = dbo.tbl_Personnel_Hours.EmployeeKey LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Hours.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_ActionNew.DivisionKey = dbo.tbl_Personnel_Division.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Hours.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
            " WHERE (dbo.tbl_Personnel_Hours.Ctrl > '" & Ctrl & "') " & _
            " ORDER BY dbo.tbl_Personnel_Hours.Ctrl "
    Case "is_END"
        If picBatchPosting.Visible = True Then Exit Sub
        If picAdd.Visible = True Then Exit Sub
        If picSLRegularHours.Visible = True Then Exit Sub
        If picSLOvertimeHours.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Hours.*, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, " & _
            " dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Division.Description AS Division, " & _
            " dbo.tbl_Personnel_Department.DepartmentName, dbo.tbl_Personnel_Position.PositionName, dbo.tbl_Personnel_Compensation_Period.DateFrom, " & _
            " dbo.tbl_Personnel_Compensation_Period.DateTo, dbo.tbl_Personnel_ActionNew.DivisionKey, dbo.tbl_Personnel_Compensation_Period.PayrollDate " & _
            " FROM  dbo.tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK RIGHT OUTER JOIN " & _
            " dbo.tbl_Personnel_Hours ON dbo.tbl_Personnel_IDNumber.PK = dbo.tbl_Personnel_Hours.EmployeeKey LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Hours.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_ActionNew.DivisionKey = dbo.tbl_Personnel_Division.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Hours.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
            " ORDER BY dbo.tbl_Personnel_Hours.Ctrl DESC"
    
    Case "is_FIND"
        s = "SELECT TOP (1) dbo.tbl_Personnel_Hours.*, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, " & _
            " dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Division.Description AS Division, " & _
            " dbo.tbl_Personnel_Department.DepartmentName, dbo.tbl_Personnel_Position.PositionName, dbo.tbl_Personnel_Compensation_Period.DateFrom, " & _
            " dbo.tbl_Personnel_Compensation_Period.DateTo, dbo.tbl_Personnel_ActionNew.DivisionKey, dbo.tbl_Personnel_Compensation_Period.PayrollDate " & _
            " FROM  dbo.tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK RIGHT OUTER JOIN " & _
            " dbo.tbl_Personnel_Hours ON dbo.tbl_Personnel_IDNumber.PK = dbo.tbl_Personnel_Hours.EmployeeKey LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Hours.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_ActionNew.DivisionKey = dbo.tbl_Personnel_Division.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Hours.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
            " WHERE (dbo.tbl_Personnel_Hours.PK = " & Ctrl & ") " & _
            " ORDER BY dbo.tbl_Personnel_Hours.Ctrl DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    locEmployeePK = rs!EmployeeKey
    locPayrollPeroid = rs!PayrollPeriodKey
    locActionMemoKey = rs!ActionMemoKey
    locDivision = rs!DivisionKey
    locPayrollKey = IIf(IsNull(rs!PayrollKey), 0, rs!PayrollKey)
    dPerfectHoursMaster = rs!PerfectHours
    txtControl.Text = rs!Ctrl
    txtName.Text = rs!IDNumber & " - " & rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
    txtPayrollPeriod.Text = Format(rs!PayrollDate, "mm/dd/yyyy")
    txtCutOffDate.Text = Format(rs!DateFrom, "mm/dd/yyyy") & " - " & Format(rs!DateTo, "mm/dd/yyyy")
    txtDivision.Text = rs!Division
    txtDepartment.Text = rs!DepartmentName
    txtPosition.Text = rs!PositionName
    txtCompType.Text = ""
    t = "SELECT dbo.tbl_Personnel_CompensationRate.Description " & _
        " FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_CompensationRate ON dbo.tbl_Personnel_ActionNew.CompensationRateKey = dbo.tbl_Personnel_CompensationRate.PK " & _
        " WHERE (dbo.tbl_Personnel_ActionNew.PK = " & rs!ActionMemoKey & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtCompType.Text = rt!Description
    End If
    rt.Close
    txtAdjustment.Text = Format(rs!Adjustment, "#,##0.00")
    txtAdjustmentRem.Text = rs!AdjustmentRem
    
    CLEAR_DETAILS_Reg
    t = "SELECT dbo.tbl_Personnel_Hours_Regular.MasterKey, dbo.tbl_Personnel_Hours_Regular.EarningKey, " & _
        " dbo.tbl_Personnel_Payroll_Earnings_Table.HoursDesc, dbo.tbl_Personnel_Hours_Regular.NoHours " & _
        " FROM  dbo.tbl_Personnel_Hours_Regular LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Payroll_Earnings_Table ON dbo.tbl_Personnel_Hours_Regular.EarningKey = dbo.tbl_Personnel_Payroll_Earnings_Table.PK " & _
        " Where (dbo.tbl_Personnel_Hours_Regular.MasterKey = " & rs!PK & ") " & _
        " ORDER BY dbo.tbl_Personnel_Payroll_Earnings_Table.Sorting"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstRegularHours.ListItems.Clear
        While Not rt.EOF
            Set x = lstRegularHours.ListItems.Add()
            x.Text = rt!EarningKey
            x.SubItems(1) = rt!HoursDesc
            If rt!EarningKey = 1 Then
                x.SubItems(2) = rt!NoHours
            Else
                x.SubItems(2) = Format(rt!NoHours, "#,##0.00")
            End If
            rt.MoveNext
        Wend
    End If
    rt.Close
    
    CLEAR_DETAILS_OT
    t = "SELECT dbo.tbl_Personnel_Hours_Overtime.MasterKey, dbo.tbl_Personnel_Hours_Overtime.EarningKey, " & _
        " dbo.tbl_Personnel_Payroll_Earnings_Table.HoursDesc, dbo.tbl_Personnel_Hours_Overtime.NoHours " & _
        " FROM  dbo.tbl_Personnel_Hours_Overtime LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Payroll_Earnings_Table ON dbo.tbl_Personnel_Hours_Overtime.EarningKey = dbo.tbl_Personnel_Payroll_Earnings_Table.PK " & _
        " Where (dbo.tbl_Personnel_Hours_Overtime.MasterKey = " & rs!PK & ") " & _
        " ORDER BY dbo.tbl_Personnel_Payroll_Earnings_Table.Sorting"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstOvertime.ListItems.Clear
        While Not rt.EOF
            Set x = lstOvertime.ListItems.Add()
            x.Text = rt!EarningKey
            x.SubItems(1) = rt!HoursDesc
            x.SubItems(2) = Format(rt!NoHours, "#,##0.00")
            rt.MoveNext
        Wend
    End If
    rt.Close
    
    imgPosted.Visible = IIf(rs!Posted = 1, True, False)
    Toolbar1.Buttons(19).Caption = IIf(rs!Posted = 1, "UnPost", " Post ")
    Toolbar1.Buttons(19).Image = IIf(rs!Posted = 1, 11, 10)
    
    If AccessRights("Personnel Compensation", "Open") = True Then
        cmdViewPayroll.Enabled = IIf(locPayrollKey = 0, False, True)
    Else
        cmdViewPayroll.Enabled = False
    End If
    
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    
    SaveSetting App.EXEName, "PersonnelHours", "PersonnelHours", rs!Ctrl
    
End If
rs.Close
End Sub

Private Sub CLEARTEXT()
locEmployeePK = 0
locPayrollPeroid = 0
locActionMemoKey = 0
locDivision = 0
locPayrollKey = 0
dPerfectHoursMaster = 0
txtControl.Text = ""
txtName.Text = ""
txtPayrollPeriod.Text = ""
txtCutOffDate.Text = ""
txtDivision.Text = ""
txtDepartment.Text = ""
txtPosition.Text = ""
txtCompType.Text = ""
txtAdjustment.Text = "0.00"
txtAdjustmentRem.Text = ""
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
CLEAR_DETAILS_Reg
CLEAR_DETAILS_OT
imgPosted.Visible = False
cmdViewPayroll.Enabled = False
End Sub

Private Sub CLEAR_DETAILS_Reg()
lstRegularHours.ListItems.Clear
Set x = lstRegularHours.ListItems.Add()
x.Text = "0"
x.SubItems(1) = " "
x.SubItems(2) = " "
End Sub

Private Sub CLEAR_DETAILS_OT()
lstOvertime.ListItems.Clear
Set x = lstOvertime.ListItems.Add()
x.Text = "0"
x.SubItems(1) = " "
x.SubItems(2) = " "
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtName.Locked = True
txtPayrollPeriod.Locked = True
txtCutOffDate.Locked = True
txtDivision.Locked = True
txtDepartment.Locked = True
txtPosition.Locked = True
txtCompType.Locked = True
txtAdjustment.Locked = bln
txtAdjustmentRem.Locked = bln
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
            .Buttons(19).Image = 10
            .Buttons(21).Image = 12
            .Buttons(23).Image = 13
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "First"
            .Buttons(9).Caption = "Back"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            .Buttons(19).Caption = "Post"
            .Buttons(21).Caption = "Refresh"
            .Buttons(23).Caption = "Close"
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
            .Buttons(23).Enabled = True
            .Buttons(1).ToolTipText = "NEW (Ins)"
            .Buttons(3).ToolTipText = "EDIT (F2)"
            .Buttons(5).ToolTipText = "DELETE (Del)"
            .Buttons(7).ToolTipText = "FIRST (Home)"
            .Buttons(9).ToolTipText = "BACK (PgUp)"
            .Buttons(11).ToolTipText = "NEXT (PgDown)"
            .Buttons(13).ToolTipText = "LAST (End)"
            .Buttons(15).ToolTipText = "FIND (F6)"
            .Buttons(17).ToolTipText = "PRINT (F9)"
            .Buttons(19).ToolTipText = "POST (F8)"
            .Buttons(21).ToolTipText = "REFRESH (F11)"
            .Buttons(23).ToolTipText = "CLOSE (Esc)"
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
            .Buttons(19).Image = 10
            .Buttons(21).Image = 12
            .Buttons(23).Image = 13
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "Save"
            .Buttons(9).Caption = "Undo"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            .Buttons(19).Caption = "Post"
            .Buttons(21).Caption = "Refresh"
            .Buttons(23).Caption = "Close"
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
            .Buttons(23).Enabled = False
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
            .Buttons(23).ToolTipText = ""
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
            .Buttons(19).Image = 10
            .Buttons(21).Image = 12
            .Buttons(23).Image = 13
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "First"
            .Buttons(9).Caption = "Undo"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            .Buttons(19).Caption = "Post"
            .Buttons(21).Caption = "Refresh"
            .Buttons(23).Caption = "Close"
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
            .Buttons(23).Enabled = False
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
            .Buttons(23).ToolTipText = ""
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
            .Buttons(19).Image = 10
            .Buttons(21).Image = 12
            .Buttons(23).Image = 13
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "Save"
            .Buttons(9).Caption = "Undo"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            .Buttons(19).Caption = "Post"
            .Buttons(21).Caption = "Refresh"
            .Buttons(23).Caption = "Close"
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
            .Buttons(23).Enabled = False
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
            .Buttons(23).ToolTipText = ""
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
            .Buttons(19).Image = 10
            .Buttons(21).Image = 12
            .Buttons(23).Image = 13
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "Save"
            .Buttons(9).Caption = "Undo"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            .Buttons(19).Caption = "Post"
            .Buttons(21).Caption = "Refresh"
            .Buttons(23).Caption = "Close"
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
            .Buttons(23).Enabled = False
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
            .Buttons(23).ToolTipText = ""
        Case 6      '=== NOT EMPTY DETAIL NAME ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 14
            .Buttons(9).Image = 15
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(21).Image = 12
            .Buttons(23).Image = 13
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "Save"
            .Buttons(9).Caption = "Undo"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            .Buttons(19).Caption = "Post"
            .Buttons(21).Caption = "Refresh"
            .Buttons(23).Caption = "Close"
            .Buttons(1).Enabled = True
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = True
            .Buttons(7).Enabled = True
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
            .Buttons(19).Enabled = False
            .Buttons(21).Enabled = False
            .Buttons(23).Enabled = False
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
            .Buttons(23).ToolTipText = ""
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
cmdCancelPostUnpost_Click
End Sub

Private Sub cmbDivisionAdd_Click()
If cmbDivisionAdd.ListIndex = -1 Then txtFrom.Text = "": txtTo.Text = "": txtSearchAdd.Text = "": lstResultAdd.Clear: Exit Sub
txtFrom.Text = "": txtTo.Text = "": txtSearchAdd.Text = "": lstResultAdd.Clear
t = "SELECT TOP (1) DateFrom, DateTo, PayrollDate " & _
    " From dbo.tbl_Personnel_Compensation_Period " & _
    " WHERE (Type = " & cmbDivisionAdd.ItemData(cmbDivisionAdd.ListIndex) & ") " & _
    " AND (DateTo <= '" & FormatDateTime(Date, vbShortDate) & "') " & _
    " ORDER BY DateTo DESC"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then
    txtFrom.Text = Format(rt!DateFrom, "mm/dd/yyyy")
    txtTo.Text = Format(rt!DateTo, "mm/dd/yyyy")
    txtPayrollDateAdd.Text = Format(rt!PayrollDate, "mm/dd/yyyy")
End If
rt.Close
End Sub

Private Sub cmbDivisionAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtPayrollDateAdd.SetFocus
End Sub

Private Sub cmbDivisionBatchPost_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtPayrollDatePostUnpost.SetFocus
End Sub

Private Sub cmbOvertimeHours_Click()
If cmbOvertimeHours.ListIndex = -1 Then Exit Sub
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    txtEarningOTKey.Text = cmbOvertimeHours.ItemData(cmbOvertimeHours.ListIndex)
    With lstOvertime.ListItems
        .Item(iRowOT).SubItems(1) = cmbOvertimeHours.List(cmbOvertimeHours.ListIndex)
    End With
End If
End Sub

Private Sub cmbOvertimeHours_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtOTValue.SetFocus
End Sub

Private Sub cmbPayrollPeriodSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKSearch_Click
End Sub

Private Sub cmbPostUnpost_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbDivisionBatchPost.SetFocus
End Sub

Private Sub cmdCancelAdd_Click()
picAdd.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True
End Sub

Private Sub cmdCancelPostUnpost_Click()
picBatchPosting.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True
End Sub

Private Sub cmdCancelSearch_Click()
picSearch.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True
End Sub

Private Sub cmdOKAdd_Click()
If cmbDivisionAdd.ListIndex = -1 Then Exit Sub
'If IsDate(txtFrom.Text) = False Then MsgBox "Please supply a valid date!                ", vbCritical, "Error...": txtPayrollDateAdd.SetFocus: Exit Sub
'If IsDate(txtTo.Text) = False Then MsgBox "Please supply a valid date!                ", vbCritical, "Error...": txtPayrollDateAdd.SetFocus: Exit Sub
If IsDate(txtPayrollDateAdd.Text) = False Then MsgBox "Please supply a valid date!                ", vbCritical, "Error...": txtPayrollDateAdd.SetFocus: Exit Sub
If lstResultAdd.ListIndex = -1 Then Exit Sub

locPayrollPeroidTmp = GET_PERIOD_V2(FormatDateTime(txtPayrollDateAdd.Text, vbShortDate), cmbDivisionAdd.ItemData(cmbDivisionAdd.ListIndex))

If locPayrollPeroidTmp = 0 Then
    MsgBox "Payroll Period Not Match to the Employee Division!      ", vbInformation, ""
    txtPayrollDateAdd.SetFocus
    HTEXT txtPayrollDateAdd
    Exit Sub
End If

iEmpStatus = GET_EMPLOYMENT_STATUS(lstResultAdd.ItemData(lstResultAdd.ListIndex), FormatDateTime(txtPayrollDateAdd.Text, vbShortDate))

If CDbl(iEmpStatus) = 0 _
Or CDbl(iEmpStatus) = 2 Then
    Array1 = Split(lstResultAdd.List(lstResultAdd.ListIndex), " - ", -1, 1)
    MsgBox Array1(1) & " WAS ALREADY INACTIVE!                      ", vbCritical, "Error..."
    Exit Sub
End If

'Check Payroll Locked
't = "SELECT COUNT(dbo.tbl_Personnel_Payroll.PK) AS RecCount " & _
    " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK " & _
    " WHERE (dbo.tbl_Personnel_Compensation_Period.PayrollDate < '" & FormatDateTime(txtPayrollDateAdd.Text, vbShortDate) & "') " & _
    " AND (dbo.tbl_Personnel_ActionNew.DivisionKey = " & cmbDivisionAdd.ItemData(cmbDivisionAdd.ListIndex) & ") " & _
    " AND (dbo.tbl_Personnel_Payroll.Locked = 0)"
t = "SELECT COUNT(dbo.tbl_Personnel_Payroll.PK) AS RecCnt " & _
    " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK " & _
    " WHERE (dbo.tbl_Personnel_Compensation_Period.PayrollDate < '" & FormatDateTime(txtPayrollDateAdd.Text, vbShortDate) & "') " & _
    " AND (dbo.tbl_Personnel_ActionNew.DivisionKey = " & cmbDivisionAdd.ItemData(cmbDivisionAdd.ListIndex) & ") " & _
    " AND (dbo.tbl_Personnel_Payroll.Locked = 0)"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then
    If CDbl(IIf(IsNull(rt!RecCnt), 0, rt!RecCnt)) > 0 Then
        MsgBox "Please locked previous payroll!             ", vbCritical, "Error..."
        rt.Close
        Exit Sub
    End If
End If
rt.Close

iChkPerfHrs = CheckPerfectDays(lstResultAdd.ItemData(lstResultAdd.ListIndex), FormatDateTime(txtPayrollDateAdd.Text, vbShortDate))
If CDbl(iChkPerfHrs) = 1 Then
    t = "SELECT tbl_Personnel_Setup_DailyPerfectDays.* " & _
        " FROM tbl_Personnel_Setup_DailyPerfectDays " & _
        " WHERE (DivisionKey = " & cmbDivisionAdd.ItemData(cmbDivisionAdd.ListIndex) & ") " & _
        " AND (PayrollPeriodKey = " & locPayrollPeroidTmp & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount = 0 Then
        MsgBox "Please setup perfect number of hours!                           ", vbCritical, "Error..."
        rt.Close
        Exit Sub
    Else
        If CDbl(rt!NoHours) <= 0 Then
            MsgBox "Invalid perfec hours value, please check set up!                    ", vbCritical, "Error..."
            rt.Close
            Exit Sub
        End If
    End If
    rt.Close
End If


CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING

locEmployeePK = lstResultAdd.ItemData(lstResultAdd.ListIndex)
locPayrollPeroid = locPayrollPeroidTmp 'GET_PERIOD_V2(FormatDateTime(txtPayrollDateAdd.Text, vbShortDate), cmbDivisionAdd.ItemData(cmbDivisionAdd.ListIndex)) 'cmbDivisionAdd.ItemData(cmbDivisionAdd.ListIndex)
txtName.Text = lstResultAdd.List(lstResultAdd.ListIndex)
txtPayrollPeriod.Text = Format(FormatDateTime(txtPayrollDateAdd.Text, vbShortDate), "mm/dd/yyyy")
Array1 = Split(GET_PERIOD_CUTOFF(locPayrollPeroid), " - ", -1, 1)
txtCutOffDate.Text = Format(Array1(0), "mm/dd/yyyy") & " - " & Format(Array1(1), "mm/dd/yyyy") 'GET_PERIOD_CUTOFF(locPayrollPeroid)
t = "SELECT TOP (1) dbo.tbl_Personnel_ActionNew.PK, dbo.tbl_Personnel_Division.Description AS Division, " & _
    " dbo.tbl_Personnel_Department.DepartmentName, dbo.tbl_Personnel_Position.PositionName, dbo.tbl_Personnel_ActionNew.DivisionKey " & _
    " FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_ActionNew.DivisionKey = dbo.tbl_Personnel_Division.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
    " WHERE (dbo.tbl_Personnel_ActionNew.EmpPK = " & locEmployeePK & ") " & _
    " AND (dbo.tbl_Personnel_ActionNew.EffectivityDate <= '" & FormatDateTime(txtPayrollPeriod.Text) & "') " & _
    " ORDER BY dbo.tbl_Personnel_ActionNew.EffectivityDate DESC"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then
    locActionMemoKey = rt!PK
    locDivision = rt!DivisionKey
    txtDivision.Text = rt!Division
    txtDepartment.Text = rt!DepartmentName
    txtPosition.Text = rt!PositionName
End If
rt.Close

t = "SELECT dbo.tbl_Personnel_CompensationRate.Description " & _
    " FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_CompensationRate ON dbo.tbl_Personnel_ActionNew.CompensationRateKey = dbo.tbl_Personnel_CompensationRate.PK " & _
    " WHERE (dbo.tbl_Personnel_ActionNew.PK = " & locActionMemoKey & ")"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then
    txtCompType.Text = rt!Description
End If
rt.Close

dPerfectHours = Get_Perfect_Hours(iChkPerfHrs, txtPayrollDateAdd.Text, cmbDivisionAdd.ItemData(cmbDivisionAdd.ListIndex), locPayrollPeroid)
dAbsentLaterUndertime = Get_AbsentLateUndertime_Hours(lstResultAdd.ItemData(lstResultAdd.ListIndex), Array1(0), Array1(1))
dNoHours = CDbl(dPerfectHours) - CDbl(dAbsentLaterUndertime)
dPerfectHoursMaster = dNoHours
With lstRegularHours.ListItems
    .Clear
    Set x = .Add()
    x.Text = "1"
    x.SubItems(1) = "No of Hours"
    x.SubItems(2) = dNoHours
End With

cmdCancelAdd_Click
txtAdjustment.SetFocus
End Sub

Private Sub cmdOKPostUnpost_Click()
If cmbPostUnpost.ListIndex = -1 Then Exit Sub
If cmbDivisionBatchPost.ListIndex = -1 Then Exit Sub
If IsDate(txtPayrollDatePostUnpost.Text) = False Then MsgBox "Please supply a valid date!                 ", vbCritical, "Error...": txtPayrollDatePostUnpost.SetFocus: Exit Sub
txtPayrollDatePostUnpost.Text = Format(FormatDateTime(txtPayrollDatePostUnpost.Text), "mm/dd/yyyy")
locPayrollPeriodTmp = GET_PERIOD_V2(FormatDateTime(txtPayrollDatePostUnpost.Text, vbShortDate), cmbDivisionBatchPost.ItemData(cmbDivisionBatchPost.ListIndex))
If locPayrollPeriodTmp = 0 Then
    MsgBox "Payroll Period Not Match to the Employee Division!      ", vbInformation, ""
    txtPayrollDatePostUnpost.SetFocus
    HTEXT txtPayrollDatePostUnpost
    Exit Sub
End If
locPayrollPeroid = locPayrollPeriodTmp
Select Case cmbPostUnpost.ItemData(cmbPostUnpost.ListIndex)
    Case 1  'Posting
        If AccessRights("Personnel - Hours", "Post") = False Then
            MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
                   "ACCESS DENIED!                                      ", vbCritical, "Alert"
            Exit Sub
        End If
        
        s = "SELECT TOP (1) PayrollPeriodKey, Locked " & _
            " From dbo.tbl_Personnel_Payroll " & _
            " WHERE (PayrollPeriodKey = " & locPayrollPeroid & ") " & _
            " AND (Locked = 1)"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            MsgBox "This payroll period was already locked!                     ", vbCritical, "Error..."
            rs.Close
            Exit Sub
        End If
        rs.Close
        
        s = "SELECT tbl_Personnel_Deduction_forPayroll.* " & _
            " FROM tbl_Personnel_Deduction_forPayroll " & _
            " WHERE (DivisionKey = " & cmbDivisionBatchPost.ItemData(cmbDivisionBatchPost.ListIndex) & ")  " & _
            " AND (PayrollPeriodKey = " & locPayrollPeroid & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            MsgBox "Please add for deduction for this division and payroll date!                    ", vbCritical, "Error..."
            rs.Close
            Exit Sub
        Else
            If CDbl(rs!Posted) = 0 Then
                MsgBox "Please post the for deduction for this division and payroll date!                    ", vbCritical, "Error..."
                rs.Close
                Exit Sub
            End If
        End If
        rs.Close
        
        If MsgBox("ARE YOU SURE IN POSTING THOSE TRANSACTION?                        ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
        
        s = "SELECT tbl_Personnel_Hours.* " & _
            " FROM tbl_Personnel_Hours " & _
            " WHERE (PayrollPeriodKey = " & locPayrollPeroid & ") " & _
            " AND (Posted = 0)"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            On Error GoTo PG:
            cnt = 0
            picBatchPosting.Visible = False
            picProgressBar.ZOrder 0
            picProgress.BackColor = &HFFFFFF
            picProgressBar.Visible = True
            While Not rs.EOF
                DoEvents
                cnt = cnt + 1
                COMPUTE_COMPENSATION rs!PK
                UpdateProgress picProgress, cnt / rs.RecordCount
                rs.MoveNext
            Wend
        Else
            MsgBox "No record to post!                 ", vbExclamation, "Info"
            rs.Close
            Exit Sub
        End If
        rs.Close
        
        picProgressBar.Visible = False
        picMain.Enabled = True
        picToolbar.Enabled = True
        BROWSER GetSetting(App.EXEName, "PersonnelHours", "PersonnelHours", ""), "is_LOAD"
        
    Case 2  'Unpost
          
        If AccessRights("Personnel - Hours", "UnPost") = False Then
            MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
                   "ACCESS DENIED!                                      ", vbCritical, "Alert"
            Exit Sub
        End If
        
        s = "SELECT TOP (1) PayrollPeriodKey, Locked " & _
            " From dbo.tbl_Personnel_Payroll " & _
            " WHERE (PayrollPeriodKey = " & locPayrollPeroid & ") " & _
            " AND (Locked = 1)"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            MsgBox "This payroll period was already locked!                     ", vbCritical, "Error..."
            rs.Close
            Exit Sub
        End If
        rs.Close
        
        
        If MsgBox("ARE YOU SURE IN UNPOSTING THOSE TRANSACTION?                        ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
        
        t = "SELECT dbo.tbl_Personnel_Hours.PK, dbo.tbl_Personnel_Hours.PayrollKey, " & _
            " dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo, " & _
            " dbo.tbl_Personnel_Compensation_Period.PayrollDate, dbo.tbl_Personnel_Hours.ActionMemoKey " & _
            " FROM  dbo.tbl_Personnel_Hours LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Hours.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
            " WHERE (dbo.tbl_Personnel_Hours.PayrollPeriodKey = " & locPayrollPeroid & ") " & _
            " AND (dbo.tbl_Personnel_Hours.Posted = 1)"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            On Error GoTo PG:
            cnt = 0
            picBatchPosting.Visible = False
            picProgressBar.ZOrder 0
            picProgress.BackColor = &HFFFFFF
            picProgressBar.Visible = True
            While Not rt.EOF
                DoEvents
                cnt = cnt + 1
                'Arr = Split(Trim(.txtCutOffDate.Text), " - ", -1, 1)
                ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll_Earnings WHERE (MasterKey = " & rt!PayrollKey & ")"
                ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll_Deductions WHERE (MasterKey = " & rt!PayrollKey & ")"
                ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll_EmployerShare WHERE (MasterKey = " & rt!PayrollKey & ")"
                ConnOmega.Execute "DELETE FROM tbl_Personnel_Loans_SL " & _
                                  " WHERE (PayrollKey = " & rt!PayrollKey & ") " & _
                                  " AND (TransactionDate = '" & FormatDateTime(rt!PayrollDate, vbShortDate) & "') " & _
                                  " AND (InOut = 'O')"
                ConnOmega.Execute "DELETE FROM tbl_Personnel_Deduction_SL " & _
                                  " WHERE (PayrollKey = " & rt!PayrollKey & ") " & _
                                  " AND (TransactionDate = '" & FormatDateTime(rt!PayrollDate, vbShortDate) & "') " & _
                                  " AND (TransactionType = 2) " & _
                                  " AND (InOut = 'O')"
                
                ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll WHERE (PK = " & rt!PayrollKey & ")"
                
                s = "SELECT COUNT(*) AS RecCnt " & _
                    " From dbo.tbl_Personnel_Payroll " & _
                    " WHERE (ActionMemoKey = " & rt!ActionMemoKey & ") "
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                If rs.RecordCount > 0 Then
                    If CDbl(rs!RecCnt) = 0 Then
                        ConnOmega.Execute "UPDATE tbl_Personnel_ActionNew SET Locked = 0 WHERE (PK = " & rt!ActionMemoKey & ")"
                    End If
                End If
                rs.Close
                
                ConnOmega.Execute "UPDATE tbl_Personnel_Hours " & _
                                  " SET PayrollKey = Null, " & _
                                  " Posted = 0, " & _
                                  " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                                  " WHERE (PK = " & rt!PK & ")"
                
                UpdateProgress picProgress, cnt / rt.RecordCount
                rt.MoveNext
            Wend
        Else
            MsgBox "No record to unpost!                 ", vbExclamation, "Info"
            rt.Close
            Exit Sub
        End If
        rt.Close
        
        picProgressBar.Visible = False
        picMain.Enabled = True
        picToolbar.Enabled = True
        BROWSER GetSetting(App.EXEName, "PersonnelHours", "PersonnelHours", ""), "is_LOAD"
        
End Select

Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error.."
Exit Sub
End Sub

Private Sub cmdOKSearch_Click()
If cmbPayrollPeriodSearch.ListIndex = -1 Then Exit Sub
'MsgBox cmbPayrollPeriodSearch.ItemData(cmbPayrollPeriodSearch.ListIndex)
BROWSER cmbPayrollPeriodSearch.ItemData(cmbPayrollPeriodSearch.ListIndex), "is_FIND"
cmdCancelSearch_Click
End Sub

Private Sub cmdViewPayroll_Click()
picSetFocus.SetFocus
If AccessRights("Personnel Compensation", "Open") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
gbl_Form_Caption = "Compensation"
If IsLoaded(frmPersonnelPayroll) Then frmPersonnelPayroll.ZOrder 0 Else frmPersonnelPayroll.Show
frmPersonnelPayroll.BROWSER locPayrollKey, "is_FIND"
End Sub

Private Sub cmRegularHours_Click()
If cmRegularHours.ListIndex = -1 Then Exit Sub
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    txtEarningRegKey.Text = cmRegularHours.ItemData(cmRegularHours.ListIndex)
    With lstRegularHours.ListItems
        .Item(iRowReg).SubItems(1) = cmRegularHours.List(cmRegularHours.ListIndex)
    End With
End If
End Sub

Private Sub cmRegularHours_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtRegValue.SetFocus
End Sub

Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
s = "SELECT dbo.tbl_Personnel_Compensation.*, dbo.tbl_Personnel_Compensation_Period.DateTo, dbo.tbl_Personnel_Compensation.PK as CompKey, " & _
    " dbo.tbl_Personnel_Compensation_Period.PayrollDate " & _
    " FROM  dbo.tbl_Personnel_Compensation LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Compensation.Period = dbo.tbl_Personnel_Compensation_Period.PK " & _
    " WHERE (dbo.tbl_Personnel_Compensation.Transfered = 0)"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    
    ' Hours
    
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Hours WHERE (EmployeeKey = " & rs!EmpPK & ") AND (PayrollPeriodKey = " & rs!Period & ")"
    
    sCtrl = ""
    t = "SELECT TOP (1) dbo.tbl_Personnel_Hours.Ctrl " & _
        " FROM  dbo.tbl_Personnel_Hours LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Hours.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
        " Where (Year(dbo.tbl_Personnel_Compensation_Period.DateTo) = " & Format(rs!DateTo, "yyyy") & ") " & _
        " ORDER BY dbo.tbl_Personnel_Hours.Ctrl DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        sCtrl = Format(CDbl(rt!Ctrl) + 1, "000000000#")
    Else
        sCtrl = Format(rs!DateTo, "yyyy") & "000000"
    End If
    rt.Close
    
    Do
        t = "SELECT tbl_Personnel_Hours.* " & _
            " FROM tbl_Personnel_Hours " & _
            " WHERE (Ctrl = '" & sCtrl & "')"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount = 0 Then
            rt.Close
            Exit Do
        End If
        rt.Close
        sCtrl = Format(CDbl(sCtrl) + 1, "000000000#")
    Loop
    
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Hours " & _
                      " (Ctrl, EmployeeKey, PayrollPeriodKey, ActionMemoKey, Adjustment, PayrollKey, LastModified, Posted) " & _
                      " VALUES ('" & sCtrl & "', " & rs!EmpPK & ", " & rs!Period & ", " & rs!ActionMemo & ", " & CDbl(rs!Adjustment) & ", " & _
                      " " & rs!PK & ", '" & IIf(IsNull(rs!LastModified), "", rs!LastModified) & "', 1)"
    
    iPK = 0
    t = "SELECT tbl_Personnel_Hours.* " & _
        " FROM tbl_Personnel_Hours " & _
        " WHERE (Ctrl = '" & sCtrl & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        iPK = rt!PK
    End If
    rt.Close
    
    If CDbl(iPK) > 0 Then
        If CDbl(rs!NoHours) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Hours_Regular " & _
                              " (MasterKey, EarningKey, NoHours) " & _
                              " VALUES (" & iPK & ", 1, " & CDbl(rs!NoHours) & ")"
        End If
        If CDbl(rs!ColaHours) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Hours_Regular " & _
                              " (MasterKey, EarningKey, NoHours) " & _
                              " VALUES (" & iPK & ", 2, " & CDbl(rs!ColaHours) & ")"
        End If
        If CDbl(rs!LH_Hours) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Hours_Regular " & _
                              " (MasterKey, EarningKey, NoHours) " & _
                              " VALUES (" & iPK & ", 4, " & CDbl(rs!LH_Hours) & ")"
        End If
        If CDbl(rs!SH_Hours) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Hours_Regular " & _
                              " (MasterKey, EarningKey, NoHours) " & _
                              " VALUES (" & iPK & ", 5, " & CDbl(rs!SH_Hours) & ")"
        End If
        If CDbl(rs!SL_Hours) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Hours_Regular " & _
                              " (MasterKey, EarningKey, NoHours) " & _
                              " VALUES (" & iPK & ", 6, " & CDbl(rs!SL_Hours) & ")"
        End If
        ' OT
        If CDbl(rs!Reg_OT_Hours) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Hours_Overtime " & _
                              " (MasterKey, EarningKey, NoHours) " & _
                              " VALUES (" & iPK & ", 8, " & CDbl(rs!Reg_OT_Hours) & ")"
        End If
        If CDbl(rs!RD_OT_Hours) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Hours_Overtime " & _
                              " (MasterKey, EarningKey, NoHours) " & _
                              " VALUES (" & iPK & ", 9, " & CDbl(rs!RD_OT_Hours) & ")"
        End If
        If CDbl(rs!LH_OT_Hours) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Hours_Overtime " & _
                              " (MasterKey, EarningKey, NoHours) " & _
                              " VALUES (" & iPK & ", 10, " & CDbl(rs!LH_OT_Hours) & ")"
        End If
        If CDbl(rs!SH_OT_Hours) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Hours_Overtime " & _
                              " (MasterKey, EarningKey, NoHours) " & _
                              " VALUES (" & iPK & ", 11, " & CDbl(rs!SH_OT_Hours) & ")"
        End If
    End If
    
    '   Payroll
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll_Earnings WHERE (MasterKey = " & rs!CompKey & ")"
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll_Deductions WHERE (MasterKey = " & rs!CompKey & ")"
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll_EmployerShare WHERE (MasterKey = " & rs!CompKey & ")"
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Loans_SL " & _
                      " WHERE (PayrollKey = " & rs!CompKey & ") " & _
                      " AND (TransactionDate = '" & FormatDateTime(rs!PayrollDate, vbShortDate) & "') " & _
                      " AND (InOut = 'O')"
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Deduction_SL " & _
                      " WHERE (PayrollKey = " & rs!CompKey & ") " & _
                      " AND (TransactionDate = '" & FormatDateTime(rs!PayrollDate, vbShortDate) & "') " & _
                      " AND (InOut = 'O')"
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll WHERE (PK = " & rs!CompKey & ")"
    
    
    sCtrl = ""
    t = "SELECT TOP (1) dbo.tbl_Personnel_Payroll.Ctrl " & _
        " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
        " Where (Year(dbo.tbl_Personnel_Compensation_Period.DateTo) = " & Format(rs!DateTo, "yyyy") & ") " & _
        " ORDER BY dbo.tbl_Personnel_Payroll.Ctrl DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        sCtrl = Format(CDbl(rt!Ctrl) + 1, "000000000#")
    Else
        sCtrl = Format(rs!DateTo, "yyyy") & "000000"
    End If
    rt.Close
    
    Do
        t = "SELECT tbl_Personnel_Payroll.* " & _
            " FROM tbl_Personnel_Payroll " & _
            " WHERE (Ctrl = '" & sCtrl & "')"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount = 0 Then
            rt.Close
            Exit Do
        End If
        rt.Close
        sCtrl = Format(CDbl(sCtrl) + 1, "000000000#")
    Loop
    
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll " & _
                      " (PK, Ctrl, EmployeeKey, PayrollPeriodKey, ActionMemoKey, Locked, LastModified) " & _
                      " VALUES (" & rs!CompKey & ", '" & sCtrl & "', " & rs!EmpPK & ", " & rs!Period & ", " & rs!ActionMemo & ", " & rs!Locked & ", " & _
                      " '" & IIf(IsNull(rs!LastModified), "", rs!LastModified) & "')"
    
    iPK = rs!CompKey
'    t = "SELECT tbl_Personnel_Payroll.* " & _
'        " FROM tbl_Personnel_Payroll " & _
'        " WHERE (Ctrl = '" & sCtrl & "')"
'    If rt.State = adStateOpen Then rt.Close
'    rt.Open t, ConnOmega
'    If rt.RecordCount > 0 Then
'        iPK = rt!PK
'    End If
'    rt.Close
    
    If CDbl(iPK) > 0 Then
        If CDbl(rs!Amount_Earned) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Earnings " & _
                              " (MasterKey, EarningKey, Taxable, NonTaxable, Hours) " & _
                              " VALUES (" & iPK & ", 1, " & CDbl(rs!Amount_Earned) & ", 0, " & CDbl(rs!NoHours) & ")"
        End If
        If CDbl(rs!TotalCola) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Earnings " & _
                              " (MasterKey, EarningKey, Taxable, NonTaxable, Hours) " & _
                              " VALUES (" & iPK & ", 2, " & CDbl(rs!TotalCola) & ", 0, " & CDbl(rs!ColaHours) & ")"
        End If
        
        If CDbl(rs!LH_Amount) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Earnings " & _
                              " (MasterKey, EarningKey, Taxable, NonTaxable, Hours) " & _
                              " VALUES (" & iPK & ", 4, " & CDbl(rs!LH_Amount) & ", 0, " & CDbl(rs!LH_Hours) & ")"
        End If
        If CDbl(rs!SH_Amount) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Earnings " & _
                              " (MasterKey, EarningKey, Taxable, NonTaxable, Hours) " & _
                              " VALUES (" & iPK & ", 5, " & CDbl(rs!SH_Amount) & ", 0, " & CDbl(rs!SH_Hours) & ")"
        End If
        If CDbl(rs!SL_Amount) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Earnings " & _
                              " (MasterKey, EarningKey, Taxable, NonTaxable, Hours) " & _
                              " VALUES (" & iPK & ", 6, " & CDbl(rs!SL_Amount) & ", 0, " & CDbl(rs!SL_Hours) & ")"
        End If
        If CDbl(rs!Adjustment) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Earnings " & _
                              " (MasterKey, EarningKey, Taxable, NonTaxable, Hours) " & _
                              " VALUES (" & iPK & ", 7, " & CDbl(rs!Adjustment) & ", 0, 0)"
        End If
        If CDbl(rs!Reg_OT_Amount) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Earnings " & _
                              " (MasterKey, EarningKey, Taxable, NonTaxable, Hours) " & _
                              " VALUES (" & iPK & ", 8, " & CDbl(rs!Reg_OT_Amount) & ", 0, " & CDbl(rs!Reg_OT_Hours) & ")"
        End If
        If CDbl(rs!RD_OT_Amount) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Earnings " & _
                              " (MasterKey, EarningKey, Taxable, NonTaxable, Hours) " & _
                              " VALUES (" & iPK & ", 9, " & CDbl(rs!RD_OT_Amount) & ", 0, " & CDbl(rs!RD_OT_Hours) & ")"
        End If
        If CDbl(rs!LH_OT_Amount) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Earnings " & _
                              " (MasterKey, EarningKey, Taxable, NonTaxable, Hours) " & _
                              " VALUES (" & iPK & ", 10, " & CDbl(rs!LH_OT_Amount) & ", 0, " & CDbl(rs!LH_OT_Hours) & ")"
        End If
        If CDbl(rs!SH_OT_Amount) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Earnings " & _
                              " (MasterKey, EarningKey, Taxable, NonTaxable, Hours) " & _
                              " VALUES (" & iPK & ", 11, " & CDbl(rs!SH_OT_Amount) & ", 0, " & CDbl(rs!SH_OT_Hours) & ")"
        End If
        
        
        '   Deduction
        
        If CDbl(rs!Mortuary) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Deductions " & _
                              " (MasterKey, DeductionKey, Amount) " & _
                              " VALUES (" & iPK & ", 13, " & CDbl(rs!Mortuary) & ")"
        End If
        If CDbl(rs!AR_Others) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Deductions " & _
                              " (MasterKey, DeductionKey, Amount) " & _
                              " VALUES (" & iPK & ", 14, " & CDbl(rs!AR_Others) & ")"
        End If
        If CDbl(rs!Advances) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Deductions " & _
                              " (MasterKey, DeductionKey, Amount) " & _
                              " VALUES (" & iPK & ", 15, " & CDbl(rs!Advances) & ")"
        End If
        If CDbl(rs!Shortages) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Deductions " & _
                              " (MasterKey, DeductionKey, Amount) " & _
                              " VALUES (" & iPK & ", 16, " & CDbl(rs!Shortages) & ")"
        End If
        If CDbl(rs!Uniforms) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Deductions " & _
                              " (MasterKey, DeductionKey, Amount) " & _
                              " VALUES (" & iPK & ", 17, " & CDbl(rs!Uniforms) & ")"
        End If
        If CDbl(rs!Others) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Deductions " & _
                              " (MasterKey, DeductionKey, Amount) " & _
                              " VALUES (" & iPK & ", 18, " & CDbl(rs!Others) & ")"
        End If
        
        '   Loans SSS
        If CDbl(rs!SSSLoan) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Deductions " & _
                              " (MasterKey, DeductionKey, Amount) " & _
                              " VALUES (" & iPK & ", 9, " & CDbl(rs!SSSLoan) & ")"
            
            t = "SELECT tbl_Personnel_Loans.* " & _
                " FROM tbl_Personnel_Loans " & _
                " WHERE (PK = " & rs!SSSLoan_No & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                ConnOmega.Execute "INSERT INTO tbl_Personnel_Loans_SL " & _
                                  " (EmpPK, LoanKey, LoanType, InOut, TransactionDate, Remarks, Credit, PayrollKey) " & _
                                  " VALUES (" & rs!EmpPK & ", " & rs!SSSLoan_No & ", 9, 'O', " & _
                                  " '" & FormatDateTime(rs!DateTo, vbShortDate) & "', " & _
                                  " 'Payroll Deduction', " & CDbl(rs!SSSLoan) & ", " & iPK & ")"
            End If
            rt.Close
        End If
        
        '   Loans PagIbig
        If CDbl(rs!PagIbigLoan) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Deductions " & _
                              " (MasterKey, DeductionKey, Amount) " & _
                              " VALUES (" & iPK & ", 11, " & CDbl(rs!PagIbigLoan) & ")"
            
            t = "SELECT tbl_Personnel_Loans.* " & _
                " FROM tbl_Personnel_Loans " & _
                " WHERE (PK = " & rs!PagIbigLoan_No & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                ConnOmega.Execute "INSERT INTO tbl_Personnel_Loans_SL " & _
                                  " (EmpPK, LoanKey, LoanType, InOut, TransactionDate, Remarks, Credit, PayrollKey) " & _
                                  " VALUES (" & rs!EmpPK & ", " & rs!PagIbigLoan_No & ", 11, 'O', " & _
                                  " '" & FormatDateTime(rs!DateTo, vbShortDate) & "', " & _
                                  " 'Payroll Deduction', " & CDbl(rs!PagIbigLoan) & ", " & iPK & ")"
            End If
            rt.Close
        End If
        
        '   Govt Cont
        If CDbl(rs!SSS) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Deductions " & _
                              " (MasterKey, DeductionKey, Amount) " & _
                              " VALUES (" & iPK & ", 1, " & CDbl(rs!SSS) & ")"
        End If
        If CDbl(rs!PHIC) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Deductions " & _
                              " (MasterKey, DeductionKey, Amount) " & _
                              " VALUES (" & iPK & ", 4, " & CDbl(rs!PHIC) & ")"
        End If
        If CDbl(rs!PAGIBIG) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Deductions " & _
                              " (MasterKey, DeductionKey, Amount) " & _
                              " VALUES (" & iPK & ", 6, " & CDbl(rs!PAGIBIG) & ")"
        End If
        If CDbl(rs!WithHeld) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Deductions " & _
                              " (MasterKey, DeductionKey, Amount) " & _
                              " VALUES (" & iPK & ", 8, " & CDbl(rs!WithHeld) & ")"
        End If
        ' Employer Share
        If CDbl(rs!SSS_Employer) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_EmployerShare " & _
                              " (MasterKey, DeductionKey, Amount) " & _
                              " VALUES (" & iPK & ", 2, " & CDbl(rs!SSS_Employer) & ")"
        End If
        If CDbl(rs!SSS_EC) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_EmployerShare " & _
                              " (MasterKey, DeductionKey, Amount) " & _
                              " VALUES (" & iPK & ", 3, " & CDbl(rs!SSS_EC) & ")"
        End If
        If CDbl(rs!PHIC_Employer) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_EmployerShare " & _
                              " (MasterKey, DeductionKey, Amount) " & _
                              " VALUES (" & iPK & ", 5, " & CDbl(rs!PHIC_Employer) & ")"
        End If
        If CDbl(rs!PagIbig_Employer) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_EmployerShare " & _
                              " (MasterKey, DeductionKey, Amount) " & _
                              " VALUES (" & iPK & ", 7, " & CDbl(rs!PagIbig_Employer) & ")"
        End If
    End If
    
    ConnOmega.Execute "UPDATE tbl_Personnel_Compensation SET Transfered = 1 WHERE (PK = " & rs!CompKey & ")"
    rs.MoveNext
Wend
rs.Close
Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Screen.MousePointer = vbHourglass
Dim iDiv, dMonthTo, dDayTo, dYearTo, dPayrollDate
iDiv = 2
s = "SELECT PK, DateTo, Type, Terms " & _
    " From dbo.tbl_Personnel_Compensation_Period"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    dMonthTo = CInt(Format(rs!DateTo, "mm"))
    dDayTo = CInt(Format(rs!DateTo, "dd"))
    dYearTo = CInt(Format(rs!DateTo, "yyyy"))
    If CDbl(rs!Type) = 1 Then
        If CDbl(dDayTo) = 8 Then
            dPayrollDate = DateSerial(dYearTo, dMonthTo, 15)
        Else
            dPayrollDate = DateSerial(dYearTo, dMonthTo + 1, 0)
        End If
    Else
        If CDbl(dDayTo) = 5 Then
            dPayrollDate = DateSerial(dYearTo, dMonthTo, 12)
        Else
            dPayrollDate = DateSerial(dYearTo, dMonthTo, 27)
        End If
    End If
    ConnOmega.Execute "UPDATE tbl_Personnel_Compensation_Period " & _
                      " SET PayrollDate = '" & FormatDateTime(dPayrollDate, vbShortDate) & "' " & _
                      " WHERE (PK = " & rs!PK & ")"
    rs.MoveNext
Wend
rs.Close
Screen.MousePointer = vbDefault
End Sub

Private Sub Command3_Click()
Screen.MousePointer = vbHourglass
Dim locLoanType
s = "SELECT tbl_Personnel_Loans.* " & _
    " FROM tbl_Personnel_Loans "
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    sCtrl = ""
    t = "SELECT TOP (1) Ctrl " & _
        " FROM tbl_Personnel_Loans " & _
        " WHERE (Year(DateGranted) = " & Format(rs!DateGranted, "yyyy") & ") " & _
        " AND (Ctrl <> '') " & _
        " ORDER BY Ctrl DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        sCtrl = Format(CDbl(rt!Ctrl) + 1, "0000000#")
    Else
        sCtrl = Format(rs!DateGranted, "yyyy") & "0000"
    End If
    rt.Close
        
    locLoanType = IIf(rs!LoanType = 1, 9, IIf(rs!LoanType = 2, 11, 0))
        
    ConnOmega.Execute "UPDATE tbl_Personnel_Loans " & _
                      " SET Ctrl = '" & sCtrl & "', " & _
                      " LoanType = " & locLoanType & " " & _
                      " WHERE (PK = " & rs!PK & ")"
    rs.MoveNext
Wend
rs.Close

Screen.MousePointer = vbDefault
End Sub

Private Sub Command4_Click()
Screen.MousePointer = vbHourglass
Dim locLoanType
s = "SELECT tbl_Personnel_Loans.* " & _
    " FROM tbl_Personnel_Loans " & _
    " WHERE (Posted = 0)"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    
    'locLoanType = IIf(rs!LoanType = 1, 9, IIf(rs!LoanType = 2, 11, 0))
    
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Loans_SL " & _
                      " (EmpPK, LoanKey, LoanType, InOut, TransactionDate, Remarks, Debit) " & _
                      " VALUES (" & rs!EmpPK & ", " & rs!PK & ", " & rs!LoanType & ", " & _
                      " 'I', '" & FormatDateTime(rs!DateGranted, vbShortDate) & "', '', " & _
                      " " & CDbl(rs!TotalAmount) & ")"
    
    ConnOmega.Execute "UPDATE tbl_Personnel_Loans " & _
                      " SET Posted = 1 " & _
                      " WHERE (PK = " & rs!PK & ")"
    rs.MoveNext
Wend
rs.Close
Screen.MousePointer = vbDefault
End Sub

Private Sub Command5_Click()
Screen.MousePointer = vbHourglass
s = "SELECT tbl_Personnel_Action.* " & _
    " FROM tbl_Personnel_Action " & _
    " WHERE (Transfered = 0)"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF

    ConnOmega.Execute "DELETE tbl_Personnel_ActionNew WHERE (PK = " & rs!PK & ")"

    t = "SELECT tbl_Personnel_IDNumber.* " & _
        " FROM tbl_Personnel_IDNumber " & _
        " WHERE (PK = " & rs!EmpPK & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        sCtrl = ""
        u = "SELECT TOP (1) CntrlNo " & _
            " FROM tbl_Personnel_ActionNew " & _
            " WHERE (Year(EffectivityDate) = " & Format(rs!EffectivityDate, "yyyy") & ") " & _
            " ORDER BY CntrlNo DESC"
        If ru.State = adStateOpen Then ru.Close
        ru.Open u, ConnOmega
        If ru.RecordCount = 0 Then
            sCtrl = Format(rs!EffectivityDate, "yyyy") & "0000"
        Else
            sCtrl = Format(CDbl(ru!CntrlNo) + 1, "0000000#")
        End If
        ru.Close
        
        Do
            u = "SELECT tbl_Personnel_ActionNew.* " & _
                " FROM tbl_Personnel_ActionNew " & _
                " WHERE (CntrlNo = '" & sCtrl & "') "
            If ru.State = adStateOpen Then ru.Close
            ru.Open u, ConnOmega
            If ru.RecordCount > 0 Then
                sCtrl = Format(CDbl(sCtrl) + 1, "0000000#")
                ru.Close
            Else
                ru.Close
                Exit Do
            End If
        Loop
        
        ConnOmega.Execute "INSERT INTO tbl_Personnel_ActionNew " & _
                          " (PK, CntrlNo, EmpPK, DivisionKey, DeptKey, EmpStatusKey, TaxStatusKey, PositionsKey, CompensationRateKey, " & _
                          " TaxCategoryKey, LoanDeductionKey, GovtDeductionKey, Is_SSS, SSS, Is_PHIC, PHIC, Is_PAGIBIG, PAGIBIG, Is_TIN, " & _
                          " TIN, EffectivityDate, Remarks, LastModified, Locked) " & _
                          " VALUES (" & rs!PK & ", '" & sCtrl & "', " & rs!EmpPK & ", " & rs!Division & ", " & rs!Dept & ", " & rs!EmpStatus & ", " & _
                          " " & rs!TaxStatus & ", " & rs!Positions & ", " & IIf(rs!CompensationRate = 1, 3, IIf(rs!CompensationRate = 2, 1, 0)) & ", " & _
                          " 4, 2, 3, " & rs!Is_SSS & ", '" & rs!SSS & "', " & rs!Is_PHIC & ", '" & rs!PHIC & "', " & rs!Is_PAGIBIG & ", '" & rs!PAGIBIG & "', " & _
                          " " & rs!Is_TIN & ", '" & rs!TIN & "', '" & FormatDateTime(rs!EffectivityDate, vbShortDate) & "', '" & FORMATSQL(rs!Remarks) & "', " & _
                          " '" & IIf(IsNull(rs!LastModified), "", rs!LastModified) & "', " & rs!Locked & ")"
        
        If CDbl(rs!Basic) > 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_ActionNew_Rate " & _
                              " (MasterKey, EarningKey, Rate, RatePerHour) " & _
                              " VALUES (" & rs!PK & ", 1, " & CDbl(rs!Basic) & ", " & CDbl(rs!RatePerHourBasic) & ")"
        End If
        If CDbl(rs!Cola) > 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_ActionNew_Rate " & _
                              " (MasterKey, EarningKey, Rate, RatePerHour) " & _
                              " VALUES (" & rs!PK & ", 2, " & CDbl(rs!Cola) & ", " & CDbl(rs!RatePerHourCola) & ")"
        End If
        If CDbl(rs!Allowance) > 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_ActionNew_Rate " & _
                              " (MasterKey, EarningKey, Rate, RatePerHour) " & _
                              " VALUES (" & rs!PK & ", 3, " & CDbl(rs!Allowance) & ", " & CDbl(rs!RatePerHourAllow) & ")"
        End If
    End If
    rt.Close
    
    ConnOmega.Execute "UPDATE tbl_Personnel_Action SET Transfered = 1 WHERE (PK = " & rs!PK & ")"
    
    rs.MoveNext
Wend
rs.Close
Screen.MousePointer = vbDefault
End Sub

Private Sub Command6_Click()
Screen.MousePointer = vbHourglass
Dim iMasterKey
s = "SELECT PK, DateApplied, LastModified " & _
    " FROM  dbo.tbl_Absent_Employee " & _
    " WHERE (Transferred = 0)"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    t = "SELECT EmpKey, AbsType, Hours, Minutes," & _
        " (SELECT TOP (1) DivisionKey " & _
        " From dbo.tbl_Personnel_ActionNew " & _
        " WHERE (EmpPK = dbo.tbl_Absent_Employee_Detail.EmpKey) " & _
        " AND (EffectivityDate <= '" & FormatDateTime(rs!DateApplied, vbShortDate) & "') " & _
        " ORDER BY EffectivityDate DESC) AS DivKey " & _
        " From dbo.tbl_Absent_Employee_Detail " & _
        " WHERE (MasterKey = " & rs!PK & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        iMasterKey = 0
        u = "SELECT PK, DateApplied, DivisionKey " & _
            " From dbo.tbl_Personnel_AbsentLateUndertime " & _
            " WHERE (DateApplied = '" & FormatDateTime(rs!DateApplied, vbShortDate) & "') " & _
            " AND (DivisionKey = " & rt!DivKey & ")"
        If ru.State = adStateOpen Then ru.Close
        ru.Open u, ConnOmega
        If ru.RecordCount = 0 Then
            sCtrl = ""
            a = "SELECT TOP (1) Ctrl " & _
                " FROM tbl_Personnel_AbsentLateUndertime " & _
                " WHERE (Year(DateApplied) = " & Format(FormatDateTime(rs!DateApplied, vbShortDate), "yyyy") & ") " & _
                " ORDER BY Ctrl DESC"
            If ra.State = adStateOpen Then ra.Close
            ra.Open a, ConnOmega
            If ra.RecordCount > 0 Then
                sCtrl = Format(CDbl(ra!Ctrl) + 1, "0000000#")
            Else
                sCtrl = Format(FormatDateTime(rs!DateApplied, vbShortDate), "yyyy") & "0000"
            End If
            ra.Close
            
            Do
                a = "SELECT tbl_Personnel_AbsentLateUndertime.* " & _
                    " FROM tbl_Personnel_AbsentLateUndertime " & _
                    " WHERE (Ctrl = '" & sCtrl & "') " & _
                    " ORDER BY Ctrl DESC"
                If ra.State = adStateOpen Then ra.Close
                ra.Open a, ConnOmega
                If ra.RecordCount = 0 Then
                    ra.Close
                    Exit Do
                End If
                ra.Close
                sCtrl = Format(CDbl(sCtrl) + 1, "0000000#")
            Loop
            
            ConnOmega.Execute "INSERT INTO tbl_Personnel_AbsentLateUndertime " & _
                              " (Ctrl, DateApplied, DivisionKey, Posted, LastModified) " & _
                              " VALUES ('" & sCtrl & "', '" & FormatDateTime(rs!DateApplied, vbShortDate) & "',  " & _
                              " " & rt!DivKey & ", 1, '" & IIf(IsNull(rs!LastModified), "", rs!LastModified) & "')"
                              
            a = "SELECT tbl_Personnel_AbsentLateUndertime.* " & _
                " FROM tbl_Personnel_AbsentLateUndertime " & _
                " WHERE (Ctrl = '" & sCtrl & "') " & _
                " ORDER BY Ctrl DESC"
            If ra.State = adStateOpen Then ra.Close
            ra.Open a, ConnOmega
            If ra.RecordCount > 0 Then
                iMasterKey = ra!PK
            End If
            ra.Close
        Else
            iMasterKey = ru!PK
        End If
        ru.Close
        
        If CDbl(iMasterKey) <> 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_AbsentLateUndertime_Details " & _
                              " (MasterKey, EmployeeKey, AbsType, Hours, Minutes) " & _
                              " VALUES (" & iMasterKey & ", " & rt!EmpKey & ", " & rt!AbsType & ", " & _
                              " " & rt!Hours & ", " & rt!Minutes & ")"
        End If
        
        rt.MoveNext
    Wend
    rt.Close
    
    ConnOmega.Execute "UPDATE tbl_Absent_Employee SET Transferred = 1 WHERE (PK = " & rs!PK & ")"
    
    rs.MoveNext
Wend
rs.Close
Screen.MousePointer = vbDefault
End Sub


Private Sub Command8_Click()
Screen.MousePointer = vbHourglass
s = "SELECT PK, SSSLoan_No, PagIbigLoan_No, LoanKeyTrans " & _
    " From dbo.tbl_Personnel_Compensation " & _
    " WHERE (LoanKeyTrans = 0) AND (PagIbigLoan_No <> 0) " & _
    " OR (LoanKeyTrans = 0) AND (SSSLoan_No <> 0)"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    
    If CDbl(rs!SSSLoan_No) <> 0 Then
        ConnOmega.Execute "UPDATE tbl_Personnel_Payroll_Deductions " & _
                          " SET LoanKey = " & rs!SSSLoan_No & " " & _
                          " WHERE (MasterKey = " & rs!PK & ") " & _
                          " AND (DeductionKey = 9)"
    End If
    If CDbl(rs!PagIbigLoan_No) <> 0 Then
        ConnOmega.Execute "UPDATE tbl_Personnel_Payroll_Deductions " & _
                          " SET LoanKey = " & rs!PagIbigLoan_No & " " & _
                          " WHERE (MasterKey = " & rs!PK & ") " & _
                          " AND (DeductionKey = 11)"
    End If
    
    ConnOmega.Execute "UPDATE tbl_Personnel_Compensation " & _
                      " SET LoanKeyTrans = 1 " & _
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
    Case vbKeyF8:       PRESS_F8
    Case vbKeyF9:
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "PersonnelHours", "PersonnelHours", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "PersonnelHours", "PersonnelHours", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "PersonnelHours", "PersonnelHours", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "PersonnelHours", "PersonnelHours", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 3
Me.Left = (MainForm.ScaleWidth - Me.Width) / 3
POPULATE_COMBO "PK", "Description", "tbl_Personnel_Division", "Description", cmbDivisionAdd
POPULATE_COMBO "PK", "Description", "tbl_Personnel_Division", "Description", cmbDivisionBatchPost
With cmbPostUnpost
    .Clear
    .AddItem "Post"
    .ItemData(.NewIndex) = 1
    .AddItem "Unpost"
    .ItemData(.NewIndex) = 2
End With
POPULATE_COMBO_EXEMPTION "PK", "HoursDesc", "tbl_Personnel_Payroll_Earnings_Table", "Sorting", "ViewInHours", 1, cmRegularHours
POPULATE_COMBO_EXEMPTION "PK", "HoursDesc", "tbl_Personnel_Payroll_Earnings_Table", "Sorting", "ViewInHours", 2, cmbOvertimeHours

isFocusReg = 0
iRowReg = 0
isFocusOT = 0
iRowOT = 0

CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANS_DETAIL = is_DET_REFRESH
BROWSER GetSetting(App.EXEName, "PersonnelHours", "PersonnelHours", ""), "is_LOAD"
If Trim(txtControl.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelHours", "PersonnelHours", ""), "is_HOME"

tmp = SetWindowLong(txtSearchAdd.hwnd, GWL_STYLE, GetWindowLong(txtSearchAdd.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearchSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearchSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtAdjustmentRem.hwnd, GWL_STYLE, GetWindowLong(txtAdjustmentRem.hwnd, GWL_STYLE) Or ES_UPPERCASE)

End Sub


Private Sub Form_Unload(Cancel As Integer)
If picAdd.Visible = True Then Cancel = -1
If picBatchPosting.Visible = True Then Cancel = -1
If picSLRegularHours.Visible = True Then Cancel = -1
If picSLOvertimeHours.Visible = True Then Cancel = -1
If picProgressBar.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstOvertime_Click()
isFocusOT = 1
iRowOT = lstOvertime.SelectedItem.Index
If imgPosted.Visible = True Then Exit Sub
With lstOvertime.ListItems
    If TRANSACTIONTYPE = is_ADDING Or _
    TRANSACTIONTYPE = is_EDITTING Then
        If CDbl(.Item(iRowOT).Text) = 0 Then TOOLBARFUNC 4 Else TOOLBARFUNC 5
        TRANS_DETAIL = is_DET_REFRESH
    End If
End With
End Sub

Private Sub lstOvertime_GotFocus()
isFocusOT = 1
iRowOT = lstOvertime.SelectedItem.Index
If imgPosted.Visible = True Then Exit Sub
With lstOvertime.ListItems
    If TRANSACTIONTYPE = is_ADDING Or _
    TRANSACTIONTYPE = is_EDITTING Then
        If CDbl(.Item(iRowOT).Text) = 0 Then TOOLBARFUNC 4 Else TOOLBARFUNC 5
        TRANS_DETAIL = is_DET_REFRESH
    End If
End With
End Sub

Private Sub lstOvertime_ItemClick(ByVal Item As MSComctlLib.ListItem)
iRowOT = lstOvertime.SelectedItem.Index
End Sub

Private Sub lstOvertime_LostFocus()
isFocusOT = 0
End Sub

Private Sub lstRegularHours_Click()
isFocusReg = 1
iRowReg = lstRegularHours.SelectedItem.Index
TRANS_DETAIL = is_DET_REFRESH
If imgPosted.Visible = True Then Exit Sub
With lstRegularHours.ListItems
    If TRANSACTIONTYPE = is_ADDING Or _
    TRANSACTIONTYPE = is_EDITTING Then
        If CDbl(.Item(iRowReg).Text) = 0 Then
            TOOLBARFUNC 4
        Else
            If CDbl(.Item(iRowReg).Text) = 1 Then
                TOOLBARFUNC 4
            Else
                TOOLBARFUNC 5
            End If
        End If
    End If
End With
End Sub

Private Sub lstRegularHours_GotFocus()
isFocusReg = 1
iRowReg = lstRegularHours.SelectedItem.Index
TRANS_DETAIL = is_DET_REFRESH
If imgPosted.Visible = True Then Exit Sub
With lstRegularHours.ListItems
    If TRANSACTIONTYPE = is_ADDING Or _
    TRANSACTIONTYPE = is_EDITTING Then
        If CDbl(.Item(iRowReg).Text) = 0 Then
            TOOLBARFUNC 4
        Else
            If CDbl(.Item(iRowReg).Text) = 1 Then
                TOOLBARFUNC 4
            Else
                TOOLBARFUNC 5
            End If
        End If
    End If
End With
End Sub

Private Sub lstRegularHours_ItemClick(ByVal Item As MSComctlLib.ListItem)
iRowReg = lstRegularHours.SelectedItem.Index
End Sub

Private Sub lstRegularHours_LostFocus()
isFocusReg = 0
End Sub

Private Sub lstResultAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAdd_Click
End Sub

Private Sub lstResultSearch_Click()
If lstResultSearch.ListIndex = -1 Then cmbPayrollPeriodSearch.Clear: Exit Sub
cmbPayrollPeriodSearch.Clear
t = "SELECT dbo.tbl_Personnel_Hours.PK, dbo.tbl_Personnel_Compensation_Period.DateFrom, " & _
    " dbo.tbl_Personnel_Compensation_Period.DateTo, dbo.tbl_Personnel_Division.Description AS Division " & _
    " FROM  dbo.tbl_Personnel_Hours LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Hours.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Hours.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_ActionNew.DivisionKey = dbo.tbl_Personnel_Division.PK " & _
    " Where (dbo.tbl_Personnel_Hours.EmployeeKey = " & lstResultSearch.ItemData(lstResultSearch.ListIndex) & ") " & _
    " ORDER BY dbo.tbl_Personnel_Compensation_Period.DateFrom DESC, dbo.tbl_Personnel_Compensation_Period.DateTo DESC"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
While Not rt.EOF
    cmbPayrollPeriodSearch.AddItem Format(rt!DateFrom, "mm/dd/yyyy") & "-" & Format(rt!DateTo, "mm/dd/yyyy") & " [" & rt!Division & "]"
    cmbPayrollPeriodSearch.ItemData(cmbPayrollPeriodSearch.NewIndex) = rt!PK
    rt.MoveNext
Wend
rt.Close
If cmbPayrollPeriodSearch.ListCount Then cmbPayrollPeriodSearch.ListIndex = 0
End Sub

Private Sub lstResultSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbPayrollPeriodSearch.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "PersonnelHours", "PersonnelHours", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "PersonnelHours", "PersonnelHours", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "PersonnelHours", "PersonnelHours", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "PersonnelHours", "PersonnelHours", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Post":    PRESS_F8
    Case "Refresh":
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtAdjustment_GotFocus()
HTEXT txtAdjustment
End Sub

Private Sub txtAdjustment_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtAdjustmentRem_GotFocus()
HTEXT txtAdjustmentRem
End Sub

Private Sub txtEarningOTKey_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstOvertime.ListItems
        .Item(iRowOT).Text = txtEarningOTKey.Text
    End With
End If
End Sub

Private Sub txtEarningRegKey_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstRegularHours.ListItems
        .Item(iRowReg).Text = txtEarningRegKey.Text
    End With
End If
End Sub

Private Sub txtFrom_GotFocus()
HTEXT txtFrom
End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtTo.SetFocus
End Sub

Private Sub txtFrom_LostFocus()
If IsDate(txtFrom.Text) = True Then txtFrom.Text = Format(FormatDateTime(txtFrom.Text, vbShortDate), "mm/dd/yyyy")
End Sub

Private Sub txtOTValue_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstOvertime.ListItems
        .Item(iRowOT).SubItems(2) = Format(RETURNTEXTVALUE(txtOTValue), "#0.00")
    End With
End If
End Sub

Private Sub txtOTValue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    picSLOvertimeHours.Visible = False
    picMain.Enabled = True
    picToolbar.Enabled = True
    lstOvertime.SetFocus
End If
End Sub

Private Sub txtOTValue_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtPayrollDateAdd_GotFocus()
HTEXT txtPayrollDateAdd
End Sub

Private Sub txtPayrollDateAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtSearchAdd.SetFocus
End Sub

Private Sub txtPayrollDateAdd_LostFocus()
If IsDate(txtPayrollDateAdd.Text) = True Then txtPayrollDateAdd.Text = Format(FormatDateTime(txtPayrollDateAdd.Text, vbShortDate), "mm/dd/yyyy")

End Sub

Private Sub txtPayrollDatePostUnpost_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKPostUnpost_Click
End Sub

Private Sub txtRegValue_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstRegularHours.ListItems
        If CheckDeductToReference(txtEarningRegKey.Text) = 1 Then
            For i = 1 To .Count
                If CLng(.Item(i).Text) = GetReferenceEarning(txtEarningRegKey.Text) Then
                    .Item(i).SubItems(2) = dPerfectHoursMaster - RETURNTEXTVALUE(txtRegValue)
                    Exit For
                End If
            Next i
        End If
        .Item(iRowReg).SubItems(2) = IIf(RETURNTEXTVALUE(txtEarningRegKey) = 1, RETURNTEXTVALUE(txtRegValue), Format(RETURNTEXTVALUE(txtRegValue), "#0.00"))
    End With
End If
End Sub

Private Sub txtRegValue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    picSLRegularHours.Visible = False
    picMain.Enabled = True
    picToolbar.Enabled = True
    lstRegularHours.SetFocus
End If
End Sub

Private Sub txtRegValue_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtSearchAdd_Change()
If cmbDivisionAdd.ListIndex = -1 Then lstResultAdd.Clear:  Exit Sub
If Trim(txtSearchAdd.Text) = "" Then lstResultAdd.Clear:  Exit Sub
lstResultAdd.Clear
If IsDate(txtPayrollDateAdd.Text) = False Then dDate = FormatDateTime(Date, vbShortDate) Else dDate = FormatDateTime(txtPayrollDateAdd.Text, vbShortDate)
t = "SELECT dbo.tbl_Personnel_IDNumber.PK, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, " & _
    " dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName " & _
    " FROM  dbo.tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
    " WHERE (ISNULL((SELECT TOP (1) tbl_Personnel_EmploymentStatus_1.Active " & _
    " FROM  dbo.tbl_Personnel_ActionNew AS tbl_Personnel_ActionNew_1 LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_EmploymentStatus AS tbl_Personnel_EmploymentStatus_1 ON tbl_Personnel_ActionNew_1.EmpStatusKey = tbl_Personnel_EmploymentStatus_1.PK " & _
    " WHERE (tbl_Personnel_ActionNew_1.EmpPK = dbo.tbl_Personnel_IDNumber.PK) " & _
    " AND (tbl_Personnel_ActionNew_1.EffectivityDate <= '" & FormatDateTime(dDate, vbShortDate) & "') " & _
    " ORDER BY tbl_Personnel_ActionNew_1.EffectivityDate DESC), 0) = 1) " & _
    " AND (ISNULL((SELECT TOP (1) DivisionKey " & _
    " FROM  dbo.tbl_Personnel_ActionNew AS tbl_Personnel_ActionNew_2 " & _
    " WHERE (EmpPK = dbo.tbl_Personnel_IDNumber.PK) " & _
    " AND (EffectivityDate <= '" & FormatDateTime(dDate, vbShortDate) & "') " & _
    " ORDER BY EffectivityDate DESC), 0) = " & cmbDivisionAdd.ItemData(cmbDivisionAdd.ListIndex) & ") " & _
    " AND (dbo.tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%') " & _
    " ORDER BY dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_IDNumber.IDNumber"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
While Not rt.EOF
    lstResultAdd.AddItem rt!IDNumber & " - " & rt!LastName & ",  " & rt!FirstName & "  " & rt!MiddleName
    lstResultAdd.ItemData(lstResultAdd.NewIndex) = rt!PK
    rt.MoveNext
Wend
rt.Close
If lstResultAdd.ListCount Then lstResultAdd.ListIndex = 0
End Sub

Private Sub txtSearchAdd_GotFocus()
HTEXT txtSearchAdd
End Sub

Private Sub txtSearchAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultAdd.SetFocus
End Sub

Private Sub txtSearchSearch_Change()
If Trim(txtSearchSearch.Text) = "" Then lstResultSearch.Clear: cmbPayrollPeriodSearch.Clear: Exit Sub
lstResultSearch.Clear: cmbPayrollPeriodSearch.Clear
s = "SELECT dbo.tbl_Personnel_Hours.EmployeeKey, dbo.tbl_Personnel_IDNumber.IDNumber, " & _
    " dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
    " dbo.tbl_Personnel_Information.MiddleName " & _
    " FROM  dbo.tbl_Personnel_Hours LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Hours.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
    " GROUP BY dbo.tbl_Personnel_Hours.EmployeeKey, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName " & _
    " HAVING (dbo.tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearchSearch.Text)) & "%') " & _
    " ORDER BY dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_IDNumber.IDNumber"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultSearch.AddItem rs!IDNumber & " - " & rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
    lstResultSearch.ItemData(lstResultSearch.NewIndex) = rs!EmployeeKey
    rs.MoveNext
Wend
rs.Close
If lstResultSearch.ListCount Then lstResultSearch.ListIndex = 0
End Sub

Private Sub txtSearchSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultSearch.SetFocus
End Sub

Private Sub txtTo_GotFocus()
HTEXT txtTo
End Sub

Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtSearchAdd.SetFocus
End Sub

Private Sub txtTo_LostFocus()
If IsDate(txtTo.Text) = True Then txtTo.Text = Format(FormatDateTime(txtTo.Text, vbShortDate), "mm/dd/yyyy")
End Sub
