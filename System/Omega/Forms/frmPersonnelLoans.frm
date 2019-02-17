VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelLoans 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPersonnelLoans.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   11445
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picBody 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   2520
      ScaleHeight     =   4335
      ScaleWidth      =   5895
      TabIndex        =   0
      Top             =   1320
      Width           =   5895
      Begin VB.ComboBox cmbInterest 
         Height          =   315
         ItemData        =   "frmPersonnelLoans.frx":08CA
         Left            =   1440
         List            =   "frmPersonnelLoans.frx":08CC
         TabIndex        =   54
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtControl 
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox cmbLoanType 
         Height          =   315
         ItemData        =   "frmPersonnelLoans.frx":08CE
         Left            =   1440
         List            =   "frmPersonnelLoans.frx":08D0
         TabIndex        =   51
         Top             =   720
         Width           =   4455
      End
      Begin VB.TextBox txtBalance 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox txtTotalPaid 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox txtDateTo 
         Height          =   315
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1560
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtDateFrom 
         Height          =   315
         Left            =   1440
         TabIndex        =   14
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox txtAmortization 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox txtNoMonths 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox txtTotalAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox txtInterest 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Top             =   2160
         Width           =   1575
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00C6B8A4&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1440
         ScaleHeight     =   255
         ScaleWidth      =   1935
         TabIndex        =   9
         Top             =   1875
         Width           =   1935
      End
      Begin VB.TextBox txtLoanAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtDateGranted 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C6B8A4&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1440
         ScaleHeight     =   255
         ScaleWidth      =   2415
         TabIndex        =   6
         Top             =   780
         Width           =   2415
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   4455
      End
      Begin VB.TextBox txtID 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   0
         Width           =   1575
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00C6B8A4&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1440
         ScaleHeight     =   255
         ScaleWidth      =   1575
         TabIndex        =   2
         Top             =   3960
         Width           =   1575
         Begin VB.CheckBox chkZeroOut 
            BackColor       =   &H00C6B8A4&
            Caption         =   "ZERO OUT"
            Height          =   195
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   1155
         End
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "BALANCE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3240
         TabIndex        =   32
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL PAID"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3240
         TabIndex        =   31
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "ZERO OUT"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE TO"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4800
         TabIndex        =   29
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE START"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "AMORTIZATION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "NO. OF MONTHS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL AMOUNT"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "INTEREST"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "INTEREST TYPE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "LOAN AMOUNT"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE GRANTED"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "LOAN TYPE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "NAME"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ID NUMBER"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   0
         Width           =   975
      End
   End
   Begin RPVGCC.b8Container picSearchAdd 
      Height          =   4575
      Left            =   3720
      TabIndex        =   33
      Top             =   1200
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   8070
      BackColor       =   15396057
      Begin VB.ListBox lstResultAdd 
         Height          =   2985
         Left            =   120
         TabIndex        =   38
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox txtSearchAdd 
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   480
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
         Picture         =   "frmPersonnelLoans.frx":08D2
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3930
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
         Picture         =   "frmPersonnelLoans.frx":102E
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   3930
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   40
         TabIndex        =   34
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
         Icon            =   "frmPersonnelLoans.frx":16A0
         ShadowVisible   =   0   'False
      End
   End
   Begin RPVGCC.b8Container picSearch 
      Height          =   4815
      Left            =   2880
      TabIndex        =   39
      Top             =   1080
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   8493
      BackColor       =   15396057
      Begin VB.ComboBox cmbPeriod 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   3720
         Width           =   5295
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
         Left            =   1080
         Picture         =   "frmPersonnelLoans.frx":1C3A
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   4170
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
         Left            =   2760
         Picture         =   "frmPersonnelLoans.frx":22AC
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   4170
         Width           =   1560
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   5295
      End
      Begin VB.ListBox lstResult 
         Height          =   2790
         Left            =   120
         TabIndex        =   40
         Top             =   840
         Width           =   5295
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   44
         Top             =   45
         Width           =   5445
         _ExtentX        =   9604
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
         Icon            =   "frmPersonnelLoans.frx":2A08
         ShadowVisible   =   0   'False
      End
      Begin VB.Label lblList 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1080
         TabIndex        =   46
         Top             =   4050
         Width           =   2775
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   9120
      TabIndex        =   52
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   48
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   49
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
         MouseIcon       =   "frmPersonnelLoans.frx":2FA2
         Begin VB.PictureBox Picture5 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   9900
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   50
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmPersonnelLoans.frx":32BC
               Top             =   120
               Visible         =   0   'False
               Width           =   1395
            End
         End
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   15000
         Y1              =   910
         Y2              =   910
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   15000
         Y1              =   90
         Y2              =   90
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   15000
         Y1              =   1005
         Y2              =   1005
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   9120
      TabIndex        =   47
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   6165
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   26458
            MinWidth        =   26458
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9840
      Top             =   1680
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
            Picture         =   "frmPersonnelLoans.frx":39CF
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelLoans.frx":46A9
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelLoans.frx":5383
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelLoans.frx":605D
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelLoans.frx":6D37
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelLoans.frx":7A11
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelLoans.frx":86EB
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelLoans.frx":93C5
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelLoans.frx":A09F
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelLoans.frx":A979
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelLoans.frx":B653
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelLoans.frx":C32D
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelLoans.frx":D007
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelLoans.frx":DCE1
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelLoans.frx":E9BB
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPersonnelLoans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public locEmployeePK As Long
Dim locLoanType As Long
Dim locInterestType As Long
Public locZeroOut As Long

Public TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2
Const is_FINDING = 3

Public isAdd_isLoan As Long

Dim tmp As Long

Dim dTotAmt, dPaidAmt, dBalAmt, iPK, iLine, dDebit, dCredit, dRunBal, sRemarks

Dim Array1, TotalAmount, Amortization, sCtrl, sLoanName_Status, iRec, iTotRec

Private Sub BROWSER(Ctrl, is_Action As String)
Select Case is_Action
    Case "is_LOAD"
        If Ctrl <> "" Then
            s = "SELECT TOP (1) dbo.tbl_Personnel_Loans.PK, dbo.tbl_Personnel_Loans.Ctrl, dbo.tbl_Personnel_Loans.EmpPK, " & _
                " dbo.tbl_Personnel_Loans.LoanType, dbo.tbl_Personnel_Loans.DateGranted, dbo.tbl_Personnel_Loans.LoanAmount, " & _
                " dbo.tbl_Personnel_Loans.InterestType, dbo.tbl_Personnel_Loans.Interest, dbo.tbl_Personnel_Loans.TotalAmount, " & _
                " dbo.tbl_Personnel_Loans.Amortization, dbo.tbl_Personnel_Loans.NoMonths, dbo.tbl_Personnel_Loans.DateFrom, " & _
                " dbo.tbl_Personnel_Loans.DateTo, dbo.tbl_Personnel_Loans.ZeroOut, dbo.tbl_Personnel_Loans.TotalPaid, " & _
                " dbo.tbl_Personnel_Loans.Balance, dbo.tbl_Personnel_Loans.Posted, dbo.tbl_Personnel_Loans.LastModified, " & _
                " dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
                " dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Payroll_Deductions_Table.Description AS LoanTypeDesc, " & _
                " dbo.tbl_Personnel_Loans_Interest.Interest AS InterestDesc " & _
                " FROM  dbo.tbl_Personnel_Loans LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Loans.LoanType = dbo.tbl_Personnel_Payroll_Deductions_Table.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Loans_Interest ON dbo.tbl_Personnel_Loans.InterestType = dbo.tbl_Personnel_Loans_Interest.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Loans.EmpPK = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
                " WHERE (dbo.tbl_Personnel_Loans.Ctrl = '" & Ctrl & "') " & _
                " ORDER BY dbo.tbl_Personnel_Loans.Ctrl"
        Else
            s = "SELECT TOP (1) dbo.tbl_Personnel_Loans.PK, dbo.tbl_Personnel_Loans.Ctrl, dbo.tbl_Personnel_Loans.EmpPK, " & _
                " dbo.tbl_Personnel_Loans.LoanType, dbo.tbl_Personnel_Loans.DateGranted, dbo.tbl_Personnel_Loans.LoanAmount, " & _
                " dbo.tbl_Personnel_Loans.InterestType, dbo.tbl_Personnel_Loans.Interest, dbo.tbl_Personnel_Loans.TotalAmount, " & _
                " dbo.tbl_Personnel_Loans.Amortization, dbo.tbl_Personnel_Loans.NoMonths, dbo.tbl_Personnel_Loans.DateFrom, " & _
                " dbo.tbl_Personnel_Loans.DateTo, dbo.tbl_Personnel_Loans.ZeroOut, dbo.tbl_Personnel_Loans.TotalPaid, " & _
                " dbo.tbl_Personnel_Loans.Balance, dbo.tbl_Personnel_Loans.Posted, dbo.tbl_Personnel_Loans.LastModified, " & _
                " dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
                " dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Payroll_Deductions_Table.Description AS LoanTypeDesc, " & _
                " dbo.tbl_Personnel_Loans_Interest.Interest AS InterestDesc " & _
                " FROM  dbo.tbl_Personnel_Loans LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Loans.LoanType = dbo.tbl_Personnel_Payroll_Deductions_Table.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Loans_Interest ON dbo.tbl_Personnel_Loans.InterestType = dbo.tbl_Personnel_Loans_Interest.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Loans.EmpPK = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
                " ORDER BY dbo.tbl_Personnel_Loans.Ctrl"
        End If
    Case "is_HOME"
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Loans.PK, dbo.tbl_Personnel_Loans.Ctrl, dbo.tbl_Personnel_Loans.EmpPK, " & _
            " dbo.tbl_Personnel_Loans.LoanType, dbo.tbl_Personnel_Loans.DateGranted, dbo.tbl_Personnel_Loans.LoanAmount, " & _
            " dbo.tbl_Personnel_Loans.InterestType, dbo.tbl_Personnel_Loans.Interest, dbo.tbl_Personnel_Loans.TotalAmount, " & _
            " dbo.tbl_Personnel_Loans.Amortization, dbo.tbl_Personnel_Loans.NoMonths, dbo.tbl_Personnel_Loans.DateFrom, " & _
            " dbo.tbl_Personnel_Loans.DateTo, dbo.tbl_Personnel_Loans.ZeroOut, dbo.tbl_Personnel_Loans.TotalPaid, " & _
            " dbo.tbl_Personnel_Loans.Balance, dbo.tbl_Personnel_Loans.Posted, dbo.tbl_Personnel_Loans.LastModified, " & _
            " dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
            " dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Payroll_Deductions_Table.Description AS LoanTypeDesc, " & _
            " dbo.tbl_Personnel_Loans_Interest.Interest AS InterestDesc " & _
            " FROM  dbo.tbl_Personnel_Loans LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Loans.LoanType = dbo.tbl_Personnel_Payroll_Deductions_Table.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Loans_Interest ON dbo.tbl_Personnel_Loans.InterestType = dbo.tbl_Personnel_Loans_Interest.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Loans.EmpPK = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
            " ORDER BY dbo.tbl_Personnel_Loans.Ctrl"
    Case "is_PAGEUP"
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Loans.PK, dbo.tbl_Personnel_Loans.Ctrl, dbo.tbl_Personnel_Loans.EmpPK, " & _
            " dbo.tbl_Personnel_Loans.LoanType, dbo.tbl_Personnel_Loans.DateGranted, dbo.tbl_Personnel_Loans.LoanAmount, " & _
            " dbo.tbl_Personnel_Loans.InterestType, dbo.tbl_Personnel_Loans.Interest, dbo.tbl_Personnel_Loans.TotalAmount, " & _
            " dbo.tbl_Personnel_Loans.Amortization, dbo.tbl_Personnel_Loans.NoMonths, dbo.tbl_Personnel_Loans.DateFrom, " & _
            " dbo.tbl_Personnel_Loans.DateTo, dbo.tbl_Personnel_Loans.ZeroOut, dbo.tbl_Personnel_Loans.TotalPaid, " & _
            " dbo.tbl_Personnel_Loans.Balance, dbo.tbl_Personnel_Loans.Posted, dbo.tbl_Personnel_Loans.LastModified, " & _
            " dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
            " dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Payroll_Deductions_Table.Description AS LoanTypeDesc, " & _
            " dbo.tbl_Personnel_Loans_Interest.Interest AS InterestDesc " & _
            " FROM  dbo.tbl_Personnel_Loans LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Loans.LoanType = dbo.tbl_Personnel_Payroll_Deductions_Table.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Loans_Interest ON dbo.tbl_Personnel_Loans.InterestType = dbo.tbl_Personnel_Loans_Interest.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Loans.EmpPK = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
            " WHERE (dbo.tbl_Personnel_Loans.Ctrl < '" & Ctrl & "') " & _
            " ORDER BY dbo.tbl_Personnel_Loans.Ctrl DESC"
    Case "is_PAGEDOWN"
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Loans.PK, dbo.tbl_Personnel_Loans.Ctrl, dbo.tbl_Personnel_Loans.EmpPK, " & _
            " dbo.tbl_Personnel_Loans.LoanType, dbo.tbl_Personnel_Loans.DateGranted, dbo.tbl_Personnel_Loans.LoanAmount, " & _
            " dbo.tbl_Personnel_Loans.InterestType, dbo.tbl_Personnel_Loans.Interest, dbo.tbl_Personnel_Loans.TotalAmount, " & _
            " dbo.tbl_Personnel_Loans.Amortization, dbo.tbl_Personnel_Loans.NoMonths, dbo.tbl_Personnel_Loans.DateFrom, " & _
            " dbo.tbl_Personnel_Loans.DateTo, dbo.tbl_Personnel_Loans.ZeroOut, dbo.tbl_Personnel_Loans.TotalPaid, " & _
            " dbo.tbl_Personnel_Loans.Balance, dbo.tbl_Personnel_Loans.Posted, dbo.tbl_Personnel_Loans.LastModified, " & _
            " dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
            " dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Payroll_Deductions_Table.Description AS LoanTypeDesc, " & _
            " dbo.tbl_Personnel_Loans_Interest.Interest AS InterestDesc " & _
            " FROM  dbo.tbl_Personnel_Loans LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Loans.LoanType = dbo.tbl_Personnel_Payroll_Deductions_Table.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Loans_Interest ON dbo.tbl_Personnel_Loans.InterestType = dbo.tbl_Personnel_Loans_Interest.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Loans.EmpPK = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
            " WHERE (dbo.tbl_Personnel_Loans.Ctrl > '" & Ctrl & "') " & _
            " ORDER BY dbo.tbl_Personnel_Loans.Ctrl"
    Case "is_END"
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Loans.PK, dbo.tbl_Personnel_Loans.Ctrl, dbo.tbl_Personnel_Loans.EmpPK, " & _
            " dbo.tbl_Personnel_Loans.LoanType, dbo.tbl_Personnel_Loans.DateGranted, dbo.tbl_Personnel_Loans.LoanAmount, " & _
            " dbo.tbl_Personnel_Loans.InterestType, dbo.tbl_Personnel_Loans.Interest, dbo.tbl_Personnel_Loans.TotalAmount, " & _
            " dbo.tbl_Personnel_Loans.Amortization, dbo.tbl_Personnel_Loans.NoMonths, dbo.tbl_Personnel_Loans.DateFrom, " & _
            " dbo.tbl_Personnel_Loans.DateTo, dbo.tbl_Personnel_Loans.ZeroOut, dbo.tbl_Personnel_Loans.TotalPaid, " & _
            " dbo.tbl_Personnel_Loans.Balance, dbo.tbl_Personnel_Loans.Posted, dbo.tbl_Personnel_Loans.LastModified, " & _
            " dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
            " dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Payroll_Deductions_Table.Description AS LoanTypeDesc, " & _
            " dbo.tbl_Personnel_Loans_Interest.Interest AS InterestDesc " & _
            " FROM  dbo.tbl_Personnel_Loans LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Loans.LoanType = dbo.tbl_Personnel_Payroll_Deductions_Table.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Loans_Interest ON dbo.tbl_Personnel_Loans.InterestType = dbo.tbl_Personnel_Loans_Interest.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Loans.EmpPK = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
            " ORDER BY dbo.tbl_Personnel_Loans.Ctrl DESC"
    Case "is_FIND"
        s = "SELECT TOP (1) dbo.tbl_Personnel_Loans.PK, dbo.tbl_Personnel_Loans.Ctrl, dbo.tbl_Personnel_Loans.EmpPK, " & _
            " dbo.tbl_Personnel_Loans.LoanType, dbo.tbl_Personnel_Loans.DateGranted, dbo.tbl_Personnel_Loans.LoanAmount, " & _
            " dbo.tbl_Personnel_Loans.InterestType, dbo.tbl_Personnel_Loans.Interest, dbo.tbl_Personnel_Loans.TotalAmount, " & _
            " dbo.tbl_Personnel_Loans.Amortization, dbo.tbl_Personnel_Loans.NoMonths, dbo.tbl_Personnel_Loans.DateFrom, " & _
            " dbo.tbl_Personnel_Loans.DateTo, dbo.tbl_Personnel_Loans.ZeroOut, dbo.tbl_Personnel_Loans.TotalPaid, " & _
            " dbo.tbl_Personnel_Loans.Balance, dbo.tbl_Personnel_Loans.Posted, dbo.tbl_Personnel_Loans.LastModified, " & _
            " dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
            " dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Payroll_Deductions_Table.Description AS LoanTypeDesc, " & _
            " dbo.tbl_Personnel_Loans_Interest.Interest AS InterestDesc " & _
            " FROM  dbo.tbl_Personnel_Loans LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Loans.LoanType = dbo.tbl_Personnel_Payroll_Deductions_Table.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Loans_Interest ON dbo.tbl_Personnel_Loans.InterestType = dbo.tbl_Personnel_Loans_Interest.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Loans.EmpPK = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
            " WHERE (dbo.tbl_Personnel_Loans.PK = " & Ctrl & ") " & _
            " ORDER BY dbo.tbl_Personnel_Loans.Ctrl DESC"
    Case Else: Exit Sub
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    locEmployeePK = rs!EmpPK
    locLoanType = rs!LoanType
    locInterestType = rs!InterestType
    locZeroOut = rs!ZeroOut
    txtControl.Text = rs!Ctrl
    txtID.Text = rs!IDNumber
    txtName.Text = rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
    txtDateGranted.Text = Format(rs!DateGranted, "mm/dd/yyyy")
    cmbLoanType.Text = rs!LoanTypeDesc
    cmbInterest.Text = rs!InterestDesc
    txtLoanAmount.Text = Format(rs!LoanAmount, "##,##0.00")
    txtInterest.Text = Format(rs!Interest, "##,##0.00")
    txtTotalAmount.Text = Format(rs!TotalAmount, "##,##0.00")
    txtAmortization.Text = Format(rs!Amortization, "##,##0.00")
    txtNoMonths.Text = rs!NoMonths
    txtDateFrom.Text = Format(rs!DateFrom, "mm/dd/yyyy")
    txtDateTo.Text = Format(rs!DateTo, "mm/dd/yyyy")
    'GET_PAID_BALANCE rs!PK, rs!LoanType
    chkZeroOut.Value = rs!ZeroOut
    
    dTotAmt = CDbl(Format(rs!TotalAmount, "##,##0.00"))
    dPaidAmt = Format(Get_Loan_Paid(rs!PK), "#,##0.00")
    dBalAmt = CDbl(dTotAmt) - CDbl(dPaidAmt)
    
    txtTotalPaid.Text = Format(dPaidAmt, "#,##0.00")
    txtBalance.Text = Format(dBalAmt, "#,##0.00")
    
    imgPosted.Visible = IIf(rs!Posted = 1, True, False)
    Toolbar1.Buttons(19).Caption = IIf(rs!Posted = 1, "UnPost", " Post ")
    Toolbar1.Buttons(19).Image = IIf(rs!Posted = 1, 11, 10)
    
    StatusBar.Panels(1).Text = rs!PK
    StatusBar.Panels(2).Text = "LAST MODIFIED BY : " & rs!LastModified
    SaveSetting App.EXEName, "LoanCtrl", "LoanCt", rs!Ctrl
End If
rs.Close
End Sub

Private Function INSERT_LOANS(strEmpPK, strEmpID, strEmpName, _
intLoanType, dtmGranted, dblLoanAmount, intInterestType, _
dblInterest, dblTotalAmount, dblAmortization, intNoMonths, _
dtmFrom, dtmTo, intZeroOut, strLastModified)
s = "INSERT INTO tbl_Personnel_Loans" & _
    " (EmpPK, EmpID, EmpName, LoanType, DateGranted, " & _
    " LoanAmount, InterestType, Interest, TotalAmount, " & _
    " Amortization, NoMonths, DateFrom, DateTo, ZeroOut, " & _
    " LastModified) " & _
    " VALUES (" & strEmpPK & ",'" & strEmpID & "', '" & strEmpName & "', " & _
    " " & intLoanType & ", '" & CDate(dtmGranted) & "', " & _
    " " & CDbl(dblLoanAmount) & ", " & CLng(intInterestType) & ", " & _
    " " & CDbl(dblInterest) & ", " & CDbl(dblTotalAmount) & ", " & _
    " " & CDbl(dblAmortization) & ", " & CLng(intNoMonths) & ", " & _
    " '" & CDate(dtmFrom) & "', '" & CDate(dtmTo) & "', " & _
    " " & CLng(intZeroOut) & ", '" & strLastModified & "')"
ConnOmega.Execute s, , -1
End Function

Private Function UPDATE_LOANS(intPK, dtmGranted, _
dblLoanAmount, intInterestType, dblInterest, _
dblTotalAmount, dblAmortization, intNoMonths, _
dtmFrom, dtmTo, intZeroOut, strLastModified, intLoanType)
s = "UPDATE tbl_Personnel_Loans" & _
    " SET DateGranted = '" & CDate(dtmGranted) & "', " & _
    " LoanType = " & intLoanType & ", " & _
    " LoanAmount = " & CDbl(dblLoanAmount) & ", " & _
    " InterestType = " & CLng(intInterestType) & ", " & _
    " Interest = " & CDbl(dblInterest) & ", " & _
    " TotalAmount = " & CDbl(dblTotalAmount) & ", " & _
    " Amortization = " & CDbl(dblAmortization) & ", " & _
    " NoMonths = " & CLng(intNoMonths) & ", " & _
    " DateFrom = '" & CDate(dtmFrom) & "', " & _
    " DateTo = '" & CDate(dtmTo) & "', " & _
    " ZeroOut = " & CLng(intZeroOut) & ", " & _
    " LastModified = '" & strLastModified & "' " & _
    " WHERE (PK = " & CLng(intPK) & ")"
ConnOmega.Execute s, , -1
End Function

Private Function CHECK_DUPLICATE(strEmpNo, intLoanType, dtmFrom, intPK) As Boolean
s = "SELECT Loans.*" & _
    " From Loans " & _
    " WHERE (EmpPK=" & strEmpNo & ") " & _
    " AND (LoanType=" & intLoanType & ") " & _
    " AND (DateFrom='" & CDate(dtmFrom) & "') " & _
    " AND (Loans.PK<>" & intPK & ")"
rs.Open s, ConnOmega
If Not rs.EOF Then
    CHECK_DUPLICATE = True
End If
End Function

Private Sub PRESS_INSERT()
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If AccessRights("Personnel Loans", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
isAdd_isLoan = 1
txtSearchAdd.Text = ""
picSearchAdd.ZOrder 0
picBody.Enabled = False
picSearchAdd.Visible = True
txtSearchAdd.SetFocus
'frmPersonnelLoansSearch.TRANSACTIONTYPE = 1
'If IsLoaded(frmPersonnelLoansSearch) Then
'    frmPersonnelLoansSearch.ZOrder 0
'Else
'    frmPersonnelLoansSearch.Show
'End If
End Sub

Private Sub PRESS_F2()
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If StatusBar.Panels(1).Text = "" Then Exit Sub
'If imgPosted.Visible = True Then MsgBox "Already Posted!             ", vbCritical, "Error...": Exit Sub
If AccessRights("Personnel Loans", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If

If imgPosted.Visible = True Then
    If MsgBox("YOU WANT TO ZERO OUT THIS LOAN?                          ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
'    Dim dBal
'    t = "SELECT ROUND(SUM(Balance), 2) AS Balance " & _
'        " From dbo.tbl_Personnel_Loans_SL " & _
'        " WHERE (LoanKey = " & StatusBar.Panels(1).Text & ") " & _
'        " AND (TransactionDate <= '" & FormatDateTime(Date, vbShortDate) & "')"
'    If rt.State = adStateOpen Then rt.Close
'    rt.Open t, ConnOmega
'    If rt.RecordCount > 0 Then
'        dBal = IIf(IsNull(rt!Balance), 0, rt!Balance)
'        If CDbl(dBal) <> 0 Then
'            ConnOmega.Execute "INSERT INTO tbl_Personnel_Loans_SL " & _
'                              " (EmpPK, LoanKey, LoanType, InOut, TransactionDate, Remarks, Credit) " & _
'                              " VALUES (" & locEmployeePK & ", " & StatusBar.Panels(1).Text & ", " & _
'                              " " & locLoanType & ", 'O', '" & FormatDateTime(Date, vbShortDate) & "', " & _
'                              " 'Zero Out', " & CDbl(dBal) & ")"
'        End If
'    End If
'    rt.Close
    
    ConnOmega.Execute "UPDATE tbl_Personnel_Loans " & _
                      " SET ZeroOut = 1 " & _
                      " WHERE (PK = " & StatusBar.Panels(1).Text & ")"
    
    BROWSER GetSetting(App.EXEName, "LoanCtrl", "LoanCt", ""), "is_LOAD"
        
    Exit Sub
End If

LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_EDITTING
End Sub

Private Function CHECK_LOANS_IN_PAYROLL(intPK, intType) As Boolean
Select Case intType
    Case 1
        s = "SELECT SSSLoan_No" & _
            " From tbl_Personnel_Compensation " & _
            " WHERE (SSSLoan_No=" & intPK & ")"
    Case 2
        s = "SELECT PagIbigLoan_No" & _
            " From tbl_Personnel_Compensation  " & _
            " WHERE (PagIbigLoan_No=" & intPK & ")"
End Select
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not rs.EOF Then
    CHECK_LOANS_IN_PAYROLL = True
End If
rs.Close
End Function

Private Function DELETE_LOANS(intPK)
s = "DELETE FROM tbl_Personnel_Loans" & _
    " WHERE (PK = " & intPK & ")"
ConnOmega.Execute s, , -1
End Function

Private Sub PRESS_DELETE()
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If StatusBar.Panels(1).Text = "" Then Exit Sub
If imgPosted.Visible = True Then MsgBox "Already Posted!             ", vbCritical, "Error...": Exit Sub
If AccessRights("Personnel Loans", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If

If MsgBox("ARE YOU SURE TO DELETE THIS RECORD?          ", vbCritical + vbYesNo + vbDefaultButton2, "CONFIRMATION") = vbNo Then Exit Sub
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_Personnel_Loans WHERE (PK =" & StatusBar.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "LoanCtrl", "LoanCt", ""), "is_PAGEDOWN"
If Trim(txtControl.Text) = "" Then BROWSER GetSetting(App.EXEName, "LoanCtrl", "LoanCt", ""), "is_HOME"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
'MsgBox locZeroOut
If IsDate(txtDateGranted.Text) = False Then MsgBox "Please supply a valid date!                 ", vbCritical, "Error...": txtDateGranted.SetFocus: Exit Sub
If locLoanType = 0 Then MsgBox "Please select Loan Type!                 ", vbCritical, "Error...": cmbLoanType.SetFocus: Exit Sub
If locInterestType = 0 Then MsgBox "Please select Interest!                 ", vbCritical, "Error...": cmbInterest.SetFocus: Exit Sub
If RETURNTEXTVALUE(txtTotalAmount) <= 0 Then MsgBox "Invalid Loan Amount!                    ", vbCritical, "Error...": txtTotalAmount.SetFocus: Exit Sub
If RETURNTEXTVALUE(txtAmortization) <= 0 Then MsgBox "Invalid Amortization Amount!                    ", vbCritical, "Error...": txtAmortization.SetFocus: Exit Sub
If RETURNTEXTVALUE(txtNoMonths) <= 0 Then MsgBox "Invalid Number of Months!                    ", vbCritical, "Error...": txtNoMonths.SetFocus: Exit Sub
If IsDate(txtDateFrom.Text) = False Then MsgBox "Please supply a valid date!                 ", vbCritical, "Error...": txtDateFrom.SetFocus: Exit Sub
If IsDate(txtDateTo.Text) = False Then MsgBox "Please supply a valid date!                 ", vbCritical, "Error...": txtDateTo.SetFocus: Exit Sub
If DateValue(CDate(txtDateFrom.Text)) > DateValue(CDate(txtDateTo.Text)) Then MsgBox "DATE TO MUST HIGHER THAN DATE FROM!         ", vbCritical, "Error...": txtDateFrom.SetFocus: Exit Sub

s = "SELECT  EmpPK, LoanType, ZeroOut" & _
    " From tbl_Personnel_Loans " & _
    " Where (EmpPK = " & locEmployeePK & ") " & _
    " And (LoanType = " & locLoanType & ") " & _
    " And (ZeroOut = 0)" & _
    " And (PK <> " & IIf(IsNumeric(StatusBar.Panels(1).Text) = False, 0, StatusBar.Panels(1).Text) & ")"
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    MsgBox "Please Zero Out Previous Loan....           ", vbCritical, "Error..."
    Exit Sub
End If
rs.Close

On Error GoTo PG:

If TRANSACTIONTYPE = is_ADDING Then
    sCtrl = ""
    t = "SELECT TOP (1) Ctrl " & _
        " FROM tbl_Personnel_Loans " & _
        " WHERE (Year(DateGranted) = " & Format(FormatDateTime(txtDateGranted.Text, vbShortDate), "yyyy") & ") " & _
        " ORDER BY Ctrl DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        sCtrl = Format(CDbl(rt!Ctrl) + 1, "0000000#")
    Else
        sCtrl = Format(FormatDateTime(txtDateGranted.Text, vbShortDate), "yyyy") & "0000"
    End If
    rt.Close
    
    Do
        t = "SELECT tbl_Personnel_Loans.* " & _
            " FROM tbl_Personnel_Loans " & _
            " WHERE (Ctrl = '" & sCtrl & "') " & _
            " ORDER BY Ctrl DESC"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount = 0 Then
            rt.Close
            Exit Do
        End If
        rt.Close
        sCtrl = Format(CDbl(sCtrl) + 1, "0000000#")
    Loop
    
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Loans " & _
                      " (EmpPK, LoanType, DateGranted, LoanAmount, InterestType, " & _
                      " Interest, TotalAmount, Amortization, NoMonths, DateFrom, " & _
                      " DateTo, LastModified, Ctrl) " & _
                      " VALUES (" & locEmployeePK & ", " & locLoanType & ", " & _
                      " '" & FormatDateTime(txtDateGranted.Text, vbShortDate) & "', " & _
                      " " & RETURNTEXTVALUE(txtLoanAmount) & ", " & locInterestType & ", " & _
                      " " & RETURNTEXTVALUE(txtInterest) & ", " & RETURNTEXTVALUE(txtTotalAmount) & ", " & _
                      " " & RETURNTEXTVALUE(txtAmortization) & ", " & RETURNTEXTVALUE(txtNoMonths) & ", " & _
                      " '" & FormatDateTime(txtDateFrom.Text, vbShortDate) & "', " & _
                      " '" & FormatDateTime(txtDateTo.Text, vbShortDate) & "', " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "', '" & sCtrl & "')"

ElseIf TRANSACTIONTYPE = is_EDITTING Then
    sCtrl = Trim(txtControl.Text)
    ConnOmega.Execute "UPDATE tbl_Personnel_Loans " & _
                      " SET LoanType = " & locLoanType & ", " & _
                      " DateGranted = '" & FormatDateTime(txtDateGranted.Text, vbShortDate) & "', " & _
                      " LoanAmount = " & RETURNTEXTVALUE(txtLoanAmount) & ", " & _
                      " InterestType = " & locInterestType & ", " & _
                      " Interest = " & RETURNTEXTVALUE(txtInterest) & ", " & _
                      " TotalAmount = " & RETURNTEXTVALUE(txtTotalAmount) & ", " & _
                      " Amortization = " & RETURNTEXTVALUE(txtAmortization) & ", " & _
                      " NoMonths = " & RETURNTEXTVALUE(txtNoMonths) & ", " & _
                      " DateFrom = '" & FormatDateTime(txtDateFrom.Text, vbShortDate) & "', " & _
                      " DateTo = '" & FormatDateTime(txtDateTo.Text, vbShortDate) & "', " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & StatusBar.Panels(1).Text & ")"
End If

CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER sCtrl, "is_LOAD"

Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
picToolbar.Enabled = False
picBody.Enabled = False
txtSearch.Text = ""
cmbLoanType.ListIndex = -1
lblList.Caption = ""
cmbPeriod.Clear
picSearch.ZOrder 0
picSearch.Visible = True
txtSearch.SetFocus
End Sub

Private Sub PRESS_F8()
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If StatusBar.Panels(1).Text = "" Then Exit Sub
If imgPosted.Visible = True Then MsgBox "Already Posted!             ", vbCritical, "Error...": Exit Sub
If AccessRights("Personnel Loans", "Post") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If MsgBox("ARE YOU SURE IN POSTING THIS TRANSACTION?                    ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
On Error GoTo PG:
s = "SELECT tbl_Personnel_Loans.* " & _
    " FROM tbl_Personnel_Loans " & _
    " WHERE (PK = " & StatusBar.Panels(1).Text & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Loans_SL " & _
                      " (EmpPK, LoanKey, LoanType, InOut, TransactionDate, Remarks, Debit) " & _
                      " VALUES (" & rs!EmpPK & ", " & rs!PK & ", " & rs!LoanType & ", " & _
                      " 'I', '" & FormatDateTime(rs!DateGranted, vbShortDate) & "', '" & FORMATSQL(cmbLoanType.Text) & "', " & _
                      " " & CDbl(rs!TotalAmount) & ")"
    
    ConnOmega.Execute "UPDATE tbl_Personnel_Loans " & _
                      " SET Posted = 1 " & _
                      " WHERE (PK = " & rs!PK & ")"
    rs.MoveNext
Wend
rs.Close
sCtrl = Trim(txtControl.Text)
CLEARTEXT
BROWSER sCtrl, "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Errir..."
Exit Sub
End Sub

Private Sub PRESS_F9()
If picSearch.Visible = True Then Exit Sub
If StatusBar.Panels(1).Text = "" Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
PopupMenu MainFormPopupF.mnuLoanRep, , picToolbar.Left + Toolbar1.Buttons(17).Left, picToolbar.Top + Toolbar1.Buttons(17).Top + Toolbar1.Buttons(17).Height
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSearchAdd.Visible = True Then cmdCancelAdd_Click: Exit Sub
    If picSearch.Visible = True Then cmdCancelSearch_Click: Exit Sub
    Unload Me
Else
    TRANSACTIONTYPE = is_REFRESH
    LOCKTEXT True
    TOOLBARFUNC 1
    BROWSER GetSetting(App.EXEName, "LoanCtrl", "LoanCt", ""), "is_LOAD"
End If
End Sub

Private Sub GET_PAID_BALANCE(intPK, intType)
Select Case intType
    Case 1
        's = "SELECT qry_SSSLoan.SumOfSSSLoan as TotalPaid, " & _
            " tbl_Personnel_Loans.TotalAmount - qry_SSSLoan.SumOfSSSLoan AS Balance" & _
            " FROM tbl_Personnel_Loans LEFT JOIN qry_SSSLoan ON tbl_Personnel_Loans.PK = qry_SSSLoan.SSSLoan_No " & _
            " WHERE (tbl_Personnel_Loans.PK = " & intPK & ") " & _
            " AND (tbl_Personnel_Loans.LoanType = 1)"
        s = "SELECT ISNULL ((SELECT SUM(tbl_Personnel_Compensation.SSSLoan) AS Payment " & _
            " From tbl_Personnel_Compensation " & _
            " WHERE (tbl_Personnel_Compensation.SSSLoan_No = tbl_Personnel_Loans.PK)), 0) AS TotalPaid, " & _
            " ROUND(TotalAmount - ISNULL((SELECT SUM(tbl_Personnel_Compensation.SSSLoan) AS Payment " & _
            " From tbl_Personnel_Compensation " & _
            " WHERE (tbl_Personnel_Compensation.SSSLoan_No = tbl_Personnel_Loans.PK)), 0), 2) AS Balance " & _
            " From tbl_Personnel_Loans " & _
            " WHERE (PK = " & intPK & ") " & _
            " AND (LoanType = 1)"
    Case 2
        's = "SELECT qry_PagIbigLoan.SumOfPagIbigLoan as TotalPaid, " & _
            " tbl_Personnel_Loans.TotalAmount - qry_PagIbigLoan.SumOfPagIbigLoan as Balance" & _
            " FROM tbl_Personnel_Loans LEFT JOIN qry_PagIbigLoan ON tbl_Personnel_Loans.PK = qry_PagIbigLoan.PagIbigLoan_No " & _
            " WHERE (tbl_Personnel_Loans.PK = " & intPK & ") " & _
            " AND (tbl_Personnel_Loans.LoanType = 2)"
        s = "SELECT ISNULL((SELECT SUM(tbl_Personnel_Compensation.PagIbigLoan) AS Payment " & _
            " From tbl_Personnel_Compensation " & _
            " WHERE (tbl_Personnel_Compensation.PagIbigLoan_No = tbl_Personnel_Loans.PK)), 0) AS TotalPaid, " & _
            " ROUND(TotalAmount - ISNULL((SELECT SUM(tbl_Personnel_Compensation.PagIbigLoan) AS Payment " & _
            " From tbl_Personnel_Compensation " & _
            " WHERE (tbl_Personnel_Compensation.PagIbigLoan_No = tbl_Personnel_Loans.PK)), 0), 2) AS Balance " & _
            " From tbl_Personnel_Loans " & _
            " WHERE (PK = " & intPK & ") " & _
            " AND (LoanType = 2)"
End Select
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    txtTotalPaid.Text = Format(IIf(IsNull(ra!TotalPaid), 0, ra!TotalPaid), "#,##0.00")
    txtBalance.Text = Format(IIf(IsNull(ra!Balance), 0, ra!Balance), "#,##0.00")
Else
    txtTotalPaid.Text = "0.00"
    txtBalance.Text = "0.00"
End If
ra.Close
End Sub

Public Sub CLEARTEXT()
locEmployeePK = 0
locLoanType = 0
locInterestType = 0
locZeroOut = 0
txtControl.Text = ""
txtID.Text = ""
txtName.Text = ""
txtDateGranted.Text = ""
txtLoanAmount.Text = ""
txtInterest.Text = ""
txtTotalAmount.Text = ""
txtAmortization.Text = ""
txtNoMonths.Text = ""
txtDateFrom.Text = ""
txtDateTo.Text = ""
txtTotalPaid.Text = ""
txtBalance.Text = ""
cmbLoanType.Text = ""
cmbLoanType.ListIndex = -1
cmbInterest.Text = ""
cmbInterest.ListIndex = -1
'chkSSS.Value = 0
'chkPagIbig.Value = 0
'chkPercent.Value = 0
'chkAmount.Value = 0
chkZeroOut.Value = 0
StatusBar.Panels(1).Text = ""
StatusBar.Panels(2).Text = ""
imgPosted.Visible = False
End Sub

Public Sub LOCKTEXT(bln As Boolean)
txtDateGranted.Locked = bln
txtLoanAmount.Locked = bln
txtInterest.Locked = bln
txtAmortization.Locked = bln
txtDateFrom.Locked = bln
cmbLoanType.Locked = bln
cmbInterest.Locked = bln
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
cmdCancelAdd_Click
End Sub

Private Sub b8TitleBar2_CLoseClick()
cmdCancelSearch_Click
End Sub

Private Sub chkZeroOut_Click()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    locZeroOut = chkZeroOut.Value
Else
    chkZeroOut.Value = locZeroOut
End If
End Sub

Private Sub chkZeroOut_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtTotalPaid.SetFocus
End If
End Sub

Private Sub cmbInterest_Click()
If cmbInterest.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    locInterestType = cmbInterest.ItemData(cmbInterest.ListIndex)
End If
End Sub

Private Sub cmbInterest_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtInterest.SetFocus
End Sub

Private Sub cmbLoanType_Click()
If cmbLoanType.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If TRANSACTIONTYPE = is_ADDING Then
        s = "SELECT EmpPK, LoanType, ZeroOut" & _
            " From tbl_Personnel_Loans " & _
            " Where (EmpPK = " & locEmployeePK & ") " & _
            " And (LoanType = " & cmbLoanType.ItemData(cmbLoanType.ListIndex) & ") " & _
            " And (ZeroOut = 0)"
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            MsgBox "Please Zero Out Previous " & cmbLoanType.List(cmbLoanType.ListIndex) & "....           ", vbCritical, "Error..."
            Exit Sub
        End If
        rs.Close
    End If
    locLoanType = cmbLoanType.ItemData(cmbLoanType.ListIndex)
End If
End Sub

Private Sub cmbLoanType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtDateGranted.SetFocus
End Sub

Private Sub cmbPeriod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKSearch_Click
End Sub

Private Sub cmdCancelAdd_Click()
picBody.Enabled = True
picSearchAdd.Visible = False
End Sub

Private Sub cmdCancelSearch_Click()
picToolbar.Enabled = True
picBody.Enabled = True
picSearch.Visible = False
End Sub

Private Sub cmdOKAdd_Click()
If lstResultAdd.ListIndex = -1 Then Exit Sub
If isAdd_isLoan = 1 Then
    CLEARTEXT
    LOCKTEXT False
    TOOLBARFUNC 2
    locEmployeePK = lstResultAdd.ItemData(lstResultAdd.ListIndex)
    Array1 = Split(lstResultAdd.List(lstResultAdd.ListIndex), " - ", -1, 1)
    txtID.Text = CStr(Array1(0))
    txtName.Text = CStr(Array1(1))
    TRANSACTIONTYPE = is_ADDING
    cmdCancelAdd_Click
    cmbLoanType.SetFocus
ElseIf isAdd_isLoan = 2 Then
    Array1 = Split(lstResultAdd.List(lstResultAdd.ListIndex), " - ", -1, 1)
    u = "SELECT dbo.tbl_Personnel_Loans.PK, dbo.tbl_Personnel_Loans.EmpPK, dbo.tbl_Personnel_Payroll_Deductions_Table.Description, " & _
        " dbo.tbl_Personnel_Loans.ZeroOut, ISNULL((SELECT ROUND(SUM(Balance), 2) AS Expr1 From dbo.tbl_Personnel_Loans_SL " & _
        " WHERE (LoanKey = dbo.tbl_Personnel_Loans.PK)), 0) AS Bal " & _
        " FROM  dbo.tbl_Personnel_Loans LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Loans.LoanType = dbo.tbl_Personnel_Payroll_Deductions_Table.PK " & _
        " WHERE (ISNULL((SELECT ROUND(SUM(Balance), 2) AS Expr1 FROM  dbo.tbl_Personnel_Loans_SL AS tbl_Personnel_Loans_SL_1 " & _
        " WHERE (LoanKey = dbo.tbl_Personnel_Loans.PK)), 0) <> 0) " & _
        " AND (dbo.tbl_Personnel_Loans.EmpPK = " & lstResultAdd.ItemData(lstResultAdd.ListIndex) & ") " & _
        " AND (dbo.tbl_Personnel_Loans.ZeroOut = 0)"
    If rs.State = adStateOpen Then rs.Close
    ru.Open u, ConnOmega
    If ru.RecordCount = 0 Then
        MsgBox "No active loan!                     ", vbExclamation, "Info"
        ru.Close
        Exit Sub
    Else
        MainForm.picProgressBar.BackColor = &H8000000F
        DoEvents
        Screen.MousePointer = vbHourglass
        
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll_Report_LoanLedger WHERE (LogInName = '" & gbl_UserName & "')"
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_LoanLedger (LogInName, CompanyKey, EmployeeName) VALUES ('" & gbl_UserName & "', 1, '" & FORMATSQL(CStr(Array1(1))) & "')"
        iPK = 0: iLine = 0
        t = "SELECT PK " & _
            " FROM tbl_Personnel_Payroll_Report_LoanLedger " & _
            " WHERE (LogInName = '" & gbl_UserName & "')"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            iPK = rt!PK
        End If
        rt.Close
        
        iRec = 0: iTotRec = 0
        ru.MoveFirst
        While Not ru.EOF
            s = "SELECT dbo.tbl_Personnel_Loans_SL.* " & _
                " From dbo.tbl_Personnel_Loans_SL " & _
                " WHERE (LoanKey = " & ru!PK & ") " & _
                " ORDER BY PK"
            If rs.State = adStateOpen Then rs.Close
            rs.Open s, ConnOmega
            iTotRec = iTotRec + rs.RecordCount
            rs.Close
            ru.MoveNext
        Wend
        
        ru.MoveFirst
        While Not ru.EOF
            sLoanName_Status = ru!Description & _
                               IIf(ru!ZeroOut = 1, " [Zero out]", "")
            s = "SELECT dbo.tbl_Personnel_Loans_SL.* " & _
                " From dbo.tbl_Personnel_Loans_SL " & _
                " WHERE (LoanKey = " & ru!PK & ") " & _
                " ORDER BY PK"
            If rs.State = adStateOpen Then rs.Close
            rs.Open s, ConnOmega
            If rs.RecordCount > 0 Then
                dRunBal = 0
                While Not rs.EOF
                    iRec = iRec + 1
                    iLine = iLine + 1
                    sRemarks = "[" & Format(rs!TransactionDate, "mm/dd/yyyy") & "]" & IIf(Trim(rs!Remarks) <> "", " - " & rs!Remarks, "")
                    dDebit = rs!Debit
                    dCredit = rs!Credit
                    dRunBal = CDbl(Format(dRunBal, "#0.00")) + (CDbl(dDebit) - CDbl(dCredit))
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_LoanLedger_Det " & _
                                      " (MasterKey, Line, LoanName, Remarks, Debit, Credit, RunBal) " & _
                                      " VALUES (" & iPK & ", " & iLine & ", '" & FORMATSQL(CStr(sLoanName_Status)) & "', " & _
                                      " '" & FORMATSQL(CStr(sRemarks)) & "'," & CDbl(dDebit) & ", " & CDbl(dCredit) & ", " & CDbl(dRunBal) & ")"
                    UpdateProgress_No_Percent MainForm.picProgressBar, iRec / iTotRec
                    rs.MoveNext
                Wend
            End If
            rs.Close
            ru.MoveNext
        Wend
        Screen.MousePointer = vbDefault
        cmdCancelAdd_Click
        MainForm.picProgressBar.BackColor = &H8000000F
        DoEvents
        frmCrystalReportViewer.PRINT_LOAN_Ledger gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
    End If
    ru.Close
End If
'chkSSS.SetFocus
End Sub

Private Sub cmdOKSearch_Click()
If cmbPeriod.ListIndex = -1 Then Exit Sub
BROWSER cmbPeriod.ItemData(cmbPeriod.ListIndex), "is_FIND"
cmdCancelSearch_Click
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
    Case vbKeyF9:       PRESS_F9
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "LoanCtrl", "LoanCt", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "LoanCtrl", "LoanCt", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "LoanCtrl", "LoanCt", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "LoanCtrl", "LoanCt", ""), "is_END"
    Case vbKeyEscape:   PRESS_ESCAPE
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Top = (MainForm.Height - Me.Height) / 4
Me.Left = (MainForm.Width - Me.Width) / 5
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
POPULATE_COMBO_EXEMPTION "PK", "Description", "tbl_Personnel_Payroll_Deductions_Table", "Sorting", "ViewInDeductionModule", 2, cmbLoanType
POPULATE_COMBO "PK", "Interest", "tbl_Personnel_Loans_Interest", "PK", cmbInterest
LOCKTEXT True
TRANSACTIONTYPE = is_REFRESH
TOOLBARFUNC 1
BROWSER GetSetting(App.EXEName, "LoanCtrl", "LoanCt", ""), "is_LOAD"
If Trim(txtName.Text) = "" Then BROWSER GetSetting(App.EXEName, "LoanCtrl", "LoanCt", ""), "is_HOME"
'With cmbLoanType
'    .Clear
'    .AddItem "SSS LOAN"
'    .AddItem "PAG IBIG LOAN"
'End With

tmp = SetWindowLong(txtSearchAdd.hwnd, GWL_STYLE, GetWindowLong(txtSearchAdd.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picSearchAdd.Visible = True Then Cancel = -1
If picSearch.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstResult_Click()
If lstResult.ListIndex = -1 Then cmbPeriod.Clear: Exit Sub
cmbPeriod.Clear
t = "SELECT dbo.tbl_Personnel_Loans.PK, dbo.tbl_Personnel_Loans.DateGranted, " & _
    " dbo.tbl_Personnel_Payroll_Deductions_Table.Description " & _
    " FROM  dbo.tbl_Personnel_Loans LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Loans.LoanType = dbo.tbl_Personnel_Payroll_Deductions_Table.PK " & _
    " Where (dbo.tbl_Personnel_Loans.EmpPK = " & lstResult.ItemData(lstResult.ListIndex) & ") " & _
    " ORDER BY dbo.tbl_Personnel_Loans.DateGranted DESC"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
While Not rt.EOF
    cmbPeriod.AddItem "[" & Format(rt!DateGranted, "mm/dd/yyyy") & "] " & rt!Description
    cmbPeriod.ItemData(cmbPeriod.NewIndex) = rt!PK
    rt.MoveNext
Wend
rt.Close
If cmbPeriod.ListCount Then cmbPeriod.ListIndex = 0
End Sub

Private Sub lstResult_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbPeriod.SetFocus
End Sub

Private Sub lstResultAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAdd_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
'        Case "Refresh"
'            'ToDo: Add 'Refresh' button code.
'            MsgBox "Add 'Refresh' button code."
'        Case "Post"
'            'ToDo: Add 'Post' button code.
'            MsgBox "Add 'Post' button code."
    Case "Add":           PRESS_INSERT
    Case "Edit":          PRESS_F2
    Case "Delete":        PRESS_DELETE
    Case "First"
        Select Case Toolbar1.Buttons(7).Caption
            Case "Save":  PRESS_F5
            Case "First": BROWSER GetSetting(App.EXEName, "LoanCtrl", "LoanCt", ""), "is_HOME"
        End Select
    Case "Back"
        Select Case Toolbar1.Buttons(9).Caption
            Case "Undo":  PRESS_ESCAPE
            Case "Back":  BROWSER GetSetting(App.EXEName, "LoanCtrl", "LoanCt", ""), "is_PAGEUP"
        End Select
    Case "Next":          BROWSER GetSetting(App.EXEName, "LoanCtrl", "LoanCt", ""), "is_PAGEDOWN"
    Case "Last":          BROWSER GetSetting(App.EXEName, "LoanCtrl", "LoanCt", ""), "is_END"
    Case "Find":          PRESS_F6
    Case "Post":          PRESS_F8
    Case "Print":         PRESS_F9
    Case "Close":         PRESS_ESCAPE
End Select
End Sub

Private Sub txtAmortization_Change()
If IsNumeric(txtAmortization.Text) And _
IsNumeric(txtTotalAmount.Text) Then
    If Trim(txtTotalAmount.Text) <> "" Then
        TotalAmount = CDbl(txtTotalAmount.Text)
    Else
        TotalAmount = 0
    End If
    If Trim(txtAmortization.Text) Then
        Amortization = CDbl(txtAmortization.Text)
    Else
        Amortization = 0
    End If
    On Error GoTo PG:
    txtNoMonths.Text = CInt(CDbl(TotalAmount) / CDbl(Amortization))
End If
Exit Sub
PG:
If Err.Number = 6 Then Exit Sub
End Sub

Private Sub txtAmortization_GotFocus()
txtAmortization.Alignment = 0
If IsNumeric(txtAmortization.Text) Then
    txtAmortization.Text = CDbl(txtAmortization.Text)
End If
HTEXT txtAmortization
End Sub

Private Sub txtAmortization_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtNoMonths.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtTotalAmount.SetFocus
End If
End Sub

Private Sub txtAmortization_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtAmortization_LostFocus()
If Trim(txtAmortization.Text) <> "" Then
    txtAmortization.Text = Format(txtAmortization.Text, "##,##0.00")
Else
    txtAmortization.Text = "0.00"
End If
txtAmortization.Alignment = 1
End Sub

Private Sub txtBalance_GotFocus()
txtBalance.Alignment = 0
If IsNumeric(txtBalance.Text) Then
    txtBalance.Text = CDbl(txtBalance.Text)
End If
HTEXT txtBalance
End Sub

Private Sub txtBalance_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtID.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtTotalPaid.SetFocus
End If
End Sub

Private Sub txtBalance_LostFocus()
If Trim(txtBalance.Text) <> "" Then
    txtBalance.Text = Format(txtBalance.Text, "##,##0.00")
Else
    txtBalance.Text = "0.00"
End If
txtBalance.Alignment = 1
End Sub

Private Sub txtDateFrom_Change()
If IsDate(txtDateFrom.Text) And _
IsNumeric(txtNoMonths.Text) Then
    txtDateTo.Text = Format(DateSerial(Year(FormatDateTime(txtDateFrom.Text, vbShortDate)), Month(FormatDateTime(txtDateFrom.Text, vbShortDate)) + (CInt(txtNoMonths.Text) - 1), Day(FormatDateTime(txtDateFrom.Text, vbShortDate))), "mm/dd/yyyy")
End If
End Sub

Private Sub txtDateFrom_GotFocus()
HTEXT txtDateFrom
End Sub

Private Sub txtDateFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
'    txtDateTo.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtNoMonths.SetFocus
End If
End Sub

Private Sub txtDateFrom_LostFocus()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If IsDate(txtDateFrom.Text) Then
        txtDateFrom.Text = Format(FormatDateTime(Trim(txtDateFrom.Text), vbShortDate), "mm/dd/yyyy")
    Else
        MsgBox "PLEASE SUPPLY A VALID DATE!         ", vbCritical, "Error.."
        txtDateFrom.SetFocus
        HTEXT txtDateFrom
    End If
End If
End Sub

Private Sub txtDateGranted_GotFocus()
HTEXT txtDateGranted
End Sub

Private Sub txtDateGranted_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtLoanAmount.SetFocus
ElseIf KeyCode = vbKeyUp Then
    cmbLoanType.SetFocus
End If
End Sub

Private Sub txtDateGranted_LostFocus()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If Trim(txtDateGranted.Text) <> "" Then
        If IsDate(txtDateGranted.Text) Then
            txtDateGranted.Text = Format(FormatDateTime(Trim(txtDateGranted.Text), vbShortDate), "mm/dd/yyyy")
        Else
            MsgBox "PLEASE SUPPLY A VALID DATE!         ", vbCritical, "Error.."
            txtDateGranted.SetFocus
            HTEXT txtDateGranted
        End If
    End If
End If
End Sub

Private Sub txtDateTo_GotFocus()
HTEXT txtDateTo
End Sub

Private Sub txtDateTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    chkZeroOut.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtDateFrom.SetFocus
End If
End Sub

Private Sub txtID_GotFocus()
HTEXT txtID
End Sub

Private Sub txtID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtName.SetFocus
End If
End Sub

Private Sub txtInterest_Change()
If locInterestType = 0 Then
    txtInterest.Text = ""
End If
If IsNumeric(txtInterest.Text) Then
    If locInterestType = 1 Then
        txtTotalAmount.Text = Format(CDbl(IIf(Trim(txtLoanAmount.Text) = "", 0, txtLoanAmount.Text)) + _
                              ((CDbl(IIf(Trim(txtInterest.Text) = "", 0, txtInterest.Text)) / 100) * CDbl(txtLoanAmount.Text)), "##,##0.00")
    ElseIf locInterestType = 2 Then
        txtTotalAmount.Text = Format(CDbl(IIf(Trim(txtLoanAmount.Text) = "", 0, txtLoanAmount.Text)) + _
                              CDbl(IIf(Trim(txtInterest.Text) = "", 0, txtInterest.Text)), "##,##0.00")
    End If
End If
End Sub

Private Sub txtInterest_GotFocus()
txtInterest.Alignment = 0
If IsNumeric(txtInterest.Text) Then
    txtInterest.Text = CDbl(txtInterest.Text)
End If
HTEXT txtInterest
End Sub

Private Sub txtInterest_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtTotalAmount.SetFocus
ElseIf KeyCode = vbKeyUp Then
    cmbLoanType.SetFocus
    'chkPercent.SetFocus
End If
End Sub

Private Sub txtInterest_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
If locInterestType = 0 Then
    MsgBox "PLEASE SELECT INTEREST TYPE!            ", vbInformation, ""
    cmbInterest.SetFocus
'    txtInterest.Text = ""
    'chkPercent.SetFocus
End If
End Sub

Private Sub txtInterest_LostFocus()
If Trim(txtInterest.Text) <> "" Then
    txtInterest.Text = Format(txtInterest.Text, "##,##0.00")
Else
    txtInterest.Text = "0.00"
End If
txtInterest.Alignment = 1
If IsNumeric(txtInterest.Text) Then
    If locInterestType = 1 Then
        txtTotalAmount.Text = Format(CDbl(IIf(Trim(txtLoanAmount.Text) = "", 0, txtLoanAmount.Text)) + _
                              ((CDbl(IIf(Trim(txtInterest.Text) = "", 0, txtInterest.Text)) / 100) * CDbl(txtLoanAmount.Text)), "##,##0.00")
    ElseIf locInterestType = 2 Then
        txtTotalAmount.Text = Format(CDbl(IIf(Trim(txtLoanAmount.Text) = "", 0, txtLoanAmount.Text)) + _
                              CDbl(IIf(Trim(txtInterest.Text) = "", 0, txtInterest.Text)), "##,##0.00")
    End If
End If
End Sub

Private Sub txtLoanAmount_Change()
If IsNumeric(txtLoanAmount.Text) Then
    If locInterestType = 1 Then
        txtTotalAmount.Text = Format(CDbl(IIf(Trim(txtLoanAmount.Text) = "", 0, txtLoanAmount.Text)) + _
                              ((CDbl(IIf(Trim(txtInterest.Text) = "", 0, txtInterest.Text)) / 100) * CDbl(txtLoanAmount.Text)), "##,##0.00")
    ElseIf locInterestType = 2 Then
        txtTotalAmount.Text = Format(CDbl(IIf(Trim(txtLoanAmount.Text) = "", 0, txtLoanAmount.Text)) + _
                              CDbl(IIf(Trim(txtInterest.Text) = "", 0, txtInterest.Text)), "##,##0.00")
    End If
End If
End Sub

Private Sub txtLoanAmount_GotFocus()
txtLoanAmount.Alignment = 0
If IsNumeric(txtLoanAmount.Text) Then
    txtLoanAmount.Text = CDbl(txtLoanAmount.Text)
End If
HTEXT txtLoanAmount
End Sub

Private Sub txtLoanAmount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    cmbInterest.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtDateGranted.SetFocus
End If
End Sub

Private Sub txtLoanAmount_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtLoanAmount_LostFocus()
If Trim(txtLoanAmount.Text) <> "" Then
    txtLoanAmount.Text = Format(txtLoanAmount.Text, "##,##0.00")
Else
    txtLoanAmount.Text = "0.00"
End If
txtLoanAmount.Alignment = 1
If IsNumeric(txtLoanAmount.Text) Then
    If locInterestType = 1 Then
        txtTotalAmount.Text = Format(CDbl(IIf(Trim(txtLoanAmount.Text) = "", 0, txtLoanAmount.Text)) + _
                              ((CDbl(IIf(Trim(txtInterest.Text) = "", 0, txtInterest.Text)) / 100) * CDbl(txtLoanAmount.Text)), "##,##0.00")
    ElseIf locInterestType = 2 Then
        txtTotalAmount.Text = Format(CDbl(IIf(Trim(txtLoanAmount.Text) = "", 0, txtLoanAmount.Text)) + _
                              CDbl(IIf(Trim(txtInterest.Text) = "", 0, txtInterest.Text)), "##,##0.00")
    End If
End If
End Sub

Private Sub txtName_GotFocus()
HTEXT txtName
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    cmbLoanType.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtID.SetFocus
End If
End Sub

Private Sub txtNoMonths_Change()
If IsDate(txtDateFrom.Text) And _
IsNumeric(txtNoMonths.Text) Then
    txtDateTo.Text = Format(DateSerial(Year(FormatDateTime(txtDateFrom.Text, vbShortDate)), Month(FormatDateTime(txtDateFrom.Text, vbShortDate)) + (CInt(txtNoMonths.Text) - 1), Day(FormatDateTime(txtDateFrom.Text, vbShortDate))), "mm/dd/yyyy")
End If
End Sub

Private Sub txtNoMonths_GotFocus()
txtNoMonths.Alignment = 0
If IsNumeric(txtNoMonths.Text) Then
    txtNoMonths.Text = CDbl(txtNoMonths.Text)
End If
HTEXT txtNoMonths
End Sub

Private Sub txtNoMonths_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtDateFrom.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtAmortization.SetFocus
End If
End Sub

Private Sub txtNoMonths_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtNoMonths_LostFocus()
txtNoMonths.Alignment = 1
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstResult.Clear: cmbPeriod.Clear: Exit Sub
lstResult.Clear: cmbPeriod.Clear
s = "SELECT dbo.tbl_Personnel_Loans.EmpPK, dbo.tbl_Personnel_IDNumber.IDNumber, " & _
    " dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
    " dbo.tbl_Personnel_Information.MiddleName " & _
    " FROM  dbo.tbl_Personnel_Loans LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Loans.EmpPK = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN  " & _
    " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
    " GROUP BY dbo.tbl_Personnel_Loans.EmpPK, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName " & _
    " HAVING (dbo.tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
    " ORDER BY dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_IDNumber.IDNumber"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResult.AddItem rs!IDNumber & " - " & rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
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
If isAdd_isLoan = 1 Then
    s = "sp_Personnel_Action_Search_Add('" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%')"
ElseIf isAdd_isLoan = 2 Then
    s = "SELECT dbo.tbl_Personnel_Loans.EmpPK AS PK, dbo.tbl_Personnel_IDNumber.IDNumber, " & _
        " dbo.tbl_Personnel_Information.LastName + ',  ' + dbo.tbl_Personnel_Information.FirstName + '  ' + dbo.tbl_Personnel_Information.MiddleName AS EmployeeName, dbo.tbl_Personnel_Information.LastName " & _
        " FROM  dbo.tbl_Personnel_Loans LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Loans.EmpPK = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
        " GROUP BY dbo.tbl_Personnel_Loans.EmpPK, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName + ',  ' + dbo.tbl_Personnel_Information.FirstName + '  ' + dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Information.LastName " & _
        " HAVING (dbo.tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%') " & _
        " ORDER BY EmployeeName, dbo.tbl_Personnel_IDNumber.IDNumber"
End If
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultAdd.AddItem rs!IDNumber & " - " & rs!EmployeeName
    lstResultAdd.ItemData(lstResultAdd.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResultAdd.ListCount Then lstResultAdd.ListIndex = 0
End Sub

Private Sub txtSearchAdd_GotFocus()
HTEXT txtSearchAdd
End Sub

Private Sub txtSearchAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultAdd.SetFocus
End Sub

Private Sub txtTotalAmount_Change()

If IsNumeric(txtAmortization.Text) And _
IsNumeric(txtTotalAmount.Text) Then
    If Trim(txtTotalAmount.Text) <> "" Then
        TotalAmount = CDbl(txtTotalAmount.Text)
    Else
        TotalAmount = 0
    End If
    If Trim(txtAmortization.Text) Then
        Amortization = CDbl(txtAmortization.Text)
    Else
        Amortization = 0
    End If
    On Error GoTo PG:
    txtNoMonths.Text = CInt(CDbl(TotalAmount) / CDbl(Amortization))
End If
Exit Sub
PG:
If Err.Number = 6 Then Exit Sub
End Sub

Private Sub txtTotalAmount_GotFocus()
txtTotalAmount.Alignment = 0
If IsNumeric(txtTotalAmount.Text) Then
    txtTotalAmount.Text = CDbl(txtTotalAmount.Text)
End If
HTEXT txtTotalAmount
End Sub

Private Sub txtTotalAmount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtAmortization.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtInterest.SetFocus
End If
End Sub

Private Sub txtTotalAmount_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtTotalAmount_LostFocus()
If Trim(txtTotalAmount.Text) <> "" Then
    txtTotalAmount.Text = Format(txtTotalAmount.Text, "##,##0.00")
Else
    txtTotalAmount.Text = "0.00"
End If
txtTotalAmount.Alignment = 1
End Sub

Private Sub txtTotalPaid_GotFocus()
txtTotalPaid.Alignment = 0
If IsNumeric(txtTotalPaid.Text) Then
    txtTotalPaid.Text = CDbl(txtTotalPaid.Text)
End If
HTEXT txtTotalPaid
End Sub

Private Sub txtTotalPaid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtBalance.SetFocus
ElseIf KeyCode = vbKeyUp Then
    chkZeroOut.SetFocus
End If
End Sub

Private Sub txtTotalPaid_LostFocus()
If Trim(txtTotalPaid.Text) <> "" Then
    txtTotalPaid.Text = Format(txtTotalPaid.Text, "##,##0.00")
Else
    txtTotalPaid.Text = "0.00"
End If
txtTotalPaid.Alignment = 1
End Sub




