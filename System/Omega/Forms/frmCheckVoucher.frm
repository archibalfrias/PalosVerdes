VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAcctgCheckVoucher 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCheckVoucher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   11760
   Begin RPVGCC.b8Container picSeachAccount 
      Height          =   3015
      Left            =   2520
      TabIndex        =   45
      Top             =   4650
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5318
      BackColor       =   15396057
      Begin VB.ListBox lstAccount 
         Height          =   1425
         Left            =   120
         TabIndex        =   49
         Top             =   840
         Width           =   4695
      End
      Begin VB.TextBox txtSearchAccount 
         Height          =   315
         Left            =   120
         TabIndex        =   48
         Top             =   480
         Width           =   4695
      End
      Begin VB.CommandButton cmdCancelAccount 
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
         Left            =   2520
         Picture         =   "frmCheckVoucher.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   2400
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKAccount 
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
         Left            =   840
         Picture         =   "frmCheckVoucher.frx":1026
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   2400
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   45
         TabIndex        =   50
         Top             =   45
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   609
         Caption         =   "Search Account"
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
         Icon            =   "frmCheckVoucher.frx":1698
         ShadowVisible   =   0   'False
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15600
      TabIndex        =   81
      Top             =   0
      Width           =   15600
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   82
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
               Caption         =   " Post   "
               Key             =   "Post"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               Key             =   "Refresh"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         MousePointer    =   99
         MouseIcon       =   "frmCheckVoucher.frx":1C32
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   10260
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   83
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmCheckVoucher.frx":1F4C
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
   Begin RPVGCC.b8Container picSearch 
      Height          =   4095
      Left            =   3600
      TabIndex        =   69
      Top             =   1680
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7223
      BackColor       =   15396057
      Begin VB.ListBox lstSearch 
         Height          =   2595
         Left            =   120
         TabIndex        =   73
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   72
         Top             =   480
         Width           =   4215
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
         Left            =   2280
         Picture         =   "frmCheckVoucher.frx":265F
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   3480
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
         Left            =   600
         Picture         =   "frmCheckVoucher.frx":2DBB
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   3480
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar3 
         Height          =   345
         Left            =   45
         TabIndex        =   74
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
         Icon            =   "frmCheckVoucher.frx":342D
         ShadowVisible   =   0   'False
      End
   End
   Begin RPVGCC.b8Container picExplainSLine 
      Height          =   855
      Left            =   360
      TabIndex        =   27
      Top             =   3360
      Visible         =   0   'False
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox txtRRNumber1 
         Height          =   315
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   80
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtInvoiceNumber1 
         Height          =   315
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   79
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtInvoiceNumber 
         Height          =   315
         Left            =   7440
         MaxLength       =   100
         TabIndex        =   76
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtRRNumber 
         Height          =   315
         Left            =   5520
         MaxLength       =   100
         TabIndex        =   75
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtSLineAmt1 
         Height          =   315
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   44
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtSLineDesc1 
         Height          =   315
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtSLineAmt 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9360
         MaxLength       =   100
         TabIndex        =   31
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtSLineDesc 
         Height          =   315
         Left            =   120
         MaxLength       =   100
         TabIndex        =   29
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Number"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7440
         TabIndex        =   78
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "RR Number"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5520
         TabIndex        =   77
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9360
         TabIndex        =   30
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   3015
      End
   End
   Begin RPVGCC.b8Container picAdd 
      Height          =   4095
      Left            =   3600
      TabIndex        =   37
      Top             =   1680
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7223
      BackColor       =   15396057
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
         Left            =   600
         Picture         =   "frmCheckVoucher.frx":39C7
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3480
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
         Left            =   2280
         Picture         =   "frmCheckVoucher.frx":4039
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   3480
         Width           =   1560
      End
      Begin VB.TextBox txtSearchAdd 
         Height          =   315
         Left            =   120
         TabIndex        =   39
         Top             =   480
         Width           =   4215
      End
      Begin VB.ListBox lstResultAdd 
         Height          =   2595
         Left            =   120
         TabIndex        =   38
         Top             =   840
         Width           =   4215
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   42
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
         Icon            =   "frmCheckVoucher.frx":4795
         ShadowVisible   =   0   'False
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   7365
      Width           =   11760
      _ExtentX        =   20743
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
   Begin VB.PictureBox picADSLine 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1080
      ScaleHeight     =   855
      ScaleWidth      =   9495
      TabIndex        =   51
      Top             =   3960
      Visible         =   0   'False
      Width           =   9495
      Begin RPVGCC.b8Container picADSLine1 
         Height          =   855
         Left            =   0
         TabIndex        =   52
         Top             =   0
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   1508
         BackColor       =   8438015
         Begin VB.TextBox txtDebit 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6480
            MaxLength       =   100
            TabIndex        =   60
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtAccCode 
            Height          =   315
            Left            =   120
            MaxLength       =   100
            TabIndex        =   59
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtAccDesc 
            Height          =   315
            Left            =   1440
            MaxLength       =   100
            TabIndex        =   58
            Top             =   360
            Width           =   4935
         End
         Begin VB.TextBox txtCredit 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7920
            MaxLength       =   100
            TabIndex        =   57
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtAccCode1 
            Height          =   315
            Left            =   4680
            MaxLength       =   100
            TabIndex        =   56
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtAccDesc1 
            Height          =   315
            Left            =   4920
            MaxLength       =   100
            TabIndex        =   55
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtDebit1 
            Height          =   315
            Left            =   5160
            MaxLength       =   100
            TabIndex        =   54
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtCredit1 
            Height          =   315
            Left            =   5400
            MaxLength       =   100
            TabIndex        =   53
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Debit"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6480
            TabIndex        =   64
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Account Code"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Account Description"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1440
            TabIndex        =   62
            Top             =   120
            Width           =   4935
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Credit"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   7920
            TabIndex        =   61
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   360
      ScaleHeight     =   5775
      ScaleWidth      =   11055
      TabIndex        =   1
      Top             =   1200
      Width           =   11055
      Begin VB.TextBox txtORDate 
         Height          =   315
         Left            =   2640
         MaxLength       =   100
         TabIndex        =   67
         Text            =   "01/01/2012"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C6B8A4&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   735
         Left            =   0
         ScaleHeight     =   735
         ScaleWidth      =   11055
         TabIndex        =   65
         Top             =   2640
         Width           =   11055
         Begin VB.TextBox txtAmtWords 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C6B8A4&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   675
            Left            =   0
            MaxLength       =   100
            TabIndex        =   66
            Top             =   50
            Width           =   10935
         End
      End
      Begin VB.TextBox txtCheckDate 
         Height          =   315
         Left            =   7200
         MaxLength       =   100
         TabIndex        =   35
         Text            =   "01/01/2012"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtEntered 
         Height          =   315
         Left            =   8400
         MaxLength       =   100
         TabIndex        =   24
         Top             =   5400
         Width           =   2655
      End
      Begin VB.TextBox txtChecked 
         Height          =   315
         Left            =   5640
         MaxLength       =   100
         TabIndex        =   22
         Top             =   5400
         Width           =   2655
      End
      Begin VB.TextBox txtPrepared 
         Height          =   315
         Left            =   2760
         MaxLength       =   100
         TabIndex        =   20
         Top             =   5400
         Width           =   2775
      End
      Begin VB.TextBox txtApproved 
         Height          =   315
         Left            =   0
         MaxLength       =   100
         TabIndex        =   18
         Text            =   "JOSE ALBERT M. CASEÑAS"
         Top             =   5400
         Width           =   2655
      End
      Begin VB.TextBox txtCheckNumber 
         Height          =   315
         Left            =   4680
         MaxLength       =   100
         TabIndex        =   13
         Top             =   2280
         Width           =   1335
      End
      Begin MSComctlLib.ListView lstExplanation 
         Height          =   1215
         Left            =   0
         TabIndex        =   11
         Top             =   960
         Width           =   11055
         _ExtentX        =   19500
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   11642
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "RR Number"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Invoice Number"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "RRNumber"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.TextBox txtORNumber 
         Height          =   315
         Left            =   480
         MaxLength       =   100
         TabIndex        =   9
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtPayeeName 
         Height          =   315
         Left            =   2640
         MaxLength       =   100
         TabIndex        =   8
         Top             =   360
         Width           =   8415
      End
      Begin VB.TextBox txtPayeeCode 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtCVDate 
         Height          =   315
         Left            =   9600
         MaxLength       =   100
         TabIndex        =   4
         Top             =   0
         Width           =   1455
      End
      Begin VB.TextBox txtCVNumber 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   2
         Top             =   0
         Width           =   1455
      End
      Begin MSComctlLib.ListView lstAccountDistribution 
         Height          =   1200
         Left            =   720
         TabIndex        =   15
         Top             =   3600
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   2117
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
            Text            =   "Account #"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Account Name"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Debit"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Credit"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "OR Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   68
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6120
         TabIndex        =   36
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblDebit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   6720
         TabIndex        =   34
         Top             =   4830
         Width           =   1575
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL >>"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5400
         TabIndex        =   33
         Top             =   4830
         Width           =   975
      End
      Begin VB.Label lblCredit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   8280
         TabIndex        =   32
         Top             =   4830
         Width           =   1455
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   9480
         TabIndex        =   26
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL >>"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8520
         TabIndex        =   25
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "ENTERED:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8400
         TabIndex        =   23
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "CHECKED:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5640
         TabIndex        =   21
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "PREPARED:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   19
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "APPROVED:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "ACCOUNT DISTRIBUTION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   3360
         Width           =   3615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque #"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   14
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "OR #"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "EXPLANATION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Payee"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CV Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8520
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "CV #"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   30
         Width           =   1095
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11280
      Top             =   600
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
            Picture         =   "frmCheckVoucher.frx":4D2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckVoucher.frx":5A09
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckVoucher.frx":66E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckVoucher.frx":73BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckVoucher.frx":8097
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckVoucher.frx":8D71
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckVoucher.frx":9A4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckVoucher.frx":A725
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckVoucher.frx":B3FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckVoucher.frx":BCD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckVoucher.frx":C9B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckVoucher.frx":D68D
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckVoucher.frx":E367
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckVoucher.frx":F041
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckVoucher.frx":FD1B
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAcctgCheckVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SearchType As Long

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim TRANS_DETAIL As Long
Const is_DET_REFRESH = 0
Const is_DET_ADDING = 1
Const is_DET_EDITTING = 2

Dim iRow            As Long
Dim iFocusE         As Long
Dim iFocusA         As Long
Dim tmp             As Long
Dim iPayee          As Double
Public iPayeeType   As Long
Dim iCVNoAuto       As Long
Dim sCVNumber       As String
Dim iPK             As Double
Dim iFocusAcc       As Long

Dim TableName, DetailTableName, Columns, ColumnsDet

Dim Arr, x, i, l, k, iDebit, iCredit, iAmount, iTotalAmount, dAmount, sBankCode, sBankName, sSearchPayeeName

Private Sub BROWSER(sCV, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If sCV <> "" Then
            s = "SELECT TOP 1 tbl_Acctg_CheckVoucher.* " & _
                " FROM tbl_Acctg_CheckVoucher " & _
                " WHERE (CVNumber = '" & FORMATSQL(CStr(sCV)) & "') " & _
                " ORDER BY CVNumber"
        Else
            s = "SELECT TOP 1 tbl_Acctg_CheckVoucher.* " & _
                " FROM tbl_Acctg_CheckVoucher " & _
                " ORDER BY CVNumber"
        End If
    Case "is_HOME"
        If picAdd.Visible = True Then Exit Sub
        If picExplainSLine.Visible = True Then Exit Sub
        If picADSLine.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Acctg_CheckVoucher.* " & _
            " FROM tbl_Acctg_CheckVoucher " & _
            " ORDER BY CVNumber"
    Case "is_PAGEUP"
        If picAdd.Visible = True Then Exit Sub
        If picExplainSLine.Visible = True Then Exit Sub
        If picADSLine.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Acctg_CheckVoucher.* " & _
            " FROM tbl_Acctg_CheckVoucher " & _
            " WHERE (CVNumber < '" & FORMATSQL(CStr(sCV)) & "') " & _
            " ORDER BY CVNumber DESC"
    Case "is_PAGEDOWN"
        If picAdd.Visible = True Then Exit Sub
        If picExplainSLine.Visible = True Then Exit Sub
        If picADSLine.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Acctg_CheckVoucher.* " & _
            " FROM tbl_Acctg_CheckVoucher " & _
            " WHERE (CVNumber > '" & FORMATSQL(CStr(sCV)) & "') " & _
            " ORDER BY CVNumber"
    Case "is_END"
        If picAdd.Visible = True Then Exit Sub
        If picExplainSLine.Visible = True Then Exit Sub
        If picADSLine.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Acctg_CheckVoucher.* " & _
            " FROM tbl_Acctg_CheckVoucher " & _
            " ORDER BY CVNumber DESC"
    Case "is_FIND"
        s = "SELECT TOP 1 tbl_Acctg_CheckVoucher.* " & _
            " FROM tbl_Acctg_CheckVoucher " & _
            " WHERE (PK = " & sCV & ") " & _
            " ORDER BY CVNumber"
End Select
If rs.State = adStateOpen Then rs.Close
'Debug.Print s
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iPayee = rs!PayeeKey
    iPayeeType = rs!PayeeType
    txtCVNumber.Text = rs!CVNumber
    txtCVDate.Text = Format(rs!CVDate, "mm/dd/yyyy")
    txtPayeeCode.Text = ""
    txtPayeeName.Text = ""
    Select Case iPayeeType
        Case 1
            t = "SELECT PK, SupplierCode as Code, SupplierName as sName " & _
                " From tbl_Inv_Supplier " & _
                " WHERE (PK = " & iPayee & ")"
        Case 2
            t = "SELECT PK, LastName + ',  ' + FirstName + '  ' + MiddleName AS sName, " & _
                " ISNULL((SELECT TOP 1 IDNumber From tbl_Member_IDNumber " & _
                " Where (MemberKey = tbl_Member_Information.PK) " & _
                " ORDER BY IDCounter DESC, IDNumber), '') AS Code " & _
                " From tbl_Member_Information " & _
                " WHERE (PK = " & iPayee & ") "
        Case 3
            t = "SELECT PK, LastName + ',  ' + FirstName + '  ' + MiddleName AS sName, " & _
                " ISNULL((SELECT TOP 1 IDNumber From tbl_Personnel_IDNumber " & _
                " Where (ProfileKey = tbl_Personnel_Information.PK) " & _
                " ORDER BY IDNumber DESC), '') AS Code " & _
                " From tbl_Personnel_Information " & _
                " WHERE (PK = " & iPayee & ") "
    End Select
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtPayeeCode.Text = rt!Code
        txtPayeeName.Text = rt!sName
    End If
    rt.Close
    
    txtORNumber.Text = rs!ORNumber
    txtCheckNumber.Text = rs!CheckNumber
    txtCheckDate.Text = IIf(IsNull(rs!CheckDate), "", Format(rs!CheckDate, "mm/dd/yyyy"))
    txtApproved.Text = rs!Approved
    txtPrepared.Text = rs!Prepared
    txtChecked.Text = rs!Checked
    txtEntered.Text = rs!Entered
    lblTotal.Caption = Format(rs!TotalAmt, "#,##0.00")
    txtAmtWords.Text = AMT2WORDS(CDbl(rs!TotalAmt))
    imgPosted.Visible = IIf(rs!Posted = 1, True, False)
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    
    CLEAR_EXPLANATION
    t = "SELECT Line, Description, Amt, RRNumber, InvNumber " & _
        " From tbl_Acctg_CheckVoucher_Detail " & _
        " Where (CVKey = " & rs!PK & ") " & _
        " ORDER BY Line"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstExplanation.ListItems.Clear
        While Not rt.EOF
            Set x = lstExplanation.ListItems.Add
            x.Text = ""
            x.SubItems(1) = rt!Description
            x.SubItems(2) = IIf(IsNull(rt!RRNumber), " ", rt!RRNumber)
            x.SubItems(3) = IIf(IsNull(rt!InvNumber), " ", rt!InvNumber)
            x.SubItems(4) = Format(rt!Amt, "#,##0.00")
            rt.MoveNext
        Wend
    End If
    rt.Close
    
    CLEAR_ACCOUNT_DIST
    iDebit = 0: iCredit = 0
    t = "SELECT tbl_Acctg_CheckVoucher_AD.AccCode, tbl_GL_Accounts.AccountName, " & _
        " tbl_Acctg_CheckVoucher_AD.Debit, tbl_Acctg_CheckVoucher_AD.Credit " & _
        " FROM tbl_Acctg_CheckVoucher_AD LEFT OUTER JOIN " & _
        " tbl_GL_Accounts ON tbl_Acctg_CheckVoucher_AD.AccCode = tbl_GL_Accounts.AccountCode " & _
        " Where (tbl_Acctg_CheckVoucher_AD.CVKey = " & rs!PK & ") " & _
        " ORDER BY tbl_Acctg_CheckVoucher_AD.Line"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstAccountDistribution.ListItems.Clear
        While Not rt.EOF
            iDebit = iDebit + CDbl(rt!Debit)
            iCredit = iCredit + CDbl(rt!Credit)
            Set x = lstAccountDistribution.ListItems.Add
            x.Text = ""
            x.SubItems(1) = rt!AccCode
            x.SubItems(2) = rt!AccountName
            x.SubItems(3) = IIf(rt!Debit = 0, " ", Format(rt!Debit, "#,##0.00"))
            x.SubItems(4) = IIf(rt!Credit = 0, " ", Format(rt!Credit, "#,##0.00"))
            rt.MoveNext
        Wend
    End If
    rt.Close
    
    lblDebit.Caption = Format(iDebit, "#,##0.00")
    lblCredit.Caption = Format(iCredit, "#,##0.00")
    
    SaveSetting App.EXEName, "CVNumber", "CVNum", rs!CVNumber
    
End If
rs.Close
End Sub

Private Sub CLEARTEXT()
iPayee = 0
'iPayeeType = 0
txtCVNumber.Text = ""
txtCVDate.Text = ""
txtPayeeCode.Text = ""
txtPayeeName.Text = ""
txtORNumber.Text = ""
txtORDate.Text = ""
txtCheckNumber.Text = ""
txtCheckDate.Text = ""
txtApproved.Text = ""
txtPrepared.Text = ""
txtChecked.Text = ""
txtEntered.Text = ""
lblTotal.Caption = "0.00"
txtAmtWords.Text = ""
lblDebit.Caption = "0.00"
lblCredit.Caption = "0.00"
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
imgPosted.Visible = False
CLEAR_EXPLANATION
CLEAR_ACCOUNT_DIST
End Sub

Private Sub CLEAR_EXPLANATION()
lstExplanation.ListItems.Clear
Set x = lstExplanation.ListItems.Add()
x.Text = " "
x.SubItems(1) = " "
x.SubItems(2) = " "
x.SubItems(3) = " "
End Sub

Private Sub CLEAR_ACCOUNT_DIST()
lstAccountDistribution.ListItems.Clear
Set x = lstAccountDistribution.ListItems.Add()
x.Text = " "
x.SubItems(1) = " "
x.SubItems(2) = " "
x.SubItems(3) = " "
x.SubItems(4) = " "
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtCVNumber.Locked = bln 'True
txtCVDate.Locked = bln 'True
txtPayeeCode.Locked = True
txtPayeeName.Locked = True
txtORNumber.Locked = bln
txtORDate.Locked = bln
txtCheckNumber.Locked = bln
txtCheckDate.Locked = bln
txtApproved.Locked = bln
txtPrepared.Locked = True
txtChecked.Locked = bln
txtEntered.Locked = bln
txtAccDesc.Locked = True
End Sub

Public Sub TOOLBARFUNC(intSelect As Integer)
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
    End Select
End With
End Sub

Private Sub PRESS_INSERT()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSearch.Visible = True Then Exit Sub
    If picAdd.Visible = True Then Exit Sub
    If picExplainSLine.Visible = True Then Exit Sub
    If picADSLine.Visible = True Then Exit Sub
    If AccessRights("Check Voucher", "Add") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    PopupMenu MainFormPopupF.mnuCVAdd, , Toolbar1.Buttons(1).Left, Toolbar1.Buttons(1).Top + Toolbar1.Buttons(1).Height
Else
    If picSearch.Visible = True Then Exit Sub
    If picAdd.Visible = True Then Exit Sub
    If picExplainSLine.Visible = True Then Exit Sub
    If picADSLine.Visible = True Then Exit Sub
    If AccessRights("Check Voucher", "Add") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If iFocusE = 1 Then
        With lstExplanation.ListItems
            If Trim(.Item(iRow).SubItems(1)) = "" Then
                .Item(iRow).SubItems(1) = " "
                .Item(iRow).SubItems(2) = ""
                .Item(iRow).SubItems(3) = ""
            Else
                Set x = .Add()
                x.Text = " "
                x.SubItems(1) = " "
                x.SubItems(2) = " "
                x.SubItems(3) = " "
                iRow = .Count
            End If
        End With
        lstExplanation.ListItems(iRow).EnsureVisible
        lstExplanation.ListItems(iRow).Selected = True
        picMain.Enabled = False
        picToolbar.Enabled = False
        txtSLineDesc.Text = ""
        txtSLineAmt.Text = ""
        picExplainSLine.ZOrder 0
        picExplainSLine.Visible = True
        TRANS_DETAIL = is_DET_ADDING
        txtSLineDesc.SetFocus
        Exit Sub
    End If
    If iFocusA = 1 Then
         With lstAccountDistribution.ListItems
            If Trim(.Item(iRow).SubItems(1)) = "" Then
                .Item(iRow).SubItems(1) = " "
                .Item(iRow).SubItems(2) = ""
                .Item(iRow).SubItems(3) = " "
                .Item(iRow).SubItems(4) = ""
            Else
                Set x = .Add()
                x.Text = " "
                x.SubItems(1) = " "
                x.SubItems(2) = " "
                x.SubItems(3) = " "
                x.SubItems(4) = " "
                iRow = .Count
            End If
        End With
        lstAccountDistribution.ListItems(iRow).EnsureVisible
        lstAccountDistribution.ListItems(iRow).Selected = True
        picMain.Enabled = False
        picToolbar.Enabled = False
        txtAccCode.Text = ""
        txtAccDesc.Text = ""
        txtDebit.Text = ""
        txtCredit.Text = ""
        picADSLine.ZOrder 0
        picADSLine.Visible = True
        TRANS_DETAIL = is_DET_ADDING
        txtAccCode.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSearch.Visible = True Then Exit Sub
    If picAdd.Visible = True Then Exit Sub
    If picExplainSLine.Visible = True Then Exit Sub
    If picADSLine.Visible = True Then Exit Sub
    If Statusbar1.Panels(1).Text = "" Then Exit Sub
    If AccessRights("Check Voucher", "Edit") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_EDITTING
Else
    If picSearch.Visible = True Then Exit Sub
    If picAdd.Visible = True Then Exit Sub
    If picExplainSLine.Visible = True Then Exit Sub
    If picADSLine.Visible = True Then Exit Sub
    If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
    If AccessRights("Check Voucher", "Edit") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If iFocusE = 1 Then
        With lstExplanation.ListItems
            txtSLineDesc.Text = .Item(iRow).SubItems(1)
            txtRRNumber.Text = .Item(iRow).SubItems(2)
            txtInvoiceNumber.Text = .Item(iRow).SubItems(3)
            txtSLineAmt.Text = .Item(iRow).SubItems(4)
            txtSLineDesc1.Text = .Item(iRow).SubItems(1)
            txtRRNumber1.Text = .Item(iRow).SubItems(2)
            txtInvoiceNumber1.Text = .Item(iRow).SubItems(3)
            txtSLineAmt1.Text = .Item(iRow).SubItems(4)
        End With
        picMain.Enabled = False
        picToolbar.Enabled = False
        picExplainSLine.ZOrder 0
        picExplainSLine.Visible = True
        TRANS_DETAIL = is_DET_EDITTING
        txtSLineDesc.SetFocus
        Exit Sub
    End If
    If iFocusA = 1 Then
        With lstAccountDistribution.ListItems
            txtAccCode.Text = .Item(iRow).SubItems(1)
            txtAccDesc.Text = .Item(iRow).SubItems(2)
            txtDebit.Text = .Item(iRow).SubItems(3)
            txtCredit.Text = .Item(iRow).SubItems(4)
            txtAccCode1.Text = .Item(iRow).SubItems(1)
            txtAccDesc1.Text = .Item(iRow).SubItems(2)
            txtDebit1.Text = .Item(iRow).SubItems(3)
            txtCredit1.Text = .Item(iRow).SubItems(4)
        End With
        picMain.Enabled = False
        picToolbar.Enabled = False
        picADSLine.ZOrder 0
        picADSLine.Visible = True
        TRANS_DETAIL = is_DET_EDITTING
        txtAccCode.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub PRESS_DELETE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSearch.Visible = True Then Exit Sub
    If picAdd.Visible = True Then Exit Sub
    If picExplainSLine.Visible = True Then Exit Sub
    If picADSLine.Visible = True Then Exit Sub
    If Statusbar1.Panels(1).Text = "" Then Exit Sub
    If AccessRights("Check Voucher", "Delete") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
    If MsgBox("ARE YOU SURE IN DELETING THIS TRANSACTION?                           ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    On Error GoTo PG:
    ConnOmega.Execute "DELETE FROM tbl_Acctg_CheckVoucher WHERE (PK = " & Statusbar1.Panels(1) & ")"
    CLEARTEXT
Else
    If picSearch.Visible = True Then Exit Sub
    If picAdd.Visible = True Then Exit Sub
    If picExplainSLine.Visible = True Then Exit Sub
    If picADSLine.Visible = True Then Exit Sub
    If AccessRights("Check Voucher", "Delete") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If iFocusE = 1 Then
        With lstExplanation.ListItems
            If .Count > 1 Then
                .Remove iRow
                If CDbl(iRow) > CDbl(.Count) Then
                    iRow = .Count
                End If
            Else
                .Item(1).SubItems(1) = " "
                .Item(1).SubItems(2) = " "
                iRow = 1
            End If
            lstExplanation.ListItems(iRow).EnsureVisible
            lstExplanation.ListItems(iRow).Selected = True
        End With
    End If
    If iFocusA = 1 Then
        With lstAccountDistribution.ListItems
            If .Count > 1 Then
                .Remove iRow
                If CDbl(iRow) > CDbl(.Count) Then
                    iRow = .Count
                End If
            Else
                .Item(1).SubItems(1) = " "
                .Item(1).SubItems(2) = " "
                .Item(1).SubItems(3) = " "
                .Item(1).SubItems(4) = " "
                iRow = 1
            End If
            lstAccountDistribution.ListItems(iRow).EnsureVisible
            lstAccountDistribution.ListItems(iRow).Selected = True
        End With
    End If
End If
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If iCVNoAuto = 0 Then If Trim(txtCVNumber.Text) = "" Then MsgBox "Please Supply CV Number!                ", vbCritical, "Error...": txtCVNumber.SetFocus: Exit Sub
If IsDate(txtCVDate.Text) = False Then MsgBox "Please Supply a Valid Date!                                ", vbCritical, "Error...": txtCVDate.SetFocus: Exit Sub
If iPayee = 0 Then MsgBox "Please Select Payee!                         ", vbCritical, "Error...": Exit Sub
If Trim(txtApproved.Text) = "" Then MsgBox "Please Supply Approved By!                      ", vbCritical, "Error...": txtApproved.SetFocus: Exit Sub
If Trim(txtChecked.Text) = "" Then MsgBox "Please Supply Checked By!                      ", vbCritical, "Error...": txtChecked.SetFocus: Exit Sub
If Trim(txtCheckDate.Text) <> "" Then
    If IsDate(txtCheckDate.Text) = False Then MsgBox "Please Supply a Valid Date Format!                      ", vbCritical, "Error...": txtCheckDate.SetFocus: Exit Sub
End If
If TRANSACTIONTYPE = is_ADDING Then
    If iCVNoAuto = 0 Then
        sCVNumber = Format(RETURNTEXTVALUE(txtCVNumber), "0000000#")
        s = "SELECT tbl_Acctg_CheckVoucher.* " & _
            " FROM tbl_Acctg_CheckVoucher " & _
            " WHERE (CVNumber = '" & sCVNumber & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            MsgBox "CV Number '" & sCVNumber & "' found duplicate!                   ", vbCritical, "Error..."
            rs.Close
            Exit Sub
        End If
        rs.Close
    End If
    If iCVNoAuto = 1 Then
        sCVNumber = ""
        s = "SELECT TOP 1 CVNumber " & _
            " FROM tbl_Acctg_CheckVoucher " & _
            " WHERE (Year(CVDate) = " & Format(FormatDateTime(txtCVDate.Text, vbShortDate), "yyyy") & ") " & _
            " AND (CVNumberAuto = 1)"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            sCVNumber = Format(CDbl(rs!CVNumber) + 1, "0000000#")
        Else
            sCVNumber = Format(FormatDateTime(txtCVDate.Text, vbShortDate), "yyyy") & "0000"
        End If
        rs.Close
        
        Do
            s = "SELECT tbl_Acctg_CheckVoucher.* " & _
                " FROM tbl_Acctg_CheckVoucher " & _
                " WHERE (CVNumber = '" & sCVNumber & "')"
            If rs.State = adStateOpen Then rs.Close
            rs.Open s, ConnOmega
            If rs.RecordCount = 0 Then
                rs.Close
                Exit Do
            End If
            rs.Close
            sCVNumber = Format(CDbl(sCVNumber) + 1, "0000000#")
        Loop
    End If
    ConnOmega.Execute "INSERT INTO tbl_Acctg_CheckVoucher " & _
                      " (CVNumber, CVNumberAuto, CVDate, PayeeKey, ORNumber, CheckNumber, " & _
                      " TotalAmt, Approved, Prepared, Checked, Entered, LastModified, PayeeType, PayeeName) " & _
                      " VALUES ('" & sCVNumber & "', " & iCVNoAuto & ", '" & FormatDateTime(txtCVDate.Text, vbShortDate) & "', " & _
                      " " & iPayee & ", '" & FORMATSQL(Trim(txtORNumber.Text)) & "', '" & FORMATSQL(Trim(txtCheckNumber.Text)) & "', " & _
                      " " & RETURNLABELVALUE(lblTotal) & ", '" & FORMATSQL(Trim(txtApproved.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtPrepared.Text)) & "', '" & FORMATSQL(Trim(txtChecked.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtEntered.Text)) & "', '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                      " " & iPayeeType & ", '" & FORMATSQL(Trim(txtPayeeName.Text)) & "')"
    iPK = 0
    s = "SELECT PK " & _
        " FROM tbl_Acctg_CheckVoucher " & _
        " WHERE (CVNumber = '" & sCVNumber & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        iPK = rs!PK
    End If
    rs.Close
    
    SaveSetting App.EXEName, "CVApprovedBy", "CVApproved", Trim(txtApproved.Text)
    SaveSetting App.EXEName, "CVCheckedBy", "CVChecked", Trim(txtChecked.Text)
    
End If
If TRANSACTIONTYPE = is_EDITTING Then
    iPK = Statusbar1.Panels(1).Text
    sCVNumber = Trim(txtCVNumber.Text)
    ConnOmega.Execute "UPDATE tbl_Acctg_CheckVoucher " & _
                      " SET CVNumber = '" & sCVNumber & "', " & _
                      " CVDate = '" & FormatDateTime(txtCVDate.Text, vbShortDate) & "', " & _
                      " PayeeKey = " & iPayee & ", ORNumber = '" & FORMATSQL(Trim(txtORNumber.Text)) & "', " & _
                      " CheckNumber = '" & FORMATSQL(Trim(txtCheckNumber.Text)) & "', " & _
                      " TotalAmt = " & RETURNLABELVALUE(lblTotal) & ", Approved = '" & FORMATSQL(Trim(txtApproved.Text)) & "', " & _
                      " Prepared = '" & FORMATSQL(Trim(txtPrepared.Text)) & "', " & _
                      " Checked = '" & FORMATSQL(Trim(txtChecked.Text)) & "', " & _
                      " Entered = '" & FORMATSQL(Trim(txtEntered.Text)) & "', " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                      " PayeeName = '" & FORMATSQL(Trim(txtPayeeName.Text)) & "', " & _
                      " WHERE (PK = " & iPK & ")"
End If


If CDbl(iPK) > 0 Then
    If IsDate(txtCheckDate.Text) = True Then
        ConnOmega.Execute "UPDATE tbl_Acctg_CheckVoucher " & _
                          " SET CheckDate = '" & FormatDateTime(txtCheckDate.Text, vbShortDate) & "'" & _
                          " WHERE (PK = " & iPK & ")"
    End If
    If IsDate(txtORDate.Text) = True Then
        ConnOmega.Execute "UPDATE tbl_Acctg_CheckVoucher " & _
                          " SET ORDate = '" & FormatDateTime(txtORDate.Text, vbShortDate) & "'" & _
                          " WHERE (PK = " & iPK & ")"
    End If
    
    With lstExplanation.ListItems
        l = 0
        ConnOmega.Execute "DELETE FROM tbl_Acctg_CheckVoucher_Detail WHERE (CVKey = " & iPK & ")"
        For i = 1 To .Count
            If Trim(.Item(i).SubItems(1)) <> "" Then
                l = l + 1
                ConnOmega.Execute "INSERT INTO tbl_Acctg_CheckVoucher_Detail " & _
                                  " (CVKey, Line, Description, RRNumber, InvNumber, Amt) " & _
                                  " VALUES (" & iPK & ", " & l & ", " & _
                                  " '" & FORMATSQL(Trim(.Item(i).SubItems(1))) & "', " & _
                                  " '" & FORMATSQL(Trim(.Item(i).SubItems(2))) & "', " & _
                                  " '" & FORMATSQL(Trim(.Item(i).SubItems(3))) & "', " & _
                                  " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(4)) = False, 0, .Item(i).SubItems(4))) & ")"
            End If
        Next i
    End With
    With lstAccountDistribution.ListItems
        l = 0
        ConnOmega.Execute "DELETE FROM tbl_Acctg_CheckVoucher_AD WHERE (CVKey = " & iPK & ")"
        For i = 1 To .Count
            If Trim(.Item(i).SubItems(1)) <> "" Then '
                l = l + 1
                ConnOmega.Execute "INSERT INTO tbl_Acctg_CheckVoucher_AD " & _
                                  " (CVKey, Line, AccCode, Debit, Credit) " & _
                                  " VALUES (" & iPK & ", " & l & ", " & _
                                  " '" & FORMATSQL(Trim(.Item(i).SubItems(1))) & "', " & _
                                  " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(3)) = False, 0, .Item(i).SubItems(3))) & ", " & _
                                  " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(4)) = False, 0, .Item(i).SubItems(4))) & ")"
            End If
        Next i
    End With
End If

SaveSetting App.EXEName, "CV_Approved", "CV_Approved", Trim(txtApproved.Text)
SaveSetting App.EXEName, "CV_Checked", "CV_Checked", Trim(txtChecked.Text)

CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANS_DETAIL = is_DET_REFRESH
BROWSER sCVNumber, "is_LOAD"
txtCVNumber.SetFocus
End Sub

Private Sub PRESS_F6()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSearch.Visible = True Then Exit Sub
    If picADSLine.Visible = True Then Exit Sub
    If picAdd.Visible = True Then Exit Sub
    PopupMenu MainFormPopupF.mnuCVFind, , Toolbar1.Buttons(15).Left, Toolbar1.Buttons(15).Top + Toolbar1.Buttons(15).Height
Else
    If iFocusAcc = 1 Then
        picADSLine.Enabled = False
        txtSearchAccount.Text = ""
        picSeachAccount.ZOrder 0
        picSeachAccount.Visible = True
        txtSearchAccount.SetFocus
    End If
End If
End Sub

Private Sub PRESS_F8()
If picSearch.Visible = True Then Exit Sub
If picADSLine.Visible = True Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
If Trim(txtCheckNumber.Text) = "" Then MsgBox "Please Supply Checked Number!                      ", vbCritical, "Error...": txtCheckNumber.SetFocus: Exit Sub
If IsDate(txtCheckDate.Text) = False Then MsgBox "Please Supply a Valid Date Format!                      ", vbCritical, "Error...": txtCheckDate.SetFocus: Exit Sub
'If Trim(txtORNumber.Text) = "" Then MsgBox "Please Supply OR Number!                              ", vbCritical, "Error...": txtORNumber.SetFocus: Exit Sub
'If IsDate(txtORDate.Text) = False Then MsgBox "Please Supply a Valid Date!                                ", vbCritical, "Error...": txtORDate.SetFocus: Exit Sub
If MsgBox("CONTINUE POSTING THIS TRANSACTION?                       ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub

If iPayeeType = 1 Then
    
    With lstAccountDistribution.ListItems
        sBankCode = "": sBankName = ""
        For i = 1 To .Count
            If CDbl(.Item(i).SubItems(1)) >= 101101 And _
            CDbl(.Item(i).SubItems(2)) <= 101105 Then
                sBankCode = .Item(i).SubItems(1)
                sBankName = .Item(i).SubItems(2)
                Exit For
            End If
        Next i
        For i = 1 To .Count
            If Trim(.Item(i).SubItems(1)) <> "" Then
                'dAmount = CDbl(IIf(IsNumeric(.Item(i).SubItems(3)) = False, 0, .Item(i).SubItems(3))) - CDbl(IIf(IsNumeric(.Item(i).SubItems(4)) = False, 0, .Item(i).SubItems(4)))
                ConnOmega.Execute "INSERT INTO tbl_GL_Transaction " & _
                                  " (GLCode, DocDate, DocNumber, PayeeKey, PayeeType, SupplierCode, " & _
                                  " SupplierName, CheckNumber, CheckDate, BookType, Debit, Credit, " & _
                                  " CVNumber, BankCode, BankName) " & _
                                  " VALUES ('" & .Item(i).SubItems(1) & "', '" & FormatDateTime(txtCVDate.Text, vbShortDate) & "', " & _
                                  " '" & Trim(txtCVNumber.Text) & "', " & iPayee & ", " & iPayeeType & ", '" & Trim(txtPayeeCode.Text) & "', " & _
                                  " '" & FORMATSQL(Trim(txtPayeeName.Text)) & "', '" & Trim(txtCheckNumber.Text) & "', " & _
                                  " '" & FormatDateTime(txtCheckDate.Text, vbShortDate) & "', 1, " & _
                                  " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(3)) = False, 0, .Item(i).SubItems(3))) & ", " & _
                                  " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(4)) = False, 0, .Item(i).SubItems(4))) & ", " & _
                                  " '" & Trim(txtCVNumber.Text) & "', '" & FORMATSQL(Trim(CStr(sBankCode))) & "', " & _
                                  " '" & FORMATSQL(Trim(CStr(sBankName))) & "')"
                
                t = "SELECT tbl_GL_Accounts.* " & _
                    " FROM tbl_GL_Accounts " & _
                    " WHERE (AccountCode = '" & .Item(i).SubItems(1) & "')"
                If rt.State = adStateOpen Then rt.Close
                rt.Open t, ConnOmega
                If rt.RecordCount > 0 Then
                    If rt!withSL = 1 Then
                        If rt!SupplierKey = 0 Then
                            ConnOmega.Execute "INSERT INTO tbl_Inv_Supplier_SL " & _
                                              " (SupplierKey, GLCode, DocNumber, DocDate, " & _
                                              " Description, CheckNumber, Reference, iType, Debit, Credit) " & _
                                              " VALUES (" & iPayee & ", '" & .Item(i).SubItems(1) & "', " & _
                                              " '" & Trim(txtCVNumber.Text) & "', " & _
                                              " '" & FormatDateTime(txtCVDate.Text, vbShortDate) & "', " & _
                                              " '', '" & Trim(txtCheckNumber.Text) & "', " & _
                                              " '" & "CV# " & Trim(txtCVNumber.Text) & " CHK# " & Trim(txtCheckNumber.Text) & "', " & _
                                              " 3, " & CDbl(IIf(IsNumeric(.Item(i).SubItems(3)) = False, 0, .Item(i).SubItems(3))) & ", " & _
                                              " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(4)) = False, 0, .Item(i).SubItems(4))) & ")"
                        Else
                            ConnOmega.Execute "INSERT INTO tbl_Inv_Supplier_SL " & _
                                              " (SupplierKey, GLCode, DocNumber, DocDate, " & _
                                              " Description, CheckNumber, Reference, iType, Debit, Credit) " & _
                                              " VALUES (" & rt!SupplierKey & ", '" & .Item(i).SubItems(1) & "', " & _
                                              " '" & Trim(txtCVNumber.Text) & "', " & _
                                              " '" & FormatDateTime(txtCVDate.Text, vbShortDate) & "', " & _
                                              " '', '" & Trim(txtCheckNumber.Text) & "', " & _
                                              " '" & "CV# " & Trim(txtCVNumber.Text) & " CHK# " & Trim(txtCheckNumber.Text) & "', " & _
                                              " 3, " & CDbl(IIf(IsNumeric(.Item(i).SubItems(3)) = False, 0, .Item(i).SubItems(3))) & ", " & _
                                              " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(4)) = False, 0, .Item(i).SubItems(4))) & ")"
                        End If
                    End If
                End If
                rt.Close
            End If
        Next i
    End With
    
End If

ConnOmega.Execute "UPDATE tbl_Acctg_CheckVoucher SET Posted = 1 WHERE (PK = " & Statusbar1.Panels(1).Text & ")"

BROWSER GetSetting(App.EXEName, "CVNumber", "CVNum", ""), "is_LOAD"

End Sub

Private Sub PRESS_F9()
If picSearch.Visible = True Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picExplainSLine.Visible = True Then Exit Sub
If picADSLine.Visible = True Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If Trim(txtCheckNumber.Text) = "" Then MsgBox "Please Supply Checked Number!                      ", vbCritical, "Error...": txtCheckNumber.SetFocus: Exit Sub
If IsDate(txtCheckDate.Text) = False Then MsgBox "Please Supply a Valid Date Format!                      ", vbCritical, "Error...": txtCheckDate.SetFocus: Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub

PopupMenu MainFormPopupF.mnuCVPrint, , Toolbar1.Buttons(17).Left, Toolbar1.Buttons(17).Top + Toolbar1.Buttons(17).Height

End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picAdd.Visible = True Then cmdCancelAdd_Click: Exit Sub
    If picSearch.Visible = True Then cmdCancelSearch_Click: Exit Sub
    Unload Me
Else
    If picSeachAccount.Visible = True Then
        cmdCancelAccount_Click
        Exit Sub
    End If
    If picExplainSLine.Visible = True Then
        With lstExplanation.ListItems
            If TRANS_DETAIL = is_DET_ADDING Then
                If .Count > 1 Then
                    .Remove .Count
                Else
                    .Item(1).SubItems(1) = " "
                    .Item(1).SubItems(2) = " "
                    .Item(1).SubItems(3) = " "
                    .Item(1).SubItems(4) = " "
                End If
                iRow = .Count
            End If
            If TRANS_DETAIL = is_DET_EDITTING Then
                .Item(iRow).SubItems(1) = txtSLineDesc1.Text
                .Item(iRow).SubItems(2) = txtRRNumber1.Text
                .Item(iRow).SubItems(3) = txtInvoiceNumber1.Text
                .Item(iRow).SubItems(4) = txtSLineAmt1.Text
            End If
        End With
        lstExplanation.ListItems(iRow).EnsureVisible
        lstExplanation.ListItems(iRow).Selected = True
        picExplainSLine.Visible = False
        picMain.Enabled = True
        picToolbar.Enabled = True
        lstExplanation.SetFocus
        Exit Sub
    End If
    If picADSLine.Visible = True Then
        With lstAccountDistribution.ListItems
            If TRANS_DETAIL = is_DET_ADDING Then
                If .Count > 1 Then
                    .Remove .Count
                Else
                    .Item(1).SubItems(1) = " "
                    .Item(1).SubItems(2) = " "
                    .Item(1).SubItems(3) = " "
                    .Item(1).SubItems(4) = " "
                End If
                iRow = .Count
            End If
            If TRANS_DETAIL = is_DET_EDITTING Then
                .Item(iRow).SubItems(1) = txtAccCode1.Text
                .Item(iRow).SubItems(2) = txtAccDesc1.Text
                .Item(iRow).SubItems(3) = txtDebit1.Text
                .Item(iRow).SubItems(4) = txtCredit1.Text
            End If
        End With
        lstAccountDistribution.ListItems(iRow).EnsureVisible
        lstAccountDistribution.ListItems(iRow).Selected = True
        picADSLine.Visible = False
        picMain.Enabled = True
        picToolbar.Enabled = True
        lstAccountDistribution.SetFocus
        Exit Sub
    End If
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    TRANS_DETAIL = is_DET_REFRESH
    BROWSER GetSetting(App.EXEName, "CVNumber", "CVNum", ""), "is_LOAD"
    If Trim(txtCVNumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "CVNumber", "CVNum", ""), "is_HOME"
    txtCVNumber.SetFocus
End If
End Sub

Private Sub b8TitleBar1_CLoseClick()
cmdCancelAccount_Click
End Sub

Private Sub b8TitleBar2_CLoseClick()
cmdCancelAdd_Click
End Sub

Private Sub b8TitleBar3_CLoseClick()
cmdCancelSearch_Click
End Sub

Private Sub cmdCancelAccount_Click()
picSeachAccount.Visible = False
picADSLine.Enabled = True
txtAccDesc.SetFocus
End Sub

Private Sub cmdCancelAdd_Click()
picAdd.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True
End Sub

Private Sub cmdCancelSearch_Click()
picSearch.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True
End Sub

Private Sub cmdOKAccount_Click()
If lstAccount.ListIndex = -1 Then Exit Sub
Arr = Split(lstAccount.List(lstAccount.ListIndex), " | ", -1, 1)
txtAccCode.Text = CStr(Arr(0))
txtAccDesc.Text = CStr(Arr(1))
cmdCancelAccount_Click
End Sub

Private Sub cmdOKAdd_Click()
If lstResultAdd.ListIndex = -1 Then Exit Sub
iCVNoAuto = 0
s = "SELECT TOP 1 tbl_Acctg_CVNumberAuto.* " & _
    " FROM tbl_Acctg_CVNumberAuto " & _
    " WHERE (EffectDate <= '" & Format(Date, "mm/dd/yyyy") & "') " & _
    " ORDER BY EffectDate DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iCVNoAuto = rs!Automatic
End If
rs.Close
CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
picMain.Enabled = True
picToolbar.Enabled = True
picAdd.Visible = False
iPayee = lstResultAdd.ItemData(lstResultAdd.ListIndex)
Arr = Split(lstResultAdd.List(lstResultAdd.ListIndex), " - ", -1, 1)
txtPayeeCode.Text = CStr(Arr(0))
txtPayeeName.Text = CStr(Arr(1))
txtCVDate.Text = Format(Date, "mm/dd/yyyy")
txtApproved.Text = GetSetting(App.EXEName, "CV_Approved", "CV_Approved", "JAC")
txtChecked.Text = GetSetting(App.EXEName, "CV_Checked", "CV_Checked", "FLS/MZA")
If iPayeeType = 1 Then
    iTotalAmount = 0
    s = "SELECT DocNumber, SUM(Balance) AS Amount " & _
        " From tbl_Inv_Supplier_SL " & _
        " Where (SupplierKey = " & iPayee & ") " & _
        " GROUP BY DocNumber " & _
        " HAVING (SUM(Balance) < 0) " & _
        " ORDER BY DocNumber"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        lstExplanation.ListItems.Clear
        While Not rs.EOF
            iAmount = CDbl(rs!Amount) * -1
            iTotalAmount = iTotalAmount + CDbl(iAmount)
            Set x = lstExplanation.ListItems.Add()
            x.Text = " "
            t = "SELECT Reference, DocNumber, InvoiceNumber " & _
                " From tbl_Inv_Supplier_SL " & _
                " WHERE ((DocNumber = '" & rs!DocNumber & "') " & _
                " AND (iType = 0)) " & _
                " OR ((DocNumber = '" & rs!DocNumber & "') " & _
                " AND (iType = 2))"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                x.SubItems(1) = IIf(IsNull(rt!Reference), rt!DocNumber, rt!Reference)
                x.SubItems(2) = IIf(IsNull(rt!DocNumber), " ", rt!DocNumber)
                x.SubItems(3) = IIf(IsNull(rt!InvoiceNumber), " ", rt!InvoiceNumber)
            Else
                x.SubItems(1) = " "
                x.SubItems(2) = " "
                x.SubItems(3) = " "
            End If
            rt.Close
            x.SubItems(4) = Format(iAmount, "#,##0.00")
            rs.MoveNext
        Wend
    End If
    rs.Close
    lblTotal.Caption = Format(iTotalAmount, "#,##0.00")
End If
If iCVNoAuto = 0 Then txtCVNumber.Locked = False: txtCVNumber.SetFocus Else txtCVNumber.Locked = True: lstExplanation.SetFocus
txtPrepared.Text = gbl_UserName
txtEntered.Text = gbl_UserName
End Sub




Private Sub cmdOKSearch_Click()
If lstSearch.ListIndex = -1 Then Exit Sub
BROWSER lstSearch.ItemData(lstSearch.ListIndex), "is_FIND"
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
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "CVNumber", "CVNum", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "CVNumber", "CVNum", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "CVNumber", "CVNum", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "CVNumber", "CVNum", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
picADSLine.Width = 9375
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
'Me.Caption = "Check Voucher"
iFocusA = 0
iFocusE = 0
iRow = 0
iFocusAcc = 0
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANS_DETAIL = is_DET_REFRESH
BROWSER GetSetting(App.EXEName, "CVNumber", "CVNum", ""), "is_LOAD"
If Trim(txtCVNumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "CVNumber", "CVNum", ""), "is_HOME"
tmp = SetWindowLong(txtSLineDesc.hwnd, GWL_STYLE, GetWindowLong(txtSLineDesc.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearchAdd.hwnd, GWL_STYLE, GetWindowLong(txtSearchAdd.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearchAccount.hwnd, GWL_STYLE, GetWindowLong(txtSearchAccount.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtApproved.hwnd, GWL_STYLE, GetWindowLong(txtApproved.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPrepared.hwnd, GWL_STYLE, GetWindowLong(txtPrepared.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtChecked.hwnd, GWL_STYLE, GetWindowLong(txtChecked.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtEntered.hwnd, GWL_STYLE, GetWindowLong(txtEntered.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picAdd.Visible = True Then Cancel = -1
If picSeachAccount.Visible = True Then Cancel = -1
If picExplainSLine.Visible = True Then Cancel = -1
If picADSLine.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lblTotal_Change()
If RETURNLABELVALUE(lblTotal) > 0 Then
    txtAmtWords.Text = AMT2WORDS(RETURNLABELVALUE(lblTotal))
End If
End Sub

Private Sub lstAccount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAccount_Click
End Sub

Private Sub lstAccountDistribution_Click()
iFocusA = 1
iRow = lstAccountDistribution.SelectedItem.Index
TRANS_DETAIL = is_DET_REFRESH
If imgPosted.Visible = True Then Exit Sub
With lstAccountDistribution.ListItems
    If .Count > 1 Then
        TOOLBARFUNC 5
    Else
        If Trim(.Item(iRow).SubItems(1)) = "" Then
            TOOLBARFUNC 4
        Else
            TOOLBARFUNC 5
        End If
    End If
    If Statusbar1.Panels(1).Text <> "" Then TRANSACTIONTYPE = is_EDITTING
End With
End Sub

Private Sub lstAccountDistribution_GotFocus()
iFocusA = 1
iRow = lstAccountDistribution.SelectedItem.Index
TRANS_DETAIL = is_DET_REFRESH
If imgPosted.Visible = True Then Exit Sub
With lstAccountDistribution.ListItems
    If .Count > 1 Then
        TOOLBARFUNC 5
    Else
        If Trim(.Item(iRow).SubItems(1)) = "" Then
            TOOLBARFUNC 4
        Else
            TOOLBARFUNC 5
        End If
    End If
    If Statusbar1.Panels(1).Text <> "" Then TRANSACTIONTYPE = is_EDITTING
End With
End Sub

Private Sub lstAccountDistribution_ItemClick(ByVal Item As MSComctlLib.ListItem)
iRow = lstAccountDistribution.SelectedItem.Index
End Sub

Private Sub lstAccountDistribution_LostFocus()
iFocusA = 0
End Sub

Private Sub lstExplanation_Click()
iFocusE = 1
iRow = lstExplanation.SelectedItem.Index
TRANS_DETAIL = is_DET_REFRESH
If imgPosted.Visible = True Then Exit Sub
With lstExplanation.ListItems
    If .Count > 1 Then
        TOOLBARFUNC 5
    Else
        If Trim(.Item(iRow).SubItems(1)) = "" Then
            TOOLBARFUNC 4
        Else
            TOOLBARFUNC 5
        End If
    End If
    If Statusbar1.Panels(1).Text <> "" Then TRANSACTIONTYPE = is_EDITTING
End With
End Sub

Private Sub lstExplanation_GotFocus()
iFocusE = 1
iRow = lstExplanation.SelectedItem.Index
TRANS_DETAIL = is_DET_REFRESH
If imgPosted.Visible = True Then Exit Sub
With lstExplanation.ListItems
    If .Count > 1 Then
        TOOLBARFUNC 5
    Else
        If Trim(.Item(iRow).SubItems(1)) = "" Then
            TOOLBARFUNC 4
        Else
            TOOLBARFUNC 5
        End If
    End If
    If Statusbar1.Panels(1).Text <> "" Then TRANSACTIONTYPE = is_EDITTING
End With
End Sub

Private Sub lstExplanation_ItemClick(ByVal Item As MSComctlLib.ListItem)
iRow = lstExplanation.SelectedItem.Index
End Sub

Private Sub lstExplanation_LostFocus()
iFocusE = 0
End Sub

Private Sub lstResultAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAdd_Click
End Sub

Private Sub lstSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKSearch_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "CVNumber", "CVNum", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "CVNumber", "CVNum", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "CVNumber", "CVNum", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "CVNumber", "CVNum", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":   PRESS_F9
    Case "Post":    PRESS_F8
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtAccCode_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstAccountDistribution.ListItems
        .Item(iRow).SubItems(1) = Trim(txtAccCode.Text)
    End With
End If
End Sub

Private Sub txtAccCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtAccDesc.SetFocus
End Sub

Private Sub txtAccCode_LostFocus()
If Trim(txtAccCode.Text) = "" Then Exit Sub
s = "SELECT tbl_GL_Accounts.* " & _
    " FROM tbl_GL_Accounts " & _
    " WHERE (AccountCode = '" & Trim(txtAccCode.Text) & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount = 0 Then MsgBox "Account Code '" & Trim(txtAccCode.Text) & "' not found!                       ", vbCritical, "Error...": txtAccCode.SetFocus: rs.Close: Exit Sub
txtAccCode.Text = rs!AccountCode
txtAccDesc.Text = rs!AccountName
rs.Close
End Sub

Private Sub txtAccDesc_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstAccountDistribution.ListItems
        .Item(iRow).SubItems(2) = Trim(txtAccDesc.Text)
    End With
End If
End Sub

Private Sub txtAccDesc_GotFocus()
iFocusAcc = 1
End Sub

Private Sub txtAccDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtDebit.SetFocus
End Sub

Private Sub txtAccDesc_LostFocus()
iFocusAcc = 0
End Sub

Private Sub txtCredit_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstAccountDistribution.ListItems
        .Item(iRow).SubItems(4) = Format(RETURNTEXTVALUE(txtCredit), "#,##0.00")
        l = 0
        For i = 1 To .Count
            l = l + CDbl(IIf(IsNumeric(.Item(i).SubItems(4)) = False, 0, .Item(i).SubItems(4)))
        Next i
    End With
    lblCredit.Caption = Format(l, "#,##0.00")
End If
End Sub

Private Sub txtCredit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If TRANS_DETAIL = is_DET_ADDING Then
        TRANS_DETAIL = is_DET_REFRESH
        With lstAccountDistribution.ListItems
            If Trim(.Item(iRow).SubItems(1)) = "" Then
                .Item(iRow).SubItems(1) = " "
                .Item(iRow).SubItems(2) = ""
                .Item(iRow).SubItems(3) = " "
                .Item(iRow).SubItems(4) = ""
            Else
                Set x = .Add()
                x.Text = " "
                x.SubItems(1) = " "
                x.SubItems(2) = " "
                x.SubItems(3) = " "
                x.SubItems(4) = " "
                iRow = .Count
            End If
        End With
        lstAccountDistribution.ListItems(iRow).EnsureVisible
        lstAccountDistribution.ListItems(iRow).Selected = True
        txtAccCode.Text = ""
        txtAccDesc.Text = ""
        txtDebit.Text = ""
        txtCredit.Text = ""
        TRANS_DETAIL = is_DET_ADDING
        txtAccCode.SetFocus
    End If
    If TRANS_DETAIL = is_DET_EDITTING Then
        picADSLine.Visible = False
        picMain.Enabled = True
        picToolbar.Enabled = True
        lstAccountDistribution.SetFocus
    End If
End If
End Sub

Private Sub txtCredit_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub



Private Sub txtCVDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstExplanation.SetFocus
End Sub

Private Sub txtCVNumber_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtCVDate.SetFocus
End Sub

Private Sub txtCVNumber_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtDebit_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstAccountDistribution.ListItems
        .Item(iRow).SubItems(3) = Format(RETURNTEXTVALUE(txtDebit), "#,##0.00")
        l = 0
        For i = 1 To .Count
            l = l + CDbl(IIf(IsNumeric(.Item(i).SubItems(3)) = False, 0, .Item(i).SubItems(3)))
        Next i
    End With
    lblDebit.Caption = Format(l, "#,##0.00")
End If
End Sub

Private Sub txtDebit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtCredit.SetFocus
End Sub

Private Sub txtDebit_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtInvoiceNumber_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstExplanation.ListItems
        .Item(iRow).SubItems(3) = Trim(txtInvoiceNumber.Text)
    End With
End If
End Sub

Private Sub txtRRNumber_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstExplanation.ListItems
        .Item(iRow).SubItems(2) = Trim(txtRRNumber.Text)
    End With
End If
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstSearch.Clear: Exit Sub
lstSearch.Clear
Select Case SearchType
    Case 1  'CV Number
        s = "SELECT PK, CVNumber as Code, PayeeKey, PayeeType " & _
            " From tbl_Acctg_CheckVoucher " & _
            " WHERE (CVNumber LIKE '%" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
            " ORDER BY CVNumber"
    Case 2  'Check Number
        s = "SELECT PK, CheckNumber as Code, PayeeKey, PayeeType " & _
            " From tbl_Acctg_CheckVoucher " & _
            " WHERE (CheckNumber LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
            " ORDER BY CheckNumber"
    Case Else: Exit Sub
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    sSearchPayeeName = ""
    Select Case rs!PayeeType
        Case 1
            t = "SELECT SupplierCode, SupplierName " & _
                " From tbl_Inv_Supplier " & _
                " WHERE (PK = " & rs!PayeeKey & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                sSearchPayeeName = rt!SupplierCode & " - " & rt!SupplierName
            End If
            rt.Close
        Case 2
            t = "SELECT PK, LastName + ',  ' + FirstName + '  ' + MiddleName AS sName, " & _
                " ISNULL((SELECT TOP 1 IDNumber From tbl_Member_IDNumber " & _
                " Where (MemberKey = tbl_Member_Information.PK) " & _
                " ORDER BY IDCounter DESC, IDNumber), '') AS Code " & _
                " From tbl_Member_Information " & _
                " WHERE (PK = " & rs!PayeeKey & ") " & _
                " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                sSearchPayeeName = rt!Code & " - " & rt!sName
            End If
            rt.Close
        Case 3
            t = "SELECT PK, LastName + ',  ' + FirstName + '  ' + MiddleName AS sName, " & _
                " ISNULL((SELECT TOP 1 IDNumber From tbl_Personnel_IDNumber " & _
                " Where (ProfileKey = tbl_Personnel_Information.PK) " & _
                " ORDER BY IDNumber DESC), '') AS Code " & _
                " From tbl_Personnel_Information " & _
                " WHERE (PK = " & rs!PayeeKey & ") " & _
                " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                sSearchPayeeName = rt!Code & " - " & rt!sName
            End If
            rt.Close
    End Select
    lstSearch.AddItem rs!Code & " | " & sSearchPayeeName
    lstSearch.ItemData(lstSearch.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstSearch.SetFocus
End Sub

Private Sub txtSearchAccount_Change()
If Trim(txtSearchAccount.Text) = "" Then lstAccount.Clear: Exit Sub
lstAccount.Clear
s = "SELECT AccountCode, AccountName " & _
    " From tbl_GL_Accounts " & _
    " WHERE (AccountName LIKE '" & FORMATSQL(Trim(txtSearchAccount.Text)) & "%') " & _
    " ORDER BY AccountName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstAccount.AddItem rs!AccountCode & " | " & rs!AccountName
    rs.MoveNext
Wend
rs.Close
If lstAccount.ListCount Then lstAccount.ListIndex = 0
End Sub

Private Sub txtSearchAccount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstAccount.SetFocus
End Sub

Private Sub txtSearchAdd_Change()
If Trim(txtSearchAdd.Text) = "" Then lstResultAdd.Clear: Exit Sub
lstResultAdd.Clear
Select Case iPayeeType
    Case 1
        s = "SELECT PK, SupplierCode as Code, SupplierName as sName " & _
            " From tbl_Inv_Supplier " & _
            " WHERE (SupplierName LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%')"
    Case 2
        s = "SELECT PK, LastName + ',  ' + FirstName + '  ' + MiddleName AS sName, " & _
            " ISNULL((SELECT TOP 1 IDNumber From tbl_Member_IDNumber " & _
            " Where (MemberKey = tbl_Member_Information.PK) " & _
            " ORDER BY IDCounter DESC, IDNumber), '') AS Code " & _
            " From tbl_Member_Information " & _
            " WHERE (LastName LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%') " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
    Case 3
        s = "SELECT PK, LastName + ',  ' + FirstName + '  ' + MiddleName AS sName, " & _
            " ISNULL((SELECT TOP 1 IDNumber From tbl_Personnel_IDNumber " & _
            " Where (ProfileKey = tbl_Personnel_Information.PK) " & _
            " ORDER BY IDNumber DESC), '') AS Code " & _
            " From tbl_Personnel_Information " & _
            " WHERE (LastName LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%') " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
    Case Else: Exit Sub
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultAdd.AddItem rs!Code & " - " & rs!sName
    lstResultAdd.ItemData(lstResultAdd.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResultAdd.ListCount Then lstResultAdd.ListIndex = 0
End Sub

Private Sub txtSearchAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultAdd.SetFocus
End Sub

Private Sub txtSLineAmt_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstExplanation.ListItems
        .Item(iRow).SubItems(4) = Format(RETURNTEXTVALUE(txtSLineAmt), "#,##0.00")
        l = 0
        For i = 1 To .Count
            l = l + CDbl(IIf(IsNumeric(.Item(i).SubItems(4)) = False, 0, .Item(i).SubItems(4)))
        Next i
        lblTotal.Caption = Format(l, "#,##0.00")
    End With
End If
End Sub

Private Sub txtSLineAmt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If TRANS_DETAIL = is_DET_ADDING Then
        TRANS_DETAIL = is_DET_REFRESH
        With lstExplanation.ListItems
            If Trim(.Item(iRow).SubItems(1)) = "" Then
                .Item(iRow).SubItems(1) = " "
                .Item(iRow).SubItems(2) = " "
                .Item(iRow).SubItems(3) = " "
                .Item(iRow).SubItems(4) = " "
            Else
                Set x = .Add()
                x.Text = " "
                x.SubItems(1) = " "
                x.SubItems(2) = " "
                x.SubItems(3) = " "
                x.SubItems(4) = " "
                iRow = .Count
            End If
        End With
        lstExplanation.ListItems(iRow).EnsureVisible
        lstExplanation.ListItems(iRow).Selected = True
        txtSLineDesc.Text = ""
        txtSLineAmt.Text = ""
        txtRRNumber.Text = ""
        txtInvoiceNumber.Text = ""
        TRANS_DETAIL = is_DET_ADDING
        txtSLineDesc.SetFocus
    End If
    If TRANS_DETAIL = is_DET_EDITTING Then
        picExplainSLine.Visible = False
        picMain.Enabled = True
        picToolbar.Enabled = True
        lstExplanation.SetFocus
    End If
End If
End Sub

Private Sub txtSLineDesc_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstExplanation.ListItems
        .Item(iRow).SubItems(1) = Trim(txtSLineDesc.Text)
    End With
End If
End Sub

Private Sub txtSLineDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtSLineAmt.SetFocus
End Sub
