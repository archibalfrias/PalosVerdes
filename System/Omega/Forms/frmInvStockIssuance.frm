VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInvStockIssuance 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvStockIssuance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   12255
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15600
      TabIndex        =   83
      Top             =   0
      Width           =   15600
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   84
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
            NumButtons      =   26
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
               Caption         =   "GL Acc"
               Key             =   "Accnt"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               Key             =   "Refresh"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         MousePointer    =   99
         MouseIcon       =   "frmInvStockIssuance.frx":08CA
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   10740
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   85
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmInvStockIssuance.frx":0BE4
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
   Begin VB.PictureBox picGLPosted 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   11160
      ScaleHeight     =   3255
      ScaleWidth      =   495
      TabIndex        =   74
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   82
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   81
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   80
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   79
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   78
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   77
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   76
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   75
         Top             =   2880
         Width           =   255
      End
   End
   Begin RPVGCC.b8Container picSearchGLAccount 
      Height          =   2955
      Left            =   3840
      TabIndex        =   68
      Top             =   2280
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5212
      BackColor       =   15396057
      Begin VB.CommandButton cmdOKGLAccount 
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
         Picture         =   "frmInvStockIssuance.frx":12F7
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   2355
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelGLAccount 
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
         Picture         =   "frmInvStockIssuance.frx":1969
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   2355
         Width           =   1560
      End
      Begin VB.TextBox txtSearchGLAccount 
         Height          =   315
         Left            =   120
         TabIndex        =   70
         Top             =   480
         Width           =   5295
      End
      Begin VB.ListBox lstResultGLAccount 
         Height          =   1425
         Left            =   120
         TabIndex        =   69
         Top             =   840
         Width           =   5295
      End
      Begin RPVGCC.b8TitleBar b8TitleBar3 
         Height          =   345
         Left            =   45
         TabIndex        =   73
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
         Icon            =   "frmInvStockIssuance.frx":20C5
         ShadowVisible   =   0   'False
      End
   End
   Begin VB.PictureBox picADSLine 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2520
      ScaleHeight     =   855
      ScaleWidth      =   7695
      TabIndex        =   54
      Top             =   1560
      Visible         =   0   'False
      Width           =   7695
      Begin RPVGCC.b8Container picADSLine1 
         Height          =   855
         Left            =   0
         TabIndex        =   55
         Top             =   0
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   1508
         BackColor       =   8438015
         Begin VB.TextBox txtCredit1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   63
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtDebit1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   62
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtAccountName1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   61
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtAccountNo1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   60
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtCredit 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5880
            TabIndex        =   59
            Top             =   360
            Width           =   1275
         End
         Begin VB.TextBox txtDebit 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4560
            TabIndex        =   58
            Top             =   360
            Width           =   1275
         End
         Begin VB.TextBox txtAccountName 
            Height          =   315
            Left            =   1320
            TabIndex        =   57
            Top             =   360
            Width           =   3195
         End
         Begin VB.TextBox txtAccountNo 
            Height          =   315
            Left            =   120
            TabIndex        =   56
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "CREDIT"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5880
            TabIndex        =   67
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DEBIT"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4560
            TabIndex        =   66
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "ACCOUNT NAME"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1320
            TabIndex        =   65
            Top             =   120
            Width           =   3135
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "ACCOUNT #"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   120
            Width           =   975
         End
      End
   End
   Begin VB.PictureBox picAccDistribution 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   2520
      ScaleHeight     =   3855
      ScaleWidth      =   7335
      TabIndex        =   34
      Top             =   1320
      Visible         =   0   'False
      Width           =   7335
      Begin RPVGCC.b8Container b8Container1 
         Height          =   3615
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6376
         BackColor       =   15396057
         Begin VB.CommandButton cmdPost 
            Caption         =   "P O S T"
            Height          =   495
            Left            =   120
            TabIndex        =   41
            Top             =   3000
            Width           =   1095
         End
         Begin VB.TextBox txtInvNetP 
            Height          =   315
            Left            =   5280
            TabIndex        =   40
            Top             =   720
            Width           =   1875
         End
         Begin VB.TextBox txtInvNumberP 
            Height          =   315
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Width           =   1515
         End
         Begin VB.TextBox txtInvDateP 
            Height          =   315
            Left            =   1680
            TabIndex        =   38
            Top             =   720
            Width           =   1515
         End
         Begin VB.TextBox txtInvGrossP 
            Height          =   315
            Left            =   3240
            TabIndex        =   37
            Top             =   720
            Width           =   1995
         End
         Begin VB.ComboBox cmbBookType 
            Height          =   315
            Left            =   2280
            TabIndex        =   36
            Text            =   "Combo1"
            Top             =   3080
            Width           =   1215
         End
         Begin MSComctlLib.ListView lstAccDistribution 
            Height          =   1815
            Left            =   120
            TabIndex        =   42
            Top             =   1080
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   3201
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
               Text            =   "Code"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Name"
               Object.Width           =   5821
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Debit"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Credit"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Amount"
               Object.Width           =   0
            EndProperty
         End
         Begin RPVGCC.b8TitleBar b8TitleBar2 
            Height          =   345
            Left            =   45
            TabIndex        =   43
            Top             =   45
            Width           =   7245
            _ExtentX        =   12779
            _ExtentY        =   609
            Caption         =   "Account Distribution"
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
            Icon            =   "frmInvStockIssuance.frx":265F
         End
         Begin VB.Label lblTotalDebit 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4440
            TabIndex        =   53
            Top             =   3000
            Width           =   1215
         End
         Begin VB.Label lblTotalCredit 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5760
            TabIndex        =   52
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL >>"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3600
            TabIndex        =   51
            Top             =   3000
            Width           =   855
         End
         Begin VB.Label lblBalance 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5760
            TabIndex        =   50
            Top             =   3240
            Width           =   1095
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "BALANCE >>"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3600
            TabIndex        =   49
            Top             =   3240
            Width           =   975
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "INVOICE NET AMOUNT"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5280
            TabIndex        =   48
            Top             =   480
            Width           =   1875
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "INVOICE NUMBER"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "INVOICE DATE"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1680
            TabIndex        =   46
            Top             =   480
            Width           =   1515
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "INVOICE GROSS AMOUNT"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3240
            TabIndex        =   45
            Top             =   480
            Width           =   1995
         End
         Begin VB.Label Label45 
            BackStyle       =   0  'Transparent
            Caption         =   "BOOK TYPE"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1320
            TabIndex        =   44
            Top             =   3120
            Width           =   975
         End
      End
   End
   Begin RPVGCC.b8Container picSLine 
      Height          =   855
      Left            =   1320
      TabIndex        =   17
      Top             =   4920
      Visible         =   0   'False
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox txtItemCode 
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtItemDescription 
         Height          =   315
         Left            =   1440
         TabIndex        =   26
         Top             =   360
         Width           =   4575
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7440
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtItemKey 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtItemKey1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtQty1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtItemDescription1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtItemCode1 
         Height          =   285
         Left            =   4080
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtUnit 
         Height          =   315
         Left            =   6120
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtUnit1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4560
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM CODE"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM DESCRIPTION"
         Height          =   255
         Left            =   1440
         TabIndex        =   30
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "QTY"
         Height          =   255
         Left            =   7440
         TabIndex        =   29
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "UNIT"
         Height          =   255
         Left            =   6120
         TabIndex        =   28
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   5625
      Width           =   12255
      _ExtentX        =   21616
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
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   1080
      ScaleHeight     =   4215
      ScaleWidth      =   9855
      TabIndex        =   1
      Top             =   1200
      Width           =   9855
      Begin VB.ComboBox cmbDepartment 
         Height          =   315
         Left            =   4080
         TabIndex        =   14
         Text            =   "cmbSource"
         Top             =   360
         Width           =   2895
      End
      Begin VB.ComboBox cmbLocation 
         Height          =   315
         Left            =   4080
         TabIndex        =   12
         Text            =   "cmbSource"
         Top             =   0
         Width           =   2895
      End
      Begin VB.TextBox txtPostedDate 
         Height          =   315
         Left            =   8280
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtPostedTime 
         Height          =   315
         Left            =   8280
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtRemarks 
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   6
         Top             =   720
         Width           =   5655
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtCtrl 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
      Begin MSComctlLib.ListView lstDetail 
         Height          =   2655
         Left            =   0
         TabIndex        =   16
         Top             =   1200
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   4683
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
         NumItems        =   7
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
            Text            =   "ItemKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ITEM CODE"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ITEM DESCRIPTION"
            Object.Width           =   9085
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "UNIT"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "QTY"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL QTY >>"
         Height          =   255
         Left            =   6960
         TabIndex        =   33
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label lblTotalQty 
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
         Height          =   255
         Left            =   8280
         TabIndex        =   32
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "DEPARTMENT"
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "LOCATION"
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "POSTED DATE"
         Height          =   255
         Left            =   7080
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "POSTED TIME"
         Height          =   255
         Left            =   7080
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "REMARKS"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ISSUANCE DATE"
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ISSUANCE #"
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1215
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
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockIssuance.frx":2BF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockIssuance.frx":38D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockIssuance.frx":45AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockIssuance.frx":5287
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockIssuance.frx":5F61
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockIssuance.frx":6C3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockIssuance.frx":7915
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockIssuance.frx":85EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockIssuance.frx":92C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockIssuance.frx":9BA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockIssuance.frx":A87D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockIssuance.frx":B557
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockIssuance.frx":C231
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockIssuance.frx":CF0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockIssuance.frx":DBE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockIssuance.frx":E8BF
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInvStockIssuance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2
Const is_FINDING = 3

Dim TRANS_DETAIL As Long
Const is_DET_REFRESH = 0
Const is_DET_ADDING = 1
Const is_DET_EDITTING = 2

Dim iRow            As Long
Dim isFocus         As Long

Dim iLocation       As Long
Dim iDepartment     As Long
Dim tmp             As Long

Dim x, sCtrl, iPK, i, iLine, dQty, dAvailableQty, a, b, iBookType

Private Sub BROWSER(Ctrl, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If Ctrl <> "" Then
            s = "SELECT TOP 1 tbl_Inv_StockIssuance.* " & _
                " FROM tbl_Inv_StockIssuance " & _
                " WHERE (CtrlNo = '" & Ctrl & "')" & _
                " ORDER BY CtrlNo"
        Else
            s = "SELECT TOP 1 tbl_Inv_StockIssuance.* " & _
                " FROM tbl_Inv_StockIssuance " & _
                " ORDER BY CtrlNo"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_StockIssuance.* " & _
            " FROM tbl_Inv_StockIssuance " & _
            " ORDER BY CtrlNo"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_StockIssuance.* " & _
            " FROM tbl_Inv_StockIssuance " & _
            " WHERE (CtrlNo < '" & Ctrl & "')" & _
            " ORDER BY CtrlNo DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_StockIssuance.* " & _
            " FROM tbl_Inv_StockIssuance " & _
            " WHERE (CtrlNo > '" & Ctrl & "')" & _
            " ORDER BY CtrlNo "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_StockIssuance.* " & _
            " FROM tbl_Inv_StockIssuance " & _
            " ORDER BY CtrlNo DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtCtrl.Text = rs!CtrlNo
    txtDate.Text = Format(rs!dDate, "mm/dd/yyyy")
    txtRemarks.Text = rs!Remarks
    iLocation = rs!Location
    iDepartment = rs!Department
    cmbLocation.Text = ""
    t = "SELECT LocName " & _
        " FROM tbl_Inv_Location " & _
        " WHERE (PK = " & iLocation & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        cmbLocation.Text = rt!LocName
    End If
    rt.Close
    cmbDepartment.Text = ""
    t = "SELECT DeptName as LocName " & _
        " FROM tbl_GL_Department " & _
        " WHERE (PK = " & iDepartment & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        cmbDepartment.Text = rt!LocName
    End If
    rt.Close
'    txtGLAccount.Text = IIf(IsNull(rs!GLCode), "", rs!GLCode)
    txtPostedDate.Text = ""
    txtPostedTime.Text = ""
    If IsNull(rs!PostedDateTime) = False Then
        txtPostedDate.Text = Format(rs!PostedDateTime, "mm/dd/yyyy")
        txtPostedTime.Text = Format(rs!PostedDateTime, "hh:mm:ss AM/PM")
    End If
    
    lblTotalQty.Caption = "0.00"
    imgPosted.Visible = IIf(rs!Posted = 1, True, False)
    picGLPosted.Visible = IIf(rs!GLPosted = 1, True, False)
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    CLEARDETAIL
    t = "SELECT tbl_Inv_StockIssuance_Detail.ItemKey, " & _
        " tbl_Inv_Items.ItemCode, " & _
        " tbl_Inv_Items.ItemDesc, " & _
        " tbl_Inv_Items.Unit, " & _
        " tbl_Inv_StockIssuance_Detail.Qty " & _
        " FROM tbl_Inv_StockIssuance_Detail LEFT OUTER JOIN " & _
        " tbl_Inv_Items ON tbl_Inv_StockIssuance_Detail.ItemKey = tbl_Inv_Items.PK " & _
        " WHERE (tbl_Inv_StockIssuance_Detail.MasterKey = " & rs!PK & ") " & _
        " ORDER BY tbl_Inv_StockIssuance_Detail.Line"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        With lstDetail.ListItems
            .Clear
            iLine = 0
            dQty = 0
            While Not rt.EOF
                iLine = iLine + 1
                dQty = dQty + CDbl(rt!Qty)
                Set x = .Add()
                x.Text = ""
                x.SubItems(1) = Format(iLine, "0#")
                x.SubItems(2) = rt!ItemKey
                x.SubItems(3) = rt!ItemCode
                x.SubItems(4) = rt!ItemDesc
                x.SubItems(5) = rt!Unit
                x.SubItems(6) = Format(rt!Qty, "#,##0.00")
                rt.MoveNext
            Wend
        End With
    End If
    rt.Close
    lblTotalQty.Caption = Format(dQty, "#,##0.00")
    SaveSetting App.EXEName, "IssuanceCtrlNo", "IssuanceCtrl", rs!CtrlNo
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSLine.Visible = True Then Exit Sub
    If AccessRights("Stock Issuance", "Add") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    CLEARTEXT
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_ADDING
    txtDate.Text = Format(Date, "mm/dd/yyyy")
    'Me.Caption = "STOCK ISSUANCE - NEW"
    txtDate.SetFocus
Else
    If picSLine.Visible = True Then Exit Sub
    If isFocus = 0 Then Exit Sub
    If AccessRights("Stock Issuance", "Add") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
    With lstDetail.ListItems
        If CDbl(.Item(.Count).SubItems(2)) = 0 Then
            .Item(.Count).SubItems(1) = Format(.Count, "0#")
            .Item(.Count).SubItems(2) = "0"
            .Item(.Count).SubItems(3) = " "
            .Item(.Count).SubItems(4) = " "
            .Item(.Count).SubItems(5) = " "
            .Item(.Count).SubItems(6) = " "
            iRow = .Count
            txtItemCode.Text = ""
            txtItemDescription.Text = ""
            txtQty.Text = ""
            picToolbar.Enabled = False
            picMain.Enabled = False
            picSLine.ZOrder 0
            picSLine.Visible = True
            TRANS_DETAIL = is_DET_ADDING
            TOOLBARFUNC 3
            txtItemCode.SetFocus
        Else
            Set x = .Add()
            x.Text = ""
            x.SubItems(1) = Format(.Count, "0#")
            x.SubItems(2) = "0"
            x.SubItems(3) = " "
            x.SubItems(4) = " "
            x.SubItems(5) = " "
            x.SubItems(6) = " "
            iRow = .Count
            lstDetail.ListItems(iRow).EnsureVisible
            lstDetail.ListItems(iRow).Selected = True
            txtItemCode.Text = ""
            txtItemDescription.Text = ""
            txtQty.Text = ""
            picToolbar.Enabled = False
            picMain.Enabled = False
            picSLine.ZOrder 0
            picSLine.Visible = True
            TRANS_DETAIL = is_DET_ADDING
            TOOLBARFUNC 3
            txtItemCode.SetFocus
        End If
    End With
End If
End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSLine.Visible = True Then Exit Sub
    If Statusbar1.Panels(1).Text = "" Then Exit Sub
    If AccessRights("Stock Issuance", "Edit") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_EDITTING
    'Me.Caption = "STOCK ISSUANCE - EDIT"
Else
    If picSLine.Visible = True Then Exit Sub
    If isFocus = 0 Then Exit Sub
    If AccessRights("Stock Issuance", "Edit") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
    With lstDetail.ListItems
        txtItemKey.Text = .Item(iRow).SubItems(2)
        txtItemCode.Text = .Item(iRow).SubItems(3)
        txtItemDescription.Text = .Item(iRow).SubItems(4)
        txtUnit.Text = .Item(iRow).SubItems(5)
        txtQty.Text = .Item(iRow).SubItems(6)
        
        txtItemKey1.Text = .Item(iRow).SubItems(2)
        txtItemCode1.Text = .Item(iRow).SubItems(3)
        txtItemDescription1.Text = .Item(iRow).SubItems(4)
        txtUnit1.Text = .Item(iRow).SubItems(5)
        txtQty1.Text = .Item(iRow).SubItems(6)
        
        picToolbar.Enabled = False
        picMain.Enabled = False
        picSLine.ZOrder 0
        picSLine.Visible = True
        TRANS_DETAIL = is_DET_EDITTING
        TOOLBARFUNC 3
        txtItemCode.SetFocus
    End With
End If
End Sub

Private Sub PRESS_DELETE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSLine.Visible = True Then Exit Sub
    If Statusbar1.Panels(1).Text = "" Then Exit Sub
    If AccessRights("Stock Issuance", "Delete") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
    If MsgBox("ARE YOU SURE IN DELETING THIS TRANSACTION?                       ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    On Error GoTo PG:
    ConnOmega.Execute "DELETE FROM tbl_Inv_StockIssuance WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    CLEARTEXT
    BROWSER GetSetting(App.EXEName, "IssuanceCtrlNo", "IssuanceCtrl", ""), "is_PAGEDOWN"
    If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "IssuanceCtrlNo", "IssuanceCtrl", ""), "is_HOME"
Else
    If picSLine.Visible = True Then Exit Sub
    If isFocus = 0 Then Exit Sub
    If AccessRights("Stock Issuance", "Delete") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
    With lstDetail.ListItems
        If .Count = 1 Then
            .Item(.Count).SubItems(1) = " "
            .Item(.Count).SubItems(2) = "0"
            .Item(.Count).SubItems(3) = " "
            .Item(.Count).SubItems(4) = " "
            .Item(.Count).SubItems(5) = " "
            .Item(.Count).SubItems(6) = " "
            iRow = 1
        ElseIf .Count > 1 Then
            .Remove iRow
            If CDbl(iRow) > CDbl(.Count) Then
                iRow = .Count
            End If
        End If
        lstDetail.ListItems(iRow).EnsureVisible
        lstDetail.ListItems(iRow).Selected = True
    End With
End If

Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub

End Sub

Private Sub PRESS_F5()
If IsDate(txtDate.Text) = False Then MsgBox "Please Supply a Valid Date!                        ", vbCritical, "Error...": txtDate.SetFocus: Exit Sub
If iLocation = 0 Then MsgBox "Please Select Location!                           ", vbCritical, "Error...": cmbLocation.SetFocus: Exit Sub
If iDepartment = 0 Then MsgBox "Please Select Department!                     ", vbCritical, "Error...": cmbDepartment.SetFocus: Exit Sub
'If Trim(txtGLAccount.Text) <> "" Then
    
'End If
txtDate.Text = Format(FormatDateTime(txtDate.Text, vbShortDate), "mm/dd/yyyy")
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    sCtrl = ""
    s = "SELECT TOP 1 tbl_Inv_StockIssuance.* " & _
        " FROM tbl_Inv_StockIssuance " & _
        " WHERE (Year(DDate) = " & Format(txtDate.Text, "yyyy") & ") " & _
        " ORDER BY CtrlNo DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        sCtrl = CDbl(rs!CtrlNo) + 1
    Else
        sCtrl = Format(txtDate.Text, "yyyy") & "0000"
    End If
    rs.Close
    Do
        s = "SELECT tbl_Inv_StockIssuance.* " & _
            " FROM tbl_Inv_StockIssuance " & _
            " WHERE (CtrlNo = '" & sCtrl & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If
        rs.Close
        sCtrl = CDbl(sCtrl) + 1
    Loop
    
    ConnOmega.Execute "INSERT INTO tbl_Inv_StockIssuance " & _
                      " (CtrlNo, DDate, Remarks, Location, Department, LastModified) " & _
                      " VALUES ('" & sCtrl & "', '" & FormatDateTime(txtDate.Text, vbShortDate) & "', " & _
                      " '" & FORMATSQL(Trim(txtRemarks.Text)) & "', " & cmbLocation.ItemData(cmbLocation.ListIndex) & ", " & _
                      " " & cmbDepartment.ItemData(cmbDepartment.ListIndex) & ", '" & CStr(Now) & " - " & gbl_CompleteName & "')"
                      
    iPK = 0
    s = "SELECT tbl_Inv_StockIssuance.* " & _
        " FROM tbl_Inv_StockIssuance " & _
        " WHERE (CtrlNo = '" & sCtrl & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        iPK = rs!PK
    End If
    rs.Close
    
    If CDbl(iPK) > 0 Then
'        If Trim(txtGLAccount.Text) <> "" Then
'            ConnOmega.Execute "UPDATE tbl_Inv_StockIssuance " & _
'                              " SET GLCode = '" & FORMATSQL(Trim(txtGLAccount.Text)) & "' " & _
'                              " WHERE (PK = " & iPK & ")"
'        End If
        iLine = 0
        With lstDetail.ListItems
            For i = 1 To .Count
                If CDbl(.Item(i).SubItems(2)) > 0 Then
                    iLine = iLine + 1
                    ConnOmega.Execute "INSERT INTO tbl_Inv_StockIssuance_Detail " & _
                                      " (STKey, Line, ItemKey, Qty) " & _
                                      " VALUES (" & iPK & ", " & iLine & ", " & _
                                      " " & .Item(i).SubItems(2) & ", " & _
                                      " " & CDbl(.Item(i).SubItems(6)) & ")"
                End If
            Next i
        End With
    End If
End If
If TRANSACTIONTYPE = is_EDITTING Then
    sCtrl = Trim(txtCtrl.Text)
    iPK = Statusbar1.Panels(1).Text
    
    ConnOmega.Execute "UPDATE tbl_Inv_StockIssuance " & _
                      " SET DDate = '" & FormatDateTime(txtDate.Text, vbShortDate) & "', " & _
                      " Remarks = '" & FORMATSQL(Trim(txtRemarks.Text)) & "', " & _
                      " Location = " & cmbLocation.ItemData(cmbLocation.ListIndex) & ", " & _
                      " Department = " & cmbDepartment.ItemData(cmbDepartment.ListIndex) & ", " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & iPK & ")"
    
'    If Trim(txtGLAccount.Text) <> "" Then
'        ConnOmega.Execute "UPDATE tbl_Inv_StockIssuance " & _
'                          " SET GLCode = '" & FORMATSQL(Trim(txtGLAccount.Text)) & "' " & _
'                          " WHERE (PK = " & iPK & ")"
'    End If
        
    ConnOmega.Execute "DELETE FROM tbl_Inv_StockIssuance_Detail WHERE (STKey = " & iPK & ")"
    iLine = 0
    With lstDetail.ListItems
        For i = 1 To .Count
            If CDbl(.Item(i).SubItems(2)) > 0 Then
                iLine = iLine + 1
                ConnOmega.Execute "INSERT INTO tbl_Inv_StockIssuance_Detail " & _
                                  " (STKey, Line, ItemKey, Qty) " & _
                                  " VALUES (" & iPK & ", " & iLine & ", " & _
                                  " " & .Item(i).SubItems(2) & ", " & _
                                  " " & CDbl(.Item(i).SubItems(6)) & ")"
            End If
        Next i
    End With
End If
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANS_DETAIL = is_DET_REFRESH
'Me.Caption = "STOCK ISSUANCE - BROWSE"
txtCtrl.SetFocus
BROWSER sCtrl, "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSLine.Visible = True Then Exit Sub
CLEARTEXT
TOOLBARFUNC 3
TRANSACTIONTYPE = is_FINDING
txtCtrl.Locked = False
'Me.Caption = "STOCK ISSUANCE - FIND"
txtCtrl.SetFocus
End Sub

Private Sub PRESS_F7()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSLine.Visible = True Then Exit Sub
If picAccDistribution.Visible = True Then Exit Sub
'If picPost.Visible = True Then Exit Sub
'If imgPosted.Visible = False Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("Stock Issuance", "Post Inv") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
picAccDistribution.ZOrder 0
lstAccDistribution.ListItems.Clear
Set x = lstAccDistribution.ListItems.Add()
x.Text = ""
x.SubItems(1) = " "
x.SubItems(2) = " "
x.SubItems(3) = " "
x.SubItems(4) = " "
x.SubItems(5) = "0"
a = 0: b = 0
s = "SELECT tbl_Inv_StockIssuance_AD.AccountCode, " & _
    " tbl_GL_Accounts.AccountName, " & _
    " tbl_Inv_StockIssuance_AD.Debit, " & _
    " tbl_Inv_StockIssuance_AD.Credit, " & _
    " tbl_Inv_StockIssuance_AD.Amount " & _
    " FROM tbl_Inv_StockIssuance_AD LEFT OUTER JOIN " & _
    " tbl_GL_Accounts ON tbl_Inv_StockIssuance_AD.AccountCode = tbl_GL_Accounts.AccountCode " & _
    " Where (tbl_Inv_StockIssuance_AD.POKey = " & Statusbar1.Panels(1).Text & ") " & _
    " ORDER BY tbl_Inv_StockIssuance_AD.Line"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    lstAccDistribution.ListItems.Clear
    While Not rs.EOF
        Set x = lstAccDistribution.ListItems.Add()
        x.Text = ""
        x.SubItems(1) = rs!AccountCode
        x.SubItems(2) = rs!AccountName
        x.SubItems(3) = IIf(CDbl(rs!Debit) = 0, " ", Format(rs!Debit, "#,##0.00"))
        x.SubItems(4) = IIf(CDbl(rs!Credit) = 0, " ", Format(rs!Credit, "#,##0.00"))
        x.SubItems(5) = rs!Amount
        a = a + CDbl(rs!Debit)
        b = b + CDbl(rs!Credit)
        rs.MoveNext
    Wend
End If
rs.Close
lblTotalDebit.Caption = Format(a, "#,##0.00")
lblTotalCredit.Caption = Format(b, "#,##0.00")

cmbBookType.Clear
s = "SELECT tbl_Acctg_Book.* " & _
    " FROM tbl_Acctg_Book " & _
    " WHERE (ViewInRR = 1) " & _
    " ORDER BY PK"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbBookType.AddItem rs!Abb
    cmbBookType.ItemData(cmbBookType.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close

s = "SELECT tbl_Acctg_Book.* " & _
    " FROM tbl_Acctg_Book " & _
    " WHERE (PK = " & iBookType & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    cmbBookType.Text = rs!Abb
End If
rs.Close

picToolbar.Enabled = False
picMain.Enabled = False
picAccDistribution.Height = 3615 '3015
picAccDistribution.Width = 7335
picAccDistribution.ZOrder 0
picAccDistribution.Visible = True
lstAccDistribution.SetFocus
End Sub


Private Sub PRESS_F8()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSLine.Visible = True Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("Stock Issuance", "Post") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
BROWSER GetSetting(App.EXEName, "IssuanceCtrlNo", "IssuanceCtrl", ""), "is_LOAD"
If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub

On Error GoTo PG:

'   Checking
With lstDetail.ListItems
    For i = 1 To .Count
        If CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2))) <> 0 Then
            dQty = CDbl(IIf(IsNumeric(.Item(i).SubItems(6)) = False, 0, .Item(i).SubItems(6)))
            If CDbl(dQty) > 0 Then
                s = "SELECT tbl_Inv_Items_Available_Location.* " & _
                    " FROM tbl_Inv_Items_Available_Location " & _
                    " WHERE (ItemKey = " & .Item(i).SubItems(2) & ") " & _
                    " AND (LocationKey = " & iLocation & ")"
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                If rs.RecordCount > 0 Then
                    If CDbl(rs!Quantity) < CDbl(dQty) Then
                        MsgBox "Not Enough Quantity to Issue!                        ", vbCritical, "Error..."
                        lstDetail.ListItems(i).EnsureVisible
                        lstDetail.ListItems(i).Selected = True
                        lstDetail.SetFocus
                        If rs.State = adStateOpen Then rs.Close
                        Exit Sub
                    End If
                Else
                    MsgBox "Not Enough Quantity to Issue!                        ", vbCritical, "Error..."
                    lstDetail.ListItems(i).EnsureVisible
                    lstDetail.ListItems(i).Selected = True
                    lstDetail.SetFocus
                    If rs.State = adStateOpen Then rs.Close
                    Exit Sub
                End If
                rs.Close
            End If
        End If
    Next i
End With


'   Posting

With lstDetail.ListItems
    For i = 1 To .Count
        If CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2))) <> 0 Then
            dQty = CDbl(IIf(IsNumeric(.Item(i).SubItems(6)) = False, 0, .Item(i).SubItems(6)))
            Do
                s = "SELECT TOP 1 PK, QuantityIn - QuantityUsed as AvQty, " & _
                    " QuantityUsed, Cost, PurcDisc, NetCost, NetVAT, QuantityIn " & _
                    " FROM tbl_Inv_Items_Transaction " & _
                    " WHERE (ItemKey = " & .Item(i).SubItems(2) & ") " & _
                    " AND (Location = " & iLocation & ") " & _
                    " AND (InOut = 'I') " & _
                    " AND (Cleared = 0) " & _
                    " ORDER BY PK"
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                If rs.RecordCount > 0 Then
                    If CDbl(rs!AvQty) > 0 Then
                        If CDbl(dQty) > CDbl(rs!AvQty) Then
                            dAvailableQty = CDbl(rs!AvQty)
                        Else
                            dAvailableQty = CDbl(dQty)
                        End If
                        ConnOmega.Execute "INSERT INTO tbl_Inv_Items_Transaction " & _
                                          " (ItemKey, Cleared, InOut, DocType, DocNumber, DocDate, Location, " & _
                                          " ReferenceKey, QuantityOut, QuantityUsed, Cost, PurcDisc, NetCost, " & _
                                          " LogInName, NetVAT) " & _
                                          " VALUES (" & .Item(i).SubItems(2) & ", 1, 'O', 6, '" & Trim(txtCtrl.Text) & "', " & _
                                          " '" & FormatDateTime(txtDate.Text, vbShortDate) & "', " & iLocation & ", " & _
                                          " " & rs!PK & ", " & CDbl(dAvailableQty) & ", " & CDbl(dAvailableQty) & ", " & _
                                          " " & CDbl(rs!Cost) & ", '" & IIf(IsNull(rs!PurcDisc), "", rs!PurcDisc) & "', " & _
                                          " " & CDbl(rs!NetCost) & ", '" & gbl_UserName & "', " & CDbl(rs!NetVAT) & ")"
                        If CDbl(rs!QuantityIn) <= (CDbl(rs!AvQty) + CDbl(dAvailableQty)) Then
                            ConnOmega.Execute "UPDATE tbl_Inv_Items_Transaction " & _
                                              " SET Cleared = 1, QuantityUsed = QuantityUsed + " & CDbl(dAvailableQty) & " " & _
                                              " WHERE (PK = " & rs!PK & ")"
                        Else
                            ConnOmega.Execute "UPDATE tbl_Inv_Items_Transaction " & _
                                              " SET QuantityUsed = QuantityUsed + " & CDbl(dAvailableQty) & " " & _
                                              " WHERE (PK = " & rs!PK & ")"
                        End If
                        dQty = dQty - CDbl(dAvailableQty)
                        If CDbl(dQty) <= 0 Then Exit Do
                    End If
                End If
                rs.Close
            Loop
        End If
    Next i
End With

ConnOmega.Execute "UPDATE tbl_Inv_StockIssuance " & _
                  " SET Posted = 1, " & _
                  " PostedDateTime = '" & Now & "', " & _
                  " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                  " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"

BROWSER GetSetting(App.EXEName, "IssuanceCtrlNo", "IssuanceCtrl", ""), "is_LOAD"

Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub

End Sub

Private Sub PRESS_F9()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSLine.Visible = True Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub

End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSLine.Visible = True Then Exit Sub
    Unload Me
Else
    If picSLine.Visible = True Then
        If TRANS_DETAIL = is_DET_ADDING Then
            With lstDetail.ListItems
                If .Count = 1 Then
                    .Item(.Count).SubItems(1) = " "
                    .Item(.Count).SubItems(2) = "0"
                    .Item(.Count).SubItems(3) = " "
                    .Item(.Count).SubItems(4) = " "
                    .Item(.Count).SubItems(5) = " "
                    .Item(.Count).SubItems(6) = " "
                ElseIf .Count > 1 Then
                    .Remove .Count
                End If
            End With
        End If
        If TRANS_DETAIL = is_DET_EDITTING Then
            With lstDetail.ListItems
                .Item(iRow).SubItems(2) = txtItemKey1.Text
                .Item(iRow).SubItems(3) = txtItemCode1.Text
                .Item(iRow).SubItems(4) = txtItemDescription1.Text
                .Item(iRow).SubItems(5) = txtUnit1.Text
                .Item(iRow).SubItems(6) = txtQty1.Text
            End With
        End If
        picSLine.Visible = False
        picToolbar.Enabled = True
        picMain.Enabled = True
        lstDetail.SetFocus
        Exit Sub
    End If
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    TRANS_DETAIL = is_DET_REFRESH
    txtCtrl.SetFocus
    'Me.Caption = "STOCK ISSUANCE - BROWSE"
    BROWSER GetSetting(App.EXEName, "IssuanceCtrlNo", "IssuanceCtrl", ""), "is_LOAD"
    If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "IssuanceCtrlNo", "IssuanceCtrl", ""), "is_HOME"
End If

End Sub

Private Sub CLEARTEXT()
iLocation = 0
iDepartment = 0
txtCtrl.Text = ""
txtDate.Text = ""
txtRemarks.Text = ""
cmbLocation.Text = ""
cmbLocation.ListIndex = -1
cmbDepartment.Text = ""
cmbDepartment.ListIndex = -1
'txtGLAccount.Text = ""
txtPostedDate.Text = ""
txtPostedTime.Text = ""
lblTotalQty.Caption = "0.00"
imgPosted.Visible = False
picGLPosted.Visible = False
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
CLEARDETAIL
End Sub

Private Sub CLEARDETAIL()
lstDetail.ListItems.Clear
Set x = lstDetail.ListItems.Add()
x.Text = ""
x.SubItems(1) = " "
x.SubItems(2) = "0"
x.SubItems(3) = " "
x.SubItems(4) = " "
x.SubItems(5) = " "
x.SubItems(6) = " "
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtCtrl.Locked = True
'txtGLAccount.Locked = True
txtPostedDate.Locked = True
txtPostedTime.Locked = True
txtDate.Locked = bln
txtRemarks.Locked = bln
cmbLocation.Locked = bln
cmbDepartment.Locked = bln
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
            .Buttons(25).Image = 14
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
            .Buttons(21).Caption = "GL Acc"
            .Buttons(23).Caption = "Refresh"
            .Buttons(25).Caption = "Close"
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
            .Buttons(25).Enabled = True
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
            .Buttons(21).ToolTipText = "ACCOUNT DISTRIBUTION (F7)"
            .Buttons(23).ToolTipText = "REFRESH (F11)"
            .Buttons(25).ToolTipText = "CLOSE (Esc)"
        Case 2      '=== ADD/EDIT ====
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 15
            .Buttons(9).Image = 16
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(21).Image = 12
            .Buttons(23).Image = 13
            .Buttons(25).Image = 14
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
            .Buttons(21).Caption = "GL Acc"
            .Buttons(23).Caption = "Refresh"
            .Buttons(25).Caption = "Close"
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
            .Buttons(25).Enabled = False
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
            .Buttons(25).ToolTipText = ""
        Case 3      '=== FIND ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 4
            .Buttons(9).Image = 16
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(21).Image = 12
            .Buttons(23).Image = 13
            .Buttons(25).Image = 14
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
            .Buttons(21).Caption = "GL Acc"
            .Buttons(23).Caption = "Refresh"
            .Buttons(25).Caption = "Close"
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
            .Buttons(25).Enabled = False
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
            .Buttons(25).ToolTipText = ""
        Case 4      '=== EMPTY DETAIL ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 15
            .Buttons(9).Image = 16
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(21).Image = 12
            .Buttons(23).Image = 13
            .Buttons(25).Image = 14
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
            .Buttons(21).Caption = "GL Acc"
            .Buttons(23).Caption = "Refresh"
            .Buttons(25).Caption = "Close"
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
            .Buttons(25).Enabled = False
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
            .Buttons(25).ToolTipText = ""
        Case 5      '=== NOT EMPTY DETAIL ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 15
            .Buttons(9).Image = 16
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(21).Image = 12
            .Buttons(23).Image = 13
            .Buttons(25).Image = 14
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
            .Buttons(21).Caption = "GL Acc"
            .Buttons(23).Caption = "Refresh"
            .Buttons(25).Caption = "Close"
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
            .Buttons(25).Enabled = False
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
            .Buttons(25).ToolTipText = ""
    End Select
End With
End Sub


Private Sub cmbDepartment_Click()
If cmbDepartment.ListIndex = -1 Then Exit Sub
iDepartment = cmbDepartment.ItemData(cmbDepartment.ListIndex)
End Sub

Private Sub cmbLocation_Click()
If cmbLocation.ListIndex = -1 Then Exit Sub
iLocation = cmbLocation.ItemData(cmbLocation.ListIndex)
End Sub

Private Sub cmdPost_Click()
If imgPosted.Visible = False Then MsgBox "Please Posted Stock Issuance!                          ", vbCritical, "Error...": Exit Sub
If picGLPosted.Visible = True Then MsgBox "Already Posted in General Ledger!                         ", vbCritical, "Error...": Exit Sub

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert:   PRESS_INSERT
    Case vbKeyF2:       PRESS_F2
    Case vbKeyDelete:   PRESS_DELETE
    Case vbKeyF5:       PRESS_F5
    Case vbKeyF6:       PRESS_F6
    Case vbKeyF7:       'PRESS_F7
    Case vbKeyF8:       PRESS_F8
    Case vbKeyF9:       PRESS_F9
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "IssuanceCtrlNo", "IssuanceCtrl", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "IssuanceCtrlNo", "IssuanceCtrl", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "IssuanceCtrlNo", "IssuanceCtrl", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "IssuanceCtrlNo", "IssuanceCtrl", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
cmbLocation.Clear
s = "SELECT tbl_Inv_Location.* " & _
    " FROM tbl_Inv_Location " & _
    " ORDER BY LocName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbLocation.AddItem rs!LocName
    cmbLocation.ItemData(cmbLocation.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
cmbDepartment.Clear
s = "SELECT tbl_GL_Department.* " & _
    " FROM tbl_GL_Department " & _
    " WHERE (Issuance = 1) " & _
    " ORDER BY DeptName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbDepartment.AddItem rs!DeptName
    cmbDepartment.ItemData(cmbDepartment.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
isFocus = 0
iRow = 0
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANS_DETAIL = is_DET_REFRESH
'Me.Caption = "STOCK ISSUANCE - BROWSE"
BROWSER GetSetting(App.EXEName, "IssuanceCtrlNo", "IssuanceCtrl", ""), "is_LOAD"
If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "IssuanceCtrlNo", "IssuanceCtrl", ""), "is_HOME"

tmp = SetWindowLong(txtRemarks.hwnd, GWL_STYLE, GetWindowLong(txtRemarks.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picSLine.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstDetail_GotFocus()
iRow = lstDetail.SelectedItem.Index
TRANS_DETAIL = is_DET_REFRESH
isFocus = 1
If TRANSACTIONTYPE = is_REFRESH Then
    If Statusbar1.Panels(1).Text = "" Then Exit Sub
    TRANSACTIONTYPE = is_EDITTING
    'Me.Caption = "STOCK ISSUANCE - EDIT"
    BROWSER GetSetting(App.EXEName, "IssuanceCtrlNo", "IssuanceCtrl", ""), "is_LOAD"
    If imgPosted.Visible = True Then TOOLBARFUNC 3: Exit Sub
End If
With lstDetail.ListItems
    If .Count = 1 Then
        If CDbl(.Item(iRow).SubItems(2)) > 0 Then
            TOOLBARFUNC 5
        Else
            TOOLBARFUNC 4
        End If
    ElseIf .Count > 1 Then
        TOOLBARFUNC 5
    End If
End With
End Sub

Private Sub lstDetail_ItemClick(ByVal Item As MSComctlLib.ListItem)
iRow = lstDetail.SelectedItem.Index
End Sub

Private Sub lstDetail_LostFocus()
isFocus = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "IssuanceCtrlNo", "IssuanceCtrl", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "IssuanceCtrlNo", "IssuanceCtrl", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "IssuanceCtrlNo", "IssuanceCtrl", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "IssuanceCtrlNo", "IssuanceCtrl", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":   PRESS_F9
    Case "Post":    PRESS_F8
    Case "Accnt":   'PRESS_F7
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtItemCode_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(3) = txtItemCode.Text
    End With
End If
End Sub

Private Sub txtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If Trim(txtItemCode.Text) = "" Then MsgBox "Please Supply ItemCode!                   ", vbCritical, "Error...": txtItemCode.SetFocus: Exit Sub
    t = "SELECT PK, ItemCode, ItemDesc, Unit " & _
        " FROM tbl_Inv_Items " & _
        " WHERE (ItemCode = '" & Trim(txtItemCode.Text) & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtItemKey.Text = rt!PK
        txtItemCode.Text = rt!ItemCode
        txtItemDescription.Text = rt!ItemDesc
        txtUnit.Text = rt!Unit
    Else
        MsgBox "'" & txtItemCode.Text & "' Not Found!                  ", vbCritical, "Error..."
        txtItemCode.SetFocus
        rt.Close
        Exit Sub
    End If
    rt.Close
    txtQty.SetFocus
End If
End Sub

Private Sub txtItemCode_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtItemDescription_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(4) = txtItemDescription.Text
    End With
End If
End Sub

Private Sub txtItemKey_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(2) = RETURNTEXTVALUE(txtItemKey)
    End With
End If
End Sub

Private Sub txtQty_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(6) = Format(RETURNTEXTVALUE(txtQty), "#,##0.00")
    End With
End If
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    picSLine.Visible = False
    picToolbar.Enabled = True
    picMain.Enabled = True
    lstDetail.SetFocus
End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtCtrl_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If TRANSACTIONTYPE <> is_FINDING Then Exit Sub
    sCtrl = Trim(txtCtrl.Text)
    s = "SELECT tbl_Inv_StockIssuance.* " & _
        " FROM tbl_Inv_StockIssuance " & _
        " WHERE (CtrlNo = '" & sCtrl & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount = 0 Then
        MsgBox "'" & Trim(txtCtrl.Text) & "' not Found!                     ", vbCritical, "Error..."
        rs.Close
        Exit Sub
    End If
    rs.Close
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    TRANS_DETAIL = is_DET_REFRESH
    'Me.Caption = "STOCK ISSUANCE - BROWSE"
    BROWSER sCtrl, "is_LOAD"
End If
End Sub

Private Sub txtUnit_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(5) = txtUnit.Text
    End With
End If
End Sub


