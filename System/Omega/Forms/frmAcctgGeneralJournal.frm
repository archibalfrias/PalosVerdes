VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAcctgGeneralJournal 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAcctgGeneralJournal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   14925
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15600
      TabIndex        =   52
      Top             =   0
      Width           =   15600
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   53
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
         MouseIcon       =   "frmAcctgGeneralJournal.frx":08CA
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   13380
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   54
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmAcctgGeneralJournal.frx":0BE4
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
   Begin RPVGCC.b8Container picSeachSL 
      Height          =   3015
      Left            =   3240
      TabIndex        =   34
      Top             =   1920
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5318
      BackColor       =   15396057
      Begin VB.ListBox lstSL 
         Height          =   1425
         Left            =   120
         TabIndex        =   38
         Top             =   840
         Width           =   4695
      End
      Begin VB.TextBox txtSearchSL 
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   4695
      End
      Begin VB.CommandButton cmdCancelSL 
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
         Picture         =   "frmAcctgGeneralJournal.frx":12F7
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2400
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKSL 
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
         Picture         =   "frmAcctgGeneralJournal.frx":1A53
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2400
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   39
         Top             =   45
         Width           =   4845
         _ExtentX        =   8546
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
         Icon            =   "frmAcctgGeneralJournal.frx":20C5
         ShadowVisible   =   0   'False
      End
   End
   Begin RPVGCC.b8Container picSeachAccount 
      Height          =   3015
      Left            =   1800
      TabIndex        =   25
      Top             =   1920
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5318
      BackColor       =   15396057
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
         Picture         =   "frmAcctgGeneralJournal.frx":265F
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2400
         Width           =   1560
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
         Picture         =   "frmAcctgGeneralJournal.frx":2CD1
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2400
         Width           =   1560
      End
      Begin VB.TextBox txtSearchAccount 
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   4695
      End
      Begin VB.ListBox lstAccount 
         Height          =   1425
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   4695
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   45
         TabIndex        =   30
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
         Icon            =   "frmAcctgGeneralJournal.frx":342D
         ShadowVisible   =   0   'False
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   6825
      Width           =   14925
      _ExtentX        =   26326
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
      Height          =   5415
      Left            =   120
      ScaleHeight     =   5415
      ScaleWidth      =   14655
      TabIndex        =   1
      Top             =   1200
      Width           =   14655
      Begin VB.TextBox txtRemarks 
         Height          =   315
         Left            =   4080
         MaxLength       =   100
         TabIndex        =   40
         Top             =   360
         Width           =   6735
      End
      Begin MSComctlLib.ListView lstDetail 
         Height          =   4215
         Left            =   0
         TabIndex        =   6
         Top             =   840
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   7435
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
         NumItems        =   12
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
            Text            =   "JVType"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Type"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Account Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Account Name"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "SLKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "SL Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "SL Name"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Description"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "Debit"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "Credit"
            Object.Width           =   2293
         EndProperty
      End
      Begin VB.TextBox txtJVNumber 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   3
         Top             =   0
         Width           =   1695
      End
      Begin VB.TextBox txtJVDate 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   41
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblBalance 
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
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   8640
         TabIndex        =   10
         Top             =   5160
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL >>"
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
         Left            =   10800
         TabIndex        =   9
         Top             =   5160
         Width           =   975
      End
      Begin VB.Label lblTotalCredit 
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
         Left            =   13080
         TabIndex        =   8
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Label lblTotalDebit 
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
         Left            =   11640
         TabIndex        =   7
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "JV #"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   30
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "JV Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   390
         Width           =   1095
      End
   End
   Begin VB.PictureBox picSLine 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   360
      ScaleHeight     =   855
      ScaleWidth      =   14295
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   14295
      Begin RPVGCC.b8Container b8Container1 
         Height          =   855
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   1508
         BackColor       =   8438015
         Begin VB.TextBox txtCredit1 
            Height          =   315
            Left            =   9000
            MaxLength       =   100
            TabIndex        =   51
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtDebit1 
            Height          =   315
            Left            =   8760
            MaxLength       =   100
            TabIndex        =   50
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtDescription1 
            Height          =   315
            Left            =   8520
            MaxLength       =   100
            TabIndex        =   49
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtSLName1 
            Height          =   315
            Left            =   8280
            MaxLength       =   100
            TabIndex        =   48
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtSLCode1 
            Height          =   315
            Left            =   8040
            MaxLength       =   100
            TabIndex        =   47
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtSLKey1 
            Height          =   315
            Left            =   7800
            MaxLength       =   100
            TabIndex        =   46
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtAccDesc1 
            Height          =   315
            Left            =   7560
            MaxLength       =   100
            TabIndex        =   45
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtAccCode1 
            Height          =   315
            Left            =   7320
            MaxLength       =   100
            TabIndex        =   44
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtType 
            Height          =   315
            Left            =   7080
            MaxLength       =   100
            TabIndex        =   43
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtJVType 
            Height          =   315
            Left            =   6840
            MaxLength       =   100
            TabIndex        =   42
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtSLKey 
            Height          =   315
            Left            =   3840
            MaxLength       =   100
            TabIndex        =   33
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtSLName 
            Height          =   315
            Left            =   4080
            MaxLength       =   100
            TabIndex        =   32
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtAccDesc 
            Height          =   315
            Left            =   2640
            MaxLength       =   100
            TabIndex        =   31
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtDescription 
            Height          =   315
            Left            =   4320
            MaxLength       =   100
            TabIndex        =   23
            Top             =   360
            Width           =   6735
         End
         Begin VB.TextBox txtCredit 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   12600
            MaxLength       =   100
            TabIndex        =   20
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtDebit 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   11160
            MaxLength       =   100
            TabIndex        =   19
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtSLCode 
            Height          =   315
            Left            =   2880
            MaxLength       =   100
            TabIndex        =   18
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtAccCode 
            Height          =   315
            Left            =   1440
            MaxLength       =   100
            TabIndex        =   16
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cmbType 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4320
            TabIndex        =   24
            Top             =   120
            Width           =   4935
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Credit"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   12600
            TabIndex        =   22
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Debit"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   11160
            TabIndex        =   21
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "SL Code"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2880
            TabIndex        =   17
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Account Code"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1440
            TabIndex        =   14
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   1215
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   13800
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
            Picture         =   "frmAcctgGeneralJournal.frx":39C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgGeneralJournal.frx":46A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgGeneralJournal.frx":537B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgGeneralJournal.frx":6055
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgGeneralJournal.frx":6D2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgGeneralJournal.frx":7A09
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgGeneralJournal.frx":86E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgGeneralJournal.frx":93BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgGeneralJournal.frx":A097
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgGeneralJournal.frx":A971
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgGeneralJournal.frx":B64B
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgGeneralJournal.frx":C325
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgGeneralJournal.frx":CFFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgGeneralJournal.frx":DCD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgGeneralJournal.frx":E9B3
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAcctgGeneralJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim TRANS_DETAIL As Long
Const is_DET_REFRESH = 0
Const is_DET_ADDING = 1
Const is_DET_EDITTING = 2

Dim iRow        As Long
Dim iFocus      As Long
Dim iFocusAcc   As Long
Dim iFocusSL    As Long
Dim tmp         As Long

Dim JVDate      As Date

Dim i, l, j, x, iPK, Arr, sJVCtrl, dJVSeries, iType, dDebit, dCredit

Private Sub BROWSER(sCtrl, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If sCtrl <> "" Then
            s = "SELECT TOP 1 tbl_Acctg_GeneralJournal.* " & _
                " FROM tbl_Acctg_GeneralJournal " & _
                " WHERE (CtrlNo = '" & sCtrl & "') " & _
                " ORDER BY CtrlNo"
        Else
            s = "SELECT TOP 1 tbl_Acctg_GeneralJournal.* " & _
                " FROM tbl_Acctg_GeneralJournal " & _
                " ORDER BY CtrlNo"
        End If
    Case "is_HOME"
        If picSLine.Visible = True Then Exit Sub
        If picSeachAccount.Visible = True Then Exit Sub
        If picSeachSL.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Acctg_GeneralJournal.* " & _
            " FROM tbl_Acctg_GeneralJournal " & _
            " ORDER BY CtrlNo"
    Case "is_PAGEUP"
        If picSLine.Visible = True Then Exit Sub
        If picSeachAccount.Visible = True Then Exit Sub
        If picSeachSL.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Acctg_GeneralJournal.* " & _
            " FROM tbl_Acctg_GeneralJournal " & _
            " WHERE (CtrlNo < '" & sCtrl & "') " & _
            " ORDER BY CtrlNo DESC"
    Case "is_PAGEDOWN"
        If picSLine.Visible = True Then Exit Sub
        If picSeachAccount.Visible = True Then Exit Sub
        If picSeachSL.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Acctg_GeneralJournal.* " & _
            " FROM tbl_Acctg_GeneralJournal " & _
            " WHERE (CtrlNo > '" & sCtrl & "') " & _
            " ORDER BY CtrlNo"
    Case "is_END"
        If picSLine.Visible = True Then Exit Sub
        If picSeachAccount.Visible = True Then Exit Sub
        If picSeachSL.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Acctg_GeneralJournal.* " & _
            " FROM tbl_Acctg_GeneralJournal " & _
            " ORDER BY CtrlNo DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtJVNumber.Text = rs!CtrlNo
    txtJVDate.Text = Format(rs!JVDate, "mm/dd/yyyy")
    txtRemarks.Text = rs!Remarks
    imgPosted.Visible = IIf(rs!Posted = 1, True, False)
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    
    lblTotalDebit.Caption = "0.00"
    lblTotalCredit.Caption = "0.00"
    
    dDebit = 0: dCredit = 0
    t = "SELECT tbl_Acctg_GeneralJournal_Detail.JVType, " & _
        " tbl_Acctg_GeneralJournal_Detail.GLCode, " & _
        " tbl_GL_Accounts.AccountName, " & _
        " tbl_Acctg_GeneralJournal_Detail.SLKey, " & _
        " tbl_Acctg_GeneralJournal_Detail.SLCode, " & _
        " tbl_Acctg_GeneralJournal_Detail.SLName, " & _
        " tbl_Acctg_GeneralJournal_Detail.Description, " & _
        " tbl_Acctg_GeneralJournal_Detail.Debit, " & _
        " tbl_Acctg_GeneralJournal_Detail.Credit " & _
        " FROM tbl_Acctg_GeneralJournal_Detail LEFT OUTER JOIN " & _
        " tbl_GL_Accounts ON tbl_Acctg_GeneralJournal_Detail.GLCode = tbl_GL_Accounts.AccountCode " & _
        " Where (tbl_Acctg_GeneralJournal_Detail.JVKey = " & rs!PK & ") " & _
        " ORDER BY tbl_Acctg_GeneralJournal_Detail.Line"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstDetail.ListItems.Clear
        l = 0
        While Not rt.EOF
            dDebit = dDebit + CDbl(rt!Debit)
            dCredit = dCredit + CDbl(rt!Credit)
            l = l + 1
            Set x = lstDetail.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = Format(l, "0#")
            x.SubItems(2) = rt!JVType
            x.SubItems(3) = IIf(rt!JVType = 0, "GL", IIf(rt!JVType = 1, "Supplier", IIf(rt!JVType = 2, "Member", IIf(rt!JVType = 3, "Employee", ""))))
            x.SubItems(4) = rt!GLCode
            x.SubItems(5) = rt!AccountName
            x.SubItems(6) = rt!SLKey
            x.SubItems(7) = rt!SLCode
            x.SubItems(8) = rt!SLName
            x.SubItems(9) = rt!Description
            x.SubItems(10) = IIf(rt!Debit = 0, " ", Format(rt!Debit, "#,##0.00"))
            x.SubItems(11) = IIf(rt!Credit = 0, " ", Format(rt!Credit, "#,##0.00"))
            rt.MoveNext
        Wend
    Else
        lstDetail.ListItems.Clear
        Set x = lstDetail.ListItems.Add()
        x.Text = ""
        x.SubItems(1) = " "
        x.SubItems(2) = " "
        x.SubItems(3) = " "
        x.SubItems(4) = " "
        x.SubItems(5) = " "
        x.SubItems(6) = "0"
        x.SubItems(7) = " "
        x.SubItems(8) = " "
        x.SubItems(9) = " "
        x.SubItems(10) = " "
        x.SubItems(11) = " "
    End If
    rt.Close
    lblTotalDebit.Caption = Format(dDebit, "#,##0.00")
    lblTotalCredit.Caption = Format(dCredit, "#,##0.00")
    
    SaveSetting App.EXEName, "GeneralJournal", "GenJour", rs!CtrlNo
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If picSLine.Visible = True Then Exit Sub
If picSeachAccount.Visible = True Then Exit Sub
If picSeachSL.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    CLEARTEXT
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_ADDING
Else
    If iFocus = 0 Then Exit Sub
    If TRANS_DETAIL <> is_DET_REFRESH Then Exit Sub
    With lstDetail.ListItems
        If Trim(.Item(iRow).SubItems(2)) = "" Then
            .Item(iRow).SubItems(1) = Format(iRow, "0#")
            .Item(iRow).SubItems(2) = " "
            .Item(iRow).SubItems(3) = " "
            .Item(iRow).SubItems(4) = " "
            .Item(iRow).SubItems(5) = " "
            .Item(iRow).SubItems(6) = "0"
            .Item(iRow).SubItems(7) = " "
            .Item(iRow).SubItems(8) = " "
            .Item(iRow).SubItems(9) = " "
            .Item(iRow).SubItems(10) = " "
            .Item(iRow).SubItems(11) = " "
        Else
            Set x = lstDetail.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = Format(.Count, "0#")
            x.SubItems(2) = " "
            x.SubItems(3) = " "
            x.SubItems(4) = " "
            x.SubItems(5) = " "
            x.SubItems(6) = "0"
            x.SubItems(7) = " "
            x.SubItems(8) = " "
            x.SubItems(9) = " "
            x.SubItems(10) = " "
            x.SubItems(11) = " "
            iRow = .Count
        End If
        lstDetail.ListItems(iRow).EnsureVisible
        lstDetail.ListItems(iRow).Selected = True
        picMain.Enabled = False
        picToolbar.Enabled = False
        picSLine.ZOrder 0
        cmbType.ListIndex = -1
        txtAccCode.Text = ""
        txtAccDesc.Text = ""
        txtSLKey.Text = "0"
        txtSLCode.Text = ""
        txtSLName.Text = ""
        txtDescription.Text = ""
        txtDebit.Text = ""
        txtCredit.Text = ""
        TRANS_DETAIL = is_DET_ADDING
        picSLine.Visible = True
        cmbType.SetFocus
    End With
End If
End Sub

Private Sub PRESS_F2()
If picSLine.Visible = True Then Exit Sub
If picSeachAccount.Visible = True Then Exit Sub
If picSeachSL.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    If Statusbar1.Panels(1).Text = "" Then Exit Sub
    If imgPosted.Visible = True Then MsgBox "Already Posted!                         ", vbCritical, "Error...": Exit Sub
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_EDITTING
Else
    If imgPosted.Visible = True Then Exit Sub
    If iFocus = 0 Then Exit Sub
    If TRANS_DETAIL <> is_DET_REFRESH Then Exit Sub
    With lstDetail.ListItems
        iType = .Item(iRow).SubItems(2)
        cmbType.ListIndex = .Item(iRow).SubItems(2)
        txtAccCode.Text = .Item(iRow).SubItems(4)
        txtAccDesc.Text = .Item(iRow).SubItems(5)
        txtSLKey.Text = .Item(iRow).SubItems(6)
        txtSLCode.Text = .Item(iRow).SubItems(7)
        txtSLName.Text = .Item(iRow).SubItems(8)
        txtDescription.Text = .Item(iRow).SubItems(9)
        txtDebit.Text = .Item(iRow).SubItems(10)
        txtCredit.Text = .Item(iRow).SubItems(11)
        
        txtJVType.Text = .Item(iRow).SubItems(2)
        txtType.Text = .Item(iRow).SubItems(3)
        txtAccCode1.Text = .Item(iRow).SubItems(4)
        txtAccDesc1.Text = .Item(iRow).SubItems(5)
        txtSLKey1.Text = .Item(iRow).SubItems(6)
        txtSLCode1.Text = .Item(iRow).SubItems(7)
        txtSLName1.Text = .Item(iRow).SubItems(8)
        txtDescription1.Text = .Item(iRow).SubItems(9)
        txtDebit1.Text = .Item(iRow).SubItems(10)
        txtCredit1.Text = .Item(iRow).SubItems(11)
    End With
    lstDetail.ListItems(iRow).EnsureVisible
    lstDetail.ListItems(iRow).Selected = True
    picMain.Enabled = False
    picToolbar.Enabled = False
    picSLine.ZOrder 0
    TRANS_DETAIL = is_DET_EDITTING
    picSLine.Visible = True
    cmbType.SetFocus
End If
End Sub

Private Sub PRESS_DELETE()
If picSLine.Visible = True Then Exit Sub
If picSeachAccount.Visible = True Then Exit Sub
If picSeachSL.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    If Statusbar1.Panels(1).Text = "" Then Exit Sub
    If imgPosted.Visible = True Then MsgBox "Already Posted!                         ", vbCritical, "Error...": Exit Sub
    If MsgBox("ARE YOU SURE IN DELETING THIS TRANSACTION?                   ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    On Error GoTo PG:
    ConnOmega.Execute "DELETE FROM tbl_Acctg_GeneralJournal WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    CLEARTEXT
    BROWSER GetSetting(App.EXEName, "GeneralJournal", "GenJour", ""), "is_PAGEDOWN"
    If Trim(txtJVNumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "GeneralJournal", "GenJour", ""), "is_HOME"
Else
    If imgPosted.Visible = True Then Exit Sub
    If iFocus = 0 Then Exit Sub
    If TRANS_DETAIL <> is_DET_REFRESH Then Exit Sub
    With lstDetail.ListItems
        If .Count = 1 Then
            .Item(iRow).SubItems(1) = " "
            .Item(iRow).SubItems(2) = " "
            .Item(iRow).SubItems(3) = " "
            .Item(iRow).SubItems(4) = " "
            .Item(iRow).SubItems(5) = " "
            .Item(iRow).SubItems(6) = "0"
            .Item(iRow).SubItems(7) = " "
            .Item(iRow).SubItems(8) = " "
            .Item(iRow).SubItems(9) = " "
            .Item(iRow).SubItems(10) = " "
            .Item(iRow).SubItems(11) = " "
        Else
            .Remove iRow
            If CDbl(.Count) > CDbl(iRow) Then
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
If picSLine.Visible = True Then Exit Sub
If picSeachAccount.Visible = True Then Exit Sub
If picSeachSL.Visible = True Then Exit Sub
If IsDate(txtJVDate.Text) = False Then MsgBox "Please Supply a Valid Date!                        ", vbCritical, "Error...": txtJVDate.SetFocus: Exit Sub
JVDate = FormatDateTime(txtJVDate, vbShortDate)
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    sJVCtrl = "": dJVSeries = 0
    s = "SELECT TOP 1 JVSeries " & _
        " FROM tbl_Acctg_GeneralJournal " & _
        " WHERE (JVYear = " & Year(JVDate) & ") " & _
        " AND (JVMonth = " & Month(JVDate) & ") " & _
        " ORDER BY JVSeries DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount = 0 Then
        dJVSeries = 1
        sJVCtrl = Format(JVDate, "yyyy") & "-" & Format(JVDate, "mm") & "-" & "0001"
    Else
        dJVSeries = CDbl(rs!JVSeries) + 1
        sJVCtrl = Format(JVDate, "yyyy") & "-" & Format(JVDate, "mm") & "-" & Format(dJVSeries, "000#")
    End If
    rs.Close
    Do
        s = "SELECT tbl_Acctg_GeneralJournal.* " & _
            " FROM tbl_Acctg_GeneralJournal " & _
            " WHERE (CtrlNo = '" & sJVCtrl & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If
        rs.Close
        dJVSeries = CDbl(dJVSeries) + 1
        sJVCtrl = Format(JVDate, "yyyy") & "-" & Format(JVDate, "mm") & "-" & Format(dJVSeries, "000#")
    Loop
    
    ConnOmega.Execute "INSERT INTO tbl_Acctg_GeneralJournal " & _
                      " (CtrlNo, JVDate, JVYear, JVMonth, JVSeries, Remarks, LastModified) " & _
                      " VALUES ('" & sJVCtrl & "', '" & FormatDateTime(JVDate, vbShortDate) & "', " & _
                      " " & Year(JVDate) & ", " & Month(JVDate) & ", '" & Format(dJVSeries, "000#") & "', " & _
                      " '" & FORMATSQL(Trim(txtRemarks.Text)) & "', '" & CStr(Now) & " - " & gbl_CompleteName & "')"
    
    iPK = 0
    s = "SELECT PK " & _
        " FROM tbl_Acctg_GeneralJournal " & _
        " WHERE (CtrlNo = '" & sJVCtrl & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        iPK = rs!PK
    End If
    rs.Close
    
    If CDbl(iPK) <> 0 Then
        With lstDetail.ListItems
            l = 0
            For i = 1 To .Count
                If Trim(.Item(i).SubItems(4)) <> "" Then
                    l = l + 1
                    ConnOmega.Execute "INSERT INTO tbl_Acctg_GeneralJournal_Detail " & _
                                      " (JVKey, Line, JVType, GLCode, SLKey, SLCode, SLName, Description, Debit, Credit) " & _
                                      " VALUES (" & iPK & ", " & l & ", " & .Item(i).SubItems(2) & ", '" & .Item(i).SubItems(4) & "', " & _
                                      " " & .Item(i).SubItems(6) & ", '" & .Item(i).SubItems(7) & "', '" & FORMATSQL(.Item(i).SubItems(8)) & "', " & _
                                      " '" & FORMATSQL(.Item(i).SubItems(9)) & "', " & CDbl(IIf(IsNumeric(.Item(i).SubItems(10)) = False, 0, .Item(i).SubItems(10))) & ", " & _
                                      " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(11)) = False, 0, .Item(i).SubItems(11))) & ")"
                End If
            Next i
        End With
    End If
    
End If
If TRANSACTIONTYPE = is_EDITTING Then
    
    sJVCtrl = txtJVNumber.Text
    iPK = Statusbar1.Panels(1).Text
    
    ConnOmega.Execute "UPDATE tbl_Acctg_GeneralJournal " & _
                      " SET JVDate = " & FormatDateTime(JVDate, vbShortDate) & ", " & _
                      " Remarks = '" & FORMATSQL(Trim(txtRemarks.Text)) & "', " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & iPK & ")"
    
    If CDbl(iPK) <> 0 Then
        ConnOmega.Execute "DELETE FROM tbl_Acctg_GeneralJournal_Detail WHERE (JVKey = " & iPK & ")"
        With lstDetail.ListItems
            l = 0
            For i = 1 To .Count
                If Trim(.Item(i).SubItems(4)) <> "" Then
                    l = l + 1
                    ConnOmega.Execute "INSERT INTO tbl_Acctg_GeneralJournal_Detail " & _
                                      " (JVKey, Line, JVType, GLCode, SLKey, SLCode, SLName, Description, Debit, Credit) " & _
                                      " VALUES (" & iPK & ", " & l & ", " & .Item(i).SubItems(2) & ", '" & .Item(i).SubItems(4) & "', " & _
                                      " " & .Item(i).SubItems(6) & ", '" & .Item(i).SubItems(7) & "', '" & FORMATSQL(.Item(i).SubItems(8)) & "', " & _
                                      " '" & FORMATSQL(.Item(i).SubItems(9)) & "', " & CDbl(IIf(IsNumeric(.Item(i).SubItems(10)) = False, 0, .Item(i).SubItems(10))) & ", " & _
                                      " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(11)) = False, 0, .Item(i).SubItems(11))) & ")"
                End If
            Next i
        End With
    End If
End If
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER sJVCtrl, "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
'If picSLine.Visible = True Then Exit Sub
If picSeachAccount.Visible = True Then Exit Sub
If picSeachSL.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    CLEARTEXT
    TOOLBARFUNC 3
Else
    If picSLine.Visible = True Then
        If iFocusAcc = 1 Then
            picSLine.Enabled = False
            picSeachAccount.ZOrder 0
            txtSearchAccount.Text = ""
            picSeachAccount.Visible = True
            txtSearchAccount.SetFocus
        End If
        If iFocusSL = 1 Then
            If cmbType.ListIndex = 0 Then Exit Sub
            picSLine.Enabled = False
            picSeachSL.ZOrder 0
            txtSearchSL.Text = ""
            picSeachSL.Visible = True
            txtSearchSL.SetFocus
        End If
    End If
End If
End Sub

Private Sub PRESS_F8()
If picSLine.Visible = True Then Exit Sub
If picSeachAccount.Visible = True Then Exit Sub
If picSeachSL.Visible = True Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If RETURNLABELVALUE(lblBalance) <> 0 Then MsgBox "Please Check Journal Details!                         ", vbCritical, "Error...": Exit Sub
If MsgBox("ARE YOU SURE IN POSTING THIS TRANSACTION?                       ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
With lstDetail.ListItems
    '== Checking
    For i = 1 To .Count
        s = "SELECT tbl_GL_Accounts.* " & _
            " FROM tbl_GL_Accounts " & _
            " WHERE (AccountCode = '" & .Item(i).SubItems(4) & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            MsgBox "Invalid Account Code!                       ", vbCritical, "Error..."
            rs.Close
            Exit Sub
        End If
        rs.Close
        Select Case CLng(.Item(i).SubItems(2))
            Case 1
                t = "SELECT tbl_Inv_Supplier.* " & _
                    " FROM tbl_Inv_Supplier " & _
                    " WHERE (PK = " & .Item(i).SubItems(6) & ")"
                rt.Open t, ConnOmega
                If rt.RecordCount = 0 Then
                    MsgBox "Invalid Supplier!                   ", vbCritical, "Error...'"
                    rt.Close
                    Exit Sub
                End If
                rt.Close
            Case 2
                t = "SELECT tbl_Member_Information.* " & _
                    " FROM tbl_Member_Information " & _
                    " WHERE (PK = " & .Item(i).SubItems(6) & ")"
                rt.Open t, ConnOmega
                If rt.RecordCount = 0 Then
                    MsgBox "Invalid Supplier!                   ", vbCritical, "Error...'"
                    rt.Close
                    Exit Sub
                End If
                rt.Close
            Case 3
                t = "SELECT tbl_Personnel_Information.* " & _
                    " FROM tbl_Personnel_Information " & _
                    " WHERE (PK = " & .Item(i).SubItems(6) & ")"
                rt.Open t, ConnOmega
                If rt.RecordCount = 0 Then
                    MsgBox "Invalid Supplier!                   ", vbCritical, "Error...'"
                    rt.Close
                    Exit Sub
                End If
                rt.Close
        End Select
    Next i
    '=== Posting
    For i = 1 To .Count
        s = "SELECT tbl_GL_Accounts.* " & _
            " FROM tbl_GL_Accounts " & _
            " WHERE (AccountCode = '" & .Item(i).SubItems(4) & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            ConnOmega.Execute "INSERT INTO tbl_GL_Transaction " & _
                              " (GLCode, DocDate, DocNumber, PayeeKey, PayeeType, SupplierCode, " & _
                              " SupplierName, BookType, Debit, Credit, Particulars) " & _
                              " VALUES ('" & .Item(i).SubItems(4) & "', '" & FormatDateTime(txtJVDate.Text, vbShortDate) & "', " & _
                              " '" & Trim(txtJVNumber.Text) & "', " & .Item(i).SubItems(6) & ", " & .Item(i).SubItems(2) & ", " & _
                              " '" & .Item(i).SubItems(7) & "', '" & FORMATSQL(.Item(i).SubItems(8)) & "', 6, " & _
                              " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(10)) = False, 0, .Item(i).SubItems(10))) & ", " & _
                              " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(11)) = False, 0, .Item(i).SubItems(11))) & ", " & _
                              " '" & FORMATSQL(.Item(i).SubItems(9)) & "')"
            If rs!withSL = 1 Then
                Select Case CLng(.Item(i).SubItems(2))
                    Case 1
                        'ConnOmega.Execute "INSERT INTO "
                    Case 2
                    
                    Case 3
                    
                End Select
            End If
        End If
        rs.Close
    Next i
End With
ConnOmega.Execute "UPDATE tbl_Acctg_GeneralJournal SET Posted = 1 WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
BROWSER GetSetting(App.EXEName, "GeneralJournal", "GenJour", ""), "is_LOAD"
End Sub

Private Sub PRESS_F9()
If picSLine.Visible = True Then Exit Sub
If picSeachAccount.Visible = True Then Exit Sub
If picSeachSL.Visible = True Then Exit Sub
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    Unload Me
Else
    If picSeachAccount.Visible = True Then
        cmdCancelAccount_Click
        Exit Sub
    End If
    If picSeachSL.Visible = True Then
        cmdCancelSL_Click
        Exit Sub
    End If
    If picSLine.Visible = True Then
        If TRANS_DETAIL = is_DET_ADDING Then
            With lstDetail.ListItems
                If .Count = 1 Then
                    .Item(iRow).SubItems(1) = " "
                    .Item(iRow).SubItems(2) = " "
                    .Item(iRow).SubItems(3) = " "
                    .Item(iRow).SubItems(4) = " "
                    .Item(iRow).SubItems(5) = " "
                    .Item(iRow).SubItems(6) = "0"
                    .Item(iRow).SubItems(7) = " "
                    .Item(iRow).SubItems(8) = " "
                    .Item(iRow).SubItems(9) = " "
                    .Item(iRow).SubItems(10) = " "
                    .Item(iRow).SubItems(11) = " "
                Else
                    .Remove .Count
                    iRow = .Count
                End If
            End With
        End If
        If TRANS_DETAIL = is_DET_EDITTING Then
            With lstDetail.ListItems
                .Item(iRow).SubItems(2) = txtJVType.Text
                .Item(iRow).SubItems(3) = txtType.Text
                .Item(iRow).SubItems(4) = txtAccCode1.Text
                .Item(iRow).SubItems(5) = txtAccDesc1.Text
                .Item(iRow).SubItems(6) = txtSLKey1.Text
                .Item(iRow).SubItems(7) = txtSLCode1.Text
                .Item(iRow).SubItems(8) = txtSLName1.Text
                .Item(iRow).SubItems(9) = txtDescription1.Text
                .Item(iRow).SubItems(10) = txtDebit1.Text
                .Item(iRow).SubItems(11) = txtCredit1.Text
            End With
        End If
        picSLine.Visible = False
        picMain.Enabled = True
        picToolbar.Enabled = True
        lstDetail.ListItems(iRow).EnsureVisible
        lstDetail.ListItems(iRow).Selected = True
        lstDetail.SetFocus
        Exit Sub
    Else
        CLEARTEXT
        LOCKTEXT True
        TOOLBARFUNC 1
        TRANSACTIONTYPE = is_REFRESH
        BROWSER GetSetting(App.EXEName, "GeneralJournal", "GenJour", ""), "is_LOAD"
        If Trim(txtJVNumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "GeneralJournal", "GenJour", ""), "is_HOME"
    End If
End If
End Sub

Private Sub CLEARTEXT()
iType = 0
txtJVNumber.Text = ""
txtJVDate.Text = ""
txtRemarks.Text = ""
lblTotalDebit.Caption = "0.00"
lblTotalCredit.Caption = "0.00"
imgPosted.Visible = False
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
lstDetail.ListItems.Clear
Set x = lstDetail.ListItems.Add()
x.Text = ""
x.SubItems(1) = " "
x.SubItems(2) = " "
x.SubItems(3) = " "
x.SubItems(4) = " "
x.SubItems(5) = " "
x.SubItems(6) = "0"
x.SubItems(7) = " "
x.SubItems(8) = " "
x.SubItems(9) = " "
x.SubItems(10) = " "
x.SubItems(11) = " "
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtJVNumber.Locked = True
txtJVDate.Locked = bln
txtRemarks.Locked = bln
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

Private Sub b8TitleBar1_CLoseClick()
cmdCancelAccount_Click
End Sub

Private Sub b8TitleBar2_CLoseClick()
cmdCancelSL_Click
End Sub

Private Sub cmbType_Click()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        iType = cmbType.ListIndex
        .Item(iRow).SubItems(2) = cmbType.ListIndex
        .Item(iRow).SubItems(3) = cmbType.List(cmbType.ListIndex)
        txtSLCode.Locked = IIf(iType = 0, True, False)
    End With
End If
End Sub

Private Sub cmbType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtAccCode.SetFocus
End Sub

Private Sub cmdCancelAccount_Click()
picSeachAccount.Visible = False
picSLine.Enabled = True
txtAccCode.SetFocus
End Sub

Private Sub cmdCancelSL_Click()
picSeachSL.Visible = False
picSLine.Enabled = True
txtSLCode.SetFocus
End Sub

Private Sub cmdOKAccount_Click()
If lstAccount.ListIndex = -1 Then Exit Sub
Arr = Split(lstAccount.List(lstAccount.ListIndex), " | ", -1, 1)
txtAccCode.Text = CStr(Arr(0))
txtAccDesc.Text = CStr(Arr(1))
cmdCancelAccount_Click
End Sub

Private Sub cmdOKSL_Click()
If lstSL.ListIndex = -1 Then Exit Sub
Arr = Split(lstSL.List(lstSL.ListIndex), " - ", -1, 1)
txtSLKey.Text = lstSL.ItemData(lstSL.ListIndex)
txtSLCode.Text = CStr(Arr(0))
txtSLName.Text = CStr(Arr(1))
cmdCancelSL_Click
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
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "GeneralJournal", "GenJour", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "GeneralJournal", "GenJour", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "GeneralJournal", "GenJour", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "GeneralJournal", "GenJour", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
picSLine.Width = 14055
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
With cmbType
    .Clear
    .AddItem "GL"
    .AddItem "Supplier"
    .AddItem "Member"
    .AddItem "Employee"
End With

CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANS_DETAIL = is_DET_REFRESH

BROWSER GetSetting(App.EXEName, "GeneralJournal", "GenJour", ""), "is_LOAD"
If Trim(txtJVNumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "GeneralJournal", "GenJour", ""), "is_HOME"

tmp = SetWindowLong(txtDescription.hwnd, GWL_STYLE, GetWindowLong(txtDescription.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearchSL.hwnd, GWL_STYLE, GetWindowLong(txtSearchSL.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearchAccount.hwnd, GWL_STYLE, GetWindowLong(txtSearchAccount.hwnd, GWL_STYLE) Or ES_UPPERCASE)

End Sub

Private Sub Form_Unload(Cancel As Integer)
If picSLine.Visible = True Then Cancel = -1
If picSeachAccount.Visible = True Then Cancel = -1
If picSeachSL.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lblTotalCredit_Change()
lblBalance.Caption = RETURNLABELVALUE(lblTotalDebit) - RETURNLABELVALUE(lblTotalCredit)
End Sub

Private Sub lblTotalDebit_Change()
lblBalance.Caption = RETURNLABELVALUE(lblTotalDebit) - RETURNLABELVALUE(lblTotalCredit)
End Sub

Private Sub lstAccount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAccount_Click
End Sub

Private Sub lstDetail_GotFocus()
iFocus = 1
iRow = lstDetail.SelectedItem.Index
TRANS_DETAIL = is_DET_REFRESH
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
With lstDetail.ListItems
    If Trim(.Item(iRow).SubItems(2)) = "" Then
        TOOLBARFUNC 4
    Else
        TOOLBARFUNC 5
    End If
End With
End Sub

Private Sub lstDetail_ItemClick(ByVal Item As MSComctlLib.ListItem)
iRow = lstDetail.SelectedItem.Index
End Sub

Private Sub lstDetail_LostFocus()
iFocus = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "GeneralJournal", "GenJour", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "GeneralJournal", "GenJour", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "GeneralJournal", "GenJour", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "GeneralJournal", "GenJour", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":   PRESS_F9
    Case "Post":    PRESS_F8
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtAccCode_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(4) = Trim(txtAccCode.Text)
    End With
End If
End Sub

Private Sub txtAccCode_GotFocus()
iFocusAcc = 1
End Sub

Private Sub txtAccCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtSLCode.SetFocus
End Sub

Private Sub txtAccCode_LostFocus()
iFocusAcc = 0
End Sub

Private Sub txtAccDesc_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(5) = Trim(txtAccDesc.Text)
    End With
End If
End Sub

Private Sub txtCredit_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(11) = IIf(RETURNTEXTVALUE(txtCredit) = 0, " ", Format(RETURNTEXTVALUE(txtCredit), "#,##0.00"))
        dCredit = 0
        For i = 1 To .Count
            dCredit = dCredit + CDbl(IIf(IsNumeric(.Item(i).SubItems(11)) = False, 0, .Item(i).SubItems(11)))
        Next i
        lblTotalCredit.Caption = Format(dCredit, "#,##0.00")
    End With
End If
End Sub

Private Sub txtCredit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If TRANS_DETAIL = is_DET_ADDING Then
        TRANS_DETAIL = is_DET_REFRESH
        Set x = lstDetail.ListItems.Add()
        x.Text = ""
        x.SubItems(1) = Format(lstDetail.ListItems.Count, "0#")
        x.SubItems(2) = " "
        x.SubItems(3) = " "
        x.SubItems(4) = " "
        x.SubItems(5) = " "
        x.SubItems(6) = "0"
        x.SubItems(7) = " "
        x.SubItems(8) = " "
        x.SubItems(9) = " "
        x.SubItems(10) = " "
        x.SubItems(11) = " "
        iRow = lstDetail.ListItems.Count
        lstDetail.ListItems(iRow).EnsureVisible
        lstDetail.ListItems(iRow).Selected = True
        
        cmbType.ListIndex = -1
        txtAccCode.Text = ""
        txtAccDesc.Text = ""
        txtSLKey.Text = "0"
        txtSLCode.Text = ""
        txtSLName.Text = ""
        txtDescription.Text = ""
        txtDebit.Text = ""
        txtCredit.Text = ""
        TRANS_DETAIL = is_DET_ADDING
        cmbType.SetFocus
        Exit Sub
    End If
    If TRANS_DETAIL = is_DET_EDITTING Then
        picSLine.Visible = False
        picMain.Enabled = True
        picToolbar.Enabled = True
        lstDetail.ListItems(iRow).EnsureVisible
        lstDetail.ListItems(iRow).Selected = True
        lstDetail.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub txtDebit_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(10) = IIf(RETURNTEXTVALUE(txtDebit) = 0, " ", Format(RETURNTEXTVALUE(txtDebit), "#,##0.00"))
        dDebit = 0
        For i = 1 To .Count
            dDebit = dDebit + CDbl(IIf(IsNumeric(.Item(i).SubItems(10)) = False, 0, .Item(i).SubItems(10)))
        Next i
        lblTotalDebit.Caption = Format(dDebit, "#,##0.00")
    End With
End If
End Sub

Private Sub txtDebit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtCredit.SetFocus
End Sub

Private Sub txtDescription_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(9) = Trim(txtDescription.Text)
    End With
End If
End Sub

Private Sub txtDescription_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtDebit.SetFocus
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

Private Sub txtSearchSL_Change()
If Trim(txtSearchSL.Text) = "" Then lstSL.Clear: Exit Sub
lstSL.Clear
Select Case iType
    Case 1
        s = "SELECT PK, SupplierCode as Code, SupplierName as sName " & _
            " From tbl_Inv_Supplier " & _
            " WHERE (SupplierName LIKE '" & FORMATSQL(Trim(txtSearchSL.Text)) & "%')"
    Case 2
        s = "SELECT PK, LastName + ',  ' + FirstName + '  ' + MiddleName AS sName, " & _
            " ISNULL((SELECT TOP 1 IDNumber From tbl_Member_IDNumber " & _
            " Where (MemberKey = tbl_Member_Information.PK) " & _
            " ORDER BY IDCounter DESC, IDNumber), '') AS Code " & _
            " From tbl_Member_Information " & _
            " WHERE (LastName LIKE '" & FORMATSQL(Trim(txtSearchSL.Text)) & "%') " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
    Case 3
        s = "SELECT PK, LastName + ',  ' + FirstName + '  ' + MiddleName AS sName, " & _
            " ISNULL((SELECT TOP 1 IDNumber From tbl_Personnel_IDNumber " & _
            " Where (ProfileKey = tbl_Personnel_Information.PK) " & _
            " ORDER BY IDNumber DESC), '') AS Code " & _
            " From tbl_Personnel_Information " & _
            " WHERE (LastName LIKE '" & FORMATSQL(Trim(txtSearchSL.Text)) & "%') " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
    Case Else: Exit Sub
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstSL.AddItem rs!Code & " - " & rs!sName
    lstSL.ItemData(lstSL.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstSL.ListCount Then lstSL.ListIndex = 0
End Sub

Private Sub txtSearchSL_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstSL.SetFocus
End Sub

Private Sub txtSLCode_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(7) = Trim(txtSLCode.Text)
    End With
End If
End Sub

Private Sub txtSLCode_GotFocus()
iFocusSL = 1
End Sub

Private Sub txtSLCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtDescription.SetFocus
End Sub

Private Sub txtSLCode_LostFocus()
iFocusSL = 0
End Sub

Private Sub txtSLKey_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(6) = Trim(txtSLKey.Text)
    End With
End If
End Sub

Private Sub txtSLName_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(8) = Trim(txtSLName.Text)
    End With
End If
End Sub
