VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelDeduction 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPersonnelDeduction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picSearchGLAccount 
      Height          =   2955
      Left            =   5520
      TabIndex        =   46
      Top             =   2160
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
         Picture         =   "frmPersonnelDeduction.frx":1CFA
         Style           =   1  'Graphical
         TabIndex        =   50
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
         Picture         =   "frmPersonnelDeduction.frx":236C
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   2355
         Width           =   1560
      End
      Begin VB.TextBox txtSearchGLAccount 
         Height          =   315
         Left            =   120
         TabIndex        =   48
         Top             =   480
         Width           =   5295
      End
      Begin VB.ListBox lstResultGLAccount 
         Height          =   1425
         Left            =   120
         TabIndex        =   47
         Top             =   840
         Width           =   5295
      End
      Begin RPVGCC.b8TitleBar b8TitleBar3 
         Height          =   345
         Left            =   45
         TabIndex        =   51
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
         Icon            =   "frmPersonnelDeduction.frx":2AC8
         ShadowVisible   =   0   'False
      End
   End
   Begin VB.PictureBox picADSLine 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4200
      ScaleHeight     =   855
      ScaleWidth      =   7695
      TabIndex        =   32
      Top             =   1440
      Visible         =   0   'False
      Width           =   7695
      Begin RPVGCC.b8Container picADSLine1 
         Height          =   855
         Left            =   0
         TabIndex        =   33
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
            TabIndex        =   41
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtDebit1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtAccountName1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtAccountNo1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtCredit 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5880
            TabIndex        =   37
            Top             =   360
            Width           =   1275
         End
         Begin VB.TextBox txtDebit 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4560
            TabIndex        =   36
            Top             =   360
            Width           =   1275
         End
         Begin VB.TextBox txtAccountName 
            Height          =   315
            Left            =   1320
            TabIndex        =   35
            Top             =   360
            Width           =   3195
         End
         Begin VB.TextBox txtAccountNo 
            Height          =   315
            Left            =   120
            TabIndex        =   34
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
            TabIndex        =   45
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
            TabIndex        =   44
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "ACCOUNT NAME"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1320
            TabIndex        =   43
            Top             =   120
            Width           =   3135
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "ACCOUNT #"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   120
            Width           =   975
         End
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   360
      ScaleHeight     =   3375
      ScaleWidth      =   11055
      TabIndex        =   4
      Top             =   1200
      Width           =   11055
      Begin VB.PictureBox picOneTime 
         BackColor       =   &H00C6B8A4&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   1935
         TabIndex        =   28
         Top             =   2160
         Width           =   1935
         Begin VB.CheckBox chkOneTime 
            BackColor       =   &H00C6B8A4&
            Caption         =   "Onetime Deduction"
            Height          =   375
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.TextBox txtMonthAmount 
         Height          =   315
         Left            =   240
         TabIndex        =   27
         Top             =   2160
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txtDateStart 
         Height          =   315
         Left            =   1920
         TabIndex        =   19
         Top             =   1800
         Width           =   1905
      End
      Begin VB.TextBox txtAmount 
         Height          =   315
         Left            =   1920
         TabIndex        =   17
         Top             =   3000
         Width           =   1905
      End
      Begin VB.TextBox txtNoofMonths 
         Height          =   315
         Left            =   1920
         TabIndex        =   15
         Top             =   2640
         Width           =   1905
      End
      Begin VB.TextBox txtTotalAmount 
         Height          =   315
         Left            =   1920
         TabIndex        =   13
         Top             =   1440
         Width           =   1905
      End
      Begin VB.ComboBox cmbDeductionType 
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Text            =   "cmbDeductionType"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   1920
         TabIndex        =   9
         Top             =   720
         Width           =   1905
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1920
         TabIndex        =   7
         Top             =   360
         Width           =   5505
      End
      Begin VB.TextBox txtCtrl 
         Height          =   315
         Left            =   1920
         TabIndex        =   5
         Top             =   0
         Width           =   1905
      End
      Begin MSComctlLib.ListView lstAccDistribution 
         Height          =   2250
         Left            =   3960
         TabIndex        =   31
         Top             =   1080
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   3969
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
      Begin VB.Label lblTotalDebit 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8280
         TabIndex        =   53
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblTotalCredit 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10080
         TabIndex        =   52
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "ACCOUNT DISTRIBUTION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3960
         TabIndex        =   30
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE START"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "DEDUCTION per SALARY"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "NO. OF MONTHS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL AMOUNT"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "DEDUCTION TYPE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NAME"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "CTRL NO."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
   End
   Begin RPVGCC.b8Container picAdd 
      Height          =   4335
      Left            =   3480
      TabIndex        =   21
      Top             =   480
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   7646
      BackColor       =   15266266
      Begin VB.CommandButton cmdOKAdd 
         Height          =   480
         Left            =   120
         Picture         =   "frmPersonnelDeduction.frx":3062
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3720
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelAdd 
         Height          =   480
         Left            =   1920
         Picture         =   "frmPersonnelDeduction.frx":36D4
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3720
         Width           =   1560
      End
      Begin VB.ListBox lstResultAdd 
         Height          =   2790
         Left            =   120
         TabIndex        =   23
         Top             =   890
         Width           =   3375
      End
      Begin VB.TextBox txtSearchAdd 
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   3375
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   40
         TabIndex        =   26
         Top             =   40
         Width           =   3530
         _ExtentX        =   6218
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
         Icon            =   "frmPersonnelDeduction.frx":3E30
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   770
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   15000
      TabIndex        =   0
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   570
         Left            =   0
         TabIndex        =   1
         Top             =   105
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   1005
         ButtonWidth     =   1058
         ButtonHeight    =   1005
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
               Caption         =   "Post"
               Key             =   "Post"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   8460
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   2
            Top             =   0
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmPersonnelDeduction.frx":43CA
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
         Y1              =   690
         Y2              =   690
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
         Y1              =   750
         Y2              =   750
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8880
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483648
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeduction.frx":4ADD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeduction.frx":4BDF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeduction.frx":4D63
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeduction.frx":507D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeduction.frx":5436
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeduction.frx":5888
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeduction.frx":5CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeduction.frx":6092
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeduction.frx":61A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeduction.frx":66E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeduction.frx":6B38
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeduction.frx":6C92
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeduction.frx":71D4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   5070
      Width           =   11790
      _ExtentX        =   20796
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
End
Attribute VB_Name = "frmPersonnelDeduction"
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

Dim tmp             As Long

Dim is_DET_FOCUS    As Long
Dim ROW             As Long
Dim isGLCodeFocus   As Long
Dim isGLCodeChange  As Long

Dim a, b, x, i, iEmployee, iDeductionType, sCtrl

Private Sub BROWSER(Ctrl, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If Ctrl <> "" Then
            s = "SELECT TOP 1 tbl_Personnel_Deduction.* " & _
                " From tbl_Personnel_Deduction " & _
                " WHERE (Ctrl = '" & Ctrl & "') " & _
                " ORDER BY Ctrl"
        Else
            s = "SELECT TOP 1 tbl_Personnel_Deduction.* " & _
                " From tbl_Personnel_Deduction " & _
                " ORDER BY Ctrl"
        End If
    Case "is_HOME"
        If picAdd.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Personnel_Deduction.* " & _
            " From tbl_Personnel_Deduction " & _
            " ORDER BY Ctrl"
    Case "is_PAGEUP"
        If picAdd.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Personnel_Deduction.* " & _
            " From tbl_Personnel_Deduction " & _
            " WHERE (Ctrl < '" & Ctrl & "') " & _
            " ORDER BY Ctrl DESC"
    Case "is_PAGEDOWN"
        If picAdd.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Personnel_Deduction.* " & _
            " From tbl_Personnel_Deduction " & _
            " WHERE (Ctrl > '" & Ctrl & "') " & _
            " ORDER BY Ctrl"
    Case "is_END"
        If picAdd.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Personnel_Deduction.* " & _
            " From tbl_Personnel_Deduction " & _
            " ORDER BY Ctrl DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iEmployee = rs!EmployeeKey
    iDeductionType = rs!DeductionType
    txtCtrl.Text = rs!Ctrl
    
    txtName.Text = ""
    t = "SELECT tbl_Personnel_IDNumber.IDNumber, " & _
        " tbl_Personnel_Information.LastName, " & _
        " tbl_Personnel_Information.FirstName, " & _
        " tbl_Personnel_Information.MiddleName " & _
        " FROM tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
        " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
        " WHERE (tbl_Personnel_IDNumber.PK = " & iEmployee & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtName.Text = rt!IDNumber & " - " & rt!LastName & ",  " & rt!FirstName & "  " & rt!MiddleName
    End If
    rt.Close
    
    txtDate.Text = Format(rs!TransDate, "mm/dd/yyyy")
    
    cmbDeductionType.ListIndex = -1
    cmbDeductionType.Text = ""
    
    t = "SELECT DeductionTypeName " & _
        " From tbl_Personnel_Deduction_Type " & _
        " WHERE (PK = " & iDeductionType & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        cmbDeductionType.Text = rt!DeductionTypeName
    End If
    rt.Close
    
    txtTotalAmount.Text = Format(rs!TotalAmount, "#,##0.00")
    chkOneTime.Value = rs!OneTimeDed
    txtDateStart.Text = Format(rs!DateStart, "mm/dd/yyyy")
    txtNoofMonths.Text = rs!NoMonths
    txtAmount.Text = Format(rs!Amount, "#,##0.00")
    imgPosted.Visible = IIf(rs!Posted = 1, True, False)
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    
    SaveSetting App.EXEName, "DeductionCtrl", "DedCtrl", rs!Ctrl
    
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If picAdd.Visible = True Then Exit Sub
If picADSLine.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then

    If AccessRights("Personnel Deduction", "Add") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    
    picMain.Enabled = False
    picToolbar.Enabled = False
    picAdd.ZOrder 0
    txtSearchAdd.Text = ""
    picAdd.Visible = True
    txtSearchAdd.SetFocus

Else
    If is_DET_FOCUS = 0 Then Exit Sub
    With lstAccDistribution.ListItems
        If .Count > 1 Then
            Set x = .Add()
            x.Text = ""
            x.SubItems(1) = " "
            x.SubItems(2) = " "
            x.SubItems(3) = " "
            x.SubItems(4) = " "
            ROW = .Count
        Else
            If Trim(.Item(.Count).SubItems(1)) <> "" Then
                Set x = .Add()
                x.Text = ""
                x.SubItems(1) = " "
                x.SubItems(2) = " "
                x.SubItems(3) = " "
                x.SubItems(4) = " "
                ROW = .Count
            Else
                ROW = 1
            End If
        End If
        lstAccDistribution.ListItems(ROW).EnsureVisible
        lstAccDistribution.ListItems(ROW).Selected = True
        TRANS_DETAIL = is_DET_ADDING
        isGLCodeChange = 1
        txtAccountNo.Text = ""
        txtAccountName.Text = ""
        txtDebit.Text = ""
        txtCredit.Text = ""
        picADSLine.Height = 855
        picADSLine.Width = 7335
        picADSLine.ZOrder 0
        picMain.Enabled = False
        picToolbar.Enabled = False
'        picAccDistribution.Enabled = False
        picADSLine.Visible = True
        txtAccountNo.SetFocus
    End With
End If
End Sub

Private Sub PRESS_F2()
If picAdd.Visible = True Then Exit Sub
If picADSLine.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    If Statusbar1.Panels(1).Text = "" Then Exit Sub
    
    If AccessRights("Personnel Deduction", "Edit") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    
    If imgPosted.Visible = True Then MsgBox "Already Posted!                         ", vbCritical, "Error...": Exit Sub
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_EDITTING
Else
    If is_DET_FOCUS = 0 Then Exit Sub
    With lstAccDistribution.ListItems
        If Trim(.Item(ROW).SubItems(1)) = "" Then Exit Sub
        txtAccountNo.Text = .Item(ROW).SubItems(1)
        txtAccountName.Text = .Item(ROW).SubItems(2)
        txtDebit.Text = .Item(ROW).SubItems(3)
        txtCredit.Text = .Item(ROW).SubItems(4)
        txtAccountNo1.Text = .Item(ROW).SubItems(1)
        txtAccountName1.Text = .Item(ROW).SubItems(2)
        txtDebit1.Text = .Item(ROW).SubItems(3)
        txtCredit1.Text = .Item(ROW).SubItems(4)
    End With
    TRANS_DETAIL = is_DET_EDITTING
    isGLCodeChange = 1
    picADSLine.Height = 855
    picADSLine.Width = 7335
    picADSLine.ZOrder 0
    picMain.Enabled = False
    picToolbar.Enabled = False
'    picAccDistribution.Enabled = False
    picADSLine.Visible = True
    txtAccountNo.SetFocus
End If
End Sub

Private Sub PRESS_DELETE()
If picAdd.Visible = True Then Exit Sub
If picADSLine.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    If Statusbar1.Panels(1).Text = "" Then Exit Sub
    
    If AccessRights("Personnel Deduction", "Delete") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    
    If imgPosted.Visible = True Then MsgBox "Already Posted!                         ", vbCritical, "Error...": Exit Sub
    If MsgBox("ARE YOU SURE IN DELETING THIS RECORD?                            ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    On Error GoTo PG:
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Deduction WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    CLEARTEXT
    BROWSER GetSetting(App.EXEName, "DeductionCtrl", "DedCtrl", ""), "is_PAGEDOWN"
    If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "DeductionCtrl", "DedCtrl", ""), "is_HOME"
Else
    If is_DET_FOCUS = 0 Then Exit Sub
    With lstAccDistribution.ListItems
    If .Count > 1 Then
        .Remove ROW
        If CDbl(ROW) > CDbl(.Count) Then ROW = .Count
    Else
        .Item(1).SubItems(1) = " "
        .Item(1).SubItems(2) = " "
        .Item(1).SubItems(3) = " "
        .Item(1).SubItems(4) = " "
        ROW = 1
    End If
    lstAccDistribution.ListItems(ROW).EnsureVisible
    lstAccDistribution.ListItems(ROW).Selected = True
    isGLCodeChange = 1
End With
End If
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If iEmployee = 0 Then MsgBox "Please Select Employee!                       ", vbCritical, "Error...": Exit Sub
If iDeductionType = 0 Then MsgBox "Please Select Deduction Type!                     ", vbCritical, "Error...": cmbDeductionType.SetFocus: Exit Sub
If IsDate(txtDate.Text) = False Then MsgBox "Please Supply a Valid Date!                      ", vbCritical, "Error...": txtDate.SetFocus: Exit Sub
If RETURNTEXTVALUE(txtTotalAmount) <= 0 Then MsgBox "Please Supply a Valid Value!                         ", vbCritical, "Error...": txtTotalAmount.SetFocus: Exit Sub
If IsDate(txtDateStart.Text) = False Then MsgBox "Please Supply a Valid Date!                      ", vbCritical, "Error...": txtDateStart.SetFocus: Exit Sub
If chkOneTime.Value = 0 Then If Val(txtNoofMonths) <= 0 Then MsgBox "Please Supply a Valid Value!                         ", vbCritical, "Error...": txtNoofMonths.SetFocus: Exit Sub
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    sCtrl = ""
    s = "SELECT TOP 1 Ctrl " & _
        " FROM tbl_Personnel_Deduction " & _
        " WHERE (Year(TransDate) = " & Format(txtDate.Text, "yyyy") & ") " & _
        " ORDER BY Ctrl DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        sCtrl = Format(CDbl(rs!Ctrl) + 1, "0000000#")
    Else
        sCtrl = Format(txtDate.Text, "yyyy") & "0000"
    End If
    rs.Close
    Do
        s = "SELECT tbl_Personnel_Deduction.* " & _
            " FROM tbl_Personnel_Deduction " & _
            " WHERE (Ctrl = '" & sCtrl & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If
        rs.Close
        sCtrl = Format(CDbl(sCtrl) + 1, "0000000#")
    Loop
    
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Deduction " & _
                      " (Ctrl, EmployeeKey, TransDate, DeductionType, TotalAmount, " & _
                      " DateStart, OneTimeDed, NoMonths, Amount, LastModified) " & _
                      " VALUES ('" & sCtrl & "', " & iEmployee & ", " & _
                      " '" & FormatDateTime(txtDate.Text, vbShortDate) & "', " & _
                      " " & iDeductionType & ", " & RETURNTEXTVALUE(txtTotalAmount) & ", " & _
                      " '" & FormatDateTime(txtDateStart.Text, vbShortDate) & "', " & _
                      " " & chkOneTime.Value & ", " & RETURNTEXTVALUE(txtNoofMonths) & ", " & _
                      " " & RETURNTEXTVALUE(txtAmount) & ", '" & CStr(Now) & " - " & gbl_CompleteName & "')"
                      
End If
If TRANSACTIONTYPE = is_EDITTING Then
    sCtrl = GetSetting(App.EXEName, "DeductionCtrl", "DedCtrl", "")
    ConnOmega.Execute "UPDATE tbl_Personnel_Deduction " & _
                      " SET TransDate = '" & FormatDateTime(txtDate.Text, vbShortDate) & "', " & _
                      " DeductionType = " & iDeductionType & ", " & _
                      " TotalAmount = " & RETURNTEXTVALUE(txtTotalAmount) & ", " & _
                      " DateStart = '" & FormatDateTime(txtDateStart.Text, vbShortDate) & "', " & _
                      " OneTimeDed = " & chkOneTime.Value & ", " & _
                      " NoMonths = " & RETURNTEXTVALUE(txtNoofMonths) & ", " & _
                      " Amount = " & RETURNTEXTVALUE(txtAmount) & ", " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
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
If picAdd.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub

End Sub

Private Sub PRESS_F8()
If picAdd.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub

If AccessRights("Personnel Deduction", "Post") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If

If imgPosted.Visible = True Then MsgBox "Already Posted!                         ", vbCritical, "Error...": Exit Sub

With lstAccDistribution.ListItems
    a = 0
    For i = 1 To .Count
        If Trim(.Item(i).SubItems(1)) <> "" Then
            a = a + 1
        End If
    Next i
End With
If CDbl(a) = 0 Then MsgBox "Check Account Distribution!                       ", vbCritical, "Error...": Exit Sub

If MsgBox("ARE YOU SURE IN POSTING THIS TRANSACTION?                        ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
On Error GoTo PG:
ConnOmega.Execute "INSERT INTO tbl_Personnel_Deduction_Summary " & _
                  " (EmployeeKey, DeductionType, TransDate, InOut, AmountIn, DocNumber, InRefKey) " & _
                  " VALUES (" & iEmployee & ", " & iDeductionType & ", " & _
                  " '" & FormatDateTime(txtDate.Text, vbShortDate) & "', 'I', " & _
                  " " & RETURNTEXTVALUE(txtTotalAmount) & ", '" & txtCtrl.Text & "', " & _
                  " " & Statusbar1.Panels(1).Text & ")"
ConnOmega.Execute "UPDATE tbl_Personnel_Deduction " & _
                  " SET Posted = 1 " & _
                  " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
BROWSER GetSetting(App.EXEName, "DeductionCtrl", "DedCtrl", ""), "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F9()
If picAdd.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub

End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picAdd.Visible = True Then cmdCancelAdd_Click: Exit Sub
    Unload Me
Else
    If picADSLine.Visible = True Then
        With lstAccDistribution.ListItems
            If TRANS_DETAIL = is_DET_ADDING Then
                If .Count > 1 Then
                    .Remove .Count
                    ROW = .Count
                Else
                    ROW = 1
                    .Item(ROW).SubItems(1) = " "
                    .Item(ROW).SubItems(2) = " "
                    .Item(ROW).SubItems(3) = " "
                    .Item(ROW).SubItems(4) = " "
                End If
            ElseIf TRANS_DETAIL = is_DET_EDITTING Then
                .Item(ROW).SubItems(1) = txtAccountNo1.Text
                .Item(ROW).SubItems(2) = txtAccountName1.Text
                .Item(ROW).SubItems(3) = txtDebit1.Text
                .Item(ROW).SubItems(4) = txtCredit1.Text
            End If
        End With
        picADSLine.Visible = False
        'picAccDistribution.Enabled = True
        picMain.Enabled = True
        picToolbar.Enabled = True
        lstAccDistribution.SetFocus
        Exit Sub
    End If
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "DeductionCtrl", "DedCtrl", ""), "is_LOAD"
    If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "DeductionCtrl", "DedCtrl", ""), "is_HOME"
End If
End Sub

Private Sub CLEARTEXT()
iEmployee = 0
iDeductionType = 0
txtCtrl.Text = ""
txtName.Text = ""
txtDate.Text = ""
cmbDeductionType.ListIndex = -1
cmbDeductionType.Text = ""
txtTotalAmount.Text = ""
chkOneTime.Value = 0
txtDateStart.Text = ""
txtNoofMonths.Text = ""
txtAmount.Text = ""
imgPosted.Visible = False
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
lstAccDistribution.ListItems.Clear
Set x = lstAccDistribution.ListItems.Add()
x.Text = ""
x.SubItems(1) = " "
x.SubItems(2) = " "
x.SubItems(3) = " "
x.SubItems(4) = " "
x.SubItems(5) = "0"
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtCtrl.Locked = True
txtName.Locked = True
txtDate.Locked = bln
cmbDeductionType.Locked = bln
txtTotalAmount.Locked = bln
txtDateStart.Locked = bln
txtNoofMonths.Locked = bln
txtAmount.Locked = True
picOneTime.Enabled = IIf(bln = True, False, True)
End Sub


Private Sub TOOLBARFUNC(intSelect As Integer)
With Toolbar1
    Select Case intSelect
        Case 1      '=== REFRESH ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(1).Enabled = True
            .Buttons(3).Enabled = True
            .Buttons(5).Enabled = True
            .Buttons(7).Image = 4
            .Buttons(7).Caption = "First"
            .Buttons(9).Image = 5
            .Buttons(9).Caption = "Back"
            .Buttons(7).Enabled = True
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = True
            .Buttons(13).Enabled = True
            .Buttons(15).Enabled = True
            .Buttons(17).Enabled = True
            .Buttons(19).Enabled = True
            .Buttons(21).Enabled = True
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
            .Buttons(21).ToolTipText = "CLOSE (Esc)"
        Case 2      '=== ADD/EDIT ====
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(1).Enabled = False
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = False
            .Buttons(7).Image = 12
            .Buttons(7).Caption = "Save"
            .Buttons(9).Image = 13
            .Buttons(9).Caption = "Undo"
            .Buttons(7).Enabled = True
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
            .Buttons(19).Enabled = False
            .Buttons(21).Enabled = False
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
        Case 3      '=== FIND ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(1).Enabled = False
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = False
            .Buttons(7).Image = 4
            .Buttons(7).Caption = "First"
            .Buttons(9).Image = 13
            .Buttons(9).Caption = "Undo"
            .Buttons(7).Enabled = False
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
            .Buttons(19).Enabled = False
            .Buttons(21).Enabled = False
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
        Case 4      '=== EMPTY DETAIL ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(1).Enabled = True
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = False
            .Buttons(7).Image = 12
            .Buttons(7).Caption = "Save"
            .Buttons(9).Image = 13
            .Buttons(9).Caption = "Undo"
            .Buttons(7).Enabled = True
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
            .Buttons(19).Enabled = False
            .Buttons(21).Enabled = False
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
        Case 5      '=== NOT EMPTY DETAIL ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(1).Enabled = True
            .Buttons(3).Enabled = True
            .Buttons(5).Enabled = True
            .Buttons(7).Image = 12
            .Buttons(7).Caption = "Save"
            .Buttons(9).Image = 13
            .Buttons(9).Caption = "Undo"
            .Buttons(7).Enabled = True
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
            .Buttons(19).Enabled = False
            .Buttons(21).Enabled = False
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
    End Select
End With
End Sub

Private Sub b8TitleBar1_CLoseClick()
cmdCancelAdd_Click
End Sub

Private Sub chkOneTime_Click()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If chkOneTime.Value = 0 Then Exit Sub
    txtNoofMonths.Text = ""
    txtAmount.Text = txtTotalAmount.Text
End If
End Sub

Private Sub cmbDeductionType_Click()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If cmbDeductionType.ListIndex = -1 Then Exit Sub
    iDeductionType = cmbDeductionType.ItemData(cmbDeductionType.ListIndex)
End If
End Sub

Private Sub cmdCancelAdd_Click()
picAdd.Visible = False
picToolbar.Enabled = True
picMain.Enabled = True
End Sub

Private Sub cmdOKAdd_Click()
If lstResultAdd.ListIndex = -1 Then Exit Sub
CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
iEmployee = lstResultAdd.ItemData(lstResultAdd.ListIndex)
txtName.Text = lstResultAdd.List(lstResultAdd.ListIndex)
cmdCancelAdd_Click
txtDate.SetFocus
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
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "DeductionCtrl", "DedCtrl", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "DeductionCtrl", "DedCtrl", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "DeductionCtrl", "DedCtrl", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "DeductionCtrl", "DedCtrl", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 2
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2

POPULATE_COMBO "PK", "DeductionTypeName", "tbl_Personnel_Deduction_Type", "PK", cmbDeductionType
isGLCodeFocus = 0
ROW = 0
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANS_DETAIL = is_DET_REFRESH
BROWSER GetSetting(App.EXEName, "DeductionCtrl", "DedCtrl", ""), "is_LOAD"
If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "DeductionCtrl", "DedCtrl", ""), "is_HOME"

tmp = SetWindowLong(txtSearchAdd.hwnd, GWL_STYLE, GetWindowLong(txtSearchAdd.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub


Private Sub Form_Unload(Cancel As Integer)
If picAdd.Visible = True Then Cancel = -1
If picADSLine.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstAccDistribution_GotFocus()
is_DET_FOCUS = 1
ROW = lstAccDistribution.SelectedItem.Index
TRANS_DETAIL = is_DET_REFRESH
End Sub

Private Sub lstAccDistribution_ItemClick(ByVal Item As MSComctlLib.ListItem)
ROW = lstAccDistribution.SelectedItem.Index
End Sub

Private Sub lstAccDistribution_LostFocus()
is_DET_FOCUS = 0
End Sub

Private Sub lstResultAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAdd_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "DeductionCtrl", "DedCtrl", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "DeductionCtrl", "DedCtrl", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "DeductionCtrl", "DedCtrl", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "DeductionCtrl", "DedCtrl", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":   PRESS_F9
    Case "Post":    PRESS_F8
    Case "Close":   PRESS_ESCAPE
End Select
End Sub


Private Sub txtAccountName_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstAccDistribution.ListItems
        .Item(ROW).SubItems(2) = Trim(txtAccountName.Text)
    End With
End If
End Sub

Private Sub txtAccountName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF6 Then
    If picADSLine.Visible = False Then Exit Sub
    picADSLine.Enabled = False
    picSearchGLAccount.ZOrder 0
    txtSearchGLAccount.Text = ""
    picSearchGLAccount.Visible = True
    txtSearchGLAccount.SetFocus
ElseIf KeyCode = vbKeyReturn Then
    txtDebit.SetFocus
End If
End Sub

Private Sub txtAccountNo_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstAccDistribution.ListItems
        .Item(ROW).SubItems(1) = Trim(txtAccountNo.Text)
    End With
End If
End Sub

Private Sub txtAccountNo_GotFocus()
isGLCodeFocus = 1
End Sub

Private Sub txtAccountNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If Trim(txtAccountNo.Text) <> "" Then
        s = "SELECT tbl_GL_Accounts.* " & _
            " FROM tbl_GL_Accounts " & _
            " WHERE (AccountCode = '" & Trim(txtAccountNo.Text) & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            txtAccountNo.Text = rs!AccountCode
            txtAccountName.Text = rs!AccountName
        Else
            MsgBox "Account Code '" & Trim(txtAccountNo.Text) & "' not Found!                           ", vbCritical, "Error..."
            rs.Close
            Exit Sub
        End If
        rs.Close
    End If
    txtDebit.SetFocus
End If
If KeyCode = vbKeyF6 Then
    picADSLine.Enabled = False
    txtSearchGLAccount.Text = ""
    picSearchGLAccount.ZOrder 0
    picSearchGLAccount.Visible = True
    txtSearchGLAccount.SetFocus
End If
End Sub

Private Sub txtAccountNo_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtAccountNo_LostFocus()
isGLCodeFocus = 0
End Sub


Private Sub txtCredit_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstAccDistribution.ListItems
        .Item(ROW).SubItems(4) = IIf(RETURNTEXTVALUE(txtCredit) = 0, " ", Format(RETURNTEXTVALUE(txtCredit), "#,##0.00"))
        .Item(ROW).SubItems(5) = Format(RETURNTEXTVALUE(txtDebit) - RETURNTEXTVALUE(txtCredit), "#,##0.00")
        b = 0
        For i = 1 To .Count
            b = b + CDbl(IIf(IsNumeric(.Item(i).SubItems(4)) = False, 0, .Item(i).SubItems(4)))
        Next i
        lblTotalCredit.Caption = Format(b, "#,##0.00")
    End With
End If
End Sub

Private Sub txtCredit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    picADSLine.Visible = False
    'picAccDistribution.Enabled = True
    picMain.Enabled = True
    picToolbar.Enabled = True
    lstAccDistribution.SetFocus
End If
End Sub

Private Sub txtCredit_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub
Private Sub txtDebit_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstAccDistribution.ListItems
        .Item(ROW).SubItems(3) = IIf(RETURNTEXTVALUE(txtDebit) = 0, " ", Format(RETURNTEXTVALUE(txtDebit), "#,##0.00"))
        .Item(ROW).SubItems(5) = Format(RETURNTEXTVALUE(txtDebit) - RETURNTEXTVALUE(txtCredit), "#,##0.00")
        b = 0
        For i = 1 To .Count
            b = b + CDbl(IIf(IsNumeric(.Item(i).SubItems(3)) = False, 0, .Item(i).SubItems(3)))
        Next i
        lblTotalDebit.Caption = Format(b, "#,##0.00")
    End With
End If
End Sub

Private Sub txtDebit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtCredit.SetFocus
End Sub

Private Sub txtDebit_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub
Private Sub txtSearchGLAccount_Change()
If Trim(txtSearchGLAccount.Text) = "" Then lstResultGLAccount.Clear: Exit Sub
lstResultGLAccount.Clear
s = "SELECT tbl_GL_Accounts.* " & _
    " FROM tbl_GL_Accounts " & _
    " WHERE (AccountName LIKE '" & FORMATSQL(Trim(txtSearchGLAccount.Text)) & "%') " & _
    " ORDER BY AccountName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultGLAccount.AddItem rs!AccountCode & " : " & rs!AccountName
    rs.MoveNext
Wend
rs.Close
If lstResultGLAccount.ListCount Then lstResultGLAccount.ListIndex = 0
End Sub

Private Sub txtSearchGLAccount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultGLAccount.SetFocus
End Sub

Private Sub txtDate_LostFocus()
If IsDate(txtDate.Text) = True Then
    txtDate.Text = Format(FormatDateTime(txtDate.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub

Private Sub txtDateStart_LostFocus()
If IsDate(txtDateStart.Text) = True Then
    txtDateStart.Text = Format(FormatDateTime(txtDateStart.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub

Private Sub txtMonthAmount_Change()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If chkOneTime.Value = 1 Then Exit Sub
    txtAmount.Text = Format(RETURNTEXTVALUE(txtMonthAmount) / 2, "#,##0.00")
End If
End Sub

Private Sub txtNoofMonths_Change()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If chkOneTime.Value = 1 Then Exit Sub
    On Error Resume Next
    txtMonthAmount.Text = Format(RETURNTEXTVALUE(txtTotalAmount) / RETURNTEXTVALUE(txtNoofMonths), "#,##0.00")
End If
End Sub

Private Sub txtSearchAdd_Change()
If Trim(txtSearchAdd.Text) = "" Then lstResultAdd.Clear: Exit Sub
lstResultAdd.Clear
s = "SELECT tbl_Personnel_IDNumber.PK, " & _
    " tbl_Personnel_IDNumber.IDNumber, " & _
    " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName " & _
    " FROM tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
    " WHERE (tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%') " & _
    " AND (ISNULL((SELECT TOP 1 tbl_Personnel_EmploymentStatus.Active " & _
    " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
    " tbl_Personnel_EmploymentStatus ON tbl_Personnel_Action.EmpStatus = tbl_Personnel_EmploymentStatus.PK " & _
    " WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_IDNumber.PK) " & _
    " AND (tbl_Personnel_Action.EffectivityDate <= CONVERT(DATETIME, CONVERT(char(6), getdate(), 12), 102)) ORDER BY tbl_Personnel_Action.EffectivityDate DESC), 0) = 1) " & _
    " ORDER BY tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName"
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

Private Sub txtSearchAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultAdd.SetFocus
End Sub

Private Sub txtTotalAmount_LostFocus()
txtTotalAmount.Text = Format(RETURNTEXTVALUE(txtTotalAmount), "#,##0.00")
End Sub
