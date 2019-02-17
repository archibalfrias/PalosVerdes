VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelDeductions 
   Appearance      =   0  'Flat
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11505
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
   ScaleHeight     =   6120
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picPrintSumm 
      Height          =   2175
      Left            =   4080
      TabIndex        =   51
      Top             =   1920
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3836
      BackColor       =   15396057
      Begin VB.ComboBox cmbDivision 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   600
         Width           =   2655
      End
      Begin VB.CommandButton cmdOKSumm 
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
         Picture         =   "frmPersonnelDeductions.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   1440
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelSumm 
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
         Picture         =   "frmPersonnelDeductions.frx":0672
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   1440
         Width           =   1560
      End
      Begin VB.TextBox txtAsOf 
         Height          =   315
         Left            =   1080
         TabIndex        =   52
         Top             =   960
         Width           =   2655
      End
      Begin VB.Timer TimerSummary 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2880
         Top             =   1440
      End
      Begin RPVGCC.b8TitleBar b8TitleBar3 
         Height          =   345
         Left            =   45
         TabIndex        =   55
         Top             =   45
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   609
         Caption         =   "Print Summary"
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
         Icon            =   "frmPersonnelDeductions.frx":0DCE
         ShadowVisible   =   0   'False
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   58
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "as of"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   56
         Top             =   960
         Width           =   1215
      End
   End
   Begin RPVGCC.b8Container picAdd 
      Height          =   4095
      Left            =   4080
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   7223
      BackColor       =   15396057
      Begin VB.Timer TimerOutStanding 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   120
         Top             =   3120
      End
      Begin VB.TextBox txtPayrollDateAdd 
         Height          =   315
         Left            =   1680
         TabIndex        =   41
         Top             =   3120
         Width           =   1575
      End
      Begin VB.ListBox lstResultAdd 
         Height          =   2205
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox txtSearchAdd 
         Height          =   315
         Left            =   120
         TabIndex        =   14
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
         Picture         =   "frmPersonnelDeductions.frx":1368
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3480
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
         Picture         =   "frmPersonnelDeductions.frx":1AC4
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3480
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   16
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
         Icon            =   "frmPersonnelDeductions.frx":2136
         ShadowVisible   =   0   'False
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   42
         Top             =   3120
         Width           =   1215
      End
   End
   Begin RPVGCC.b8Container picSearch 
      Height          =   3855
      Left            =   2880
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6800
      BackColor       =   15396057
      Begin VB.ComboBox cmbResultSearch 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2880
         Width           =   5895
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
         Left            =   1320
         Picture         =   "frmPersonnelDeductions.frx":26D0
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3240
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
         Left            =   3000
         Picture         =   "frmPersonnelDeductions.frx":2D42
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3240
         Width           =   1560
      End
      Begin VB.TextBox txtSearchSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   5895
      End
      Begin VB.ListBox lstResultSearch 
         Height          =   2010
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   5895
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   45
         TabIndex        =   22
         Top             =   45
         Width           =   6045
         _ExtentX        =   10663
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
         Icon            =   "frmPersonnelDeductions.frx":349E
         ShadowVisible   =   0   'False
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   600
      ScaleHeight     =   4335
      ScaleWidth      =   10335
      TabIndex        =   4
      Top             =   1320
      Width           =   10335
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   8880
         TabIndex        =   59
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtCutOffDate 
         Height          =   315
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   360
         Width           =   4215
      End
      Begin MSComctlLib.ListView lstDetails 
         Height          =   3135
         Left            =   0
         TabIndex        =   24
         Top             =   840
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   5530
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
            Text            =   "DeductionKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Deduction Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "DedPerKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Deduction Period"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Ded. Per Period"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Remarks"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   0
         Width           =   6975
      End
      Begin VB.TextBox txtControl 
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPerPayroll 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "100,000.00"
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
         Left            =   6840
         TabIndex        =   27
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label lblTotalAmount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "100,000.00"
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
         Left            =   3960
         TabIndex        =   26
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
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
         Left            =   2520
         TabIndex        =   25
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cut-Off Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3360
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10200
      Top             =   1200
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
            Picture         =   "frmPersonnelDeductions.frx":3A38
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductions.frx":4712
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductions.frx":53EC
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductions.frx":60C6
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductions.frx":6DA0
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductions.frx":7A7A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductions.frx":8754
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductions.frx":942E
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductions.frx":A108
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductions.frx":A9E2
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductions.frx":B6BC
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductions.frx":C396
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductions.frx":D070
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductions.frx":DD4A
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductions.frx":EA24
            Key             =   "IMG15"
         EndProperty
      EndProperty
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
         MouseIcon       =   "frmPersonnelDeductions.frx":F6FE
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
               Picture         =   "frmPersonnelDeductions.frx":FA18
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
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   5820
      Width           =   11505
      _ExtentX        =   20294
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
   Begin RPVGCC.b8Container picSLLines 
      Height          =   855
      Left            =   480
      TabIndex        =   28
      Top             =   1320
      Visible         =   0   'False
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox cmbDeductionPeriod1 
         Height          =   315
         Left            =   9360
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   50
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtDeductionPeriodKey1 
         Height          =   315
         Left            =   9120
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   49
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtDeductionPeriodKey 
         Height          =   315
         Left            =   5520
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.ComboBox cmbDeductionPeriod 
         Height          =   315
         ItemData        =   "frmPersonnelDeductions.frx":1012B
         Left            =   5400
         List            =   "frmPersonnelDeductions.frx":1012D
         TabIndex        =   46
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtDeductionKey1 
         Height          =   315
         Left            =   8640
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   45
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox cmbDeductionName1 
         Height          =   315
         Left            =   8880
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   44
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtAmount1 
         Height          =   315
         Left            =   9600
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtSLRemarks 
         Height          =   315
         Left            =   8160
         TabIndex        =   38
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtPerPayroll 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7200
         TabIndex        =   36
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox cmbDeductionName 
         Height          =   315
         ItemData        =   "frmPersonnelDeductions.frx":1012F
         Left            =   120
         List            =   "frmPersonnelDeductions.frx":10131
         TabIndex        =   35
         Top             =   360
         Width           =   3975
      End
      Begin VB.TextBox txtSLRemarks1 
         Height          =   315
         Left            =   10080
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   32
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtPerPayroll1 
         Height          =   315
         Left            =   9840
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtDeductionKey 
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4200
         TabIndex        =   29
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Deduction Period"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5400
         TabIndex        =   47
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8160
         TabIndex        =   39
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Per Payroll"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7200
         TabIndex        =   37
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4200
         TabIndex        =   34
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Deduction Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   120
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmPersonnelDeductions"
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

Private TRANS_DETAIL As Long
Const is_DET_REFRESH = 0
Const is_DET_ADDING = 1
Const is_DET_EDITTING = 2


Public isAddPrint As Long
Dim isFocus, iRow       As Long

Public locEmployeePK, locPayrollPeroid    As Long

Dim Array1, i, x, sCtrl, iPK, iEmpStatus, locPayrollPeroidTmp, Array2
Dim dTotalAmt, dTotalPerPayroll, iMasterKey, iLine, dRunBal, iDedKey, strDedKeySourceKey


Private Sub PRESS_INSERT()
If picPrintSumm.Visible = True Then Exit Sub
If picSLLines.Visible = True Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    If AccessRights("Personnel Deduction", "Add") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    isAddPrint = 1
    b8TitleBar2.Caption = "Search"
    Label4.Caption = "Payroll Date"
    picAdd.ZOrder 0
    txtSearchAdd.Text = ""
    txtPayrollDateAdd.Text = ""
    picMain.Enabled = False
    picToolbar.Enabled = False
    picAdd.Visible = True
    txtSearchAdd.SetFocus
Else
    If isFocus = 0 Then Exit Sub
    With lstDetails.ListItems
        If CDbl(.Item(.Count).Text) <> 0 Then
            Set x = .Add()
            x.Text = "0"
            x.SubItems(1) = " "
            x.SubItems(2) = " "
            x.SubItems(3) = "-1"
            x.SubItems(4) = " "
            x.SubItems(5) = " "
            x.SubItems(6) = " "
        Else
            .Item(.Count).Text = "0"
            .Item(.Count).SubItems(1) = " "
            .Item(.Count).SubItems(2) = " "
            .Item(.Count).SubItems(3) = "-1"
            .Item(.Count).SubItems(4) = " "
            .Item(.Count).SubItems(5) = " "
            .Item(.Count).SubItems(6) = " "
        End If
        iRow = .Count
    End With
    lstDetails.ListItems(iRow).EnsureVisible
    lstDetails.ListItems(iRow).Selected = True
    picSLLines.ZOrder 0
    txtDeductionKey.Text = ""
    cmbDeductionName.Text = ""
    cmbDeductionName.ListIndex = -1
    txtDeductionPeriodKey = "-1"
    cmbDeductionPeriod.Text = ""
    cmbDeductionPeriod.ListIndex = -1
    txtAmount.Text = ""
    txtPerPayroll.Text = ""
    txtSLRemarks.Text = ""
    picMain.Enabled = False
    picToolbar.Enabled = False
    picSLLines.Visible = True
    TRANS_DETAIL = is_DET_ADDING
    cmbDeductionName.SetFocus
End If
End Sub

Private Sub PRESS_F2()
If picPrintSumm.Visible = True Then Exit Sub
If picSLLines.Visible = True Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    If Trim(Statusbar1.Panels(1).Text) = "" Then Exit Sub
    If imgPosted.Visible = True Then MsgBox "ALREADY POSTED!                     ", vbCritical, "Error...": Exit Sub
    If AccessRights("Personnel Deduction", "Edit") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_EDITTING
    If isFocus = 1 Then lstDetails_Click
Else
    If isFocus = 0 Then Exit Sub
    With lstDetails.ListItems
        txtDeductionKey.Text = .Item(iRow).Text
        cmbDeductionName.Text = .Item(iRow).SubItems(1)
        txtAmount.Text = .Item(iRow).SubItems(2)
        txtDeductionPeriodKey.Text = .Item(iRow).SubItems(3)
        cmbDeductionPeriod.Text = .Item(iRow).SubItems(4)
        txtPerPayroll.Text = .Item(iRow).SubItems(5)
        txtSLRemarks.Text = .Item(iRow).SubItems(6)
        
        txtDeductionKey1.Text = .Item(iRow).Text
        cmbDeductionName1.Text = .Item(iRow).SubItems(1)
        txtAmount1.Text = .Item(iRow).SubItems(2)
        txtDeductionPeriodKey1.Text = .Item(iRow).SubItems(3)
        cmbDeductionPeriod1.Text = .Item(iRow).SubItems(4)
        txtPerPayroll1.Text = .Item(iRow).SubItems(5)
        txtSLRemarks1.Text = .Item(iRow).SubItems(6)
    End With
    picSLLines.ZOrder 0
    picMain.Enabled = False
    picToolbar.Enabled = False
    picSLLines.Visible = True
    TRANS_DETAIL = is_DET_EDITTING
    cmbDeductionName.SetFocus
End If
End Sub

Private Sub PRESS_DELETE()
If picPrintSumm.Visible = True Then Exit Sub
If picSLLines.Visible = True Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    If Trim(Statusbar1.Panels(1).Text) = "" Then Exit Sub
    If imgPosted.Visible = True Then MsgBox "ALREADY POSTED!                     ", vbCritical, "Error...": Exit Sub
    If AccessRights("Personnel Deduction", "Delete") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If MsgBox("ARE YOU SURE IN DELETING THIS TRANSACTION?                   ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    On Error GoTo PG:
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Deduction WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    CLEARTEXT
    BROWSER GetSetting(App.EXEName, "PersonnelDeductionCtrl", "PersonnelDeductionCtrl", ""), "is_PAGEDOWN"
    If Trim(txtControl.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelDeductionCtrl", "PersonnelDeductionCtrl", ""), "is_HOME"
Else
    If isFocus = 0 Then Exit Sub
    With lstDetails.ListItems
        If .Count > 1 Then
            .Remove iRow
            If CDbl(iRow) > CDbl(.Count) Then iRow = .Count
        Else
            .Item(1).Text = "0"
            .Item(1).SubItems(1) = " "
            .Item(1).SubItems(2) = " "
            .Item(1).SubItems(3) = "-1"
            .Item(1).SubItems(4) = " "
            .Item(1).SubItems(5) = " "
            .Item(1).SubItems(6) = " "
            iRow = 1
        End If
    End With
    lstDetails.ListItems(iRow).EnsureVisible
    lstDetails.ListItems(iRow).Selected = True
End If
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If picPrintSumm.Visible = True Then Exit Sub
If picSLLines.Visible = True Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If IsDate(txtDate.Text) = False Then MsgBox "Please supply a valid date!                    ", vbCritical, "Error...": txtDate.SetFocus: Exit Sub
locPayrollPeroid = GET_PERIOD_V2(FormatDateTime(txtDate.Text, vbShortDate), GET_DIVISION_V2(locEmployeePK, FormatDateTime(txtDate.Text, vbShortDate)))
'If locPayrollPeroid = 0 Then MsgBox "Please supply a valid payroll period!                    ", vbCritical, "Error...": txtDate.SetFocus: Exit Sub
If locPayrollPeroid = 0 Then
    MsgBox "Payroll Period Not Match to the Employee Division!      ", vbInformation, ""
    txtDate.SetFocus
    HTEXT txtDate
    Exit Sub
End If
'If RETURNTEXTVALUE(txtTotalAmount) <= 0 Then MsgBox "Please supply a value higher than zero!                 ", vbCritical, "Error...": txtTotalAmount.SetFocus: Exit Sub
'If RETURNTEXTVALUE(txtDedPerPayroll) <= 0 Then MsgBox "Please supply a value higher than zero!                 ", vbCritical, "Error...": txtDedPerPayroll.SetFocus: Exit Sub
With lstDetails.ListItems
    For i = 1 To .Count
        If CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2))) <= 0 Then MsgBox "Invalid Amount!                    ", vbCritical, "Error...": lstDetails.ListItems(i).EnsureVisible: lstDetails.ListItems(i).Selected = True: Exit Sub
        If CDbl(IIf(IsNumeric(.Item(i).SubItems(5)) = False, 0, .Item(i).SubItems(5))) <= 0 Then MsgBox "Invalid Amount!                    ", vbCritical, "Error...": lstDetails.ListItems(i).EnsureVisible: lstDetails.ListItems(i).Selected = True: Exit Sub
    Next i
End With
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    sCtrl = ""
    s = "SELECT TOP (1) dbo.tbl_Personnel_Deduction.Ctrl " & _
        " FROM  dbo.tbl_Personnel_Deduction LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Deduction.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
        " Where (Year(dbo.tbl_Personnel_Compensation_Period.PayrollDate) = " & Format(FormatDateTime(txtDate.Text), "yyyy") & ") " & _
        " ORDER BY dbo.tbl_Personnel_Deduction.Ctrl DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        sCtrl = Format(CDbl(rs!Ctrl) + 1, "000000000#")
    Else
        sCtrl = Format(FormatDateTime(txtDate.Text), "yyyy") & "000000"
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
        sCtrl = Format(CDbl(sCtrl) + 1, "000000000#")
    Loop
                      
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Deduction " & _
                      " (Ctrl, EmployeeKey, PayrollPeriodKey, LastModified) " & _
                      " VALUES ('" & sCtrl & "', " & locEmployeePK & ", " & locPayrollPeroid & ", " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
    
    iPK = 0
    s = "SELECT tbl_Personnel_Deduction.* " & _
        " FROM tbl_Personnel_Deduction " & _
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
    ConnOmega.Execute "UPDATE tbl_Personnel_Deduction " & _
                      " SET PayrollPeriodKey = " & locPayrollPeroid & ", " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
End If

If CDbl(iPK) <> 0 Then
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Deduction_Details WHERE (MasterKey = " & iPK & ")"
    With lstDetails.ListItems
        For i = 1 To .Count
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Deduction_Details " & _
                              " (MasterKey, DeductionKey, Amount, DeductionPeriodKey, DedPerPayroll, Remarks) " & _
                              " VALUES (" & iPK & ", " & .Item(i).Text & ", " & _
                              " " & CDbl(.Item(i).SubItems(2)) & ", " & CDbl(.Item(i).SubItems(3)) & ", " & _
                              " " & CDbl(.Item(i).SubItems(5)) & ", '" & FORMATSQL(.Item(i).SubItems(6)) & "')"
        Next i
    End With
End If
txtDate.SetFocus
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANS_DETAIL = is_DET_REFRESH
BROWSER sCtrl, "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If picPrintSumm.Visible = True Then Exit Sub
If picSLLines.Visible = True Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Trim(Statusbar1.Panels(1).Text) = "" Then Exit Sub
picSearch.ZOrder 0
txtSearchSearch.Text = ""
picMain.Enabled = False
picToolbar.Enabled = False
picSearch.Visible = True
txtSearchSearch.SetFocus
End Sub

Private Sub PRESS_F8()
If picPrintSumm.Visible = True Then Exit Sub
If picSLLines.Visible = True Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Trim(Statusbar1.Panels(1).Text) = "" Then Exit Sub
On Error GoTo PG:
If imgPosted.Visible = False Then
    If AccessRights("Personnel Deduction", "Post") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If MsgBox("ARE YOU SURE IN POSTING THIS TRANSACTION?                   ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    
    s = "SELECT tbl_Personnel_Deduction_Details.* " & _
        " FROM tbl_Personnel_Deduction_Details " & _
        " WHERE (MasterKey = " & Statusbar1.Panels(1).Text & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Deduction_SL " & _
                          " (SourceKey, EmployeeKey, DeductionKey, TransactionDate, Remarks, InOut, Debit, TransactionType) " & _
                          " VALUES (" & Statusbar1.Panels(1).Text & ", " & locEmployeePK & ", " & rs!DeductionKey & ", " & _
                          " '" & FormatDateTime(txtDate.Text, vbShortDate) & "', '" & FORMATSQL(rs!Remarks) & "', " & _
                          " 'I', " & CDbl(rs!Amount) & ", 1)"
        rs.MoveNext
    Wend
    rs.Close
    
    ConnOmega.Execute "UPDATE tbl_Personnel_Deduction " & _
                      " SET Posted = 1 " & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
End If
If imgPosted.Visible = True Then
    If AccessRights("Personnel Deduction", "UnPost") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    
    t = "SELECT dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.SourceKey " & _
        " FROM  dbo.tbl_Personnel_Deduction_forPayroll_Det_Det LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Deduction_forPayroll_Det ON dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.MasterKey = dbo.tbl_Personnel_Deduction_forPayroll_Det.MasterKey AND dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.EmployeeKey = dbo.tbl_Personnel_Deduction_forPayroll_Det.EmployeeKey LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Deduction_forPayroll ON dbo.tbl_Personnel_Deduction_forPayroll_Det.MasterKey = dbo.tbl_Personnel_Deduction_forPayroll.PK " & _
        " WHERE (dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.SourceKey = " & Statusbar1.Panels(1).Text & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        MsgBox "Can't Unpost this transaction!              " & vbCrLf & _
               "Already created in for deduction module!                ", vbCritical, "Error..."
        rt.Close
        Exit Sub
    End If
    rt.Close
    
    If MsgBox("ARE YOU SURE IN UNPOSTING THIS TRANSACTION?                   ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Deduction_SL WHERE (SourceKey = " & Statusbar1.Panels(1).Text & ")"
    
    ConnOmega.Execute "UPDATE tbl_Personnel_Deduction " & _
                      " SET Posted = 0 " & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    
End If

CLEARTEXT
BROWSER GetSetting(App.EXEName, "PersonnelDeductionCtrl", "PersonnelDeductionCtrl", ""), "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F9()
If picPrintSumm.Visible = True Then Exit Sub
If picSLLines.Visible = True Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Trim(Statusbar1.Panels(1).Text) = "" Then Exit Sub

PopupMenu MainFormPopupF.mnuPayrollDeductionReport, , Toolbar1.Buttons(17).Left, Toolbar1.Buttons(17).Top + Toolbar1.Buttons(17).Height

'isAddPrint = 2
'b8TitleBar2.Caption = "Employee Active Deduction Balance"
'Label4.Caption = "as of"
'picAdd.ZOrder 0
'txtSearchAdd.Text = ""
'txtPayrollDateAdd.Text = Format(Date, "mm/dd/yyyy")
'picMain.Enabled = False
'picToolbar.Enabled = False
'picAdd.Visible = True
'txtSearchAdd.SetFocus
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picPrintSumm.Visible = True Then cmdCancelSumm_Click: Exit Sub
    If picAdd.Visible = True Then cmdCancelAdd_Click: Exit Sub
    If picSearch.Visible = True Then cmdCancelSearch_Click: Exit Sub
    Unload Me
Else
    If picSLLines.Visible = True Then
        If TRANS_DETAIL = is_DET_ADDING Then
            With lstDetails.ListItems
                If .Count > 1 Then
                    .Remove .Count
                Else
                    .Item(1).Text = "0"
                    .Item(1).SubItems(1) = " "
                    .Item(1).SubItems(2) = " "
                    .Item(1).SubItems(3) = "-1"
                    .Item(1).SubItems(4) = " "
                    .Item(1).SubItems(5) = " "
                    .Item(1).SubItems(6) = " "
                End If
                iRow = .Count
            End With
        End If
        If TRANS_DETAIL = is_DET_EDITTING Then
            With lstDetails.ListItems
                .Item(iRow).Text = txtDeductionKey1.Text
                .Item(iRow).SubItems(1) = cmbDeductionName1.Text
                .Item(iRow).SubItems(2) = txtAmount1.Text
                .Item(iRow).SubItems(3) = txtDeductionPeriodKey1.Text
                .Item(iRow).SubItems(4) = cmbDeductionPeriod1.Text
                .Item(iRow).SubItems(5) = txtPerPayroll1.Text
                .Item(iRow).SubItems(6) = txtSLRemarks1.Text
            End With
        End If
        picSLLines.Visible = False
        picMain.Enabled = True
        picToolbar.Enabled = True
        lstDetails.ListItems(iRow).EnsureVisible
        lstDetails.ListItems(iRow).Selected = True
        lstDetails.SetFocus
        Exit Sub
    End If
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "PersonnelDeductionCtrl", "PersonnelDeductionCtrl", ""), "is_LOAD"
    If Trim(txtControl.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelDeductionCtrl", "PersonnelDeductionCtrl", ""), "is_HOME"
End If
End Sub

Private Sub BROWSER(Ctrl, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If Ctrl <> "" Then
            s = "SELECT TOP (1) dbo.tbl_Personnel_Deduction.PK, dbo.tbl_Personnel_Deduction.Ctrl, dbo.tbl_Personnel_Deduction.EmployeeKey, " & _
                " dbo.tbl_Personnel_Deduction.PayrollPeriodKey, dbo.tbl_Personnel_Deduction.Posted, dbo.tbl_Personnel_Deduction.LastModified, " & _
                " dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
                " dbo.tbl_Personnel_Information.MiddleName , dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo, " & _
                " dbo.tbl_Personnel_Compensation_Period.PayrollDate  " & _
                " FROM  dbo.tbl_Personnel_Deduction LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Deduction.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Deduction.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
                " WHERE (dbo.tbl_Personnel_Deduction.Ctrl = '" & Ctrl & "') " & _
                " ORDER BY dbo.tbl_Personnel_Deduction.Ctrl"
        Else
            s = "SELECT TOP (1) dbo.tbl_Personnel_Deduction.PK, dbo.tbl_Personnel_Deduction.Ctrl, dbo.tbl_Personnel_Deduction.EmployeeKey, " & _
                " dbo.tbl_Personnel_Deduction.PayrollPeriodKey, dbo.tbl_Personnel_Deduction.Posted, dbo.tbl_Personnel_Deduction.LastModified, " & _
                " dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
                " dbo.tbl_Personnel_Information.MiddleName , dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo, " & _
                " dbo.tbl_Personnel_Compensation_Period.PayrollDate  " & _
                " FROM  dbo.tbl_Personnel_Deduction LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Deduction.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Deduction.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
                " ORDER BY dbo.tbl_Personnel_Deduction.Ctrl"
        End If
    Case "is_HOME"
        If picPrintSumm.Visible = True Then Exit Sub
        If picSLLines.Visible = True Then Exit Sub
        If picAdd.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Deduction.PK, dbo.tbl_Personnel_Deduction.Ctrl, dbo.tbl_Personnel_Deduction.EmployeeKey, " & _
            " dbo.tbl_Personnel_Deduction.PayrollPeriodKey, dbo.tbl_Personnel_Deduction.Posted, dbo.tbl_Personnel_Deduction.LastModified, " & _
            " dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
            " dbo.tbl_Personnel_Information.MiddleName , dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo, " & _
            " dbo.tbl_Personnel_Compensation_Period.PayrollDate  " & _
            " FROM  dbo.tbl_Personnel_Deduction LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Deduction.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Deduction.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
            " ORDER BY dbo.tbl_Personnel_Deduction.Ctrl"
    Case "is_PAGEUP"
        If picPrintSumm.Visible = True Then Exit Sub
        If picSLLines.Visible = True Then Exit Sub
        If picAdd.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Deduction.PK, dbo.tbl_Personnel_Deduction.Ctrl, dbo.tbl_Personnel_Deduction.EmployeeKey, " & _
            " dbo.tbl_Personnel_Deduction.PayrollPeriodKey, dbo.tbl_Personnel_Deduction.Posted, dbo.tbl_Personnel_Deduction.LastModified, " & _
            " dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
            " dbo.tbl_Personnel_Information.MiddleName , dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo, " & _
            " dbo.tbl_Personnel_Compensation_Period.PayrollDate  " & _
            " FROM  dbo.tbl_Personnel_Deduction LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Deduction.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Deduction.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
            " WHERE (dbo.tbl_Personnel_Deduction.Ctrl < '" & Ctrl & "') " & _
            " ORDER BY dbo.tbl_Personnel_Deduction.Ctrl DESC"
    Case "is_PAGEDOWN"
        If picPrintSumm.Visible = True Then Exit Sub
        If picSLLines.Visible = True Then Exit Sub
        If picAdd.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Deduction.PK, dbo.tbl_Personnel_Deduction.Ctrl, dbo.tbl_Personnel_Deduction.EmployeeKey, " & _
            " dbo.tbl_Personnel_Deduction.PayrollPeriodKey, dbo.tbl_Personnel_Deduction.Posted, dbo.tbl_Personnel_Deduction.LastModified, " & _
            " dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
            " dbo.tbl_Personnel_Information.MiddleName , dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo, " & _
            " dbo.tbl_Personnel_Compensation_Period.PayrollDate  " & _
            " FROM  dbo.tbl_Personnel_Deduction LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Deduction.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Deduction.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
            " WHERE (dbo.tbl_Personnel_Deduction.Ctrl > '" & Ctrl & "') " & _
            " ORDER BY dbo.tbl_Personnel_Deduction.Ctrl"
    Case "is_END"
        If picPrintSumm.Visible = True Then Exit Sub
        If picSLLines.Visible = True Then Exit Sub
        If picAdd.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Deduction.PK, dbo.tbl_Personnel_Deduction.Ctrl, dbo.tbl_Personnel_Deduction.EmployeeKey, " & _
            " dbo.tbl_Personnel_Deduction.PayrollPeriodKey, dbo.tbl_Personnel_Deduction.Posted, dbo.tbl_Personnel_Deduction.LastModified, " & _
            " dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
            " dbo.tbl_Personnel_Information.MiddleName , dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo, " & _
            " dbo.tbl_Personnel_Compensation_Period.PayrollDate  " & _
            " FROM  dbo.tbl_Personnel_Deduction LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Deduction.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Deduction.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
            " ORDER BY dbo.tbl_Personnel_Deduction.Ctrl DESC"
    Case "is_FIND"
        s = "SELECT TOP (1) dbo.tbl_Personnel_Deduction.PK, dbo.tbl_Personnel_Deduction.Ctrl, dbo.tbl_Personnel_Deduction.EmployeeKey, " & _
            " dbo.tbl_Personnel_Deduction.PayrollPeriodKey, dbo.tbl_Personnel_Deduction.Posted, dbo.tbl_Personnel_Deduction.LastModified, " & _
            " dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
            " dbo.tbl_Personnel_Information.MiddleName , dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo, " & _
            " dbo.tbl_Personnel_Compensation_Period.PayrollDate  " & _
            " FROM  dbo.tbl_Personnel_Deduction LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Deduction.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Deduction.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
            " WHERE (dbo.tbl_Personnel_Deduction.PK = " & Ctrl & ") " & _
            " ORDER BY dbo.tbl_Personnel_Deduction.Ctrl DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    locEmployeePK = rs!EmployeeKey
    locPayrollPeroid = rs!PayrollPeriodKey
    txtControl.Text = rs!Ctrl
    txtName.Text = rs!IDNumber & " - " & rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
    txtDate.Text = Format(rs!PayrollDate, "mm/dd/yyyy")
    txtCutOffDate.Text = Format(rs!DateFrom, "mm/dd/yyyy") & " - " & Format(rs!DateTo, "mm/dd/yyyy")
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    
    imgPosted.Visible = IIf(rs!Posted = 1, True, False)
    Toolbar1.Buttons(19).Caption = IIf(rs!Posted = 1, "UnPost", " Post ")
    Toolbar1.Buttons(19).Image = IIf(rs!Posted = 1, 11, 10)
    
    dTotalAmt = 0: dTotalPerPayroll = 0: CLEAR_Details
    't = "SELECT dbo.tbl_Personnel_Deduction_Details.DeductionKey, dbo.tbl_Personnel_Payroll_Deductions_Table.Description, " & _
        " dbo.tbl_Personnel_Deduction_Details.Amount, dbo.tbl_Personnel_Deduction_Details.DedPerPayroll, " & _
        " dbo.tbl_Personnel_Deduction_Details.Remarks " & _
        " FROM  dbo.tbl_Personnel_Deduction_Details LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Deduction_Details.DeductionKey = dbo.tbl_Personnel_Payroll_Deductions_Table.PK " & _
        " Where (dbo.tbl_Personnel_Deduction_Details.MasterKey = " & rs!PK & ") " & _
        " ORDER BY dbo.tbl_Personnel_Payroll_Deductions_Table.Sorting"
    t = "SELECT dbo.tbl_Personnel_Deduction_Details.DeductionKey, dbo.tbl_Personnel_Payroll_Deductions_Table.Description, " & _
        " dbo.tbl_Personnel_Deduction_Details.Amount, dbo.tbl_Personnel_Deduction_Details.DeductionPeriodKey, " & _
        " dbo.tbl_Personnel_Deduction_Period.DeductionPeriod, dbo.tbl_Personnel_Deduction_Details.DedPerPayroll, " & _
        " dbo.tbl_Personnel_Deduction_Details.Remarks " & _
        " FROM  dbo.tbl_Personnel_Deduction_Details LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Deduction_Period ON dbo.tbl_Personnel_Deduction_Details.DeductionPeriodKey = dbo.tbl_Personnel_Deduction_Period.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Deduction_Details.DeductionKey = dbo.tbl_Personnel_Payroll_Deductions_Table.PK " & _
        " Where (dbo.tbl_Personnel_Deduction_Details.MasterKey = " & rs!PK & ") " & _
        " ORDER BY dbo.tbl_Personnel_Payroll_Deductions_Table.Sorting"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstDetails.ListItems.Clear
        While Not rt.EOF
            dTotalAmt = dTotalAmt + CDbl(rt!Amount)
            dTotalPerPayroll = dTotalPerPayroll + CDbl(rt!DedPerPayroll)
            Set x = lstDetails.ListItems.Add()
            x.Text = rt!DeductionKey
            x.SubItems(1) = rt!Description
            x.SubItems(2) = Format(rt!Amount, "#,##0.00")
            x.SubItems(3) = rt!DeductionPeriodKey
            x.SubItems(4) = rt!DeductionPeriod
            x.SubItems(5) = Format(rt!DedPerPayroll, "#,##0.00")
            x.SubItems(6) = rt!Remarks
            rt.MoveNext
        Wend
    End If
    rt.Close
    
    lblTotalAmount.Caption = Format(dTotalAmt, "#,##0.00")
    lblPerPayroll.Caption = Format(dTotalPerPayroll, "#,##0.00")

    
    SaveSetting App.EXEName, "PersonnelDeductionCtrl", "PersonnelDeductionCtrl", rs!Ctrl
End If
rs.Close
End Sub

Private Sub CLEARTEXT()
locEmployeePK = 0
locPayrollPeroid = 0
txtControl.Text = ""
txtName.Text = ""
txtDate.Text = ""
txtCutOffDate.Text = ""
lblTotalAmount.Caption = "0.00"
lblPerPayroll.Caption = "0.00"
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
imgPosted.Visible = False
CLEAR_Details
End Sub

Private Sub CLEAR_Details()
With lstDetails.ListItems
    .Clear
    Set x = .Add()
    x.Text = "0"
    x.SubItems(1) = " "
    x.SubItems(2) = " "
    x.SubItems(3) = "-1"
    x.SubItems(4) = " "
    x.SubItems(5) = " "
    x.SubItems(6) = " "
End With
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtName.Locked = True
txtCutOffDate.Locked = True
txtDate.Locked = bln
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
cmdCancelSumm_Click
End Sub

Private Sub cmbDeductionName_Click()
If cmbDeductionName.ListIndex = -1 Then Exit Sub
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    txtDeductionKey.Text = cmbDeductionName.ItemData(cmbDeductionName.ListIndex)
    With lstDetails.ListItems
        .Item(iRow).SubItems(1) = cmbDeductionName.List(cmbDeductionName.ListIndex)
    End With
End If
End Sub

Private Sub cmbDeductionName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtAmount.SetFocus
End Sub



Private Sub cmbDeductionPeriod_Click()
If cmbDeductionPeriod.ListIndex = -1 Then Exit Sub
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    txtDeductionPeriodKey.Text = cmbDeductionPeriod.ItemData(cmbDeductionPeriod.ListIndex)
    With lstDetails.ListItems
        .Item(iRow).SubItems(4) = cmbDeductionPeriod.List(cmbDeductionPeriod.ListIndex)
    End With
End If
End Sub

Private Sub cmbDeductionPeriod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtPerPayroll.SetFocus
End Sub

Private Sub cmbDivision_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtAsOf.SetFocus
End Sub

Private Sub cmbResultSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKSearch_Click
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

Private Sub cmdCancelSumm_Click()
picPrintSumm.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True
End Sub

Private Sub cmdOKAdd_Click()
If lstResultAdd.ListIndex = -1 Then Exit Sub
If IsDate(txtPayrollDateAdd.Text) = False Then MsgBox "Please supply a valid date!                      ", vbCritical, "Error...": txtPayrollDateAdd.SetFocus: Exit Sub
If isAddPrint = 1 Then
    iEmpStatus = GET_EMPLOYMENT_STATUS(lstResultAdd.ItemData(lstResultAdd.ListIndex), FormatDateTime(txtPayrollDateAdd.Text, vbShortDate))
    If iEmpStatus = 0 Or iEmpStatus = 2 Then
        Array1 = Split(lstResultAdd.List(lstResultAdd.ListIndex), " - ", -1, 1)
        If MsgBox(Array1(1) & " WAS ALREADY INACTIVE!                       ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    End If
    locPayrollPeroidTmp = GET_PERIOD_V2(FormatDateTime(txtPayrollDateAdd.Text, vbShortDate), GET_DIVISION_V2(lstResultAdd.ItemData(lstResultAdd.ListIndex), FormatDateTime(txtPayrollDateAdd.Text, vbShortDate)))
    If locPayrollPeroidTmp = 0 Then
        MsgBox "Payroll Period Not Match to the Employee Division!      ", vbInformation, ""
        txtPayrollDateAdd.SetFocus
        HTEXT txtPayrollDateAdd
        Exit Sub
    End If
    CLEARTEXT
    LOCKTEXT False
    TOOLBARFUNC 2
    locEmployeePK = lstResultAdd.ItemData(lstResultAdd.ListIndex)
    locPayrollPeroid = locPayrollPeroidTmp
    txtName.Text = lstResultAdd.List(lstResultAdd.ListIndex)
    txtDate.Text = Format(FormatDateTime(txtPayrollDateAdd.Text, vbShortDate), "mm/dd/yyyy")
    txtCutOffDate.Text = GET_PERIOD_CUTOFF(locPayrollPeroid)
    TRANSACTIONTYPE = is_ADDING
    cmdCancelAdd_Click
    txtDate.SetFocus
ElseIf isAddPrint = 2 Then
    TimerOutStanding.Enabled = True
End If
End Sub

Private Sub cmdOKSearch_Click()
If cmbResultSearch.ListIndex = -1 Then Exit Sub
BROWSER cmbResultSearch.ItemData(cmbResultSearch.ListIndex), "is_FIND"
cmdCancelSearch_Click
End Sub

Private Sub cmdOKSumm_Click()
If cmbDivision.ListIndex = -1 Then MsgBox "Please select division!                   ", vbCritical, "Error...": cmbDivision.SetFocus: Exit Sub
If IsDate(txtAsOf.Text) = False Then MsgBox "Please supply a valid date!                    ", vbCritical, "Error...": txtAsOf.SetFocus: Exit Sub

cmdCancelSumm_Click
End Sub

Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
s = "SELECT ROUND(SUM(Balance), 2) AS Balance, EmployeeKey, DeductionKey " & _
    " From dbo.tbl_Personnel_Deduction_SL " & _
    " GROUP BY EmployeeKey, DeductionKey " & _
    " HAVING (ROUND(SUM(Balance), 2) < 0)"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    
    t = "SELECT TOP (1) dbo.tbl_Personnel_Deduction_SL.SourceKey, dbo.tbl_Personnel_Compensation_Period.PayrollDate " & _
        " FROM  dbo.tbl_Personnel_Deduction_SL LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Deduction ON dbo.tbl_Personnel_Deduction_SL.SourceKey = dbo.tbl_Personnel_Deduction.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Deduction.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
        " Where (dbo.tbl_Personnel_Deduction_SL.EmployeeKey = " & rs!EmployeeKey & ") " & _
        " And (dbo.tbl_Personnel_Deduction_SL.DeductionKey = " & rs!DeductionKey & ") " & _
        " ORDER BY TransactionDate"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Deduction_SL " & _
                          " (SourceKey, EmployeeKey, DeductionKey, TransactionDate, Remarks, InOut, Debit) " & _
                          " VALUES (" & rt!SourceKey & ", " & rs!EmployeeKey & ", " & rs!DeductionKey & ", " & _
                          " '" & rt!PayrollDate & "', 'Adjusment - System', " & _
                          " 'I', " & CDbl(rs!Balance) * -1 & ")"
    End If
    rt.Close
    
    'ConnOmega.Execute "INSERT INTO tbl_Personnel_Deduction_SL " & _
                          " (SourceKey, EmployeeKey, DeductionKey, TransactionDate, Remarks, InOut, Debit) " & _
                          " VALUES (" & Statusbar1.Panels(1).Text & ", " & locEmployeePK & ", " & rs!DeductionKey & ", " & _
                          " '" & FormatDateTime(txtDate.Text, vbShortDate) & "', '" & FORMATSQL(rs!Remarks) & "', " & _
                          " 'I', " & CDbl(rs!Amount) & ")"
    
'    ConnOmega.Execute "INSERT INTO tbl_Personnel_Deduction_SL " & _
                      " () " & _
                      " VALUES ()"
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
    Case vbKeyF9:       PRESS_F9
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "PersonnelDeductionCtrl", "PersonnelDeductionCtrl", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "PersonnelDeductionCtrl", "PersonnelDeductionCtrl", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "PersonnelDeductionCtrl", "PersonnelDeductionCtrl", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "PersonnelDeductionCtrl", "PersonnelDeductionCtrl", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 3
Me.Left = (MainForm.ScaleWidth - Me.Width) / 3
POPULATE_COMBO_EXEMPTION "PK", "Description", "tbl_Personnel_Payroll_Deductions_Table", "Sorting", "ViewInDeductionModule", 1, cmbDeductionName
POPULATE_COMBO "PK", "DeductionPeriod", "tbl_Personnel_Deduction_Period", "PK", cmbDeductionPeriod
POPULATE_COMBO "PK", "Description", "tbl_Personnel_Division", "Description", cmbDivision
isFocus = 0
iRow = 0
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANS_DETAIL = is_DET_REFRESH
BROWSER GetSetting(App.EXEName, "PersonnelDeductionCtrl", "PersonnelDeductionCtrl", ""), "is_LOAD"
If Trim(txtControl.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelDeductionCtrl", "PersonnelDeductionCtrl", ""), "is_HOME"
tmp = SetWindowLong(txtSearchAdd.hwnd, GWL_STYLE, GetWindowLong(txtSearchAdd.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearchSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearchSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
'tmp = SetWindowLong(txtRemarks.hwnd, GWL_STYLE, GetWindowLong(txtRemarks.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picPrintSumm.Visible = True Then Cancel = -1
If picSLLines.Visible = True Then Cancel = -1
If picAdd.Visible = True Then Cancel = -1
If picSearch.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstDetails_Click()
iRow = lstDetails.SelectedItem.Index
isFocus = 1
TRANS_DETAIL = is_DET_REFRESH
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If CDbl(lstDetails.ListItems.Item(iRow).Text) <> 0 Then TOOLBARFUNC 5 Else TOOLBARFUNC 4
End If
End Sub

Private Sub lstDetails_GotFocus()
iRow = lstDetails.SelectedItem.Index
isFocus = 1
TRANS_DETAIL = is_DET_REFRESH
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If CDbl(lstDetails.ListItems.Item(iRow).Text) <> 0 Then TOOLBARFUNC 5 Else TOOLBARFUNC 4
End If
End Sub

Private Sub lstDetails_ItemClick(ByVal Item As MSComctlLib.ListItem)
iRow = lstDetails.SelectedItem.Index
End Sub

Private Sub lstDetails_LostFocus()
isFocus = 0
End Sub

Private Sub lstResultAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtPayrollDateAdd.SetFocus
End Sub

Private Sub lstResultSearch_Click()
If lstResultSearch.ListIndex = -1 Then cmbResultSearch.Clear: Exit Sub
cmbResultSearch.Clear
't = "SELECT dbo.tbl_Personnel_Deduction.PK, dbo.tbl_Personnel_Deduction.TransDate, " & _
    " dbo.tbl_Personnel_Payroll_Deductions_Table.Description, dbo.tbl_Personnel_Deduction.Remarks " & _
    " FROM  dbo.tbl_Personnel_Deduction LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Deduction.DeductionKey = dbo.tbl_Personnel_Payroll_Deductions_Table.PK " & _
    " Where (dbo.tbl_Personnel_Deduction.EmployeeKey = " & lstResultSearch.ItemData(lstResultSearch.ListIndex) & ") " & _
    " ORDER BY dbo.tbl_Personnel_Deduction.TransDate DESC"
t = "SELECT dbo.tbl_Personnel_Deduction.PK, dbo.tbl_Personnel_Compensation_Period.PayrollDate, dbo.tbl_Personnel_Deduction.EmployeeKey " & _
    " FROM  dbo.tbl_Personnel_Deduction LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Deduction.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
    " Where (dbo.tbl_Personnel_Deduction.EmployeeKey = " & lstResultSearch.ItemData(lstResultSearch.ListIndex) & ") " & _
    " ORDER BY dbo.tbl_Personnel_Compensation_Period.PayrollDate DESC"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
While Not rt.EOF
    cmbResultSearch.AddItem Format(rt!PayrollDate, "mm/dd/yyyy")
    cmbResultSearch.ItemData(cmbResultSearch.NewIndex) = rt!PK
    rt.MoveNext
Wend
rt.Close
If cmbResultSearch.ListCount Then cmbResultSearch.ListIndex = 0
End Sub

Private Sub lstResultSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbResultSearch.SetFocus
End Sub

Private Sub TimerOutStanding_Timer()
TimerOutStanding.Enabled = False
picAdd.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True
Screen.MousePointer = vbHourglass
s = "SELECT dbo.tbl_Personnel_Deduction_SL.DeductionKey, " & _
    " dbo.tbl_Personnel_Payroll_Deductions_Table.Description, " & _
    " ROUND(SUM(dbo.tbl_Personnel_Deduction_SL.Balance), 2) AS Balance " & _
    " FROM  dbo.tbl_Personnel_Deduction_SL LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Deduction_SL.DeductionKey = dbo.tbl_Personnel_Payroll_Deductions_Table.PK " & _
    " Where (dbo.tbl_Personnel_Deduction_SL.EmployeeKey = " & lstResultAdd.ItemData(lstResultAdd.ListIndex) & ") " & _
    " AND (dbo.tbl_Personnel_Deduction_SL.TransactionDate <= '" & FormatDateTime(txtPayrollDateAdd.Text, vbShortDate) & "') " & _
    " GROUP BY dbo.tbl_Personnel_Deduction_SL.DeductionKey, dbo.tbl_Personnel_Payroll_Deductions_Table.Description, dbo.tbl_Personnel_Payroll_Deductions_Table.Sorting " & _
    " Having (ROUND(SUM(dbo.tbl_Personnel_Deduction_SL.Balance),2) > 0) " & _
    " ORDER BY dbo.tbl_Personnel_Payroll_Deductions_Table.Sorting"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount = 0 Then
    Screen.MousePointer = vbDefault
    MsgBox "No outstanding balance!                         ", vbInformation, "Info"
    rs.Close
    Exit Sub
Else
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Deduction_Report WHERE (LogIn = '" & gbl_UserName & "')"
    Array2 = Split(lstResultAdd.List(lstResultAdd.ListIndex), " - ", -1, 1)
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Deduction_Report " & _
                      " (LogIn, EmployeeID, EmployeeName, AsOf, CompanyKey) " & _
                      " VALUES ('" & gbl_UserName & "', '" & FORMATSQL(CStr(Array2(0))) & "','" & FORMATSQL(CStr(Array2(1))) & "', " & _
                      " '" & Format(FormatDateTime(txtPayrollDateAdd.Text, vbShortDate), "mmmm dd, yyyy") & "', 1)"
    
    iMasterKey = 0: iLine = 0: dRunBal = 0: iDedKey = 0: strDedKeySourceKey = ""
    t = "SELECT tbl_Personnel_Deduction_Report.* " & _
        " FROM tbl_Personnel_Deduction_Report " & _
        " WHERE (LogIn = '" & gbl_UserName & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        iMasterKey = rt!PK
    End If
    rt.Close
    
    While Not rs.EOF
        u = "SELECT dbo.tbl_Personnel_Deduction_SL.EmployeeKey, dbo.tbl_Personnel_Deduction_SL.DeductionKey, " & _
            " dbo.tbl_Personnel_Deduction_SL.SourceKey, dbo.tbl_Personnel_Payroll_Deductions_Table.Description, " & _
            " ROUND(SUM(dbo.tbl_Personnel_Deduction_SL.Balance), 2) AS Balance " & _
            " FROM  dbo.tbl_Personnel_Deduction_SL LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Deduction_SL.DeductionKey = dbo.tbl_Personnel_Payroll_Deductions_Table.PK " & _
            " WHERE (dbo.tbl_Personnel_Deduction_SL.TransactionDate <= '" & FormatDateTime(txtPayrollDateAdd.Text, vbShortDate) & "') " & _
            " GROUP BY dbo.tbl_Personnel_Deduction_SL.EmployeeKey, dbo.tbl_Personnel_Deduction_SL.DeductionKey, dbo.tbl_Personnel_Deduction_SL.SourceKey, dbo.tbl_Personnel_Payroll_Deductions_Table.Description " & _
            " HAVING (dbo.tbl_Personnel_Deduction_SL.EmployeeKey = " & lstResultAdd.ItemData(lstResultAdd.ListIndex) & ") " & _
            " AND (dbo.tbl_Personnel_Deduction_SL.DeductionKey = " & rs!DeductionKey & ") " & _
            " AND (ROUND(SUM(dbo.tbl_Personnel_Deduction_SL.Balance), 2) > 0)"
        If ru.State = adStateOpen Then ru.Close
        ru.Open u, ConnOmega
        While Not ru.EOF
            t = "SELECT dbo.tbl_Personnel_Deduction_SL.EmployeeKey, dbo.tbl_Personnel_Deduction_SL.DeductionKey, " & _
                " dbo.tbl_Personnel_Payroll_Deductions_Table.Description, dbo.tbl_Personnel_Deduction_SL.TransactionDate, " & _
                " dbo.tbl_Personnel_Deduction_SL.Remarks, dbo.tbl_Personnel_Deduction_SL.InOut, dbo.tbl_Personnel_Deduction_SL.Debit, " & _
                " dbo.tbl_Personnel_Deduction_SL.Credit, dbo.tbl_Personnel_Deduction_SL.Balance " & _
                " FROM  dbo.tbl_Personnel_Deduction_SL LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Deduction_SL.DeductionKey = dbo.tbl_Personnel_Payroll_Deductions_Table.PK " & _
                " WHERE (dbo.tbl_Personnel_Deduction_SL.EmployeeKey = " & lstResultAdd.ItemData(lstResultAdd.ListIndex) & ") " & _
                " AND (dbo.tbl_Personnel_Deduction_SL.DeductionKey = " & rs!DeductionKey & ") " & _
                " AND (dbo.tbl_Personnel_Deduction_SL.SourceKey = " & ru!SourceKey & ") " & _
                " AND (dbo.tbl_Personnel_Deduction_SL.TransactionDate <= '" & FormatDateTime(txtPayrollDateAdd.Text, vbShortDate) & "') " & _
                " ORDER BY dbo.tbl_Personnel_Deduction_SL.TransactionDate"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            While Not rt.EOF
                iLine = iLine + 1
'                If CDbl(iDedKey) <> CDbl(rt!DeductionKey) Then
'                    dRunBal = 0
'                End If
'                iDedKey = rt!DeductionKey
                If Trim(CStr(strDedKeySourceKey)) <> CStr(rt!DeductionKey) & "-" & CStr(ru!SourceKey) Then
                    dRunBal = 0
                End If
                strDedKeySourceKey = CStr(rt!DeductionKey) & "-" & CStr(ru!SourceKey)
                dRunBal = dRunBal + CDbl(rt!Balance)
                ConnOmega.Execute "INSERT INTO tbl_Personnel_Deduction_Report_Det " & _
                                  " (MasterKey, Line, AccountName, TransDate, Remarks, Debit, Credit, RunBal, AccountNameSourceKey) " & _
                                  " VALUES (" & iMasterKey & ", " & iLine & ", '" & FORMATSQL(rt!Description) & "', " & _
                                  " '" & FormatDateTime(rt!TransactionDate, vbShortDate) & "', " & _
                                  " '" & FORMATSQL(rt!Remarks) & "', " & CDbl(rt!Debit) & ", " & CDbl(rt!Credit) & ", " & _
                                  " " & CDbl(dRunBal) & ", '" & FORMATSQL(rt!Description) & "-" & ru!SourceKey & "')"
                rt.MoveNext
            Wend
            rt.Close
            ru.MoveNext
        Wend
        ru.Close
        rs.MoveNext
    Wend
    rs.Close
    Screen.MousePointer = vbDefault
    frmCrystalReportViewer.PRINT_Employee_OutStandingBal gbl_UserName
    If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
End If
'Screen.MousePointer = vbDefault
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "PersonnelDeductionCtrl", "PersonnelDeductionCtrl", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "PersonnelDeductionCtrl", "PersonnelDeductionCtrl", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "PersonnelDeductionCtrl", "PersonnelDeductionCtrl", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "PersonnelDeductionCtrl", "PersonnelDeductionCtrl", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Post":    PRESS_F8
    Case "Print":   PRESS_F9
    Case "Refresh":
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtAmount_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    dTotalAmt = 0
    With lstDetails.ListItems
        .Item(iRow).SubItems(2) = Format(RETURNTEXTVALUE(txtAmount), "#,##0.00")
        For i = 1 To .Count
            dTotalAmt = dTotalAmt + CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2)))
        Next i
    End With
    lblTotalAmount.Caption = Format(dTotalAmt, "#,##0.00")
End If
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbDeductionPeriod.SetFocus
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtDate_GotFocus()
HTEXT txtDate
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then lstDetails.SetFocus
End Sub

Private Sub txtDate_LostFocus()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If IsDate(txtDate.Text) Then txtDate.Text = Format(FormatDateTime(txtDate.Text, vbShortDate), "mm/dd/yyyy") Else Exit Sub
    locPayrollPeroidTmp = GET_PERIOD_V2(FormatDateTime(txtDate.Text, vbShortDate), GET_DIVISION_V2(locEmployeePK, FormatDateTime(txtDate.Text, vbShortDate)))
    If locPayrollPeroidTmp = 0 Then
        MsgBox "Payroll Period Not Match to the Employee Division!      ", vbInformation, ""
        txtPayrollDateAdd.SetFocus
        HTEXT txtPayrollDateAdd
        Exit Sub
    End If
    txtCutOffDate.Text = GET_PERIOD_CUTOFF(locPayrollPeroidTmp)
End If
End Sub

Private Sub txtDeductionKey_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetails.ListItems
        .Item(iRow).Text = txtDeductionKey.Text
    End With
End If
End Sub

Private Sub txtDeductionPeriodKey_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetails.ListItems
        .Item(iRow).SubItems(3) = txtDeductionPeriodKey.Text
    End With
End If
End Sub

Private Sub txtPayrollDateAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAdd_Click
End Sub

Private Sub txtPerPayroll_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    dTotalPerPayroll = 0
    With lstDetails.ListItems
        .Item(iRow).SubItems(5) = Format(RETURNTEXTVALUE(txtPerPayroll), "#,##0.00")
        For i = 1 To .Count
            dTotalPerPayroll = dTotalPerPayroll + CDbl(IIf(IsNumeric(.Item(i).SubItems(5)) = False, 0, .Item(i).SubItems(5)))
        Next i
    End With
    lblPerPayroll.Caption = Format(dTotalPerPayroll, "#,##0.00")
End If
End Sub

Private Sub txtPerPayroll_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtSLRemarks.SetFocus
End Sub

Private Sub txtPerPayroll_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtSearchAdd_Change()
If Trim(txtSearchAdd.Text) = "" Then lstResultAdd.Clear: Exit Sub
lstResultAdd.Clear
If isAddPrint = 1 Then
    s = "SELECT dbo.tbl_Personnel_IDNumber.PK, dbo.tbl_Personnel_IDNumber.IDNumber, " & _
        " dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
        " dbo.tbl_Personnel_Information.MiddleName " & _
        " FROM  dbo.tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
        " WHERE (ISNULL((SELECT TOP (1) tbl_Personnel_EmploymentStatus_1.Active " & _
        " FROM  dbo.tbl_Personnel_ActionNew AS tbl_Personnel_ActionNew_1 LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_EmploymentStatus AS tbl_Personnel_EmploymentStatus_1 ON tbl_Personnel_ActionNew_1.EmpStatusKey = tbl_Personnel_EmploymentStatus_1.PK " & _
        " WHERE (tbl_Personnel_ActionNew_1.EmpPK = dbo.tbl_Personnel_IDNumber.PK) " & _
        " AND (tbl_Personnel_ActionNew_1.EffectivityDate <= '" & FormatDateTime(Date, vbShortDate) & "') " & _
        " ORDER BY tbl_Personnel_ActionNew_1.EffectivityDate DESC), 0) = 1) " & _
        " AND (dbo.tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%') " & _
        " ORDER BY dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName"
ElseIf isAddPrint = 2 Then
    s = "SELECT dbo.tbl_Personnel_Deduction.EmployeeKey as PK, dbo.tbl_Personnel_IDNumber.IDNumber, " & _
        " dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
        " dbo.tbl_Personnel_Information.MiddleName " & _
        " FROM  dbo.tbl_Personnel_Deduction LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Deduction.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
        " GROUP BY dbo.tbl_Personnel_Deduction.EmployeeKey, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName " & _
        " HAVING (dbo.tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%') " & _
        " ORDER BY dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_IDNumber.IDNumber"
End If
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultAdd.AddItem rs!IDNumber & " - " & rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
    lstResultAdd.ItemData(lstResultAdd.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResultAdd.ListCount Then lstResultAdd.ListIndex = 0
End Sub

Private Sub txtSearchAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultAdd.SetFocus
End Sub

Private Sub txtSearchSearch_Change()
If Trim(txtSearchSearch.Text) = "" Then lstResultSearch.Clear: cmbResultSearch.Clear:  Exit Sub
lstResultSearch.Clear: cmbResultSearch.Clear
s = "SELECT dbo.tbl_Personnel_Deduction.EmployeeKey, dbo.tbl_Personnel_IDNumber.IDNumber, " & _
    " dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
    " dbo.tbl_Personnel_Information.MiddleName " & _
    " FROM  dbo.tbl_Personnel_Deduction LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Deduction.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
    " GROUP BY dbo.tbl_Personnel_Deduction.EmployeeKey, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName " & _
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

Private Sub txtSLRemarks_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetails.ListItems
        .Item(iRow).SubItems(6) = txtSLRemarks.Text
    End With
End If
End Sub

Private Sub txtSLRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If RETURNTEXTVALUE(txtDeductionKey) = 0 Then MsgBox "Please select deduction name!              ", vbCritical, "Error...": cmbDeductionName.SetFocus: Exit Sub
    If RETURNTEXTVALUE(txtDeductionPeriodKey) = -1 Then MsgBox "Please select deduction period!              ", vbCritical, "Error...": cmbDeductionPeriod.SetFocus: Exit Sub
    picSLLines.Visible = False
    picMain.Enabled = True
    picToolbar.Enabled = True
    lstDetails.SetFocus
End If
End Sub
