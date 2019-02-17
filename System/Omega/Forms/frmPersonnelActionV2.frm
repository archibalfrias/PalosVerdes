VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelActionV2 
   Appearance      =   0  'Flat
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9300
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
   ScaleHeight     =   5970
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picSLRates 
      Height          =   855
      Left            =   5280
      TabIndex        =   53
      Top             =   2040
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox txtRate1 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   69
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox cmbEarningDesc1 
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   68
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtEarningKey1 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   67
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtEarningKey 
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   66
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2400
         TabIndex        =   56
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cmbEarningDesc 
         Height          =   315
         Left            =   120
         TabIndex        =   54
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         TabIndex        =   57
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   120
         Width           =   1095
      End
   End
   Begin RPVGCC.b8Container picSearch 
      Height          =   4335
      Left            =   3000
      TabIndex        =   58
      Top             =   1080
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   7646
      BackColor       =   15266266
      Begin VB.TextBox txtSearchSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   63
         Top             =   480
         Width           =   3375
      End
      Begin VB.ListBox lstResultSearch 
         Height          =   2400
         Left            =   120
         TabIndex        =   62
         Top             =   885
         Width           =   3375
      End
      Begin VB.CommandButton cmdCancelSearch 
         Height          =   480
         Left            =   1920
         Picture         =   "frmPersonnelActionV2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   3720
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKSearch 
         Height          =   480
         Left            =   120
         Picture         =   "frmPersonnelActionV2.frx":075C
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   3720
         Width           =   1560
      End
      Begin VB.ComboBox cmbEffectivityDate 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   3360
         Width           =   2055
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   40
         TabIndex        =   64
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
         Icon            =   "frmPersonnelActionV2.frx":0DCE
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Effectivity Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   3360
         Width           =   1335
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   360
      ScaleHeight     =   3975
      ScaleWidth      =   8535
      TabIndex        =   3
      Top             =   1320
      Width           =   8535
      Begin VB.TextBox txtEffectDate 
         Height          =   315
         Left            =   1680
         TabIndex        =   43
         Top             =   3300
         Width           =   3255
      End
      Begin VB.TextBox txtControl 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSComctlLib.ListView lstRates 
         Height          =   1335
         Left            =   5040
         TabIndex        =   39
         Top             =   1650
         Width           =   3465
         _ExtentX        =   6112
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "MasterKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "EarningKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   3475
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Rates"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.TextBox txtRemarks 
         Height          =   315
         Left            =   1680
         TabIndex        =   37
         Top             =   3630
         Width           =   6825
      End
      Begin VB.ComboBox cmbLoanDeduction 
         Height          =   315
         Left            =   1680
         TabIndex        =   34
         Top             =   2640
         Width           =   3255
      End
      Begin VB.ComboBox cmbGovtDeduction 
         Height          =   315
         ItemData        =   "frmPersonnelActionV2.frx":1368
         Left            =   1680
         List            =   "frmPersonnelActionV2.frx":136A
         TabIndex        =   33
         Top             =   2970
         Width           =   3255
      End
      Begin VB.ComboBox cmbTaxCategory 
         Height          =   315
         ItemData        =   "frmPersonnelActionV2.frx":136C
         Left            =   1680
         List            =   "frmPersonnelActionV2.frx":136E
         TabIndex        =   31
         Top             =   2310
         Width           =   3255
      End
      Begin VB.PictureBox picGovt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C6B8A4&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   5040
         ScaleHeight     =   1335
         ScaleWidth      =   4455
         TabIndex        =   18
         Top             =   330
         Width           =   4455
         Begin VB.TextBox txtSSS 
            Height          =   315
            Left            =   1200
            TabIndex        =   26
            Top             =   0
            Width           =   1935
         End
         Begin VB.TextBox txtPHIC 
            Height          =   315
            Left            =   1200
            TabIndex        =   25
            Top             =   330
            Width           =   1935
         End
         Begin VB.TextBox txtPagIbig 
            Height          =   315
            Left            =   1200
            TabIndex        =   24
            Top             =   660
            Width           =   1935
         End
         Begin VB.TextBox txtTIN 
            Height          =   315
            Left            =   1200
            TabIndex        =   23
            Top             =   990
            Width           =   1935
         End
         Begin VB.CheckBox chkSSS 
            BackColor       =   &H00FDE9C6&
            Height          =   200
            Left            =   3240
            TabIndex        =   22
            Top             =   40
            Width           =   200
         End
         Begin VB.CheckBox chkPHIC 
            BackColor       =   &H00FDE9C6&
            Height          =   200
            Left            =   3240
            TabIndex        =   21
            Top             =   390
            Width           =   200
         End
         Begin VB.CheckBox chkPagIbig 
            BackColor       =   &H00FDE9C6&
            Height          =   200
            Left            =   3240
            TabIndex        =   20
            Top             =   720
            Width           =   200
         End
         Begin VB.CheckBox chkTIN 
            BackColor       =   &H00FDE9C6&
            Height          =   200
            Left            =   3240
            TabIndex        =   19
            Top             =   1080
            Width           =   200
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "SSS NO."
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   30
            Top             =   45
            Width           =   1095
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "PHIL HEALTH "
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   29
            Top             =   390
            Width           =   1095
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "PAG IBIG"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   28
            Top             =   735
            Width           =   1095
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "TIN"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   27
            Top             =   1065
            Width           =   1095
         End
      End
      Begin VB.ComboBox cmdDept 
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Top             =   660
         Width           =   3255
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Top             =   0
         Width           =   6825
      End
      Begin VB.ComboBox cmdStatus 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   990
         Width           =   3255
      End
      Begin VB.ComboBox cmbPost 
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   1320
         Width           =   3255
      End
      Begin VB.ComboBox cmbComp 
         Height          =   315
         ItemData        =   "frmPersonnelActionV2.frx":1370
         Left            =   1680
         List            =   "frmPersonnelActionV2.frx":1372
         TabIndex        =   6
         Top             =   1650
         Width           =   3255
      End
      Begin VB.ComboBox cmbTaxStatus 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   1980
         Width           =   3255
      End
      Begin VB.ComboBox cmbDivision 
         Height          =   315
         ItemData        =   "frmPersonnelActionV2.frx":1374
         Left            =   1680
         List            =   "frmPersonnelActionV2.frx":1376
         TabIndex        =   4
         Top             =   330
         Width           =   3255
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "EFFECTIVITY DATE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   44
         Top             =   3330
         Width           =   1575
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
         Left            =   6480
         TabIndex        =   41
         Top             =   3030
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL >>"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5400
         TabIndex        =   40
         Top             =   3030
         Width           =   975
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "REMARKS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   38
         Top             =   3660
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "LOAN DEDUCTION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   36
         Top             =   2670
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "GOVT DEDUCTION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   35
         Top             =   3010
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "TAX CATEGORY"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   32
         Top             =   2340
         Width           =   1695
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "POSITION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   1350
         Width           =   1095
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "DEPARTMENT"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   705
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "EMP STATUS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   1050
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NAME"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   30
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "RATE COMPENSATION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "TAX STATUS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   2010
         Width           =   1095
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "DIVISION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   375
         Width           =   1095
      End
   End
   Begin RPVGCC.b8Container picAdd 
      Height          =   4335
      Left            =   3000
      TabIndex        =   45
      Top             =   1200
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   7646
      BackColor       =   15266266
      Begin VB.TextBox txtEffectDateAdd 
         Height          =   315
         Left            =   1800
         TabIndex        =   51
         Top             =   3360
         Width           =   1695
      End
      Begin VB.CommandButton cmdOKAdd 
         Height          =   480
         Left            =   120
         Picture         =   "frmPersonnelActionV2.frx":1378
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   3720
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelAdd 
         Height          =   480
         Left            =   1920
         Picture         =   "frmPersonnelActionV2.frx":19EA
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   3720
         Width           =   1560
      End
      Begin VB.ListBox lstResultAdd 
         Height          =   2400
         Left            =   120
         TabIndex        =   47
         Top             =   885
         Width           =   3375
      End
      Begin VB.TextBox txtSearchAdd 
         Height          =   315
         Left            =   120
         TabIndex        =   46
         Top             =   480
         Width           =   3375
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   40
         TabIndex        =   50
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
         Icon            =   "frmPersonnelActionV2.frx":2146
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Effectivity Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   3435
         Width           =   1575
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   1
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   2
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
               Caption         =   "Refresh"
               Key             =   "Refresh"
               ImageKey        =   "IMG12"
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
               ImageKey        =   "IMG13"
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         MousePointer    =   99
         MouseIcon       =   "frmPersonnelActionV2.frx":26E0
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
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   5655
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11959
            MinWidth        =   11959
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10320
      Top             =   1320
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
            Picture         =   "frmPersonnelActionV2.frx":29FA
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelActionV2.frx":36D4
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelActionV2.frx":43AE
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelActionV2.frx":5088
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelActionV2.frx":5D62
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelActionV2.frx":6A3C
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelActionV2.frx":7716
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelActionV2.frx":83F0
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelActionV2.frx":90CA
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelActionV2.frx":99A4
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelActionV2.frx":A67E
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelActionV2.frx":B358
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelActionV2.frx":C032
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelActionV2.frx":CD0C
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelActionV2.frx":D9E6
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPersonnelActionV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public locEmployeePK As Long
Public locTransNo As Long

Dim locDiv As Long
Dim locDivTmp As Long
Dim locDept As Long
Dim locPost As Long
Dim locEmpStatus As Long
Dim locTaxStatus As Long
Dim iSupervisory As Long
Dim locTaxCat As Long
Dim locLoanDed As Long
Dim locGovtDed As Long
Dim locCompKey As Long

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

Dim isFocus, iRow As Long


Dim Array1, x, sCtrl, iPK, i, dblTotalAmt, dblRatePerHour

Private Sub PRESS_INSERT()
If picAdd.Visible = True Then Exit Sub
If picSLRates.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    If AccessRights("Personnel Action Memo", "Add") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    picAdd.ZOrder 0
    txtSearchAdd.Text = ""
    picAdd.Visible = True
    picMain.Enabled = False
    picToolbar.Enabled = False
    txtSearchAdd.SetFocus
Else
    If isFocus = 0 Then Exit Sub
    With lstRates.ListItems
        If CDbl(.Item(.Count).SubItems(1)) <> 0 Then
            Set x = .Add()
            x.Text = "0"
            x.SubItems(1) = "0"
            x.SubItems(2) = " "
            x.SubItems(3) = " "
        Else
            .Item(.Count).SubItems(2) = " "
            .Item(.Count).SubItems(3) = " "
        End If
        iRow = .Count
    End With
    lstRates.ListItems(iRow).EnsureVisible
    lstRates.ListItems(iRow).Selected = True
    picSLRates.ZOrder 0
    cmbEarningDesc.Text = ""
    cmbEarningDesc.ListIndex = -1
    txtRate.Text = ""
    picMain.Enabled = False
    picToolbar.Enabled = False
    picSLRates.Visible = True
    TRANS_DETAIL = is_DET_ADDING
    cmbEarningDesc.SetFocus
End If
End Sub

Private Sub PRESS_F2()
If picAdd.Visible = True Then Exit Sub
If picSLRates.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    If AccessRights("Personnel Action Memo", "Edit") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If Trim(StatusBar.Panels(3).Text) = "LOCKED" Then MsgBox "Action Memo already locked!                 ", vbCritical, "Error...": Exit Sub
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_EDITTING
Else
    If isFocus = 0 Then Exit Sub
    With lstRates.ListItems
        txtEarningKey.Text = .Item(iRow).SubItems(1)
        cmbEarningDesc.Text = .Item(iRow).SubItems(2)
        txtRate.Text = .Item(iRow).SubItems(3)
        txtEarningKey1.Text = .Item(iRow).SubItems(1)
        cmbEarningDesc1.Text = .Item(iRow).SubItems(2)
        txtRate1.Text = .Item(iRow).SubItems(3)
    End With
    lstRates.ListItems(iRow).EnsureVisible
    lstRates.ListItems(iRow).Selected = True
    picSLRates.ZOrder 0
    picMain.Enabled = False
    picToolbar.Enabled = False
    picSLRates.Visible = True
    TRANS_DETAIL = is_DET_EDITTING
    cmbEarningDesc.SetFocus
End If
End Sub

Private Sub PRESS_DELETE()
If picAdd.Visible = True Then Exit Sub
If picSLRates.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    If AccessRights("Personnel Action Memo", "Delete") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    
    If Trim(StatusBar.Panels(3).Text) = "LOCKED" Then MsgBox "Action Memo already locked!                 ", vbCritical, "Error...": Exit Sub
    
    s = "SELECT tbl_Personnel_Hours.* " & _
        " From tbl_Personnel_Hours " & _
        " WHERE (ActionMemoKey = " & StatusBar.Panels(1).Text & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        MsgBox "CANNOT BE DELETED!          " & vbCrLf & _
               "" & vbCrLf & _
               "Action used by the Payroll Transaction.   " & vbCrLf & _
               "Any changes have an effect on Payroll Computation.  ", vbCritical, "Error..."
        Exit Sub
    End If
    rs.Close
    
    s = "SELECT tbl_Personnel_Payroll.* " & _
        " From tbl_Personnel_Payroll " & _
        " WHERE (ActionMemoKey = " & StatusBar.Panels(1).Text & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        MsgBox "CANNOT BE DELETED!          " & vbCrLf & _
               "" & vbCrLf & _
               "Action used by the Payroll Transaction.   " & vbCrLf & _
               "Any changes have an effect on Payroll Computation.  ", vbCritical, "Error..."
        Exit Sub
    End If
    rs.Close
    If MsgBox("ARE YOU SURE TO DELETE THIS RECORD?          ", vbCritical + vbYesNo, "Confirm") = vbNo Then Exit Sub
    ConnOmega.Execute "DELETE FROM tbl_Personnel_ActionNew WHERE (PK = " & StatusBar.Panels(1).Text & ")"
    CLEARTEXT
    BROWSER GetSetting(App.EXEName, "PersonnelActionCtrlV2", "PerActCtrlV2", ""), "is_PAGEDOWN"
    If Trim(txtControl.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelActionCtrlV2", "PerActCtrlV2", ""), "is_HOME"
    
    
'    If MsgBox("ARE YOU SURE IN DELETING THIS TRANSACTION?                   ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
'    On Error GoTo PG:
    
Else
    If isFocus = 0 Then Exit Sub
    With lstRates.ListItems
        If .Count > 1 Then
            .Remove iRow
            If CDbl(iRow) > CDbl(.Count) Then iRow = .Count
        Else
            .Item(1).SubItems(1) = "0"
            .Item(1).SubItems(2) = " "
            .Item(1).SubItems(3) = " "
            iRow = 1
        End If
        dblTotalAmt = 0
        For i = 1 To .Count
            dblTotalAmt = dblTotalAmt + CDbl(IIf(IsNumeric(.Item(iRow).SubItems(3)) = False, 0, .Item(iRow).SubItems(3)))
        Next i
    End With
    lstRates.ListItems(iRow).EnsureVisible
    lstRates.ListItems(iRow).Selected = True
    lblTotal.Caption = Format(dblTotalAmt, "#,##0.00")
End If
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If picAdd.Visible = True Then Exit Sub
If picSLRates.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If locDiv = 0 Then MsgBox "Please Select Division!              ", vbCritical, "Error...": cmbDivision.SetFocus: Exit Sub
If locDept = 0 Then MsgBox "Please Supply Department!                   ", vbCritical, "Error...": cmdDept.SetFocus: Exit Sub
If locPost = 0 Then MsgBox "Please Supply Position!                 ", vbCritical, "Error...": cmbPost.SetFocus: Exit Sub
If locEmpStatus = 0 Then MsgBox "Please Supply Employment Status!               ", vbCritical, "Error...": cmdStatus.SetFocus: Exit Sub
If locTaxStatus = 0 Then MsgBox "Please Supply Tax Status!                  ", vbCritical, "Error...": cmbTaxStatus.SetFocus: Exit Sub
If locCompKey = 0 Then MsgBox "Please Supply Compensation Rate!                 ", vbCritical, "Error...": cmbComp.SetFocus: Exit Sub
If locTaxCat = 0 Then MsgBox "Please Supply Tax Category!               ", vbCritical, "Error...": cmbTaxCategory.SetFocus: Exit Sub
If locLoanDed = 0 Then MsgBox "Please Supply Loan Deduction Schedule!                  ", vbCritical, "Error...": cmbLoanDeduction.SetFocus: Exit Sub
If locGovtDed = 0 Then MsgBox "Please Supply Government Deduction Schedule!                 ", vbCritical, "Error...": cmbGovtDeduction.SetFocus: Exit Sub
If IsDate(txtEffectDate.Text) = False Then MsgBox "Please supply a valid date!                    ", vbCritical, "Error...": txtEffectDate.SetFocus: Exit Sub
If RETURNLABELVALUE(lblTotal) <= 0 Then MsgBox "Invalid rate value!                  ", vbCritical, "Error...": Exit Sub
If iSupervisory = 2 Then
    If AccessRights("Personnel Action Memo", "Supervisory") = False Then
        MsgBox "YOU DON'T HAVE ACCESS ON SUPERVISORY LEVEL POSITION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
End If

If TRANSACTIONTYPE = is_ADDING Then
    
    If DateValue(GET_LAST_ACTION_EFFECTIVITY_NEW(locEmployeePK)) > DateValue(CDate(Trim(txtEffectDate.Text))) Then
        MsgBox "EFFECTIVITY DATE MUST BE HIGHER THAN THE LAST ACTION MEMO!          ", vbInformation, "Error..."
        Exit Sub
    End If

    sCtrl = ""
    u = "SELECT TOP (1) CntrlNo " & _
        " FROM tbl_Personnel_ActionNew " & _
        " WHERE (Year(EffectivityDate) = " & Format(FormatDateTime(txtEffectDate.Text, vbShortDate), "yyyy") & ") " & _
        " ORDER BY CntrlNo DESC"
    If ru.State = adStateOpen Then ru.Close
    ru.Open u, ConnOmega
    If ru.RecordCount = 0 Then
        sCtrl = Format(FormatDateTime(txtEffectDate.Text, vbShortDate), "yyyy") & "0000"
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
    
    ConnOmega.Execute "INSERT INTO tbl_Personnel_ActionNew" & _
                      " (CntrlNo, EmpPK, DivisionKey, DeptKey, EmpStatusKey, TaxStatusKey," & _
                      " PositionsKey, CompensationRateKey, Is_PAGIBIG, Is_PHIC, Is_SSS, " & _
                      " Is_TIN, SSS, PAGIBIG, PHIC, TIN, Remarks, EffectivityDate, " & _
                      " LastModified, TaxCategoryKey, LoanDeductionKey, GovtDeductionKey)" & _
                      " VALUES ('" & sCtrl & "', " & locEmployeePK & ", " & locDiv & ", " & _
                      " " & locDept & ", " & locEmpStatus & ", " & locTaxStatus & ", " & locPost & ", " & _
                      " " & locCompKey & ", " & chkPagIbig.Value & ", " & chkPHIC.Value & ", " & chkSSS.Value & ", " & _
                      " " & chkTIN.Value & ", '" & Trim(txtSSS.Text) & "', '" & Trim(txtPagIbig.Text) & "', '" & Trim(txtPHIC.Text) & "', " & _
                      " '" & Trim(txtTIN.Text) & "', '" & FORMATSQL(Trim(txtRemarks.Text)) & "', '" & FormatDateTime(txtEffectDate.Text, vbShortDate) & "', " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "', " & locTaxCat & ", " & locLoanDed & ", " & locGovtDed & ")"
    
    iPK = 0
    s = "SELECT PK " & _
        " FROM tbl_Personnel_ActionNew " & _
        " WHERE (CntrlNo = '" & sCtrl & "') "
    If ru.State = adStateOpen Then ru.Close
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        iPK = rs!PK
    End If
    rs.Close
    
End If
If TRANSACTIONTYPE = is_EDITTING Then
    
    If CDbl(locDivTmp) <> CDbl(locDiv) Then
        
    End If
    
    sCtrl = txtControl.Text
    iPK = StatusBar.Panels(1).Text
    
    If locDivTmp <> locDiv Then
        t = "SELECT dbo.tbl_Personnel_Hours.ActionMemoKey, dbo.tbl_Personnel_Compensation_Period.PayrollDate " & _
            " FROM  dbo.tbl_Personnel_Hours LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Hours.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
            " WHERE (dbo.tbl_Personnel_Hours.ActionMemoKey = " & iPK & ")"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            MsgBox "Division can't be changed it has already linked to personnel hours under " & Format(rt!PayrollDate, "mm/dd/yyyy") & " payroll date!              ", vbExclamation, "Can't Edit"
            rt.Close
            Exit Sub
        End If
        rt.Close
    End If
    
    ConnOmega.Execute "UPDATE tbl_Personnel_ActionNew" & _
                      " SET DivisionKey = " & locDiv & ", DeptKey = " & locDept & ", EmpStatusKey = " & locEmpStatus & ", " & _
                      " TaxStatusKey = " & locTaxStatus & ", PositionsKey = " & locPost & ", CompensationRateKey = " & locCompKey & ", " & _
                      " Is_PAGIBIG = " & chkPagIbig.Value & ", Is_PHIC = " & chkPHIC.Value & ", Is_SSS = " & chkSSS.Value & ", " & _
                      " Is_TIN = " & chkTIN.Value & ", SSS = '" & Trim(txtSSS.Text) & "', PAGIBIG = '" & Trim(txtPagIbig.Text) & "', " & _
                      " PHIC = '" & Trim(txtPHIC.Text) & "', TIN = '" & Trim(txtTIN.Text) & "', Remarks = '" & FORMATSQL(Trim(txtRemarks.Text)) & "', " & _
                      " EffectivityDate = '" & FormatDateTime(txtEffectDate.Text, vbShortDate) & "', LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                      " TaxCategoryKey = " & locTaxCat & ", LoanDeductionKey = " & locLoanDed & ", GovtDeductionKey = " & locGovtDed & "" & _
                      " WHERE (PK = " & iPK & ")"
    
End If

If CDbl(iPK) > 0 Then
    t = "SELECT ProfileKey " & _
        " FROM tbl_Personnel_IDNumber " & _
        " WHERE (PK = " & locEmployeePK & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        ConnOmega.Execute "UPDATE tbl_Personnel_Information " & _
                          " SET SSSNumber = '" & Trim(txtSSS.Text) & "', " & _
                          " PHICNumber = '" & Trim(txtPHIC.Text) & "', " & _
                          " HDMFNumber = '" & Trim(txtPagIbig.Text) & "', " & _
                          " TIN = '" & Trim(txtTIN.Text) & "', " & _
                          " TaxStatus = " & locTaxStatus & " " & _
                          " WHERE (PK = " & rt!ProfileKey & ")"
    End If
    rt.Close
    ConnOmega.Execute "DELETE FROM tbl_Personnel_ActionNew_Rate WHERE (MasterKey = " & iPK & ")"
    With lstRates.ListItems
        For i = 1 To .Count
            If CDbl(.Item(i).SubItems(1)) <> 0 Then
                
                dblRatePerHour = 0
                If CDbl(.Item(i).SubItems(3)) > 0 Then
                    If locCompKey = 3 Then
                        dblRatePerHour = ((CDbl(.Item(i).SubItems(3)) / 2) / gbl_MpnthlyDivisor) / 8
                    Else
                        dblRatePerHour = CDbl(.Item(i).SubItems(3)) / 8
                    End If
                End If
            
                ConnOmega.Execute "INSERT INTO tbl_Personnel_ActionNew_Rate " & _
                                  " (MasterKey, EarningKey, Rate, RatePerHour) " & _
                                  " VALUES (" & iPK & ", " & CDbl(.Item(i).SubItems(1)) & ", " & _
                                  " " & CDbl(.Item(i).SubItems(3)) & ", " & CDbl(dblRatePerHour) & ")"
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
End Sub

Private Sub PRESS_F6()
If picAdd.Visible = True Then Exit Sub
If picSLRates.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
picSearch.ZOrder 0
txtSearchSearch.Text = ""
picMain.Enabled = False
picToolbar.Enabled = False
picSearch.Visible = True
txtSearchSearch.SetFocus
End Sub

Private Sub PRESS_F8()
If picAdd.Visible = True Then Exit Sub
If picSLRates.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
End Sub

Private Sub PRESS_F9()
If picAdd.Visible = True Then Exit Sub
If picSLRates.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If StatusBar.Panels(1).Text = "" Then Exit Sub

Dim iPKfrom, iPKto, iLine, iLine2
iPKto = StatusBar.Panels(1).Text
ConnOmega.Execute "DELETE FROM tbl_Personnel_ActionNew_Report WHERE (LogIn = '" & gbl_UserName & "')"
s = "SELECT dbo.tbl_Personnel_ActionNew.CntrlNo, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, " & _
    " dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_IDNumber.DateHired, " & _
    " dbo.tbl_Personnel_ActionNew.EffectivityDate, dbo.tbl_Personnel_Information.SSSNumber, dbo.tbl_Personnel_Information.PHICNumber, " & _
    " dbo.tbl_Personnel_Information.HDMFNumber, dbo.tbl_Personnel_Information.TIN, dbo.tbl_Personnel_ActionNew.Remarks, " & _
    " dbo.tbl_Personnel_ActionNew.EmpPK " & _
    " FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_ActionNew.EmpPK = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_ActionNew.DivisionKey = dbo.tbl_Personnel_Division.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_EmploymentStatus ON dbo.tbl_Personnel_ActionNew.EmpStatusKey = dbo.tbl_Personnel_EmploymentStatus.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Govt_TaxStatus ON dbo.tbl_Personnel_ActionNew.TaxStatusKey = dbo.tbl_Govt_TaxStatus.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Position_Level ON dbo.tbl_Personnel_Position.PositionLevel = dbo.tbl_Personnel_Position_Level.PK " & _
    " WHERE (dbo.tbl_Personnel_ActionNew.PK = " & iPKto & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    ConnOmega.Execute "INSERT INTO tbl_Personnel_ActionNew_Report " & _
                      " (LogIn, Ctrl, IDNumber, EmployeeName, DateHired, EffectDate, SSSNum, PHICNum, PagIbigNum, TIN, Company, Remarks) " & _
                      " VALUES ('" & gbl_UserName & "', '" & rs!CntrlNo & "', '" & rs!IDNumber & "', '" & FORMATSQL(rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName) & "', " & _
                      " '" & FormatDateTime(rs!DateHired, vbShortDate) & "', '" & FormatDateTime(rs!EffectivityDate, vbShortDate) & "', " & _
                      " '" & FORMATSQL(rs!SSSNumber) & "', '" & FORMATSQL(rs!PHICNumber) & "', '" & FORMATSQL(rs!HDMFNumber) & "', " & _
                      " '" & FORMATSQL(rs!TIN) & "', 1, '" & FORMATSQL(rs!Remarks) & "')"
    
    iPK = 0
    t = "SELECT PK " & _
        " FROM tbl_Personnel_ActionNew_Report " & _
        " WHERE (LogIn = '" & gbl_UserName & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        iPK = rt!PK
    End If
    rt.Close
    
    iPKfrom = 0:
    t = "SELECT dbo.tbl_Personnel_Division.Description AS Division, dbo.tbl_Personnel_Department.DepartmentName AS Department, " & _
        " dbo.tbl_Personnel_Position.PositionName AS Position, dbo.tbl_Personnel_Position_Level.LevelName AS [Level], " & _
        " dbo.tbl_Personnel_EmploymentStatus.StatusName AS [Employment Status], dbo.tbl_Personnel_CompensationRate.Description as [Basis of Salary], " & _
        " dbo.tbl_Govt_TaxStatus.TaxStatus AS [Tax Status]" & _
        " FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
        " dbo.tbl_Govt_TaxStatus ON dbo.tbl_Personnel_ActionNew.TaxStatusKey = dbo.tbl_Govt_TaxStatus.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_ActionNew.EmpPK = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_ActionNew.DivisionKey = dbo.tbl_Personnel_Division.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_EmploymentStatus ON dbo.tbl_Personnel_ActionNew.EmpStatusKey = dbo.tbl_Personnel_EmploymentStatus.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Position_Level ON dbo.tbl_Personnel_Position.PositionLevel = dbo.tbl_Personnel_Position_Level.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_CompensationRate ON dbo.tbl_Personnel_ActionNew.CompensationRateKey = dbo.tbl_Personnel_CompensationRate.PK " & _
        " WHERE (dbo.tbl_Personnel_ActionNew.PK = " & iPKto & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        iLine = 0
        For i = 1 To rt.Fields.Count
            iLine = iLine + 1
            ConnOmega.Execute "INSERT INTO tbl_Personnel_ActionNew_Report_Det " & _
                              " (MasterKey, Line, NatureOfChange, ValueTo, LeftRightPos) " & _
                              " VALUES (" & iPK & ", " & iLine & ", " & _
                              " '" & rt.Fields(i - 1).Name & "', " & _
                              " '" & FORMATSQL(CStr(rt.Fields(i - 1).Value)) & "', 1)"
        Next i
    End If
    rt.Close
    
    'From
    iPKfrom = 0:
    t = "SELECT TOP (1) dbo.tbl_Personnel_Division.Description AS Division, dbo.tbl_Personnel_Department.DepartmentName AS Department, " & _
        " dbo.tbl_Personnel_Position.PositionName AS Position, dbo.tbl_Personnel_Position_Level.LevelName AS [Level], " & _
        " dbo.tbl_Personnel_EmploymentStatus.StatusName AS [Employment Status], dbo.tbl_Personnel_CompensationRate.Description as [Basis of Salary], " & _
        " dbo.tbl_Govt_TaxStatus.TaxStatus AS [Tax Status], dbo.tbl_Personnel_ActionNew.PK " & _
        " FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
        " dbo.tbl_Govt_TaxStatus ON dbo.tbl_Personnel_ActionNew.TaxStatusKey = dbo.tbl_Govt_TaxStatus.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_ActionNew.EmpPK = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_ActionNew.DivisionKey = dbo.tbl_Personnel_Division.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_EmploymentStatus ON dbo.tbl_Personnel_ActionNew.EmpStatusKey = dbo.tbl_Personnel_EmploymentStatus.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Position_Level ON dbo.tbl_Personnel_Position.PositionLevel = dbo.tbl_Personnel_Position_Level.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_CompensationRate ON dbo.tbl_Personnel_ActionNew.CompensationRateKey = dbo.tbl_Personnel_CompensationRate.PK " & _
        " WHERE (dbo.tbl_Personnel_ActionNew.EmpPK = " & rs!EmpPK & ") " & _
        " AND (dbo.tbl_Personnel_ActionNew.EffectivityDate < '" & FormatDateTime(rs!EffectivityDate, vbShortDate) & "') " & _
        " ORDER BY dbo.tbl_Personnel_ActionNew.EffectivityDate DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        iPKfrom = rt!PK
        iLine = 0
        For i = 1 To rt.Fields.Count - 1
            iLine = iLine + 1
            'ConnOmega.Execute "INSERT INTO tbl_Personnel_ActionNew_Report_Det " & _
                              " (MasterKey, Line, NatureOfChange, ValueTo) " & _
                              " VALUES (" & iPK & ", " & iLine & ", " & _
                              " '" & rt.Fields(i - 1).Name & "', " & _
                              " '" & FORMATSQL(CStr(rt.Fields(i - 1).Value)) & "')"
            ConnOmega.Execute "UPDATE tbl_Personnel_ActionNew_Report_Det " & _
                              " SET ValueFrom = '" & FORMATSQL(CStr(rt.Fields(i - 1).Value)) & "' " & _
                              " WHERE (MasterKey = " & iPK & ") " & _
                              " AND (Line = " & iLine & ")"
        Next i
    End If
    rt.Close
    
    'Rate To
    iLine2 = iLine
    t = "SELECT dbo.tbl_Personnel_Payroll_Earnings_Table.Description, dbo.tbl_Personnel_ActionNew_Rate.Rate " & _
        " FROM  dbo.tbl_Personnel_ActionNew_Rate LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_ActionNew_Rate.MasterKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Payroll_Earnings_Table ON dbo.tbl_Personnel_ActionNew_Rate.EarningKey = dbo.tbl_Personnel_Payroll_Earnings_Table.PK " & _
        " WHERE (dbo.tbl_Personnel_ActionNew.PK = " & iPKto & ") " & _
        " ORDER BY dbo.tbl_Personnel_Payroll_Earnings_Table.Sorting"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        iLine2 = iLine2 + 1
        ConnOmega.Execute "INSERT INTO tbl_Personnel_ActionNew_Report_Det " & _
                          " (MasterKey, Line, NatureOfChange, ValueTo, LeftRightPos) " & _
                          " VALUES (" & iPK & ", " & iLine2 & ", " & _
                          " '" & rt.Fields(0).Value & "', " & _
                          " '" & FORMATSQL(CStr(Format(rt.Fields(1).Value, "#,##0.00"))) & "', 2)"
        rt.MoveNext
    Wend
    rt.Close
    
    iLine2 = iLine
    t = "SELECT dbo.tbl_Personnel_Payroll_Earnings_Table.Description, dbo.tbl_Personnel_ActionNew_Rate.Rate " & _
        " FROM  dbo.tbl_Personnel_ActionNew_Rate LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_ActionNew_Rate.MasterKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Payroll_Earnings_Table ON dbo.tbl_Personnel_ActionNew_Rate.EarningKey = dbo.tbl_Personnel_Payroll_Earnings_Table.PK " & _
        " WHERE (dbo.tbl_Personnel_ActionNew.PK = " & iPKfrom & ") " & _
        " ORDER BY dbo.tbl_Personnel_Payroll_Earnings_Table.Sorting"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        iLine2 = iLine2 + 1
        
        u = "SELECT tbl_Personnel_ActionNew_Report_Det.* " & _
            " FROM tbl_Personnel_ActionNew_Report_Det " & _
            " WHERE (MasterKey = " & iPK & ") " & _
            " AND (NatureOfChange = '" & rt.Fields(0).Value & "')"
        If ru.State = adStateOpen Then ru.Close
        ru.Open u, ConnOmega
        If ru.RecordCount = 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Personnel_ActionNew_Report_Det " & _
                              " (MasterKey, Line, NatureOfChange, ValueFrom, LeftRightPos) " & _
                              " VALUES (" & iPK & ", " & iLine2 & ", " & _
                              " '" & rt.Fields(0).Value & "', " & _
                              " '" & FORMATSQL(CStr(Format(rt.Fields(1).Value, "#,##0.00"))) & "', 2)"
        Else
            ConnOmega.Execute "UPDATE tbl_Personnel_ActionNew_Report_Det " & _
                              " SET ValueFrom = '" & FORMATSQL(CStr(Format(rt.Fields(1).Value, "#,##0.00"))) & "' " & _
                              " WHERE (MasterKey = " & iPK & ") " & _
                              " AND (NatureOfChange = '" & rt.Fields(0).Value & "')"
        End If
        ru.Close
        
        rt.MoveNext
    Wend
    rt.Close
    
    t = "SELECT tbl_Personnel_ActionNew_Report_Det.* " & _
        " FROM tbl_Personnel_ActionNew_Report_Det " & _
        " WHERE (MasterKey = " & iPK & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        If CStr(rt!ValueFrom) <> CStr(rt!ValueTo) Then
            ConnOmega.Execute "UPDATE tbl_Personnel_ActionNew_Report_Det " & _
                              " SET isBold = 1 " & _
                              " WHERE (MasterKey = " & rt!MasterKey & ") " & _
                              " AND (Line = " & rt!line & ")"
        End If
        rt.MoveNext
    Wend
    rt.Close
    
End If
rs.Close

s = "SELECT tbl_Personnel_ActionNew_Report.*" & _
    " FROM tbl_Personnel_ActionNew_Report" & _
    " WHERE (LogIn = '" & gbl_UserName & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
rs.Requery
rs.Close

frmCrystalReportViewer.PRINT_ACTION_MEMO_V2 gbl_UserName
If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show

End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picAdd.Visible = True Then cmdCancelAdd_Click: Exit Sub
    If picSearch.Visible = True Then cmdCancelSearch_Click: Exit Sub
    Unload Me
Else
    If picSLRates.Visible = True Then
        With lstRates.ListItems
            If TRANS_DETAIL = is_DET_ADDING Then
                If .Count > 1 Then
                    .Remove .Count
                Else
                    .Item(1).SubItems(1) = "0"
                    .Item(1).SubItems(2) = " "
                    .Item(1).SubItems(3) = " "
                End If
                iRow = .Count
            End If
            If TRANS_DETAIL = is_DET_EDITTING Then
                .Item(iRow).SubItems(1) = txtEarningKey1.Text
                .Item(iRow).SubItems(2) = cmbEarningDesc.Text
                .Item(iRow).SubItems(3) = txtRate1.Text
            End If
        End With
        picSLRates.Visible = False
        picMain.Enabled = True
        picToolbar.Enabled = True
        lstRates.ListItems(iRow).EnsureVisible
        lstRates.ListItems(iRow).Selected = True
        lstRates.SetFocus
        Exit Sub
    End If
    txtName.SetFocus
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    TRANS_DETAIL = is_DET_REFRESH
    BROWSER GetSetting(App.EXEName, "PersonnelActionCtrlV2", "PerActCtrlV2", ""), "is_LOAD"
    If Trim(txtControl.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelActionCtrlV2", "PerActCtrlV2", ""), "is_HOME"
End If
End Sub

Private Sub BROWSER(Ctrl, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If Ctrl <> "" Then
            If AccessRights("Personnel Action Memo", "Supervisory") = False Then
                s = "SELECT TOP (1) dbo.tbl_Personnel_ActionNew.PK, dbo.tbl_Personnel_ActionNew.CntrlNo, dbo.tbl_Personnel_ActionNew.EmpPK, dbo.tbl_Personnel_ActionNew.DivisionKey, dbo.tbl_Personnel_ActionNew.DeptKey, dbo.tbl_Personnel_ActionNew.EmpStatusKey, dbo.tbl_Personnel_ActionNew.TaxStatusKey, dbo.tbl_Personnel_ActionNew.PositionsKey, " & _
                    " dbo.tbl_Personnel_ActionNew.CompensationRateKey, dbo.tbl_Personnel_ActionNew.TaxCategoryKey, dbo.tbl_Personnel_ActionNew.LoanDeductionKey, dbo.tbl_Personnel_ActionNew.GovtDeductionKey, dbo.tbl_Personnel_ActionNew.Is_SSS, dbo.tbl_Personnel_ActionNew.SSS, dbo.tbl_Personnel_ActionNew.Is_PHIC, dbo.tbl_Personnel_ActionNew.PHIC, " & _
                    " dbo.tbl_Personnel_ActionNew.Is_PAGIBIG , dbo.tbl_Personnel_ActionNew.PAGIBIG, dbo.tbl_Personnel_ActionNew.Is_TIN, dbo.tbl_Personnel_ActionNew.TIN, dbo.tbl_Personnel_ActionNew.EffectivityDate, dbo.tbl_Personnel_ActionNew.Remarks, dbo.tbl_Personnel_ActionNew.LastModified, dbo.tbl_Personnel_ActionNew.Locked " & _
                    " FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
                    " WHERE (dbo.tbl_Personnel_Position.PositionLevel = 1) " & _
                    " AND (dbo.tbl_Personnel_ActionNew.CntrlNo = '" & Ctrl & "') " & _
                    " ORDER BY dbo.tbl_Personnel_ActionNew.CntrlNo"
            Else
                s = "SELECT tbl_Personnel_ActionNew.* " & _
                    " FROM tbl_Personnel_ActionNew " & _
                    " WHERE (CntrlNo = '" & Ctrl & "')"
            End If
        Else
            If AccessRights("Personnel Action Memo", "Supervisory") = False Then
                s = "SELECT TOP (1) dbo.tbl_Personnel_ActionNew.PK, dbo.tbl_Personnel_ActionNew.CntrlNo, dbo.tbl_Personnel_ActionNew.EmpPK, dbo.tbl_Personnel_ActionNew.DivisionKey, dbo.tbl_Personnel_ActionNew.DeptKey, dbo.tbl_Personnel_ActionNew.EmpStatusKey, dbo.tbl_Personnel_ActionNew.TaxStatusKey, dbo.tbl_Personnel_ActionNew.PositionsKey, " & _
                    " dbo.tbl_Personnel_ActionNew.CompensationRateKey, dbo.tbl_Personnel_ActionNew.TaxCategoryKey, dbo.tbl_Personnel_ActionNew.LoanDeductionKey, dbo.tbl_Personnel_ActionNew.GovtDeductionKey, dbo.tbl_Personnel_ActionNew.Is_SSS, dbo.tbl_Personnel_ActionNew.SSS, dbo.tbl_Personnel_ActionNew.Is_PHIC, dbo.tbl_Personnel_ActionNew.PHIC, " & _
                    " dbo.tbl_Personnel_ActionNew.Is_PAGIBIG , dbo.tbl_Personnel_ActionNew.PAGIBIG, dbo.tbl_Personnel_ActionNew.Is_TIN, dbo.tbl_Personnel_ActionNew.TIN, dbo.tbl_Personnel_ActionNew.EffectivityDate, dbo.tbl_Personnel_ActionNew.Remarks, dbo.tbl_Personnel_ActionNew.LastModified, dbo.tbl_Personnel_ActionNew.Locked " & _
                    " FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
                    " WHERE (dbo.tbl_Personnel_Position.PositionLevel = 1) " & _
                    " ORDER BY dbo.tbl_Personnel_ActionNew.CntrlNo"
            Else
                s = "SELECT tbl_Personnel_ActionNew.* " & _
                    " FROM tbl_Personnel_ActionNew " & _
                    " ORDER BY CntrlNo"
            End If
        End If
    Case "is_HOME"
        If picAdd.Visible = True Then Exit Sub
        If picSLRates.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If AccessRights("Personnel Action Memo", "Supervisory") = False Then
                s = "SELECT TOP (1) dbo.tbl_Personnel_ActionNew.PK, dbo.tbl_Personnel_ActionNew.CntrlNo, dbo.tbl_Personnel_ActionNew.EmpPK, dbo.tbl_Personnel_ActionNew.DivisionKey, dbo.tbl_Personnel_ActionNew.DeptKey, dbo.tbl_Personnel_ActionNew.EmpStatusKey, dbo.tbl_Personnel_ActionNew.TaxStatusKey, dbo.tbl_Personnel_ActionNew.PositionsKey, " & _
                " dbo.tbl_Personnel_ActionNew.CompensationRateKey, dbo.tbl_Personnel_ActionNew.TaxCategoryKey, dbo.tbl_Personnel_ActionNew.LoanDeductionKey, dbo.tbl_Personnel_ActionNew.GovtDeductionKey, dbo.tbl_Personnel_ActionNew.Is_SSS, dbo.tbl_Personnel_ActionNew.SSS, dbo.tbl_Personnel_ActionNew.Is_PHIC, dbo.tbl_Personnel_ActionNew.PHIC, " & _
                " dbo.tbl_Personnel_ActionNew.Is_PAGIBIG , dbo.tbl_Personnel_ActionNew.PAGIBIG, dbo.tbl_Personnel_ActionNew.Is_TIN, dbo.tbl_Personnel_ActionNew.TIN, dbo.tbl_Personnel_ActionNew.EffectivityDate, dbo.tbl_Personnel_ActionNew.Remarks, dbo.tbl_Personnel_ActionNew.LastModified, dbo.tbl_Personnel_ActionNew.Locked " & _
                " FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
                " WHERE (dbo.tbl_Personnel_Position.PositionLevel = 1) " & _
                " ORDER BY dbo.tbl_Personnel_ActionNew.CntrlNo"
        Else
            s = "SELECT tbl_Personnel_ActionNew.* " & _
                " FROM tbl_Personnel_ActionNew " & _
                " ORDER BY CntrlNo"
        End If
    Case "is_PAGEUP"
        If picAdd.Visible = True Then Exit Sub
        If picSLRates.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If AccessRights("Personnel Action Memo", "Supervisory") = False Then
            s = "SELECT TOP (1) dbo.tbl_Personnel_ActionNew.PK, dbo.tbl_Personnel_ActionNew.CntrlNo, dbo.tbl_Personnel_ActionNew.EmpPK, dbo.tbl_Personnel_ActionNew.DivisionKey, dbo.tbl_Personnel_ActionNew.DeptKey, dbo.tbl_Personnel_ActionNew.EmpStatusKey, dbo.tbl_Personnel_ActionNew.TaxStatusKey, dbo.tbl_Personnel_ActionNew.PositionsKey, " & _
                " dbo.tbl_Personnel_ActionNew.CompensationRateKey, dbo.tbl_Personnel_ActionNew.TaxCategoryKey, dbo.tbl_Personnel_ActionNew.LoanDeductionKey, dbo.tbl_Personnel_ActionNew.GovtDeductionKey, dbo.tbl_Personnel_ActionNew.Is_SSS, dbo.tbl_Personnel_ActionNew.SSS, dbo.tbl_Personnel_ActionNew.Is_PHIC, dbo.tbl_Personnel_ActionNew.PHIC, " & _
                " dbo.tbl_Personnel_ActionNew.Is_PAGIBIG , dbo.tbl_Personnel_ActionNew.PAGIBIG, dbo.tbl_Personnel_ActionNew.Is_TIN, dbo.tbl_Personnel_ActionNew.TIN, dbo.tbl_Personnel_ActionNew.EffectivityDate, dbo.tbl_Personnel_ActionNew.Remarks, dbo.tbl_Personnel_ActionNew.LastModified, dbo.tbl_Personnel_ActionNew.Locked " & _
                " FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
                " WHERE (dbo.tbl_Personnel_Position.PositionLevel = 1) " & _
                " AND (dbo.tbl_Personnel_ActionNew.CntrlNo < '" & Ctrl & "') " & _
                " ORDER BY dbo.tbl_Personnel_ActionNew.CntrlNo DESC"
        Else
            s = "SELECT tbl_Personnel_ActionNew.* " & _
                " FROM tbl_Personnel_ActionNew " & _
                " WHERE (CntrlNo < '" & Ctrl & "') " & _
                " ORDER BY CntrlNo DESC"
        End If
    Case "is_PAGEDOWN"
        If picAdd.Visible = True Then Exit Sub
        If picSLRates.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If AccessRights("Personnel Action Memo", "Supervisory") = False Then
            s = "SELECT TOP (1) dbo.tbl_Personnel_ActionNew.PK, dbo.tbl_Personnel_ActionNew.CntrlNo, dbo.tbl_Personnel_ActionNew.EmpPK, dbo.tbl_Personnel_ActionNew.DivisionKey, dbo.tbl_Personnel_ActionNew.DeptKey, dbo.tbl_Personnel_ActionNew.EmpStatusKey, dbo.tbl_Personnel_ActionNew.TaxStatusKey, dbo.tbl_Personnel_ActionNew.PositionsKey, " & _
                " dbo.tbl_Personnel_ActionNew.CompensationRateKey, dbo.tbl_Personnel_ActionNew.TaxCategoryKey, dbo.tbl_Personnel_ActionNew.LoanDeductionKey, dbo.tbl_Personnel_ActionNew.GovtDeductionKey, dbo.tbl_Personnel_ActionNew.Is_SSS, dbo.tbl_Personnel_ActionNew.SSS, dbo.tbl_Personnel_ActionNew.Is_PHIC, dbo.tbl_Personnel_ActionNew.PHIC, " & _
                " dbo.tbl_Personnel_ActionNew.Is_PAGIBIG , dbo.tbl_Personnel_ActionNew.PAGIBIG, dbo.tbl_Personnel_ActionNew.Is_TIN, dbo.tbl_Personnel_ActionNew.TIN, dbo.tbl_Personnel_ActionNew.EffectivityDate, dbo.tbl_Personnel_ActionNew.Remarks, dbo.tbl_Personnel_ActionNew.LastModified, dbo.tbl_Personnel_ActionNew.Locked " & _
                " FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
                " WHERE (dbo.tbl_Personnel_Position.PositionLevel = 1) " & _
                " AND (dbo.tbl_Personnel_ActionNew.CntrlNo > '" & Ctrl & "') " & _
                " ORDER BY dbo.tbl_Personnel_ActionNew.CntrlNo"
        Else
            s = "SELECT tbl_Personnel_ActionNew.* " & _
                " FROM tbl_Personnel_ActionNew " & _
                " WHERE (CntrlNo > '" & Ctrl & "') " & _
                " ORDER BY CntrlNo"
        End If
    Case "is_END"
        If picAdd.Visible = True Then Exit Sub
        If picSLRates.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If AccessRights("Personnel Action Memo", "Supervisory") = False Then
            s = "SELECT TOP (1) dbo.tbl_Personnel_ActionNew.PK, dbo.tbl_Personnel_ActionNew.CntrlNo, dbo.tbl_Personnel_ActionNew.EmpPK, dbo.tbl_Personnel_ActionNew.DivisionKey, dbo.tbl_Personnel_ActionNew.DeptKey, dbo.tbl_Personnel_ActionNew.EmpStatusKey, dbo.tbl_Personnel_ActionNew.TaxStatusKey, dbo.tbl_Personnel_ActionNew.PositionsKey, " & _
                " dbo.tbl_Personnel_ActionNew.CompensationRateKey, dbo.tbl_Personnel_ActionNew.TaxCategoryKey, dbo.tbl_Personnel_ActionNew.LoanDeductionKey, dbo.tbl_Personnel_ActionNew.GovtDeductionKey, dbo.tbl_Personnel_ActionNew.Is_SSS, dbo.tbl_Personnel_ActionNew.SSS, dbo.tbl_Personnel_ActionNew.Is_PHIC, dbo.tbl_Personnel_ActionNew.PHIC, " & _
                " dbo.tbl_Personnel_ActionNew.Is_PAGIBIG , dbo.tbl_Personnel_ActionNew.PAGIBIG, dbo.tbl_Personnel_ActionNew.Is_TIN, dbo.tbl_Personnel_ActionNew.TIN, dbo.tbl_Personnel_ActionNew.EffectivityDate, dbo.tbl_Personnel_ActionNew.Remarks, dbo.tbl_Personnel_ActionNew.LastModified, dbo.tbl_Personnel_ActionNew.Locked " & _
                " FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
                " WHERE (dbo.tbl_Personnel_Position.PositionLevel = 1) " & _
                " ORDER BY dbo.tbl_Personnel_ActionNew.CntrlNo DESC"
        Else
            s = "SELECT tbl_Personnel_ActionNew.* " & _
                " FROM tbl_Personnel_ActionNew " & _
                " ORDER BY CntrlNo DESC"
        End If
    Case "is_FIND"
        s = "SELECT tbl_Personnel_ActionNew.* " & _
            " FROM tbl_Personnel_ActionNew " & _
            " WHERE (PK = " & Ctrl & ") " & _
            " ORDER BY CntrlNo DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    locTransNo = rs!PK
    locEmployeePK = rs!EmpPK
    locDept = rs!DeptKey
    locPost = rs!PositionsKey
    locEmpStatus = rs!EmpStatusKey
    locTaxStatus = rs!TaxStatusKey
    locTaxCat = rs!TaxCategoryKey
    locLoanDed = rs!LoanDeductionKey
    locGovtDed = rs!GovtDeductionKey
    locCompKey = rs!CompensationRateKey
    txtControl.Text = rs!CntrlNo
    locDivTmp = rs!DivisionKey
    locDiv = rs!DivisionKey
    
    'txtID.Text = rs!IDNumber
    txtName.Text = "" 'rs!EmployeeName
    t = "SELECT dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, " & _
        " dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName " & _
        " FROM  dbo.tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
        " WHERE (dbo.tbl_Personnel_IDNumber.PK = " & rs!EmpPK & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtName.Text = rt!IDNumber & " - " & rt!LastName & ",  " & rt!FirstName & "  " & rt!MiddleName
    End If
    rt.Close
    
    If DIV_NAME(rs!DivisionKey) <> "" Then
        Array1 = Split(DIV_NAME(rs!DivisionKey), ";", -1, 1)
        cmbDivision.Text = CStr(Array1(1))
    Else
        cmbDivision.Text = ""
    End If
    
    If DEPT_NAME(rs!DeptKey) <> "" Then
        Array1 = Split(DEPT_NAME(rs!DeptKey), ";", -1, 1)
        cmdDept.Text = CStr(Array1(1))
    Else
        cmdDept.Text = ""
    End If
    If EMP_STATUS(rs!EmpStatusKey) <> "" Then
        Array1 = Split(EMP_STATUS(rs!EmpStatusKey), ";", -1, 1)
        cmdStatus = CStr(Array1(1))
    Else
        cmdStatus.Text = ""
    End If
    If TAX_STATUS_NAME(rs!TaxStatusKey) <> "" Then
        Array1 = Split(TAX_STATUS_NAME(rs!TaxStatusKey), ";", -1, 1)
        cmbTaxStatus.Text = CStr(Array1(1))
    Else
        cmbTaxStatus.Text = ""
    End If
    If POSITION_NAME(rs!PositionsKey) <> "" Then
        Array1 = Split(POSITION_NAME(rs!PositionsKey), ";", -1, 1)
        cmbPost.Text = CStr(Array1(1))
    Else
        cmbPost.Text = ""
    End If
    If DEDUCTION_TABLE(rs!LoanDeductionKey) <> "" Then
        Array1 = Split(DEDUCTION_TABLE(rs!LoanDeductionKey), ";", -1, 1)
        cmbLoanDeduction.Text = CStr(Array1(1))
    Else
        cmbLoanDeduction.Text = ""
    End If
    If DEDUCTION_TABLE(rs!GovtDeductionKey) <> "" Then
        Array1 = Split(DEDUCTION_TABLE(rs!GovtDeductionKey), ";", -1, 1)
        cmbGovtDeduction.Text = CStr(Array1(1))
    Else
        cmbGovtDeduction.Text = ""
    End If
    If COMPENSATION_RATE(rs!CompensationRateKey) <> "" Then
        Array1 = Split(COMPENSATION_RATE(rs!CompensationRateKey), ";", -1, 1)
        cmbComp.Text = CStr(Array1(1))
    Else
        cmbComp.Text = ""
    End If
    If TAX_CATEGORY(rs!TaxCategoryKey) <> "" Then
        Array1 = Split(TAX_CATEGORY(rs!TaxCategoryKey), ";", -1, 1)
        cmbTaxCategory.Text = CStr(Array1(1))
    Else
        cmbTaxCategory.Text = ""
    End If
    
    chkSSS.Value = rs!Is_SSS
    chkPHIC.Value = rs!Is_PHIC
    chkPagIbig.Value = rs!Is_PAGIBIG
    chkTIN.Value = rs!Is_TIN
    txtSSS.Text = rs!SSS
    txtPHIC.Text = rs!PHIC
    txtPagIbig.Text = rs!PAGIBIG
    txtTIN.Text = rs!TIN
    txtRemarks.Text = rs!Remarks
    txtEffectDate.Text = Format(rs!EffectivityDate, "mm/dd/yyyy")
    
    dblTotalAmt = 0: CLEAR_Details
    t = "SELECT dbo.tbl_Personnel_ActionNew_Rate.MasterKey, " & _
        " dbo.tbl_Personnel_ActionNew_Rate.EarningKey, " & _
        " dbo.tbl_Personnel_Payroll_Earnings_Table.Description, " & _
        " dbo.tbl_Personnel_ActionNew_Rate.Rate, " & _
        " dbo.tbl_Personnel_ActionNew_Rate.RatePerHour " & _
        " FROM  dbo.tbl_Personnel_ActionNew_Rate LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Payroll_Earnings_Table ON dbo.tbl_Personnel_ActionNew_Rate.EarningKey = dbo.tbl_Personnel_Payroll_Earnings_Table.PK " & _
        " Where (dbo.tbl_Personnel_ActionNew_Rate.MasterKey = " & rs!PK & ") " & _
        " ORDER BY dbo.tbl_Personnel_Payroll_Earnings_Table.Sorting"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstRates.ListItems.Clear
        While Not rt.EOF
            dblTotalAmt = dblTotalAmt + CDbl(rt!Rate)
            Set x = lstRates.ListItems.Add()
            x.Text = rt!MasterKey
            x.SubItems(1) = rt!EarningKey
            x.SubItems(2) = rt!Description
            x.SubItems(3) = Format(rt!Rate, "#,##0.00")
            'x.SubItems(4) = rt!RatePerHour
            rt.MoveNext
        Wend
    End If
    rt.Close
    lblTotal.Caption = Format(dblTotalAmt, "#,##0.00")
    
    StatusBar.Panels(1).Text = rs!PK
    StatusBar.Panels(2).Text = IIf(IsNull(rs!LastModified), "", "LAST MODIFIED BY : " & rs!LastModified)
    StatusBar.Panels(3).Text = IIf(rs!Locked = 0, "UNLOCKED", "LOCKED")
    SaveSetting App.EXEName, "PersonnelActionCtrlV2", "PerActCtrlV2", rs!CntrlNo
    
End If
rs.Close
End Sub

Private Sub CLEAR_Details()
With lstRates.ListItems
    .Clear
    Set x = .Add()
    x.Text = "0"
    x.SubItems(1) = "0"
    x.SubItems(2) = " "
    x.SubItems(3) = " "
    'x.SubItems(4) = "0"
End With
End Sub

Public Function CLEARTEXT()
iSupervisory = 0
locEmployeePK = 0
locDept = 0
locPost = 0
locEmpStatus = 0
locTaxStatus = 0
locTaxCat = 0
locLoanDed = 0
locGovtDed = 0
locCompKey = 0
locDiv = 0
locDivTmp = 0

txtControl.Text = ""
txtName.Text = ""
cmbDivision.Text = ""
cmbDivision.ListIndex = -1
cmdDept.Text = ""
cmdDept.ListIndex = -1
cmdStatus.Text = ""
cmdStatus.ListIndex = -1
cmbTaxStatus.Text = ""
cmbTaxStatus.ListIndex = -1
cmbPost.Text = ""
cmbPost.ListIndex = -1
cmbComp.Text = ""
cmbComp.ListIndex = -1
cmbTaxCategory.Text = ""
cmbTaxCategory.ListIndex = -1
cmbLoanDeduction.Text = ""
cmbLoanDeduction.ListIndex = -1
cmbGovtDeduction.Text = ""
cmbGovtDeduction.ListIndex = -1
txtSSS.Text = ""
txtPHIC.Text = ""
txtPagIbig.Text = ""
txtTIN.Text = ""
txtEffectDate.Text = ""
txtRemarks.Text = ""
chkSSS.Value = 0
chkPHIC.Value = 0
chkPagIbig.Value = 0
chkTIN.Value = 0
lblTotal.Caption = "0.00"
StatusBar.Panels(1).Text = ""
StatusBar.Panels(2).Text = ""
StatusBar.Panels(3).Text = ""
CLEAR_Details
End Function

Public Sub LOCKTEXT(bln As Boolean)
txtName.Locked = True
cmbDivision.Locked = bln
cmdDept.Locked = bln
cmdStatus.Locked = bln
cmbTaxStatus.Locked = bln
cmbPost.Locked = bln
cmbComp.Locked = bln
cmbTaxCategory.Locked = bln
cmbLoanDeduction.Locked = bln
cmbGovtDeduction.Locked = bln
txtSSS.Locked = bln
txtPHIC.Locked = bln
txtPagIbig.Locked = bln
txtTIN.Locked = bln
txtEffectDate.Locked = bln
txtRemarks.Locked = bln
chkSSS.Enabled = IIf(bln = True, False, True)
chkPHIC.Enabled = IIf(bln = True, False, True)
chkPagIbig.Enabled = IIf(bln = True, False, True)
chkTIN.Enabled = IIf(bln = True, False, True)
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

Private Sub b8TitleBar1_CLoseClick()
cmdCancelAdd_Click
End Sub

Private Sub b8TitleBar2_CLoseClick()
cmdCancelSearch_Click
End Sub

Private Sub cmbComp_Click()
If cmbComp.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    locCompKey = cmbComp.ItemData(cmbComp.ListIndex)
End If
End Sub

Private Sub cmbDivision_Click()
If cmbDivision.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    locDiv = cmbDivision.ItemData(cmbDivision.ListIndex)
End If
End Sub

Private Sub cmbEarningDesc_Click()
If cmbEarningDesc.ListIndex = -1 Then Exit Sub
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    txtEarningKey.Text = cmbEarningDesc.ItemData(cmbEarningDesc.ListIndex)
    With lstRates.ListItems
        .Item(iRow).SubItems(2) = cmbEarningDesc.List(cmbEarningDesc.ListIndex)
    End With
End If
End Sub

Private Sub cmbEarningDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtRate.SetFocus
End Sub

Private Sub cmbEffectivityDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKSearch_Click
End Sub

Private Sub cmbGovtDeduction_Click()
If cmbGovtDeduction.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    locGovtDed = cmbGovtDeduction.ItemData(cmbGovtDeduction.ListIndex)
End If
End Sub

Private Sub cmbLoanDeduction_Click()
If cmbLoanDeduction.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    locLoanDed = cmbLoanDeduction.ItemData(cmbLoanDeduction.ListIndex)
End If
End Sub

Private Sub cmbPost_Click()
If cmbPost.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    locPost = cmbPost.ItemData(cmbPost.ListIndex)
    t = "SELECT dbo.tbl_Personnel_Position.PositionLevel " & _
        " FROM  dbo.tbl_Personnel_Position LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Position_Level ON dbo.tbl_Personnel_Position.PositionLevel = dbo.tbl_Personnel_Position_Level.PK " & _
        " WHERE (dbo.tbl_Personnel_Position.PK = " & locPost & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        iSupervisory = rt!PositionLevel
    End If
    rt.Close
End If
End Sub

Private Sub cmbTaxCategory_Click()
If cmbTaxCategory.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    locTaxCat = cmbTaxCategory.ItemData(cmbTaxCategory.ListIndex)
End If
End Sub

Private Sub cmbTaxStatus_Click()
If cmbTaxStatus.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    locTaxStatus = cmbTaxStatus.ItemData(cmbTaxStatus.ListIndex)
End If
End Sub

Private Sub cmdCancelAdd_Click()
picMain.Enabled = True
picToolbar.Enabled = True
picAdd.Visible = False
End Sub

Private Sub cmdCancelSearch_Click()
picSearch.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True
End Sub

Private Sub cmdDept_Click()
If cmdDept.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    locDept = cmdDept.ItemData(cmdDept.ListIndex)
End If
End Sub

Private Sub cmdOKAdd_Click()
If lstResultAdd.ListIndex = -1 Then MsgBox "Please select employee!                  ", vbCritical, "Error...": txtSearchAdd.SetFocus: Exit Sub
If IsDate(txtEffectDateAdd.Text) = False Then MsgBox "Please supply a valid date!             ", vbCritical, "Error...": txtEffectDateAdd.SetFocus: Exit Sub
CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING

locEmployeePK = lstResultAdd.ItemData(lstResultAdd.ListIndex)
txtEffectDate.Text = Format(FormatDateTime(txtEffectDateAdd.Text, vbShortDate), "mm/dd/yyyy")
txtName.Text = lstResultAdd.List(lstResultAdd.ListIndex)
s = "SELECT tbl_Personnel_Information.SSSNumber AS SSS, " & _
    " tbl_Personnel_Information.PHICNumber AS PHIC, " & _
    " tbl_Personnel_Information.HDMFNumber AS PagIbig, " & _
    " tbl_Personnel_Information.TIN AS TIN " & _
    " FROM tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
    " WHERE (tbl_Personnel_IDNumber.PK = " & locEmployeePK & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtSSS.Text = rs!SSS
    txtPHIC.Text = rs!PHIC
    txtPagIbig.Text = rs!PAGIBIG
    txtTIN.Text = rs!TIN
Else
    txtSSS.Text = "ON PROCESS"
    txtPHIC.Text = "ON PROCESS"
    txtPagIbig.Text = "ON PROCESS"
    txtTIN.Text = "ON PROCESS"
End If
rs.Close
s = "SELECT TOP (1) PK, Is_SSS, Is_PHIC, Is_PAGIBIG, Is_TIN, " & _
    " DivisionKey, DeptKey, EmpStatusKey, TaxStatusKey, PositionsKey, CompensationRateKey, " & _
    " TaxCategoryKey, LoanDeductionKey, GovtDeductionKey " & _
    " From tbl_Personnel_ActionNew " & _
    " Where (EmpPK = " & locEmployeePK & ") " & _
    " And (EffectivityDate <= '" & FormatDateTime(txtEffectDateAdd.Text, vbShortDate) & "') " & _
    " ORDER BY EffectivityDate DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    chkSSS.Value = rs!Is_SSS
    chkPHIC.Value = rs!Is_PHIC
    chkPagIbig.Value = rs!Is_PAGIBIG
    chkTIN.Value = rs!Is_TIN
    
    Array1 = Split(DIV_NAME(rs!DivisionKey), ";", -1, 1)
    cmbDivision.Text = CStr(Array1(1))
    Array1 = Split(DEPT_NAME(rs!DeptKey), ";", -1, 1)
    cmdDept.Text = CStr(Array1(1))
    Array1 = Split(EMP_STATUS(rs!EmpStatusKey), ";", -1, 1)
    cmdStatus.Text = CStr(Array1(1))
    Array1 = Split(TAX_STATUS_NAME(rs!TaxStatusKey), ";", -1, 1)
    cmbTaxStatus.Text = CStr(Array1(1))
    Array1 = Split(POSITION_NAME(rs!PositionsKey), ";", -1, 1)
    cmbPost.Text = CStr(Array1(1))
    Array1 = Split(DEDUCTION_TABLE(rs!LoanDeductionKey), ";", -1, 1)
    cmbLoanDeduction.Text = CStr(Array1(1))
    Array1 = Split(DEDUCTION_TABLE(rs!GovtDeductionKey), ";", -1, 1)
    cmbGovtDeduction.Text = CStr(Array1(1))
    Array1 = Split(COMPENSATION_RATE(rs!CompensationRateKey), ";", -1, 1)
    cmbComp.Text = CStr(Array1(1))
    Array1 = Split(TAX_CATEGORY(rs!TaxCategoryKey), ";", -1, 1)
    cmbTaxCategory.Text = CStr(Array1(1))
    
    locDiv = rs!DivisionKey
    locDept = rs!DeptKey
    locPost = rs!PositionsKey
    locCompKey = rs!CompensationRateKey
    locEmpStatus = rs!EmpStatusKey
    locTaxStatus = rs!TaxStatusKey
    locTaxCat = rs!TaxCategoryKey
    locLoanDed = rs!LoanDeductionKey
    locGovtDed = rs!GovtDeductionKey
    
    CLEAR_Details
    t = "SELECT dbo.tbl_Personnel_ActionNew_Rate.MasterKey, " & _
        " dbo.tbl_Personnel_ActionNew_Rate.EarningKey, " & _
        " dbo.tbl_Personnel_Payroll_Earnings_Table.Description, " & _
        " dbo.tbl_Personnel_ActionNew_Rate.Rate, " & _
        " dbo.tbl_Personnel_ActionNew_Rate.RatePerHour " & _
        " FROM  dbo.tbl_Personnel_ActionNew_Rate LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Payroll_Earnings_Table ON dbo.tbl_Personnel_ActionNew_Rate.EarningKey = dbo.tbl_Personnel_Payroll_Earnings_Table.PK " & _
        " Where (dbo.tbl_Personnel_ActionNew_Rate.MasterKey = " & rs!PK & ") " & _
        " ORDER BY dbo.tbl_Personnel_Payroll_Earnings_Table.Sorting"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstRates.ListItems.Clear
        While Not rt.EOF
            Set x = lstRates.ListItems.Add()
            x.Text = rt!MasterKey
            x.SubItems(1) = rt!EarningKey
            x.SubItems(2) = rt!Description
            x.SubItems(3) = Format(rt!Rate, "#,##0.00")
            'x.SubItems(4) = rt!RatePerHour
            rt.MoveNext
        Wend
    End If
    rt.Close
    
End If
rs.Close

cmdCancelAdd_Click
txtRemarks.SetFocus

End Sub

Private Sub cmdOKSearch_Click()
If cmbEffectivityDate.ListIndex = -1 Then Exit Sub
BROWSER cmbEffectivityDate.ItemData(cmbEffectivityDate.ListIndex), "is_FIND"
cmdCancelSearch_Click
End Sub

Private Sub cmdStatus_Click()
If cmdStatus.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    locEmpStatus = cmdStatus.ItemData(cmdStatus.ListIndex)
End If
End Sub

Private Sub Command1_Click()

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
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "PersonnelActionCtrlV2", "PerActCtrlV2", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "PersonnelActionCtrlV2", "PerActCtrlV2", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "PersonnelActionCtrlV2", "PerActCtrlV2", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "PersonnelActionCtrlV2", "PerActCtrlV2", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption & " [V2]"
Me.Top = (MainForm.ScaleHeight - Me.Height) / 3
Me.Left = (MainForm.ScaleWidth - Me.Width) / 3
POPULATE_COMBO "PK", "Description", "tbl_Personnel_Division", "Description", cmbDivision
POPULATE_COMBO "PK", "DepartmentName", "tbl_Personnel_Department", "DepartmentName", cmdDept
POPULATE_COMBO "PK", "PositionName", "tbl_Personnel_Position", "PositionName", cmbPost
POPULATE_COMBO "PK", "TaxStatus", "tbl_Govt_TaxStatus", "PK", cmbTaxStatus
POPULATE_COMBO "PK", "StatusName", "tbl_Personnel_EmploymentStatus", "StatusName", cmdStatus
POPULATE_COMBO "PK", "Description", "tbl_Personnel_CompensationRate", "PK", cmbComp
POPULATE_COMBO "PK", "TaxCategory", "tbl_Govt_TaxCategory", "PK", cmbTaxCategory
POPULATE_COMBO "PK", "Description", "tbl_Personnel_ActionNew_DedTable", "PK", cmbLoanDeduction
POPULATE_COMBO "PK", "Description", "tbl_Personnel_ActionNew_DedTable", "PK", cmbGovtDeduction

cmbEarningDesc.Clear
s = "SELECT tbl_Personnel_Payroll_Earnings_Table.* " & _
    " FROM tbl_Personnel_Payroll_Earnings_Table " & _
    " WHERE (ViewInActionModule = 1) " & _
    " ORDER BY Sorting"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbEarningDesc.AddItem rs!Description
    cmbEarningDesc.ItemData(cmbEarningDesc.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close

iRow = 0
isFocus = 0
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANS_DETAIL = is_DET_REFRESH
BROWSER GetSetting(App.EXEName, "PersonnelActionCtrlV2", "PerActCtrlV2", ""), "is_LOAD"
If Trim(txtControl.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelActionCtrlV2", "PerActCtrlV2", ""), "is_HOME"

tmp = SetWindowLong(txtSSS.hwnd, GWL_STYLE, GetWindowLong(txtSSS.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPHIC.hwnd, GWL_STYLE, GetWindowLong(txtPHIC.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPagIbig.hwnd, GWL_STYLE, GetWindowLong(txtPagIbig.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtTIN.hwnd, GWL_STYLE, GetWindowLong(txtTIN.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtRemarks.hwnd, GWL_STYLE, GetWindowLong(txtRemarks.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearchAdd.hwnd, GWL_STYLE, GetWindowLong(txtSearchAdd.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearchSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearchSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)

End Sub

Private Sub Form_Unload(Cancel As Integer)
If picAdd.Visible = True Then Cancel = -1
If picSLRates.Visible = True Then Cancel = -1
If picSearch.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstRates_Click()
isFocus = 1
TRANS_DETAIL = is_DET_REFRESH
If lstRates.ListItems.Count = 0 Then
    iRow = 0: Exit Sub
End If
iRow = lstRates.SelectedItem.Index
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If CDbl(lstRates.ListItems.Item(iRow).SubItems(1)) = 0 Then TOOLBARFUNC 4 Else TOOLBARFUNC 5
End If
End Sub

Private Sub lstRates_GotFocus()
isFocus = 1
TRANS_DETAIL = is_DET_REFRESH
If lstRates.ListItems.Count = 0 Then
    iRow = 0: Exit Sub
End If
iRow = lstRates.SelectedItem.Index
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If CDbl(lstRates.ListItems.Item(iRow).SubItems(1)) = 0 Then TOOLBARFUNC 4 Else TOOLBARFUNC 5
End If
End Sub

Private Sub lstRates_ItemClick(ByVal Item As MSComctlLib.ListItem)
If lstRates.ListItems.Count = 0 Then
    iRow = 0: Exit Sub
End If
iRow = lstRates.SelectedItem.Index
End Sub

Private Sub lstRates_LostFocus()
isFocus = 0
End Sub

Private Sub lstResultAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtEffectDateAdd.SetFocus
End Sub

Private Sub lstResultSearch_Click()
If lstResultSearch.ListIndex = -1 Then cmbEffectivityDate.Clear: Exit Sub
cmbEffectivityDate.Clear
s = "SELECT PK, EffectivityDate" & _
    " From tbl_Personnel_ActionNew  " & _
    " Where (EmpPK = " & lstResultSearch.ItemData(lstResultSearch.ListIndex) & ") " & _
    " ORDER BY EffectivityDate DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbEffectivityDate.AddItem Format(rs!EffectivityDate, "mm/dd/yyyy")
    cmbEffectivityDate.ItemData(cmbEffectivityDate.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If cmbEffectivityDate.ListCount Then cmbEffectivityDate.ListIndex = 0
End Sub

Private Sub lstResultSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbEffectivityDate.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "PersonnelActionCtrlV2", "PerActCtrlV2", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "PersonnelActionCtrlV2", "PerActCtrlV2", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "PersonnelActionCtrlV2", "PerActCtrlV2", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "PersonnelActionCtrlV2", "PerActCtrlV2", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":   PRESS_F9
    Case "Refresh":
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtEarningKey_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstRates.ListItems
        .Item(iRow).SubItems(1) = txtEarningKey.Text
    End With
End If
End Sub

Private Sub txtEffectDateAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAdd_Click
End Sub

Private Sub txtEffectDateAdd_LostFocus()
If IsDate(txtEffectDateAdd.Text) Then
    txtEffectDateAdd.Text = Format(FormatDateTime(txtEffectDateAdd.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub

Private Sub txtRate_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstRates.ListItems
        .Item(iRow).SubItems(3) = Format(RETURNTEXTVALUE(txtRate), "#,##0.00")
        dblTotalAmt = 0
        For i = 1 To .Count
            dblTotalAmt = dblTotalAmt + CDbl(IIf(IsNumeric(.Item(iRow).SubItems(3)) = False, 0, .Item(iRow).SubItems(3)))
        Next i
        lblTotal.Caption = Format(dblTotalAmt, "#,##0.00")
    End With
End If
End Sub

Private Sub txtRate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    picSLRates.Visible = False
    picMain.Enabled = True
    picToolbar.Enabled = True
    lstRates.SetFocus
End If
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtSearchAdd_Change()
If Trim(txtSearchAdd.Text) = "" Then lstResultAdd.Clear: Exit Sub
lstResultAdd.Clear
's = "sp_Personnel_Action_Search_Add('" & FORMATSQL(Trim(txtSearch.Text)) & "%')"
s = "SELECT tbl_Personnel_IDNumber.PK, " & _
    " tbl_Personnel_IDNumber.IDNumber, " & _
    " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName " & _
    " FROM tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
    " WHERE (tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%') " & _
    " AND (ISNULL((SELECT TOP 1 tbl_Personnel_EmploymentStatus.Active " & _
    " FROM tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
    " tbl_Personnel_EmploymentStatus ON tbl_Personnel_ActionNew.EmpStatusKey = tbl_Personnel_EmploymentStatus.PK " & _
    " WHERE (tbl_Personnel_ActionNew.EmpPK = tbl_Personnel_IDNumber.PK) " & _
    " AND (tbl_Personnel_ActionNew.EffectivityDate <= CONVERT(DATETIME, CONVERT(char(6), getdate(), 12), 102)) ORDER BY tbl_Personnel_ActionNew.EffectivityDate DESC), 0) = 1) " & _
    " OR (tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%') " & _
    " AND (ISNULL((SELECT TOP 1 tbl_Personnel_EmploymentStatus.Active " & _
    " FROM tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
    " tbl_Personnel_EmploymentStatus ON tbl_Personnel_ActionNew.EmpStatusKey = tbl_Personnel_EmploymentStatus.PK " & _
    " WHERE (tbl_Personnel_ActionNew.EmpPK = tbl_Personnel_IDNumber.PK) " & _
    " AND (tbl_Personnel_ActionNew.EffectivityDate <= CONVERT(DATETIME, CONVERT(char(6), getdate(), 12), 102)) ORDER BY tbl_Personnel_ActionNew.EffectivityDate DESC), 0) = 0) " & _
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

Private Sub txtSearchSearch_Change()
If Trim(txtSearchSearch.Text) = "" Then lstResultSearch.Clear: cmbEffectivityDate.Clear: Exit Sub
lstResultSearch.Clear: cmbEffectivityDate.Clear
's = "SELECT tbl_Personnel_IDNumber.PK, tbl_Personnel_IDNumber.IDNumber, " & _
    " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName " & _
    " FROM tbl_Personnel_ActionNew AS tbl_Personnel_ActionNew_1 LEFT OUTER JOIN " & _
    " tbl_Personnel_IDNumber ON tbl_Personnel_ActionNew_1.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
    " WHERE (tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearchSearch.Text)) & "%') " & _
    " GROUP BY tbl_Personnel_IDNumber.PK, tbl_Personnel_IDNumber.IDNumber, " & _
    " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName " & _
    " HAVING (ISNULL ((SELECT TOP 1 tbl_Personnel_EmploymentStatus_1.Active " & _
    " FROM tbl_Personnel_ActionNew AS tbl_Personnel_ActionNew_2 LEFT OUTER JOIN " & _
    " tbl_Personnel_EmploymentStatus AS tbl_Personnel_EmploymentStatus_1 ON " & _
    " tbl_Personnel_ActionNew_2.EmpStatus = tbl_Personnel_EmploymentStatus_1.PK " & _
    " Where (tbl_Personnel_ActionNew_2.EmpPK = tbl_Personnel_IDNumber.PK) " & _
    " ORDER BY tbl_Personnel_ActionNew_2.EffectivityDate DESC), 0) = 1) " & _
    " ORDER BY tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName, " & _
    " tbl_Personnel_IDNumber.IDNumber "
If AccessRights("Personnel Action Memo", "Supervisory") = False Then
    s = "SELECT dbo.tbl_Personnel_IDNumber.PK, dbo.tbl_Personnel_IDNumber.IDNumber, " & _
        " dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
        " dbo.tbl_Personnel_Information.MiddleName " & _
        " FROM  dbo.tbl_Personnel_ActionNew INNER JOIN " & _
        " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_ActionNew.EmpPK = dbo.tbl_Personnel_IDNumber.PK INNER JOIN " & _
        " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK INNER JOIN " & _
        " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
        " Where (dbo.tbl_Personnel_Position.PositionLevel = 1) " & _
        " GROUP BY dbo.tbl_Personnel_IDNumber.PK, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName " & _
        " HAVING (dbo.tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearchSearch.Text)) & "%') " & _
        " ORDER BY dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_IDNumber.IDNumber"
Else
    s = "SELECT tbl_Personnel_IDNumber.PK, tbl_Personnel_IDNumber.IDNumber, " & _
        " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName " & _
        " FROM tbl_Personnel_ActionNew AS tbl_Personnel_ActionNew_1 LEFT OUTER JOIN " & _
        " tbl_Personnel_IDNumber ON tbl_Personnel_ActionNew_1.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
        " WHERE (tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearchSearch.Text)) & "%') " & _
        " GROUP BY tbl_Personnel_IDNumber.PK, tbl_Personnel_IDNumber.IDNumber, " & _
        " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName " & _
        " ORDER BY tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName, " & _
        " tbl_Personnel_IDNumber.IDNumber "
End If
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultSearch.AddItem rs!IDNumber & " - " & rs!EmployeeName
    lstResultSearch.ItemData(lstResultSearch.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResultSearch.ListCount Then lstResultSearch.ListIndex = 0
End Sub
