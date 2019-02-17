VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelDeductionsForPayroll 
   Appearance      =   0  'Flat
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12645
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
   ScaleHeight     =   8685
   ScaleWidth      =   12645
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picAdd 
      Height          =   2295
      Left            =   4680
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4048
      BackColor       =   15396057
      Begin VB.ComboBox cmbDivision 
         Height          =   315
         Left            =   240
         TabIndex        =   25
         Text            =   "Combo1"
         Top             =   600
         Width           =   3495
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
         Picture         =   "frmPersonnelDeductionsForPayroll.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1560
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
         Picture         =   "frmPersonnelDeductionsForPayroll.frx":0672
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1560
         Width           =   1560
      End
      Begin VB.TextBox txtPayrollDateAdd 
         Height          =   315
         Left            =   1680
         TabIndex        =   20
         Top             =   1080
         Width           =   1575
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   23
         Top             =   45
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   609
         Caption         =   "Add"
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
         Icon            =   "frmPersonnelDeductionsForPayroll.frx":0DCE
         ShadowVisible   =   0   'False
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   24
         Top             =   1080
         Width           =   1215
      End
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
         MouseIcon       =   "frmPersonnelDeductionsForPayroll.frx":1368
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   10980
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   2
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmPersonnelDeductionsForPayroll.frx":1682
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
      Top             =   720
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
            Picture         =   "frmPersonnelDeductionsForPayroll.frx":1D95
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductionsForPayroll.frx":2A6F
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductionsForPayroll.frx":3749
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductionsForPayroll.frx":4423
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductionsForPayroll.frx":50FD
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductionsForPayroll.frx":5DD7
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductionsForPayroll.frx":6AB1
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductionsForPayroll.frx":778B
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductionsForPayroll.frx":8465
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductionsForPayroll.frx":8D3F
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductionsForPayroll.frx":9A19
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductionsForPayroll.frx":A6F3
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductionsForPayroll.frx":B3CD
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductionsForPayroll.frx":C0A7
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeductionsForPayroll.frx":CD81
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   8385
      Width           =   12645
      _ExtentX        =   22304
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
      Left            =   6840
      TabIndex        =   26
      Top             =   1200
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox cmbDeductionName 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtPayrollDateSL 
         Height          =   315
         Left            =   2760
         TabIndex        =   35
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4200
         TabIndex        =   32
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtDeductionKey 
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtPayrollDateSL1 
         Height          =   315
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtAmount1 
         Height          =   315
         Left            =   4800
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox cmbDeductionName1 
         Height          =   315
         Left            =   4320
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtDeductionKey1 
         Height          =   315
         Left            =   4080
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   36
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Deduction Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4200
         TabIndex        =   33
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   360
      ScaleHeight     =   6855
      ScaleWidth      =   11895
      TabIndex        =   4
      Top             =   1200
      Width           =   11895
      Begin VB.TextBox txtEmpKey 
         Height          =   315
         Left            =   600
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   38
         Top             =   360
         Visible         =   0   'False
         Width           =   150
      End
      Begin MSComctlLib.ListView lstGlobalList 
         Height          =   375
         Left            =   840
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "EmployeeKey"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Employee Name"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Total"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "DeductionKey"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DeductionName"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "SourceKey"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "PayrollDate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstEmployee 
         Height          =   5655
         Left            =   0
         TabIndex        =   12
         Top             =   840
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   9975
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "EmployeeKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Employee Name"
            Object.Width           =   10760
         EndProperty
      End
      Begin VB.TextBox txtDivisionName 
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Top             =   0
         Width           =   6975
      End
      Begin VB.TextBox txtPayrollDate 
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtCutOffDate 
         Height          =   315
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox txtCtrl 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   150
      End
      Begin MSComctlLib.ListView lstDeductionList 
         Height          =   5655
         Left            =   6600
         TabIndex        =   13
         Top             =   840
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   9975
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
            Text            =   "DeductionKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Deduction Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "SourceKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Payroll Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Total >>"
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
         TabIndex        =   17
         Top             =   6600
         Width           =   735
      End
      Begin VB.Label lblTotalEmployee 
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
         Left            =   10440
         TabIndex        =   16
         Top             =   6600
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
         Left            =   3600
         TabIndex        =   15
         Top             =   6600
         Visible         =   0   'False
         Width           =   1215
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
         Left            =   5040
         TabIndex        =   14
         Top             =   6600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Division Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   0
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cut-Off Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3360
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmPersonnelDeductionsForPayroll"
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

Dim isFocusEmp, iRowEmp       As Long
Dim isFocusDed, iRowDed       As Long

Public locDivisionKey, locPayrollPeriod    As Long

Dim Array1, x, y, i, sCtrl, iPK, iEmpStatus, locPayrollPeriodTmp, dEmpForDed
Dim dTotDed, dTotBal, iLineCnt, dDedPerPay, cnt, sDeductionKey, iDeductionKeyCnt
Dim sPayDate, iPeriodTerms
Dim sFileName, RowCnt, ColCnt, iWorkSheet, RowFreeze, ColFreeze, WorkbookName, strRange

Private Sub BROWSER(Ctrl, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If Ctrl <> "" Then
            s = "SELECT TOP (1) dbo.tbl_Personnel_Deduction_forPayroll.Ctrl, dbo.tbl_Personnel_Deduction_forPayroll.DivisionKey, " & _
                " dbo.tbl_Personnel_Deduction_forPayroll.PayrollPeriodKey, dbo.tbl_Personnel_Deduction_forPayroll.Remarks, " & _
                " dbo.tbl_Personnel_Deduction_forPayroll.Posted, dbo.tbl_Personnel_Deduction_forPayroll.LastModified, " & _
                " dbo.tbl_Personnel_Division.Description as Division, dbo.tbl_Personnel_Compensation_Period.PayrollDate, " & _
                " dbo.tbl_Personnel_Deduction_forPayroll.PK, dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo " & _
                " FROM  dbo.tbl_Personnel_Deduction_forPayroll LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Deduction_forPayroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_Deduction_forPayroll.DivisionKey = dbo.tbl_Personnel_Division.PK " & _
                " WHERE (dbo.tbl_Personnel_Deduction_forPayroll.Ctrl = '" & Ctrl & "') " & _
                " ORDER BY dbo.tbl_Personnel_Deduction_forPayroll.Ctrl"
        Else
            s = "SELECT TOP (1) dbo.tbl_Personnel_Deduction_forPayroll.Ctrl, dbo.tbl_Personnel_Deduction_forPayroll.DivisionKey, " & _
                " dbo.tbl_Personnel_Deduction_forPayroll.PayrollPeriodKey, dbo.tbl_Personnel_Deduction_forPayroll.Remarks, " & _
                " dbo.tbl_Personnel_Deduction_forPayroll.Posted, dbo.tbl_Personnel_Deduction_forPayroll.LastModified, " & _
                " dbo.tbl_Personnel_Division.Description as Division, dbo.tbl_Personnel_Compensation_Period.PayrollDate, " & _
                " dbo.tbl_Personnel_Deduction_forPayroll.PK, dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo " & _
                " FROM  dbo.tbl_Personnel_Deduction_forPayroll LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Deduction_forPayroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_Deduction_forPayroll.DivisionKey = dbo.tbl_Personnel_Division.PK " & _
                " ORDER BY dbo.tbl_Personnel_Deduction_forPayroll.Ctrl"
        End If
    Case "is_HOME"
        If picAdd.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Deduction_forPayroll.Ctrl, dbo.tbl_Personnel_Deduction_forPayroll.DivisionKey, " & _
            " dbo.tbl_Personnel_Deduction_forPayroll.PayrollPeriodKey, dbo.tbl_Personnel_Deduction_forPayroll.Remarks, " & _
            " dbo.tbl_Personnel_Deduction_forPayroll.Posted, dbo.tbl_Personnel_Deduction_forPayroll.LastModified, " & _
            " dbo.tbl_Personnel_Division.Description as Division, dbo.tbl_Personnel_Compensation_Period.PayrollDate, " & _
            " dbo.tbl_Personnel_Deduction_forPayroll.PK, dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo " & _
            " FROM  dbo.tbl_Personnel_Deduction_forPayroll LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Deduction_forPayroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_Deduction_forPayroll.DivisionKey = dbo.tbl_Personnel_Division.PK " & _
            " ORDER BY dbo.tbl_Personnel_Deduction_forPayroll.Ctrl"
    Case "is_PAGEUP"
        If picAdd.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Deduction_forPayroll.Ctrl, dbo.tbl_Personnel_Deduction_forPayroll.DivisionKey, " & _
            " dbo.tbl_Personnel_Deduction_forPayroll.PayrollPeriodKey, dbo.tbl_Personnel_Deduction_forPayroll.Remarks, " & _
            " dbo.tbl_Personnel_Deduction_forPayroll.Posted, dbo.tbl_Personnel_Deduction_forPayroll.LastModified, " & _
            " dbo.tbl_Personnel_Division.Description as Division, dbo.tbl_Personnel_Compensation_Period.PayrollDate, " & _
            " dbo.tbl_Personnel_Deduction_forPayroll.PK, dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo " & _
            " FROM  dbo.tbl_Personnel_Deduction_forPayroll LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Deduction_forPayroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_Deduction_forPayroll.DivisionKey = dbo.tbl_Personnel_Division.PK " & _
            " WHERE (dbo.tbl_Personnel_Deduction_forPayroll.Ctrl < '" & Ctrl & "') " & _
            " ORDER BY dbo.tbl_Personnel_Deduction_forPayroll.Ctrl DESC"
    Case "is_PAGEDOWN"
        If picAdd.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Deduction_forPayroll.Ctrl, dbo.tbl_Personnel_Deduction_forPayroll.DivisionKey, " & _
            " dbo.tbl_Personnel_Deduction_forPayroll.PayrollPeriodKey, dbo.tbl_Personnel_Deduction_forPayroll.Remarks, " & _
            " dbo.tbl_Personnel_Deduction_forPayroll.Posted, dbo.tbl_Personnel_Deduction_forPayroll.LastModified, " & _
            " dbo.tbl_Personnel_Division.Description as Division, dbo.tbl_Personnel_Compensation_Period.PayrollDate, " & _
            " dbo.tbl_Personnel_Deduction_forPayroll.PK, dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo " & _
            " FROM  dbo.tbl_Personnel_Deduction_forPayroll LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Deduction_forPayroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_Deduction_forPayroll.DivisionKey = dbo.tbl_Personnel_Division.PK " & _
            " WHERE (dbo.tbl_Personnel_Deduction_forPayroll.Ctrl > '" & Ctrl & "') " & _
            " ORDER BY dbo.tbl_Personnel_Deduction_forPayroll.Ctrl"
    Case "is_END"
        If picAdd.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Deduction_forPayroll.Ctrl, dbo.tbl_Personnel_Deduction_forPayroll.DivisionKey, " & _
            " dbo.tbl_Personnel_Deduction_forPayroll.PayrollPeriodKey, dbo.tbl_Personnel_Deduction_forPayroll.Remarks, " & _
            " dbo.tbl_Personnel_Deduction_forPayroll.Posted, dbo.tbl_Personnel_Deduction_forPayroll.LastModified, " & _
            " dbo.tbl_Personnel_Division.Description as Division, dbo.tbl_Personnel_Compensation_Period.PayrollDate, " & _
            " dbo.tbl_Personnel_Deduction_forPayroll.PK, dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo " & _
            " FROM  dbo.tbl_Personnel_Deduction_forPayroll LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Deduction_forPayroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_Deduction_forPayroll.DivisionKey = dbo.tbl_Personnel_Division.PK " & _
            " ORDER BY dbo.tbl_Personnel_Deduction_forPayroll.Ctrl DESC"
    Case "is_FIND"
        s = "SELECT TOP (1) dbo.tbl_Personnel_Deduction_forPayroll.Ctrl, dbo.tbl_Personnel_Deduction_forPayroll.DivisionKey, " & _
            " dbo.tbl_Personnel_Deduction_forPayroll.PayrollPeriodKey, dbo.tbl_Personnel_Deduction_forPayroll.Remarks, " & _
            " dbo.tbl_Personnel_Deduction_forPayroll.Posted, dbo.tbl_Personnel_Deduction_forPayroll.LastModified, " & _
            " dbo.tbl_Personnel_Division.Description as Division, dbo.tbl_Personnel_Compensation_Period.PayrollDate, " & _
            " dbo.tbl_Personnel_Deduction_forPayroll.PK, dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo " & _
            " FROM  dbo.tbl_Personnel_Deduction_forPayroll LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Deduction_forPayroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_Deduction_forPayroll.DivisionKey = dbo.tbl_Personnel_Division.PK " & _
            " WHERE (dbo.tbl_Personnel_Deduction_forPayroll.PK = " & Ctrl & ") " & _
            " ORDER BY dbo.tbl_Personnel_Deduction_forPayroll.Ctrl DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    locDivisionKey = rs!DivisionKey
    locPayrollPeriod = rs!PayrollPeriodKey
    txtCtrl.Text = rs!Ctrl
    txtDivisionName.Text = rs!Division
    txtPayrollDate.Text = Format(rs!PayrollDate, "mm/dd/yyyy")
    txtCutOffDate.Text = Format(rs!DateFrom, "mm/dd/yyyy") & " - " & Format(rs!DateTo, "mm/dd/yyyy")
    
    dTotBal = 0: dTotDed = 0
    CLEAR_Details_Emp
    CLEAR_Details_Ded
    
    
    t = "SELECT dbo.tbl_Personnel_Deduction_forPayroll_Det.EmployeeKey, dbo.tbl_Personnel_IDNumber.IDNumber, " & _
        " dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
        " dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Deduction_forPayroll_Det.Balance " & _
        " FROM  dbo.tbl_Personnel_Deduction_forPayroll_Det RIGHT OUTER JOIN " & _
        " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Deduction_forPayroll_Det.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
        " Where (dbo.tbl_Personnel_Deduction_forPayroll_Det.MasterKey = " & rs!PK & ") " & _
        " ORDER BY dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_IDNumber.IDNumber"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstEmployee.ListItems.Clear
        While Not rt.EOF
            Set x = lstEmployee.ListItems.Add()
            x.Text = rt!EmployeeKey
            x.SubItems(1) = rt!IDNumber & " - " & rt!LastName & ",  " & rt!FirstName & "  " & rt!MiddleName
            'x.SubItems(2) = "" 'Format(rt!Balance, "#,##0.00")
            dTotBal = dTotBal + CDbl(rt!Balance)
            rt.MoveNext
        Wend
    End If
    rt.Close
    
    lstGlobalList.ListItems.Clear
    t = "SELECT dbo.tbl_Personnel_Deduction_forPayroll_Det.EmployeeKey, dbo.tbl_Personnel_IDNumber.IDNumber, " & _
        " dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
        " dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Deduction_forPayroll_Det.Balance, " & _
        " dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.DeductionKey, dbo.tbl_Personnel_Payroll_Deductions_Table.Description, " & _
        " dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.SourceKey, " & _
        " dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.Amount, dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.PayrollDate " & _
        " FROM  dbo.tbl_Personnel_Deduction_forPayroll_Det LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Deduction_forPayroll_Det_Det ON dbo.tbl_Personnel_Deduction_forPayroll_Det.MasterKey = dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.MasterKey AND dbo.tbl_Personnel_Deduction_forPayroll_Det.EmployeeKey = dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.EmployeeKey LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Deduction_forPayroll_Det.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.DeductionKey = dbo.tbl_Personnel_Payroll_Deductions_Table.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Deduction ON dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.SourceKey = dbo.tbl_Personnel_Deduction.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Deduction.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
        " Where (dbo.tbl_Personnel_Deduction_forPayroll_Det.MasterKey = " & rs!PK & ") " & _
        " ORDER BY dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_IDNumber.IDNumber"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        Set y = lstGlobalList.ListItems.Add()
        y.Text = rt!EmployeeKey
        y.SubItems(1) = rt!IDNumber & " - " & rt!LastName & ",  " & rt!FirstName & "  " & rt!MiddleName
        y.SubItems(2) = Format(rt!Balance, "#,##0.00")
        y.SubItems(3) = rt!DeductionKey 'DeductionKey
        y.SubItems(4) = rt!Description 'DeductionName
        y.SubItems(5) = rt!SourceKey 'SourceKey
        y.SubItems(6) = IIf(IsNull(rt!PayrollDate), "", rt!PayrollDate) 'PayrollDate
        y.SubItems(7) = Format(rt!Amount, "#,##0.00") 'Amount
        rt.MoveNext
    Wend
    rt.Close
    
    imgPosted.Visible = IIf(rs!Posted = 1, True, False)
    Toolbar1.Buttons(19).Caption = IIf(rs!Posted = 1, "UnPost", " Post ")
    Toolbar1.Buttons(19).Image = IIf(rs!Posted = 1, 11, 10)
    
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    
    iRowEmp = 1
    lstEmployee_Click
    
    lblTotalAmount.Caption = "0.00" 'Format(dTotBal, "#,##0.00")
    'lblTotalEmployee.Caption = Format(dTotDed, "#,##0.00")
    
    SaveSetting App.EXEName, "PersonnelDeductionForPayroll", "PersonnelDeductionForPayroll", rs!Ctrl
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If picAdd.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If AccessRights("Personnel - For Deduction", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
picAdd.ZOrder 0
cmbDivision.Text = ""
cmbDivision.ListIndex = -1
txtPayrollDateAdd.Text = ""
picMain.Enabled = False
picToolbar.Enabled = False
picAdd.Visible = True
cmbDivision.SetFocus
End Sub

Private Sub PRESS_F2()
If picAdd.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    If Statusbar1.Panels(1).Text = "" Then Exit Sub
    If imgPosted.Visible = True Then MsgBox "Already Posted!                         ", vbCritical, "Error...": Exit Sub
    If AccessRights("Personnel - For Deduction", "Edit") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_EDITTING
Else
    If isFocusDed = 0 Then Exit Sub
    If Toolbar1.Buttons(3).Enabled = False Then Exit Sub
    With lstDeductionList.ListItems
        txtDeductionKey.Text = .Item(iRowDed).Text
        cmbDeductionName.Text = .Item(iRowDed).SubItems(1)
        txtPayrollDateSL.Text = .Item(iRowDed).SubItems(3)
        txtAmount.Text = .Item(iRowDed).SubItems(4)
        
        txtAmount1.Text = .Item(iRowDed).SubItems(4)
    End With
    picSLLines.ZOrder 0
    picMain.Enabled = False
    picToolbar.Enabled = False
    picSLLines.Visible = True
    TRANS_DETAIL = is_DET_EDITTING
    txtAmount.SetFocus
End If
End Sub

Private Sub PRESS_DELETE()
If picAdd.Visible = True Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If AccessRights("Personnel - For Deduction", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If MsgBox("ARE YOU SURE IN DELETING THIS TRANSACTION!                   ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_Personnel_Deduction_forPayroll WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "PersonnelDeductionForPayroll", "PersonnelDeductionForPayroll", ""), "is_PAGEDOWN"
If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelDeductionForPayroll", "PersonnelDeductionForPayroll", ""), "is_HOME"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
If picAdd.Visible = True Then Exit Sub
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    sCtrl = ""
    s = "SELECT TOP (1) dbo.tbl_Personnel_Deduction_forPayroll.Ctrl " & _
        " FROM  dbo.tbl_Personnel_Deduction_forPayroll LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Deduction_forPayroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
        " Where (Year(dbo.tbl_Personnel_Compensation_Period.PayrollDate) = " & Format(FormatDateTime(txtPayrollDate.Text, vbShortDate), "yyyy") & ") " & _
        " ORDER BY dbo.tbl_Personnel_Deduction_forPayroll.Ctrl DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        sCtrl = Format(CDbl(rs!Ctrl) + 1, "0000000#")
    Else
        sCtrl = Format(FormatDateTime(txtPayrollDate.Text, vbShortDate), "yyyy") & "0000"
    End If
    rs.Close
    
    Do
        s = "SELECT tbl_Personnel_Deduction_forPayroll.* " & _
            " FROM tbl_Personnel_Deduction_forPayroll " & _
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
    
    
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Deduction_forPayroll " & _
                      " (Ctrl, DivisionKey, PayrollPeriodKey, Remarks, LastModified) " & _
                      " VALUES ('" & sCtrl & "', " & locDivisionKey & ", " & _
                      " " & locPayrollPeriod & ", '', '" & CStr(Now) & " - " & gbl_CompleteName & "')"
    
    iPK = 0
    s = "SELECT tbl_Personnel_Deduction_forPayroll.* " & _
        " FROM tbl_Personnel_Deduction_forPayroll " & _
        " WHERE (Ctrl = '" & sCtrl & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        iPK = rs!PK
    End If
    rs.Close
    
End If
If TRANSACTIONTYPE = is_EDITTING Then
    sCtrl = Trim(txtCtrl.Text)
    iPK = Statusbar1.Panels(1).Text
    
    ConnOmega.Execute "UPDATE tbl_Personnel_Deduction_forPayroll " & _
                      " SET Remarks = '', " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & iPK & ")"
End If

If CDbl(iPK) > 0 Then
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Deduction_forPayroll_Det WHERE (MasterKey = " & iPK & ")"
    With lstEmployee.ListItems
        For i = 1 To .Count
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Deduction_forPayroll_Det " & _
                              " (MasterKey, EmployeeKey) " & _
                              " VALUES (" & iPK & ", " & .Item(i).Text & ")"
        Next i
    End With
    
    With lstGlobalList.ListItems
        For i = 1 To .Count
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Deduction_forPayroll_Det_Det " & _
                              " (MasterKey, EmployeeKey, DeductionKey, SourceKey, Amount, PayrollDate) " & _
                              " VALUES (" & iPK & ", " & .Item(i).Text & ", " & _
                              " " & .Item(i).SubItems(3) & ", " & .Item(i).SubItems(5) & ", " & _
                              " " & CDbl(.Item(i).SubItems(7)) & ", '" & .Item(i).SubItems(6) & "')"
        Next i
    End With
End If

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
If picAdd.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
End Sub

Private Sub PRESS_F8()
If picAdd.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
On Error GoTo PG:
If imgPosted.Visible = False Then
    If AccessRights("Personnel - For Deduction", "Post") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    
    s = "SELECT COUNT(dbo.tbl_Personnel_Payroll.PK) AS RecCnt " & _
        " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK " & _
        " WHERE (dbo.tbl_Personnel_ActionNew.DivisionKey = " & locDivisionKey & ") " & _
        " AND (dbo.tbl_Personnel_Payroll.PayrollPeriodKey = " & locPayrollPeriod & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        If CDbl(rs!RecCnt) > 0 Then
            MsgBox "Please delete all transaction in compensation under this division and payroll date!                 ", vbCritical, "Error..."
            rs.Close
            Exit Sub
        End If
    End If
    rs.Close
    
    If MsgBox("ARE YOU SURE IN POSTING THIS TRANSACTION?                   ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    
    s = "SELECT dbo.tbl_Personnel_Deduction_forPayroll.DivisionKey, dbo.tbl_Personnel_Deduction_forPayroll.PayrollPeriodKey, " & _
        " dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.EmployeeKey, dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.DeductionKey, " & _
        " dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.SourceKey, dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.Amount " & _
        " FROM  dbo.tbl_Personnel_Deduction_forPayroll_Det_Det LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Deduction_forPayroll ON dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.MasterKey = dbo.tbl_Personnel_Deduction_forPayroll.PK " & _
        " WHERE (dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.MasterKey = " & Statusbar1.Panels(1).Text & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Deduction_Payroll " & _
                          " (DivisionKey, PayrollPeriodKey, EmployeeKey, DeductionKey, SourceKey, Amount) " & _
                          " VALUES (" & rs!DivisionKey & ", " & rs!PayrollPeriodKey & ", " & rs!EmployeeKey & ", " & _
                          " " & rs!DeductionKey & ", " & rs!SourceKey & ", " & CDbl(rs!Amount) & ")"
        rs.MoveNext
    Wend
    rs.Close
    
    ConnOmega.Execute "UPDATE tbl_Personnel_Deduction_forPayroll " & _
                      " SET Posted = 1, " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
End If

If imgPosted.Visible = True Then
    If AccessRights("Personnel - For Deduction", "UnPost") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
        
    s = "SELECT COUNT(dbo.tbl_Personnel_Payroll.PK) AS RecCnt " & _
        " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK " & _
        " WHERE (dbo.tbl_Personnel_ActionNew.DivisionKey = " & locDivisionKey & ") " & _
        " AND (dbo.tbl_Personnel_Payroll.PayrollPeriodKey = " & locPayrollPeriod & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        If CDbl(rs!RecCnt) > 0 Then
            MsgBox "Please delete all transaction in compensation under this division and payroll date!                 ", vbCritical, "Error..."
            rs.Close
            Exit Sub
        End If
    End If
    rs.Close
    
    If MsgBox("ARE YOU SURE IN UNPOSTING THIS TRANSACTION?                   ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    
    s = "SELECT dbo.tbl_Personnel_Deduction_forPayroll.DivisionKey, dbo.tbl_Personnel_Deduction_forPayroll.PayrollPeriodKey, " & _
        " dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.EmployeeKey, dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.DeductionKey, " & _
        " dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.SourceKey, dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.Amount " & _
        " FROM  dbo.tbl_Personnel_Deduction_forPayroll_Det_Det LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Deduction_forPayroll ON dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.MasterKey = dbo.tbl_Personnel_Deduction_forPayroll.PK " & _
        " WHERE (dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.MasterKey = " & Statusbar1.Panels(1).Text & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Deduction_Payroll " & _
                          " WHERE (DivisionKey = " & rs!DivisionKey & ") " & _
                          " AND (PayrollPeriodKey = " & rs!PayrollPeriodKey & ")"
        rs.MoveNext
    Wend
    rs.Close
    
    ConnOmega.Execute "UPDATE tbl_Personnel_Deduction_forPayroll " & _
                      " SET Posted = 0, " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    
    
End If
CLEARTEXT
BROWSER GetSetting(App.EXEName, "PersonnelDeductionForPayroll", "PersonnelDeductionForPayroll", ""), "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F9()
If picAdd.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub

With MainForm.CommonDialog1
    .CancelError = True
    On Error GoTo ErrorHandler
    .DialogTitle = "Save"
    .Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
    .ShowSave
    sFileName = Trim(.Filename)
End With

On Error GoTo PG:
WorkbookName = sFileName

Set xlsApp = CreateObject("Excel.Application")
With xlsApp
    .Visible = False
    .Workbooks.Add
    .DisplayAlerts = False
    iWorkSheet = 0
    If .Workbooks(1).Sheets.Count = 3 Then
        .Workbooks(1).Sheets(3).Delete
        .Workbooks(1).Sheets(2).Delete
    End If
    iWorkSheet = iWorkSheet + 1
    .Workbooks(1).Sheets(iWorkSheet).Activate
    .Workbooks(1).Sheets(iWorkSheet).Name = "for Deduction"
    
    RowCnt = 0: ColCnt = 0: RowFreeze = 0: ColFreeze = 0
    RowCnt = RowCnt + 1
    ColCnt = ColCnt + 1
    RowFreeze = RowFreeze + 1
    ColFreeze = ColFreeze + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "RANCHO PALOS VERDES"
    .Range(strRange).Font.Name = "Calibri"
    .Range(strRange).Font.Size = 10
    .Range(strRange).Font.Bold = True
    
    ColCnt = 0
    RowCnt = RowCnt + 1
    ColCnt = ColCnt + 1
    RowFreeze = RowFreeze + 1
    'ColFreeze = ColFreeze + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "for Deduction"
    .Range(strRange).Font.Name = "Calibri"
    .Range(strRange).Font.Size = 10
    .Range(strRange).Font.Bold = False
    
    ColCnt = 0
    RowCnt = RowCnt + 1
    ColCnt = ColCnt + 1
    RowFreeze = RowFreeze + 1
    'ColFreeze = ColFreeze + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Division"
    .Range(strRange).Font.Name = "Calibri"
    .Range(strRange).Font.Size = 10
    .Range(strRange).Font.Bold = False
    
    ColCnt = 0
    RowCnt = RowCnt + 1
    ColCnt = ColCnt + 1
    RowFreeze = RowFreeze + 1
    'ColFreeze = ColFreeze + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Payroll Date"
    .Range(strRange).Font.Name = "Calibri"
    .Range(strRange).Font.Size = 10
    .Range(strRange).Font.Bold = False
    
    ColCnt = 0
    RowCnt = RowCnt + 1
    ColCnt = ColCnt + 1
    RowFreeze = RowFreeze + 1
    'ColFreeze = ColFreeze + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = ""
    .Range(strRange).Font.Name = "Calibri"
    .Range(strRange).Font.Size = 10
    .Range(strRange).Font.Bold = False
    
    ColCnt = 0
    RowCnt = RowCnt + 1
    ColCnt = ColCnt + 1
    RowFreeze = RowFreeze + 1
    ColFreeze = ColFreeze + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "#"
    .Range(strRange).Font.Name = "Calibri"
    .Range(strRange).Font.Size = 10
    .Range(strRange).Font.Bold = True
    
    ColCnt = ColCnt + 1
    ColFreeze = ColFreeze + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Employee Name"
    .Range(strRange).Font.Name = "Calibri"
    .Range(strRange).Font.Size = 10
    .Range(strRange).Font.Bold = True
    
    ColCnt = ColCnt + 1
    'ColFreeze = ColFreeze + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Balance"
    .Range(strRange).Font.Name = "Calibri"
    .Range(strRange).Font.Size = 10
    .Range(strRange).Font.Bold = True
    
    sDeductionKey = "": iDeductionKeyCnt = 0
    s = "SELECT dbo.tbl_Personnel_Payroll_Deductions_Table.Description, dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.DeductionKey " & _
        " FROM  dbo.tbl_Personnel_Deduction_forPayroll_Det_Det INNER JOIN " & _
        " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.DeductionKey = dbo.tbl_Personnel_Payroll_Deductions_Table.PK " & _
        " Where (dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.MasterKey = 10) " & _
        " GROUP BY dbo.tbl_Personnel_Payroll_Deductions_Table.Description, dbo.tbl_Personnel_Payroll_Deductions_Table.Sorting, dbo.tbl_Personnel_Deduction_forPayroll_Det_Det.DeductionKey " & _
        " ORDER BY dbo.tbl_Personnel_Payroll_Deductions_Table.Sorting"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        iDeductionKeyCnt = iDeductionKeyCnt + 1
        sDeductionKey = sDeductionKey & rs!DeductionKey & "|"
        ColCnt = ColCnt + 1
        'ColFreeze = ColFreeze + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = rs!Description
        .Range(strRange).Font.Name = "Calibri"
        .Range(strRange).Font.Size = 10
        .Range(strRange).Font.Bold = True
        rs.MoveNext
    Wend
    rs.Close
    
    sDeductionKey = Mid(sDeductionKey, 1, Len(sDeductionKey) - 1)
    
    cnt = 0
    s = "SELECT dbo.tbl_Personnel_Deduction_forPayroll_Det.EmployeeKey, dbo.tbl_Personnel_IDNumber.IDNumber, " & _
        " dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
        " dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Deduction_forPayroll_Det.Balance " & _
        " FROM  dbo.tbl_Personnel_Deduction_forPayroll_Det LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Deduction_forPayroll_Det.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
        " Where (dbo.tbl_Personnel_Deduction_forPayroll_Det.MasterKey = " & Statusbar1.Panels(1).Text & ") " & _
        " ORDER BY dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_IDNumber.IDNumber"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        cnt = cnt + 1
        ColCnt = 0
        RowCnt = RowCnt + 1
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = cnt
        .Range(strRange).Font.Name = "Calibri"
        .Range(strRange).Font.Size = 10
        .Range(strRange).Font.Bold = False
        
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = rs!IDNumber & " - " & rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
        .Range(strRange).Font.Name = "Calibri"
        .Range(strRange).Font.Size = 10
        .Range(strRange).Font.Bold = False
        
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = rs!Balance
        .Range(strRange).Font.Name = "Calibri"
        .Range(strRange).Font.Size = 10
        .Range(strRange).Font.Bold = False
        .Range(strRange).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        
        If CDbl(iDeductionKeyCnt) = 1 Then
            t = "SELECT SUM(Amount) AS Amount " & _
                " From dbo.tbl_Personnel_Deduction_forPayroll_Det_Det " & _
                " WHERE (MasterKey = " & Statusbar1.Panels(1).Text & ") " & _
                " AND (EmployeeKey = " & rs!EmployeeKey & ") " & _
                " AND (DeductionKey = " & sDeductionKey & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rs.RecordCount > 0 Then
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = IIf(IsNull(rt!Amount), 0, rt!Amount)
                .Range(strRange).Font.Name = "Calibri"
                .Range(strRange).Font.Size = 10
                .Range(strRange).Font.Bold = False
                .Range(strRange).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            End If
            rt.Close
        ElseIf CDbl(iDeductionKeyCnt) > 1 Then
            Array1 = Split(sDeductionKey, "|", -1, 1)
            For i = 0 To UBound(Array1)
                t = "SELECT SUM(Amount) AS Amount " & _
                    " From dbo.tbl_Personnel_Deduction_forPayroll_Det_Det " & _
                    " WHERE (MasterKey = " & Statusbar1.Panels(1).Text & ") " & _
                    " AND (EmployeeKey = " & rs!EmployeeKey & ") " & _
                    " AND (DeductionKey = " & Array1(i) & ")"
                If rt.State = adStateOpen Then rt.Close
                rt.Open t, ConnOmega
                If rs.RecordCount > 0 Then
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = IIf(IsNull(rt!Amount), 0, rt!Amount)
                    .Range(strRange).Font.Name = "Calibri"
                    .Range(strRange).Font.Size = 10
                    .Range(strRange).Font.Bold = False
                    .Range(strRange).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
                End If
                rt.Close
            Next i
        End If
        
        rs.MoveNext
    Wend
    rs.Close
    
    
    RowFreeze = RowFreeze + 1
    strRange = EXCEL_RANGE(ColFreeze, RowFreeze)
    .Range(strRange).Select
    .ActiveWindow.FreezePanes = True
    
    If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
    .ActiveWorkbook.SaveAs Filename:=WorkbookName
    .Visible = True
    Set xlsApp = Nothing
    
End With

Exit Sub
ErrorHandler:
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picAdd.Visible = True Then cmdCancelAdd_Click: Exit Sub
    Unload Me
Else
    If picSLLines.Visible = True Then
        With lstDeductionList.ListItems
            .Item(iRowDed).SubItems(4) = txtAmount1.Text
        End With
        picSLLines.Visible = False
        picMain.Enabled = True
        picToolbar.Enabled = True
        lstDeductionList.SetFocus
        Exit Sub
    End If
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    TRANS_DETAIL = is_DET_REFRESH
    BROWSER GetSetting(App.EXEName, "PersonnelDeductionForPayroll", "PersonnelDeductionForPayroll", ""), "is_LOAD"
    If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelDeductionForPayroll", "PersonnelDeductionForPayroll", ""), "is_HOME"
End If
End Sub

Private Sub CLEARTEXT()
locDivisionKey = 0
locPayrollPeriod = 0
txtEmpKey.Text = ""
txtCtrl.Text = ""
txtDivisionName.Text = ""
txtPayrollDate.Text = ""
txtCutOffDate.Text = ""
lblTotalAmount.Caption = "0.00"
lblTotalEmployee.Caption = "0.00"
imgPosted.Visible = False
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
CLEAR_Details_Emp
CLEAR_Details_Ded
lstGlobalList.ListItems.Clear
End Sub

Private Sub CLEAR_Details_Emp()
lstEmployee.ListItems.Clear
Set x = lstEmployee.ListItems.Add()
x.Text = "0"
x.SubItems(1) = " "
'x.SubItems(2) = " "
End Sub

Private Sub CLEAR_Details_Ded()
lstDeductionList.ListItems.Clear
Set x = lstDeductionList.ListItems.Add()
x.Text = "0"
x.SubItems(1) = " "
x.SubItems(2) = "0"
x.SubItems(3) = " "
x.SubItems(4) = " "
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtDivisionName.Locked = True
txtPayrollDate.Locked = True
txtCutOffDate.Locked = True
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
            .Buttons(1).Enabled = False
            .Buttons(3).Enabled = True
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
            .Buttons(3).ToolTipText = "EDIT (F2)"
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

Private Sub b8TitleBar2_CLoseClick()
cmdCancelAdd_Click
End Sub

Private Sub cmbDivision_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtPayrollDateAdd.SetFocus
End Sub

Private Sub cmdCancelAdd_Click()
picAdd.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True
End Sub

Private Sub cmdOKAdd_Click()
If cmbDivision.ListIndex = -1 Then MsgBox "Please select Division!                       ", vbCritical, "Error....": cmbDivision.SetFocus: Exit Sub
If IsDate(txtPayrollDateAdd.Text) = False Then MsgBox "Please supply a valid date!                  ", vbCritical, "Error...": txtPayrollDateAdd.SetFocus: Exit Sub
locPayrollPeriodTmp = GET_PERIOD_V2(FormatDateTime(txtPayrollDateAdd.Text, vbShortDate), cmbDivision.ItemData(cmbDivision.ListIndex))
If locPayrollPeriodTmp = 0 Then
    MsgBox "Payroll Period Not Match to the Employee Division!      ", vbInformation, ""
    txtPayrollDateAdd.SetFocus
    HTEXT txtPayrollDateAdd
    Exit Sub
End If

s = "SELECT COUNT(dbo.tbl_Personnel_Payroll.PK) AS RecCnt " & _
    " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK " & _
    " WHERE (dbo.tbl_Personnel_ActionNew.DivisionKey = " & cmbDivision.ItemData(cmbDivision.ListIndex) & ") " & _
    " AND (dbo.tbl_Personnel_Payroll.PayrollPeriodKey = " & locPayrollPeriodTmp & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    If CDbl(rs!RecCnt) > 0 Then
        MsgBox "Please delete all transaction in compensation under this division and payroll date!                 " & vbCrLf & _
               "Before adding this transaction!                                                                     ", vbCritical, "Error..."
        rs.Close
        Exit Sub
    End If
End If
rs.Close

CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
locDivisionKey = cmbDivision.ItemData(cmbDivision.ListIndex)
locPayrollPeriod = locPayrollPeriodTmp
iPeriodTerms = Get_Period_Terms(locPayrollPeriod)
txtDivisionName.Text = cmbDivision.List(cmbDivision.ListIndex)
txtPayrollDate.Text = Format(FormatDateTime(txtPayrollDateAdd.Text, vbShortDate), "mm/dd/yyyy")
txtCutOffDate.Text = GET_PERIOD_CUTOFF(locPayrollPeriod)


lstEmployee.ListItems.Clear
lstDeductionList.ListItems.Clear
lstGlobalList.ListItems.Clear

DoEvents

ConnOmega.Execute "DELETE FROM tbl_Personnel_Deduction_forPayroll_Tmp WHERE (LogIn = '" & gbl_UserName & "')"

b = "SELECT dbo.tbl_Personnel_IDNumber.PK as EmployeeKey, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName, " & _
    " ISNULL((SELECT TOP (1) dbo.tbl_Personnel_Position.PositionLevel FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
    " WHERE (dbo.tbl_Personnel_ActionNew.EmpPK = dbo.tbl_Personnel_IDNumber.PK) AND (dbo.tbl_Personnel_ActionNew.EffectivityDate <= '" & FormatDateTime(txtPayrollDateAdd.Text, vbShortDate) & "') " & _
    " ORDER BY dbo.tbl_Personnel_ActionNew.EffectivityDate DESC), 0) AS PositionLevel, " & _
    " ISNULL((SELECT TOP (1) dbo.tbl_Personnel_EmploymentStatus.WithMortuary FROM dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN dbo.tbl_Personnel_EmploymentStatus ON dbo.tbl_Personnel_ActionNew.EmpStatusKey = dbo.tbl_Personnel_EmploymentStatus.PK " & _
    " WHERE (dbo.tbl_Personnel_ActionNew.EmpPK = dbo.tbl_Personnel_IDNumber.PK) AND (dbo.tbl_Personnel_ActionNew.EffectivityDate <= '" & FormatDateTime(txtPayrollDateAdd.Text, vbShortDate) & "') " & _
    " ORDER BY dbo.tbl_Personnel_ActionNew.EffectivityDate DESC),0) as WithMoruary " & _
    " FROM  dbo.tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
    " WHERE (ISNULL((SELECT TOP (1) tbl_Personnel_EmploymentStatus_1.Active FROM dbo.tbl_Personnel_ActionNew AS tbl_Personnel_ActionNew_1 LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_EmploymentStatus AS tbl_Personnel_EmploymentStatus_1 ON tbl_Personnel_ActionNew_1.EmpStatusKey = tbl_Personnel_EmploymentStatus_1.PK " & _
    " WHERE (tbl_Personnel_ActionNew_1.EmpPK = dbo.tbl_Personnel_IDNumber.PK) AND (tbl_Personnel_ActionNew_1.EffectivityDate <= '" & FormatDateTime(txtPayrollDateAdd.Text, vbShortDate) & "') " & _
    " ORDER BY tbl_Personnel_ActionNew_1.EffectivityDate DESC), 0) = 1) " & _
    " AND (ISNULL((SELECT TOP (1) DivisionKey FROM  dbo.tbl_Personnel_ActionNew AS tbl_Personnel_ActionNew_2 WHERE (EmpPK = dbo.tbl_Personnel_IDNumber.PK) AND (EffectivityDate <= '" & FormatDateTime(txtPayrollDateAdd.Text, vbShortDate) & "') " & _
    " ORDER BY EffectivityDate DESC), 0) = " & cmbDivision.ItemData(cmbDivision.ListIndex) & ") " & _
    " ORDER BY dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName"
If rb.State = adStateOpen Then rb.Close
rb.Open b, ConnOmega
While Not rb.EOF

'    MsgBox rb!LastName & ",  " & rb!FirstName & "  " & rb!MiddleName
    
    dTotDed = 0
    If CDbl(rb!WithMoruary) = 1 Then
        
'        MsgBox "pass mortuary"
        
        If CDbl(locDivisionKey) = 1 Then
            sPayDate = Format(FormatDateTime(txtPayrollDateAdd.Text, vbShortDate), "mm/dd/yyyy")
        Else
            If Day(txtPayrollDate.Text) = 12 Then
                sPayDate = Format(DateSerial(Year(FormatDateTime(txtPayrollDateAdd.Text, vbShortDate)), Month(FormatDateTime(txtPayrollDateAdd.Text, vbShortDate)), 15), "mm/dd/yyyy")
            Else
                sPayDate = Format(DateSerial(Year(FormatDateTime(txtPayrollDateAdd.Text, vbShortDate)), Month(FormatDateTime(txtPayrollDateAdd.Text, vbShortDate)) + 1, 0), "mm/dd/yyyy")
            End If
        End If
        
        s = "SELECT dbo.tbl_Personnel_Mortuary_Det_Det.Amount " & _
            " FROM  dbo.tbl_Personnel_Mortuary_Det_Det LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Mortuary ON dbo.tbl_Personnel_Mortuary_Det_Det.MasterKey = dbo.tbl_Personnel_Mortuary.PK " & _
            " WHERE (dbo.tbl_Personnel_Mortuary_Det_Det.PayrollDate = '" & FormatDateTime(sPayDate, vbShortDate) & "') " & _
            " AND (dbo.tbl_Personnel_Mortuary_Det_Det.PositionLevelKey = " & rb!PositionLevel & ") " & _
            " AND (dbo.tbl_Personnel_Mortuary.Posted = 1) " & _
            " AND (dbo.tbl_Personnel_Mortuary_Det_Det.Amount > 0)"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
        
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Deduction_forPayroll_Tmp " & _
                              " (LogIn, EmployeeKey, EmployeeName, EmployeeID, Balance, DeductionKey, DeductionName, SourceKey, PayrollDate, Amount) " & _
                              " VALUES ('" & gbl_UserName & "', " & rb!EmployeeKey & ", '" & FORMATSQL(rb!LastName & ",  " & rb!FirstName & "  " & rb!MiddleName) & "', '" & rb!IDNumber & "', " & _
                              " 0, 13, 'Mortuary', 0, '" & Format(FormatDateTime(txtPayrollDateAdd.Text, vbShortDate), "mm/dd/yyyy") & "', " & _
                              " " & CDbl(Format(rs!Amount, "#,##0.00")) & ")"
        End If
        rs.Close
    End If
    
    s = "SELECT ISNULL(SUM(Balance),0) AS Balance " & _
        " From dbo.tbl_Personnel_Deduction_SL " & _
        " Where (EmployeeKey = " & rb!EmployeeKey & ") " & _
        " HAVING (ISNULL(SUM(Balance),0) > 0)"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        t = "SELECT dbo.tbl_Personnel_Deduction_SL.DeductionKey, dbo.tbl_Personnel_Payroll_Deductions_Table.Description, " & _
            " dbo.tbl_Personnel_Deduction_SL.SourceKey, dbo.tbl_Personnel_Compensation_Period.PayrollDate, " & _
            " ROUND(ISNULL(SUM(dbo.tbl_Personnel_Deduction_SL.Balance),0),2) AS Balance " & _
            " FROM  dbo.tbl_Personnel_Deduction RIGHT OUTER JOIN " & _
            " dbo.tbl_Personnel_Deduction_SL LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Deduction_SL.DeductionKey = dbo.tbl_Personnel_Payroll_Deductions_Table.PK ON dbo.tbl_Personnel_Deduction.PK = dbo.tbl_Personnel_Deduction_SL.SourceKey LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Deduction.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
            " WHERE (dbo.tbl_Personnel_Deduction_SL.EmployeeKey = " & rb!EmployeeKey & ") " & _
            " AND (dbo.tbl_Personnel_Deduction_SL.TransactionDate <= '" & FormatDateTime(txtPayrollDateAdd.Text, vbShortDate) & "') " & _
            " GROUP BY dbo.tbl_Personnel_Deduction_SL.DeductionKey, dbo.tbl_Personnel_Payroll_Deductions_Table.Description, dbo.tbl_Personnel_Deduction_SL.SourceKey, dbo.tbl_Personnel_Compensation_Period.PayrollDate, dbo.tbl_Personnel_Payroll_Deductions_Table.Sorting " & _
            " HAVING (ROUND(ISNULL(SUM(dbo.tbl_Personnel_Deduction_SL.Balance),0),2) > 0) " & _
            " ORDER BY dbo.tbl_Personnel_Payroll_Deductions_Table.Sorting, dbo.tbl_Personnel_Compensation_Period.PayrollDate"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        While Not rt.EOF
            
            dDedPerPay = 0
            'u = "SELECT DedPerPayroll " & _
                " From dbo.tbl_Personnel_Deduction_Details " & _
                " WHERE (MasterKey = " & rt!SourceKey & ") " & _
                " AND (DeductionKey = " & rt!DeductionKey & ")"
            u = "SELECT DedPerPayroll " & _
                " From dbo.tbl_Personnel_Deduction_Details " & _
                " WHERE ((MasterKey = " & rt!SourceKey & ") " & _
                " AND (DeductionKey = " & rt!DeductionKey & ") " & _
                " AND (DeductionPeriodKey = 0)) OR " & _
                " ((MasterKey = " & rt!SourceKey & ") " & _
                " AND (DeductionKey = " & rt!DeductionKey & ") " & _
                " AND (DeductionPeriodKey = " & iPeriodTerms & "))"
            If ru.State = adStateOpen Then ru.Close
            ru.Open u, ConnOmega
            If ru.RecordCount > 0 Then
                dDedPerPay = Format(ru!DedPerPayroll, "#,##0.00")
            End If
            ru.Close
            
            If CDbl(rt!Balance) < CDbl(dDedPerPay) Then
                dDedPerPay = CDbl(rt!Balance)
            End If
            
            If CDbl(dDedPerPay) <> 0 Then
                ConnOmega.Execute "INSERT INTO tbl_Personnel_Deduction_forPayroll_Tmp " & _
                                  " (LogIn, EmployeeKey, EmployeeName, EmployeeID, Balance, DeductionKey, DeductionName, SourceKey, PayrollDate, Amount) " & _
                                  " VALUES ('" & gbl_UserName & "', " & rb!EmployeeKey & ", '" & FORMATSQL(rb!LastName & ",  " & rb!FirstName & "  " & rb!MiddleName) & "', '" & rb!IDNumber & "', " & _
                                  " " & CDbl(Format(rs!Balance, "#,##0.00")) & ", " & rt!DeductionKey & ", '" & FORMATSQL(rt!Description) & "', " & rt!SourceKey & ", " & _
                                  " '" & Format(FormatDateTime(rt!PayrollDate, vbShortDate), "mm/dd/yyyy") & "', " & _
                                  " " & CDbl(dDedPerPay) & ")"
            End If
            
            rt.MoveNext
        Wend
        rt.Close
    End If
    rs.Close
    rb.MoveNext
Wend
rb.Close

s = "SELECT EmployeeKey, EmployeeID, EmployeeName, SUM(Balance) AS Balance " & _
    " From dbo.tbl_Personnel_Deduction_forPayroll_Tmp " & _
    " WHERE (LogIn = '" & gbl_UserName & "') " & _
    " GROUP BY EmployeeKey, EmployeeName, EmployeeID " & _
    " ORDER BY EmployeeName, EmployeeID"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    Set x = lstEmployee.ListItems.Add()
    x.Text = rs!EmployeeKey
    x.SubItems(1) = rs!EmployeeID & " - " & rs!EmployeeName
    'x.SubItems(2) = " " 'Format(rs!Balance, "#,##0.00")
    rs.MoveNext
Wend
rs.Close

s = "SELECT EmployeeKey, EmployeeID, EmployeeName, Balance, DeductionKey, DeductionName, SourceKey, PayrollDate, Amount " & _
    " From dbo.tbl_Personnel_Deduction_forPayroll_Tmp " & _
    " WHERE (LogIn = '" & gbl_UserName & "') " & _
    " ORDER BY EmployeeName, EmployeeID"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    Set y = lstGlobalList.ListItems.Add()
    y.Text = rs!EmployeeKey
    y.SubItems(1) = rs!EmployeeID & " - " & rs!EmployeeName
    y.SubItems(2) = Format(rs!Balance, "#,##0.00")
    y.SubItems(3) = rs!DeductionKey 'DeductionKey
    y.SubItems(4) = rs!DeductionName 'DeductionName
    y.SubItems(5) = rs!SourceKey 'SourceKey
    y.SubItems(6) = Format(rs!PayrollDate, "mm/dd/yyyy") 'PayrollDate
    y.SubItems(7) = Format(rs!Amount, "#,##0.00")
    rs.MoveNext
Wend
rs.Close

cmdCancelAdd_Click
lblTotalAmount.Caption = "0.00" 'Format(dTotBal, "#,##0.00")
iRowEmp = 1
lstEmployee_Click
End Sub

'Private Function CheckDuplicateEmp(iEmp) As Boolean
'CheckDuplicateEmp = False
'With lstEmployee.ListItems
'    For i = 1 To .Count
'        If CDbl(iEmp) = CDbl(.Item(i).Text) Then
'            CheckDuplicateEmp = True
'            Exit For
'        End If
'    Next i
'End With
'End Function

'Private Function GetListEmployeeIndex(iEmp) As Integer
'GetListEmployeeIndex = 0
'With lstEmployee.ListItems
'    For i = 1 To .Count
'        If CDbl(iEmp) = CDbl(.Item(i).Text) Then
'            GetListEmployeeIndex = i
'            Exit For
'        End If
'    Next i
'End With
'End Function

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
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "PersonnelDeductionForPayroll", "PersonnelDeductionForPayroll", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "PersonnelDeductionForPayroll", "PersonnelDeductionForPayroll", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "PersonnelDeductionForPayroll", "PersonnelDeductionForPayroll", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "PersonnelDeductionForPayroll", "PersonnelDeductionForPayroll", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 3
Me.Left = (MainForm.ScaleWidth - Me.Width) / 3
POPULATE_COMBO "PK", "Description", "tbl_Personnel_Division", "Description", cmbDivision
'POPULATE_COMBO_EXEMPTION "PK", "Description", "tbl_Personnel_Payroll_Deductions_Table", "Sorting", "ViewInDeductionModule", 1, cmbDeductionName
isFocusEmp = 0
iRowEmp = 0
isFocusDed = 0
iRowDed = 0
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANS_DETAIL = is_DET_REFRESH
BROWSER GetSetting(App.EXEName, "PersonnelDeductionForPayroll", "PersonnelDeductionForPayroll", ""), "is_LOAD"
If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelDeductionForPayroll", "PersonnelDeductionForPayroll", ""), "is_HOME"
End Sub

Private Sub lstDeductionList_Click()
isFocusDed = 1
iRowDed = lstDeductionList.SelectedItem.Index
TRANS_DETAIL = is_DET_REFRESH
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If CDbl(lstDeductionList.ListItems.Item(iRowDed).Text) <> 0 Then
        TOOLBARFUNC 5
    Else
        TOOLBARFUNC 4
    End If
End If
End Sub

Private Sub lstDeductionList_GotFocus()
isFocusDed = 1
iRowDed = lstDeductionList.SelectedItem.Index
TRANS_DETAIL = is_DET_REFRESH
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If CDbl(lstDeductionList.ListItems.Item(iRowDed).Text) <> 0 Then
        TOOLBARFUNC 5
    Else
        TOOLBARFUNC 4
    End If
End If
End Sub

Private Sub lstDeductionList_ItemClick(ByVal Item As MSComctlLib.ListItem)
iRowDed = lstDeductionList.SelectedItem.Index
End Sub

Private Sub lstDeductionList_LostFocus()
isFocusDed = 0
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    TOOLBARFUNC 2
End If
End Sub

Private Sub lstEmployee_Click()
isFocusEmp = 1
iRowEmp = lstEmployee.SelectedItem.Index
CLEAR_Details_Ded
TRANS_DETAIL = is_DET_REFRESH
dTotDed = 0
If CDbl(lstEmployee.ListItems.Item(iRowEmp).Text) <> 0 Then
    iLineCnt = 0: txtEmpKey.Text = lstEmployee.ListItems.Item(iRowEmp).Text
    For i = 1 To lstGlobalList.ListItems.Count
        If CDbl(lstEmployee.ListItems.Item(iRowEmp).Text) = CDbl(lstGlobalList.ListItems.Item(i).Text) Then
            iLineCnt = iLineCnt + 1
        End If
    Next i
    
    If CDbl(iLineCnt) > 0 Then
        lstDeductionList.ListItems.Clear
        For i = 1 To lstGlobalList.ListItems.Count
            If CDbl(lstEmployee.ListItems.Item(iRowEmp).Text) = CDbl(lstGlobalList.ListItems.Item(i).Text) Then
                Set x = lstDeductionList.ListItems.Add()
                x.Text = lstGlobalList.ListItems.Item(i).SubItems(3)
                x.SubItems(1) = lstGlobalList.ListItems.Item(i).SubItems(4)
                x.SubItems(2) = lstGlobalList.ListItems.Item(i).SubItems(5)
                x.SubItems(3) = lstGlobalList.ListItems.Item(i).SubItems(6)
                x.SubItems(4) = lstGlobalList.ListItems.Item(i).SubItems(7)
                dTotDed = dTotDed + CDbl(IIf(IsNumeric(lstGlobalList.ListItems.Item(i).SubItems(7)) = False, 0, lstGlobalList.ListItems.Item(i).SubItems(7)))
            End If
        Next i
    End If
End If
lblTotalEmployee.Caption = Format(dTotDed, "#,##0.00")
End Sub

Private Sub lstEmployee_GotFocus()
isFocusEmp = 1
iRowEmp = lstEmployee.SelectedItem.Index
CLEAR_Details_Ded
TRANS_DETAIL = is_DET_REFRESH
dTotDed = 0
If CDbl(lstEmployee.ListItems.Item(iRowEmp).Text) <> 0 Then
    iLineCnt = 0: txtEmpKey.Text = lstEmployee.ListItems.Item(iRowEmp).Text
    For i = 1 To lstGlobalList.ListItems.Count
        If CDbl(lstEmployee.ListItems.Item(iRowEmp).Text) = CDbl(lstGlobalList.ListItems.Item(i).Text) Then
            iLineCnt = iLineCnt + 1
        End If
    Next i
    
    If CDbl(iLineCnt) > 0 Then
        lstDeductionList.ListItems.Clear
        For i = 1 To lstGlobalList.ListItems.Count
            If CDbl(lstEmployee.ListItems.Item(iRowEmp).Text) = CDbl(lstGlobalList.ListItems.Item(i).Text) Then
                Set x = lstDeductionList.ListItems.Add()
                x.Text = lstGlobalList.ListItems.Item(i).SubItems(3)
                x.SubItems(1) = lstGlobalList.ListItems.Item(i).SubItems(4)
                x.SubItems(2) = lstGlobalList.ListItems.Item(i).SubItems(5)
                x.SubItems(3) = lstGlobalList.ListItems.Item(i).SubItems(6)
                x.SubItems(4) = lstGlobalList.ListItems.Item(i).SubItems(7)
                dTotDed = dTotDed + CDbl(IIf(IsNumeric(lstGlobalList.ListItems.Item(i).SubItems(7)) = False, 0, lstGlobalList.ListItems.Item(i).SubItems(7)))
            End If
        Next i
    End If
End If
lblTotalEmployee.Caption = Format(dTotDed, "#,##0.00")
End Sub

Private Sub lstEmployee_ItemClick(ByVal Item As MSComctlLib.ListItem)
iRowEmp = lstEmployee.SelectedItem.Index
iRowEmp = lstEmployee.SelectedItem.Index
CLEAR_Details_Ded
dTotDed = 0
If CDbl(lstEmployee.ListItems.Item(iRowEmp).Text) <> 0 Then
    iLineCnt = 0: txtEmpKey.Text = lstEmployee.ListItems.Item(iRowEmp).Text
    For i = 1 To lstGlobalList.ListItems.Count
        If CDbl(lstEmployee.ListItems.Item(iRowEmp).Text) = CDbl(lstGlobalList.ListItems.Item(i).Text) Then
            iLineCnt = iLineCnt + 1
        End If
    Next i
    
    If CDbl(iLineCnt) > 0 Then
        lstDeductionList.ListItems.Clear
        For i = 1 To lstGlobalList.ListItems.Count
            If CDbl(lstEmployee.ListItems.Item(iRowEmp).Text) = CDbl(lstGlobalList.ListItems.Item(i).Text) Then
                Set x = lstDeductionList.ListItems.Add()
                x.Text = lstGlobalList.ListItems.Item(i).SubItems(3)
                x.SubItems(1) = lstGlobalList.ListItems.Item(i).SubItems(4)
                x.SubItems(2) = lstGlobalList.ListItems.Item(i).SubItems(5)
                x.SubItems(3) = lstGlobalList.ListItems.Item(i).SubItems(6)
                x.SubItems(4) = lstGlobalList.ListItems.Item(i).SubItems(7)
                dTotDed = dTotDed + CDbl(IIf(IsNumeric(lstGlobalList.ListItems.Item(i).SubItems(7)) = False, 0, lstGlobalList.ListItems.Item(i).SubItems(7)))
            End If
        Next i
    End If
End If
lblTotalEmployee.Caption = Format(dTotDed, "#,##0.00")
End Sub

Private Sub lstEmployee_LostFocus()
isFocusEmp = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "PersonnelDeductionForPayroll", "PersonnelDeductionForPayroll", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "PersonnelDeductionForPayroll", "PersonnelDeductionForPayroll", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "PersonnelDeductionForPayroll", "PersonnelDeductionForPayroll", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "PersonnelDeductionForPayroll", "PersonnelDeductionForPayroll", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Post":    PRESS_F8
    Case "Print":   PRESS_F9
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtAmount_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    dTotDed = 0
    With lstDeductionList.ListItems
        .Item(iRowDed).SubItems(4) = Format(RETURNTEXTVALUE(txtAmount), "#,##0.00")
        For i = 1 To .Count
            dTotDed = dTotDed + CDbl(IIf(IsNumeric(.Item(i).SubItems(4)) = False, 0, .Item(i).SubItems(4)))
        Next i
    End With
    
    With lstGlobalList.ListItems
        For i = 1 To .Count
            If RETURNTEXTVALUE(txtEmpKey) = CDbl(.Item(i).Text) Then
                If CDbl(lstDeductionList.ListItems.Item(iRowDed).Text) = CDbl(.Item(i).SubItems(3)) Then
                    If CDbl(lstDeductionList.ListItems.Item(iRowDed).SubItems(2)) = CDbl(.Item(i).SubItems(5)) Then
                        If DateValue(lstDeductionList.ListItems.Item(iRowDed).SubItems(3)) = DateValue(.Item(i).SubItems(6)) Then
                            .Item(i).SubItems(7) = RETURNTEXTVALUE(txtAmount)
                        End If
                    End If
                End If
            End If
        Next i
    End With
    lblTotalEmployee.Caption = Format(dTotDed, "#,##0.00")
End If
End Sub

Private Sub txtAmount_GotFocus()
HTEXT txtAmount
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    picSLLines.Visible = False
    picMain.Enabled = True
    picToolbar.Enabled = True
    lstDeductionList.SetFocus
End If
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtPayrollDateAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAdd_Click
End Sub

