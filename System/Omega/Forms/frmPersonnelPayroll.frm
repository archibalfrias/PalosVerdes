VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelPayroll 
   Appearance      =   0  'Flat
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7125
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
   ScaleHeight     =   7125
   ScaleWidth      =   11490
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   600
      ScaleHeight     =   4935
      ScaleWidth      =   10215
      TabIndex        =   4
      Top             =   1440
      Width           =   10215
      Begin VB.TextBox txtCutOffDate 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "12/15/2017 - 12/31/2017"
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   0
         Width           =   7095
      End
      Begin VB.TextBox txtPayrollDate 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "12/15/2017 - 12/31/2017"
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtDepartment 
         Height          =   315
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtDivision 
         Height          =   315
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox txtPosition 
         Height          =   315
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C6B8A4&
         Caption         =   "Earnings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   0
         TabIndex        =   7
         Top             =   1560
         Width           =   5415
         Begin MSComctlLib.ListView lstEarnings 
            Height          =   2415
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   4260
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
               Text            =   "EarningKey"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   4586
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "# of Hours"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Amount"
               Object.Width           =   1940
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C6B8A4&
         Caption         =   "Deductions"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   5520
         TabIndex        =   5
         Top             =   1560
         Width           =   4695
         Begin MSComctlLib.ListView lstDeductions 
            Height          =   2415
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   4260
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
               Object.Width           =   4587
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Amount"
               Object.Width           =   2646
            EndProperty
         End
      End
      Begin VB.Label lblLocked 
         BackStyle       =   0  'Transparent
         Caption         =   "L O C K E D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1320
         TabIndex        =   35
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Cut-Off Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   34
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblNetPay 
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
         Left            =   8520
         TabIndex        =   24
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Pay >>"
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
         TabIndex        =   23
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label lblTotDeductions 
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
         Left            =   8520
         TabIndex        =   22
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Deductions >>"
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
         TabIndex        =   21
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label lblTotEarnings 
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
         Left            =   3720
         TabIndex        =   20
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Earnings >>"
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
         Left            =   2280
         TabIndex        =   19
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   14
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
         MouseIcon       =   "frmPersonnelPayroll.frx":0000
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
               Picture         =   "frmPersonnelPayroll.frx":031A
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11880
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
            Picture         =   "frmPersonnelPayroll.frx":0A2D
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPayroll.frx":1707
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPayroll.frx":23E1
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPayroll.frx":30BB
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPayroll.frx":3D95
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPayroll.frx":4A6F
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPayroll.frx":5749
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPayroll.frx":6423
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPayroll.frx":70FD
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPayroll.frx":79D7
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPayroll.frx":86B1
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPayroll.frx":938B
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPayroll.frx":A065
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPayroll.frx":AD3F
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPayroll.frx":BA19
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   6825
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
   Begin RPVGCC.b8Container picSearch 
      Height          =   4455
      Left            =   3360
      TabIndex        =   25
      Top             =   1080
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   7858
      BackColor       =   15396057
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
         Picture         =   "frmPersonnelPayroll.frx":C6F3
         Style           =   1  'Graphical
         TabIndex        =   30
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
         Left            =   2640
         Picture         =   "frmPersonnelPayroll.frx":CD65
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3840
         Width           =   1560
      End
      Begin VB.TextBox txtSearchSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   4815
      End
      Begin VB.ListBox lstResultSearch 
         Height          =   2595
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   4815
      End
      Begin VB.ComboBox cmbPayrollPeriodSearch 
         Height          =   315
         ItemData        =   "frmPersonnelPayroll.frx":D4C1
         Left            =   1440
         List            =   "frmPersonnelPayroll.frx":D4C3
         TabIndex        =   26
         Top             =   3480
         Width           =   3495
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   45
         TabIndex        =   31
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
         Icon            =   "frmPersonnelPayroll.frx":D4C5
         ShadowVisible   =   0   'False
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Period - Division"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   3480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmPersonnelPayroll"
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

Dim x, iPK, dTotEarnings, dTotDeductions, dNetPay, iDivKey

Private Sub PRESS_INSERT()


End Sub

Private Sub PRESS_DELETE()

End Sub

Private Sub PRESS_F5()

End Sub

Private Sub PRESS_F6()
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
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub

End Sub

Private Sub PRESS_F9()
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1) = "" Then Exit Sub
Screen.MousePointer = vbHourglass
GeneratePayslipSignLedger gbl_UserName, 1, Statusbar1.Panels(1).Text, txtPayrollDate.Text, 2, iDivKey
frmCrystalReportViewer.PRINT_PAYROLL_PAYSLIP_V3 gbl_UserName
If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
Screen.MousePointer = vbDefault
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSearch.Visible = True Then cmdCancelSearch_Click: Exit Sub
    Unload Me
Else
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "PersonnelPayroll", "PersonnelPayroll", ""), "is_LOAD"
    If Trim(txtName.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelPayroll", "PersonnelPayroll", ""), "is_HOME"
End If
End Sub

Public Sub BROWSER(Ctrl, isAction)
Select Case isAction
    Case "is_LOAD"
        If Ctrl <> "" Then
            s = "SELECT TOP (1) dbo.tbl_Personnel_Payroll.PK, dbo.tbl_Personnel_Payroll.Ctrl, " & _
                " dbo.tbl_Personnel_Payroll.EmployeeKey, dbo.tbl_Personnel_Payroll.PayrollPeriodKey, " & _
                " dbo.tbl_Personnel_Payroll.ActionMemoKey, dbo.tbl_Personnel_Payroll.Locked, " & _
                " dbo.tbl_Personnel_Payroll.LastModified, dbo.tbl_Personnel_IDNumber.IDNumber, " & _
                " dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
                " dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Division.Description AS Division, " & _
                " dbo.tbl_Personnel_Department.DepartmentName, dbo.tbl_Personnel_Position.PositionName, " & _
                " dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo, " & _
                " dbo.tbl_Personnel_Compensation_Period.PayrollDate, dbo.tbl_Personnel_ActionNew.DivisionKey " & _
                " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Payroll.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Position RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_Division RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Division.PK = dbo.tbl_Personnel_ActionNew.DivisionKey LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK ON dbo.tbl_Personnel_Position.PK = dbo.tbl_Personnel_ActionNew.PositionsKey ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK " & _
                " WHERE (dbo.tbl_Personnel_Payroll.Ctrl = '" & Ctrl & "') " & _
                " AND ((SELECT COUNT(*) AS RecCnt From dbo.tbl_Personnel_Payroll_Earnings Where (dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK)) > 0) " & _
                " ORDER BY dbo.tbl_Personnel_Payroll.Ctrl"
        Else
            s = "SELECT TOP (1) dbo.tbl_Personnel_Payroll.PK, dbo.tbl_Personnel_Payroll.Ctrl, " & _
                " dbo.tbl_Personnel_Payroll.EmployeeKey, dbo.tbl_Personnel_Payroll.PayrollPeriodKey, " & _
                " dbo.tbl_Personnel_Payroll.ActionMemoKey, dbo.tbl_Personnel_Payroll.Locked, " & _
                " dbo.tbl_Personnel_Payroll.LastModified, dbo.tbl_Personnel_IDNumber.IDNumber, " & _
                " dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
                " dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Division.Description AS Division, " & _
                " dbo.tbl_Personnel_Department.DepartmentName, dbo.tbl_Personnel_Position.PositionName, " & _
                " dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo, " & _
                " dbo.tbl_Personnel_Compensation_Period.PayrollDate, dbo.tbl_Personnel_ActionNew.DivisionKey " & _
                " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Payroll.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Position RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_Division RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Division.PK = dbo.tbl_Personnel_ActionNew.DivisionKey LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK ON dbo.tbl_Personnel_Position.PK = dbo.tbl_Personnel_ActionNew.PositionsKey ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK " & _
                " WHERE ((SELECT COUNT(*) AS RecCnt From dbo.tbl_Personnel_Payroll_Earnings Where (dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK)) > 0) " & _
                " ORDER BY dbo.tbl_Personnel_Payroll.Ctrl"
        End If
    Case "is_HOME"
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Payroll.PK, dbo.tbl_Personnel_Payroll.Ctrl, " & _
            " dbo.tbl_Personnel_Payroll.EmployeeKey, dbo.tbl_Personnel_Payroll.PayrollPeriodKey, " & _
            " dbo.tbl_Personnel_Payroll.ActionMemoKey, dbo.tbl_Personnel_Payroll.Locked, " & _
            " dbo.tbl_Personnel_Payroll.LastModified, dbo.tbl_Personnel_IDNumber.IDNumber, " & _
            " dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
            " dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Division.Description AS Division, " & _
            " dbo.tbl_Personnel_Department.DepartmentName, dbo.tbl_Personnel_Position.PositionName, " & _
            " dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo, " & _
            " dbo.tbl_Personnel_Compensation_Period.PayrollDate, dbo.tbl_Personnel_ActionNew.DivisionKey " & _
            " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Payroll.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Position RIGHT OUTER JOIN " & _
            " dbo.tbl_Personnel_Division RIGHT OUTER JOIN " & _
            " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Division.PK = dbo.tbl_Personnel_ActionNew.DivisionKey LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK ON dbo.tbl_Personnel_Position.PK = dbo.tbl_Personnel_ActionNew.PositionsKey ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK " & _
            " WHERE ((SELECT COUNT(*) AS RecCnt From dbo.tbl_Personnel_Payroll_Earnings Where (dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK)) > 0) " & _
            " ORDER BY dbo.tbl_Personnel_Payroll.Ctrl"
    Case "is_PAGEUP"
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Payroll.PK, dbo.tbl_Personnel_Payroll.Ctrl, " & _
                " dbo.tbl_Personnel_Payroll.EmployeeKey, dbo.tbl_Personnel_Payroll.PayrollPeriodKey, " & _
                " dbo.tbl_Personnel_Payroll.ActionMemoKey, dbo.tbl_Personnel_Payroll.Locked, " & _
                " dbo.tbl_Personnel_Payroll.LastModified, dbo.tbl_Personnel_IDNumber.IDNumber, " & _
                " dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
                " dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Division.Description AS Division, " & _
                " dbo.tbl_Personnel_Department.DepartmentName, dbo.tbl_Personnel_Position.PositionName, " & _
                " dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo, " & _
                " dbo.tbl_Personnel_Compensation_Period.PayrollDate, dbo.tbl_Personnel_ActionNew.DivisionKey " & _
                " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Payroll.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Position RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_Division RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Division.PK = dbo.tbl_Personnel_ActionNew.DivisionKey LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK ON dbo.tbl_Personnel_Position.PK = dbo.tbl_Personnel_ActionNew.PositionsKey ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK " & _
                " WHERE (dbo.tbl_Personnel_Payroll.Ctrl < '" & Ctrl & "') " & _
                " AND ((SELECT COUNT(*) AS RecCnt From dbo.tbl_Personnel_Payroll_Earnings Where (dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK)) > 0) " & _
                " ORDER BY dbo.tbl_Personnel_Payroll.Ctrl DESC"
    Case "is_PAGEDOWN"
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Payroll.PK, dbo.tbl_Personnel_Payroll.Ctrl, " & _
                " dbo.tbl_Personnel_Payroll.EmployeeKey, dbo.tbl_Personnel_Payroll.PayrollPeriodKey, " & _
                " dbo.tbl_Personnel_Payroll.ActionMemoKey, dbo.tbl_Personnel_Payroll.Locked, " & _
                " dbo.tbl_Personnel_Payroll.LastModified, dbo.tbl_Personnel_IDNumber.IDNumber, " & _
                " dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
                " dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Division.Description AS Division, " & _
                " dbo.tbl_Personnel_Department.DepartmentName, dbo.tbl_Personnel_Position.PositionName, " & _
                " dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo, " & _
                " dbo.tbl_Personnel_Compensation_Period.PayrollDate, dbo.tbl_Personnel_ActionNew.DivisionKey " & _
                " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Payroll.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Position RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_Division RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Division.PK = dbo.tbl_Personnel_ActionNew.DivisionKey LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK ON dbo.tbl_Personnel_Position.PK = dbo.tbl_Personnel_ActionNew.PositionsKey ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK " & _
                " WHERE (dbo.tbl_Personnel_Payroll.Ctrl > '" & Ctrl & "') " & _
                " AND ((SELECT COUNT(*) AS RecCnt From dbo.tbl_Personnel_Payroll_Earnings Where (dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK)) > 0) " & _
                " ORDER BY dbo.tbl_Personnel_Payroll.Ctrl "
    Case "is_END"
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Payroll.PK, dbo.tbl_Personnel_Payroll.Ctrl, " & _
                " dbo.tbl_Personnel_Payroll.EmployeeKey, dbo.tbl_Personnel_Payroll.PayrollPeriodKey, " & _
                " dbo.tbl_Personnel_Payroll.ActionMemoKey, dbo.tbl_Personnel_Payroll.Locked, " & _
                " dbo.tbl_Personnel_Payroll.LastModified, dbo.tbl_Personnel_IDNumber.IDNumber, " & _
                " dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
                " dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Division.Description AS Division, " & _
                " dbo.tbl_Personnel_Department.DepartmentName, dbo.tbl_Personnel_Position.PositionName, " & _
                " dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo, " & _
                " dbo.tbl_Personnel_Compensation_Period.PayrollDate, dbo.tbl_Personnel_ActionNew.DivisionKey " & _
                " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Payroll.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Position RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_Division RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Division.PK = dbo.tbl_Personnel_ActionNew.DivisionKey LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK ON dbo.tbl_Personnel_Position.PK = dbo.tbl_Personnel_ActionNew.PositionsKey ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK " & _
                " WHERE ((SELECT COUNT(*) AS RecCnt From dbo.tbl_Personnel_Payroll_Earnings Where (dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK)) > 0) " & _
                " ORDER BY dbo.tbl_Personnel_Payroll.Ctrl DESC"
    Case "is_FIND"
        s = "SELECT TOP (1) dbo.tbl_Personnel_Payroll.PK, dbo.tbl_Personnel_Payroll.Ctrl, " & _
            " dbo.tbl_Personnel_Payroll.EmployeeKey, dbo.tbl_Personnel_Payroll.PayrollPeriodKey, " & _
            " dbo.tbl_Personnel_Payroll.ActionMemoKey, dbo.tbl_Personnel_Payroll.Locked, " & _
            " dbo.tbl_Personnel_Payroll.LastModified, dbo.tbl_Personnel_IDNumber.IDNumber, " & _
            " dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
            " dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Division.Description AS Division, " & _
            " dbo.tbl_Personnel_Department.DepartmentName, dbo.tbl_Personnel_Position.PositionName, " & _
            " dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo, " & _
            " dbo.tbl_Personnel_Compensation_Period.PayrollDate, dbo.tbl_Personnel_ActionNew.DivisionKey " & _
            " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Payroll.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Position RIGHT OUTER JOIN " & _
            " dbo.tbl_Personnel_Division RIGHT OUTER JOIN " & _
            " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Division.PK = dbo.tbl_Personnel_ActionNew.DivisionKey LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK ON dbo.tbl_Personnel_Position.PK = dbo.tbl_Personnel_ActionNew.PositionsKey ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK " & _
            " WHERE (dbo.tbl_Personnel_Payroll.PK = " & Ctrl & ") " & _
            " AND ((SELECT COUNT(*) AS RecCnt From dbo.tbl_Personnel_Payroll_Earnings Where (dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK)) > 0) " & _
            " ORDER BY dbo.tbl_Personnel_Payroll.Ctrl DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iDivKey = rs!DivisionKey
    txtName.Text = rs!IDNumber & " - " & rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
    txtPayrollDate.Text = Format(rs!PayrollDate, "mm/dd/yyyy")
    txtCutOffDate.Text = Format(rs!DateFrom, "mm/dd/yyyy") & " - " & Format(rs!DateTo, "mm/dd/yyyy")
    txtDivision.Text = rs!Division
    txtDepartment.Text = rs!DepartmentName
    txtPosition.Text = rs!PositionName
    
    dTotEarnings = 0: dTotDeductions = 0: dNetPay = 0
    CLEAR_DETAILS_Earnings
    t = "SELECT dbo.tbl_Personnel_Payroll_Earnings.EarningKey, dbo.tbl_Personnel_Payroll_Earnings_Table.Description, " & _
        " dbo.tbl_Personnel_Payroll_Earnings.Hours, dbo.tbl_Personnel_Payroll_Earnings.TotalAmount, " & _
        " dbo.tbl_Personnel_Payroll_Earnings.Remarks, dbo.tbl_Personnel_Payroll_Earnings.EarningKey " & _
        " FROM  dbo.tbl_Personnel_Payroll_Earnings LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Payroll_Earnings_Table ON dbo.tbl_Personnel_Payroll_Earnings.EarningKey = dbo.tbl_Personnel_Payroll_Earnings_Table.PK " & _
        " Where (dbo.tbl_Personnel_Payroll_Earnings.MasterKey = " & rs!PK & ") " & _
        " ORDER BY dbo.tbl_Personnel_Payroll_Earnings_Table.Sorting"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstEarnings.ListItems.Clear
        While Not rt.EOF
            dTotEarnings = dTotEarnings + CDbl(Format(rt!TotalAmount, "#,##0.00"))
            Set x = lstEarnings.ListItems.Add()
            x.Text = rt!EarningKey
            x.SubItems(1) = rt!Description & IIf(Trim(rt!Remarks) <> "", " [" & rt!Remarks & "]", "")
            'x.SubItems(2) = Format(rt!Hours, "#0.00")
            If rt!EarningKey = 1 Then
                x.SubItems(2) = rt!Hours
            Else
                x.SubItems(2) = Format(rt!Hours, "#,##0.00")
            End If
            x.SubItems(3) = Format(rt!TotalAmount, "#,##0.00")
            rt.MoveNext
        Wend
    End If
    rt.Close
    
    CLEAR_DETAILS_Deductions
    t = "SELECT dbo.tbl_Personnel_Payroll_Deductions.DeductionKey, " & _
        " dbo.tbl_Personnel_Payroll_Deductions_Table.Description, " & _
        " dbo.tbl_Personnel_Payroll_Deductions.Amount " & _
        " FROM  dbo.tbl_Personnel_Payroll_Deductions LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Payroll_Deductions.DeductionKey = dbo.tbl_Personnel_Payroll_Deductions_Table.PK " & _
        " Where (dbo.tbl_Personnel_Payroll_Deductions.MasterKey = " & rs!PK & ") " & _
        " ORDER BY dbo.tbl_Personnel_Payroll_Deductions_Table.Sorting"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstDeductions.ListItems.Clear
        While Not rt.EOF
            dTotDeductions = dTotDeductions + CDbl(Format(rt!Amount, "#,##0.00"))
            Set x = lstDeductions.ListItems.Add()
            x.Text = rt!DeductionKey
            x.SubItems(1) = rt!Description
            x.SubItems(2) = Format(rt!Amount, "#,##0.00")
            rt.MoveNext
        Wend
    End If
    rt.Close
    
    lblLocked.Visible = IIf(rs!Locked = 1, True, False)
    
    lblTotEarnings.Caption = Format(dTotEarnings, "#,##0.00")
    lblTotDeductions.Caption = Format(dTotDeductions, "#,##0.00")
    lblNetPay.Caption = Format(CDbl(Format(dTotEarnings, "#,##0.00")) - CDbl(Format(dTotDeductions, "#,##0.00")), "#,##0.00")
    
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    
    SaveSetting App.EXEName, "PersonnelPayroll", "PersonnelPayroll", rs!Ctrl
End If
rs.Close
End Sub

Private Sub CLEARTEXT()
iDivKey = 0
txtName.Text = ""
txtPayrollDate.Text = ""
txtCutOffDate.Text = ""
txtDivision.Text = ""
txtDepartment.Text = ""
txtPosition.Text = ""
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
lblTotEarnings.Caption = "0.00"
lblTotDeductions.Caption = "0.00"
lblNetPay.Caption = "0.00"
lblLocked.Visible = False
CLEAR_DETAILS_Earnings
CLEAR_DETAILS_Deductions
End Sub

Private Sub CLEAR_DETAILS_Earnings()
lstEarnings.ListItems.Clear
Set x = lstEarnings.ListItems.Add()
x.Text = ""
x.SubItems(1) = " "
x.SubItems(2) = " "
x.SubItems(3) = " "
End Sub

Private Sub CLEAR_DETAILS_Deductions()
lstDeductions.ListItems.Clear
Set x = lstDeductions.ListItems.Add()
x.Text = ""
x.SubItems(1) = " "
x.SubItems(2) = " "
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtName.Locked = True
txtPayrollDate.Locked = True
txtCutOffDate.Locked = True
txtDivision.Locked = True
txtDepartment.Locked = True
txtPosition.Locked = True
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

Private Sub cmbPayrollPeriodSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKSearch_Click
End Sub

Private Sub cmdCancelSearch_Click()
picSearch.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True
End Sub

Private Sub cmdOKSearch_Click()
If cmbPayrollPeriodSearch.ListIndex = -1 Then Exit Sub
BROWSER cmbPayrollPeriodSearch.ItemData(cmbPayrollPeriodSearch.ListIndex), "is_FIND"
cmdCancelSearch_Click
End Sub

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert
    Case vbKeyF2
    Case vbKeyDelete
    Case vbKeyF5
    Case vbKeyF6:       PRESS_F6
    Case vbKeyF8
    Case vbKeyF9:       PRESS_F9
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "PersonnelPayroll", "PersonnelPayroll", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "PersonnelPayroll", "PersonnelPayroll", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "PersonnelPayroll", "PersonnelPayroll", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "PersonnelPayroll", "PersonnelPayroll", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption & " [V2]"
Me.Top = (MainForm.ScaleHeight - Me.Height) / 3
Me.Left = (MainForm.ScaleWidth - Me.Width) / 3
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "PersonnelPayroll", "PersonnelPayroll", ""), "is_LOAD"
If Trim(txtName.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelPayroll", "PersonnelPayroll", ""), "is_HOME"

tmp = SetWindowLong(txtSearchSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearchSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)

End Sub

Private Sub Form_Unload(Cancel As Integer)
If picSearch.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstResultSearch_Click()
If lstResultSearch.ListIndex = -1 Then cmbPayrollPeriodSearch.Clear: Exit Sub
cmbPayrollPeriodSearch.Clear
t = "SELECT dbo.tbl_Personnel_Payroll.PK, dbo.tbl_Personnel_Compensation_Period.DateFrom, " & _
    " dbo.tbl_Personnel_Compensation_Period.DateTo, dbo.tbl_Personnel_Division.Description AS Division, " & _
    " dbo.tbl_Personnel_Compensation_Period.PayrollDate " & _
    " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_ActionNew.DivisionKey = dbo.tbl_Personnel_Division.PK " & _
    " Where (dbo.tbl_Personnel_Payroll.EmployeeKey = " & lstResultSearch.ItemData(lstResultSearch.ListIndex) & ") " & _
    " AND ((SELECT COUNT(*) AS RecCnt From dbo.tbl_Personnel_Payroll_Earnings Where (dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK)) > 0) " & _
    " ORDER BY dbo.tbl_Personnel_Compensation_Period.DateFrom DESC, dbo.tbl_Personnel_Compensation_Period.DateTo DESC"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
While Not rt.EOF
    cmbPayrollPeriodSearch.AddItem Format(rt!PayrollDate, "mm/dd/yyyy") & " [" & rt!Division & "]"
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
    Case "Add"
    Case "Edit"
    Case "Delete"
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then Else BROWSER GetSetting(App.EXEName, "PersonnelPayroll", "PersonnelPayroll", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then Else BROWSER GetSetting(App.EXEName, "PersonnelPayroll", "PersonnelPayroll", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "PersonnelPayroll", "PersonnelPayroll", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "PersonnelPayroll", "PersonnelPayroll", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":   PRESS_F9
    Case "Post":
    Case "Refresh":
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtSearchSearch_Change()
If Trim(txtSearchSearch.Text) = "" Then lstResultSearch.Clear: cmbPayrollPeriodSearch.Clear: Exit Sub
lstResultSearch.Clear: cmbPayrollPeriodSearch.Clear
s = "SELECT dbo.tbl_Personnel_Payroll.EmployeeKey, dbo.tbl_Personnel_IDNumber.IDNumber, " & _
    " dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, " & _
    " dbo.tbl_Personnel_Information.MiddleName " & _
    " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Payroll.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
    " WHERE ((SELECT COUNT(*) AS RecCnt From dbo.tbl_Personnel_Payroll_Earnings Where (dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK)) > 0) " & _
    " GROUP BY dbo.tbl_Personnel_Payroll.EmployeeKey, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName " & _
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
