VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelAllowanceBrowse 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPersonnelAllowanceBrowse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picGenerate 
      Height          =   2415
      Left            =   2400
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4260
      BackColor       =   15396057
      Begin VB.TextBox txtFrom 
         Height          =   315
         Left            =   840
         TabIndex        =   29
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtTo 
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelGenerate 
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
         Picture         =   "frmPersonnelAllowanceBrowse.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1680
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKGenerate 
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
         Picture         =   "frmPersonnelAllowanceBrowse.frx":1026
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1680
         Width           =   1560
      End
      Begin VB.ComboBox cmbDivision 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   720
         Width           =   3495
      End
      Begin RPVGCC.b8TitleBar b8TitleBar4 
         Height          =   345
         Left            =   40
         TabIndex        =   26
         Top             =   40
         Width           =   3890
         _ExtentX        =   6853
         _ExtentY        =   609
         Caption         =   "Generate"
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
         Icon            =   "frmPersonnelAllowanceBrowse.frx":1698
         ShadowVisible   =   0   'False
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "PERIOD"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TO"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   30
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "SELECT DIVISION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   480
         Width           =   3375
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   38
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   39
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
         MouseIcon       =   "frmPersonnelAllowanceBrowse.frx":1C32
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
   Begin RPVGCC.b8Container picProgress 
      Height          =   975
      Left            =   1560
      TabIndex        =   32
      Top             =   1440
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1720
      BackColor       =   13023396
      Begin VB.Timer TimerGenerate 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   4800
         Top             =   480
      End
      Begin VB.PictureBox picProgressBar 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5235
         TabIndex        =   33
         Top             =   120
         Width           =   5295
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   720
      ScaleHeight     =   1815
      ScaleWidth      =   7455
      TabIndex        =   1
      Top             =   1320
      Width           =   7455
      Begin VB.TextBox txtRate 
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   36
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtHours 
         Height          =   315
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   34
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   7
         Top             =   0
         Width           =   6495
      End
      Begin VB.TextBox txtDepartment 
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   6
         Top             =   720
         Width           =   6495
      End
      Begin VB.TextBox txtPosition 
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   5
         Top             =   1080
         Width           =   6495
      End
      Begin VB.TextBox txtAmount 
         Height          =   315
         Left            =   3960
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtPeriod 
         Height          =   315
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtDivision 
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   2
         Top             =   360
         Width           =   6495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   37
         Top             =   1470
         Width           =   495
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Hours"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   35
         Top             =   1470
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   30
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Allowance"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3120
         TabIndex        =   10
         Top             =   1470
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Period"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4920
         TabIndex        =   9
         Top             =   1470
         Width           =   615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   3570
      Width           =   9120
      _ExtentX        =   16087
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
   Begin RPVGCC.b8Container picSearch 
      Height          =   3075
      Left            =   2160
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5424
      BackColor       =   15396057
      Begin VB.CommandButton cmdOK 
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
         Picture         =   "frmPersonnelAllowanceBrowse.frx":1F4C
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2480
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancel 
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
         Picture         =   "frmPersonnelAllowanceBrowse.frx":25BE
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2480
         Width           =   1560
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   4215
      End
      Begin VB.ListBox lstResult 
         Height          =   1230
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   4215
      End
      Begin VB.ComboBox cmbEffectDate 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2100
         Width           =   3135
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   45
         TabIndex        =   20
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
         Icon            =   "frmPersonnelAllowanceBrowse.frx":2D1A
         ShadowVisible   =   0   'False
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Period"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   2100
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10080
      Top             =   2160
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
            Picture         =   "frmPersonnelAllowanceBrowse.frx":32B4
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowanceBrowse.frx":3F8E
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowanceBrowse.frx":4C68
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowanceBrowse.frx":5942
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowanceBrowse.frx":661C
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowanceBrowse.frx":72F6
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowanceBrowse.frx":7FD0
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowanceBrowse.frx":8CAA
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowanceBrowse.frx":9984
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowanceBrowse.frx":A25E
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowanceBrowse.frx":AF38
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowanceBrowse.frx":BC12
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowanceBrowse.frx":C8EC
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowanceBrowse.frx":D5C6
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowanceBrowse.frx":E2A0
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPersonnelAllowanceBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_GENERATING = 1

Public iType As Long

Dim dtmFrom As Date
Dim dtmTo As Date
Dim tmp As Long

Dim Filename As String
Dim WorkbookName As String
Dim iWorkSheet As Integer
Dim RowCnt, ColCnt, strRange, k, strValue, iCompensationRate, dRate, _
iPeriod, iHour, iDivision, i


Private Sub BROWSER(iPK, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If iPK <> "" Then
            s = "SELECT TOP 1 tbl_Personnel_Allowance_Per_Period.PK, tbl_Personnel_IDNumber.IDNumber + ' - ' + tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
                " (SELECT TOP 1 Division From tbl_Personnel_Action Where (EmpPK = tbl_Personnel_Allowance_Per_Period.EmpPK) And (EffectivityDate <= tbl_Personnel_Compensation_Period.DateTo) ORDER BY EffectivityDate DESC) AS Division, " & _
                " (SELECT TOP 1 tbl_Personnel_Department.DepartmentCode + ' - ' + tbl_Personnel_Department.DepartmentName AS Department FROM tbl_Personnel_Action AS tbl_Personnel_Action_1 LEFT OUTER JOIN tbl_Personnel_Department ON tbl_Personnel_Action_1.Dept = tbl_Personnel_Department.PK " & _
                " WHERE (tbl_Personnel_Action_1.EmpPK = tbl_Personnel_Allowance_Per_Period.EmpPK) AND (tbl_Personnel_Action_1.EffectivityDate <= tbl_Personnel_Compensation_Period.DateTo) ORDER BY tbl_Personnel_Action_1.EffectivityDate DESC) AS Department, " & _
                " (SELECT TOP 1 tbl_Personnel_Position.PositionCode + ' - ' + tbl_Personnel_Position.PositionName AS Position FROM tbl_Personnel_Action AS tbl_Personnel_Action_2 LEFT OUTER JOIN tbl_Personnel_Position ON tbl_Personnel_Action_2.Positions = tbl_Personnel_Position.PK " & _
                " WHERE (tbl_Personnel_Action_2.EmpPK = tbl_Personnel_Allowance_Per_Period.EmpPK) AND (tbl_Personnel_Action_2.EffectivityDate <= tbl_Personnel_Compensation_Period.DateTo) ORDER BY tbl_Personnel_Action_2.EffectivityDate DESC) AS Position, " & _
                " tbl_Personnel_Allowance_Per_Period.Amount, tbl_Personnel_Compensation_Period.DateFrom , tbl_Personnel_Compensation_Period.DateTo, tbl_Personnel_Allowance_Per_Period.LastModified, tbl_Personnel_Allowance_Per_Period.NoHours, " & _
                " tbl_Personnel_Allowance_Per_Period.RatePerHour " & _
                " FROM tbl_Personnel_Allowance_Per_Period LEFT OUTER JOIN " & _
                " tbl_Personnel_Compensation_Period ON tbl_Personnel_Allowance_Per_Period.Period = tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " tbl_Personnel_IDNumber ON tbl_Personnel_Allowance_Per_Period.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
                " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
                " WHERE (tbl_Personnel_Allowance_Per_Period.PK = " & iPK & ") " & _
                " ORDER BY tbl_Personnel_Allowance_Per_Period.PK"
        Else
            s = "SELECT TOP 1 tbl_Personnel_Allowance_Per_Period.PK, tbl_Personnel_IDNumber.IDNumber + ' - ' + tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
                " (SELECT TOP 1 Division From tbl_Personnel_Action Where (EmpPK = tbl_Personnel_Allowance_Per_Period.EmpPK) And (EffectivityDate <= tbl_Personnel_Compensation_Period.DateTo) ORDER BY EffectivityDate DESC) AS Division, " & _
                " (SELECT TOP 1 tbl_Personnel_Department.DepartmentCode + ' - ' + tbl_Personnel_Department.DepartmentName AS Department FROM tbl_Personnel_Action AS tbl_Personnel_Action_1 LEFT OUTER JOIN tbl_Personnel_Department ON tbl_Personnel_Action_1.Dept = tbl_Personnel_Department.PK " & _
                " WHERE (tbl_Personnel_Action_1.EmpPK = tbl_Personnel_Allowance_Per_Period.EmpPK) AND (tbl_Personnel_Action_1.EffectivityDate <= tbl_Personnel_Compensation_Period.DateTo) ORDER BY tbl_Personnel_Action_1.EffectivityDate DESC) AS Department, " & _
                " (SELECT TOP 1 tbl_Personnel_Position.PositionCode + ' - ' + tbl_Personnel_Position.PositionName AS Position FROM tbl_Personnel_Action AS tbl_Personnel_Action_2 LEFT OUTER JOIN tbl_Personnel_Position ON tbl_Personnel_Action_2.Positions = tbl_Personnel_Position.PK " & _
                " WHERE (tbl_Personnel_Action_2.EmpPK = tbl_Personnel_Allowance_Per_Period.EmpPK) AND (tbl_Personnel_Action_2.EffectivityDate <= tbl_Personnel_Compensation_Period.DateTo) ORDER BY tbl_Personnel_Action_2.EffectivityDate DESC) AS Position, " & _
                " tbl_Personnel_Allowance_Per_Period.Amount, tbl_Personnel_Compensation_Period.DateFrom , tbl_Personnel_Compensation_Period.DateTo, tbl_Personnel_Allowance_Per_Period.LastModified, tbl_Personnel_Allowance_Per_Period.NoHours, " & _
                " tbl_Personnel_Allowance_Per_Period.RatePerHour " & _
                " FROM tbl_Personnel_Allowance_Per_Period LEFT OUTER JOIN " & _
                " tbl_Personnel_Compensation_Period ON tbl_Personnel_Allowance_Per_Period.Period = tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " tbl_Personnel_IDNumber ON tbl_Personnel_Allowance_Per_Period.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
                " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
                " ORDER BY tbl_Personnel_Allowance_Per_Period.PK DESC"
        End If
    Case "is_HOME"
        s = "SELECT TOP 1 tbl_Personnel_Allowance_Per_Period.PK, tbl_Personnel_IDNumber.IDNumber + ' - ' + tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
            " (SELECT TOP 1 Division From tbl_Personnel_Action Where (EmpPK = tbl_Personnel_Allowance_Per_Period.EmpPK) And (EffectivityDate <= tbl_Personnel_Compensation_Period.DateTo) ORDER BY EffectivityDate DESC) AS Division, " & _
            " (SELECT TOP 1 tbl_Personnel_Department.DepartmentCode + ' - ' + tbl_Personnel_Department.DepartmentName AS Department FROM tbl_Personnel_Action AS tbl_Personnel_Action_1 LEFT OUTER JOIN tbl_Personnel_Department ON tbl_Personnel_Action_1.Dept = tbl_Personnel_Department.PK " & _
            " WHERE (tbl_Personnel_Action_1.EmpPK = tbl_Personnel_Allowance_Per_Period.EmpPK) AND (tbl_Personnel_Action_1.EffectivityDate <= tbl_Personnel_Compensation_Period.DateTo) ORDER BY tbl_Personnel_Action_1.EffectivityDate DESC) AS Department, " & _
            " (SELECT TOP 1 tbl_Personnel_Position.PositionCode + ' - ' + tbl_Personnel_Position.PositionName AS Position FROM tbl_Personnel_Action AS tbl_Personnel_Action_2 LEFT OUTER JOIN tbl_Personnel_Position ON tbl_Personnel_Action_2.Positions = tbl_Personnel_Position.PK " & _
            " WHERE (tbl_Personnel_Action_2.EmpPK = tbl_Personnel_Allowance_Per_Period.EmpPK) AND (tbl_Personnel_Action_2.EffectivityDate <= tbl_Personnel_Compensation_Period.DateTo) ORDER BY tbl_Personnel_Action_2.EffectivityDate DESC) AS Position, " & _
            " tbl_Personnel_Allowance_Per_Period.Amount, tbl_Personnel_Compensation_Period.DateFrom , tbl_Personnel_Compensation_Period.DateTo, tbl_Personnel_Allowance_Per_Period.LastModified, tbl_Personnel_Allowance_Per_Period.NoHours, " & _
            " tbl_Personnel_Allowance_Per_Period.RatePerHour " & _
            " FROM tbl_Personnel_Allowance_Per_Period LEFT OUTER JOIN " & _
            " tbl_Personnel_Compensation_Period ON tbl_Personnel_Allowance_Per_Period.Period = tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_IDNumber ON tbl_Personnel_Allowance_Per_Period.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
            " ORDER BY tbl_Personnel_Allowance_Per_Period.PK DESC"
    Case "is_PAGEUP"
        s = "SELECT TOP 1 tbl_Personnel_Allowance_Per_Period.PK, tbl_Personnel_IDNumber.IDNumber + ' - ' + tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
            " (SELECT TOP 1 Division From tbl_Personnel_Action Where (EmpPK = tbl_Personnel_Allowance_Per_Period.EmpPK) And (EffectivityDate <= tbl_Personnel_Compensation_Period.DateTo) ORDER BY EffectivityDate DESC) AS Division, " & _
            " (SELECT TOP 1 tbl_Personnel_Department.DepartmentCode + ' - ' + tbl_Personnel_Department.DepartmentName AS Department FROM tbl_Personnel_Action AS tbl_Personnel_Action_1 LEFT OUTER JOIN tbl_Personnel_Department ON tbl_Personnel_Action_1.Dept = tbl_Personnel_Department.PK " & _
            " WHERE (tbl_Personnel_Action_1.EmpPK = tbl_Personnel_Allowance_Per_Period.EmpPK) AND (tbl_Personnel_Action_1.EffectivityDate <= tbl_Personnel_Compensation_Period.DateTo) ORDER BY tbl_Personnel_Action_1.EffectivityDate DESC) AS Department, " & _
            " (SELECT TOP 1 tbl_Personnel_Position.PositionCode + ' - ' + tbl_Personnel_Position.PositionName AS Position FROM tbl_Personnel_Action AS tbl_Personnel_Action_2 LEFT OUTER JOIN tbl_Personnel_Position ON tbl_Personnel_Action_2.Positions = tbl_Personnel_Position.PK " & _
            " WHERE (tbl_Personnel_Action_2.EmpPK = tbl_Personnel_Allowance_Per_Period.EmpPK) AND (tbl_Personnel_Action_2.EffectivityDate <= tbl_Personnel_Compensation_Period.DateTo) ORDER BY tbl_Personnel_Action_2.EffectivityDate DESC) AS Position, " & _
            " tbl_Personnel_Allowance_Per_Period.Amount, tbl_Personnel_Compensation_Period.DateFrom , tbl_Personnel_Compensation_Period.DateTo, tbl_Personnel_Allowance_Per_Period.LastModified, tbl_Personnel_Allowance_Per_Period.NoHours, " & _
            " tbl_Personnel_Allowance_Per_Period.RatePerHour " & _
            " FROM tbl_Personnel_Allowance_Per_Period LEFT OUTER JOIN " & _
            " tbl_Personnel_Compensation_Period ON tbl_Personnel_Allowance_Per_Period.Period = tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_IDNumber ON tbl_Personnel_Allowance_Per_Period.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
            " WHERE (tbl_Personnel_Allowance_Per_Period.PK > " & iPK & ") " & _
            " ORDER BY tbl_Personnel_Allowance_Per_Period.PK"
    Case "is_PAGEDOWN"
        s = "SELECT TOP 1 tbl_Personnel_Allowance_Per_Period.PK, tbl_Personnel_IDNumber.IDNumber + ' - ' + tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
            " (SELECT TOP 1 Division From tbl_Personnel_Action Where (EmpPK = tbl_Personnel_Allowance_Per_Period.EmpPK) And (EffectivityDate <= tbl_Personnel_Compensation_Period.DateTo) ORDER BY EffectivityDate DESC) AS Division, " & _
            " (SELECT TOP 1 tbl_Personnel_Department.DepartmentCode + ' - ' + tbl_Personnel_Department.DepartmentName AS Department FROM tbl_Personnel_Action AS tbl_Personnel_Action_1 LEFT OUTER JOIN tbl_Personnel_Department ON tbl_Personnel_Action_1.Dept = tbl_Personnel_Department.PK " & _
            " WHERE (tbl_Personnel_Action_1.EmpPK = tbl_Personnel_Allowance_Per_Period.EmpPK) AND (tbl_Personnel_Action_1.EffectivityDate <= tbl_Personnel_Compensation_Period.DateTo) ORDER BY tbl_Personnel_Action_1.EffectivityDate DESC) AS Department, " & _
            " (SELECT TOP 1 tbl_Personnel_Position.PositionCode + ' - ' + tbl_Personnel_Position.PositionName AS Position FROM tbl_Personnel_Action AS tbl_Personnel_Action_2 LEFT OUTER JOIN tbl_Personnel_Position ON tbl_Personnel_Action_2.Positions = tbl_Personnel_Position.PK " & _
            " WHERE (tbl_Personnel_Action_2.EmpPK = tbl_Personnel_Allowance_Per_Period.EmpPK) AND (tbl_Personnel_Action_2.EffectivityDate <= tbl_Personnel_Compensation_Period.DateTo) ORDER BY tbl_Personnel_Action_2.EffectivityDate DESC) AS Position, " & _
            " tbl_Personnel_Allowance_Per_Period.Amount, tbl_Personnel_Compensation_Period.DateFrom , tbl_Personnel_Compensation_Period.DateTo, tbl_Personnel_Allowance_Per_Period.LastModified, tbl_Personnel_Allowance_Per_Period.NoHours, " & _
            " tbl_Personnel_Allowance_Per_Period.RatePerHour " & _
            " FROM tbl_Personnel_Allowance_Per_Period LEFT OUTER JOIN " & _
            " tbl_Personnel_Compensation_Period ON tbl_Personnel_Allowance_Per_Period.Period = tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_IDNumber ON tbl_Personnel_Allowance_Per_Period.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
            " WHERE (tbl_Personnel_Allowance_Per_Period.PK < " & iPK & ") " & _
            " ORDER BY tbl_Personnel_Allowance_Per_Period.PK DESC"
    Case "is_END"
        s = "SELECT TOP 1 tbl_Personnel_Allowance_Per_Period.PK, tbl_Personnel_IDNumber.IDNumber + ' - ' + tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
            " (SELECT TOP 1 Division From tbl_Personnel_Action Where (EmpPK = tbl_Personnel_Allowance_Per_Period.EmpPK) And (EffectivityDate <= tbl_Personnel_Compensation_Period.DateTo) ORDER BY EffectivityDate DESC) AS Division, " & _
            " (SELECT TOP 1 tbl_Personnel_Department.DepartmentCode + ' - ' + tbl_Personnel_Department.DepartmentName AS Department FROM tbl_Personnel_Action AS tbl_Personnel_Action_1 LEFT OUTER JOIN tbl_Personnel_Department ON tbl_Personnel_Action_1.Dept = tbl_Personnel_Department.PK " & _
            " WHERE (tbl_Personnel_Action_1.EmpPK = tbl_Personnel_Allowance_Per_Period.EmpPK) AND (tbl_Personnel_Action_1.EffectivityDate <= tbl_Personnel_Compensation_Period.DateTo) ORDER BY tbl_Personnel_Action_1.EffectivityDate DESC) AS Department, " & _
            " (SELECT TOP 1 tbl_Personnel_Position.PositionCode + ' - ' + tbl_Personnel_Position.PositionName AS Position FROM tbl_Personnel_Action AS tbl_Personnel_Action_2 LEFT OUTER JOIN tbl_Personnel_Position ON tbl_Personnel_Action_2.Positions = tbl_Personnel_Position.PK " & _
            " WHERE (tbl_Personnel_Action_2.EmpPK = tbl_Personnel_Allowance_Per_Period.EmpPK) AND (tbl_Personnel_Action_2.EffectivityDate <= tbl_Personnel_Compensation_Period.DateTo) ORDER BY tbl_Personnel_Action_2.EffectivityDate DESC) AS Position, " & _
            " tbl_Personnel_Allowance_Per_Period.Amount, tbl_Personnel_Compensation_Period.DateFrom , tbl_Personnel_Compensation_Period.DateTo, tbl_Personnel_Allowance_Per_Period.LastModified, tbl_Personnel_Allowance_Per_Period.NoHours, " & _
            " tbl_Personnel_Allowance_Per_Period.RatePerHour " & _
            " FROM tbl_Personnel_Allowance_Per_Period LEFT OUTER JOIN " & _
            " tbl_Personnel_Compensation_Period ON tbl_Personnel_Allowance_Per_Period.Period = tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_IDNumber ON tbl_Personnel_Allowance_Per_Period.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
            " ORDER BY tbl_Personnel_Allowance_Per_Period.PK"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtName.Text = rs!EmployeeName
    txtDivision.Text = IIf(rs!Division = 1, "CLUB HOUSE", "MAINTENANCE")
    txtDepartment.Text = rs!Department
    txtPosition.Text = rs!Position
    txtRate.Text = Format(rs!RatePerHour, "#,##0.00")
    txtHours.Text = Format(rs!NoHours, "#,##0.00")
    txtAmount.Text = Format(rs!Amount, "#,##0.00")
    txtPeriod.Text = Format(rs!DateFrom, "mm/dd/yyyy") & " - " & Format(rs!DateTo, "mm/dd/yyyy")
    StatusBar1.Panels(1).Text = rs!PK
    StatusBar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    SaveSetting App.EXEName, "AllowancePK", "AllowKey", rs!PK
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picGenerate.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If picProgress.Visible = True Then Exit Sub
iType = 1       'Generate
b8TitleBar4.Caption = "Generate"
DoEvents
picMain.Enabled = False
picToolbar.Enabled = False
picGenerate.ZOrder 0
cmbDivision.ListIndex = -1
txtFrom.Text = ""
txtTo.Text = ""
picGenerate.Visible = True
cmbDivision.SetFocus
End Sub

Private Sub PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picGenerate.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If picProgress.Visible = True Then Exit Sub
picSearch.ZOrder 0
txtSearch.Text = ""
picSearch.Visible = True
txtSearch.SetFocus
End Sub

Private Sub PRESS_F9()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picGenerate.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If picProgress.Visible = True Then Exit Sub

PopupMenu MainFormPopupF.mnuAllowancePrint, , Toolbar1.Buttons(17).Left, 500

End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picGenerate.Visible = True Then cmdCancelGenerate_Click: Exit Sub
    If picSearch.Visible = True Then cmdCancel_Click: Exit Sub
    Unload Me
Else

End If
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
cmdCancel_Click
End Sub

Private Sub b8TitleBar4_CLoseClick()
cmdCancelGenerate_Click
End Sub

Private Sub cmbEffectDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub cmdCancel_Click()
picSearch.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True
End Sub


Private Sub cmdCancelGenerate_Click()
picGenerate.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True
End Sub

Private Sub cmdOK_Click()
If cmbEffectDate.ListIndex = -1 Then Exit Sub
BROWSER cmbEffectDate.ItemData(cmbEffectDate.ListIndex), "is_LOAD"
cmdCancel_Click
End Sub

Private Sub cmdOKGenerate_Click()
If cmbDivision.ListIndex = -1 Then Exit Sub
If IsDate(txtFrom.Text) = False Then MsgBox "Please Supply a Valid Date Range!                        ", vbCritical, "Error...": txtFrom.SetFocus: Exit Sub
If IsDate(txtTo.Text) = False Then MsgBox "Please Supply a Valid Date Range!                        ", vbCritical, "Error...": txtTo.SetFocus: Exit Sub

iPeriod = 0: dtmFrom = FormatDateTime(txtFrom.Text, vbShortDate): dtmTo = FormatDateTime(txtTo.Text, vbShortDate)
t = "SELECT tbl_Personnel_Compensation_Period.* " & _
    " FROM tbl_Personnel_Compensation_Period " & _
    " WHERE (DateFrom = '" & FormatDateTime(txtFrom.Text, vbShortDate) & "') " & _
    " AND (DateTo = '" & FormatDateTime(txtTo.Text, vbShortDate) & "')"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then
    iPeriod = rt!PK
End If
rt.Close

If CDbl(iPeriod) = 0 Then MsgBox "Invalid period!                   ", vbCritical, "Error...": Exit Sub

If iType = 1 Then
    s = "SELECT tbl_Personnel_Allowance_Per_Period.* " & _
        " FROM tbl_Personnel_Allowance_Per_Period " & _
        " WHERE (Period = " & iPeriod & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        If MsgBox("Has existing record! Regenerate Employee Allowance?                      ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then rs.Close: Exit Sub
    End If
    rs.Close
ElseIf iType = 2 Or iType = 3 Then
    s = "SELECT tbl_Personnel_Allowance_Per_Period.* " & _
        " FROM tbl_Personnel_Allowance_Per_Period " & _
        " WHERE (Period = " & iPeriod & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount = 0 Then
        MsgBox "No Record Found!                ", vbCritical, "Error..."
        rs.Close
        Exit Sub
    End If
    rs.Close
End If

iDivision = cmbDivision.ListIndex + 1
picGenerate.Visible = False

picProgress.ZOrder 0
picProgressBar.BackColor = &HFFFFFF
picProgress.Visible = True
TimerGenerate.Enabled = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert:   PRESS_INSERT
    Case vbKeyF6:       PRESS_F6
    Case vbKeyF9:       PRESS_F9
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "AllowancePK", "AllowKey", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "AllowancePK", "AllowKey", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "AllowancePK", "AllowKey", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "AllowancePK", "AllowKey", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.Height - Me.Height) / 3
Me.Left = (MainForm.Width - Me.Width) / 5

With cmbDivision
    .Clear
    .AddItem "CLUB HOUSE"
    .AddItem "MAINTENANCE"
    .ListIndex = 0
End With

TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "AllowancePK", "AllowKey", ""), "is_LOAD"
If Trim(txtName.Text) = "" Then BROWSER GetSetting(App.EXEName, "AllowancePK", "AllowKey", ""), "is_HOME"


'tmp = SetWindowLong(txtName.hWnd, GWL_STYLE, GetWindowLong(txtName.hWnd, GWL_STYLE) Or ES_UPPERCASE)
'tmp = SetWindowLong(txtSearchAdd.hWnd, GWL_STYLE, GetWindowLong(txtSearchAdd.hWnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picProgress.Visible = True Then Cancel = -1
If picGenerate.Visible = True Then Cancel = -1
If picSearch.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstResult_Click()
If lstResult.ListIndex = -1 Then cmbEffectDate.Clear: Exit Sub
cmbEffectDate.Clear
t = "SELECT tbl_Personnel_Allowance_Per_Period.PK, " & _
    " tbl_Personnel_Compensation_Period.DateFrom, " & _
    " tbl_Personnel_Compensation_Period.DateTo " & _
    " FROM tbl_Personnel_Allowance_Per_Period LEFT OUTER JOIN " & _
    " tbl_Personnel_Compensation_Period ON tbl_Personnel_Allowance_Per_Period.Period = tbl_Personnel_Compensation_Period.PK " & _
    " Where (tbl_Personnel_Allowance_Per_Period.EmpPK = " & lstResult.ItemData(lstResult.ListIndex) & ") " & _
    " ORDER BY tbl_Personnel_Compensation_Period.DateFrom DESC, tbl_Personnel_Compensation_Period.DateTo DESC"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
While Not rt.EOF
    cmbEffectDate.AddItem Format(rt!DateFrom, "mm/dd/yyyy") & " - " & Format(rt!DateTo, "mm/dd/yyyy")
    cmbEffectDate.ItemData(cmbEffectDate.NewIndex) = rt!PK
    rt.MoveNext
Wend
rt.Close
If cmbEffectDate.ListCount Then cmbEffectDate.ListIndex = 0
End Sub

Private Sub lstResult_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbEffectDate.SetFocus
End Sub

Private Sub TimerGenerate_Timer()
TimerGenerate.Enabled = False
If iType = 1 Then
    TRANSACTIONTYPE = is_GENERATING
    
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Allowance_Per_Period WHERE (Period = " & iPeriod & ")"
    i = 0
    's = "SELECT tbl_Personnel_Allowance.* " & _
        " FROM tbl_Personnel_Allowance "
    s = "SELECT PK AS EmpPK, " & _
        " (SELECT TOP 1 RatePerHour " & _
        " From tbl_Personnel_Allowance " & _
        " WHERE (EmpPK = tbl_Personnel_IDNumber.PK) " & _
        " AND (EffectDate <= '" & FormatDateTime(dtmTo, vbShortDate) & "') " & _
        " ORDER BY EffectDate DESC) AS RatePerHour " & _
        " From tbl_Personnel_IDNumber " & _
        " WHERE ((SELECT TOP 1 RatePerHour " & _
        " FROM tbl_Personnel_Allowance AS tbl_Personnel_Allowance_1 " & _
        " WHERE (EmpPK = tbl_Personnel_IDNumber.PK) " & _
        " AND (EffectDate <= '" & FormatDateTime(dtmTo, vbShortDate) & "') " & _
        " ORDER BY EffectDate DESC) IS NOT NULL)"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        i = i + 1
        iCompensationRate = 0
        dRate = 0
        
        t = "SELECT TOP 1 CompensationRateKey " & _
            " From tbl_Personnel_ActionNew " & _
            " WHERE (EmpPK = " & rs!EmpPK & ") " & _
            " AND (EffectivityDate <= '" & FormatDateTime(dtmTo, vbShortDate) & "') " & _
            " ORDER BY EffectivityDate DESC"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            iCompensationRate = rt!CompensationRateKey
        End If
        rt.Close
        
        t = "SELECT TOP 1 Rate, RatePerHour " & _
            " FROM tbl_Personnel_Allowance " & _
            " WHERE (EmpPK = " & rs!EmpPK & ") " & _
            " AND (EffectDate <= '" & FormatDateTime(dtmTo, vbShortDate) & "') " & _
            " ORDER BY EffectDate DESC"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
'            If CInt(iCompensationRate) = 1 Then
'                dRate = ((CDbl(rt!Rate) / 2) / 13.08333) / 8
'            ElseIf CInt(iCompensationRate) = 2 Then
'                dRate = CDbl(rt!Rate) / 8
'            End If
            If CInt(iCompensationRate) = 3 Then
                dRate = ((CDbl(rt!Rate) / 2) / 13.08333) / 8
            Else
                dRate = CDbl(rt!Rate) / 8
            End If
        End If
        rt.Close
        
        'MsgBox dRate
        
        If CDbl(dRate) > 0 Then
            iHour = 0
            't = "SELECT tbl_Personnel_Compensation.* " & _
                " FROM tbl_Personnel_Compensation " & _
                " WHERE (EmpPK = " & rs!EmpPK & ") " & _
                " AND (Period = " & iPeriod & ") "
            t = "SELECT SUM(dbo.tbl_Personnel_Hours_Regular.NoHours) AS NoHours " & _
                " FROM  dbo.tbl_Personnel_Hours_Regular LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Hours ON dbo.tbl_Personnel_Hours_Regular.MasterKey = dbo.tbl_Personnel_Hours.PK " & _
                " WHERE (dbo.tbl_Personnel_Hours.EmployeeKey = " & rs!EmpPK & ") " & _
                " AND (dbo.tbl_Personnel_Hours.PayrollPeriodKey = " & iPeriod & ") " & _
                " AND (dbo.tbl_Personnel_Hours_Regular.EarningKey = 1) OR " & _
                " (dbo.tbl_Personnel_Hours.EmployeeKey = " & rs!EmpPK & ") " & _
                " AND (dbo.tbl_Personnel_Hours.PayrollPeriodKey = " & iPeriod & ") " & _
                " AND (dbo.tbl_Personnel_Hours_Regular.EarningKey = 6)"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                'iHour = CDbl(rt!NoHours) + CDbl(rt!SL_Hours)
                iHour = IIf(IsNull(rt!NoHours), 0, rt!NoHours)
            End If
            rt.Close
            
            If CDbl(iHour) > 0 Then
                ConnOmega.Execute "INSERT INTO tbl_Personnel_Allowance_Per_Period " & _
                                  " (EmpPK, Period, Division, NoHours, RatePerHour, LastModified) " & _
                                  " VALUES (" & rs!EmpPK & ", " & iPeriod & ", " & iDivision & ", " & _
                                  " " & CDbl(iHour) & ", " & CDbl(dRate) & ", " & _
                                  " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
            End If
            
        End If
        UpdateProgress picProgressBar, i / rs.RecordCount
        rs.MoveNext
    Wend
    rs.Close
    
    picMain.Enabled = True
    picToolbar.Enabled = True
    picProgress.Visible = False
    
    TRANSACTIONTYPE = is_REFRESH
    t = "SELECT TOP 1 PK " & _
        " FROM tbl_Personnel_Allowance_Per_Period " & _
        " ORDER BY PK DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        BROWSER rt!PK, "is_HOME"
    End If
    rt.Close
    
    Exit Sub
    
End If

If iType = 2 Then
    
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Allowance_Report WHERE (LogIn = '" & gbl_UserName & "')"
    i = 0
    's = "SELECT tbl_Personnel_Allowance_Per_Period.PK, tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
        " (SELECT TOP 1 Division From tbl_Personnel_Action Where (EmpPK = tbl_Personnel_Allowance_Per_Period.EmpPK) And (EffectivityDate <= tbl_Personnel_Compensation_Period.DateTo) ORDER BY EffectivityDate DESC) AS Division, " & _
        " (SELECT TOP 1 tbl_Personnel_Department.DepartmentCode + ' - ' + tbl_Personnel_Department.DepartmentName AS Department FROM tbl_Personnel_Action AS tbl_Personnel_Action_1 LEFT OUTER JOIN tbl_Personnel_Department ON tbl_Personnel_Action_1.Dept = tbl_Personnel_Department.PK " & _
        " WHERE (tbl_Personnel_Action_1.EmpPK = tbl_Personnel_Allowance_Per_Period.EmpPK) AND (tbl_Personnel_Action_1.EffectivityDate <= tbl_Personnel_Compensation_Period.DateTo) ORDER BY tbl_Personnel_Action_1.EffectivityDate DESC) AS Department, " & _
        " (SELECT TOP 1 tbl_Personnel_Position.PositionCode + ' - ' + tbl_Personnel_Position.PositionName AS Position FROM tbl_Personnel_Action AS tbl_Personnel_Action_2 LEFT OUTER JOIN tbl_Personnel_Position ON tbl_Personnel_Action_2.Positions = tbl_Personnel_Position.PK " & _
        " WHERE (tbl_Personnel_Action_2.EmpPK = tbl_Personnel_Allowance_Per_Period.EmpPK) AND (tbl_Personnel_Action_2.EffectivityDate <= tbl_Personnel_Compensation_Period.DateTo) ORDER BY tbl_Personnel_Action_2.EffectivityDate DESC) AS Position, " & _
        " tbl_Personnel_Allowance_Per_Period.Amount, tbl_Personnel_Compensation_Period.DateFrom , tbl_Personnel_Compensation_Period.DateTo, tbl_Personnel_Allowance_Per_Period.LastModified, " & _
        " tbl_Personnel_Allowance_Per_Period.NoHours, tbl_Personnel_Allowance_Per_Period.RatePerHour " & _
        " FROM tbl_Personnel_Allowance_Per_Period LEFT OUTER JOIN " & _
        " tbl_Personnel_Compensation_Period ON tbl_Personnel_Allowance_Per_Period.Period = tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_IDNumber ON tbl_Personnel_Allowance_Per_Period.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
        " WHERE (tbl_Personnel_Allowance_Per_Period.Division = " & iDivision & ") " & _
        " AND (tbl_Personnel_Allowance_Per_Period.Period = " & iPeriod & ") " & _
        " ORDER BY tbl_Personnel_Allowance_Per_Period.PK "
    s = "SELECT dbo.tbl_Personnel_Allowance_Per_Period.PK, dbo.tbl_Personnel_Information.LastName + ',  ' + dbo.tbl_Personnel_Information.FirstName + '  ' + dbo.tbl_Personnel_Information.MiddleName AS EmployeeName," & _
        " (SELECT TOP (1) dbo.tbl_Personnel_Division.Description FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_ActionNew.DivisionKey = dbo.tbl_Personnel_Division.PK Where (dbo.tbl_Personnel_ActionNew.EffectivityDate <= dbo.tbl_Personnel_Compensation_Period.DateTo) And (dbo.tbl_Personnel_ActionNew.EmpPK = dbo.tbl_Personnel_Allowance_Per_Period.EmpPK) ORDER BY dbo.tbl_Personnel_ActionNew.EffectivityDate) AS Division, " & _
        " (SELECT TOP (1) dbo.tbl_Personnel_Department.DepartmentCode + ' - ' + dbo.tbl_Personnel_Department.DepartmentName AS Dept FROM  dbo.tbl_Personnel_ActionNew AS tbl_Personnel_ActionNew_1 LEFT OUTER JOIN dbo.tbl_Personnel_Department ON tbl_Personnel_ActionNew_1.DeptKey = dbo.tbl_Personnel_Department.PK Where (tbl_Personnel_ActionNew_1.EffectivityDate <= dbo.tbl_Personnel_Compensation_Period.DateTo) And (tbl_Personnel_ActionNew_1.EmpPK = dbo.tbl_Personnel_Allowance_Per_Period.EmpPK) ORDER BY tbl_Personnel_ActionNew_1.EffectivityDate) AS Department, " & _
        " (SELECT TOP (1) dbo.tbl_Personnel_Position.PositionCode + ' - ' + dbo.tbl_Personnel_Position.PositionName AS Position FROM  dbo.tbl_Personnel_ActionNew AS tbl_Personnel_ActionNew_2 LEFT OUTER JOIN dbo.tbl_Personnel_Position ON tbl_Personnel_ActionNew_2.PositionsKey = dbo.tbl_Personnel_Position.PK  Where (tbl_Personnel_ActionNew_2.EffectivityDate <= dbo.tbl_Personnel_Compensation_Period.DateTo) And (tbl_Personnel_ActionNew_2.EmpPK = dbo.tbl_Personnel_Allowance_Per_Period.EmpPK) ORDER BY tbl_Personnel_ActionNew_2.EffectivityDate) AS Position, " & _
        " dbo.tbl_Personnel_Allowance_Per_Period.Amount, dbo.tbl_Personnel_Compensation_Period.DateFrom, dbo.tbl_Personnel_Compensation_Period.DateTo, dbo.tbl_Personnel_Allowance_Per_Period.LastModified, dbo.tbl_Personnel_Allowance_Per_Period.NoHours, dbo.tbl_Personnel_Allowance_Per_Period.RatePerHour " & _
        " FROM  dbo.tbl_Personnel_Allowance_Per_Period LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Allowance_Per_Period.Period = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Allowance_Per_Period.EmpPK = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
        " WHERE (dbo.tbl_Personnel_Allowance_Per_Period.Division = " & iDivision & ") " & _
        " AND (dbo.tbl_Personnel_Allowance_Per_Period.Period = " & iPeriod & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        i = i + 1
        
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Allowance_Report " & _
                          " (LogIn, EmployeeName, Department, Positions, Amount, NoHours, RateperHour) " & _
                          " VALUES ('" & gbl_UserName & "', '" & FORMATSQL(rs!EmployeeName) & "', " & _
                          " '" & FORMATSQL(rs!Department) & "', '" & FORMATSQL(rs!Position) & "', " & _
                          " " & CDbl(rs!Amount) & ", " & CDbl(rs!NoHours) & ", " & CDbl(rs!RatePerHour) & ")"
        
        UpdateProgress picProgressBar, i / rs.RecordCount
        rs.MoveNext
    Wend
    rs.Close
    
    picMain.Enabled = True
    picToolbar.Enabled = True
    picProgress.Visible = False
    
    frmCrystalReportViewer.PRINT_ALLOWANCE_REPORT gbl_CompanyName, iDivision, "PERIOD : " & txtFrom.Text & " - " & txtTo.Text, gbl_UserName
    If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
    
End If

If iType = 3 Then
    
    MainForm.CommonDialog1.CancelError = True
    On Error GoTo ErrorHandler
    MainForm.CommonDialog1.DialogTitle = "Save"
    MainForm.CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
    MainForm.CommonDialog1.ShowSave
    Filename = Trim(MainForm.CommonDialog1.Filename)
    
    On Error GoTo PG:
    
    WorkbookName = CStr(Filename)
            
    iWorkSheet = 1
    Set xlsApp = CreateObject("Excel.Application")
    xlsApp.Visible = False
    xlsApp.Workbooks.Add
    xlsApp.DisplayAlerts = False
'    xlsApp.Workbooks(1).Sheets(2).Delete
'    xlsApp.Workbooks(1).Sheets(2).Delete
    xlsApp.Workbooks(1).Sheets(iWorkSheet).Activate
    xlsApp.Workbooks(1).Sheets(iWorkSheet).Name = "B P I"
    
    RowCnt = RowCnt + 1
    ColCnt = 0
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = gbl_CompanyName
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
    
    RowCnt = RowCnt + 1
    ColCnt = 0
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "For B P I ATM"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
    
    RowCnt = RowCnt + 1
    ColCnt = 0
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
    
    RowCnt = RowCnt + 1
    For k = 1 To 5
        Select Case k
            Case 1: strValue = "Account Number"
            Case 2: strValue = "Last Name"
            Case 3: strValue = "First Name"
            Case 4: strValue = "Middle Name"
            Case 5: strValue = "Allowance"
        End Select
        ColCnt = k
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = strValue
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
        If k = 5 Then
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(ColCnt).ColumnWidth = 13
        ElseIf k = 1 Then
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(ColCnt).ColumnWidth = 20
        Else
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 2
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(ColCnt).ColumnWidth = 20
        End If
    Next k
    
    i = 0
    s = "SELECT tbl_Personnel_IDNumber.AccountNumber, " & _
        " tbl_Personnel_Information.LastName, " & _
        " tbl_Personnel_Information.FirstName, " & _
        " tbl_Personnel_Information.MiddleName, " & _
        " tbl_Personnel_Allowance_Per_Period.Amount " & _
        " FROM tbl_Personnel_Allowance_Per_Period LEFT OUTER JOIN " & _
        " tbl_Personnel_Compensation_Period ON tbl_Personnel_Allowance_Per_Period.Period = tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_IDNumber ON tbl_Personnel_Allowance_Per_Period.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
        " WHERE (tbl_Personnel_Allowance_Per_Period.Division = " & iDivision & ") " & _
        " AND (tbl_Personnel_Allowance_Per_Period.Period = " & iPeriod & ") " & _
        " ORDER BY tbl_Personnel_Information.LastName, tbl_Personnel_Information.FirstName, tbl_Personnel_Information.MiddleName "
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        i = i + 1
        RowCnt = RowCnt + 1
        For k = 1 To 5
            strValue = rs.Fields(k - 1).Value
            ColCnt = k
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = strValue
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            If k = 5 Then
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "#,##0.00"
            ElseIf k = 1 Then
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "0000-0000-00"
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3
            Else
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 2
            End If
        Next k
        UpdateProgress picProgressBar, i / rs.RecordCount
        rs.MoveNext
    Wend
    rs.Close
    
SAVING:
        On Error GoTo err_saving:
        If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
        xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName
        
        xlsApp.Visible = True
        
        picMain.Enabled = True
        picToolbar.Enabled = True
        picProgress.Visible = False
    
End If

Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
picMain.Enabled = True
picToolbar.Enabled = True
picProgress.Visible = False
Exit Sub

Exit Sub
ErrorHandler:
picMain.Enabled = True
picToolbar.Enabled = True
picProgress.Visible = False
Exit Sub

Exit Sub
err_saving:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING:

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
        Case "Refresh"
            'ToDo: Add 'Refresh' button code.
            MsgBox "Add 'Refresh' button code."
        Case "Delete"
            'ToDo: Add 'Delete' button code.
            MsgBox "Add 'Delete' button code."
        Case "Edit"
            'ToDo: Add 'Edit' button code.
            MsgBox "Add 'Edit' button code."
    Case "Add":     PRESS_INSERT
    Case "Find":    PRESS_F6
    Case "Print":   PRESS_F9
    Case "Close":   PRESS_ESCAPE
    Case "First":   BROWSER GetSetting(App.EXEName, "AllowancePK", "AllowKey", ""), "is_HOME"
    Case "Back":    BROWSER GetSetting(App.EXEName, "AllowancePK", "AllowKey", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "AllowancePK", "AllowKey", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "AllowancePK", "AllowKey", ""), "is_END"
End Select
End Sub

Private Sub txtFrom_GotFocus()
HTEXT txtFrom
End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtTo.SetFocus
End Sub

Private Sub txtFrom_LostFocus()
If IsDate(txtFrom.Text) = True Then
    txtFrom.Text = Format(FormatDateTime(txtFrom.Text, vbShortDate), "mm/dd/yyyy")
    s = "SELECT tbl_Personnel_Compensation_Period.* " & _
        " FROM tbl_Personnel_Compensation_Period " & _
        " WHERE (Type = " & cmbDivision.ListIndex + 1 & ") " & _
        " AND (DateFrom = '" & FormatDateTime(txtFrom.Text, vbShortDate) & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        txtTo.Text = Format(rs!DateTo, "mm/dd/yyyy")
    Else
        MsgBox "Invalid Range!                      ", vbCritical, "Error..."
        txtTo.Text = ""
        rs.Close
        Exit Sub
    End If
    rs.Close
End If
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstResult.Clear: Exit Sub
lstResult.Clear
s = "SELECT tbl_Personnel_Allowance_Per_Period.EmpPK, tbl_Personnel_IDNumber.IDNumber, " & _
    " tbl_Personnel_Information.LastName, tbl_Personnel_Information.FirstName, " & _
    " tbl_Personnel_Information.MiddleName " & _
    " FROM tbl_Personnel_Allowance_Per_Period LEFT OUTER JOIN " & _
    " tbl_Personnel_IDNumber ON tbl_Personnel_Allowance_Per_Period.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
    " GROUP BY tbl_Personnel_Allowance_Per_Period.EmpPK, tbl_Personnel_IDNumber.IDNumber, tbl_Personnel_Information.LastName, " & _
    " tbl_Personnel_Information.FirstName , tbl_Personnel_Information.MiddleName " & _
    " HAVING (tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
    " ORDER BY tbl_Personnel_Information.LastName, tbl_Personnel_Information.FirstName, tbl_Personnel_Information.MiddleName, " & _
    " tbl_Personnel_IDNumber.IDNumber "
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

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResult.SetFocus
End Sub
