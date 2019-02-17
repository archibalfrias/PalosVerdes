VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelAllowance 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   9180
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPersonnelAllowance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picAdd 
      Height          =   2955
      Left            =   2040
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5212
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
         Picture         =   "frmPersonnelAllowance.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2355
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
         Picture         =   "frmPersonnelAllowance.frx":0F3C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2355
         Width           =   1560
      End
      Begin VB.TextBox txtSearchAdd 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   4215
      End
      Begin VB.ListBox lstResultAdd 
         Height          =   1425
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   4215
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   17
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
         Icon            =   "frmPersonnelAllowance.frx":1698
         ShadowVisible   =   0   'False
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   28
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   29
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
         MouseIcon       =   "frmPersonnelAllowance.frx":1C32
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
   Begin RPVGCC.b8Container picSearch 
      Height          =   3075
      Left            =   2040
      TabIndex        =   18
      Top             =   360
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5424
      BackColor       =   15396057
      Begin VB.ComboBox cmbEffectDate 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2100
         Width           =   1935
      End
      Begin VB.ListBox lstResult 
         Height          =   1230
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   4215
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
         Picture         =   "frmPersonnelAllowance.frx":1F4C
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2480
         Width           =   1560
      End
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
         Picture         =   "frmPersonnelAllowance.frx":26A8
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2480
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   45
         TabIndex        =   23
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
         Icon            =   "frmPersonnelAllowance.frx":2D1A
         ShadowVisible   =   0   'False
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Effectivity Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   24
         Top             =   2100
         Width           =   1215
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   840
      ScaleHeight     =   1815
      ScaleWidth      =   7335
      TabIndex        =   6
      Top             =   1320
      Width           =   7335
      Begin VB.TextBox txtDivision 
         Height          =   315
         Left            =   960
         MaxLength       =   100
         TabIndex        =   26
         Top             =   360
         Width           =   6375
      End
      Begin VB.TextBox txtEffectDate 
         Height          =   315
         Left            =   3720
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtAmount 
         Height          =   315
         Left            =   960
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtPosition 
         Height          =   315
         Left            =   960
         MaxLength       =   100
         TabIndex        =   2
         Top             =   1080
         Width           =   6375
      End
      Begin VB.TextBox txtDepartment 
         Height          =   315
         Left            =   960
         MaxLength       =   100
         TabIndex        =   1
         Top             =   720
         Width           =   6375
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   960
         MaxLength       =   100
         TabIndex        =   0
         Top             =   0
         Width           =   6375
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Effectivity Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   1470
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   1470
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   30
         Width           =   1095
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   3525
      Width           =   9180
      _ExtentX        =   16193
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
      Left            =   11520
      Top             =   1440
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
            Picture         =   "frmPersonnelAllowance.frx":32B4
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowance.frx":3F8E
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowance.frx":4C68
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowance.frx":5942
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowance.frx":661C
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowance.frx":72F6
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowance.frx":7FD0
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowance.frx":8CAA
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowance.frx":9984
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowance.frx":A25E
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowance.frx":AF38
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowance.frx":BC12
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowance.frx":C8EC
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowance.frx":D5C6
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAllowance.frx":E2A0
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPersonnelAllowance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim tmp As Long

Dim iPK, iEmpPK, dRatePerHour, iCompensationRate

Private Sub BROWSER(iPK, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If iPK <> "" Then
            s = "SELECT TOP 1 tbl_Personnel_Allowance.PK, " & _
                " tbl_Personnel_IDNumber.IDNumber + ' - ' + tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
                " ISNULL((SELECT TOP 1 tbl_Personnel_Department.DepartmentName FROM tbl_Personnel_Action LEFT OUTER JOIN tbl_Personnel_Department ON tbl_Personnel_Action.Dept = tbl_Personnel_Department.PK " & _
                " WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC), '') AS Department, " & _
                " ISNULL((SELECT TOP 1 tbl_Personnel_Position.PositionName FROM tbl_Personnel_Action LEFT OUTER JOIN tbl_Personnel_Position ON tbl_Personnel_Action.Positions = tbl_Personnel_Position.PK " & _
                " WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC), '') AS Positions, " & _
                " (SELECT TOP 1 tbl_Personnel_Action.CompensationRate From tbl_Personnel_Action WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC) AS CompensationRate, " & _
                " (SELECT TOP 1 tbl_Personnel_Action.Division From tbl_Personnel_Action WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC) AS Division, " & _
                " tbl_Personnel_Allowance.Rate, tbl_Personnel_Allowance.EffectDate , tbl_Personnel_Allowance.LastModified, tbl_Personnel_Allowance.EmpPK " & _
                " FROM tbl_Personnel_Allowance LEFT OUTER JOIN " & _
                " tbl_Personnel_IDNumber ON tbl_Personnel_Allowance.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
                " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
                " WHERE (tbl_Personnel_Allowance.PK = " & iPK & ") " & _
                " ORDER BY tbl_Personnel_Allowance.PK"
        Else
            s = "SELECT TOP 1 tbl_Personnel_Allowance.PK, " & _
                " tbl_Personnel_IDNumber.IDNumber + ' - ' + tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
                " ISNULL((SELECT TOP 1 tbl_Personnel_Department.DepartmentName FROM tbl_Personnel_Action LEFT OUTER JOIN tbl_Personnel_Department ON tbl_Personnel_Action.Dept = tbl_Personnel_Department.PK " & _
                " WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC), '') AS Department, " & _
                " ISNULL((SELECT TOP 1 tbl_Personnel_Position.PositionName FROM tbl_Personnel_Action LEFT OUTER JOIN tbl_Personnel_Position ON tbl_Personnel_Action.Positions = tbl_Personnel_Position.PK " & _
                " WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC), '') AS Positions, " & _
                " (SELECT TOP 1 tbl_Personnel_Action.CompensationRate From tbl_Personnel_Action WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC) AS CompensationRate, " & _
                " (SELECT TOP 1 tbl_Personnel_Action.Division From tbl_Personnel_Action WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC) AS Division, " & _
                " tbl_Personnel_Allowance.Rate, tbl_Personnel_Allowance.EffectDate , tbl_Personnel_Allowance.LastModified, tbl_Personnel_Allowance.EmpPK " & _
                " FROM tbl_Personnel_Allowance LEFT OUTER JOIN " & _
                " tbl_Personnel_IDNumber ON tbl_Personnel_Allowance.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
                " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
                " ORDER BY tbl_Personnel_Allowance.PK"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picAdd.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Personnel_Allowance.PK, " & _
            " tbl_Personnel_IDNumber.IDNumber + ' - ' + tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
            " ISNULL((SELECT TOP 1 tbl_Personnel_Department.DepartmentName FROM tbl_Personnel_Action LEFT OUTER JOIN tbl_Personnel_Department ON tbl_Personnel_Action.Dept = tbl_Personnel_Department.PK " & _
            " WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC), '') AS Department, " & _
            " ISNULL((SELECT TOP 1 tbl_Personnel_Position.PositionName FROM tbl_Personnel_Action LEFT OUTER JOIN tbl_Personnel_Position ON tbl_Personnel_Action.Positions = tbl_Personnel_Position.PK " & _
            " WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC), '') AS Positions, " & _
            " (SELECT TOP 1 tbl_Personnel_Action.CompensationRate From tbl_Personnel_Action WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC) AS CompensationRate, " & _
            " (SELECT TOP 1 tbl_Personnel_Action.Division From tbl_Personnel_Action WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC) AS Division, " & _
            " tbl_Personnel_Allowance.Rate, tbl_Personnel_Allowance.EffectDate , tbl_Personnel_Allowance.LastModified, tbl_Personnel_Allowance.EmpPK " & _
            " FROM tbl_Personnel_Allowance LEFT OUTER JOIN " & _
            " tbl_Personnel_IDNumber ON tbl_Personnel_Allowance.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
            " ORDER BY tbl_Personnel_Allowance.PK"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picAdd.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Personnel_Allowance.PK, " & _
            " tbl_Personnel_IDNumber.IDNumber + ' - ' + tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
            " ISNULL((SELECT TOP 1 tbl_Personnel_Department.DepartmentName FROM tbl_Personnel_Action LEFT OUTER JOIN tbl_Personnel_Department ON tbl_Personnel_Action.Dept = tbl_Personnel_Department.PK " & _
            " WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC), '') AS Department, " & _
            " ISNULL((SELECT TOP 1 tbl_Personnel_Position.PositionName FROM tbl_Personnel_Action LEFT OUTER JOIN tbl_Personnel_Position ON tbl_Personnel_Action.Positions = tbl_Personnel_Position.PK " & _
            " WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC), '') AS Positions, " & _
            " (SELECT TOP 1 tbl_Personnel_Action.CompensationRate From tbl_Personnel_Action WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC) AS CompensationRate, " & _
            " (SELECT TOP 1 tbl_Personnel_Action.Division From tbl_Personnel_Action WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC) AS Division, " & _
            " tbl_Personnel_Allowance.Rate, tbl_Personnel_Allowance.EffectDate , tbl_Personnel_Allowance.LastModified, tbl_Personnel_Allowance.EmpPK " & _
            " FROM tbl_Personnel_Allowance LEFT OUTER JOIN " & _
            " tbl_Personnel_IDNumber ON tbl_Personnel_Allowance.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
            " WHERE (tbl_Personnel_Allowance.PK < " & iPK & ") " & _
            " ORDER BY tbl_Personnel_Allowance.PK DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picAdd.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Personnel_Allowance.PK, " & _
            " tbl_Personnel_IDNumber.IDNumber + ' - ' + tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
            " ISNULL((SELECT TOP 1 tbl_Personnel_Department.DepartmentName FROM tbl_Personnel_Action LEFT OUTER JOIN tbl_Personnel_Department ON tbl_Personnel_Action.Dept = tbl_Personnel_Department.PK " & _
            " WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC), '') AS Department, " & _
            " ISNULL((SELECT TOP 1 tbl_Personnel_Position.PositionName FROM tbl_Personnel_Action LEFT OUTER JOIN tbl_Personnel_Position ON tbl_Personnel_Action.Positions = tbl_Personnel_Position.PK " & _
            " WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC), '') AS Positions, " & _
            " (SELECT TOP 1 tbl_Personnel_Action.CompensationRate From tbl_Personnel_Action WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC) AS CompensationRate, " & _
            " (SELECT TOP 1 tbl_Personnel_Action.Division From tbl_Personnel_Action WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC) AS Division, " & _
            " tbl_Personnel_Allowance.Rate, tbl_Personnel_Allowance.EffectDate , tbl_Personnel_Allowance.LastModified, tbl_Personnel_Allowance.EmpPK " & _
            " FROM tbl_Personnel_Allowance LEFT OUTER JOIN " & _
            " tbl_Personnel_IDNumber ON tbl_Personnel_Allowance.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
            " WHERE (tbl_Personnel_Allowance.PK > " & iPK & ") " & _
            " ORDER BY tbl_Personnel_Allowance.PK"
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picAdd.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Personnel_Allowance.PK, " & _
            " tbl_Personnel_IDNumber.IDNumber + ' - ' + tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
            " ISNULL((SELECT TOP 1 tbl_Personnel_Department.DepartmentName FROM tbl_Personnel_Action LEFT OUTER JOIN tbl_Personnel_Department ON tbl_Personnel_Action.Dept = tbl_Personnel_Department.PK " & _
            " WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC), '') AS Department, " & _
            " ISNULL((SELECT TOP 1 tbl_Personnel_Position.PositionName FROM tbl_Personnel_Action LEFT OUTER JOIN tbl_Personnel_Position ON tbl_Personnel_Action.Positions = tbl_Personnel_Position.PK " & _
            " WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC), '') AS Positions, " & _
            " (SELECT TOP 1 tbl_Personnel_Action.CompensationRate From tbl_Personnel_Action WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC) AS CompensationRate, " & _
            " (SELECT TOP 1 tbl_Personnel_Action.Division From tbl_Personnel_Action WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_Allowance.EmpPK) AND (tbl_Personnel_Action.EffectivityDate <= tbl_Personnel_Allowance.EffectDate) ORDER BY tbl_Personnel_Action.EffectivityDate DESC) AS Division, " & _
            " tbl_Personnel_Allowance.Rate, tbl_Personnel_Allowance.EffectDate, tbl_Personnel_Allowance.LastModified, tbl_Personnel_Allowance.EmpPK " & _
            " FROM tbl_Personnel_Allowance LEFT OUTER JOIN " & _
            " tbl_Personnel_IDNumber ON tbl_Personnel_Allowance.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
            " ORDER BY tbl_Personnel_Allowance.PK DESC"
    Case Else: Exit Sub
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iEmpPK = rs!EmpPK
    iCompensationRate = rs!CompensationRate
    txtName.Text = rs!EmployeeName
    txtDivision.Text = IIf(rs!Division = 1, "CLUB HOUSE", "MAINTENANCE")
    txtDepartment.Text = rs!Department
    txtPosition.Text = rs!Positions
    txtAmount.Text = Format(rs!Rate, "#,##0.00")
    txtEffectDate.Text = Format(rs!Effectdate, "mm/dd/yyyy")
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    SaveSetting App.EXEName, "AllowancePK", "AllowanceKey", rs!PK
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
txtSearchAdd.Text = ""
picAdd.ZOrder 0
picAdd.Visible = True
txtSearchAdd.SetFocus
End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_EDITTING
End Sub

Private Sub PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If MsgBox("ARE YOU SURE IN DELETING THIS RECORD?                    ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_Personnel_Allowance WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
CLEARTEXT
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If IsDate(txtEffectDate.Text) = False Then MsgBox "Please Supply a Valid Date!                    ", vbCritical, "Error...": txtEffectDate.SetFocus: Exit Sub
If RETURNTEXTVALUE(txtAmount) < 0 Then MsgBox "Invalid Amount!                       ", vbCritical, "Error...": txtAmount.SetFocus: Exit Sub
On Error GoTo PG:
If CInt(iCompensationRate) = 1 Then
    dRatePerHour = ((CDbl(RETURNTEXTVALUE(txtAmount)) / 2) / 13.08333) / 8
ElseIf CInt(iCompensationRate) = 2 Then
    dRatePerHour = CDbl(RETURNTEXTVALUE(txtAmount)) / 8
End If
If TRANSACTIONTYPE = is_ADDING Then
    If CDbl(iEmpPK) = 0 Then MsgBox "Please Select Employee!                          ", vbCritical, "Error...": Exit Sub
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Allowance " & _
                      " (EmpPK, EffectDate, Rate, RatePerHour, LastModified) " & _
                      " VALUES (" & iEmpPK & ", '" & FormatDateTime(txtEffectDate.Text, vbShortDate) & "', " & _
                      " " & RETURNTEXTVALUE(txtAmount) & ", " & CDbl(dRatePerHour) & ", '" & CStr(Now) & " - " & gbl_CompleteName & "')"
    iPK = 0
    s = "SELECT TOP 1 PK " & _
        " FROM tbl_Personnel_Allowance " & _
        " ORDER BY PK DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        iPK = rs!PK
    End If
    rs.Close
End If
If TRANSACTIONTYPE = is_EDITTING Then
    iPK = Statusbar1.Panels(1).Text
    ConnOmega.Execute "UPDATE tbl_Personnel_Allowance " & _
                      " SET EffectDate = '" & FormatDateTime(txtEffectDate.Text, vbShortDate) & "', " & _
                      " Rate = " & RETURNTEXTVALUE(txtAmount) & ", " & _
                      " RatePerHour = " & CDbl(dRatePerHour) & ", " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & " )"
End If
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER iPK, "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
picSearch.ZOrder 0
txtSearch.Text = ""
picSearch.Visible = True
txtSearch.SetFocus
End Sub

Private Sub PRESS_F9()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picAdd.Visible = True Then cmdCancelAdd_Click: Exit Sub
    If picSearch.Visible = True Then cmdCancel_Click: Exit Sub
    Unload Me
Else
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "AllowancePK", "AllowanceKey", 0), "is_LOAD"
    If Trim(txtName.Text) = "" Then BROWSER GetSetting(App.EXEName, "AllowancePK", "AllowanceKey", 0), "is_HOME"
End If
End Sub


Private Sub CLEARTEXT()
txtName.Text = ""
txtDivision.Text = ""
txtDepartment.Text = ""
txtPosition.Text = ""
txtAmount.Text = ""
txtEffectDate.Text = ""
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtName.Locked = True
txtDivision.Locked = True
txtDepartment.Locked = True
txtPosition.Locked = True
txtAmount.Locked = bln
txtEffectDate.Locked = bln
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

Private Sub b8TitleBar2_CLoseClick()
cmdCancelAdd_Click
End Sub

Private Sub cmbEffectDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub cmdCancel_Click()
picToolbar.Enabled = True
picMain.Enabled = True
picSearch.Visible = False
End Sub

Private Sub cmdCancelAdd_Click()
picAdd.Visible = False
picToolbar.Enabled = True
picMain.Enabled = True
End Sub

Private Sub cmdOK_Click()
If cmbEffectDate.ListIndex = -1 Then Exit Sub
cmdCancel_Click
BROWSER cmbEffectDate.ItemData(cmbEffectDate.ListIndex), "is_LOAD"
End Sub

Private Sub cmdOKAdd_Click()
If lstResultAdd.ListIndex = -1 Then Exit Sub
iEmpPK = lstResultAdd.ItemData(lstResultAdd.ListIndex)
s = "SELECT TOP 1 tbl_Personnel_Department.DepartmentName, " & _
    " tbl_Personnel_Position.PositionName, " & _
    " tbl_Personnel_Action.CompensationRate, " & _
    " tbl_Personnel_Action.Division " & _
    " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
    " tbl_Personnel_Position ON tbl_Personnel_Action.Positions = tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
    " tbl_Personnel_Department ON tbl_Personnel_Action.Dept = tbl_Personnel_Department.PK " & _
    " Where (tbl_Personnel_Action.EmpPK = " & iEmpPK & ") " & _
    " ORDER BY tbl_Personnel_Action.EffectivityDate DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    CLEARTEXT
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_ADDING
    iCompensationRate = rs!CompensationRate
    txtName.Text = lstResultAdd.List(lstResultAdd.ListIndex)
    txtDivision.Text = IIf(rs!Division = 1, "CLUB HOUSE", "MAINTENANCE")
    txtDepartment.Text = rs!DepartmentName
    txtPosition.Text = rs!PositionName
    cmdCancelAdd_Click
    txtAmount.SetFocus
End If
rs.Close
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
    Case vbKeyF9:
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "AllowancePK", "AllowanceKey", 0), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "AllowancePK", "AllowanceKey", 0), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "AllowancePK", "AllowanceKey", 0), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "AllowancePK", "AllowanceKey", 0), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.Height - Me.Height) / 3
Me.Left = (MainForm.Width - Me.Width) / 5
iEmpPK = 0
iCompensationRate = 0
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "AllowancePK", "AllowanceKey", 0), "is_LOAD"
If Trim(txtName.Text) = "" Then BROWSER GetSetting(App.EXEName, "AllowancePK", "AllowanceKey", 0), "is_HOME"

tmp = SetWindowLong(txtName.hwnd, GWL_STYLE, GetWindowLong(txtName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearchAdd.hwnd, GWL_STYLE, GetWindowLong(txtSearchAdd.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picAdd.Visible = True Then Cancel = -1
If picSearch.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstResult_Click()
If lstResult.ListIndex = -1 Then cmbEffectDate.Clear: Exit Sub
cmbEffectDate.Clear
t = "SELECT PK, EffectDate " & _
    " FROM tbl_Personnel_Allowance " & _
    " WHERE (EmpPK = " & lstResult.ItemData(lstResult.ListIndex) & ") " & _
    " ORDER BY EffectDate DESC"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
While Not rt.EOF
    cmbEffectDate.AddItem Format(rt!Effectdate, "mm/dd/yyyy")
    cmbEffectDate.ItemData(cmbEffectDate.NewIndex) = rt!PK
    rt.MoveNext
Wend
rt.Close
If cmbEffectDate.ListCount Then cmbEffectDate.ListIndex = 0
End Sub

Private Sub lstResult_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbEffectDate.SetFocus
End Sub

Private Sub lstResultAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAdd_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
        Case "Refresh"
            'ToDo: Add 'Refresh' button code.
            MsgBox "Add 'Refresh' button code."
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "AllowancePK", "AllowanceKey", 0), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "AllowancePK", "AllowanceKey", 0), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "AllowancePK", "AllowanceKey", 0), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "AllowancePK", "AllowanceKey", 0), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":
    Case "Close":   PRESS_ESCAPE
End Select
End Sub



Private Sub txtAmount_GotFocus()
HTEXT txtAmount
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtEffectDate.SetFocus
End Sub

Private Sub txtAmount_LostFocus()
txtAmount.Text = RETURNTEXTVALUE(txtAmount)
End Sub

Private Sub txtEffectDate_GotFocus()
HTEXT txtEffectDate
End Sub

Private Sub txtEffectDate_LostFocus()
If IsDate(txtEffectDate.Text) = True Then
    txtEffectDate.Text = Format(FormatDateTime(txtEffectDate.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstResult.Clear: Exit Sub
lstResult.Clear
s = "SELECT tbl_Personnel_Allowance.EmpPK, " & _
    " tbl_Personnel_IDNumber.IDNumber + ' - ' + tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName " & _
    " FROM tbl_Personnel_Allowance LEFT OUTER JOIN " & _
    " tbl_Personnel_IDNumber ON tbl_Personnel_Allowance.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
    " WHERE (tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
    " GROUP BY tbl_Personnel_Allowance.EmpPK, " & _
    " tbl_Personnel_IDNumber.IDNumber + ' - ' + tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName " & _
    " ORDER BY tbl_Personnel_IDNumber.IDNumber + ' - ' + tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResult.AddItem rs!EmployeeName
    lstResult.ItemData(lstResult.NewIndex) = rs!EmpPK
    rs.MoveNext
Wend
rs.Close
If lstResult.ListCount Then lstResult.ListIndex = 0
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResult.SetFocus
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

Private Sub txtSearchAdd_GotFocus()
HTEXT txtSearchAdd
End Sub

Private Sub txtSearchAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultAdd.SetFocus
End Sub
