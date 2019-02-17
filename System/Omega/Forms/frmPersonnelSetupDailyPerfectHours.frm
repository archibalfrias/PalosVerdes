VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelSetupDailyPerfectHours 
   Appearance      =   0  'Flat
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9210
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
   ScaleHeight     =   4125
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   1920
      ScaleHeight     =   1815
      ScaleWidth      =   4815
      TabIndex        =   3
      Top             =   1440
      Width           =   4815
      Begin VB.TextBox txtNoOfHours 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   12
         Top             =   1440
         Width           =   1260
      End
      Begin VB.TextBox txtPayrollCutOff 
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtPayrollDate 
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Top             =   720
         Width           =   3735
      End
      Begin VB.ComboBox cmbDivision 
         Height          =   315
         ItemData        =   "frmPersonnelSetupDailyPerfectHours.frx":0000
         Left            =   1080
         List            =   "frmPersonnelSetupDailyPerfectHours.frx":0002
         TabIndex        =   6
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox txtCtrl 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   4
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "No of Hours"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   13
         Top             =   1485
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cut Off"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   5
         Top             =   45
         Width           =   735
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
         MouseIcon       =   "frmPersonnelSetupDailyPerfectHours.frx":0004
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
      Left            =   7920
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
            Picture         =   "frmPersonnelSetupDailyPerfectHours.frx":031E
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelSetupDailyPerfectHours.frx":0FF8
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelSetupDailyPerfectHours.frx":1CD2
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelSetupDailyPerfectHours.frx":29AC
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelSetupDailyPerfectHours.frx":3686
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelSetupDailyPerfectHours.frx":4360
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelSetupDailyPerfectHours.frx":503A
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelSetupDailyPerfectHours.frx":5D14
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelSetupDailyPerfectHours.frx":69EE
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelSetupDailyPerfectHours.frx":72C8
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelSetupDailyPerfectHours.frx":7FA2
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelSetupDailyPerfectHours.frx":8C7C
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelSetupDailyPerfectHours.frx":9956
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelSetupDailyPerfectHours.frx":A630
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelSetupDailyPerfectHours.frx":B30A
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   3810
      Width           =   9210
      _ExtentX        =   16245
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
Attribute VB_Name = "frmPersonnelSetupDailyPerfectHours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE     As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim sCtrl, locDivKey, locPayrollPeriodKey

Private Sub BROWSER(Ctrl, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If Ctrl <> "" Then
            s = "SELECT TOP (1) dbo.tbl_Personnel_Setup_DailyPerfectDays.PK, dbo.tbl_Personnel_Setup_DailyPerfectDays.Ctrl, " & _
                " dbo.tbl_Personnel_Setup_DailyPerfectDays.DivisionKey, dbo.tbl_Personnel_Division.Description AS DivisionName, " & _
                " dbo.tbl_Personnel_Setup_DailyPerfectDays.PayrollPeriodKey, dbo.tbl_Personnel_Compensation_Period.DateFrom, " & _
                " dbo.tbl_Personnel_Compensation_Period.DateTo , dbo.tbl_Personnel_Compensation_Period.PayrollDate, " & _
                " dbo.tbl_Personnel_Setup_DailyPerfectDays.NoHours, dbo.tbl_Personnel_Setup_DailyPerfectDays.LastModified " & _
                " FROM  dbo.tbl_Personnel_Setup_DailyPerfectDays LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Setup_DailyPerfectDays.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_Setup_DailyPerfectDays.DivisionKey = dbo.tbl_Personnel_Division.PK " & _
                " WHERE (dbo.tbl_Personnel_Setup_DailyPerfectDays.Ctrl = '" & Ctrl & "') " & _
                " ORDER BY dbo.tbl_Personnel_Setup_DailyPerfectDays.Ctrl"
        Else
            s = "SELECT TOP (1) dbo.tbl_Personnel_Setup_DailyPerfectDays.PK, dbo.tbl_Personnel_Setup_DailyPerfectDays.Ctrl, " & _
                " dbo.tbl_Personnel_Setup_DailyPerfectDays.DivisionKey, dbo.tbl_Personnel_Division.Description AS DivisionName, " & _
                " dbo.tbl_Personnel_Setup_DailyPerfectDays.PayrollPeriodKey, dbo.tbl_Personnel_Compensation_Period.DateFrom, " & _
                " dbo.tbl_Personnel_Compensation_Period.DateTo , dbo.tbl_Personnel_Compensation_Period.PayrollDate, " & _
                " dbo.tbl_Personnel_Setup_DailyPerfectDays.NoHours, dbo.tbl_Personnel_Setup_DailyPerfectDays.LastModified " & _
                " FROM  dbo.tbl_Personnel_Setup_DailyPerfectDays LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Setup_DailyPerfectDays.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_Setup_DailyPerfectDays.DivisionKey = dbo.tbl_Personnel_Division.PK " & _
                " ORDER BY dbo.tbl_Personnel_Setup_DailyPerfectDays.Ctrl"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Setup_DailyPerfectDays.PK, dbo.tbl_Personnel_Setup_DailyPerfectDays.Ctrl, " & _
            " dbo.tbl_Personnel_Setup_DailyPerfectDays.DivisionKey, dbo.tbl_Personnel_Division.Description AS DivisionName, " & _
            " dbo.tbl_Personnel_Setup_DailyPerfectDays.PayrollPeriodKey, dbo.tbl_Personnel_Compensation_Period.DateFrom, " & _
            " dbo.tbl_Personnel_Compensation_Period.DateTo , dbo.tbl_Personnel_Compensation_Period.PayrollDate, " & _
            " dbo.tbl_Personnel_Setup_DailyPerfectDays.NoHours, dbo.tbl_Personnel_Setup_DailyPerfectDays.LastModified " & _
            " FROM  dbo.tbl_Personnel_Setup_DailyPerfectDays LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Setup_DailyPerfectDays.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_Setup_DailyPerfectDays.DivisionKey = dbo.tbl_Personnel_Division.PK " & _
            " ORDER BY dbo.tbl_Personnel_Setup_DailyPerfectDays.Ctrl"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Setup_DailyPerfectDays.PK, dbo.tbl_Personnel_Setup_DailyPerfectDays.Ctrl, " & _
            " dbo.tbl_Personnel_Setup_DailyPerfectDays.DivisionKey, dbo.tbl_Personnel_Division.Description AS DivisionName, " & _
            " dbo.tbl_Personnel_Setup_DailyPerfectDays.PayrollPeriodKey, dbo.tbl_Personnel_Compensation_Period.DateFrom, " & _
            " dbo.tbl_Personnel_Compensation_Period.DateTo , dbo.tbl_Personnel_Compensation_Period.PayrollDate, " & _
            " dbo.tbl_Personnel_Setup_DailyPerfectDays.NoHours, dbo.tbl_Personnel_Setup_DailyPerfectDays.LastModified " & _
            " FROM  dbo.tbl_Personnel_Setup_DailyPerfectDays LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Setup_DailyPerfectDays.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_Setup_DailyPerfectDays.DivisionKey = dbo.tbl_Personnel_Division.PK " & _
            " WHERE (dbo.tbl_Personnel_Setup_DailyPerfectDays.Ctrl < '" & Ctrl & "') " & _
            " ORDER BY dbo.tbl_Personnel_Setup_DailyPerfectDays.Ctrl DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Setup_DailyPerfectDays.PK, dbo.tbl_Personnel_Setup_DailyPerfectDays.Ctrl, " & _
            " dbo.tbl_Personnel_Setup_DailyPerfectDays.DivisionKey, dbo.tbl_Personnel_Division.Description AS DivisionName, " & _
            " dbo.tbl_Personnel_Setup_DailyPerfectDays.PayrollPeriodKey, dbo.tbl_Personnel_Compensation_Period.DateFrom, " & _
            " dbo.tbl_Personnel_Compensation_Period.DateTo , dbo.tbl_Personnel_Compensation_Period.PayrollDate, " & _
            " dbo.tbl_Personnel_Setup_DailyPerfectDays.NoHours, dbo.tbl_Personnel_Setup_DailyPerfectDays.LastModified " & _
            " FROM  dbo.tbl_Personnel_Setup_DailyPerfectDays LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Setup_DailyPerfectDays.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_Setup_DailyPerfectDays.DivisionKey = dbo.tbl_Personnel_Division.PK " & _
            " WHERE (dbo.tbl_Personnel_Setup_DailyPerfectDays.Ctrl > '" & Ctrl & "') " & _
            " ORDER BY dbo.tbl_Personnel_Setup_DailyPerfectDays.Ctrl"
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Setup_DailyPerfectDays.PK, dbo.tbl_Personnel_Setup_DailyPerfectDays.Ctrl, " & _
            " dbo.tbl_Personnel_Setup_DailyPerfectDays.DivisionKey, dbo.tbl_Personnel_Division.Description AS DivisionName, " & _
            " dbo.tbl_Personnel_Setup_DailyPerfectDays.PayrollPeriodKey, dbo.tbl_Personnel_Compensation_Period.DateFrom, " & _
            " dbo.tbl_Personnel_Compensation_Period.DateTo , dbo.tbl_Personnel_Compensation_Period.PayrollDate, " & _
            " dbo.tbl_Personnel_Setup_DailyPerfectDays.NoHours, dbo.tbl_Personnel_Setup_DailyPerfectDays.LastModified " & _
            " FROM  dbo.tbl_Personnel_Setup_DailyPerfectDays LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Setup_DailyPerfectDays.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_Setup_DailyPerfectDays.DivisionKey = dbo.tbl_Personnel_Division.PK " & _
            " ORDER BY dbo.tbl_Personnel_Setup_DailyPerfectDays.Ctrl DESC"
    Case "is_FIND"
        s = "SELECT TOP (1) dbo.tbl_Personnel_Setup_DailyPerfectDays.PK, dbo.tbl_Personnel_Setup_DailyPerfectDays.Ctrl, " & _
            " dbo.tbl_Personnel_Setup_DailyPerfectDays.DivisionKey, dbo.tbl_Personnel_Division.Description AS DivisionName, " & _
            " dbo.tbl_Personnel_Setup_DailyPerfectDays.PayrollPeriodKey, dbo.tbl_Personnel_Compensation_Period.DateFrom, " & _
            " dbo.tbl_Personnel_Compensation_Period.DateTo , dbo.tbl_Personnel_Compensation_Period.PayrollDate, " & _
            " dbo.tbl_Personnel_Setup_DailyPerfectDays.NoHours, dbo.tbl_Personnel_Setup_DailyPerfectDays.LastModified " & _
            " FROM  dbo.tbl_Personnel_Setup_DailyPerfectDays LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Setup_DailyPerfectDays.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_Setup_DailyPerfectDays.DivisionKey = dbo.tbl_Personnel_Division.PK " & _
            " WHERE (dbo.tbl_Personnel_Setup_DailyPerfectDays.PK = " & Ctrl & ") " & _
            " ORDER BY dbo.tbl_Personnel_Setup_DailyPerfectDays.Ctrl DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    locDivKey = rs!DivisionKey
    locPayrollPeriodKey = rs!PayrollPeriodKey
    txtCtrl.Text = rs!Ctrl
    txtPayrollDate.Text = Format(rs!PayrollDate, "mm/dd/yyyy")
    txtPayrollCutOff.Text = Format(rs!DateFrom, "mm/dd/yyyy") & " - " & Format(rs!DateTo, "mm/dd/yyyy")
    txtNoOfHours.Text = Format(rs!NoHours, "#0.00")
    cmbDivision.Text = rs!DivisionName
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    SaveSetting App.EXEName, "PerfectDaysCtrl", "PerfectDaysCtrl", rs!Ctrl
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If AccessRights("Perfect Days (Daily)", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
cmbDivision.SetFocus
End Sub

Private Sub PRESS_F2()
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If AccessRights("Perfect Days (Daily)", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_EDITTING
End Sub

Private Sub PRESS_DELETE()
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If AccessRights("Perfect Days (Daily)", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If MsgBox("ARE YOU SURE IN DELETING THIS RECORD?                        ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_Personnel_Setup_DailyPerfectDays WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "PerfectDaysCtrl", "PerfectDaysCtrl", ""), "is_PAGEDOWN"
If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "PerfectDaysCtrl", "PerfectDaysCtrl", ""), "is_HOME"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If CDbl(locDivKey) = 0 Then MsgBox "Please select Division!                     ", vbCritical, "Error...": cmbDivision.SetFocus: Exit Sub
If IsDate(txtPayrollDate.Text) = False Then MsgBox "Please supply a valid date value!                     ", vbCritical, "Error...": txtPayrollDate.SetFocus: Exit Sub
t = "SELECT PK, DateFrom, DateTo, PayrollDate " & _
    " From dbo.tbl_Personnel_Compensation_Period " & _
    " WHERE (Type = " & locDivKey & ") " & _
    " AND (PayrollDate = '" & FormatDateTime(txtPayrollDate.Text, vbShortDate) & "')"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then
    locPayrollPeriodKey = rt!PK
    txtPayrollDate.Text = Format(rt!PayrollDate, "mm/dd/yyyy")
    txtPayrollCutOff.Text = Format(rt!DateFrom, "mm/dd/yyyy") & " - " & Format(rt!DateTo, "mm/dd/yyyy")
End If
rt.Close
If CDbl(locPayrollPeriodKey) = 0 Then MsgBox "Please supply a valid payroll period!                     ", vbCritical, "Error...": txtPayrollDate.SetFocus: Exit Sub
If RETURNTEXTVALUE(txtNoOfHours) <= 0 Then MsgBox "Invalid Hours Value!                  ", vbCritical, "Error...": txtNoOfHours.SetFocus: Exit Sub


On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    sCtrl = ""
    s = "SELECT TOP (1) dbo.tbl_Personnel_Setup_DailyPerfectDays.Ctrl " & _
        " FROM  dbo.tbl_Personnel_Setup_DailyPerfectDays LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Setup_DailyPerfectDays.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
        " Where (Year(dbo.tbl_Personnel_Compensation_Period.PayrollDate) = " & Format(FormatDateTime(txtPayrollDate.Text, vbShortDate), "yyyy") & ") " & _
        " ORDER BY dbo.tbl_Personnel_Setup_DailyPerfectDays.Ctrl DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        sCtrl = Format(CDbl(rs!Ctrl) + 1, "0000000#")
    Else
        sCtrl = Format(FormatDateTime(txtPayrollDate.Text, vbShortDate), "yyyy") & "0000"
    End If
    rs.Close
    
    Do
        s = "SELECT tbl_Personnel_Setup_DailyPerfectDays.* " & _
            " FROM tbl_Personnel_Setup_DailyPerfectDays " & _
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
    
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Setup_DailyPerfectDays " & _
                      " (Ctrl, DivisionKey, PayrollPeriodKey, NoHours, LastModified) " & _
                      " VALUES ('" & sCtrl & "', " & locDivKey & ", " & locPayrollPeriodKey & ", " & _
                      " " & RETURNTEXTVALUE(txtNoOfHours) & ", '" & CStr(Now) & " - " & gbl_CompleteName & "')"
    
    
End If
If TRANSACTIONTYPE = is_EDITTING Then
    sCtrl = Trim(txtCtrl.Text)
    ConnOmega.Execute "UPDATE tbl_Personnel_Setup_DailyPerfectDays " & _
                      " SET DivisionKey = " & locDivKey & ", " & _
                      " PayrollPeriodKey = " & locPayrollPeriodKey & ", " & _
                      " NoHours = " & RETURNTEXTVALUE(txtNoOfHours) & ", " & _
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

End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    Unload Me
Else
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "PerfectDaysCtrl", "PerfectDaysCtrl", ""), "is_LOAD"
    If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "PerfectDaysCtrl", "PerfectDaysCtrl", ""), "is_HOME"
End If
End Sub

Private Sub CLEARTEXT()
locDivKey = 0
locPayrollPeriodKey = 0
txtCtrl.Text = ""
txtPayrollDate.Text = ""
txtPayrollCutOff.Text = ""
txtNoOfHours.Text = ""
cmbDivision.Text = ""
cmbDivision.ListIndex = -1
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtCtrl.Locked = True
txtPayrollDate.Locked = bln
txtPayrollCutOff.Locked = True
txtNoOfHours.Locked = bln
cmbDivision.Locked = bln
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

Private Sub cmbDivision_Click()
If cmbDivision.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    locDivKey = cmbDivision.ItemData(cmbDivision.ListIndex)
End If
End Sub

Private Sub cmbDivision_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtPayrollDate.SetFocus
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
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "PerfectDaysCtrl", "PerfectDaysCtrl", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "PerfectDaysCtrl", "PerfectDaysCtrl", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "PerfectDaysCtrl", "PerfectDaysCtrl", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "PerfectDaysCtrl", "PerfectDaysCtrl", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
POPULATE_COMBO "PK", "Description", "tbl_Personnel_Division", "Description", cmbDivision
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "PerfectDaysCtrl", "PerfectDaysCtrl", ""), "is_LOAD"
If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "PerfectDaysCtrl", "PerfectDaysCtrl", ""), "is_HOME"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "PerfectDaysCtrl", "PerfectDaysCtrl", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "PerfectDaysCtrl", "PerfectDaysCtrl", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "PerfectDaysCtrl", "PerfectDaysCtrl", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "PerfectDaysCtrl", "PerfectDaysCtrl", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtNoOfHours_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtPayrollDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtNoOfHours.SetFocus
End Sub

Private Sub txtPayrollDate_LostFocus()
If IsDate(txtPayrollDate.Text) = False Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If cmbDivision.ListIndex = -1 Then Exit Sub
    t = "SELECT PK, DateFrom, DateTo, PayrollDate " & _
        " From dbo.tbl_Personnel_Compensation_Period " & _
        " WHERE (Type = " & cmbDivision.ItemData(cmbDivision.ListIndex) & ") " & _
        " AND (PayrollDate = '" & FormatDateTime(txtPayrollDate.Text, vbShortDate) & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        locPayrollPeriodKey = rt!PK
        txtPayrollDate.Text = Format(rt!PayrollDate, "mm/dd/yyyy")
        txtPayrollCutOff.Text = Format(rt!DateFrom, "mm/dd/yyyy") & " - " & Format(rt!DateTo, "mm/dd/yyyy")
    End If
    rt.Close
End If
End Sub
