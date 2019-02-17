VERSION 5.00
Begin VB.Form frmPersonnelCompensationReport 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPersonnelCompensationReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picTaxWithHeldAlpha 
      Height          =   1815
      Left            =   1680
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3201
      BackColor       =   15396057
      Begin VB.TextBox txtTaxAlphaYear 
         Height          =   315
         Left            =   1680
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelTaxAlpha 
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
         Picture         =   "frmPersonnelCompensationReport.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1080
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKTaxAlpha 
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
         Picture         =   "frmPersonnelCompensationReport.frx":2EFE
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1080
         Width           =   1560
      End
      Begin VB.ComboBox cmbDivisionAlpha 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2640
         Width           =   3495
      End
      Begin RPVGCC.b8TitleBar b8TitleBar4 
         Height          =   345
         Left            =   40
         TabIndex        =   21
         Top             =   40
         Width           =   3890
         _ExtentX        =   6853
         _ExtentY        =   609
         Caption         =   "Tax WithHeld (Alpha List)"
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
         Icon            =   "frmPersonnelCompensationReport.frx":3570
         ShadowVisible   =   0   'False
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "SELECT DIVISION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2400
         Width           =   3375
      End
   End
   Begin RPVGCC.b8Container picProgress 
      Height          =   975
      Left            =   960
      TabIndex        =   14
      Top             =   2640
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1720
      BackColor       =   15396057
      Begin VB.PictureBox picProgressBar 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5235
         TabIndex        =   15
         Top             =   120
         Width           =   5295
      End
   End
   Begin VB.PictureBox picPrint 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   120
      ScaleHeight     =   6375
      ScaleWidth      =   7215
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.TextBox txtTerms 
         Height          =   315
         Left            =   3480
         TabIndex        =   9
         Top             =   5880
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.ComboBox cmbPeriodPrint 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   6000
         Width           =   3375
      End
      Begin VB.ComboBox cmbGroup 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   3735
      End
      Begin VB.ComboBox cmbDivision 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   5280
         Width           =   3375
      End
      Begin VB.ListBox lstReportType 
         Height          =   4740
         Left            =   0
         TabIndex        =   4
         Top             =   240
         Width           =   3375
      End
      Begin VB.CommandButton cmdCancelPrint 
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
         Left            =   5400
         Picture         =   "frmPersonnelCompensationReport.frx":3B0A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5835
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKPrint 
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
         Left            =   3720
         Picture         =   "frmPersonnelCompensationReport.frx":4266
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5835
         Width           =   1560
      End
      Begin VB.ListBox lstResultPrint 
         Height          =   4740
         Left            =   3480
         TabIndex        =   1
         Top             =   1035
         Width           =   3735
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   3480
         TabIndex        =   24
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox txtSearchPrint 
         Height          =   315
         Left            =   3480
         TabIndex        =   7
         Top             =   675
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "GROUP BY"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3480
         TabIndex        =   13
         Top             =   0
         Width           =   3375
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "PAYROLL PERIOD"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   5760
         Width           =   1335
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "SELECT DIVISION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   5040
         Width           =   3375
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "REPORT TYPE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmPersonnelCompensationReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmp             As Long


Dim Filename        As String
Dim WorkbookName    As String
Dim iWorkSheet      As Integer

Dim RowCnt, ColCnt, strRange, iReset, strGrossTot, staTaxTot, strSSSTot, strPHICTot, _
strHDMFTot, strColaTot, strAllowTot, strValue, iCnt, strRange1, strRange2, sTaxStatus, _
i, j, k, PK, iDivision, Arr, iMonth1, iMonth2, iMonth3, iYear1, iYear2, iYear3, _
dtmDateTo, sDeptName, Array1, dtmTo

Private Sub LOAD_GROUP_BY(intIndex)
With cmbGroup
    .Clear
    If intIndex = 0 Then
        .AddItem "DEPARTMENT"
        .AddItem "STATUS"
        .AddItem "POSITION"
        .AddItem "EMPLOYEE"
    Else
        .AddItem cmbDivision.Text
    End If
    .ListIndex = 0
End With
End Sub

Private Sub PAYROLL_TEMP_RF_SV(Division, Period, PostLevel)

picProgressBar.BackColor = &HFFFFFF
picPrint.Enabled = False
picProgress.ZOrder 0
picProgress.Visible = True
i = 0

ConnOmega.Execute "DELETE FROM tbl_PersonnelPayroll_Tmp WHERE (LogInName = '" & gbl_UserName & "')"

s = "sp_Personnel_Compensation_Print_RF_SV(" & Division & ", " & Period & ", " & PostLevel & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
While Not ra.EOF
    DoEvents
    i = i + 1
    
    ConnOmega.Execute "INSERT INTO tbl_PersonnelPayroll_Tmp " & _
                      " (LogInName, EmpPK, Division, Dept, Status, Positions, Period, ActionMemo, NoHours, SH_Hours, LH_Hours, SL_Hours, Adjustment, Reg_OT_Hours, " & _
                      " RD_OT_Hours, SH_OT_Hours, LH_OT_Hours, Amount_Earned, SH_Amount, LH_Amount, SL_Amount, Reg_OT_Amount, RD_OT_Amount, SH_OT_Amount, LH_OT_Amount, " & _
                      " TotalEarning, Mortuary, AR_Others, Advances, Shortages, Uniforms, Others, Is_Have_Loan, SSSLoan_No, SSSLoan, SSSBalance, PagIbigLoan_No, " & _
                      " PagIbigLoan, PagIbigBalance, Is_Have_Cont, SSS, SSS_Employer, SSS_EC, PHIC, PHIC_Employer, PagIbig, PagIbig_Employer, WithHeld, TotalDeduction, " & _
                      " NetEarning, Locked, IDNumber, LName, FName, MName, BDate, DepartmentName, StatusName, PositionName, DateFrom, DateTo, Type, CompensationRate, " & _
                      " SSSNo, Is_PHIC, PHICNo, IDName, Is_TIN, TIN, Basic, RatePerHour, TotalCola, TotalAllowance, PostLevel) " & _
                      " VALUES ('" & gbl_UserName & "', " & ra!EmpPK & ", " & ra!Division & ", " & ra!Dept & ", " & ra!Status & ", " & ra!Positions & ", " & ra!Period & ", " & ra!ActionMemo & ", " & _
                      " " & CDbl(ra!NoHours) & ", " & CDbl(ra!SH_Hours) & ", " & CDbl(ra!LH_Hours) & ", " & CDbl(ra!SL_Hours) & ", " & CDbl(ra!Adjustment) & ", " & _
                      " " & CDbl(ra!Reg_OT_Hours) & ", " & CDbl(ra!RD_OT_Hours) & ", " & CDbl(ra!SH_OT_Hours) & ", " & CDbl(ra!LH_OT_Hours) & ", " & CDbl(ra!Amount_Earned) & ", " & _
                      " " & CDbl(ra!SH_Amount) & ", " & CDbl(ra!LH_Amount) & ", " & CDbl(ra!SL_Amount) & ", " & CDbl(ra!Reg_OT_Amount) & ", " & CDbl(ra!RD_OT_Amount) & ", " & _
                      " " & CDbl(ra!SH_OT_Amount) & ", " & CDbl(ra!LH_OT_Amount) & ", " & CDbl(ra!TotalEarning) & ", " & CDbl(ra!Mortuary) & ", " & CDbl(ra!AR_Others) & ", " & _
                      " " & CDbl(ra!Advances) & ", " & CDbl(ra!Shortages) & ", " & CDbl(ra!Uniforms) & ", " & CDbl(ra!Others) & ", " & ra!Is_Have_Loan & ", " & _
                      " " & ra!SSSLoan_No & ", " & CDbl(ra!SSSLoan) & ", " & CDbl(ra!SSSBalance) & ", " & ra!PagIbigLoan_No & ", " & CDbl(ra!PagIbigLoan) & ", " & _
                      " " & CDbl(ra!PagIbigBalance) & ", " & ra!Is_Have_Cont & ", " & CDbl(ra!SSS) & ", " & CDbl(ra!SSS_Employer) & ", " & CDbl(ra!SSS_EC) & ", " & CDbl(ra!PHIC) & ", " & _
                      " " & CDbl(ra!PHIC_Employer) & ", " & CDbl(ra!PAGIBIG) & ", " & CDbl(ra!PagIbig_Employer) & ", " & CDbl(ra!WithHeld) & ", " & CDbl(ra!TotalDeduction) & ", " & _
                      " " & CDbl(ra!NetEarning) & ", " & ra!Locked & ", '" & ra!IDNumber & "', '" & FORMATSQL(ra!LName) & "', '" & FORMATSQL(ra!FName) & "', '" & FORMATSQL(ra!MName) & "', " & _
                      " '" & FormatDateTime(ra!BDate, vbShortDate) & "' , '" & FORMATSQL(ra!DepartmentName) & "', '" & FORMATSQL(ra!StatusName) & "', '" & FORMATSQL(ra!PositionName) & "', " & _
                      " '" & FormatDateTime(ra!DateFrom, vbShortDate) & "', '" & FormatDateTime(ra!DateTo, vbShortDate) & "', " & ra!Type & ", " & ra!CompensationRate & ", " & _
                      " '" & ra!SSSNo & "', " & ra!Is_PHIC & ", '" & ra!PHICNo & "', '" & FORMATSQL(ra!IDName) & "', " & ra!Is_TIN & ", '" & ra!TIN & "', " & CDbl(ra!Basic) & ", " & _
                      " " & CDbl(ra!RatePerHour) & ", " & CDbl(ra!TotalCola) & ", " & CDbl(ra!TotalAllowance) & ", " & ra!PositionLevel & ")"
    
    UpdateProgress picProgressBar, i / ra.RecordCount
    ra.MoveNext
Wend
ra.Close

s = "SELECT tbl_PersonnelPayroll_Tmp.* " & _
    " FROM tbl_PersonnelPayroll_Tmp"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
ra.Requery
ra.Close

picProgressBar.BackColor = &HFFFFFF
picProgress.Visible = False
picPrint.Enabled = True

End Sub

Private Sub PAYROLL_TEMP(Division, Period)

picProgressBar.BackColor = &HFFFFFF
picPrint.Enabled = False
picProgress.ZOrder 0
picProgress.Visible = True
i = 0

ConnOmega.Execute "DELETE FROM tbl_PersonnelPayroll_Tmp WHERE (LogInName = '" & gbl_UserName & "')"

s = "sp_Personnel_Compensation_Print(" & Division & ", " & Period & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
While Not ra.EOF
    DoEvents
    i = i + 1
    
    ConnOmega.Execute "INSERT INTO tbl_PersonnelPayroll_Tmp " & _
                      " (LogInName, EmpPK, Division, Dept, Status, Positions, Period, ActionMemo, NoHours, SH_Hours, LH_Hours, SL_Hours, Adjustment, Reg_OT_Hours, " & _
                      " RD_OT_Hours, SH_OT_Hours, LH_OT_Hours, Amount_Earned, SH_Amount, LH_Amount, SL_Amount, Reg_OT_Amount, RD_OT_Amount, SH_OT_Amount, LH_OT_Amount, " & _
                      " TotalEarning, Mortuary, AR_Others, Advances, Shortages, Uniforms, Others, Is_Have_Loan, SSSLoan_No, SSSLoan, SSSBalance, PagIbigLoan_No, " & _
                      " PagIbigLoan, PagIbigBalance, Is_Have_Cont, SSS, SSS_Employer, SSS_EC, PHIC, PHIC_Employer, PagIbig, PagIbig_Employer, WithHeld, TotalDeduction, " & _
                      " NetEarning, Locked, IDNumber, LName, FName, MName, BDate, DepartmentName, StatusName, PositionName, DateFrom, DateTo, Type, CompensationRate, " & _
                      " SSSNo, Is_PHIC, PHICNo, IDName, Is_TIN, TIN, Basic, RatePerHour, TotalCola, TotalAllowance) " & _
                      " VALUES ('" & gbl_UserName & "', " & ra!EmpPK & ", " & ra!Division & ", " & ra!Dept & ", " & ra!Status & ", " & ra!Positions & ", " & ra!Period & ", " & ra!ActionMemo & ", " & _
                      " " & CDbl(ra!NoHours) & ", " & CDbl(ra!SH_Hours) & ", " & CDbl(ra!LH_Hours) & ", " & CDbl(ra!SL_Hours) & ", " & CDbl(ra!Adjustment) & ", " & _
                      " " & CDbl(ra!Reg_OT_Hours) & ", " & CDbl(ra!RD_OT_Hours) & ", " & CDbl(ra!SH_OT_Hours) & ", " & CDbl(ra!LH_OT_Hours) & ", " & CDbl(ra!Amount_Earned) & ", " & _
                      " " & CDbl(ra!SH_Amount) & ", " & CDbl(ra!LH_Amount) & ", " & CDbl(ra!SL_Amount) & ", " & CDbl(ra!Reg_OT_Amount) & ", " & CDbl(ra!RD_OT_Amount) & ", " & _
                      " " & CDbl(ra!SH_OT_Amount) & ", " & CDbl(ra!LH_OT_Amount) & ", " & CDbl(ra!TotalEarning) & ", " & CDbl(ra!Mortuary) & ", " & CDbl(ra!AR_Others) & ", " & _
                      " " & CDbl(ra!Advances) & ", " & CDbl(ra!Shortages) & ", " & CDbl(ra!Uniforms) & ", " & CDbl(ra!Others) & ", " & ra!Is_Have_Loan & ", " & _
                      " " & ra!SSSLoan_No & ", " & CDbl(ra!SSSLoan) & ", " & CDbl(ra!SSSBalance) & ", " & ra!PagIbigLoan_No & ", " & CDbl(ra!PagIbigLoan) & ", " & _
                      " " & CDbl(ra!PagIbigBalance) & ", " & ra!Is_Have_Cont & ", " & CDbl(ra!SSS) & ", " & CDbl(ra!SSS_Employer) & ", " & CDbl(ra!SSS_EC) & ", " & CDbl(ra!PHIC) & ", " & _
                      " " & CDbl(ra!PHIC_Employer) & ", " & CDbl(ra!PAGIBIG) & ", " & CDbl(ra!PagIbig_Employer) & ", " & CDbl(ra!WithHeld) & ", " & CDbl(ra!TotalDeduction) & ", " & _
                      " " & CDbl(ra!NetEarning) & ", " & ra!Locked & ", '" & ra!IDNumber & "', '" & FORMATSQL(ra!LName) & "', '" & FORMATSQL(ra!FName) & "', '" & FORMATSQL(ra!MName) & "', " & _
                      " '" & FormatDateTime(ra!BDate, vbShortDate) & "' , '" & FORMATSQL(ra!DepartmentName) & "', '" & FORMATSQL(ra!StatusName) & "', '" & FORMATSQL(ra!PositionName) & "', " & _
                      " '" & FormatDateTime(ra!DateFrom, vbShortDate) & "', '" & FormatDateTime(ra!DateTo, vbShortDate) & "', " & ra!Type & ", " & ra!CompensationRate & ", " & _
                      " '" & ra!SSSNo & "', " & ra!Is_PHIC & ", '" & ra!PHICNo & "', '" & FORMATSQL(ra!IDName) & "', " & ra!Is_TIN & ", '" & ra!TIN & "', " & CDbl(ra!Basic) & ", " & CDbl(ra!RatePerHour) & ", " & CDbl(ra!TotalCola) & ", " & _
                      " " & CDbl(ra!TotalAllowance) & ")"
    
    UpdateProgress picProgressBar, i / ra.RecordCount
    ra.MoveNext
Wend
ra.Close

s = "SELECT tbl_PersonnelPayroll_Tmp.* " & _
    " FROM tbl_PersonnelPayroll_Tmp"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
ra.Requery
ra.Close

picProgressBar.BackColor = &HFFFFFF
picProgress.Visible = False
picPrint.Enabled = True


End Sub

Private Sub b8TitleBar4_CLoseClick()
cmdCancelTaxAlpha_Click
End Sub

Private Sub cmbDivision_Click()
If cmbDivision.ListIndex = -1 Then cmbPeriodPrint.Clear: txtTerms.Text = "": Exit Sub

cmbGroup_Click

LOAD_GROUP_BY lstReportType.ListIndex
Array1 = Split(FIND_PAYROLL_PERIOD(Date, cmbDivision.ListIndex + 1), ";", -1)
txtTerms.Text = CLng(Array1(3))
cmbPeriodPrint.Clear
If lstReportType.ListIndex = 19 Then
    s = "SELECT TOP 1 tbl_Personnel_Compensation_Period.DateFrom, " & _
        " tbl_Personnel_Compensation_Period.DateTo, " & _
        " tbl_Personnel_Compensation.Period " & _
        " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
        " tbl_Personnel_Compensation_Period ON tbl_Personnel_Compensation.Period = tbl_Personnel_Compensation_Period.PK " & _
        " Where (tbl_Personnel_Compensation.Division = " & cmbDivision.ListIndex + 1 & ") " & _
        " GROUP BY tbl_Personnel_Compensation_Period.DateFrom, tbl_Personnel_Compensation_Period.DateTo, tbl_Personnel_Compensation.Period " & _
        " ORDER BY tbl_Personnel_Compensation_Period.DateFrom DESC, tbl_Personnel_Compensation_Period.DateTo DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        dtmTo = rs!DateTo
    End If
    rs.Close
    
    s = "SELECT TOP 1 PK, DateFrom, DateTo " & _
        " From tbl_Personnel_Compensation_Period " & _
        " WHERE (Type = " & cmbDivision.ListIndex + 1 & ") " & _
        " AND (DateTo > '" & FormatDateTime(dtmTo, vbShortDate) & "') " & _
        " ORDER BY DateFrom, DateTo"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        cmbPeriodPrint.AddItem Format(rs!DateFrom, "mm/dd/yyyy") & " - " & Format(rs!DateTo, "mm/dd/yyyy")
        cmbPeriodPrint.ItemData(cmbPeriodPrint.NewIndex) = rs!PK
        rs.MoveNext
    Wend
    rs.Close
End If
s = "SELECT tbl_Personnel_Compensation_Period.DateFrom, " & _
    " tbl_Personnel_Compensation_Period.DateTo, " & _
    " tbl_Personnel_Compensation.Period " & _
    " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
    " tbl_Personnel_Compensation_Period ON tbl_Personnel_Compensation.Period = tbl_Personnel_Compensation_Period.PK " & _
    " Where (tbl_Personnel_Compensation.Division = " & cmbDivision.ListIndex + 1 & ") " & _
    " GROUP BY tbl_Personnel_Compensation_Period.DateFrom, tbl_Personnel_Compensation_Period.DateTo, tbl_Personnel_Compensation.Period " & _
    " ORDER BY tbl_Personnel_Compensation_Period.DateFrom DESC, tbl_Personnel_Compensation_Period.DateTo DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbPeriodPrint.AddItem Format(rs!DateFrom, "mm/dd/yyyy") & " - " & Format(rs!DateTo, "mm/dd/yyyy")
    cmbPeriodPrint.ItemData(cmbPeriodPrint.NewIndex) = rs!Period
    dtmTo = rs!DateTo
    rs.MoveNext
Wend
rs.Close

If cmbPeriodPrint.ListCount Then cmbPeriodPrint.ListIndex = 0
End Sub


Private Sub cmbDivision_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbPeriodPrint.SetFocus
End Sub

Private Sub cmbGroup_Click()
If cmbGroup.ListIndex = -1 Then Exit Sub
If lstReportType.ListIndex = -1 Then Exit Sub
If cmbPeriodPrint.ListIndex = -1 Then Exit Sub

If lstReportType.ListIndex = 0 Then
    If cmbGroup.ListIndex <> 3 Then
        Select Case cmbGroup.ListIndex
            Case 0
                s = "SELECT tbl_Personnel_Compensation.Dept as PK, " & _
                    " tbl_Personnel_Department.DepartmentName as Name " & _
                    " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
                    " tbl_Personnel_Department ON tbl_Personnel_Compensation.Dept = tbl_Personnel_Department.PK " & _
                    " Where (tbl_Personnel_Compensation.Period = " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ") " & _
                    " GROUP BY tbl_Personnel_Compensation.Dept, tbl_Personnel_Department.DepartmentName " & _
                    " ORDER BY tbl_Personnel_Department.DepartmentName"
            Case 1
                s = "SELECT tbl_Personnel_Compensation.Status AS PK, " & _
                    " tbl_Personnel_EmploymentStatus.StatusName AS Name " & _
                    " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
                    " tbl_Personnel_EmploymentStatus ON tbl_Personnel_Compensation.Status = tbl_Personnel_EmploymentStatus.PK " & _
                    " Where (tbl_Personnel_Compensation.Period = " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ") " & _
                    " GROUP BY tbl_Personnel_Compensation.Status, tbl_Personnel_EmploymentStatus.StatusName " & _
                    " ORDER BY tbl_Personnel_EmploymentStatus.StatusName"
            Case 2
                s = "SELECT tbl_Personnel_Compensation.Positions as PK, " & _
                    " tbl_Personnel_Position.PositionName as Name " & _
                    " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
                    " tbl_Personnel_Position ON tbl_Personnel_Compensation.Positions = tbl_Personnel_Position.PK " & _
                    " Where (tbl_Personnel_Compensation.Period = " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ") " & _
                    " GROUP BY tbl_Personnel_Compensation.Positions, tbl_Personnel_Position.PositionName " & _
                    " ORDER BY tbl_Personnel_Position.PositionName"
        End Select
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        txtSearchPrint.Text = ""
        txtSearch.Visible = False
        lstResultPrint.Top = picPrint.Top + 630 '1155
        lstResultPrint.Height = 4935 '4350 '4155
        With lstResultPrint
            .Clear
            While Not rs.EOF
                .AddItem rs!Name
                .ItemData(.NewIndex) = rs!PK
                rs.MoveNext
            Wend
            .AddItem "SUPERVISORY"
            .ItemData(.NewIndex) = 0
            If .ListCount Then .ListIndex = 0
        End With
        rs.Close
        
    Else
        txtSearchPrint.Text = ""
        txtSearchPrint.Visible = True
        lstResultPrint.Clear
        lstResultPrint.Top = picPrint.Top + 1010 '1560
        lstResultPrint.Height = 4545 '3765
    End If
    
Else
    txtSearchPrint.Text = ""
    txtSearchPrint.Visible = False
    lstResultPrint.Clear
    lstResultPrint.Top = picPrint.Top + 630 '1155
    lstResultPrint.Height = 4935 '4350 '4155
End If
End Sub

Private Sub cmbPeriodPrint_Click()
cmbGroup_Click
End Sub

Private Sub cmdCancelPrint_Click()
Unload Me
End Sub

Private Sub cmdCancelTaxAlpha_Click()
picTaxWithHeldAlpha.Visible = False
picPrint.Enabled = True
End Sub

Private Sub cmdOKPrint_Click()
If cmbPeriodPrint.ListIndex = -1 Then Exit Sub
txtTerms.Text = GET_TERMS(cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex))
Select Case lstReportType.ListIndex
    Case 0  'PAYSLIP
        
        Select Case cmbGroup.ListIndex
            Case 0  'DEPT
                If lstResultPrint.ItemData(lstResultPrint.ListIndex) = 0 Then
                    PAYROLL_TEMP_RF_SV cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex), 2
                    frmCrystalReportViewer.PRINT_PAYSLIP_SUPERVISORY gbl_CompanyName, gbl_UserName, 2
                    If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
                Else
                    PAYROLL_TEMP_RF_SV cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex), 1
                    frmCrystalReportViewer.PRINT_PAYSLIP_DEPT gbl_CompanyName, gbl_UserName, lstResultPrint.ItemData(lstResultPrint.ListIndex)
                    If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
                End If
            Case 1  'STATUS
                If lstResultPrint.ItemData(lstResultPrint.ListIndex) = 0 Then
                    PAYROLL_TEMP_RF_SV cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex), 2
                    frmCrystalReportViewer.PRINT_PAYSLIP_SUPERVISORY gbl_CompanyName, gbl_UserName, 2
                    If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
                Else
                    PAYROLL_TEMP_RF_SV cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex), 1
                    frmCrystalReportViewer.PRINT_PAYSLIP_STATUS gbl_CompanyName, gbl_UserName, lstResultPrint.ItemData(lstResultPrint.ListIndex)
                    If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
                End If
                
            Case 2  'POSITION
                If lstResultPrint.ItemData(lstResultPrint.ListIndex) = 0 Then
                    PAYROLL_TEMP_RF_SV cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex), 2
                    frmCrystalReportViewer.PRINT_PAYSLIP_SUPERVISORY gbl_CompanyName, gbl_UserName, 2
                    If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
                Else
                    PAYROLL_TEMP_RF_SV cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex), 1
                    frmCrystalReportViewer.PRINT_PAYSLIP_POST gbl_CompanyName, gbl_UserName, lstResultPrint.ItemData(lstResultPrint.ListIndex)
                If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
                End If
                
            Case 3  'EMPLOYEE
                frmCrystalReportViewer.PRINT_PAYSLIP_EMPLOYEE gbl_CompanyName, gbl_UserName, lstResultPrint.ItemData(lstResultPrint.ListIndex)
                If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
            Case Else: Exit Sub
        End Select
    Case 1  'SIGNATURE LEDGER
        
        PopupMenu MainFormPopupF.mnuCompensationPrint, , picPrint.Left + cmdOKPrint.Left + 200, picPrint.Top + cmdOKPrint.Top + 200
        
    Case 2  'COMPENSATION (TOP)
        
        PopupMenu MainFormPopupF.mnuCompensationPrint, , picPrint.Left + cmdOKPrint.Left + 200, picPrint.Top + cmdOKPrint.Top + 200
        
    Case 3  'COMPENSATION
        
        PopupMenu MainFormPopupF.mnuCompensationPrint, , picPrint.Left + cmdOKPrint.Left + 200, picPrint.Top + cmdOKPrint.Top + 200
        
    Case 4  'DEDUCTION (TOP)
    
        PopupMenu MainFormPopupF.mnuCompensationPrint, , picPrint.Left + cmdOKPrint.Left + 200, picPrint.Top + cmdOKPrint.Top + 200
                
    Case 5  'DEDUCTION
        
        PopupMenu MainFormPopupF.mnuCompensationPrint, , picPrint.Left + cmdOKPrint.Left + 200, picPrint.Top + cmdOKPrint.Top + 200
                
    Case 6  'SSS LOANS TOP
    
    Case 7  'SSS LOANS
        
        PAYROLL_TEMP cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex)
        frmCrystalReportViewer.PRINT_SSS_LOAN gbl_CompanyName, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
        
    Case 8  'PAG IBIG LOANS TOP
            
    Case 9  'PAG IBIG LOANS
        
        PAYROLL_TEMP cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex)
        frmCrystalReportViewer.PRINT_PAGIBIG_LOAN gbl_CompanyName, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
        
    Case 10 'SSS TOP
    
    Case 11 'SSS
        
        PAYROLL_TEMP cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex)
        frmCrystalReportViewer.PRINT_SSS_COLLECTION gbl_CompanyName, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
        
    Case 12 'PHIC TOP
    
    Case 13 'PHIC
        
        PAYROLL_TEMP cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex)
        frmCrystalReportViewer.PRINT_PHIC_COLLECTION gbl_CompanyName, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
        
    Case 14 'PAG IBIG TOP
    
    Case 15 'PAG IBIG
    
        PAYROLL_TEMP cmbDivision.ListIndex + 1, cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex)
        frmCrystalReportViewer.PRINT_PAGIBIG_COLLECTION gbl_CompanyName, gbl_CompanyTelNo, gbl_CompanySSSNo, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
        
    Case 16 'WITHHELD TOP
        
        picTaxWithHeldAlpha.ZOrder 0
        picPrint.Enabled = False
        txtTaxAlphaYear.Text = Year(Now)
        picTaxWithHeldAlpha.Visible = True
        txtTaxAlphaYear.SetFocus
        
    Case 17 'WITHHELD
        
        
        picProgressBar.BackColor = &HFFFFFF
        picPrint.Enabled = False
        picProgress.ZOrder 0
        picProgress.Visible = True
        i = 0
        
        ConnOmega.Execute "DELETE FROM tbl_Personnel_TaxWithHeld WHERE(LogInName = '" & gbl_UserName & "')"
        
        s = "SELECT ISNULL(tbl_Personnel_Department.DepartmentName, '') AS DepartmentName, " & _
            " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
            " tbl_Personnel_Information.TIN as TinNumber, tbl_Personnel_Compensation.TotalEarning, tbl_Personnel_Compensation.WithHeld, " & _
            " tbl_Personnel_Compensation.Division, tbl_Personnel_Compensation.Period, tbl_Personnel_Compensation_Period.DateFrom, " & _
            " tbl_Personnel_Compensation_Period.DateTo , tbl_Personnel_Compensation.EmpPK " & _
            " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
            " tbl_Personnel_Action ON tbl_Personnel_Compensation.ActionMemo = tbl_Personnel_Action.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_IDNumber ON tbl_Personnel_Compensation.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_Department ON tbl_Personnel_Compensation.Dept = tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_Compensation_Period ON " & _
            " tbl_Personnel_Compensation.Period = tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
            " WHERE (tbl_Personnel_Compensation.Is_Have_Cont = 1) " & _
            " AND (tbl_Personnel_Compensation.Division = " & cmbDivision.ListIndex + 1 & ") " & _
            " AND (tbl_Personnel_Compensation.Period = " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ") " & _
            " AND (tbl_Personnel_Action.Is_TIN = 1) " & _
            " ORDER BY tbl_Personnel_Department.DepartmentName, " & _
            " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName"
        If ra.State = adStateOpen Then ra.Close
        ra.Open s, ConnOmega
        While Not ra.EOF
            DoEvents
            i = i + 1
            
            ConnOmega.Execute "INSERT INTO tbl_Personnel_TaxWithHeld" & _
                              " (LogInName, Department, sTIN, sName, Gross1, Gross2, Tax, DateFrom, DateTo)" & _
                              " VALUES('" & gbl_UserName & "', '" & FORMATSQL(ra!DepartmentName) & "', " & _
                              " '" & ra!TinNumber & "', '" & FORMATSQL(ra!EmployeeName) & "', " & _
                              " " & CDbl(ra!TotalEarning) & ", " & GET_PREVIOUS_GROSS(ra!Period, cmbDivision.ListIndex + 1, ra!EmpPK) & ", " & _
                              " " & CDbl(ra!WithHeld) & ", '" & FormatDateTime(ra!DateFrom, vbShortDate) & "', " & _
                              " '" & FormatDateTime(ra!DateTo, vbShortDate) & "')"
            
            UpdateProgress picProgressBar, i / ra.RecordCount
            ra.MoveNext
        Wend
        ra.Close
        
        picProgressBar.BackColor = &HFFFFFF
        picProgress.Visible = False
        picPrint.Enabled = True
        
        frmCrystalReportViewer.PRINT_TAX_COLLECTION cmbDivision.ListIndex + 1, gbl_CompanyName, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
        
    Case 18 '13th MONTH (TOP SHEET)
    
    Case 19 '13th MONTH
        
        Arr = Split(cmbPeriodPrint.List(cmbPeriodPrint.ListIndex), " - ", -1, 1)
        
        Select Case Month(FormatDateTime(Arr(1), vbShortDate))
            Case 12
                iMonth1 = 9
                iMonth2 = 10
                iMonth3 = 11
                iYear1 = Year(FormatDateTime(Arr(1), vbShortDate))
                iYear2 = Year(FormatDateTime(Arr(1), vbShortDate))
                iYear3 = Year(FormatDateTime(Arr(1), vbShortDate))
            Case 3
                iMonth1 = 12
                iMonth2 = 1
                iMonth3 = 2
                iYear1 = Year(FormatDateTime(Arr(1), vbShortDate)) - 1
                iYear2 = Year(FormatDateTime(Arr(1), vbShortDate))
                iYear3 = Year(FormatDateTime(Arr(1), vbShortDate))
            Case 6
                iMonth1 = 3
                iMonth2 = 4
                iMonth3 = 5
                iYear1 = Year(FormatDateTime(Arr(1), vbShortDate))
                iYear2 = Year(FormatDateTime(Arr(1), vbShortDate))
                iYear3 = Year(FormatDateTime(Arr(1), vbShortDate))
            Case 9
                iMonth1 = 6
                iMonth2 = 7
                iMonth3 = 8
                iYear1 = Year(FormatDateTime(Arr(1), vbShortDate))
                iYear2 = Year(FormatDateTime(Arr(1), vbShortDate))
                iYear3 = Year(FormatDateTime(Arr(1), vbShortDate))
            Case Else: Exit Sub
        End Select
        
        picProgressBar.BackColor = &HFFFFFF
        picPrint.Enabled = False
        picProgress.ZOrder 0
        picProgress.Visible = True
        i = 0
        
        dtmDateTo = DateSerial(iYear3, iMonth3 + 1, 0)
        
        ConnOmega.Execute "DELETE FROM tbl_Personnel_13thMonth WHERE (LogInName = '" & gbl_UserName & "')"
        
        s = "sp_13th_Month_Report(" & iMonth1 & ", " & iYear1 & ", " & iMonth2 & ", " & iYear2 & ", " & iMonth3 & ", " & iYear3 & ", " & cmbDivision.ListIndex + 1 & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        While Not rs.EOF
            DoEvents
            i = i + 1
            sDeptName = "" 'rs!DepartmentName
            
            t = "SELECT TOP 1 tbl_Personnel_Department.DepartmentName " & _
                " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
                " tbl_Personnel_Compensation_Period ON " & _
                " tbl_Personnel_Compensation.Period = tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " tbl_Personnel_Department ON tbl_Personnel_Compensation.Dept = tbl_Personnel_Department.PK " & _
                " WHERE (tbl_Personnel_Compensation.EmpPK = " & rs!EmpPK & ") " & _
                " AND (tbl_Personnel_Compensation_Period.DateTo <= '" & DateSerial(iYear3, iMonth3 + 1, 0) & "') " & _
                " ORDER BY tbl_Personnel_Compensation_Period.DateTo DESC"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                sDeptName = rt!DepartmentName
            End If
            rt.Close
            
            ConnOmega.Execute "INSERT INTO tbl_Personnel_13thMonth " & _
                              " (LogInName, Department, IDNumber, sName, " & _
                              " Basic1, Basic2, Basic3) " & _
                              " VALUES ('" & gbl_UserName & "', " & _
                              " '" & FORMATSQL(CStr(sDeptName)) & "', '" & rs!IDNumber & "', " & _
                              " '" & FORMATSQL(rs!EmployeeName) & "', " & CDbl(rs!iMonth1) & ", " & _
                              " " & CDbl(rs!iMonth2) & ", " & CDbl(rs!iMonth3) & ")"
            
            UpdateProgress_Caption rs!EmployeeName, picProgressBar, i / rs.RecordCount
            rs.MoveNext
        Wend
        rs.Close
        
        picProgressBar.BackColor = &HFFFFFF
        picProgress.Visible = False
        picPrint.Enabled = True
        
        frmCrystalReportViewer.PRINT_13TH_MONTH cmbDivision.ListIndex + 1, Month(FormatDateTime(Arr(1), vbShortDate)), gbl_CompanyName, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
    
    Case 20     'Cola Summary
        
        picProgressBar.BackColor = &HFFFFFF
        picPrint.Enabled = False
        picProgress.ZOrder 0
        picProgress.Visible = True
        i = 0

        ConnOmega.Execute "DELETE FROM tbl_PersonnelPayroll_Cola_Tmp WHERE (LogInName = '" & gbl_UserName & "')"
        
        s = "sp_Personnel_Compensation_Print_Cola(" & cmbDivision.ListIndex + 1 & ", " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        While Not rs.EOF
            i = i + 1
            
            ConnOmega.Execute "INSERT INTO tbl_PersonnelPayroll_Cola_Tmp " & _
                              " (LogInName, EmpPK, Division, Dept, Status, Positions, " & _
                              " Period, ActionMemo, IDNumber, LName, FName, MName, " & _
                              " BDate, DepartmentName, StatusName, PositionName, " & _
                              " DateFrom, DateTo, ColaPerDay, ColaPerHour, " & _
                              " ColaHour, TotalCola) " & _
                              " VALUES ('" & gbl_UserName & "', " & rs!EmpPK & ", " & _
                              " " & rs!Division & ", " & rs!Dept & ", " & rs!Status & ", " & _
                              " " & rs!Positions & ", " & rs!Period & ", " & rs!ActionMemo & ", " & _
                              " '" & rs!IDNumber & "', '" & FORMATSQL(rs!LName) & "', " & _
                              " '" & FORMATSQL(rs!FName) & "', '" & FORMATSQL(rs!MName) & "', " & _
                              " '" & FormatDateTime(rs!BDate, vbShortDate) & "' , " & _
                              " '" & FORMATSQL(rs!DepartmentName) & "', '" & FORMATSQL(rs!StatusName) & "', " & _
                              " '" & FORMATSQL(rs!PositionName) & "', '" & FormatDateTime(rs!DateFrom, vbShortDate) & "', " & _
                              " '" & FormatDateTime(rs!DateTo, vbShortDate) & "', " & _
                              " " & CDbl(rs!ColaPerDay) & ", " & CDbl(rs!ColaPerHour) & ", " & _
                              " " & CDbl(rs!ColaHours) & ", " & CDbl(rs!TotalCola) & ")"
            
            UpdateProgress picProgressBar, i / rs.RecordCount
            rs.MoveNext
        Wend
        rs.Close
        
        picProgressBar.BackColor = &HFFFFFF
        picProgress.Visible = False
        picPrint.Enabled = True
        
        If i > 0 Then
            frmCrystalReportViewer.PRINT_COLA_SUMMARY gbl_CompanyName, gbl_UserName
            If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
        End If
        
        
    Case 21     'To B P I
        
        PopupMenu MainFormPopupF.mnuCompensationPrint, , picPrint.Left + cmdOKPrint.Left + 200, picPrint.Top + cmdOKPrint.Top + 200
                
    Case Else: Exit Sub
End Select
Exit Sub
ErrorHandler:
Exit Sub

End Sub

Private Sub cmdOKTaxAlpha_Click()
If RETURNTEXTVALUE(txtTaxAlphaYear) <= 0 Then Exit Sub

MainForm.CommonDialog1.CancelError = True
On Error GoTo ErrorHandler
MainForm.CommonDialog1.DialogTitle = "Save"
MainForm.CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
MainForm.CommonDialog1.ShowSave
Filename = Trim(MainForm.CommonDialog1.Filename)

WorkbookName = CStr(Filename)

On Error GoTo PG:

picTaxWithHeldAlpha.Visible = False
picProgressBar.BackColor = &HFFFFFF
picProgress.ZOrder 0
picProgress.Visible = True
i = 0

ConnOmega.Execute "DELETE FROM tbl_Personnel_Tax_Alphalist WHERE (LogInName = '" & gbl_UserName & "')"

s = "SELECT tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
    " tbl_Personnel_Information.TIN " & _
    " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
    " tbl_Personnel_IDNumber ON tbl_Personnel_Compensation.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK LEFT OUTER JOIN " & _
    " tbl_Personnel_Compensation_Period ON tbl_Personnel_Compensation.Period = tbl_Personnel_Compensation_Period.PK " & _
    " Where (Year(tbl_Personnel_Compensation_Period.DateTo) = " & RETURNTEXTVALUE(txtTaxAlphaYear) & ") " & _
    " GROUP BY tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName, " & _
    " tbl_Personnel_Information.TIN " & _
    " ORDER BY tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    DoEvents
    i = i + 1
    iDivision = 0
    t = "SELECT TOP 1 tbl_Personnel_Action.Division " & _
        " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
        " tbl_Personnel_IDNumber ON tbl_Personnel_Action.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
        " WHERE (YEAR(tbl_Personnel_Action.EffectivityDate) <= " & RETURNTEXTVALUE(txtTaxAlphaYear) & ") " & _
        " AND (tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName = '" & FORMATSQL(rs!EmployeeName) & "') " & _
        " ORDER BY tbl_Personnel_Action.EffectivityDate DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        iDivision = rt!Division
    End If
    rt.Close
    
    sTaxStatus = ""
    t = "SELECT TOP 1 tbl_Personnel_TaxStatus.TaxStatus " & _
        " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
        " tbl_Personnel_IDNumber ON tbl_Personnel_Action.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_TaxStatus ON tbl_Personnel_Action.TaxStatus = tbl_Personnel_TaxStatus.PK " & _
        " WHERE (YEAR(tbl_Personnel_Action.EffectivityDate) <= " & RETURNTEXTVALUE(txtTaxAlphaYear) & ") " & _
        " AND (tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName = '" & FORMATSQL(rs!EmployeeName) & "') " & _
        " ORDER BY tbl_Personnel_Action.EffectivityDate DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        sTaxStatus = rt!TaxStatus
    End If
    rt.Close
    
    For j = 1 To 12
        t = "SELECT SUM(tbl_Personnel_Compensation.TotalEarning) AS Gross, " & _
            " SUM(tbl_Personnel_Compensation.SSS) AS SSS, " & _
            " SUM(tbl_Personnel_Compensation.PHIC) AS PHIC, " & _
            " SUM(tbl_Personnel_Compensation.PagIbig) AS PagIbig, " & _
            " SUM(tbl_Personnel_Compensation.WithHeld) AS WithHeld, " & _
            " SUM(tbl_Personnel_Compensation.TotalCola) AS Cola, " & _
            " SUM(tbl_Personnel_Compensation.TotalAllowance) AS Allowance " & _
            " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
            " tbl_Personnel_Compensation_Period ON " & _
            " tbl_Personnel_Compensation.Period = tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_IDNumber ON tbl_Personnel_Compensation.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
            " WHERE (tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName = '" & FORMATSQL(rs!EmployeeName) & "') " & _
            " AND (YEAR(tbl_Personnel_Compensation_Period.DateTo) = " & RETURNTEXTVALUE(txtTaxAlphaYear) & ") " & _
            " AND (MONTH(tbl_Personnel_Compensation_Period.DateTo) = " & j & ")"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            u = "SELECT tbl_Personnel_Tax_Alphalist.* " & _
                " FROM tbl_Personnel_Tax_Alphalist " & _
                " WHERE (LogInName = '" & gbl_UserName & "') " & _
                " AND (EmployeeName = '" & FORMATSQL(rs!EmployeeName) & "')"
            If ru.State = adStateOpen Then ru.Close
            ru.Open u, ConnOmega
            If ru.RecordCount = 0 Then
                ConnOmega.Execute "INSERT INTO tbl_Personnel_Tax_Alphalist " & _
                                  " (LogInName, EmployeeName, Tin, TaxStatus, Division) " & _
                                  " VALUES ('" & gbl_UserName & "', '" & FORMATSQL(rs!EmployeeName) & "', '" & rs!TIN & "', '" & FORMATSQL(CStr(sTaxStatus)) & "', " & iDivision & ")"
            End If
            ru.Close
            
            ConnOmega.Execute "UPDATE tbl_Personnel_Tax_Alphalist " & _
                                      " SET " & "Gross" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!Gross), 0, rt!Gross)) & ", " & _
                                      " " & "Tax" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!WithHeld), 0, rt!WithHeld)) & ", " & _
                                      " " & "SSS" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!SSS), 0, rt!SSS)) & ", " & _
                                      " " & "PHIC" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!PHIC), 0, rt!PHIC)) & ", " & _
                                      " " & "HDMF" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!PAGIBIG), 0, rt!PAGIBIG)) & ", " & _
                                      " " & "Cola" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!Cola), 0, rt!Cola)) & ", " & _
                                      " " & "Allow" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!Allowance), 0, rt!Allowance)) & " " & _
                                      " WHERE (LogInName = '" & gbl_UserName & "') " & _
                                      " AND (EmployeeName = '" & FORMATSQL(rs!EmployeeName) & "')"
            
            
        End If
        rt.Close
        
    Next j
    
    UpdateProgress_Caption rs!EmployeeName, picProgressBar, i / rs.RecordCount
    rs.MoveNext
Wend
rs.Close


picProgressBar.BackColor = &HFFFFFF

i = 0: RowCnt = 0: iDivision = 0: iCnt = 0

iWorkSheet = 1
Set xlsApp = CreateObject("Excel.Application")
xlsApp.Visible = False
xlsApp.Workbooks.Add
xlsApp.DisplayAlerts = False
If xlsApp.Workbooks(1).Sheets.Count = 3 Then
    xlsApp.Workbooks(1).Sheets(2).Delete
    xlsApp.Workbooks(1).Sheets(2).Delete
End If
xlsApp.Workbooks(1).Sheets(iWorkSheet).Activate
xlsApp.Workbooks(1).Sheets(iWorkSheet).Name = "AlphaList"

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = gbl_CompanyName
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Alpha List for the year " & txtTaxAlphaYear.Text
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(ColCnt).ColumnWidth = 3
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4

ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(ColCnt).ColumnWidth = 1
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 2

ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(ColCnt).ColumnWidth = 30
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 2

ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3

ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3

For k = 1 To 12
    For j = 1 To 7
        ColCnt = ColCnt + 1
        Select Case j
            Case 1: strValue = "": strRange1 = EXCEL_RANGE(ColCnt, RowCnt)
            Case 2: strValue = ""
            Case 3: strValue = IIf(CDbl(k) = 1, "J A N U A R Y", _
                               IIf(CDbl(k) = 2, "F E B R U A R Y", _
                               IIf(CDbl(k) = 3, "M A R C H", _
                               IIf(CDbl(k) = 4, "A P R I L", _
                               IIf(CDbl(k) = 5, "M A Y", _
                               IIf(CDbl(k) = 6, "J U N E", _
                               IIf(CDbl(k) = 7, "J U L Y", _
                               IIf(CDbl(k) = 8, "A U G U S T", _
                               IIf(CDbl(k) = 9, "S E P T E M B E R", _
                               IIf(CDbl(k) = 10, "O C T O B E R", _
                               IIf(CDbl(k) = 11, "N O V E M B E R", _
                               IIf(CDbl(k) = 12, "D E C E M B E R", ""))))))))))))
            Case 4: strValue = ""
            Case 5: strValue = ""
            Case 6: strValue = ""
            Case 7: strValue = "": strRange2 = EXCEL_RANGE(ColCnt, RowCnt)
        End Select
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = strValue
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3
    Next j
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1, strRange2).Select
    xlsApp.Selection.Merge
    If k = 1 Or k = 3 Or k = 5 Or k = 7 Or k = 9 Or k = 11 Then
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).Interior.ColorIndex = 15
    Else
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).Interior.ColorIndex = 28
    End If
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).Font.Color = vbRed
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).HorizontalAlignment = 3
Next k

ColCnt = ColCnt + 1
strRange1 = EXCEL_RANGE(ColCnt, RowCnt)
ColCnt = ColCnt + 6
strRange2 = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1, strRange2).Select
xlsApp.Selection.Merge
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).Value = "TOTAL"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).HorizontalAlignment = 3
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).Interior.ColorIndex = 15
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).Font.Color = vbRed
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange1).HorizontalAlignment = 3
    
RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "#"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(ColCnt).ColumnWidth = 3
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4

ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(ColCnt).ColumnWidth = 1
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 2

ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Employee Name"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(ColCnt).ColumnWidth = 30
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 2

ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "TIN"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3

ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Tax Status"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3

For k = 1 To 12
    For j = 1 To 7
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        Select Case j
            Case 1: strValue = "Gross": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 10
            Case 2: strValue = "Tax": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 12
            Case 3: strValue = "SSS": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 14
            Case 4: strValue = "PHIC": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 16
            Case 5: strValue = "HDMF": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 17
            Case 6: strValue = "Cola": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 18
            Case 7: strValue = "Allowance": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 19
        End Select
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = strValue
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
    Next j
Next k

For j = 1 To 7
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    Select Case j
        Case 1: strValue = "Gross": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 10
        Case 2: strValue = "Tax": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 12
        Case 3: strValue = "SSS": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 14
        Case 4: strValue = "PHIC": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 16
        Case 5: strValue = "HDMF": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 17
        Case 6: strValue = "Cola": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 18
        Case 7: strValue = "Allowance": xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Interior.ColorIndex = 19
    End Select
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = strValue
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
Next j


s = "SELECT Division, EmployeeName, Tin, TaxStatus, Gross01, Tax01, SSS01, PHIC01, HDMF01, Cola01, Allow01, Gross02, Tax02, SSS02, PHIC02, HDMF02, Cola02, " & _
    " Allow02, Gross03, Tax03, SSS03, PHIC03, HDMF03, Cola03, Allow03, Gross04, Tax04, SSS04, PHIC04, HDMF04, Cola04, Allow04, Gross05, Tax05, " & _
    " SSS05, PHIC05, HDMF05, Cola05, Allow05, Gross06, Tax06, SSS06, PHIC06, HDMF06, Cola06, Allow06, Gross07, Tax07, SSS07, PHIC07, HDMF07, " & _
    " Cola07, Allow07, Gross08, Tax08, SSS08, PHIC08, HDMF08, Cola08, Allow08, Gross09, Tax09, SSS09, PHIC09, HDMF09, Cola09, Allow09, Gross10, " & _
    " Tax10, SSS10, PHIC10, HDMF10, Cola10, Allow10, Gross11, Tax11, SSS11, PHIC11, HDMF11, Cola11, Allow11, Gross12, Tax12, SSS12, PHIC12, " & _
    " HDMF12 , Cola12, Allow12 " & _
    " From tbl_Personnel_Tax_Alphalist " & _
    " WHERE (LogInName = '" & gbl_UserName & "') " & _
    " ORDER BY Division, EmployeeName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    DoEvents
    i = i + 1
    iReset = 0: strGrossTot = "=": staTaxTot = "=": strSSSTot = "=": strPHICTot = "="
    strHDMFTot = "=": strColaTot = "=": strAllowTot = "="
    
    If CDbl(iDivision) <> CDbl(rs!Division) Then
        If iDivision <> 0 Then
            RowCnt = RowCnt + 1
            ColCnt = 0
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
        End If
        iDivision = rs!Division
        iCnt = 0
    End If
    
    iCnt = iCnt + 1
    RowCnt = RowCnt + 1
    ColCnt = 0
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = iCnt
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "."
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
    
    For j = 1 To rs.Fields.Count - 1
        Select Case j
            Case 1
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs.Fields(j).Value
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(ColCnt).ColumnWidth = 30
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Color = vbBlue
            Case 2
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs.Fields(j).Value
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Color = vbBlue
            Case 3
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs.Fields(j).Value
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Color = vbRed
            Case Else
                ColCnt = ColCnt + 1
                iReset = iReset + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "#,##0.00"
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs.Fields(j).Value
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
                Select Case iReset
                    Case 1: strGrossTot = strGrossTot & strRange & "+"
                    Case 2: staTaxTot = staTaxTot & strRange & "+"
                    Case 3: strSSSTot = strSSSTot & strRange & "+"
                    Case 4: strPHICTot = strPHICTot & strRange & "+"
                    Case 5: strHDMFTot = strHDMFTot & strRange & "+"
                    Case 6: strColaTot = strColaTot & strRange & "+"
                    Case 7: strAllowTot = strAllowTot & strRange & "+"
                End Select
                If iReset = 7 Then
                    iReset = 0
                End If
        End Select
    Next j
    
    For j = 1 To 7
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        Select Case j
            Case 1: strValue = Mid(strGrossTot, 1, Len(strGrossTot) - 1)
            Case 2: strValue = Mid(staTaxTot, 1, Len(staTaxTot) - 1)
            Case 3: strValue = Mid(strSSSTot, 1, Len(strSSSTot) - 1)
            Case 4: strValue = Mid(strPHICTot, 1, Len(strPHICTot) - 1)
            Case 5: strValue = Mid(strHDMFTot, 1, Len(strHDMFTot) - 1)
            Case 6: strValue = Mid(strColaTot, 1, Len(strColaTot) - 1)
            Case 7: strValue = Mid(strAllowTot, 1, Len(strAllowTot) - 1)
        End Select
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "#,##0.00"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = strValue
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
    Next j
    
    UpdateProgress_Caption "Generating Excel Report", picProgressBar, i / rs.RecordCount
    rs.MoveNext
Wend
rs.Close

SAVING:
On Error GoTo err_saving:
If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

xlsApp.Visible = True

picProgressBar.BackColor = &HFFFFFF
picProgress.Visible = False
picPrint.Enabled = True

Exit Sub
err_saving:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING:

Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub

Exit Sub
ErrorHandler:
Exit Sub
End Sub

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.Height - Me.Height) / 10
Me.Left = (MainForm.Width - Me.Width) / 5

With lstReportType
    .Clear
    .AddItem "PAYSLIP"
    .AddItem "SIGNATURE LEDGER"
    .AddItem "COMPENSATION SUMMARY (TOP SHEET)"
    .AddItem "COMPENSATION SUMMARY"
    .AddItem "DEDUCTION SUMMARY (TOP SHEET)"
    .AddItem "DEDUCTION SUMMARY"
    .AddItem "SSS LOANS (TOP SHEET)"
    .AddItem "SSS LOANS"
    .AddItem "PAG-IBIG LOANS (TOP SHEET)"
    .AddItem "PAG-IBIG LOANS"
    .AddItem "SSS COLLECTIONS (TOP SHEET)"
    .AddItem "SSS COLLECTIONS"
    .AddItem "PHIC COLLECTIONS (TOP SHEET)"
    .AddItem "PHIC COLLECTIONS"
    .AddItem "PAG-IBIG COLLECTIONS (TOP SHEET)"
    .AddItem "PAG-IBIG COLLECTIONS"
    .AddItem "TAX WITHHELD (Alpha List)" '(TOP SHEET)"
    .AddItem "TAX WITHHELD"
    .AddItem "13th MONTH (TOP SHEET)"
    .AddItem "13th MONTH"
    .AddItem "COLA SUMMARY"
    '.AddItem "ALLOWANCE SUMMARY"
    .AddItem "FOR ATM"
    .ListIndex = 0
End With

With cmbDivision
    .Clear
    .AddItem "CLUB HOUSE"
    .AddItem "MAINTENANCE"
    .ListIndex = 0
End With

tmp = SetWindowLong(txtSearchPrint.hwnd, GWL_STYLE, GetWindowLong(txtSearchPrint.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picTaxWithHeldAlpha.Visible = True Then Cancel = -1
If picProgress.Visible = True Then Cancel = -1
End Sub

Private Sub lstReportType_Click()
If lstReportType.ListIndex = -1 Then Exit Sub
LOAD_GROUP_BY lstReportType.ListIndex
cmbDivision_Click
End Sub

Private Sub lstResultPrint_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKPrint_Click
End Sub

Private Sub txtSearchPrint_Change()
If Trim(txtSearchPrint.Text) = "" Then lstResultPrint.Clear:  Exit Sub
lstResultPrint.Clear
's = "SELECT tbl_Personnel_Compensation.EmpPK, " & _
    " tbl_PersonnelProfile.IDNumber, " & _
    " tbl_PersonnelProfile.LName + ',  ' + tbl_PersonnelProfile.FName + '  ' + tbl_PersonnelProfile.MName AS EmpName " & _
    " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
    " tbl_PersonnelProfile ON tbl_Personnel_Compensation.EmpPK = tbl_PersonnelProfile.PK " & _
    " WHERE (tbl_Personnel_Compensation.Division = " & cmbDivision.ListIndex + 1 & ") " & _
    " AND (tbl_Personnel_Compensation.Period = " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ") " & _
    " AND (tbl_PersonnelProfile.LName LIKE '" & FORMATSQL(Trim(txtSearchPrint.Text)) & "%') " & _
    " ORDER BY tbl_PersonnelProfile.LName + ',  ' + tbl_PersonnelProfile.FName + '  ' + tbl_PersonnelProfile.MName"
s = ""
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultPrint.AddItem rs!IDNumber & " - " & rs!EmpName
    lstResultPrint.ItemData(lstResultPrint.NewIndex) = rs!EmpPK
    rs.MoveNext
Wend
rs.Close
If lstResultPrint.ListCount Then lstResultPrint.ListIndex = 0
End Sub

Private Sub txtTaxAlphaYear_GotFocus()
HTEXT txtTaxAlphaYear
End Sub

Private Sub txtTaxAlphaYear_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKTaxAlpha_Click
End Sub

Private Sub txtTaxAlphaYear_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

