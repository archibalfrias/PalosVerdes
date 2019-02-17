VERSION 5.00
Begin VB.Form MainFormPopupF 
   Caption         =   "Form1"
   ClientHeight    =   2385
   ClientLeft      =   165
   ClientTop       =   2310
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu ProfileSearch 
      Caption         =   "ProfileSearch"
      Begin VB.Menu ProfileSearchLName 
         Caption         =   "Last Name"
      End
      Begin VB.Menu ProfileSearchFName 
         Caption         =   "First Name"
      End
      Begin VB.Menu ProfileSearchMName 
         Caption         =   "Middle Name"
      End
   End
   Begin VB.Menu ProfilePrint 
      Caption         =   "ProfilePrint"
      Begin VB.Menu ProfilePrintProfile 
         Caption         =   "Profile"
      End
      Begin VB.Menu ProfilePrintHistory 
         Caption         =   "History"
      End
      Begin VB.Menu ProfilePrintBar1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu ProfilePrintActive 
         Caption         =   "Active"
         Visible         =   0   'False
      End
      Begin VB.Menu ProfilePrintInactive 
         Caption         =   "Inactive"
         Visible         =   0   'False
      End
      Begin VB.Menu ProfilePrintHeadCount 
         Caption         =   "Head Count"
         Visible         =   0   'False
      End
      Begin VB.Menu ProfilePrintBar2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu ProfilePrintAlphalistActive 
         Caption         =   "Alphalist"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuIDSearch 
      Caption         =   "mnuIDSearch"
      Begin VB.Menu mnuIDSearchIDNumber 
         Caption         =   "ID Number"
      End
      Begin VB.Menu mnuIDSearchEmployee 
         Caption         =   "Employee"
      End
   End
   Begin VB.Menu mnuMemberDetails 
      Caption         =   "mnuMemberDetails"
      Begin VB.Menu mnuMemberDetailsAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuMemberDetailsEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuMemberDetailsDelete 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu mnuMemberFind 
      Caption         =   "mnuMemberFind"
      Begin VB.Menu mnuMemberFindLName 
         Caption         =   "Last Name"
      End
      Begin VB.Menu mnuMemberFindFName 
         Caption         =   "First Name"
      End
      Begin VB.Menu mnuMemberFindMName 
         Caption         =   "Middle Name"
      End
   End
   Begin VB.Menu mnuMemberIDFind 
      Caption         =   "mnuMemberIDFind"
      Begin VB.Menu mnuMemberIDFindLName 
         Caption         =   "Last Name"
      End
      Begin VB.Menu mnuMemberIDFindIDNumber 
         Caption         =   "ID Number"
      End
   End
   Begin VB.Menu mnuPrintSystem36 
      Caption         =   "mnuPrintSystem36"
      Begin VB.Menu mnuPrintSystem36Result 
         Caption         =   "Result"
      End
      Begin VB.Menu mnuPrintSystem36Bar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintSystem36Class 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuActionMemo 
      Caption         =   "mnuActionMemo"
      Begin VB.Menu mnuActionMemoRankNFile 
         Caption         =   "Rank in File"
      End
      Begin VB.Menu mnuActionMemoSupervisory 
         Caption         =   "Supervisory"
      End
   End
   Begin VB.Menu mnuCompensationPrint 
      Caption         =   "mnuCompensationPrint"
      Begin VB.Menu mnuCompensationPrintRankNFile 
         Caption         =   "Rank In File"
      End
      Begin VB.Menu mnuCompensationPrintSupervisory 
         Caption         =   "Supervisory"
      End
   End
   Begin VB.Menu mnuPlayerAdd 
      Caption         =   "mnuPlayerAdd"
      Begin VB.Menu mnuPlayerAddIndividual 
         Caption         =   "Individual"
      End
      Begin VB.Menu mnuPlayerAddFromExcel 
         Caption         =   "From Excel"
      End
   End
   Begin VB.Menu mnuAllowancePrint 
      Caption         =   "mnuAllowancePrint"
      Begin VB.Menu mnuAllowancePrintPreview 
         Caption         =   "Preview"
      End
      Begin VB.Menu mnuAllowancePrinttoBank 
         Caption         =   "To Bank"
      End
   End
   Begin VB.Menu mnuMemberActionAdd 
      Caption         =   "mnuMemberActionAdd"
      Begin VB.Menu mnuMemberActionAddAssignee 
         Caption         =   "ASSIGNEE"
      End
      Begin VB.Menu mnuMemberActionAddShareHolder 
         Caption         =   "SHARE HOLDER"
      End
      Begin VB.Menu mnuMemberActionAddBoughtShare 
         Caption         =   "BOUGHT SHARE"
      End
   End
   Begin VB.Menu mnuItemFind 
      Caption         =   "mnuItemFind"
      Begin VB.Menu mnuItemFindItemCode 
         Caption         =   "Item Code"
      End
      Begin VB.Menu mnuItemFindDescription 
         Caption         =   "Description"
      End
   End
   Begin VB.Menu mnuItemReport 
      Caption         =   "mnuItemReport"
      Begin VB.Menu mnuItemReportTransaction 
         Caption         =   "Transactions"
      End
      Begin VB.Menu mnuItemReportSummary 
         Caption         =   "Summary"
      End
      Begin VB.Menu mnuItemReportBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItemReportBySection 
         Caption         =   "by Section"
         Begin VB.Menu mnuItemReportSections 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuItemReportBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItemReportExportToExcel 
         Caption         =   "Export Items to Excel"
      End
   End
   Begin VB.Menu mnuRRFind 
      Caption         =   "mnuRRFind"
      Begin VB.Menu mnuRRFindPONumber 
         Caption         =   "PO Number"
      End
      Begin VB.Menu mnuRRFindRRNumber 
         Caption         =   "RR Number"
      End
   End
   Begin VB.Menu mnuRRPosting 
      Caption         =   "mnuRRPosting"
      Begin VB.Menu mnuRRPostingReceived 
         Caption         =   "Received"
      End
      Begin VB.Menu mnuRRPostingInvoice 
         Caption         =   "G L"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuSupplierReport 
      Caption         =   "mnuSupplierReport"
      Begin VB.Menu mnuSupplierReportItems 
         Caption         =   "List of Items"
      End
      Begin VB.Menu mnuSupplierReportBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSupplierReportHistory 
         Caption         =   "History"
      End
      Begin VB.Menu mnuSupplierReportSL 
         Caption         =   "Subsidiary Ledger"
      End
      Begin VB.Menu mnuSupplierReportBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSupplierReportML 
         Caption         =   "MasterList"
      End
   End
   Begin VB.Menu mnuCVAdd 
      Caption         =   "mnuCVAdd"
      Begin VB.Menu mnuCVAddSupplier 
         Caption         =   "Supplier"
      End
      Begin VB.Menu mnuCVAddMember 
         Caption         =   "Member"
      End
      Begin VB.Menu mnuCVAddEmployee 
         Caption         =   "Employee"
      End
   End
   Begin VB.Menu mnuCVFind 
      Caption         =   "mnuCVFind"
      Begin VB.Menu mnuCVFindCVNumber 
         Caption         =   "CV Number"
      End
      Begin VB.Menu mnuCVFindCheckNumber 
         Caption         =   "Check Number"
      End
      Begin VB.Menu mnuCVFindPayee 
         Caption         =   "Payee"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuCVPrint 
      Caption         =   "mnuCVPrint"
      Begin VB.Menu mnuCVPrintVoucher 
         Caption         =   "Voucher"
      End
      Begin VB.Menu mnuCVPrintCheck 
         Caption         =   "Check"
      End
   End
   Begin VB.Menu mnuChartOfAccounts 
      Caption         =   "mnuChartOfAccounts"
      Begin VB.Menu mnuChartOfAccountsSubsidiary 
         Caption         =   "Subsidiary"
      End
      Begin VB.Menu mnuChartOfAccountsModify 
         Caption         =   "Modify"
      End
   End
   Begin VB.Menu mnuScoringLocation 
      Caption         =   "mnuScoringLocation"
      Begin VB.Menu mnuScoringLocationName 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuScoringLocationAdd 
      Caption         =   "mnuScoringLocationAdd"
      Begin VB.Menu mnuScoringLocationNameAdd 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuRegistrationAdd 
      Caption         =   "mnuRegistrationAdd"
      Begin VB.Menu mnuRegistrationAddBagTagNo 
         Caption         =   "Bag Tag Number"
      End
      Begin VB.Menu mnuRegistrationAddPlayerName 
         Caption         =   "Player Name"
      End
   End
   Begin VB.Menu mnuTournamentInfoPrint 
      Caption         =   "mnuTournamentInfoPrint"
      Begin VB.Menu mnuTournamentInfoPrintScoreCard 
         Caption         =   "Score Card"
      End
   End
   Begin VB.Menu mnuPayrollHourPosting 
      Caption         =   "mnuPayrollHourPosting"
      Begin VB.Menu mnuPayrollHourPostingSingleTrans 
         Caption         =   "this Transaction"
      End
      Begin VB.Menu mnuPayrollHourPostingBatch 
         Caption         =   "Batch"
      End
   End
   Begin VB.Menu mnuPayrollPrint 
      Caption         =   "mnuPayrollPrint"
      Begin VB.Menu mnuPayrollPrintRankNFile 
         Caption         =   "Rank N File"
      End
      Begin VB.Menu mnuPayrollPrintSupervisory 
         Caption         =   "Supervisory"
      End
   End
   Begin VB.Menu mnuLoanRep 
      Caption         =   "mnuLoanRep"
      Begin VB.Menu mnuLoanRepSubsidiary 
         Caption         =   "Subsidiary"
      End
      Begin VB.Menu mnuLoanRepEmpActiveLoan 
         Caption         =   "Employee Active Loan"
      End
   End
   Begin VB.Menu mnuPayrollPrint1 
      Caption         =   "mnuPayrollPrint1"
      Begin VB.Menu mnuPayrollPrint1RnFSup 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuPayrollDeductionReport 
      Caption         =   "mnuPayrollDeductionReport"
      Begin VB.Menu mnuPayrollDeductionReportEmployee 
         Caption         =   "per Employee"
      End
      Begin VB.Menu mnuPayrollDeductionReportSummary 
         Caption         =   "Summary"
      End
   End
End
Attribute VB_Name = "MainFormPopupF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Trans_Rank_Supervisory As Long

Dim Filename As String
Dim WorkbookName As String
Dim iWorkSheet As Integer
Dim RowCnt, ColCnt, strRange, i, l, k, x, strValue, iReset, strAmount, _
iPK, Arr, Arr1, iLocationKey, iFilterIndex, sLine, strPath, sLoanName_Status, iRec, iLine, dDebit, dCredit, dRunBal, sRemarks
        
Private Sub PAYROLL_TEMP_RF_SV(Division, Period, PostLevel)

frmPersonnelCompensationReport.picProgressBar.BackColor = &HFFFFFF
frmPersonnelCompensationReport.picPrint.Enabled = False
frmPersonnelCompensationReport.picProgress.ZOrder 0
frmPersonnelCompensationReport.picProgress.Visible = True
i = 0

ConnOmega.Execute "DELETE FROM tbl_PersonnelPayroll_Tmp WHERE (LogInName = '" & gbl_UserName & "')"

's = "SELECT qry_Payroll_Transaction.*" & _
    " From qry_Payroll_Transaction " & _
    " WHERE (Division = " & Division & ") " & _
    " AND (Period = " & Period & ")"
    
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
    
    UpdateProgress frmPersonnelCompensationReport.picProgressBar, i / ra.RecordCount
    ra.MoveNext
Wend
ra.Close

s = "SELECT tbl_PersonnelPayroll_Tmp.* " & _
    " FROM tbl_PersonnelPayroll_Tmp"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
ra.Requery
ra.Close

frmPersonnelCompensationReport.picProgressBar.BackColor = &HFFFFFF
frmPersonnelCompensationReport.picProgress.Visible = False
frmPersonnelCompensationReport.picPrint.Enabled = True


End Sub


Private Sub mnuActionMemoRankNFile_Click()
    
    With frmPersonnelAction
        .cmbPost.Clear
        s = "SELECT PK, PositionName " & _
            " From tbl_Personnel_Position " & _
            " Where (PositionLevel = 1) " & _
            " ORDER BY PositionName"
        If ra.State = adStateOpen Then ra.Close
        ra.Open s, ConnOmega
        While Not ra.EOF
            .cmbPost.AddItem ra!PositionName
            .cmbPost.ItemData(.cmbPost.NewIndex) = ra!PK
            ra.MoveNext
        Wend
        ra.Close
        .HIDE_SALARY_RATE 0
        .picToolbar.Enabled = False
        .picMain.Enabled = False
        .picAdd.ZOrder 0
        .txtSearch.Text = ""
        .picAdd.Visible = True
        .txtSearch.SetFocus
    End With
End Sub

Private Sub mnuActionMemoSupervisory_Click()
If AccessRights("Personnel Action Memo", "Supervisory") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
With frmPersonnelAction
    .cmbPost.Clear
    s = "SELECT PK, PositionName " & _
        " From tbl_Personnel_Position " & _
        " Where (PositionLevel = 2) " & _
        " ORDER BY PositionName"
    If ra.State = adStateOpen Then ra.Close
    ra.Open s, ConnOmega
    While Not ra.EOF
        .cmbPost.AddItem ra!PositionName
        .cmbPost.ItemData(.cmbPost.NewIndex) = ra!PK
        ra.MoveNext
    Wend
    ra.Close
    .picToolbar.Enabled = False
    .picMain.Enabled = False
    .HIDE_SALARY_RATE 2
    .picAdd.ZOrder 0
    .txtSearch.Text = ""
    .picAdd.Visible = True
    .txtSearch.SetFocus
End With
End Sub

Private Sub mnuAllowancePrintPreview_Click()
With frmPersonnelAllowanceBrowse
    .iType = 2       'Preview
    .b8TitleBar4.Caption = "Print"
    DoEvents
    .picMain.Enabled = False
    .picToolbar.Enabled = False
    .picGenerate.ZOrder 0
    .cmbDivision.ListIndex = -1
    .txtFrom.Text = ""
    .txtTo.Text = ""
    .picGenerate.Visible = True
    .cmbDivision.SetFocus
End With
End Sub

Private Sub mnuAllowancePrinttoBank_Click()
With frmPersonnelAllowanceBrowse
    .iType = 3       'to Bank
    .b8TitleBar4.Caption = "Print"
    DoEvents
    .picMain.Enabled = False
    .picToolbar.Enabled = False
    .picGenerate.ZOrder 0
    .cmbDivision.ListIndex = -1
    .txtFrom.Text = ""
    .txtTo.Text = ""
    .picGenerate.Visible = True
    .cmbDivision.SetFocus
End With
End Sub

Private Sub mnuChartOfAccountsModify_Click()
If RETURNTEXTVALUE(frmAcctgChartOfAccounts.txtCurrPK) = 0 Then Exit Sub
SaveSetting App.EXEName, "ChartOfAccount", "ChartOfAccount", frmAcctgChartOfAccounts.txtCurrCode.Text
If IsLoaded(frmAcctgChartOfAccountsCard) Then frmAcctgChartOfAccountsCard.ZOrder 0 Else frmAcctgChartOfAccountsCard.Show
frmAcctgChartOfAccountsCard.BROWSER GetSetting(App.EXEName, "ChartOfAccount", "ChartOfAccount", ""), "is_LOAD"
End Sub

Private Sub mnuChartOfAccountsSubsidiary_Click()
If RETURNTEXTVALUE(frmAcctgChartOfAccounts.txtCurrPK) = 0 Then Exit Sub
frmAcctgChartOfAccountsSL.sFormCaption = Trim(frmAcctgChartOfAccounts.FGrid.TextMatrix(frmAcctgChartOfAccounts.FGrid.ROW, 1)) & " - " & Trim(frmAcctgChartOfAccounts.FGrid.TextMatrix(frmAcctgChartOfAccounts.FGrid.ROW, 2))
frmAcctgChartOfAccountsSL.sAccCode = Trim(frmAcctgChartOfAccounts.FGrid.TextMatrix(frmAcctgChartOfAccounts.FGrid.ROW, 1))
frmAcctgChartOfAccountsSL.txtDateFrom.Text = ""
frmAcctgChartOfAccountsSL.txtDateTo.Text = ""
frmAcctgChartOfAccountsSL.TimerLoadSL.Enabled = True
If IsLoaded(frmAcctgChartOfAccountsSL) Then frmAcctgChartOfAccountsSL.ZOrder 0 Else frmAcctgChartOfAccountsSL.Show
End Sub

Private Sub mnuCompensationPrintRankNFile_Click()
Select Case frmPersonnelCompensationReport.lstReportType.ListIndex
    Case 1  'SIGNATURE LEDGER
        
        PAYROLL_TEMP_RF_SV frmPersonnelCompensationReport.cmbDivision.ListIndex + 1, _
            frmPersonnelCompensationReport.cmbPeriodPrint.ItemData(frmPersonnelCompensationReport.cmbPeriodPrint.ListIndex), _
            1
        
        frmCrystalReportViewer.PRINT_SIGNATURE_LEDGER gbl_CompanyName, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
    
    Case 2  'COMPENSATION (TOP)
        
        PAYROLL_TEMP_RF_SV frmPersonnelCompensationReport.cmbDivision.ListIndex + 1, _
            frmPersonnelCompensationReport.cmbPeriodPrint.ItemData(frmPersonnelCompensationReport.cmbPeriodPrint.ListIndex), _
            1
            
        frmCrystalReportViewer.PRINT_COMPENSATION_SUMMARY_TOP gbl_CompanyName, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
        
    Case 3  'COMPENSATION
        
        PAYROLL_TEMP_RF_SV frmPersonnelCompensationReport.cmbDivision.ListIndex + 1, _
            frmPersonnelCompensationReport.cmbPeriodPrint.ItemData(frmPersonnelCompensationReport.cmbPeriodPrint.ListIndex), _
            1
            
        frmCrystalReportViewer.PRINT_COMPENSATION_SUMMARY gbl_CompanyName, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
    
    Case 4  'DEDUCTION (TOP)
        
        PAYROLL_TEMP_RF_SV frmPersonnelCompensationReport.cmbDivision.ListIndex + 1, _
            frmPersonnelCompensationReport.cmbPeriodPrint.ItemData(frmPersonnelCompensationReport.cmbPeriodPrint.ListIndex), _
            1
            
        frmCrystalReportViewer.PRINT_DEDUCTION_SUMMARY_TOP gbl_CompanyName, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
    
    Case 5  'DEDUCTION
        
        PAYROLL_TEMP_RF_SV frmPersonnelCompensationReport.cmbDivision.ListIndex + 1, _
            frmPersonnelCompensationReport.cmbPeriodPrint.ItemData(frmPersonnelCompensationReport.cmbPeriodPrint.ListIndex), _
            1
            
        frmCrystalReportViewer.PRINT_DEDUCTION_SUMMARY gbl_CompanyName, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
    
    Case 21 'For BPI
        
              
        i = 0
        
        s = "SELECT tbl_Personnel_IDNumber.AccountNumber, " & _
            " tbl_Personnel_Information.FirstName, " & _
            " tbl_Personnel_Information.MiddleName, " & _
            " tbl_Personnel_Information.LastName, " & _
            " tbl_Personnel_Compensation.TotalEarning, " & _
            " tbl_Personnel_Compensation.TotalDeduction, " & _
            " tbl_Personnel_Compensation.TotalCola " & _
            " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
            " tbl_Personnel_Position ON tbl_Personnel_Compensation.Positions = tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_IDNumber ON tbl_Personnel_Compensation.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
            " WHERE (tbl_Personnel_IDNumber.AccountNumber <> '') " & _
            " AND (tbl_Personnel_Compensation.Division = " & frmPersonnelCompensationReport.cmbDivision.ListIndex + 1 & ") " & _
            " AND (tbl_Personnel_Compensation.Period = " & frmPersonnelCompensationReport.cmbPeriodPrint.ItemData(frmPersonnelCompensationReport.cmbPeriodPrint.ListIndex) & ") " & _
            " AND (tbl_Personnel_Position.PositionLevel = 1) " & _
            " ORDER BY tbl_Personnel_Information.LastName, tbl_Personnel_Information.FirstName, tbl_Personnel_Information.MiddleName"
        If ra.State = adStateOpen Then ra.Close
        ra.Open s, ConnOmega
        If ra.RecordCount > 0 Then
        
            MainForm.CommonDialog1.CancelError = True
            On Error GoTo ErrorHandler
            MainForm.CommonDialog1.DialogTitle = "Save"
            MainForm.CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
            MainForm.CommonDialog1.ShowSave
            Filename = Trim(MainForm.CommonDialog1.Filename)
            
            WorkbookName = CStr(Filename)
            
            frmPersonnelCompensationReport.picProgressBar.BackColor = &HFFFFFF
            frmPersonnelCompensationReport.picPrint.Enabled = False
            frmPersonnelCompensationReport.picProgress.ZOrder 0
            frmPersonnelCompensationReport.picProgress.Visible = True
            
            iWorkSheet = 1
            Set xlsApp = CreateObject("Excel.Application")
            xlsApp.Visible = False
            xlsApp.Workbooks.Add
            xlsApp.DisplayAlerts = False
            xlsApp.Workbooks(1).Sheets(2).Delete
            xlsApp.Workbooks(1).Sheets(2).Delete
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
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "For B P I ATM (Rank in File)"
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
            For k = 1 To 8
                Select Case k
                    Case 1: strValue = "Account Number"
                    Case 2: strValue = "First Name"
                    Case 3: strValue = "Middle Name"
                    Case 4: strValue = "Last Name"
                    Case 5: strValue = "Total Earnings"
                    Case 6: strValue = "Total Deduction"
                    Case 7: strValue = "C O L A"
                    Case 8: strValue = "Amount" '"Allowance"
                    'Case 9: strValue = "Amount"
                End Select
                ColCnt = k
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = strValue
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
                If k >= 5 And k <= 8 Then
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
            While Not ra.EOF
                i = i + 1
                strAmount = "="
                RowCnt = RowCnt + 1
                For k = 1 To 7
                    strValue = ra.Fields(k - 1).Value
                    ColCnt = k
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = strValue
                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
                    Select Case k
                        Case 5: strAmount = strAmount & strRange
                        Case 6: strAmount = strAmount & "-" & strRange
                        Case 7: strAmount = strAmount & "+" & strRange
                        'Case 8: strAmount = strAmount & "+" & strRange
                    End Select
                    
                    If k >= 5 And k <= 7 Then
                        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
                        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "#,##0.00"
                    ElseIf k = 1 Then
                        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "0000-0000-00"
                        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3
                    Else
                        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 2
                    End If
                Next k
                strValue = strAmount 'Mid(strAmount, 1, Len(strAmount) - 1)
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = strValue
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "#,##0.00"
                
                UpdateProgress frmPersonnelCompensationReport.picProgressBar, i / ra.RecordCount
                
                ra.MoveNext
            Wend
        End If
        ra.Close
        
SAVING:
        On Error GoTo err_saving:
        If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
        xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName
        
        xlsApp.Visible = True
        
        frmPersonnelCompensationReport.picProgressBar.BackColor = &HFFFFFF
        frmPersonnelCompensationReport.picProgress.Visible = False
        frmPersonnelCompensationReport.picPrint.Enabled = True
        
End Select
Exit Sub
ErrorHandler:
Exit Sub

Exit Sub
err_saving:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING:
End Sub

Private Sub mnuCompensationPrintSupervisory_Click()

If AccessRights("Personnel Compensation", "Supervisory") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If

Select Case frmPersonnelCompensationReport.lstReportType.ListIndex
    Case 1  'SIGNATURE LEDGER
        
        PAYROLL_TEMP_RF_SV frmPersonnelCompensationReport.cmbDivision.ListIndex + 1, _
            frmPersonnelCompensationReport.cmbPeriodPrint.ItemData(frmPersonnelCompensationReport.cmbPeriodPrint.ListIndex), _
            2
        
        frmCrystalReportViewer.PRINT_SIGNATURE_LEDGER gbl_CompanyName, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
    
    Case 2  'COMPENSATION (TOP)
        
        PAYROLL_TEMP_RF_SV frmPersonnelCompensationReport.cmbDivision.ListIndex + 1, _
            frmPersonnelCompensationReport.cmbPeriodPrint.ItemData(frmPersonnelCompensationReport.cmbPeriodPrint.ListIndex), _
            2
            
        frmCrystalReportViewer.PRINT_COMPENSATION_SUMMARY_TOP gbl_CompanyName, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
        
    Case 3  'COMPENSATION
        
        PAYROLL_TEMP_RF_SV frmPersonnelCompensationReport.cmbDivision.ListIndex + 1, _
            frmPersonnelCompensationReport.cmbPeriodPrint.ItemData(frmPersonnelCompensationReport.cmbPeriodPrint.ListIndex), _
            2
            
        frmCrystalReportViewer.PRINT_COMPENSATION_SUMMARY gbl_CompanyName, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
    
    Case 4  'DEDUCTION (TOP)
        
        PAYROLL_TEMP_RF_SV frmPersonnelCompensationReport.cmbDivision.ListIndex + 1, _
            frmPersonnelCompensationReport.cmbPeriodPrint.ItemData(frmPersonnelCompensationReport.cmbPeriodPrint.ListIndex), _
            2
            
        frmCrystalReportViewer.PRINT_DEDUCTION_SUMMARY_TOP gbl_CompanyName, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
    
    Case 5  'DEDUCTION
        
        PAYROLL_TEMP_RF_SV frmPersonnelCompensationReport.cmbDivision.ListIndex + 1, _
            frmPersonnelCompensationReport.cmbPeriodPrint.ItemData(frmPersonnelCompensationReport.cmbPeriodPrint.ListIndex), _
            2
            
        frmCrystalReportViewer.PRINT_DEDUCTION_SUMMARY gbl_CompanyName, gbl_UserName
        If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
    
    Case 21 'For BPI
        

        
        i = 0
        
        s = "SELECT tbl_Personnel_IDNumber.AccountNumber, " & _
            " tbl_Personnel_Information.FirstName, " & _
            " tbl_Personnel_Information.MiddleName, " & _
            " tbl_Personnel_Information.LastName, " & _
            " tbl_Personnel_Compensation.TotalEarning, " & _
            " tbl_Personnel_Compensation.TotalDeduction, " & _
            " tbl_Personnel_Compensation.TotalCola " & _
            " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
            " tbl_Personnel_Position ON tbl_Personnel_Compensation.Positions = tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_IDNumber ON tbl_Personnel_Compensation.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
            " WHERE (tbl_Personnel_IDNumber.AccountNumber <> '') " & _
            " AND (tbl_Personnel_Compensation.Division = " & frmPersonnelCompensationReport.cmbDivision.ListIndex + 1 & ") " & _
            " AND (tbl_Personnel_Compensation.Period = " & frmPersonnelCompensationReport.cmbPeriodPrint.ItemData(frmPersonnelCompensationReport.cmbPeriodPrint.ListIndex) & ") " & _
            " AND (tbl_Personnel_Position.PositionLevel = 2) " & _
            " ORDER BY tbl_Personnel_Information.LastName, tbl_Personnel_Information.FirstName, tbl_Personnel_Information.MiddleName"
        If ra.State = adStateOpen Then ra.Close
        ra.Open s, ConnOmega
        If ra.RecordCount > 0 Then
        
            MainForm.CommonDialog1.CancelError = True
            On Error GoTo ErrorHandler
            MainForm.CommonDialog1.DialogTitle = "Save"
            MainForm.CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
            MainForm.CommonDialog1.ShowSave
            Filename = Trim(MainForm.CommonDialog1.Filename)
            
            WorkbookName = CStr(Filename)
            
            frmPersonnelCompensationReport.picProgressBar.BackColor = &HFFFFFF
            frmPersonnelCompensationReport.picPrint.Enabled = False
            frmPersonnelCompensationReport.picProgress.ZOrder 0
            frmPersonnelCompensationReport.picProgress.Visible = True
            
            iWorkSheet = 1
            Set xlsApp = CreateObject("Excel.Application")
            xlsApp.Visible = False
            xlsApp.Workbooks.Add
            xlsApp.DisplayAlerts = False
            xlsApp.Workbooks(1).Sheets(2).Delete
            xlsApp.Workbooks(1).Sheets(2).Delete
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
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "For B P I ATM (Supervisory)"
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
            For k = 1 To 8
                Select Case k
                    Case 1: strValue = "Account Number"
                    Case 2: strValue = "First Name"
                    Case 3: strValue = "Middle Name"
                    Case 4: strValue = "Last Name"
                    Case 5: strValue = "Total Earnings"
                    Case 6: strValue = "Total Deduction"
                    Case 7: strValue = "C O L A"
                    Case 8: strValue = "Amount" '"Allowance"
                    'Case 9: strValue = "Amount"
                End Select
                ColCnt = k
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = strValue
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
                If k >= 5 And k <= 8 Then
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
            While Not ra.EOF
                i = i + 1
                strAmount = "="
                RowCnt = RowCnt + 1
                For k = 1 To 7
                    strValue = ra.Fields(k - 1).Value
                    ColCnt = k
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = strValue
                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
                    Select Case k
                        Case 5: strAmount = strAmount & strRange
                        Case 6: strAmount = strAmount & "-" & strRange
                        Case 7: strAmount = strAmount & "+" & strRange
                        'Case 8: strAmount = strAmount & "+" & strRange
                    End Select
                    
                    If k >= 5 And k <= 7 Then
                        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
                        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "#,##0.00"
                    ElseIf k = 1 Then
                        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "0000-0000-00"
                        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3
                    Else
                        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 2
                    End If
                Next k
                strValue = strAmount 'Mid(strAmount, 1, Len(strAmount) - 1)
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = strValue
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "#,##0.00"
                
                UpdateProgress frmPersonnelCompensationReport.picProgressBar, i / ra.RecordCount
                
                ra.MoveNext
            Wend
        End If
        ra.Close
        
SAVING:
        On Error GoTo err_saving:
        If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
        xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName
        
        xlsApp.Visible = True
        
        frmPersonnelCompensationReport.picProgressBar.BackColor = &HFFFFFF
        frmPersonnelCompensationReport.picProgress.Visible = False
        frmPersonnelCompensationReport.picPrint.Enabled = True
End Select

Exit Sub
ErrorHandler:
Exit Sub

Exit Sub
err_saving:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING:
End Sub

Private Sub mnuCVAddEmployee_Click()
With frmAcctgCheckVoucher
    .iPayeeType = 3
    .picAdd.ZOrder 0
    .txtSearchAdd.Text = ""
    .picAdd.Visible = True
    .txtSearchAdd.SetFocus
End With
End Sub

Private Sub mnuCVAddMember_Click()
With frmAcctgCheckVoucher
    .iPayeeType = 2
    .picAdd.ZOrder 0
    .txtSearchAdd.Text = ""
    .picAdd.Visible = True
    .txtSearchAdd.SetFocus
End With
End Sub

Private Sub mnuCVAddSupplier_Click()
With frmAcctgCheckVoucher
    .iPayeeType = 1
    .picAdd.ZOrder 0
    .txtSearchAdd.Text = ""
    .picAdd.Visible = True
    .txtSearchAdd.SetFocus
End With
End Sub

Private Sub mnuCVFindCheckNumber_Click()
With frmAcctgCheckVoucher
    .picMain.Enabled = False
    .picToolbar.Enabled = False
    .picSearch.ZOrder 0
    .txtSearch.Text = ""
    .picSearch.Visible = True
    .SearchType = 2
    .txtSearch.SetFocus
End With
End Sub

Private Sub mnuCVFindCVNumber_Click()
With frmAcctgCheckVoucher
    .picMain.Enabled = False
    .picToolbar.Enabled = False
    .picSearch.ZOrder 0
    .txtSearch.Text = ""
    .picSearch.Visible = True
    .SearchType = 1
    .txtSearch.SetFocus
End With
End Sub

Private Sub mnuCVPrintCheck_Click()
With frmAcctgCheckVoucher
    
    If .imgPosted.Visible = True Then MsgBox "Already Posted!                        ", vbCritical, "Error...": Exit Sub
    
    If Trim(.txtCheckNumber.Text) = "" Then MsgBox "Please Supply Check Number!                   ", vbCritical, "Error...": .txtCheckNumber.SetFocus: Exit Sub
    
    If IsDate(.txtCheckDate.Text) = False Then MsgBox "Please Supply Check Date!                      ", vbCritical, "Error...": .txtCheckDate.SetFocus: Exit Sub
    
    CREATE_CV_CHECK "tbl_Acctg_CheckVoucher_Check"
    
    ConnOmega.Execute "DELETE FROM tbl_Acctg_CheckVoucher_Check WHERE (LogInName = '" & gbl_UserName & "')"
    
    ConnOmega.Execute "INSERT INTO tbl_Acctg_CheckVoucher_Check " & _
                      " (LogInName, CheckDate, CheckAmt, PayTo, Pesos) " & _
                      " VALUES ('" & gbl_UserName & "', '" & Format(FormatDateTime(.txtCheckDate.Text, vbShortDate), "mmmm dd, yyyy") & "', " & _
                      " '" & "***" & .lblTotal.Caption & "***" & "', '" & FORMATSQL(Trim(.txtPayeeName.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(.txtAmtWords.Text)) & "')"
    

    frmPrinter.PRINT_TRANSACTION = 5
    frmPrinter.Show 1
 
End With
End Sub

Private Sub mnuCVPrintVoucher_Click()
With frmAcctgCheckVoucher
    
    If .imgPosted.Visible = True Then MsgBox "Already Posted!                        ", vbCritical, "Error...": Exit Sub

    CREATE_CV_TABLES "tbl_Acctg_CheckVoucher_Report"
    
    ConnOmega.Execute "DELETE FROM tbl_Acctg_CheckVoucher_Report WHERE (LogInName = '" & gbl_UserName & "')"
    
    ConnOmega.Execute "INSERT INTO tbl_Acctg_CheckVoucher_Report " & _
                      " (LogInName, CVNumber, CVDate, PayTo, Pesos, ORNumber, TotAmt, ChkNumber, Approved, Received, Prepared, Checked, Entered) " & _
                      " VALUES ('" & gbl_UserName & "', '" & Trim(.txtCVNumber.Text) & "', '" & Format(FormatDateTime(.txtCVDate.Text, vbShortDate), "mmm dd, yyyy") & "', " & _
                      " '" & FORMATSQL(Trim(.txtPayeeName.Text)) & "', '" & FORMATSQL(Trim(.txtAmtWords.Text)) & "', '" & FORMATSQL(Trim(.txtORNumber.Text)) & "', " & _
                      " '" & .lblTotal.Caption & "', '" & Format(Trim(.txtCheckNumber.Text)) & "', '" & FORMATSQL(Trim(.txtApproved.Text)) & "', " & _
                      " '','" & FORMATSQL(Trim(.txtPrepared.Text)) & "', '" & FORMATSQL(Trim(.txtChecked.Text)) & "', '" & FORMATSQL(Trim(.txtEntered.Text)) & "')"
                      
    iPK = 0
    s = "SELECT tbl_Acctg_CheckVoucher_Report.* " & _
        " FROM tbl_Acctg_CheckVoucher_Report " & _
        " WHERE (LogInName = '" & gbl_UserName & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        iPK = rs!PK
    End If
    rs.Close
    
    With .lstExplanation.ListItems
        l = 0
        For i = 1 To .Count
            If Trim(.Item(i).SubItems(1)) <> "" And _
            Trim(.Item(i).SubItems(2)) <> "" Then
                l = l + 1
                ConnOmega.Execute "INSERT INTO tbl_Acctg_CheckVoucher_Report_Explanation " & _
                                  " (MasterKey, Line, Description, Amount) " & _
                                  " VALUES (" & iPK & ", " & l & ", '" & FORMATSQL(Trim(.Item(i).SubItems(1))) & "', " & _
                                  " '" & FORMATSQL(Trim(.Item(i).SubItems(2))) & "')"
            End If
        Next i
    End With
    
    l = 0
    With .lstAccountDistribution.ListItems
        For i = 1 To .Count
            If Trim(.Item(i).SubItems(1)) <> "" Then
                l = l + 1
                ConnOmega.Execute "INSERT INTO tbl_Acctg_CheckVoucher_Report_AD " & _
                                  " (MasterKey, Line, AccountCode, AccountName, Debit, Credit) " & _
                                  " VALUES (" & iPK & ", " & l & ", " & _
                                  " '" & FORMATSQL(Trim(.Item(i).SubItems(1))) & "', " & _
                                  " '" & FORMATSQL(Trim(.Item(i).SubItems(2))) & "', " & _
                                  " '" & FORMATSQL(Trim(.Item(i).SubItems(3))) & "', " & _
                                  " '" & FORMATSQL(Trim(.Item(i).SubItems(4))) & "')"
            End If
        Next i
    End With
    
    s = "SELECT tbl_Acctg_CheckVoucher_Report.* " & _
        " FROM tbl_Acctg_CheckVoucher_Report " & _
        " WHERE (LogInName = '" & gbl_UserName & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    rs.Requery
    rs.Close

    frmPrinter.PRINT_TRANSACTION = 4
    frmPrinter.Show 1
    
    
    ''frmCrystalReportViewer.PRINT_CHECK_VOUCHER gbl_UserName
'    frmCrystalReportViewer.PRINT_CHECK_VOUCHER_PREPRINTED gbl_UserName
'    If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show

End With
End Sub

Private Sub mnuIDSearchEmployee_Click()
With frmPersonnelIDNumber
    .picSearch1.ZOrder 0
    .picMain.Enabled = False
    .picToolbar.Enabled = False
    .txtSearch1.Text = ""
    .picSearch1.Visible = True
    .txtSearch1.SetFocus
End With
End Sub

Private Sub mnuIDSearchIDNumber_Click()
With frmPersonnelIDNumber
    .SearchType = 1
    .picSearch.ZOrder 0
    .picMain.Enabled = False
    .picToolbar.Enabled = False
    .txtSearch.Text = ""
    .picSearch.Visible = True
    .txtSearch.SetFocus
End With
End Sub

Private Sub mnuItemFindDescription_Click()
With frmInvItems
    .picSearch.Top = (.ScaleHeight - .picSearch.Height) / 2
    .picSearch.Left = (.ScaleWidth - .picSearch.Width) / 2
    .picSearch.ZOrder 0
    .picBody.Enabled = False
    .picToolbar.Enabled = False
    .picSearch.Visible = True
    .txtSearch.Text = ""
    .txtSearch.SetFocus
End With
End Sub

Private Sub mnuItemFindItemCode_Click()
With frmInvItems
    .TRANSACTIONTYPE = 3
    .CLEARTEXT
    .TOOLBARFUNC 3
    .Caption = "Items - Find"
    .txtItemCode.Locked = False
    .txtItemCode.SetFocus
End With
End Sub

Private Sub mnuItemReportExportToExcel_Click()
frmInvItems.TimerExporttoExcel.Enabled = True
End Sub

Private Sub mnuItemReportSections_Click(Index As Integer)
Arr = Split(mnuItemReportSections(Index).Caption, " - ", -1, 1)
frmPreview.sSectCode = Arr(0)
frmPreview.sSectName = Arr(1)
frmPreview.Timer_Items_Section.Enabled = True
If IsLoaded(frmPreview) Then frmPreview.ZOrder 0 Else frmPreview.Show
End Sub

Private Sub mnuItemReportTransaction_Click()
frmPreview.iItemKey = frmInvItems.Statusbar1.Panels(1).Text
frmPreview.sItemCode = frmInvItems.txtItemCode.Text
frmPreview.sItemDesc = frmInvItems.txtItemDesc.Text
frmPreview.Timer_ItemsTransaction.Enabled = True
If IsLoaded(frmPreview) Then frmPreview.ZOrder 0 Else frmPreview.Show
End Sub

Private Sub mnuLoanRepEmpActiveLoan_Click()
With frmPersonnelLoans
    .isAdd_isLoan = 2
    .txtSearchAdd.Text = ""
    .picSearchAdd.ZOrder 0
    .picBody.Enabled = False
    .picSearchAdd.Visible = True
    .txtSearchAdd.SetFocus
End With
End Sub

Private Sub mnuLoanRepSubsidiary_Click()
With frmPersonnelLoans
    If .imgPosted.Visible = False Then MsgBox "Not yet posted!                  ", vbInformation, "Info": Exit Sub
    MainForm.picProgressBar.BackColor = &H8000000F
    DoEvents
    Screen.MousePointer = vbHourglass
    sLoanName_Status = .cmbLoanType.Text & _
                       IIf(.locZeroOut = 1, " [Zero out]", "")
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll_Report_LoanLedger WHERE (LogInName = '" & gbl_UserName & "')"
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_LoanLedger (LogInName, CompanyKey, EmployeeName) VALUES ('" & gbl_UserName & "', 1, '" & FORMATSQL(Trim(.txtName.Text)) & "')"
    iPK = 0: iLine = 0: dRunBal = 0
    t = "SELECT PK " & _
        " FROM tbl_Personnel_Payroll_Report_LoanLedger " & _
        " WHERE (LogInName = '" & gbl_UserName & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        iPK = rt!PK
    End If
    rt.Close
    iRec = 0
    s = "SELECT dbo.tbl_Personnel_Loans_SL.* " & _
        " From dbo.tbl_Personnel_Loans_SL " & _
        " WHERE (LoanKey = " & .StatusBar.Panels(1).Text & ") " & _
        " ORDER BY PK"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        While Not rs.EOF
            iRec = iRec + 1
            iLine = iLine + 1
            sRemarks = "[" & Format(rs!TransactionDate, "mm/dd/yyyy") & "]" & IIf(Trim(rs!Remarks) <> "", " - " & rs!Remarks, "")
            dDebit = rs!Debit
            dCredit = rs!Credit
            dRunBal = CDbl(Format(dRunBal, "#0.00")) + (CDbl(dDebit) - CDbl(dCredit))
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_LoanLedger_Det " & _
                              " (MasterKey, Line, LoanName, Remarks, Debit, Credit, RunBal) " & _
                              " VALUES (" & iPK & ", " & iLine & ", '" & FORMATSQL(CStr(sLoanName_Status)) & "', " & _
                              " '" & FORMATSQL(CStr(sRemarks)) & "'," & CDbl(dDebit) & ", " & CDbl(dCredit) & ", " & CDbl(dRunBal) & ")"
            UpdateProgress_No_Percent MainForm.picProgressBar, iRec / rs.RecordCount
            rs.MoveNext
        Wend
    End If
    rs.Close
    Screen.MousePointer = vbDefault
    MainForm.picProgressBar.BackColor = &H8000000F
    DoEvents
    frmCrystalReportViewer.PRINT_LOAN_Ledger gbl_UserName
    If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
End With
End Sub

Private Sub mnuMemberActionAddAssignee_Click()
frmMembershipAction.iActionType = 1
frmMembershipAction.iSearchAdd = 1
'frmMembershipAction.b8TitleBar2.Caption = "Search Assignor"
frmMembershipAction.Label11.Caption = "Select Assignor"
DoEvents
frmMembershipAction.picMain.Enabled = False
frmMembershipAction.picToolbar.Enabled = False
frmMembershipAction.picAdd.ZOrder 0
frmMembershipAction.txtSearchAdd.Text = ""
frmMembershipAction.lstResultAdd.Height = 1230 '1425
frmMembershipAction.txtSearchAssignor.Visible = True
frmMembershipAction.lstResultAssignorAdd.Visible = True
frmMembershipAction.picAdd.Visible = True
frmMembershipAction.txtSearchAdd.SetFocus
End Sub

Private Sub mnuMemberActionAddBoughtShare_Click()
frmMembershipAction.iActionType = 3 '1
frmMembershipAction.iSearchAdd = 3 '1
'frmMembershipAction.b8TitleBar2.Caption = "Search Assignor"
frmMembershipAction.Label11.Caption = "Select Share Holder"
DoEvents
frmMembershipAction.picMain.Enabled = False
frmMembershipAction.picToolbar.Enabled = False
frmMembershipAction.picAdd.ZOrder 0
frmMembershipAction.txtSearchAdd.Text = ""
frmMembershipAction.lstResultAdd.Height = 1230 '1425
frmMembershipAction.txtSearchAssignor.Visible = True
frmMembershipAction.lstResultAssignorAdd.Visible = True
frmMembershipAction.picAdd.Visible = True
frmMembershipAction.txtSearchAdd.SetFocus
End Sub

Private Sub mnuMemberActionAddShareHolder_Click()
frmMembershipAction.iActionType = 2
frmMembershipAction.iSearchAdd = 1
frmMembershipAction.b8TitleBar2.Caption = "Search"
DoEvents
frmMembershipAction.picMain.Enabled = False
frmMembershipAction.picToolbar.Enabled = False
frmMembershipAction.picAdd.ZOrder 0
frmMembershipAction.txtSearchAdd.Text = ""
frmMembershipAction.lstResultAdd.Height = 3375
frmMembershipAction.txtSearchAssignor.Visible = False
frmMembershipAction.lstResultAssignorAdd.Visible = False
frmMembershipAction.picAdd.Visible = True
frmMembershipAction.txtSearchAdd.SetFocus
End Sub

Private Sub mnuMemberDetailsAdd_Click()

With frmMembershipInformation
    Select Case .FocusDetail
        Case 1  'Child
            With .lstChildren.ListItems
                If Trim(.Item(.Count).SubItems(2)) <> "" Then
                    Set x = .Add()
                    x.Text = ""
                    x.SubItems(1) = " "
                    x.SubItems(2) = " "
                    x.SubItems(3) = " "
                    x.SubItems(4) = " "
                    x.SubItems(5) = " "
                    x.SubItems(6) = " "
                    x.SubItems(7) = " "
                    x.SubItems(8) = " "
                    x.SubItems(9) = "0"
                End If
            End With
            .ChildRow = .lstChildren.ListItems.Count
            .lstChildren.ListItems.Item(.ChildRow).SubItems(1) = Format(.ChildRow, "0#")
            .lstChildren.ListItems(.ChildRow).EnsureVisible
            .lstChildren.ListItems(.ChildRow).Selected = True
            .picToolbar.Enabled = False
            .picMain.Enabled = False
            .picSLChild.ZOrder 0
            .txtChildFName.Text = ""
            .txtChildGName.Text = ""
            .txtChildMName.Text = ""
            .txtChildBDate.Text = ""
            .txtChildPicturePath.Text = ""
            .cmbChildStatus.ListIndex = 0
            .imgImageChild.Picture = LoadPicture("")
            .imgImageChild.Visible = False
            .imgPicture2.Visible = True
            .picSLChild.Visible = True
            .TRANSDetail = 1
            .txtChildFName.SetFocus
        Case 2  'Golf
            With .lstOtherGolf.ListItems
                If Trim(.Item(.Count).SubItems(2)) <> "" Then
                    Set x = .Add()
                    x.Text = ""
                    x.SubItems(1) = " "
                    x.SubItems(2) = " "
                    x.SubItems(3) = " "
                End If
            End With
            .MemberRow = .lstOtherGolf.ListItems.Count
            .lstOtherGolf.ListItems.Item(.MemberRow).SubItems(1) = Format(.MemberRow, "0#")
            .lstOtherGolf.ListItems(.MemberRow).EnsureVisible
            .lstOtherGolf.ListItems(.MemberRow).Selected = True
            .picToolbar.Enabled = False
            .picMain.Enabled = False
            .picSLGolf.ZOrder 0
            .txtGolf.Text = ""
            .txtMemberSince.Text = ""
            .picSLGolf.Visible = True
            .TRANSDetail = 1
            .txtGolf.SetFocus
        Case 3  'Card
            With .lstCards.ListItems
                If Trim(.Item(.Count).SubItems(2)) <> "" Then
                    Set x = .Add()
                    x.Text = ""
                    x.SubItems(1) = " "
                    x.SubItems(2) = " "
                    x.SubItems(3) = " "
                End If
            End With
            .CardRow = .lstCards.ListItems.Count
            .lstCards.ListItems.Item(.CardRow).SubItems(1) = Format(.CardRow, "0#")
            .lstCards.ListItems(.CardRow).EnsureVisible
            .lstCards.ListItems(.CardRow).Selected = True
            .picToolbar.Enabled = False
            .picMain.Enabled = False
            .picSLCreditCard.ZOrder 0
            .txtCreditCard.Text = ""
            .txtTypeCredit.Text = ""
            .picSLCreditCard.Visible = True
            .TRANSDetail = 1
            .txtCreditCard.SetFocus
    End Select
End With
End Sub

Private Sub mnuMemberDetailsDelete_Click()
With frmMembershipInformation
    Select Case .FocusDetail
        Case 1  'Child
            With .lstChildren.ListItems
                If .Count > 1 Then
                    .Remove frmMembershipInformation.ChildRow
                    If frmMembershipInformation.ChildRow > .Count Then
                        frmMembershipInformation.ChildRow = .Count
                    End If
                Else
                    .Item(1).SubItems(1) = " "
                    .Item(1).SubItems(2) = " "
                    .Item(1).SubItems(3) = " "
                    .Item(1).SubItems(4) = " "
                    .Item(1).SubItems(5) = " "
                    .Item(1).SubItems(6) = " "
                    .Item(1).SubItems(7) = " "
                    frmMembershipInformation.ChildRow = 1
                End If
            End With
            .lstChildren.ListItems(.ChildRow).EnsureVisible
            .lstChildren.ListItems(.ChildRow).Selected = True
            .lstChildren.SetFocus
        Case 2  'Golf
            With .lstOtherGolf.ListItems
                If .Count > 1 Then
                    .Remove frmMembershipInformation.MemberRow
                    If frmMembershipInformation.MemberRow > .Count Then
                        frmMembershipInformation.MemberRow = .Count
                    End If
                Else
                    .Item(1).SubItems(1) = " "
                    .Item(1).SubItems(2) = " "
                    .Item(1).SubItems(3) = " "
                    frmMembershipInformation.MemberRow = 1
                End If
            End With
            .lstOtherGolf.ListItems(.MemberRow).EnsureVisible
            .lstOtherGolf.ListItems(.MemberRow).Selected = True
            .lstOtherGolf.SetFocus
        Case 3  'Card
            With .lstCards.ListItems
                If .Count > 1 Then
                    .Remove frmMembershipInformation.CardRow
                    If frmMembershipInformation.CardRow > .Count Then
                        frmMembershipInformation.CardRow = .Count
                    End If
                Else
                    .Item(1).SubItems(1) = " "
                    .Item(1).SubItems(2) = " "
                    .Item(1).SubItems(3) = " "
                    frmMembershipInformation.CardRow = 1
                End If
            End With
            .lstCards.ListItems(.CardRow).EnsureVisible
            .lstCards.ListItems(.CardRow).Selected = True
            .lstCards.SetFocus
    End Select
End With
End Sub

Private Sub mnuMemberDetailsEdit_Click()
With frmMembershipInformation
    Select Case .FocusDetail
        Case 1  'Child
            .picToolbar.Enabled = False
            .picMain.Enabled = False
            .picSLChild.ZOrder 0
            .txtChildFName1.Text = .lstChildren.ListItems.Item(.ChildRow).SubItems(6)
            .txtChildGName1.Text = .lstChildren.ListItems.Item(.ChildRow).SubItems(7)
            .txtChildMName1.Text = .lstChildren.ListItems.Item(.ChildRow).SubItems(8)
            .txtChildBDate1.Text = .lstChildren.ListItems.Item(.ChildRow).SubItems(3)
            .txtChildPicturePath1.Text = .lstChildren.ListItems.Item(.ChildRow).SubItems(5)
            .txtChildStatus1.Text = .lstChildren.ListItems.Item(.ChildRow).SubItems(9)
            .txtChildStatusKey1.Text = .lstChildren.ListItems.Item(.ChildRow).SubItems(10)
            
            .txtChildFName.Text = .lstChildren.ListItems.Item(.ChildRow).SubItems(6)
            .txtChildGName.Text = .lstChildren.ListItems.Item(.ChildRow).SubItems(7)
            .txtChildMName.Text = .lstChildren.ListItems.Item(.ChildRow).SubItems(8)
            .txtChildBDate.Text = .lstChildren.ListItems.Item(.ChildRow).SubItems(3)
            .txtChildPicturePath.Text = .lstChildren.ListItems.Item(.ChildRow).SubItems(5)
            .cmbChildStatus.ListIndex = .lstChildren.ListItems.Item(.ChildRow).SubItems(10)
            If Trim(.txtChildPicturePath.Text) = "" Then
                .imgImageChild.Picture = LoadPicture("")
                .imgImageChild.Visible = False
                .imgPicture2.Visible = True
            Else
                .imgImageChild.Picture = LoadPicture(Trim(.txtChildPicturePath.Text))
                .imgImageChild.Visible = True
                .imgPicture2.Visible = False
            End If
            .picSLChild.Visible = True
            .TRANSDetail = 2
            .txtChildFName.SetFocus
        Case 2  'Golf
            .picToolbar.Enabled = False
            .picMain.Enabled = False
            .picSLGolf.ZOrder 0
            .txtGolf1.Text = .lstOtherGolf.ListItems.Item(.MemberRow).SubItems(2)
            .txtMemberSince1.Text = .lstOtherGolf.ListItems.Item(.MemberRow).SubItems(3)
            .txtGolf.Text = .lstOtherGolf.ListItems.Item(.MemberRow).SubItems(2)
            .txtMemberSince.Text = .lstOtherGolf.ListItems.Item(.MemberRow).SubItems(3)
            .picSLGolf.Visible = True
            .TRANSDetail = 2
            .txtGolf.SetFocus
        Case 3  'Card
            .picToolbar.Enabled = False
            .picMain.Enabled = False
            .picSLCreditCard.ZOrder 0
            .txtCreditCard1.Text = .lstCards.ListItems.Item(.CardRow).SubItems(2)
            .txtTypeCredit1.Text = .lstCards.ListItems.Item(.CardRow).SubItems(3)
            .txtCreditCard.Text = .lstCards.ListItems.Item(.CardRow).SubItems(2)
            .txtTypeCredit.Text = .lstCards.ListItems.Item(.CardRow).SubItems(3)
            .picSLCreditCard.Visible = True
            .TRANSDetail = 2
            .txtCreditCard.SetFocus
    End Select
End With

End Sub

Private Sub mnuMemberFindFName_Click()
With frmMembershipInformation
    .isFind = 2
    .picToolbar.Enabled = False
    .picMain.Enabled = False
    .picSearch.ZOrder 0
    .txtSearch.Text = ""
    .picSearch.Visible = True
    .txtSearch.SetFocus
End With
End Sub

Private Sub mnuMemberFindLName_Click()
With frmMembershipInformation
    .isFind = 1
    .picToolbar.Enabled = False
    .picMain.Enabled = False
    .picSearch.ZOrder 0
    .txtSearch.Text = ""
    .picSearch.Visible = True
    .txtSearch.SetFocus
End With
End Sub

Private Sub mnuMemberFindMName_Click()
With frmMembershipInformation
    .isFind = 3
    .picToolbar.Enabled = False
    .picMain.Enabled = False
    .picSearch.ZOrder 0
    .txtSearch.Text = ""
    .picSearch.Visible = True
    .txtSearch.SetFocus
End With
End Sub

Private Sub mnuMemberIDFindIDNumber_Click()
With frmMembershipIDNumber
    '.SearchAdd = 2
    .isFind = 2
    .picSearch.ZOrder 0
    .txtSearch.Text = ""
    .picSearch.Visible = True
    .txtSearch.SetFocus
End With
End Sub

Private Sub mnuMemberIDFindLName_Click()
With frmMembershipIDNumber
    '.SearchAdd = 2
    .isFind = 1
    .picSearch.ZOrder 0
    .txtSearch.Text = ""
    .picSearch.Visible = True
    .txtSearch.SetFocus
End With
End Sub

Private Sub mnuPayrollDeductionReportEmployee_Click()
With frmPersonnelDeductions
    .isAddPrint = 2
    .b8TitleBar2.Caption = "Employee Active Deduction Balance"
    .Label4.Caption = "as of"
    .picAdd.ZOrder 0
    .txtSearchAdd.Text = ""
    .txtPayrollDateAdd.Text = Format(Date, "mm/dd/yyyy")
    .picMain.Enabled = False
    .picToolbar.Enabled = False
    .picAdd.Visible = True
    .txtSearchAdd.SetFocus
End With
End Sub

Private Sub mnuPayrollDeductionReportSummary_Click()
With frmPersonnelDeductions
    .picPrintSumm.ZOrder 0
    .cmbDivision.ListIndex = -1
    .txtAsOf.Text = Format(Date, "mm/dd/yyyy")
    .picMain.Enabled = False
    .picToolbar.Enabled = False
    .picPrintSumm.Visible = True
    .cmbDivision.SetFocus
End With
End Sub

Private Sub mnuPayrollHourPostingBatch_Click()
With frmPersonnelHours
    .picBatchPosting.ZOrder 0
    .picMain.Enabled = False
    .picToolbar.Enabled = False
    .cmbPostUnpost.Text = ""
    .cmbPostUnpost.ListIndex = -1
    .cmbDivisionBatchPost.Text = " "
    .cmbDivisionBatchPost.ListIndex = -1
    .txtPayrollDatePostUnpost.Text = ""
    .picBatchPosting.Visible = True
    .cmbPostUnpost.SetFocus
End With
End Sub

Private Sub mnuPayrollHourPostingSingleTrans_Click()
With frmPersonnelHours
    On Error GoTo PG:
    If .imgPosted.Visible = True Then
        If AccessRights("Personnel - Hours", "UnPost") = False Then
            MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
                   "ACCESS DENIED!                                      ", vbCritical, "Alert"
            Exit Sub
        End If
        
        If MsgBox("ARE YOU SURE IN UNPOSTING THIS TRANSACTION?                        ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
        
        s = "SELECT TOP (1) PayrollPeriodKey, Locked " & _
            " From dbo.tbl_Personnel_Payroll " & _
            " WHERE (PK = " & .locPayrollKey & ") " & _
            " AND (Locked = 1)"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            MsgBox "This payroll was already locked!                     ", vbCritical, "Error..."
            rs.Close
            Exit Sub
        End If
        rs.Close
        Arr = Split(Trim(.txtCutOffDate.Text), " - ", -1, 1)
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll_Earnings WHERE (MasterKey = " & .locPayrollKey & ")"
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll_Deductions WHERE (MasterKey = " & .locPayrollKey & ")"
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll_EmployerShare WHERE (MasterKey = " & .locPayrollKey & ")"
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Loans_SL " & _
                          " WHERE (PayrollKey = " & .locPayrollKey & ") " & _
                          " AND (TransactionDate = '" & FormatDateTime(.txtPayrollPeriod.Text, vbShortDate) & "') " & _
                          " AND (InOut = 'O')"
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Deduction_SL " & _
                          " WHERE (PayrollKey = " & .locPayrollKey & ") " & _
                          " AND (TransactionDate = '" & FormatDateTime(.txtPayrollPeriod.Text, vbShortDate) & "') " & _
                          " AND (TransactionType = 2) " & _
                          " AND (InOut = 'O')"
        
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll WHERE (PK = " & .locPayrollKey & ")"
        
        s = "SELECT COUNT(*) AS RecCnt " & _
            " From dbo.tbl_Personnel_Payroll " & _
            " WHERE (ActionMemoKey = " & .locActionMemoKey & ") "
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            If CDbl(rs!RecCnt) = 0 Then
                ConnOmega.Execute "UPDATE tbl_Personnel_ActionNew SET Locked = 0 WHERE (PK = " & .locActionMemoKey & ")"
            End If
        End If
        rs.Close
        
        ConnOmega.Execute "UPDATE tbl_Personnel_Hours " & _
                          " SET PayrollKey = Null, " & _
                          " Posted = 0, " & _
                          " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                          " WHERE (PK = " & .Statusbar1.Panels(1).Text & ")"
        
        .BROWSER GetSetting(App.EXEName, "PersonnelHours", "PersonnelHours", ""), "is_LOAD"
        
    Else
        If AccessRights("Personnel - Hours", "Post") = False Then
            MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
                   "ACCESS DENIED!                                      ", vbCritical, "Alert"
            Exit Sub
        End If
        
        s = "SELECT tbl_Personnel_Deduction_forPayroll.* " & _
            " FROM tbl_Personnel_Deduction_forPayroll " & _
            " WHERE (DivisionKey = " & .locDivision & ")  " & _
            " AND (PayrollPeriodKey = " & .locPayrollPeroid & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            MsgBox "Please add for deduction for this division and payroll date!                    ", vbCritical, "Error..."
            rs.Close
            Exit Sub
        Else
            If CDbl(rs!Posted) = 0 Then
                MsgBox "Please post the for deduction for this division and payroll date!                    ", vbCritical, "Error..."
                rs.Close
                Exit Sub
            End If
        End If
        rs.Close
        
        If MsgBox("ARE YOU SURE IN POSTING THIS TRANSACTION?                        ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
        
        s = "SELECT TOP (1) PayrollPeriodKey, Locked " & _
            " From dbo.tbl_Personnel_Payroll " & _
            " WHERE (PayrollPeriodKey = " & .locPayrollPeroid & ") " & _
            " AND (Locked = 1)"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            MsgBox "This payroll period was already locked!                     ", vbCritical, "Error..."
            rs.Close
            Exit Sub
        End If
        rs.Close
        
        COMPUTE_COMPENSATION .Statusbar1.Panels(1).Text
        .BROWSER GetSetting(App.EXEName, "PersonnelHours", "PersonnelHours", ""), "is_LOAD"
        
        If AccessRights("Personnel Compensation", "Open") = False Then Exit Sub
        
        If MsgBox("Successfully Posted!                     " & vbCrLf & vbCrLf & _
                  "View Compensation Module?                ", vbInformation + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
        
        gbl_Form_Caption = "Compensation"
        If IsLoaded(frmPersonnelPayroll) Then frmPersonnelPayroll.ZOrder 0 Else frmPersonnelPayroll.Show
        frmPersonnelPayroll.BROWSER .locPayrollKey, "is_FIND"
    End If
End With
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub mnuPayrollPrint1RnFSup_Click(Index As Integer)
With frmPersonnelPayrollReport
    t = "SELECT PK " & _
        " FROM tbl_Personnel_Position_Level " & _
        " WHERE (LevelName = '" & FORMATSQL(CStr(mnuPayrollPrint1RnFSup(Index).Caption)) & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        .PostLevel = rt!PK
    End If
    rt.Close
    Select Case .lstReportType.ItemData(.lstReportType.ListIndex)
        Case 1: .TimerPaySlip.Enabled = True
        Case 2: .TimerPaySlip.Enabled = True
        Case 3: .TimerEarnings.Enabled = True
        Case 4: .TimerDeductions.Enabled = True
        Case 5: .TimerLoans.Enabled = True
        Case 6: .TimerContri.Enabled = True
        Case 8: .Timer13Month.Enabled = True
        Case 9: .TimerForATM2.Enabled = True
    End Select
End With
End Sub

Private Sub mnuPayrollPrintRankNFile_Click()
With frmPersonnelPayrollReport
    .PostLevel = 1
    Select Case .lstReportType.ItemData(.lstReportType.ListIndex)
        Case 1: .TimerPaySlip.Enabled = True
        Case 2: .TimerPaySlip.Enabled = True
        Case 3: .TimerEarnings.Enabled = True
        Case 4: .TimerDeductions.Enabled = True
        Case 5: .TimerLoans.Enabled = True
        Case 6: .TimerContri.Enabled = True
        Case 8: .Timer13Month.Enabled = True
        Case 9: .TimerForATM2.Enabled = True
    End Select
End With
End Sub

Private Sub mnuPayrollPrintSupervisory_Click()
With frmPersonnelPayrollReport
    .PostLevel = 2
    Select Case .lstReportType.ItemData(.lstReportType.ListIndex)
        Case 1: .TimerPaySlip.Enabled = True
        Case 2: .TimerPaySlip.Enabled = True
        Case 3: .TimerEarnings.Enabled = True
        Case 4: .TimerDeductions.Enabled = True
        Case 5: .TimerLoans.Enabled = True
        Case 6: .TimerContri.Enabled = True
        Case 8: .Timer13Month.Enabled = True
        Case 9: .TimerForATM.Enabled = True
    End Select
End With
End Sub

Private Sub mnuPlayerAddFromExcel_Click()
With frmPlayerSetup
    .CommonDialog1.DialogTitle = "OPEN FILE"
    .CommonDialog1.Filename = ""
    .CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
    .CommonDialog1.FilterIndex = 1
    .CommonDialog1.ShowOpen
    strPath = .CommonDialog1.Filename
    If Trim(strPath) = "" Then Exit Sub
    .txtPath.Text = strPath
'    .txtPath.SetFocus
    
    'Timer1.Enabled = True
    .Timer2.Enabled = True
End With
End Sub

Private Sub mnuPlayerAddIndividual_Click()
With frmPlayerSetup
    .CLEARTEXT
    .LOCKTEXT False
    .TOOLBARFUNC 2
    .TRANSACTIONTYPE = 1
    .txtLName.SetFocus
End With
End Sub

Private Sub mnuPrintSystem36Class_Click(Index As Integer)
With frmScoreCardsSystem36
    .ReportClass = mnuPrintSystem36Class(Index).Caption
    .TimerPrintSummary.Enabled = True
'    .picToolbar.Enabled = False
'    .picMain.Enabled = False
'    .picPrint.ZOrder 0
'    .cmbGrossNet.ListIndex = -1
'    .picPrint.Visible = True
'    .cmbGrossNet.SetFocus
End With
End Sub

Private Sub mnuPrintSystem36Result_Click()
frmScoreCardsSystem36.TimerPrintResult.Enabled = True
End Sub

Private Sub mnuRegistrationAddBagTagNo_Click()
With frmOperationRegistration
    If AccessRights("Registration", "Add") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    .iAddType = 1
    .picAddBagDrop.ZOrder 0
    .txtSearchBagTag.Text = ""
    .picAddBagDrop.Visible = True
    .txtSearchBagTag.SetFocus
End With
End Sub

Private Sub mnuRegistrationAddPlayerName_Click()
With frmOperationRegistration
    If AccessRights("Registration", "Add") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    .iAddType = 2
    .picAddBagDrop.ZOrder 0
    .txtSearchBagTag.Text = ""
    .picAddBagDrop.Visible = True
    .txtSearchBagTag.SetFocus
End With
End Sub

Private Sub mnuRRFindPONumber_Click()
With frmInvRR
    .CLEARTEXT
    .TOOLBARFUNC 3
    .TRANSACTIONTYPE = 3
    .Caption = "RECEIVING REPORT - FIND"
    .txtPONumber.Locked = False
    .txtPONumber.BackColor = &H80000005
    .txtRRNumber.BackColor = &HE0E0E0
    .txtPONumber.SetFocus
End With
End Sub

Private Sub mnuRRFindRRNumber_Click()
With frmInvRR
    .CLEARTEXT
    .TOOLBARFUNC 3
    .TRANSACTIONTYPE = 3
    .Caption = "RECEIVING REPORT - FIND"
    .txtRRNumber.Locked = False
    .txtRRNumber.BackColor = &H80000005
    .txtPONumber.BackColor = &HE0E0E0
    .txtRRNumber.SetFocus
End With
End Sub

Private Sub mnuRRPostingInvoice_Click()
If AccessRights("Receiving Report", "Post Inv") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
End Sub

Private Sub mnuRRPostingReceived_Click()
With frmInvRR
    If .Statusbar1.Panels(1).Text = "" Then Exit Sub
    If .TRANSACTIONTYPE <> 0 Then Exit Sub
    If .picSLine.Visible = True Then Exit Sub
    If .picPost.Visible = True Then Exit Sub
    If AccessRights("Receiving Report", "Post Rcd") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If Trim(.txtInvNumber.Text) = "" Then MsgBox "Please Supply Invoice Number!              ", vbCritical, "Error...": .txtInvNumber.SetFocus: Exit Sub
    If IsDate(.txtInvDate.Text) = False Then MsgBox "Please Supply a Valid Invoice Date!                 ", vbCritical, "Error...": .txtInvDate.SetFocus: Exit Sub
    If RETURNTEXTVALUE(.txtInvGross) <= 0 Then MsgBox "Please Supply a Valid Amount!               ", vbCritical, "Error...": .txtInvGross.SetFocus: Exit Sub
    If RETURNTEXTVALUE(.txtInvNet) <= 0 Then MsgBox "Please Supply a Valid Amount!               ", vbCritical, "Error...": .txtInvNet.SetFocus: Exit Sub
    .BROWSER GetSetting(App.EXEName, "PONumberRR", "PONumRR", ""), "is_LOAD"
    If .imgPosted.Visible = True Then MsgBox "ALREADY POSTED!                     ", vbCritical, "Error...": Exit Sub
    
    .cmbLocation.Clear
    s = "SELECT tbl_Inv_Location.* " & _
        " FROM tbl_Inv_Location " & _
        " ORDER BY LocName"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        .cmbLocation.AddItem rs!LocName
        .cmbLocation.ItemData(.cmbLocation.NewIndex) = rs!PK
        rs.MoveNext
    Wend
    rs.Close
    .txtRRDatePosting.Text = Format(Now, "mm/dd/yyyy")
    .picToolbar.Enabled = False
    .picBody.Enabled = False
    .picPost.ZOrder 0
    .picPost.Visible = True
    .cmbLocation.SetFocus

End With
End Sub

Private Sub mnuScoringLocationName_Click(Index As Integer)

s = "SELECT dbo.tbl_Scoring_TournamentInfo_Location.MasterKey, " & _
    " dbo.tbl_Scoring_Location.ScoringLocation, " & _
    " dbo.tbl_Scoring_Location.PK " & _
    " FROM dbo.tbl_Scoring_TournamentInfo_Location LEFT OUTER JOIN " & _
    " dbo.tbl_Scoring_Location ON dbo.tbl_Scoring_TournamentInfo_Location.LocationKey = dbo.tbl_Scoring_Location.PK " & _
    " WHERE (dbo.tbl_Scoring_TournamentInfo_Location.MasterKey = " & TournamentKey & ") " & _
    " AND (dbo.tbl_Scoring_Location.ScoringLocation = '" & mnuScoringLocationName(Index).Caption & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iLocationKey = rs!PK
End If
rs.Close

iFilterIndex = 0
With MainForm.CommonDialog1
    .CancelError = True
    On Error GoTo ErrorHandler
    .DialogTitle = "Save"
    .Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx|Text File|*.txt"
    iFilterIndex = .FilterIndex
    .ShowSave
    Filename = Trim(.Filename)
End With
Arr = Split(Trim(Filename), "\", -1, 1)
Arr1 = Split(CStr(Arr(UBound(Arr))), ".", -1, 1)

On Error GoTo PG:
Screen.MousePointer = vbHourglass
If Arr1(UBound(Arr1)) = "txt" Then    'Text File
    Open Filename For Output As #1
        'Location
        s = "SELECT tbl_Scoring_Location.* " & _
            " FROM tbl_Scoring_Location " & _
            " WHERE (PK = " & iLocationKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        While Not rs.EOF
            sLine = "LOCATION["
            For i = 0 To rs.Fields.Count - 1
                sLine = sLine & rs.Fields(i).Value & "|"
            Next i
            Print #1, Mid(CStr(sLine), 1, Len(sLine) - 1)
            rs.MoveNext
        Wend
        rs.Close
        'Location Details
        s = "SELECT tbl_Scoring_Location_Details.* " & _
            " FROM tbl_Scoring_Location_Details " & _
            " WHERE (MasterKey = " & iLocationKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        While Not rs.EOF
            sLine = "LOCATION_DETAILS["
            For i = 0 To rs.Fields.Count - 1
                sLine = sLine & rs.Fields(i).Value & "|"
            Next i
            Print #1, Mid(CStr(sLine), 1, Len(sLine) - 1)
            rs.MoveNext
        Wend
        rs.Close
        'Scoring System
        s = "SELECT tbl_Scoring_System.* " & _
            " FROM tbl_Scoring_System "
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        While Not rs.EOF
            sLine = "SCORING_SYSTEM["
            For i = 0 To rs.Fields.Count - 1
                sLine = sLine & rs.Fields(i).Value & "|"
            Next i
            Print #1, Mid(CStr(sLine), 1, Len(sLine) - 1)
            rs.MoveNext
        Wend
        rs.Close
        'Tournament Info
        s = "SELECT tbl_Scoring_TournamentInfo.* " & _
            " FROM tbl_Scoring_TournamentInfo " & _
            " WHERE (PK = " & TournamentKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        While Not rs.EOF
            sLine = "TOURNAMENT_INFO["
            For i = 0 To rs.Fields.Count - 1
                sLine = sLine & rs.Fields(i).Value & "|"
            Next i
            Print #1, Mid(CStr(sLine), 1, Len(sLine) - 1)
            rs.MoveNext
        Wend
        rs.Close
        'Tournament Info Class
        s = "SELECT tbl_Scoring_TournamentInfo_Class.* " & _
            " FROM tbl_Scoring_TournamentInfo_Class " & _
            " WHERE (TournamentKey = " & TournamentKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        While Not rs.EOF
            sLine = "TOURNAMENT_INFO_CLASS["
            For i = 0 To rs.Fields.Count - 1
                sLine = sLine & rs.Fields(i).Value & "|"
            Next i
            Print #1, Mid(CStr(sLine), 1, Len(sLine) - 1)
            rs.MoveNext
        Wend
        rs.Close
        'Tournament Info Index
        s = "SELECT tbl_Scoring_TournamentInfo_Index.* " & _
            " FROM tbl_Scoring_TournamentInfo_Index " & _
            " WHERE (TournamentKey = " & TournamentKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        While Not rs.EOF
            sLine = "TOURNAMENT_INFO_INDEX["
            For i = 0 To rs.Fields.Count - 1
                sLine = sLine & rs.Fields(i).Value & "|"
            Next i
            Print #1, Mid(CStr(sLine), 1, Len(sLine) - 1)
            rs.MoveNext
        Wend
        rs.Close
        'Tournament Info Location
        s = "SELECT tbl_Scoring_TournamentInfo_Location.* " & _
            " FROM tbl_Scoring_TournamentInfo_Location " & _
            " WHERE (MasterKey = " & TournamentKey & ") " & _
            " AND (LocationKey = " & iLocationKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        While Not rs.EOF
            sLine = "TOURNAMENT_INFO_LOCATION["
            For i = 0 To rs.Fields.Count - 1
                sLine = sLine & rs.Fields(i).Value & "|"
            Next i
            Print #1, Mid(CStr(sLine), 1, Len(sLine) - 1)
            rs.MoveNext
        Wend
        rs.Close
        'Player Name
        s = "SELECT tbl_Scoring_PlayerName.* " & _
            " FROM tbl_Scoring_PlayerName " & _
            " WHERE (TournamentKey = " & TournamentKey & ") "
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        While Not rs.EOF
            sLine = "PLAYER_NAME["
            For i = 0 To rs.Fields.Count - 1
                sLine = sLine & rs.Fields(i).Value & "|"
            Next i
            Print #1, Mid(CStr(sLine), 1, Len(sLine) - 1)
            rs.MoveNext
        Wend
        rs.Close
    Close #1
    
    Screen.MousePointer = vbDefault
    
    If MsgBox("Would you like to open the file just saved?              ", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm") = vbYes Then
        Shell "Notepad.exe " & Filename, vbMaximizedFocus
    End If
    
Else    'Excel File
    WorkbookName = Filename
    ColCnt = 0: RowCnt = 0
    Set xlsApp = CreateObject("Excel.Application")
    With xlsApp
        .Visible = False
        .Workbooks.Add
        .DisplayAlerts = False
        iWorkSheet = 0
        .Workbooks(1).Sheets(2).Delete
        .Workbooks(1).Sheets(2).Delete
        
        For i = 1 To 7
            xlsApp.Workbooks(1).Sheets.Add
        Next i
        
        s = "SELECT tbl_Scoring_Location.* " & _
            " FROM tbl_Scoring_Location " & _
            " WHERE (PK = " & iLocationKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            iWorkSheet = iWorkSheet + 1
            .Workbooks(1).Sheets(iWorkSheet).Activate
            .Workbooks(1).Sheets(iWorkSheet).Name = "Location"
            RowCnt = 0: ColCnt = 0
            With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                RowCnt = RowCnt + 1
                For i = 0 To rs.Fields.Count - 2 '1
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = rs.Fields(i).Name
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                Next i
                RowCnt = RowCnt + 1: ColCnt = 0
                For i = 0 To rs.Fields.Count - 2 '1
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = rs.Fields(i).Value
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                Next i
            End With
        End If
        rs.Close
        
        
        s = "SELECT tbl_Scoring_Location_Details.* " & _
            " FROM tbl_Scoring_Location_Details " & _
            " WHERE (MasterKey = " & iLocationKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            iWorkSheet = iWorkSheet + 1
            .Workbooks(1).Sheets(iWorkSheet).Activate
            .Workbooks(1).Sheets(iWorkSheet).Name = "Location_Details"
            RowCnt = 0: ColCnt = 0
            With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                RowCnt = RowCnt + 1
                For i = 0 To rs.Fields.Count - 1
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = rs.Fields(i).Name
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                Next i
                While Not rs.EOF
                    RowCnt = RowCnt + 1: ColCnt = 0
                    For i = 0 To rs.Fields.Count - 1
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = rs.Fields(i).Value
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                    Next i
                    rs.MoveNext
                Wend
            End With
        End If
        rs.Close
        
        s = "SELECT tbl_Scoring_System.* " & _
            " FROM tbl_Scoring_System "
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            iWorkSheet = iWorkSheet + 1
            .Workbooks(1).Sheets(iWorkSheet).Activate
            .Workbooks(1).Sheets(iWorkSheet).Name = "Scoring_System"
            RowCnt = 0: ColCnt = 0
            With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                RowCnt = RowCnt + 1
                For i = 0 To rs.Fields.Count - 1
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = rs.Fields(i).Name
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                Next i
                While Not rs.EOF
                    RowCnt = RowCnt + 1: ColCnt = 0
                    For i = 0 To rs.Fields.Count - 1
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = rs.Fields(i).Value
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                    Next i
                    rs.MoveNext
                Wend
            End With
        End If
        rs.Close
        
        s = "SELECT tbl_Scoring_TournamentInfo.* " & _
            " FROM tbl_Scoring_TournamentInfo " & _
            " WHERE (PK = " & TournamentKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            iWorkSheet = iWorkSheet + 1
            .Workbooks(1).Sheets(iWorkSheet).Activate
            .Workbooks(1).Sheets(iWorkSheet).Name = "TournamentInfo"
            RowCnt = 0: ColCnt = 0
            With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                RowCnt = RowCnt + 1
                For i = 0 To rs.Fields.Count - 1
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = rs.Fields(i).Name
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                Next i
                RowCnt = RowCnt + 1: ColCnt = 0
                For i = 0 To rs.Fields.Count - 1
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = rs.Fields(i).Value
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                Next i
            End With
        End If
        rs.Close
        
        s = "SELECT tbl_Scoring_TournamentInfo_Class.* " & _
            " FROM tbl_Scoring_TournamentInfo_Class " & _
            " WHERE (TournamentKey = " & TournamentKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            iWorkSheet = iWorkSheet + 1
            .Workbooks(1).Sheets(iWorkSheet).Activate
            .Workbooks(1).Sheets(iWorkSheet).Name = "TournamentInfo_Class"
            RowCnt = 0: ColCnt = 0
            With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                RowCnt = RowCnt + 1
                For i = 0 To rs.Fields.Count - 1
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = rs.Fields(i).Name
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                Next i
                While Not rs.EOF
                    RowCnt = RowCnt + 1: ColCnt = 0
                    For i = 0 To rs.Fields.Count - 1
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = rs.Fields(i).Value
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                    Next i
                    rs.MoveNext
                Wend
            End With
        End If
        rs.Close
        
        s = "SELECT tbl_Scoring_TournamentInfo_Index.* " & _
            " FROM tbl_Scoring_TournamentInfo_Index " & _
            " WHERE (TournamentKey = " & TournamentKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            iWorkSheet = iWorkSheet + 1
            .Workbooks(1).Sheets(iWorkSheet).Activate
            .Workbooks(1).Sheets(iWorkSheet).Name = "TournamentInfo_Index"
            RowCnt = 0: ColCnt = 0
            With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                RowCnt = RowCnt + 1
                For i = 0 To rs.Fields.Count - 1
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = rs.Fields(i).Name
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                Next i
                While Not rs.EOF
                    RowCnt = RowCnt + 1: ColCnt = 0
                    For i = 0 To rs.Fields.Count - 1
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = rs.Fields(i).Value
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                    Next i
                    rs.MoveNext
                Wend
            End With
        End If
        rs.Close
        
        s = "SELECT tbl_Scoring_TournamentInfo_Location.* " & _
            " FROM tbl_Scoring_TournamentInfo_Location " & _
            " WHERE (MasterKey = " & TournamentKey & ") " & _
            " AND (LocationKey = " & iLocationKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            iWorkSheet = iWorkSheet + 1
            .Workbooks(1).Sheets(iWorkSheet).Activate
            .Workbooks(1).Sheets(iWorkSheet).Name = "TournamentInfo_Location"
            RowCnt = 0: ColCnt = 0
            With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                RowCnt = RowCnt + 1
                For i = 0 To rs.Fields.Count - 1
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = rs.Fields(i).Name
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                Next i
                While Not rs.EOF
                    RowCnt = RowCnt + 1: ColCnt = 0
                    For i = 0 To rs.Fields.Count - 1
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = rs.Fields(i).Value
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                    Next i
                    rs.MoveNext
                Wend
            End With
        End If
        rs.Close
        
        s = "SELECT tbl_Scoring_PlayerName.* " & _
            " FROM tbl_Scoring_PlayerName " & _
            " WHERE (TournamentKey = " & TournamentKey & ") "
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            iWorkSheet = iWorkSheet + 1
            .Workbooks(1).Sheets(iWorkSheet).Activate
            .Workbooks(1).Sheets(iWorkSheet).Name = "PlayerName"
            RowCnt = 0: ColCnt = 0
            With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                RowCnt = RowCnt + 1
                For i = 0 To rs.Fields.Count - 1
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = rs.Fields(i).Name
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                Next i
                While Not rs.EOF
                    RowCnt = RowCnt + 1: ColCnt = 0
                    For i = 0 To rs.Fields.Count - 1
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = rs.Fields(i).Value
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                    Next i
                    rs.MoveNext
                Wend
            End With
        End If
        rs.Close
        
        If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
        .ActiveWorkbook.SaveAs Filename:=WorkbookName
        .Visible = True
        Set xlsApp = Nothing
        
    End With
    Screen.MousePointer = vbDefault
End If

Exit Sub
ErrorHandler:
Exit Sub

Exit Sub
PG:
Screen.MousePointer = vbDefault
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub mnuScoringLocationNameAdd_Click(Index As Integer)
s = "SELECT dbo.tbl_Scoring_TournamentInfo_Location.MasterKey, " & _
    " dbo.tbl_Scoring_Location.ScoringLocation, " & _
    " dbo.tbl_Scoring_Location.PK " & _
    " FROM dbo.tbl_Scoring_TournamentInfo_Location LEFT OUTER JOIN " & _
    " dbo.tbl_Scoring_Location ON dbo.tbl_Scoring_TournamentInfo_Location.LocationKey = dbo.tbl_Scoring_Location.PK " & _
    " WHERE (dbo.tbl_Scoring_TournamentInfo_Location.MasterKey = " & TournamentKey & ") " & _
    " AND (dbo.tbl_Scoring_Location.ScoringLocation = '" & mnuScoringLocationNameAdd(Index).Caption & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iLocationKey = rs!PK
End If
rs.Close

s = "SELECT tbl_Scoring_TournamentInfo_Location.* " & _
    " FROM tbl_Scoring_TournamentInfo_Location " & _
    " WHERE (LocationKey = " & iLocationKey & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    If rs!HomeCourt = 1 Then
        'With frmScoreCard
        With frmScoreCardAll
            .picMain.Enabled = False
            .picToolbar.Enabled = False
            .picSearchAdd.ZOrder 0
            .txtSearchAdd.Text = ""
            .txtDateAdd.Text = Format(FormatDateTime(Date, vbShortDate), "mm/dd/yyyy")
            .picSearchAdd.Visible = True
            .txtSearchAdd.SetFocus
            LocationKey = iLocationKey
        End With
    Else
        
        If MsgBox("LOAD FROM FILE?                     ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbYes Then
        
            MainForm.CommonDialog1.DialogTitle = "OPEN FILE"
            MainForm.CommonDialog1.Filename = ""
            MainForm.CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx|Text File|*.txt"
            MainForm.CommonDialog1.FilterIndex = 1
            MainForm.CommonDialog1.ShowSave
            frmScoreCardAll.txtPath.Text = MainForm.CommonDialog1.Filename
            If Trim(frmScoreCardAll.txtPath.Text) = "" Then Exit Sub
            frmScoreCardAll.TimerAddLocation.Enabled = True
            
'            frmScoreCard.txtPath.Text = MainForm.CommonDialog1.Filename
'            If Trim(frmScoreCard.txtPath.Text) = "" Then Exit Sub
'            frmScoreCard.TimerAddLocation.Enabled = True
            
        Else
            'With frmScoreCard
            With frmScoreCardAll
                .picMain.Enabled = False
                .picToolbar.Enabled = False
                .picSearchAdd.ZOrder 0
                .txtSearchAdd.Text = ""
                .txtDateAdd.Text = Format(FormatDateTime(Date, vbShortDate), "mm/dd/yyyy")
                .picSearchAdd.Visible = True
                .txtSearchAdd.SetFocus
                LocationKey = iLocationKey
            End With
        End If
        
    End If
End If
rs.Close

Exit Sub
ErrorHandler:
Exit Sub
End Sub

Private Sub mnuSupplierReportSL_Click()
frmInvSupplierSL.iSupplier = frmInvSupplier.Statusbar1.Panels(1).Text
frmInvSupplierSL.txtSupplier.Text = frmInvSupplier.txtSuppCode & " - " & frmInvSupplier.txtSuppName.Text
If IsLoaded(frmInvSupplierSL) Then frmInvSupplierSL.ZOrder 0 Else frmInvSupplierSL.Show
End Sub

Private Sub mnuTournamentInfoPrintScoreCard_Click()
If frmTournamentSetup.cmbScoring.ListIndex <= 0 Then Exit Sub
If frmTournamentSetup.cmbScoring.ListIndex = 4 Then Exit Sub
With frmTournamentSetup
    .picPrint.Visible = True
    .picMain.Enabled = False
    .picToolbar.Enabled = False
End With
End Sub

Private Sub ProfilePrintActive_Click()
frmPersonnelInformation.TimerActive.Enabled = True
End Sub

Private Sub ProfilePrintAlphalistActive_Click()
With frmPersonnelInformation
    .picMain.Enabled = False
    .picToolbar.Enabled = False
    .picAlphalist.ZOrder 0
    .txtAsOf.Text = Format(Date, "mm/dd/yyyy")
    .picAlphalist.Visible = True
    .txtAsOf.SetFocus
End With
End Sub

Private Sub ProfilePrintHeadCount_Click()
frmPersonnelInformation.TimerHeadCount.Enabled = True
End Sub

Private Sub ProfilePrintInactive_Click()
frmPersonnelInformation.TimerInactive.Enabled = True
End Sub

Private Sub ProfilePrintProfile_Click()
frmProgressBar.picAlphalist.Visible = False
frmProgressBar.iEmployee = frmPersonnelInformation.Statusbar1.Panels(1).Text
frmProgressBar.TimerProfile.Enabled = True
frmProgressBar.Width = 6260
frmProgressBar.Height = 970
frmProgressBar.Show 1
End Sub

Private Sub ProfileSearchFName_Click()
With frmPersonnelInformation
    .SearchType = 11
    .picToolbar.Enabled = False
    .picMain.Enabled = False
    .picSearch.ZOrder 0
    .txtSearch.Text = ""
    .picSearch.Visible = True
    .txtSearch.SetFocus
End With
End Sub

Private Sub ProfileSearchLName_Click()
With frmPersonnelInformation
    .SearchType = 10
    .picToolbar.Enabled = False
    .picMain.Enabled = False
    .picSearch.ZOrder 0
    .txtSearch.Text = ""
    .picSearch.Visible = True
    .txtSearch.SetFocus
End With
End Sub

Private Sub ProfileSearchMName_Click()
With frmPersonnelInformation
    .SearchType = 12
    .picToolbar.Enabled = False
    .picMain.Enabled = False
    .picSearch.ZOrder 0
    .txtSearch.Text = ""
    .picSearch.Visible = True
    .txtSearch.SetFocus
End With
End Sub
