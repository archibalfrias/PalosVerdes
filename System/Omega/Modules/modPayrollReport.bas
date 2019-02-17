Attribute VB_Name = "modPayrollReport"
Option Explicit

Dim i, k, l, iPK, sCompRate, dGross, dDed, iLine, iLineLoanBal, iEarnCol, iDedCol, sAccName, sAccKey
Dim sQry, iColMax, iColCnt, iRowCnt, ArrAccName, ArrAccName1, ArrAccKey, ArrAccKey1, iFieldCnt
Dim sHours, ArrAccKeyHours, ArrAccKeyHours1, dLoanBal
Dim dblSSS, dblSSSEmpr, dblSSSEC, dblPHIC, dblPHICEmpr, dblPagIbig, dblPagIbigEmpr, dblWHT
Dim dblSSSLoan, dblSSSLoanBal, dblPagIbigLoan, dblPagIbigLoanBal
Dim iMonth1, iMonth2, iMonth3, iYear1, iYear2, iYear3, sDeptName

Dim iRec, dblGrossIncome, iTerms, dTimeSumm, dDateFrom, dDateTo

Dim dSSSLoan, dPGBGLoan, dSSS, dPHIC, dPGBG, dWHT, sRate

Public Sub Generate13thMonth(sUser, DivKey, DivName, PostLevel, iMonth, iYear)
ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll_13thMonth WHERE (LogInName = '" & sUser & "')"
MainForm.picProgressBar.BackColor = &H8000000F
Dim PayrollDate
PayrollDate = DateSerial(iYear, iMonth, 1)
DoEvents
'Select Case Month(FormatDateTime(PayrollDate, vbShortDate))
Select Case iMonth
    Case 3
        iMonth1 = 12
        iMonth2 = 1
        iMonth3 = 2
        iYear1 = CDbl(iYear) - 1
        iYear2 = iYear
        iYear3 = iYear
    Case 6
        iMonth1 = 3
        iMonth2 = 4
        iMonth3 = 5
        iYear1 = iYear
        iYear2 = iYear
        iYear3 = iYear
    Case 9
        iMonth1 = 6
        iMonth2 = 7
        iMonth3 = 8
        iYear1 = iYear
        iYear2 = iYear
        iYear3 = iYear
    Case 12
        iMonth1 = 9
        iMonth2 = 10
        iMonth3 = 11
        iYear1 = iYear
        iYear2 = iYear
        iYear3 = iYear
    Case Else: Exit Sub
End Select

'Select Case Month(FormatDateTime(PayrollDate, vbShortDate))
'    Case 12
'        iMonth1 = 9
'        iMonth2 = 10
'        iMonth3 = 11
'        iYear1 = Year(FormatDateTime(PayrollDate, vbShortDate))
'        iYear2 = Year(FormatDateTime(PayrollDate, vbShortDate))
'        iYear3 = Year(FormatDateTime(PayrollDate, vbShortDate))
'    Case 3
'        iMonth1 = 12
'        iMonth2 = 1
'        iMonth3 = 2
'        iYear1 = Year(FormatDateTime(PayrollDate, vbShortDate)) - 1
'        iYear2 = Year(FormatDateTime(PayrollDate, vbShortDate))
'        iYear3 = Year(FormatDateTime(PayrollDate, vbShortDate))
'    Case 6
'        iMonth1 = 3
'        iMonth2 = 4
'        iMonth3 = 5
'        iYear1 = Year(FormatDateTime(PayrollDate, vbShortDate))
'        iYear2 = Year(FormatDateTime(PayrollDate, vbShortDate))
'        iYear3 = Year(FormatDateTime(PayrollDate, vbShortDate))
'    Case 9
'        iMonth1 = 6
'        iMonth2 = 7
'        iMonth3 = 8
'        iYear1 = Year(FormatDateTime(PayrollDate, vbShortDate))
'        iYear2 = Year(FormatDateTime(PayrollDate, vbShortDate))
'        iYear3 = Year(FormatDateTime(PayrollDate, vbShortDate))
'    Case Else: Exit Sub
'End Select

a = "SELECT tbl_Personnel_Payroll_1.EmployeeKey, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName + ',  ' + dbo.tbl_Personnel_Information.FirstName + '  ' + dbo.tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
    " ISNULL((SELECT SUM(dbo.tbl_Personnel_Payroll_Earnings.TotalAmount) AS Amount FROM  dbo.tbl_Personnel_Payroll_Earnings LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN dbo.tbl_Personnel_Payroll_Earnings_Table ON dbo.tbl_Personnel_Payroll_Earnings.EarningKey = dbo.tbl_Personnel_Payroll_Earnings_Table.PK LEFT OUTER JOIN dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
    " WHERE (dbo.tbl_Personnel_Payroll_Earnings_Table.Month13 = 1) " & _
    " AND (dbo.tbl_Personnel_Payroll.EmployeeKey = tbl_Personnel_Payroll_1.EmployeeKey) " & _
    " AND (MONTH(dbo.tbl_Personnel_Compensation_Period.PayrollDate) = " & iMonth1 & ") " & _
    " AND (YEAR(dbo.tbl_Personnel_Compensation_Period.PayrollDate) = " & iYear1 & ")), 0) AS Month1, " & _
    " ISNULL((SELECT SUM(tbl_Personnel_Payroll_Earnings_1.TotalAmount) AS Amount FROM  dbo.tbl_Personnel_Payroll_Earnings AS tbl_Personnel_Payroll_Earnings_1 LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Payroll AS tbl_Personnel_Payroll_2 ON tbl_Personnel_Payroll_Earnings_1.MasterKey = tbl_Personnel_Payroll_2.PK LEFT OUTER JOIN dbo.tbl_Personnel_Payroll_Earnings_Table AS tbl_Personnel_Payroll_Earnings_Table_1 ON tbl_Personnel_Payroll_Earnings_1.EarningKey = tbl_Personnel_Payroll_Earnings_Table_1.PK LEFT OUTER JOIN dbo.tbl_Personnel_Compensation_Period AS tbl_Personnel_Compensation_Period_2 ON tbl_Personnel_Payroll_2.PayrollPeriodKey = tbl_Personnel_Compensation_Period_2.PK " & _
    " WHERE (tbl_Personnel_Payroll_Earnings_Table_1.Month13 = 1) " & _
    " AND (tbl_Personnel_Payroll_2.EmployeeKey = tbl_Personnel_Payroll_1.EmployeeKey) " & _
    " AND (MONTH(tbl_Personnel_Compensation_Period_2.PayrollDate) = " & iMonth2 & ") " & _
    " AND (YEAR(tbl_Personnel_Compensation_Period_2.PayrollDate) = " & iYear2 & ")), 0) AS Month2, " & _
    " ISNULL((SELECT SUM(tbl_Personnel_Payroll_Earnings_1.TotalAmount) AS Amount FROM  dbo.tbl_Personnel_Payroll_Earnings AS tbl_Personnel_Payroll_Earnings_1 LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Payroll AS tbl_Personnel_Payroll_2 ON tbl_Personnel_Payroll_Earnings_1.MasterKey = tbl_Personnel_Payroll_2.PK LEFT OUTER JOIN dbo.tbl_Personnel_Payroll_Earnings_Table AS tbl_Personnel_Payroll_Earnings_Table_1 ON tbl_Personnel_Payroll_Earnings_1.EarningKey = tbl_Personnel_Payroll_Earnings_Table_1.PK LEFT OUTER JOIN dbo.tbl_Personnel_Compensation_Period AS tbl_Personnel_Compensation_Period_2 ON tbl_Personnel_Payroll_2.PayrollPeriodKey = tbl_Personnel_Compensation_Period_2.PK " & _
    " WHERE (tbl_Personnel_Payroll_Earnings_Table_1.Month13 = 1) " & _
    " AND (tbl_Personnel_Payroll_2.EmployeeKey = tbl_Personnel_Payroll_1.EmployeeKey) " & _
    " AND (MONTH(tbl_Personnel_Compensation_Period_2.PayrollDate) = " & iMonth3 & ") " & _
    " AND (YEAR(tbl_Personnel_Compensation_Period_2.PayrollDate) = " & iYear3 & ")), 0) AS Month3 " & _
    " FROM  dbo.tbl_Personnel_Payroll AS tbl_Personnel_Payroll_1 LEFT OUTER JOIN dbo.tbl_Personnel_IDNumber ON tbl_Personnel_Payroll_1.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK LEFT OUTER JOIN dbo.tbl_Personnel_Compensation_Period AS tbl_Personnel_Compensation_Period_1 ON tbl_Personnel_Payroll_1.PayrollPeriodKey = tbl_Personnel_Compensation_Period_1.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_ActionNew ON tbl_Personnel_Payroll_1.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK LEFT OUTER JOIN dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
    " WHERE (MONTH(tbl_Personnel_Compensation_Period_1.PayrollDate) = " & iMonth1 & ") AND (YEAR(tbl_Personnel_Compensation_Period_1.PayrollDate) = " & iYear1 & ") AND (dbo.tbl_Personnel_ActionNew.DivisionKey = " & DivKey & ") AND (dbo.tbl_Personnel_Position.PositionLevel = " & PostLevel & ") " & _
    " OR (MONTH(tbl_Personnel_Compensation_Period_1.PayrollDate) = " & iMonth2 & ") AND (YEAR(tbl_Personnel_Compensation_Period_1.PayrollDate) = " & iYear2 & ") AND (dbo.tbl_Personnel_ActionNew.DivisionKey = " & DivKey & ") AND (dbo.tbl_Personnel_Position.PositionLevel = " & PostLevel & ") " & _
    " OR (MONTH(tbl_Personnel_Compensation_Period_1.PayrollDate) = " & iMonth3 & ") AND (YEAR(tbl_Personnel_Compensation_Period_1.PayrollDate) = " & iYear3 & ") AND (dbo.tbl_Personnel_ActionNew.DivisionKey = " & DivKey & ") AND (dbo.tbl_Personnel_Position.PositionLevel = " & PostLevel & ")" & _
    " GROUP BY tbl_Personnel_Payroll_1.EmployeeKey, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName + ',  ' + dbo.tbl_Personnel_Information.FirstName + '  ' + dbo.tbl_Personnel_Information.MiddleName"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_13thMonth " & _
                      " (LogInName, CompanyKey, DivisionName, MonthYear, PostLevel) " & _
                      " VALUES ('" & sUser & "', 1, '" & FORMATSQL(CStr(DivName)) & "', '" & UCase(Format(PayrollDate, "mmmm yyyy")) & "', " & _
                      " '" & IIf(PostLevel = 1, "Rank in File", "Supervisory") & "')"
    iPK = 0
    t = "SELECT PK " & _
        " FROM tbl_Personnel_Payroll_13thMonth " & _
        " WHERE (LogInName = '" & sUser & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        iPK = rt!PK
    End If
    rt.Close
    
    iRec = 0
    If CDbl(iPK) <> 0 Then
        While Not ra.EOF
            DoEvents
            iRec = iRec + 1
            t = "SELECT TOP (1) dbo.tbl_Personnel_Department.DepartmentName " & _
                " FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK " & _
                " WHERE (dbo.tbl_Personnel_ActionNew.EmpPK = " & ra!EmployeeKey & ") " & _
                " AND (dbo.tbl_Personnel_ActionNew.EffectivityDate <= '" & DateSerial(iYear3, iMonth3 + 1, 0) & "') " & _
                " ORDER BY dbo.tbl_Personnel_ActionNew.EffectivityDate DESC"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                sDeptName = rt!DepartmentName
            End If
            rt.Close
            
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_13thMonth_Det " & _
                              " (MasterKey, EmployeeKey, Department, IDNumber, EmployeeName, Basic1, Basic2, Basic3) " & _
                              " VALUES (" & iPK & ", " & ra!EmployeeKey & ", '" & FORMATSQL(CStr(sDeptName)) & "', " & _
                              " '" & ra!IDNumber & "', '" & FORMATSQL(ra!EmployeeName) & "', " & CDbl(ra!Month1) & ", " & _
                              " " & CDbl(ra!Month2) & ", " & CDbl(ra!Month3) & ")"
            UpdateProgress_No_Percent MainForm.picProgressBar, iRec / ra.RecordCount
            ra.MoveNext
        Wend
    End If
End If
ra.Close
MainForm.picProgressBar.BackColor = &H8000000F
DoEvents
End Sub

Public Sub GenerateLoanContri(sUser, iDivKey, PayrollKey, LoanContri, sDivName, iPostLevel, iMonth, iYear)
ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll_ContriLoanRep WHERE (LogInName = '" & sUser & "')"
'&H8000000F&
MainForm.picProgressBar.BackColor = &H8000000F
DoEvents
Select Case LoanContri
    Case 5  'loans
        a = "SELECT dbo.tbl_Personnel_Payroll.EmployeeKey, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName + ',  ' + dbo.tbl_Personnel_Information.FirstName + '  ' + dbo.tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
            " dbo.tbl_Personnel_Department.DepartmentName, dbo.tbl_Personnel_Information.SSSNumber, dbo.tbl_Personnel_Information.PHICNumber , dbo.tbl_Personnel_Information.HDMFNumber, dbo.tbl_Personnel_Information.TIN, dbo.tbl_Personnel_Payroll.PayrollPeriodKey, " & _
            " dbo.tbl_Personnel_ActionNew.DivisionKey, dbo.tbl_Personnel_Payroll_Deductions.MasterKey " & _
            " FROM  dbo.tbl_Personnel_Payroll_Deductions LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Payroll_Deductions.DeductionKey = dbo.tbl_Personnel_Payroll_Deductions_Table.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Deductions.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Payroll.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
            " Where (dbo.tbl_Personnel_Payroll.PayrollPeriodKey = " & PayrollKey & ") " & _
            " And (dbo.tbl_Personnel_ActionNew.DivisionKey = " & iDivKey & ") " & _
            " And (dbo.tbl_Personnel_Payroll_Deductions_Table.GovtDed = 1) " & _
            " And (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
            " GROUP BY dbo.tbl_Personnel_Payroll.EmployeeKey, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName + ',  ' + dbo.tbl_Personnel_Information.FirstName + '  ' + dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Department.DepartmentName, " & _
            " dbo.tbl_Personnel_Information.SSSNumber, dbo.tbl_Personnel_Information.PHICNumber, dbo.tbl_Personnel_Information.HDMFNumber , dbo.tbl_Personnel_Information.TIN, dbo.tbl_Personnel_Payroll.PayrollPeriodKey, dbo.tbl_Personnel_ActionNew.DivisionKey, dbo.tbl_Personnel_Payroll_Deductions.MasterKey " & _
            " HAVING (SUM(dbo.tbl_Personnel_Payroll_Deductions.Amount) <> 0)"
    Case 6  'contri
        a = "SELECT dbo.tbl_Personnel_Payroll.EmployeeKey, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName + ',  ' + dbo.tbl_Personnel_Information.FirstName + '  ' + dbo.tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
            " dbo.tbl_Personnel_Department.DepartmentName, dbo.tbl_Personnel_Information.SSSNumber, dbo.tbl_Personnel_Information.PHICNumber , dbo.tbl_Personnel_Information.HDMFNumber, dbo.tbl_Personnel_Information.TIN, dbo.tbl_Personnel_Payroll.PayrollPeriodKey, " & _
            " dbo.tbl_Personnel_ActionNew.DivisionKey, dbo.tbl_Personnel_Payroll_Deductions.MasterKey " & _
            " FROM  dbo.tbl_Personnel_Payroll_Deductions LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Payroll_Deductions.DeductionKey = dbo.tbl_Personnel_Payroll_Deductions_Table.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Deductions.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Payroll.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
            " Where (dbo.tbl_Personnel_Payroll.PayrollPeriodKey = " & PayrollKey & ") " & _
            " And (dbo.tbl_Personnel_ActionNew.DivisionKey = " & iDivKey & ") " & _
            " And (dbo.tbl_Personnel_Payroll_Deductions_Table.GovtDed = 2) " & _
            " And (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
            " GROUP BY dbo.tbl_Personnel_Payroll.EmployeeKey, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName + ',  ' + dbo.tbl_Personnel_Information.FirstName + '  ' + dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Department.DepartmentName, " & _
            " dbo.tbl_Personnel_Information.SSSNumber, dbo.tbl_Personnel_Information.PHICNumber, dbo.tbl_Personnel_Information.HDMFNumber , dbo.tbl_Personnel_Information.TIN, dbo.tbl_Personnel_Payroll.PayrollPeriodKey, dbo.tbl_Personnel_ActionNew.DivisionKey, dbo.tbl_Personnel_Payroll_Deductions.MasterKey " & _
            " HAVING (SUM(dbo.tbl_Personnel_Payroll_Deductions.Amount) <> 0)"
End Select
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_ContriLoanRep " & _
                      " (LogInName, CompanyKey, PayrollPeriodKey, DivisionName, PostLevel) " & _
                      " VALUES ('" & sUser & "', 1, " & PayrollKey & ", '" & FORMATSQL(CStr(sDivName)) & "', '" & IIf(iPostLevel = 1, "Rank in File", "Supervisory") & "')"
    iPK = 0
    t = "SELECT PK " & _
        " FROM tbl_Personnel_Payroll_ContriLoanRep " & _
        " WHERE (LogInName = '" & sUser & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        iPK = rt!PK
    End If
    rt.Close
    iRec = 0
    If CDbl(iPK) <> 0 Then
        While Not ra.EOF
            DoEvents
            iRec = iRec + 1
            
            dblGrossIncome = 0
            If LoanContri = 6 Then
                t = "SELECT ROUND(SUM(dbo.tbl_Personnel_Payroll_Earnings.Taxable), 2) AS Gross " & _
                    " FROM  dbo.tbl_Personnel_Payroll_Earnings LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
                    " WHERE (dbo.tbl_Personnel_Payroll.EmployeeKey = " & ra!EmployeeKey & ") " & _
                    " AND (MONTH(dbo.tbl_Personnel_Compensation_Period.PayrollDate) = " & iMonth & ") " & _
                    " AND (YEAR(dbo.tbl_Personnel_Compensation_Period.PayrollDate) = " & iYear & ")"
                If rt.State = adStateOpen Then rt.Close
                rt.Open t, ConnOmega
                If rt.RecordCount > 0 Then
                    dblGrossIncome = IIf(IsNull(rt!Gross), 0, rt!Gross)
                End If
                rt.Close
            End If
            
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_ContriLoanRep_Det " & _
                              " (MasterKey, EmployeeKey, EmployeeName, EmployeeID, Department, SSSNum, PHICNum, PagIbigNum, TIN, GrossIncome) " & _
                              " VALUES (" & iPK & ", " & ra!EmployeeKey & ", '" & FORMATSQL(CStr(ra!EmployeeName)) & "', '" & ra!IDNumber & "', " & _
                              " '" & FORMATSQL(CStr(ra!DepartmentName)) & "', '" & Replace(FORMATSQL(CStr(ra!SSSNumber)), "-", "") & "', " & _
                              " '" & Replace(FORMATSQL(CStr(ra!PHICNumber)), "-", "") & "', '" & Replace(FORMATSQL(CStr(ra!HDMFNumber)), "-", "") & "', " & _
                              " '" & Replace(FORMATSQL(CStr(ra!TIN)), "-", "") & "', " & CDbl(dblGrossIncome) & ")"
                              
            dblSSSLoan = 0: dblSSSLoanBal = 0: dblPagIbigLoan = 0: dblPagIbigLoanBal = 0
            dblSSS = 0: dblSSSEmpr = 0: dblSSSEC = 0: dblPHIC = 0: dblPHICEmpr = 0: dblPagIbig = 0: dblPagIbigEmpr = 0: dblWHT = 0
            
            Select Case LoanContri
                Case 5  'loans
                    
                    t = "SELECT PK, Description, GovtDedEmpr " & _
                        " From dbo.tbl_Personnel_Payroll_Deductions_Table " & _
                        " Where (GovtDed = 1) " & _
                        " ORDER BY Sorting"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        u = "SELECT Amount " & _
                            " From dbo.tbl_Personnel_Payroll_Deductions " & _
                            " WHERE (MasterKey = " & ra!MasterKey & ") " & _
                            " AND (DeductionKey = " & rt!PK & ")"
                        If ru.State = adStateOpen Then ru.Close
                        ru.Open u, ConnOmega
                        If ru.RecordCount > 0 Then
                            Select Case rt!PK
                                Case 9  'SSS Loans
                                    dblSSSLoan = ru!Amount
                                    
                                Case 11 'PagIbig Loans
                                    dblPagIbigLoan = ru!Amount
                            End Select
                        End If
                        ru.Close
                        'Balance
                        u = "SELECT LoanKey, TransactionDate " & _
                            " From dbo.tbl_Personnel_Loans_SL " & _
                            " WHERE (PayrollKey = " & ra!MasterKey & ") " & _
                            " AND (LoanType = " & rt!PK & ")"
                        If ru.State = adStateOpen Then ru.Close
                        ru.Open u, ConnOmega
                        If ru.RecordCount > 0 Then
                            v = "SELECT SUM(Balance) AS Amount " & _
                                " From dbo.tbl_Personnel_Loans_SL " & _
                                " WHERE (LoanKey = " & ru!LoanKey & ") " & _
                                " AND (TransactionDate <= '" & FormatDateTime(ru!TransactionDate, vbShortDate) & "')"
                            If rv.State = adStateOpen Then rv.Close
                            rv.Open v, ConnOmega
                            If rv.RecordCount > 0 Then
                                Select Case rt!PK
                                    Case 9  'SSS Loans
                                        dblSSSLoanBal = IIf(IsNull(rv!Amount), 0, rv!Amount)
                                    Case 11 'PagIbig Loans
                                        dblPagIbigLoanBal = IIf(IsNull(rv!Amount), 0, rv!Amount)
                                End Select
                            End If
                            rv.Close
                        End If
                        ru.Close
                        
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ConnOmega.Execute "UPDATE tbl_Personnel_Payroll_ContriLoanRep_Det " & _
                                      " SET SSSLoan = " & CDbl(dblSSSLoan) & ", SSSLoanBal = " & CDbl(dblSSSLoanBal) & ", " & _
                                      " PagIbigLoan = " & CDbl(dblPagIbigLoan) & ", PagIbigLoanBal = " & CDbl(dblPagIbigLoanBal) & " " & _
                                      " WHERE (MasterKey = " & iPK & ") " & _
                                      " AND (EmployeeKey = " & ra!EmployeeKey & ")"
                Case 6  'contri
                    
                    t = "SELECT PK, Description, GovtDedEmpr " & _
                        " From dbo.tbl_Personnel_Payroll_Deductions_Table " & _
                        " Where (GovtDed = 2) " & _
                        " ORDER BY Sorting"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        If CDbl(rt!GovtDedEmpr) = 1 Then
                            u = "SELECT Amount " & _
                                " From dbo.tbl_Personnel_Payroll_EmployerShare " & _
                                " WHERE (MasterKey = " & ra!MasterKey & ") " & _
                                " AND (DeductionKey = " & rt!PK & ")"
                        Else
                            u = "SELECT Amount " & _
                                " From dbo.tbl_Personnel_Payroll_Deductions " & _
                                " WHERE (MasterKey = " & ra!MasterKey & ") " & _
                                " AND (DeductionKey = " & rt!PK & ")"
                        End If
                        If ru.State = adStateOpen Then ru.Close
                        ru.Open u, ConnOmega
                        If ru.RecordCount > 0 Then
                            Select Case rt!PK
                                Case 1  'SSS
                                    dblSSS = ru!Amount
                                Case 2  'SSSEmpr
                                    dblSSSEmpr = ru!Amount
                                Case 3  'SSSEC
                                    dblSSSEC = ru!Amount
                                Case 4  'PHIC
                                    dblPHIC = ru!Amount
                                Case 5  'PHICEmpr
                                    dblPHICEmpr = ru!Amount
                                Case 6  'PagIbig
                                    dblPagIbig = ru!Amount
                                Case 7  'PagIbigEmpr
                                    dblPagIbigEmpr = ru!Amount
                                Case 8  'WHT
                                    dblWHT = ru!Amount
                            End Select
                        End If
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ConnOmega.Execute "UPDATE tbl_Personnel_Payroll_ContriLoanRep_Det " & _
                                      " SET SSS = " & CDbl(dblSSS) & ", SSSEmp = " & CDbl(dblSSSEmpr) & ", " & _
                                      " SSSEC = " & CDbl(dblSSSEC) & ", PHIC = " & CDbl(dblPHIC) & ", " & _
                                      " PHICEmp = " & CDbl(dblPHICEmpr) & ", PagIbig = " & CDbl(dblPagIbig) & ", " & _
                                      " PagIbigEmp = " & CDbl(dblPagIbigEmpr) & ", WHT = " & CDbl(dblWHT) & " " & _
                                      " WHERE (MasterKey = " & iPK & ") " & _
                                      " AND (EmployeeKey = " & ra!EmployeeKey & ")"
                    
            End Select
            
            UpdateProgress_No_Percent MainForm.picProgressBar, iRec / ra.RecordCount
            ra.MoveNext
        Wend
    End If
End If
ra.Close
MainForm.picProgressBar.BackColor = &H8000000F
DoEvents
End Sub

Public Sub GenerateLedger(sUser, iGroup, iKey, PayrollDate, iEarnDed, iDivKey, iPostLevel)
MainForm.picProgressBar.BackColor = &H8000000F
DoEvents
a = ""
Select Case iGroup
    Case 2  'Department
        a = "SELECT COUNT(*) AS RecCount " & _
            " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
            " WHERE (dbo.tbl_Personnel_ActionNew.DeptKey = " & iKey & ") " & _
            " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
            " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
    Case 3  'Division
        a = "SELECT COUNT(*) AS RecCount " & _
            " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
            " WHERE (dbo.tbl_Personnel_ActionNew.DivisionKey = " & iKey & ") " & _
            " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
            " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
End Select
If a = "" Then Exit Sub
If rs.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    If CDbl(ra!RecCount) > 0 Then
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll_Report_Ledger WHERE (LogInName = '" & sUser & "')"
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger " & _
                          " (LogInName, CompanyKey, PayrollPeriod, GroupType, PostLevelDesc) " & _
                          " VALUES ('" & sUser & "', 1, '" & FormatDateTime(PayrollDate, vbShortDate) & "', " & iGroup & ", " & _
                          " '" & IIf(iPostLevel = 1, "Rank in File", "Supervisory") & "')"
        iPK = 0
        t = "SELECT PK " & _
            " FROM tbl_Personnel_Payroll_Report_Ledger " & _
            " WHERE (LogInName = '" & sUser & "')"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            iPK = rt!PK
        End If
        rt.Close
        iTerms = 0
        ' Header
        'iColMax = 10: iColCnt = 0
        'iColMax = 9
        If CDbl(iEarnDed) = 1 Then
            iColMax = 9
        Else
            iColMax = 10
        End If
        iColCnt = 0
        sAccName = "": sAccKey = ""
        
        t = "SELECT tbl_Personnel_Compensation_Period.* " & _
            " FROM tbl_Personnel_Compensation_Period " & _
            " WHERE (PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            iTerms = rt!Terms
        End If
        rt.Close
        If CDbl(iEarnDed) = 1 Then  'Earning
            t = "SELECT dbo.tbl_Personnel_Payroll_Earnings_Table.Abbvt as Description, dbo.tbl_Personnel_Payroll_Earnings.EarningKey " & _
                " FROM  dbo.tbl_Personnel_Payroll_Earnings LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Payroll_Earnings_Table ON dbo.tbl_Personnel_Payroll_Earnings.EarningKey = dbo.tbl_Personnel_Payroll_Earnings_Table.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
                " WHERE (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "') " & _
                " AND (dbo.tbl_Personnel_ActionNew.DivisionKey = " & iDivKey & ") " & _
                " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
                " GROUP BY dbo.tbl_Personnel_Payroll_Earnings_Table.Abbvt, dbo.tbl_Personnel_Payroll_Earnings_Table.Sorting, dbo.tbl_Personnel_Payroll_Earnings.EarningKey " & _
                " ORDER BY dbo.tbl_Personnel_Payroll_Earnings_Table.Sorting"
        Else    'Deductions
            
            If iTerms = 1 Then
                iColMax = iColMax - 2
            ElseIf iTerms = 2 Then
                iColMax = iColMax - 4
            End If
        
            t = "SELECT dbo.tbl_Personnel_Payroll_Deductions_Table.Abbvt as Description, dbo.tbl_Personnel_Payroll_Deductions.DeductionKey as EarningKey " & _
                " FROM  dbo.tbl_Personnel_Payroll_Deductions LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Payroll_Deductions.DeductionKey = dbo.tbl_Personnel_Payroll_Deductions_Table.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Deductions.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
                " WHERE (dbo.tbl_Personnel_ActionNew.DivisionKey = " & iDivKey & ") " & _
                " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "') " & _
                " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
                " AND (dbo.tbl_Personnel_Payroll_Deductions_Table.DedSched = 0) " & _
                " GROUP BY dbo.tbl_Personnel_Payroll_Deductions_Table.Abbvt, dbo.tbl_Personnel_Payroll_Deductions.DeductionKey, dbo.tbl_Personnel_Payroll_Deductions_Table.Sorting " & _
                " ORDER BY dbo.tbl_Personnel_Payroll_Deductions_Table.Sorting"
            't = "SELECT dbo.tbl_Personnel_Payroll_Deductions_Table.Abbvt as Description, dbo.tbl_Personnel_Payroll_Deductions.DeductionKey as EarningKey " & _
                " FROM  dbo.tbl_Personnel_Payroll_Deductions LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Payroll_Deductions.DeductionKey = dbo.tbl_Personnel_Payroll_Deductions_Table.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Deductions.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
                " WHERE (dbo.tbl_Personnel_ActionNew.DivisionKey = " & iDivKey & ") " & _
                " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "') " & _
                " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
                " GROUP BY dbo.tbl_Personnel_Payroll_Deductions_Table.Abbvt, dbo.tbl_Personnel_Payroll_Deductions.DeductionKey, dbo.tbl_Personnel_Payroll_Deductions_Table.Sorting " & _
                " ORDER BY dbo.tbl_Personnel_Payroll_Deductions_Table.Sorting"
        End If
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        While Not rt.EOF
            If CDbl(iColCnt) = CDbl(iColMax) Then
                iColCnt = 0
                sAccKey = sAccKey & "|" & rt!EarningKey
                sAccName = sAccName & "|" & rt!Description
            Else
                sAccKey = sAccKey & "{" & rt!EarningKey
                sAccName = sAccName & "{" & rt!Description
            End If
            iColCnt = iColCnt + 1
            rt.MoveNext
        Wend
        rt.Close
        
        sAccKey = Mid(sAccKey, 2, Len(sAccKey))
        sAccName = Mid(sAccName, 2, Len(sAccName))
                
        ArrAccName = Split(sAccName, "|", -1, 1)
        
        iRowCnt = 0
        If UBound(ArrAccName) = -1 Then
            iColCnt = 0
            ArrAccName1 = Split(ArrAccName, "{", -1, 1)
            iRowCnt = iRowCnt + 1
            sQry = "" & iPK & ", " & iRowCnt & ""
            For i = 0 To UBound(ArrAccName1)
                iColCnt = iColCnt + 1
                sQry = sQry & ",'" & ArrAccName1(i) & "'"
            Next i
            
            If CDbl(iColCnt) < CDbl(iColMax) Then
                For k = 1 To CDbl(iColMax) - CDbl(iColCnt)
                    sQry = sQry & ",''"
                Next k
            End If
            
            If CDbl(iEarnDed) = 1 Then 'earnings
                ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_Header " & _
                                  " (MasterKey, Line, Header1, Header2, Header3, Header4, Header5, Header6, Header7, Header8, Header9) " & _
                                  " VALUES (" & sQry & ")"
            Else
                If CDbl(iTerms) = 1 Then
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_Header " & _
                                      " (MasterKey, Line, Header1, Header2, Header3, Header4, Header5, Header6, Header7, Header8) " & _
                                      " VALUES (" & sQry & ")"
                ElseIf CDbl(iTerms) = 2 Then
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_Header " & _
                                      " (MasterKey, Line, Header1, Header2, Header3, Header4, Header5, Header6) " & _
                                      " VALUES (" & sQry & ")"
                End If
            End If
            
        Else
            For l = 0 To UBound(ArrAccName)
                iColCnt = 0
                ArrAccName1 = Split(ArrAccName(l), "{", -1, 1)
                iRowCnt = iRowCnt + 1
                sQry = "" & iPK & ", " & iRowCnt & ""
                For i = 0 To UBound(ArrAccName1)
                    iColCnt = iColCnt + 1
                    sQry = sQry & ",'" & ArrAccName1(i) & "'"
                Next i
                
                If CDbl(iColCnt) < CDbl(iColMax) Then
                    For k = 1 To CDbl(iColMax) - CDbl(iColCnt)
                        sQry = sQry & ",''"
                    Next k
                End If
                
                If CDbl(iEarnDed) = 1 Then 'earnings
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_Header " & _
                                      " (MasterKey, Line, Header1, Header2, Header3, Header4, Header5, Header6, Header7, Header8, Header9) " & _
                                      " VALUES (" & sQry & ")"
                Else
                    'Debug.Print sQry
                    
                    If CDbl(iTerms) = 1 Then
                        ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_Header " & _
                                          " (MasterKey, Line, Header1, Header2, Header3, Header4, Header5, Header6, Header7, Header8) " & _
                                          " VALUES (" & sQry & ")"
                    ElseIf CDbl(iTerms) = 2 Then
                        ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_Header " & _
                                          " (MasterKey, Line, Header1, Header2, Header3, Header4, Header5, Header6) " & _
                                          " VALUES (" & sQry & ")"
                    End If
                End If
                
            Next l
        End If
    End If
End If
ra.Close

If CDbl(iPK) <> 0 Then
    Select Case iGroup
        Case 2  'department
            a = "SELECT dbo.tbl_Personnel_Payroll.EmployeeKey, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName + ',  ' + dbo.tbl_Personnel_Information.FirstName + '  ' + dbo.tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
                " dbo.tbl_Personnel_Division.Description AS Division, dbo.tbl_Personnel_Department.DepartmentName AS Department, " & _
                " dbo.tbl_Personnel_Position.PositionName AS Position, dbo.tbl_Govt_TaxStatus.TaxStatus, dbo.tbl_Personnel_EmploymentStatus.StatusName AS EmploymentStatus, " & _
                " dbo.tbl_Personnel_CompensationRate.Description AS CompensationRate,  dbo.tbl_Personnel_Compensation_Period.PayrollDate, dbo.tbl_Personnel_Compensation_Period.DateFrom, " & _
                " dbo.tbl_Personnel_Compensation_Period.DateTo, dbo.tbl_Personnel_Payroll.ActionMemoKey, dbo.tbl_Personnel_Payroll.PK, dbo.tbl_Personnel_ActionNew.DeptKey, " & _
                " dbo.tbl_Personnel_ActionNew.CompensationRateKey " & _
                " FROM  dbo.tbl_Personnel_Information RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Payroll.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK ON dbo.tbl_Personnel_Information.PK = dbo.tbl_Personnel_IDNumber.ProfileKey LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_CompensationRate RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_CompensationRate.PK = dbo.tbl_Personnel_ActionNew.CompensationRateKey LEFT OUTER JOIN " & _
                " dbo.tbl_Govt_TaxStatus ON dbo.tbl_Personnel_ActionNew.TaxStatusKey = dbo.tbl_Govt_TaxStatus.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_EmploymentStatus ON dbo.tbl_Personnel_ActionNew.EmpStatusKey = dbo.tbl_Personnel_EmploymentStatus.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK FULL OUTER JOIN " & _
                " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_ActionNew.DivisionKey = dbo.tbl_Personnel_Division.PK " & _
                " WHERE (dbo.tbl_Personnel_ActionNew.DeptKey = " & iKey & ") " & _
                " AND ((dbo.tbl_Personnel_ActionNew.DivisionKey = " & iDivKey & ") " & _
                " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
                " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
        Case 3  'Division ' iDivKey
            a = "SELECT dbo.tbl_Personnel_Payroll.EmployeeKey, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName + ',  ' + dbo.tbl_Personnel_Information.FirstName + '  ' + dbo.tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
                " dbo.tbl_Personnel_Division.Description AS Division, dbo.tbl_Personnel_Department.DepartmentName AS Department, " & _
                " dbo.tbl_Personnel_Position.PositionName AS Position, dbo.tbl_Govt_TaxStatus.TaxStatus, dbo.tbl_Personnel_EmploymentStatus.StatusName AS EmploymentStatus, " & _
                " dbo.tbl_Personnel_CompensationRate.Description AS CompensationRate,  dbo.tbl_Personnel_Compensation_Period.PayrollDate, dbo.tbl_Personnel_Compensation_Period.DateFrom, " & _
                " dbo.tbl_Personnel_Compensation_Period.DateTo, dbo.tbl_Personnel_Payroll.ActionMemoKey, dbo.tbl_Personnel_Payroll.PK, dbo.tbl_Personnel_ActionNew.DeptKey, " & _
                " dbo.tbl_Personnel_ActionNew.CompensationRateKey " & _
                " FROM  dbo.tbl_Personnel_Information RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Payroll.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK ON dbo.tbl_Personnel_Information.PK = dbo.tbl_Personnel_IDNumber.ProfileKey LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_CompensationRate RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_CompensationRate.PK = dbo.tbl_Personnel_ActionNew.CompensationRateKey LEFT OUTER JOIN " & _
                " dbo.tbl_Govt_TaxStatus ON dbo.tbl_Personnel_ActionNew.TaxStatusKey = dbo.tbl_Govt_TaxStatus.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_EmploymentStatus ON dbo.tbl_Personnel_ActionNew.EmpStatusKey = dbo.tbl_Personnel_EmploymentStatus.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK FULL OUTER JOIN " & _
                " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_ActionNew.DivisionKey = dbo.tbl_Personnel_Division.PK " & _
                " WHERE (dbo.tbl_Personnel_ActionNew.DivisionKey = " & iKey & ") " & _
                " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
                " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
    End Select
    If ra.State = adStateOpen Then ra.Close
    ra.Open a, ConnOmega
    If ra.RecordCount > 0 Then
        ConnOmega.Execute "UPDATE tbl_Personnel_Payroll_Report_Ledger " & _
                          " SET PayrollPeriodFrom = '" & FormatDateTime(ra!DateFrom, vbShortDate) & "', " & _
                          " PayrollPeriodTo = '" & FormatDateTime(ra!DateTo, vbShortDate) & "', " & _
                          " PayrollRange = '" & Format(FormatDateTime(ra!DateFrom, vbShortDate), "mm/dd/yyyy") & " - " & Format(FormatDateTime(ra!DateTo, vbShortDate), "mm/dd/yyyy") & "' " & _
                          " WHERE (PK = " & iPK & ")"
        
        iRec = 0
        While Not ra.EOF
            DoEvents
            iRec = iRec + 1
            sRate = ""
            sCompRate = "RATE : " & ra!CompensationRate & " ("
            t = "SELECT dbo.tbl_Personnel_Payroll_Earnings_Table.Description, dbo.tbl_Personnel_ActionNew_Rate.Rate " & _
                " FROM  dbo.tbl_Personnel_ActionNew_Rate LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Payroll_Earnings_Table ON dbo.tbl_Personnel_ActionNew_Rate.EarningKey = dbo.tbl_Personnel_Payroll_Earnings_Table.PK " & _
                " WHERE (dbo.tbl_Personnel_ActionNew_Rate.MasterKey = " & ra!ActionMemoKey & ") " & _
                " ORDER BY dbo.tbl_Personnel_Payroll_Earnings_Table.Sorting"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                While Not rt.EOF
                    sCompRate = sCompRate & rt!Description & " = " & Format(rt!Rate, "#,##0.00") & " | "
                    sRate = sRate & IIf(ra!CompensationRateKey = 3, Format(CDbl(rt!Rate) / 2, "#,##0.00"), Format(rt!Rate, "#,##0.00"))
                    rt.MoveNext
                Wend
            End If
            rt.Close
            sCompRate = Mid(sCompRate, 1, Len(sCompRate) - 3) & ")"
            
            dGross = 0: dDed = 0
            dSSSLoan = 0: dPGBGLoan = 0: dSSS = 0: dPHIC = 0: dPGBG = 0: dWHT = 0
            
            If CDbl(iEarnDed) <> 1 Then
                t = "SELECT dbo.tbl_Personnel_Payroll_Deductions_Table.* " & _
                    " From dbo.tbl_Personnel_Payroll_Deductions_Table " & _
                    " WHERE (FixDed = 1) AND (DedSched = " & iTerms & ")"
                If rt.State = adStateOpen Then rt.Close
                rt.Open t, ConnOmega
                If rt.RecordCount > 0 Then
                    While Not rt.EOF
                        u = "SELECT dbo.tbl_Personnel_Payroll_Deductions.Amount " & _
                            " FROM  dbo.tbl_Personnel_Payroll_Deductions LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Deductions.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Payroll_Deductions.DeductionKey = dbo.tbl_Personnel_Payroll_Deductions_Table.PK LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
                            " WHERE (dbo.tbl_Personnel_Payroll.EmployeeKey = " & ra!EmployeeKey & ") " & _
                            " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "') " & _
                            " AND (dbo.tbl_Personnel_Payroll_Deductions.DeductionKey = " & rt!PK & ")"
                        If ru.State = adStateOpen Then ru.Close
                        ru.Open u, ConnOmega
                        If ru.RecordCount > 0 Then
                            Select Case rt!PK
                                Case 1: dSSS = ru!Amount        'sss
                                Case 4: dPHIC = ru!Amount       'phic
                                Case 6: dPGBG = ru!Amount       'pgbg
                                Case 8: dWHT = ru!Amount        'wht
                                Case 9: dSSSLoan = ru!Amount    'sss loan
                                Case 11: dPGBGLoan = ru!Amount  'pgbg loan
                            End Select
                        End If
                        ru.Close
                        rt.MoveNext
                    Wend
                    
                End If
                rt.Close
            End If
            
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det " & _
                              " (MasterKey, EmployeeKey, IDNumber, EmployeeName, Division, Department, Position, TaxStatus, " & _
                              " EmploymentStatus, CompensationRate, Gross, Deduction, DeptKey, SSSLoan, PGBGLoan, SSS, PHIC, PGBG, WHT, Rate) " & _
                              " VALUES (" & iPK & ", " & ra!EmployeeKey & ", '" & ra!IDNumber & "', '" & FORMATSQL(ra!EmployeeName) & "', " & _
                              " '" & FORMATSQL(ra!Division) & "', '" & FORMATSQL(ra!Department) & "', '" & FORMATSQL(ra!Position) & "', " & _
                              " '" & FORMATSQL(ra!TaxStatus) & "', '" & FORMATSQL(ra!EmploymentStatus) & "', '" & CStr(sCompRate) & "', " & _
                              " " & CDbl(dGross) & ", " & CDbl(dDed) & ", " & ra!DeptKey & ", " & CDbl(dSSSLoan) & ", " & CDbl(dPGBGLoan) & ", " & _
                              " " & CDbl(dSSS) & ", " & CDbl(dPHIC) & ", " & CDbl(dPGBG) & ", " & CDbl(dWHT) & ", '" & sRate & "')"
            
            If CDbl(iEarnDed) <> 1 Then
                dDed = dDed + CDbl(dSSSLoan) + CDbl(dPGBGLoan) + CDbl(dSSS) + CDbl(dPHIC) + CDbl(dPGBG) + CDbl(dWHT)
            End If
            
            ArrAccKey = Split(sAccKey, "|", -1, 1)
                
            iRowCnt = 0
            If UBound(ArrAccKey) = -1 Then
                iColCnt = 0
                ArrAccKey1 = Split(ArrAccKey, "{", -1, 1)
                iRowCnt = iRowCnt + 1
                sQry = "" & iPK & ", " & ra!EmployeeKey & ", " & iRowCnt & ""
                For i = 0 To UBound(ArrAccKey1)
                    iColCnt = iColCnt + 1
                    If CDbl(iEarnDed) = 1 Then  'Earning
                        t = "SELECT SUM(dbo.tbl_Personnel_Payroll_Earnings.TotalAmount) AS Amount " & _
                            " FROM  dbo.tbl_Personnel_Payroll_Earnings LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
                            " WHERE (dbo.tbl_Personnel_Payroll_Earnings.EarningKey = " & ArrAccKey1(i) & ") " & _
                            " AND (dbo.tbl_Personnel_Payroll.EmployeeKey = " & ra!EmployeeKey & ") " & _
                            " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
                    Else
                        t = "SELECT SUM(dbo.tbl_Personnel_Payroll_Deductions.Amount) AS Amount " & _
                            " FROM  dbo.tbl_Personnel_Payroll_Deductions LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Deductions.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
                            " WHERE (dbo.tbl_Personnel_Payroll_Deductions.DeductionKey = " & ArrAccKey1(i) & ") " & _
                            " AND (dbo.tbl_Personnel_Payroll.EmployeeKey = " & ra!EmployeeKey & ") " & _
                            " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
                    End If
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    If rt.RecordCount > 0 Then
                        If CDbl(iEarnDed) = 1 Then  'Earning
                            dGross = dGross + CDbl(IIf(IsNull(rt!Amount), 0, Format(rt!Amount, "#,##0.00")))
                        Else
                            dDed = dDed + CDbl(IIf(IsNull(rt!Amount), 0, Format(rt!Amount, "#,##0.00")))
                        End If
                        sQry = sQry & ",'" & IIf(IsNull(rt!Amount), "", Format(rt!Amount, "#,##0.00")) & "'"
                    End If
                    rt.Close
                Next i
                
                If CDbl(iColCnt) < CDbl(iColMax) Then
                    For k = 1 To CDbl(iColMax) - CDbl(iColCnt)
                        sQry = sQry & ",''"
                    Next k
                End If
                
                If CDbl(iEarnDed) = 1 Then 'earnings
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_Amount " & _
                                      " (MasterKey, EmployeeKey, Line, Value1, Value2, Value3, Value4, Value5, Value6, Value7, Value8, Value9) " & _
                                      " VALUES (" & sQry & ")"
                Else
                    If CDbl(iTerms) = 1 Then
                        ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_Amount " & _
                                          " (MasterKey, EmployeeKey, Line, Value1, Value2, Value3, Value4, Value5, Value6, Value7, Value8) " & _
                                          " VALUES (" & sQry & ")"
                    ElseIf CDbl(iTerms) = 2 Then
                        ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_Amount " & _
                                          " (MasterKey, EmployeeKey, Line, Value1, Value2, Value3, Value4, Value5, Value6) " & _
                                          " VALUES (" & sQry & ")"
                    End If
                End If
            Else
                For l = 0 To UBound(ArrAccKey)
                    iColCnt = 0
                    ArrAccKey1 = Split(ArrAccKey(l), "{", -1, 1)
                    iRowCnt = iRowCnt + 1
                    sQry = "" & iPK & ", " & ra!EmployeeKey & ", " & iRowCnt & ""
                    For i = 0 To UBound(ArrAccKey1)
                        iColCnt = iColCnt + 1
                        If CDbl(iEarnDed) = 1 Then  'Earning
                            t = "SELECT SUM(dbo.tbl_Personnel_Payroll_Earnings.TotalAmount) AS Amount " & _
                                " FROM  dbo.tbl_Personnel_Payroll_Earnings LEFT OUTER JOIN " & _
                                " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
                                " WHERE (dbo.tbl_Personnel_Payroll_Earnings.EarningKey = " & ArrAccKey1(i) & ") " & _
                                " AND (dbo.tbl_Personnel_Payroll.EmployeeKey = " & ra!EmployeeKey & ") " & _
                                " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
                        Else
                            t = "SELECT SUM(dbo.tbl_Personnel_Payroll_Deductions.Amount) AS Amount " & _
                                " FROM  dbo.tbl_Personnel_Payroll_Deductions LEFT OUTER JOIN " & _
                                " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Deductions.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
                                " WHERE (dbo.tbl_Personnel_Payroll_Deductions.DeductionKey = " & ArrAccKey1(i) & ") " & _
                                " AND (dbo.tbl_Personnel_Payroll.EmployeeKey = " & ra!EmployeeKey & ") " & _
                                " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
                        End If
                        If rt.State = adStateOpen Then rt.Close
                        rt.Open t, ConnOmega
                        If rt.RecordCount > 0 Then
                            If CDbl(iEarnDed) = 1 Then  'Earning
                                dGross = dGross + CDbl(IIf(IsNull(rt!Amount), 0, Format(rt!Amount, "#,##0.00")))
                            Else
                                dDed = dDed + CDbl(IIf(IsNull(rt!Amount), 0, Format(rt!Amount, "#,##0.00")))
                            End If
                            sQry = sQry & ",'" & IIf(IsNull(rt!Amount), "", Format(rt!Amount, "#,##0.00")) & "'"
                        End If
                        rt.Close
                    Next i
                    
                    If CDbl(iColCnt) < CDbl(iColMax) Then
                        For k = 1 To CDbl(iColMax) - CDbl(iColCnt)
                            sQry = sQry & ",''"
                        Next k
                    End If
                    
                    If CDbl(iEarnDed) = 1 Then 'earnings
                        ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_Amount " & _
                                          " (MasterKey, EmployeeKey, Line, Value1, Value2, Value3, Value4, Value5, Value6, Value7, Value8, Value9) " & _
                                          " VALUES (" & sQry & ")"
                    Else
                        If CDbl(iTerms) = 1 Then
                            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_Amount " & _
                                              " (MasterKey, EmployeeKey, Line, Value1, Value2, Value3, Value4, Value5, Value6, Value7, Value8) " & _
                                              " VALUES (" & sQry & ")"
                        ElseIf CDbl(iTerms) = 2 Then
                            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_Amount " & _
                                              " (MasterKey, EmployeeKey, Line, Value1, Value2, Value3, Value4, Value5, Value6) " & _
                                              " VALUES (" & sQry & ")"
                        End If
                    End If
                    
                Next l
            End If
            
            'ArrAccKeyHours = Split(sAccKey, "|", -1, 1)
            iRowCnt = 0
            If CDbl(iEarnDed) = 1 Then  'Earning
                If UBound(ArrAccKey) = -1 Then
                    'Hours
                    iColCnt = 0
                    ArrAccKey1 = Split(CStr(ArrAccKey), "{", -1, 1)
                    iRowCnt = iRowCnt + 1
                    sQry = "" & iPK & ", " & ra!EmployeeKey & ", " & iRowCnt & ""
                    For i = 0 To UBound(ArrAccKey1)
                        iColCnt = iColCnt + 1
                        t = "SELECT SUM(dbo.tbl_Personnel_Payroll_Earnings.Hours) AS Hours " & _
                            " FROM  dbo.tbl_Personnel_Payroll_Earnings LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
                            " WHERE (dbo.tbl_Personnel_Payroll_Earnings.EarningKey = " & ArrAccKey1(i) & ") " & _
                            " AND (dbo.tbl_Personnel_Payroll.EmployeeKey = " & ra!EmployeeKey & ") " & _
                            " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
                        If rt.State = adStateOpen Then rt.Close
                        rt.Open t, ConnOmega
                        If rt.RecordCount > 0 Then
                            If CDbl(ArrAccKey1(i)) <> 1 Then
                                sHours = IIf(IsNull(rt!Hours), 0, Format(rt!Hours, "#0.00"))
                                sHours = IIf(CDbl(sHours) = 0, "", sHours & " hrs")
                            Else
                                sHours = rt!Hours & " hrs"
                            End If
                            sQry = sQry & ",'" & sHours & "'"
                        End If
                        rt.Close
                    Next i
                    
                    If CDbl(iColCnt) < CDbl(iColMax) Then
                        For k = 1 To CDbl(iColMax) - CDbl(iColCnt)
                            sQry = sQry & ",''"
                        Next k
                    End If
                    
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_Hour " & _
                                      " (MasterKey, EmployeeKey, Line, Value1, Value2, Value3, Value4, Value5, Value6, Value7, Value8, Value9) " & _
                                      " VALUES (" & sQry & ")"
                Else
                    For l = 0 To UBound(ArrAccKey)
                        iColCnt = 0
                        ArrAccKey1 = Split(ArrAccKey(l), "{", -1, 1)
                        iRowCnt = iRowCnt + 1
                        sQry = "" & iPK & ", " & ra!EmployeeKey & ", " & iRowCnt & ""
                        For i = 0 To UBound(ArrAccKey1)
                            iColCnt = iColCnt + 1
                            t = "SELECT SUM(dbo.tbl_Personnel_Payroll_Earnings.Hours) AS Hours " & _
                                " FROM  dbo.tbl_Personnel_Payroll_Earnings LEFT OUTER JOIN " & _
                                " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
                                " WHERE (dbo.tbl_Personnel_Payroll_Earnings.EarningKey = " & ArrAccKey1(i) & ") " & _
                                " AND (dbo.tbl_Personnel_Payroll.EmployeeKey = " & ra!EmployeeKey & ") " & _
                                " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
                            If rt.State = adStateOpen Then rt.Close
                            rt.Open t, ConnOmega
                            If rt.RecordCount > 0 Then
                                If CDbl(ArrAccKey1(i)) <> 1 Then
                                    sHours = IIf(IsNull(rt!Hours), 0, Format(rt!Hours, "#0.00"))
                                    sHours = IIf(CDbl(sHours) = 0, "", sHours & " hrs")
                                Else
                                    sHours = rt!Hours & " hrs"
                                End If
                                sQry = sQry & ",'" & sHours & "'"
                            End If
                            rt.Close
                        Next i
                        
                        If CDbl(iColCnt) < CDbl(iColMax) Then
                            For k = 1 To CDbl(iColMax) - CDbl(iColCnt)
                                sQry = sQry & ",''"
                            Next k
                        End If
                        
                        ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_Hour " & _
                                          " (MasterKey, EmployeeKey, Line, Value1, Value2, Value3, Value4, Value5, Value6, Value7, Value8, Value9) " & _
                                          " VALUES (" & sQry & ")"
                        
                    Next l
                End If
            End If
            
            ConnOmega.Execute "UPDATE tbl_Personnel_Payroll_Report_Ledger_Det " & _
                              " SET Gross = " & CDbl(dGross) & ", " & _
                              " Deduction = " & CDbl(dDed) & " " & _
                              " WHERE (MasterKey = " & iPK & ") " & _
                              " AND (EmployeeKey = " & ra!EmployeeKey & ")"
            
            UpdateProgress_No_Percent MainForm.picProgressBar, iRec / ra.RecordCount
            ra.MoveNext
        Wend
    End If
    ra.Close
    
    'SubTotal
    Select Case iGroup
        Case 2 'department
            If CDbl(iEarnDed) = 1 Then  'Earning
                a = "SELECT dbo.tbl_Personnel_ActionNew.DeptKey " & _
                    " FROM  dbo.tbl_Personnel_Payroll_Earnings LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
                    " WHERE (dbo.tbl_Personnel_ActionNew.DivisionKey = " & iDivKey & ") " & _
                    " AND (dbo.tbl_Personnel_ActionNew.DeptKey = " & iKey & ") " & _
                    " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
                    " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "') " & _
                    " GROUP BY dbo.tbl_Personnel_ActionNew.DeptKey"
            Else
                a = "SELECT dbo.tbl_Personnel_ActionNew.DeptKey " & _
                    " FROM  dbo.tbl_Personnel_Payroll_Deductions LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Deductions.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
                    " WHERE (dbo.tbl_Personnel_ActionNew.DivisionKey = " & iDivKey & ") " & _
                    " AND (dbo.tbl_Personnel_ActionNew.DeptKey = " & iKey & ") " & _
                    " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
                    " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "') " & _
                    " GROUP BY dbo.tbl_Personnel_ActionNew.DeptKey"
            End If
        Case 3  'Division
            If CDbl(iEarnDed) = 1 Then  'Earning
                a = "SELECT dbo.tbl_Personnel_ActionNew.DeptKey " & _
                    " FROM  dbo.tbl_Personnel_Payroll_Earnings LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
                    " WHERE (dbo.tbl_Personnel_ActionNew.DivisionKey = " & iKey & ") " & _
                    " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
                    " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "') " & _
                    " GROUP BY dbo.tbl_Personnel_ActionNew.DeptKey"
            Else
                a = "SELECT dbo.tbl_Personnel_ActionNew.DeptKey " & _
                    " FROM  dbo.tbl_Personnel_Payroll_Deductions LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Deductions.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
                    " WHERE (dbo.tbl_Personnel_ActionNew.DivisionKey = " & iKey & ") " & _
                    " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
                    " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "') " & _
                    " GROUP BY dbo.tbl_Personnel_ActionNew.DeptKey"
            End If
    End Select
    If ra.State = adStateOpen Then ra.Close
    ra.Open a, ConnOmega
    While Not ra.EOF
        ArrAccKey = Split(sAccKey, "|", -1, 1)
        iRowCnt = 0
        If UBound(ArrAccKey) = -1 Then
            iColCnt = 0
            ArrAccKey1 = Split(ArrAccKey, "{", -1, 1)
            iRowCnt = iRowCnt + 1
            sQry = "" & iPK & ", " & ra!DeptKey & ", " & iRowCnt & ""
            For i = 0 To UBound(ArrAccKey1)
                iColCnt = iColCnt + 1
                If CDbl(iEarnDed) = 1 Then  'Earning
                    t = "SELECT SUM(dbo.tbl_Personnel_Payroll_Earnings.TotalAmount) AS Amount " & _
                        " FROM  dbo.tbl_Personnel_Payroll_Earnings LEFT OUTER JOIN " & _
                        " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                        " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                        " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
                        " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
                        " WHERE (dbo.tbl_Personnel_Payroll_Earnings.EarningKey = " & ArrAccKey1(i) & ") " & _
                        " AND (dbo.tbl_Personnel_ActionNew.DeptKey = " & ra!DeptKey & ") " & _
                        " AND (dbo.tbl_Personnel_ActionNew.DivisionKey = " & iDivKey & ") " & _
                        " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
                        " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
                Else
                    t = "SELECT SUM(dbo.tbl_Personnel_Payroll_Deductions.Amount) AS Amount " & _
                        " FROM  dbo.tbl_Personnel_Payroll_Deductions LEFT OUTER JOIN " & _
                        " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Deductions.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                        " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                        " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
                        " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
                        " WHERE (dbo.tbl_Personnel_Payroll_Deductions.DeductionKey = " & ArrAccKey1(i) & ") " & _
                        " AND (dbo.tbl_Personnel_ActionNew.DeptKey = " & ra!DeptKey & ") " & _
                        " AND (dbo.tbl_Personnel_ActionNew.DivisionKey = " & iDivKey & ") " & _
                        " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
                        " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
                End If
                If rt.State = adStateOpen Then rt.Close
                rt.Open t, ConnOmega
                If rt.RecordCount > 0 Then
                    sQry = sQry & ",'" & IIf(IsNull(rt!Amount), "0.00", Format(rt!Amount, "#,##0.00")) & "'"
                End If
                rt.Close
            Next i
            
            If CDbl(iColCnt) < CDbl(iColMax) Then
                For k = 1 To CDbl(iColMax) - CDbl(iColCnt)
                    sQry = sQry & ",''"
                Next k
            End If
            
            If CDbl(iEarnDed) = 1 Then 'earnings
                ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_SubTotal " & _
                                  " (MasterKey, DeptKey, Line, Value1, Value2, Value3, Value4, Value5, Value6, Value7, Value8, Value9) " & _
                                  " VALUES (" & sQry & ")"
            Else
                If CDbl(iTerms) = 1 Then
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_SubTotal " & _
                                      " (MasterKey, DeptKey, Line, Value1, Value2, Value3, Value4, Value5, Value6, Value7, Value8) " & _
                                      " VALUES (" & sQry & ")"
                ElseIf CDbl(iTerms) = 2 Then
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_SubTotal " & _
                                      " (MasterKey, DeptKey, Line, Value1, Value2, Value3, Value4, Value5, Value6) " & _
                                      " VALUES (" & sQry & ")"
                End If
            End If
        Else
            For l = 0 To UBound(ArrAccKey)
                iColCnt = 0
                ArrAccKey1 = Split(ArrAccKey(l), "{", -1, 1)
                iRowCnt = iRowCnt + 1
                sQry = "" & iPK & ", " & ra!DeptKey & ", " & iRowCnt & ""
                For i = 0 To UBound(ArrAccKey1)
                    iColCnt = iColCnt + 1
                    If CDbl(iEarnDed) = 1 Then  'Earning
                        t = "SELECT SUM(dbo.tbl_Personnel_Payroll_Earnings.TotalAmount) AS Amount " & _
                            " FROM  dbo.tbl_Personnel_Payroll_Earnings LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
                            " WHERE (dbo.tbl_Personnel_Payroll_Earnings.EarningKey = " & ArrAccKey1(i) & ") " & _
                            " AND (dbo.tbl_Personnel_ActionNew.DeptKey = " & ra!DeptKey & ") " & _
                            " AND (dbo.tbl_Personnel_ActionNew.DivisionKey = " & iDivKey & ") " & _
                            " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
                            " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
                    Else
                        t = "SELECT SUM(dbo.tbl_Personnel_Payroll_Deductions.Amount) AS Amount " & _
                            " FROM  dbo.tbl_Personnel_Payroll_Deductions LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Deductions.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
                            " WHERE (dbo.tbl_Personnel_Payroll_Deductions.DeductionKey = " & ArrAccKey1(i) & ") " & _
                            " AND (dbo.tbl_Personnel_ActionNew.DeptKey = " & ra!DeptKey & ") " & _
                            " AND (dbo.tbl_Personnel_ActionNew.DivisionKey = " & iDivKey & ") " & _
                            " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
                            " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
                    End If
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    If rt.RecordCount > 0 Then
                        sQry = sQry & ",'" & IIf(IsNull(rt!Amount), "0.00", Format(rt!Amount, "#,##0.00")) & "'"
                    End If
                    rt.Close
                Next i
                
                If CDbl(iColCnt) < CDbl(iColMax) Then
                    For k = 1 To CDbl(iColMax) - CDbl(iColCnt)
                        sQry = sQry & ",''"
                    Next k
                End If
                
                If CDbl(iEarnDed) = 1 Then 'earnings
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_SubTotal " & _
                                      " (MasterKey, DeptKey, Line, Value1, Value2, Value3, Value4, Value5, Value6, Value7, Value8, Value9) " & _
                                      " VALUES (" & sQry & ")"
                Else
                    If CDbl(iTerms) = 1 Then
                        ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_SubTotal " & _
                                          " (MasterKey, DeptKey, Line, Value1, Value2, Value3, Value4, Value5, Value6, Value7, Value8) " & _
                                          " VALUES (" & sQry & ")"
                    ElseIf CDbl(iTerms) = 2 Then
                        ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_SubTotal " & _
                                          " (MasterKey, DeptKey, Line, Value1, Value2, Value3, Value4, Value5, Value6) " & _
                                          " VALUES (" & sQry & ")"
                    End If
                End If
                
            Next l
        End If
        
        ra.MoveNext
    Wend
    ra.Close
    
    
    'Grand Total
    If CDbl(iGroup) = 3 Then
        ArrAccKey = Split(sAccKey, "|", -1, 1)
        iRowCnt = 0
        If UBound(ArrAccKey) = -1 Then
            iColCnt = 0
            ArrAccKey1 = Split(ArrAccKey, "{", -1, 1)
            iRowCnt = iRowCnt + 1
            sQry = "" & iPK & ", " & iRowCnt & ""
            For i = 0 To UBound(ArrAccKey1)
                iColCnt = iColCnt + 1
                If CDbl(iEarnDed) = 1 Then  'Earning
                    t = "SELECT SUM(dbo.tbl_Personnel_Payroll_Earnings.TotalAmount) AS Amount " & _
                        " FROM  dbo.tbl_Personnel_Payroll_Earnings LEFT OUTER JOIN " & _
                        " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                        " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                        " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
                        " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
                        " WHERE (dbo.tbl_Personnel_Payroll_Earnings.EarningKey = " & ArrAccKey1(i) & ") " & _
                        " AND (dbo.tbl_Personnel_ActionNew.DivisionKey = " & iDivKey & ") " & _
                        " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
                        " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
                Else
                    t = "SELECT SUM(dbo.tbl_Personnel_Payroll_Deductions.Amount) AS Amount " & _
                        " FROM  dbo.tbl_Personnel_Payroll_Deductions LEFT OUTER JOIN " & _
                        " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Deductions.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                        " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                        " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
                        " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
                        " WHERE (dbo.tbl_Personnel_Payroll_Deductions.DeductionKey = " & ArrAccKey1(i) & ") " & _
                        " AND (dbo.tbl_Personnel_ActionNew.DivisionKey = " & iDivKey & ") " & _
                        " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
                        " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
                End If
                If rt.State = adStateOpen Then rt.Close
                rt.Open t, ConnOmega
                If rt.RecordCount > 0 Then
                    sQry = sQry & ",'" & IIf(IsNull(rt!Amount), "0.00", Format(rt!Amount, "#,##0.00")) & "'"
                End If
                rt.Close
            Next i
            
            If CDbl(iColCnt) < CDbl(iColMax) Then
                For k = 1 To CDbl(iColMax) - CDbl(iColCnt)
                    sQry = sQry & ",''"
                Next k
            End If
            
            If CDbl(iEarnDed) = 1 Then 'earnings
                ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_GrandTotal " & _
                                  " (MasterKey, Line, Value1, Value2, Value3, Value4, Value5, Value6, Value7, Value8, Value9) " & _
                                  " VALUES (" & sQry & ")"
            Else
                If CDbl(iTerms) = 1 Then
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_GrandTotal " & _
                                      " (MasterKey, Line, Value1, Value2, Value3, Value4, Value5, Value6, Value7, Value8) " & _
                                      " VALUES (" & sQry & ")"
                ElseIf CDbl(iTerms) = 2 Then
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_GrandTotal " & _
                                      " (MasterKey, Line, Value1, Value2, Value3, Value4, Value5, Value6) " & _
                                      " VALUES (" & sQry & ")"
                End If
            End If
        Else
            For l = 0 To UBound(ArrAccKey)
                iColCnt = 0
                ArrAccKey1 = Split(ArrAccKey(l), "{", -1, 1)
                iRowCnt = iRowCnt + 1
                sQry = "" & iPK & ", " & iRowCnt & ""
                For i = 0 To UBound(ArrAccKey1)
                    iColCnt = iColCnt + 1
                    If CDbl(iEarnDed) = 1 Then  'Earning
                        t = "SELECT SUM(dbo.tbl_Personnel_Payroll_Earnings.TotalAmount) AS Amount " & _
                            " FROM  dbo.tbl_Personnel_Payroll_Earnings LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
                            " WHERE (dbo.tbl_Personnel_Payroll_Earnings.EarningKey = " & ArrAccKey1(i) & ") " & _
                            " AND (dbo.tbl_Personnel_ActionNew.DivisionKey = " & iDivKey & ") " & _
                            " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
                            " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
                    Else
                        t = "SELECT SUM(dbo.tbl_Personnel_Payroll_Deductions.Amount) AS Amount " & _
                            " FROM  dbo.tbl_Personnel_Payroll_Deductions LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Deductions.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
                            " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
                            " WHERE (dbo.tbl_Personnel_Payroll_Deductions.DeductionKey = " & ArrAccKey1(i) & ") " & _
                            " AND (dbo.tbl_Personnel_ActionNew.DivisionKey = " & iDivKey & ") " & _
                            " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
                            " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
                    End If
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    If rt.RecordCount > 0 Then
                        sQry = sQry & ",'" & IIf(IsNull(rt!Amount), "0.00", Format(rt!Amount, "#,##0.00")) & "'"
                    End If
                    rt.Close
                Next i
                
                If CDbl(iColCnt) < CDbl(iColMax) Then
                    For k = 1 To CDbl(iColMax) - CDbl(iColCnt)
                        sQry = sQry & ",''"
                    Next k
                End If
                
                If CDbl(iEarnDed) = 1 Then 'earnings
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_GrandTotal " & _
                                      " (MasterKey, Line, Value1, Value2, Value3, Value4, Value5, Value6, Value7, Value8, Value9) " & _
                                      " VALUES (" & sQry & ")"
                Else
                    If CDbl(iTerms) = 1 Then
                        ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_GrandTotal " & _
                                          " (MasterKey, Line, Value1, Value2, Value3, Value4, Value5, Value6, Value7, Value8) " & _
                                          " VALUES (" & sQry & ")"
                    ElseIf CDbl(iTerms) = 2 Then
                        ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_Ledger_Det_GrandTotal " & _
                                          " (MasterKey, Line, Value1, Value2, Value3, Value4, Value5, Value6) " & _
                                          " VALUES (" & sQry & ")"
                    End If
                End If
                
            Next l
        End If
    End If
End If
MainForm.picProgressBar.BackColor = &H8000000F
DoEvents
End Sub

Public Sub GeneratePayslipSignLedger(sUser, iGroup, iKey, PayrollDate, iPostLevel, iDivKey)
MainForm.picProgressBar.BackColor = &H8000000F
DoEvents
t = "SELECT tbl_Personnel_Compensation_Period.* " & _
    " FROM tbl_Personnel_Compensation_Period " & _
    " WHERE (PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then
    dDateFrom = FormatDateTime(rt!DateFrom, vbShortDate)
    dDateTo = FormatDateTime(rt!DateTo, vbShortDate)
End If
rt.Close

Select Case iGroup
    Case 1  'one record
        a = "SELECT COUNT(*) AS RecCount " & _
            " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK " & _
            " WHERE (dbo.tbl_Personnel_Payroll.PK = " & iKey & ")"
    Case 2  'Department
        a = "SELECT COUNT(*) AS RecCount " & _
            " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
            " WHERE (dbo.tbl_Personnel_ActionNew.DeptKey = " & iKey & ") " & _
            " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
            " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
    Case 3  'Division
        a = "SELECT COUNT(*) AS RecCount " & _
            " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
            " WHERE (dbo.tbl_Personnel_ActionNew.DivisionKey = " & iKey & ") " & _
            " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
            " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
End Select
'MsgBox a
If a = "" Then Exit Sub
If rs.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    If CDbl(ra!RecCount) > 0 Then
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll_Report_PaySlip WHERE (LogInName = '" & sUser & "')"
        
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_PaySlip " & _
                          " (LogInName, CompanyKey, PayrollPeriod, PostLevelDesc) " & _
                          " VALUES ('" & sUser & "', 1, '" & FormatDateTime(PayrollDate, vbShortDate) & "', " & _
                          " '" & IIf(iPostLevel = 1, "Rank in File", "Supervisory") & "')"
        iPK = 0
        t = "SELECT PK " & _
            " FROM tbl_Personnel_Payroll_Report_PaySlip " & _
            " WHERE (LogInName = '" & sUser & "')"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            iPK = rt!PK
        End If
        rt.Close
    End If
End If
ra.Close

If CDbl(iPK) <> 0 Then
    Select Case iGroup
        Case 1  'one record
            a = "SELECT dbo.tbl_Personnel_Payroll.EmployeeKey, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName + ',  ' + dbo.tbl_Personnel_Information.FirstName + '  ' + dbo.tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
                " dbo.tbl_Personnel_Division.Description AS Division, dbo.tbl_Personnel_Department.DepartmentName AS Department, " & _
                " dbo.tbl_Personnel_Position.PositionName AS Position, dbo.tbl_Govt_TaxStatus.TaxStatus, dbo.tbl_Personnel_EmploymentStatus.StatusName AS EmploymentStatus, " & _
                " dbo.tbl_Personnel_CompensationRate.Description AS CompensationRate,  dbo.tbl_Personnel_Compensation_Period.PayrollDate, dbo.tbl_Personnel_Compensation_Period.DateFrom, " & _
                " dbo.tbl_Personnel_Compensation_Period.DateTo, dbo.tbl_Personnel_Payroll.ActionMemoKey, dbo.tbl_Personnel_Payroll.PK " & _
                " FROM  dbo.tbl_Personnel_Information RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Payroll.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK ON dbo.tbl_Personnel_Information.PK = dbo.tbl_Personnel_IDNumber.ProfileKey LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_CompensationRate RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_CompensationRate.PK = dbo.tbl_Personnel_ActionNew.CompensationRateKey LEFT OUTER JOIN " & _
                " dbo.tbl_Govt_TaxStatus ON dbo.tbl_Personnel_ActionNew.TaxStatusKey = dbo.tbl_Govt_TaxStatus.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_EmploymentStatus ON dbo.tbl_Personnel_ActionNew.EmpStatusKey = dbo.tbl_Personnel_EmploymentStatus.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK FULL OUTER JOIN " & _
                " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_ActionNew.DivisionKey = dbo.tbl_Personnel_Division.PK " & _
                " WHERE (dbo.tbl_Personnel_Payroll.PK = " & iKey & ")"
        Case 2  'department
            a = "SELECT dbo.tbl_Personnel_Payroll.EmployeeKey, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName + ',  ' + dbo.tbl_Personnel_Information.FirstName + '  ' + dbo.tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
                " dbo.tbl_Personnel_Division.Description AS Division, dbo.tbl_Personnel_Department.DepartmentName AS Department, " & _
                " dbo.tbl_Personnel_Position.PositionName AS Position, dbo.tbl_Govt_TaxStatus.TaxStatus, dbo.tbl_Personnel_EmploymentStatus.StatusName AS EmploymentStatus, " & _
                " dbo.tbl_Personnel_CompensationRate.Description AS CompensationRate,  dbo.tbl_Personnel_Compensation_Period.PayrollDate, dbo.tbl_Personnel_Compensation_Period.DateFrom, " & _
                " dbo.tbl_Personnel_Compensation_Period.DateTo, dbo.tbl_Personnel_Payroll.ActionMemoKey, dbo.tbl_Personnel_Payroll.PK " & _
                " FROM  dbo.tbl_Personnel_Information RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Payroll.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK ON dbo.tbl_Personnel_Information.PK = dbo.tbl_Personnel_IDNumber.ProfileKey LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_CompensationRate RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_CompensationRate.PK = dbo.tbl_Personnel_ActionNew.CompensationRateKey LEFT OUTER JOIN " & _
                " dbo.tbl_Govt_TaxStatus ON dbo.tbl_Personnel_ActionNew.TaxStatusKey = dbo.tbl_Govt_TaxStatus.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_EmploymentStatus ON dbo.tbl_Personnel_ActionNew.EmpStatusKey = dbo.tbl_Personnel_EmploymentStatus.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK FULL OUTER JOIN " & _
                " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_ActionNew.DivisionKey = dbo.tbl_Personnel_Division.PK " & _
                " WHERE (dbo.tbl_Personnel_ActionNew.DeptKey = " & iKey & ") " & _
                " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
                " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
        Case 3  'Division
            a = "SELECT dbo.tbl_Personnel_Payroll.EmployeeKey, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName + ',  ' + dbo.tbl_Personnel_Information.FirstName + '  ' + dbo.tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
                " dbo.tbl_Personnel_Division.Description AS Division, dbo.tbl_Personnel_Department.DepartmentName AS Department, " & _
                " dbo.tbl_Personnel_Position.PositionName AS Position, dbo.tbl_Govt_TaxStatus.TaxStatus, dbo.tbl_Personnel_EmploymentStatus.StatusName AS EmploymentStatus, " & _
                " dbo.tbl_Personnel_CompensationRate.Description AS CompensationRate,  dbo.tbl_Personnel_Compensation_Period.PayrollDate, dbo.tbl_Personnel_Compensation_Period.DateFrom, " & _
                " dbo.tbl_Personnel_Compensation_Period.DateTo, dbo.tbl_Personnel_Payroll.ActionMemoKey, dbo.tbl_Personnel_Payroll.PK " & _
                " FROM  dbo.tbl_Personnel_Information RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Payroll.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK ON dbo.tbl_Personnel_Information.PK = dbo.tbl_Personnel_IDNumber.ProfileKey LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_CompensationRate RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_CompensationRate.PK = dbo.tbl_Personnel_ActionNew.CompensationRateKey LEFT OUTER JOIN " & _
                " dbo.tbl_Govt_TaxStatus ON dbo.tbl_Personnel_ActionNew.TaxStatusKey = dbo.tbl_Govt_TaxStatus.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_EmploymentStatus ON dbo.tbl_Personnel_ActionNew.EmpStatusKey = dbo.tbl_Personnel_EmploymentStatus.PK LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK FULL OUTER JOIN " & _
                " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_ActionNew.DivisionKey = dbo.tbl_Personnel_Division.PK " & _
                " WHERE (dbo.tbl_Personnel_ActionNew.DivisionKey = " & iKey & ") " & _
                " AND (dbo.tbl_Personnel_Position.PositionLevel = " & iPostLevel & ") " & _
                " AND (dbo.tbl_Personnel_Compensation_Period.PayrollDate = '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
    End Select
    If ra.State = adStateOpen Then ra.Close
    ra.Open a, ConnOmega
    If ra.RecordCount > 0 Then
        ConnOmega.Execute "UPDATE tbl_Personnel_Payroll_Report_PaySlip " & _
                          " SET PayrollPeriodFrom = '" & FormatDateTime(ra!DateFrom, vbShortDate) & "', " & _
                          " PayrollPeriodTo = '" & FormatDateTime(ra!DateTo, vbShortDate) & "', " & _
                          " PayrollRange = '" & Format(FormatDateTime(ra!DateFrom, vbShortDate), "mm/dd/yyyy") & " - " & Format(FormatDateTime(ra!DateTo, vbShortDate), "mm/dd/yyyy") & "' " & _
                          " WHERE (PK = " & iPK & ")"
        iRec = 0
        While Not ra.EOF
            DoEvents
            iRec = iRec + 1
            sCompRate = "RATE : " & ra!CompensationRate & " ("
            t = "SELECT dbo.tbl_Personnel_Payroll_Earnings_Table.Description, dbo.tbl_Personnel_ActionNew_Rate.Rate " & _
                " FROM  dbo.tbl_Personnel_ActionNew_Rate LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Payroll_Earnings_Table ON dbo.tbl_Personnel_ActionNew_Rate.EarningKey = dbo.tbl_Personnel_Payroll_Earnings_Table.PK " & _
                " WHERE (dbo.tbl_Personnel_ActionNew_Rate.MasterKey = " & ra!ActionMemoKey & ") " & _
                " ORDER BY dbo.tbl_Personnel_Payroll_Earnings_Table.Sorting"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                While Not rt.EOF
                    sCompRate = sCompRate & rt!Description & " = " & Format(rt!Rate, "#,##0.00") & " | "
                    rt.MoveNext
                Wend
            End If
            rt.Close
            sCompRate = Mid(sCompRate, 1, Len(sCompRate) - 3) & ")"
            
            dGross = 0: dDed = 0: dLoanBal = 0: dTimeSumm = 0
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_PaySlip_Det " & _
                              " (MasterKey, EmployeeKey, IDNumber, EmployeeName, Division, Department, Position, TaxStatus, " & _
                              " EmploymentStatus, CompensationRate, Gross, Deduction, LoanBal, TimeSumm) " & _
                              " VALUES (" & iPK & ", " & ra!EmployeeKey & ", '" & ra!IDNumber & "', '" & FORMATSQL(ra!EmployeeName) & "', " & _
                              " '" & FORMATSQL(ra!Division) & "', '" & FORMATSQL(ra!Department) & "', '" & FORMATSQL(ra!Position) & "', " & _
                              " '" & FORMATSQL(ra!TaxStatus) & "', '" & FORMATSQL(ra!EmploymentStatus) & "', '" & CStr(sCompRate) & "', " & _
                              " " & CDbl(dGross) & ", " & CDbl(dDed) & ", " & CDbl(dLoanBal) & ", " & CDbl(dTimeSumm) & ")"
            
            'Earnings
            iLine = 0: iEarnCol = 1
            t = "SELECT dbo.tbl_Personnel_Payroll_Earnings_Table.Description, dbo.tbl_Personnel_Payroll_Earnings.TotalAmount, " & _
                " dbo.tbl_Personnel_Payroll_Earnings.Remarks, dbo.tbl_Personnel_Payroll_Earnings_Table.Abbvt, " & _
                " dbo.tbl_Personnel_Payroll_Earnings.Hours " & _
                " FROM  dbo.tbl_Personnel_Payroll_Earnings LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Payroll_Earnings_Table ON dbo.tbl_Personnel_Payroll_Earnings.EarningKey = dbo.tbl_Personnel_Payroll_Earnings_Table.PK " & _
                " Where (dbo.tbl_Personnel_Payroll_Earnings.MasterKey = " & ra!PK & ") " & _
                " ORDER BY dbo.tbl_Personnel_Payroll_Earnings_Table.Sorting"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            While Not rt.EOF
                dGross = dGross + CDbl(Format(rt!TotalAmount, "#,##0.00"))
                iLine = iLine + 1
                If CDbl(iLine) = 9 Then iLine = 1: iEarnCol = iEarnCol + 1
                If CDbl(iEarnCol) = 1 Then
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_PaySlip_Det_Earnings " & _
                                      " (MasterKey, EmployeeKey, Line, EarnDescription, Amount, Hours) " & _
                                      " VALUES (" & iPK & ", " & ra!EmployeeKey & ", " & iLine & ", " & _
                                      " '" & FORMATSQL(rt!Abbvt & IIf(Trim(rt!Remarks) <> "", " [" & rt!Remarks & "]", "")) & "', " & _
                                      " '" & Format(rt!TotalAmount, "#,##0.00") & "', " & _
                                      " '" & Format(rt!Hours, "#0.00") & "hrs" & "')"
                ElseIf CDbl(iEarnCol) = 2 Then
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_PaySlip_Det_Earnings2 " & _
                                      " (MasterKey, EmployeeKey, Line, EarnDescription, Amount, Hours) " & _
                                      " VALUES (" & iPK & ", " & ra!EmployeeKey & ", " & iLine & ", " & _
                                      " '" & FORMATSQL(rt!Abbvt & IIf(Trim(rt!Remarks) <> "", " [" & rt!Remarks & "]", "")) & "', " & _
                                      " '" & Format(rt!TotalAmount, "#,##0.00") & "', " & _
                                      " '" & Format(rt!Hours, "#0.00") & "hrs" & "')"
                End If
                
                rt.MoveNext
            Wend
            rt.Close
            
            'add row
            If CDbl(iEarnCol) = 1 And CDbl(iLine) < 8 Then
                For i = 1 To (8 - CDbl(iLine))
                    iLine = iLine + 1
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_PaySlip_Det_Earnings " & _
                                      " (MasterKey, EmployeeKey, Line, EarnDescription, Amount) " & _
                                      " VALUES (" & iPK & ", " & ra!EmployeeKey & ", " & iLine & ", " & _
                                      " '', '')"
                Next i
            End If
            
            'Deductions
            iLine = 0: iDedCol = 1: iLineLoanBal = 0
            t = "SELECT dbo.tbl_Personnel_Payroll_Deductions_Table.Description, dbo.tbl_Personnel_Payroll_Deductions_Table.DedSched, " & _
                " dbo.tbl_Personnel_Payroll_Deductions.LoanKey, dbo.tbl_Personnel_Payroll_Deductions.Amount " & _
                " FROM  dbo.tbl_Personnel_Payroll_Deductions LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Payroll_Deductions.DeductionKey = dbo.tbl_Personnel_Payroll_Deductions_Table.PK " & _
                " Where (dbo.tbl_Personnel_Payroll_Deductions.MasterKey = " & ra!PK & ") " & _
                " ORDER BY dbo.tbl_Personnel_Payroll_Deductions_Table.Sorting"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            While Not rt.EOF
                dDed = dDed + CDbl(Format(rt!Amount, "#,##0.00"))
                
                'MsgBox rt!DedSched
                
                If CDbl(rt!DedSched) = 1 Then
                    'MsgBox IIf(IsNull(rt!LoanKey), 0, rt!LoanKey)
                    dLoanBal = dLoanBal + CDbl(Format(rt!Amount, "#,##0.00"))
                    iLineLoanBal = iLineLoanBal + 1
                    u = "SELECT ROUND(SUM(Balance), 2) AS Bal " & _
                        " From dbo.tbl_Personnel_Loans_SL " & _
                        " WHERE (LoanKey = " & IIf(IsNull(rt!LoanKey), 0, rt!LoanKey) & ") " & _
                        " AND (TransactionDate <= '" & FormatDateTime(PayrollDate, vbShortDate) & "')"
                    If ru.State = adStateOpen Then ru.Close
                    ru.Open u, ConnOmega
                    If ru.RecordCount > 0 Then
'                        MsgBox ru!Bal
                        If CDbl(ru!Bal) > 0 Then
                            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_PaySlip_Det_LoanBal " & _
                                              " (MasterKey, EmployeeKey, Line, LoanDesc, Amount) " & _
                                              " VALUES (" & iPK & ", " & ra!EmployeeKey & ", " & iLineLoanBal & ", " & _
                                              " '" & FORMATSQL(rt!Description) & "', '" & Format(ru!Bal, "#,##0.00") & "') "
                        End If
                    End If
                    ru.Close
                End If
                
                iLine = iLine + 1
                If CDbl(iLine) = 9 Then iLine = 1: iDedCol = iDedCol + 1
                If CDbl(iDedCol) = 1 Then
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_PaySlip_Det_Deductions " & _
                                      " (MasterKey, EmployeeKey, Line, DedDescription, Amount) " & _
                                      " VALUES (" & iPK & ", " & ra!EmployeeKey & ", " & iLine & ", " & _
                                      " '" & FORMATSQL(rt!Description) & "', '" & Format(rt!Amount, "#,##0.00") & "')"
                ElseIf CDbl(iDedCol) = 2 Then
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_PaySlip_Det_Deductions2 " & _
                                      " (MasterKey, EmployeeKey, Line, DedDescription, Amount) " & _
                                      " VALUES (" & iPK & ", " & ra!EmployeeKey & ", " & iLine & ", " & _
                                      " '" & FORMATSQL(rt!Description) & "', '" & Format(rt!Amount, "#,##0.00") & "')"
                End If
                rt.MoveNext
            Wend
            rt.Close
            
            'add row
            If CDbl(iDedCol) = 1 And CDbl(iLine) < 8 Then
                For i = 1 To (8 - CDbl(iLine))
                    iLine = iLine + 1
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_PaySlip_Det_Deductions " & _
                                      " (MasterKey, EmployeeKey, Line, DedDescription, Amount) " & _
                                      " VALUES (" & iPK & ", " & ra!EmployeeKey & ", " & iLine & ", " & _
                                      " '', '')"
                Next i
            End If
            
            
            iLine = 0
            t = "SELECT dbo.tbl_Personnel_AbsentLateUndertime_Details.AbsType, " & _
                " SUM(dbo.tbl_Personnel_AbsentLateUndertime_Details.TotalHours) AS TotalHours " & _
                " FROM  dbo.tbl_Personnel_AbsentLateUndertime_Details LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_AbsentLateUndertime ON dbo.tbl_Personnel_AbsentLateUndertime_Details.MasterKey = dbo.tbl_Personnel_AbsentLateUndertime.PK " & _
                " WHERE (dbo.tbl_Personnel_AbsentLateUndertime_Details.EmployeeKey = " & ra!EmployeeKey & ") " & _
                " AND (dbo.tbl_Personnel_AbsentLateUndertime.DateApplied >= '" & dDateFrom & "') " & _
                " AND (dbo.tbl_Personnel_AbsentLateUndertime.DateApplied <= '" & dDateTo & "') " & _
                " AND (dbo.tbl_Personnel_AbsentLateUndertime.Posted = 1) " & _
                " AND (dbo.tbl_Personnel_AbsentLateUndertime.DivisionKey = " & iDivKey & ") " & _
                " GROUP BY dbo.tbl_Personnel_AbsentLateUndertime_Details.AbsType " & _
                " ORDER BY dbo.tbl_Personnel_AbsentLateUndertime_Details.AbsType"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            While Not rt.EOF
                iLine = iLine + 1
                dTimeSumm = dTimeSumm + CDbl(rt!TotalHours)
                ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Report_PaySlip_Det_TimeSumm " & _
                                  " (MasterKey, EmployeeKey, Line, TimeSummDesc, Hours) " & _
                                  " VALUES (" & iPK & ", " & ra!EmployeeKey & ", " & iLine & ",  " & _
                                  " '" & IIf(rt!AbsType = 1, "Absent", IIf(rt!AbsType = 2, "Late", IIf(rt!AbsType = 3, "Undertime", ""))) & "', " & _
                                  " '" & Format(rt!TotalHours, "#0.00") & " hrs" & "')"
                rt.MoveNext
            Wend
            rt.Close
            
            
            ConnOmega.Execute "UPDATE tbl_Personnel_Payroll_Report_PaySlip_Det " & _
                              " SET Gross = " & CDbl(dGross) & ", " & _
                              " Deduction = " & CDbl(dDed) & ", " & _
                              " LoanBal = " & CDbl(dLoanBal) & ", " & _
                              " TimeSumm = " & CDbl(dTimeSumm) & " " & _
                              " WHERE (MasterKey = " & iPK & ") " & _
                              " AND (EmployeeKey = " & ra!EmployeeKey & ")"
                              
            UpdateProgress_No_Percent MainForm.picProgressBar, iRec / ra.RecordCount
            ra.MoveNext
        Wend
    End If
    ra.Close
End If
MainForm.picProgressBar.BackColor = &H8000000F
DoEvents
End Sub
