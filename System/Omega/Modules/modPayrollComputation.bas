Attribute VB_Name = "modPayrollComputation"
Option Explicit

Public gTaxable As Double
Public gSSS As Double
Public gPHIC As Double
Public gPagIbig As Double

Public gTaxableF As Double
Public gSSSF As Double
Public gPHICF As Double
Public gPagIbigF As Double

Dim iPayrollKey As Long

Dim sCtrl, dAmt, dRatePerHour, dTaxable, dNonTaxable, gTotEarnings, Array1, dDedAmt

Dim dSSSEmp, dSSSEmr, dSSSEC, dPHICEmp, dPHICEmr, dPagIbigEmp, dPagIbigEmr, dForTaxable, dTaxDue, dLoanBal
Dim gbl_TotalEarningTmp As Double

Public Function Get_Perfect_Hours(iChkPerfHrs, dPayDate, iDiv, iPayPeriod) As Double
Get_Perfect_Hours = 0
If iChkPerfHrs = 0 Then
    a = "SELECT TOP (1) PerfectHours " & _
        " FROM tbl_System_PerfectHours " & _
        " WHERE (EffectDate <= '" & FormatDateTime(dPayDate, vbShortDate) & "') " & _
        " ORDER BY EffectDate DESC"
Else
    a = "SELECT NoHours as PerfectHours " & _
        " FROM tbl_Personnel_Setup_DailyPerfectDays " & _
        " WHERE (DivisionKey = " & iDiv & ") " & _
        " AND (PayrollPeriodKey = " & iPayPeriod & ")"
End If
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    Get_Perfect_Hours = ra!PerfectHours
End If
ra.Close
End Function


Public Function Get_AbsentLateUndertime_Hours(iEmp, dFrom, dTo) As Double
Get_AbsentLateUndertime_Hours = 0
Dim dHours, dMinsToHrs
a = "SELECT SUM(dbo.tbl_Personnel_AbsentLateUndertime_Details.TotalHours) AS Hours " & _
    " FROM  dbo.tbl_Personnel_AbsentLateUndertime_Details LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_AbsentLateUndertime ON dbo.tbl_Personnel_AbsentLateUndertime_Details.MasterKey = dbo.tbl_Personnel_AbsentLateUndertime.PK " & _
    " WHERE (dbo.tbl_Personnel_AbsentLateUndertime_Details.EmployeeKey = " & iEmp & ") " & _
    " AND (dbo.tbl_Personnel_AbsentLateUndertime.DateApplied >= '" & FormatDateTime(dFrom, vbShortDate) & "') " & _
    " AND (dbo.tbl_Personnel_AbsentLateUndertime.DateApplied <= '" & FormatDateTime(dTo, vbShortDate) & "') " & _
    " AND (dbo.tbl_Personnel_AbsentLateUndertime.Posted = 1)"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    Get_AbsentLateUndertime_Hours = IIf(IsNull(ra!Hours), 0, ra!Hours)
End If
ra.Close
End Function

Public Function EarningsPerHour(ActionMemoKey, RefRate) As Double
EarningsPerHour = 0
a = "SELECT RatePerHour " & _
    " From dbo.tbl_Personnel_ActionNew_Rate " & _
    " WHERE (MasterKey = " & ActionMemoKey & ") " & _
    " AND (EarningKey = " & RefRate & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    EarningsPerHour = ra!RatePerHour
End If
ra.Clone
End Function

Public Sub COMPUTE_COMPENSATION(iHourKey)
gbl_TotalEarning = 0
gTaxable = 0: gSSS = 0: gPHIC = 0: gPagIbig = 0: dTaxable = 0: dNonTaxable = 0: gTotEarnings = 0: iPayrollKey = 0: sCtrl = ""
t = "SELECT dbo.tbl_Personnel_Hours.PK, dbo.tbl_Personnel_Hours.EmployeeKey, dbo.tbl_Personnel_Hours.PayrollPeriodKey, " & _
    " dbo.tbl_Personnel_Hours.ActionMemoKey, dbo.tbl_Personnel_Hours.Adjustment, dbo.tbl_Personnel_Hours.AdjustmentRem, " & _
    " dbo.tbl_Personnel_Compensation_Period.DateTo, dbo.tbl_Personnel_Compensation_Period.Terms AS PeriodTerms, " & _
    " tbl_Personnel_ActionNew_DedTable_1.Terms AS GovtTerms, dbo.tbl_Personnel_ActionNew_DedTable.Terms AS LoanTerms, " & _
    " dbo.tbl_Personnel_ActionNew.DivisionKey, dbo.tbl_Personnel_ActionNew.Is_SSS, dbo.tbl_Personnel_ActionNew.Is_PHIC, " & _
    " dbo.tbl_Personnel_ActionNew.Is_PAGIBIG, dbo.tbl_Personnel_ActionNew.Is_TIN, dbo.tbl_Personnel_ActionNew.TaxCategoryKey, " & _
    " dbo.tbl_Personnel_ActionNew.TaxStatusKey, dbo.tbl_Personnel_Compensation_Period.PayrollDate, dbo.tbl_Personnel_ActionNew.EmpStatusKey, " & _
    " dbo.tbl_Personnel_ActionNew.PositionsKey, dbo.tbl_Personnel_ActionNew.CompensationRateKey " & _
    " FROM  dbo.tbl_Personnel_ActionNew_DedTable RIGHT OUTER JOIN " & _
    " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_ActionNew_DedTable.PK = dbo.tbl_Personnel_ActionNew.LoanDeductionKey RIGHT OUTER JOIN " & _
    " dbo.tbl_Personnel_Hours ON dbo.tbl_Personnel_ActionNew.PK = dbo.tbl_Personnel_Hours.ActionMemoKey LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_ActionNew_DedTable AS tbl_Personnel_ActionNew_DedTable_1 ON dbo.tbl_Personnel_ActionNew.GovtDeductionKey = tbl_Personnel_ActionNew_DedTable_1.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Hours.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
    " WHERE (dbo.tbl_Personnel_Hours.PK = " & iHourKey & ")"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then
    
    v = "SELECT tbl_Personnel_Payroll.* " & _
        " FROM tbl_Personnel_Payroll " & _
        " WHERE (EmployeeKey = " & rt!EmployeeKey & ") " & _
        " AND (PayrollPeriodKey = " & rt!PayrollPeriodKey & ")"
    If rv.State = adStateOpen Then rv.Close
    rv.Open v, ConnOmega
    If rv.RecordCount = 0 Then
        u = "SELECT TOP (1) dbo.tbl_Personnel_Payroll.Ctrl " & _
            " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
            " Where (Year(dbo.tbl_Personnel_Compensation_Period.DateTo) = " & Format(rt!PayrollDate, "yyyy") & ") " & _
            " ORDER BY dbo.tbl_Personnel_Payroll.Ctrl DESC"
        If ru.State = adStateOpen Then ru.Close
        ru.Open u, ConnOmega
        If ru.RecordCount > 0 Then
            sCtrl = Format(CDbl(ru!Ctrl) + 1, "000000000#")
        Else
            sCtrl = Format(rt!PayrollDate, "yyyy") & "000000"
        End If
        ru.Close
        
        Do
            u = "SELECT tbl_Personnel_Payroll.* " & _
                " FROM tbl_Personnel_Payroll " & _
                " WHERE (Ctrl = '" & sCtrl & "')"
            If ru.State = adStateOpen Then ru.Close
            ru.Open u, ConnOmega
            If ru.RecordCount = 0 Then
                ru.Close
                Exit Do
            End If
            ru.Close
            sCtrl = Format(CDbl(sCtrl) + 1, "000000000#")
        Loop
        
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll " & _
                          " (Ctrl, EmployeeKey, PayrollPeriodKey, ActionMemoKey, LastModified) " & _
                          " VALUES ('" & sCtrl & "', " & rt!EmployeeKey & ", " & rt!PayrollPeriodKey & ", " & _
                          " " & rt!ActionMemoKey & ", '" & CStr(Now) & " - " & gbl_CompleteName & "')"
        
        u = "SELECT PK " & _
            " FROM tbl_Personnel_Payroll " & _
            " WHERE (Ctrl = '" & sCtrl & "')"
        If ru.State = adStateOpen Then ru.Close
        ru.Open u, ConnOmega
        If ru.RecordCount > 0 Then
            iPayrollKey = ru!PK
        End If
        ru.Close
    
    Else
        iPayrollKey = rv!PK
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll_Earnings WHERE (MasterKey = " & iPayrollKey & ")"
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll_Deductions WHERE (MasterKey = " & iPayrollKey & ")"
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Payroll_EmployerShare WHERE (MasterKey = " & iPayrollKey & ")"
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Loans_SL " & _
                          " WHERE (PayrollKey = " & iPayrollKey & ") " & _
                          " AND (TransactionDate = '" & FormatDateTime(rt!PayrollDate, vbShortDate) & "') " & _
                          " AND (InOut = 'O')"
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Deduction_SL " & _
                          " WHERE (PayrollKey = " & iPayrollKey & ") " & _
                          " AND (TransactionDate = '" & FormatDateTime(rt!PayrollDate, vbShortDate) & "') " & _
                          " AND (TransactionType = 2) " & _
                          " AND (InOut = 'O')"
    End If
    rv.Close
    
    If CDbl(iPayrollKey) <> 0 Then
        ' ---- Earnings (Regular Hours)
        u = "SELECT dbo.tbl_Personnel_Hours_Regular.NoHours, dbo.tbl_Personnel_Payroll_Earnings_Table.RefRate, " & _
            " dbo.tbl_Personnel_Hours_Regular.EarningKey " & _
            " FROM  dbo.tbl_Personnel_Hours_Regular LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Payroll_Earnings_Table ON dbo.tbl_Personnel_Hours_Regular.EarningKey = dbo.tbl_Personnel_Payroll_Earnings_Table.PK " & _
            " WHERE (dbo.tbl_Personnel_Hours_Regular.MasterKey = " & rt!PK & ")"
        If ru.State = adStateOpen Then ru.Close
        ru.Open u, ConnOmega
        While Not ru.EOF
            dRatePerHour = EarningsPerHour(rt!ActionMemoKey, ru!RefRate)
            dAmt = Format(CDbl(ru!NoHours) * CDbl(dRatePerHour) * Earning_Multiplier(rt!PayrollDate, ru!EarningKey, rt!CompensationRateKey), "#0.00")
            dTaxable = 0: dNonTaxable = 0
            v = "SELECT Tax, SSS, PHIC, PagIbig " & _
                " FROM tbl_Personnel_Payroll_Earnings_Table " & _
                " WHERE (PK = " & ru!EarningKey & ")"
            If rv.State = adStateOpen Then rv.Close
            rv.Open v, ConnOmega
            If rv.RecordCount > 0 Then
                'If rv!TAX = 1 Then gTaxable = gTaxable + CDbl(dAmt)
                'If rv!SSS = 1 Then gSSS = gSSS + CDbl(dAmt)
                'If rv!PHIC = 1 Then gPHIC = gPHIC + CDbl(dAmt)
                'If rv!PagIbig = 1 Then gPagIbig = gPagIbig + CDbl(dAmt)
                If rv!TAX = 1 Then dTaxable = dAmt: dNonTaxable = 0 Else dTaxable = 0: dNonTaxable = dAmt
            End If
            rv.Close
            
            gTotEarnings = gTotEarnings + CDbl(dAmt)
            gbl_TotalEarning = gbl_TotalEarning + CDbl(dAmt)
            
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Earnings " & _
                              " (MasterKey, EarningKey, Taxable, NonTaxable, Hours) " & _
                              " VALUES (" & iPayrollKey & ", " & ru!EarningKey & ", " & _
                              " " & CDbl(dTaxable) & ", " & CDbl(dNonTaxable) & ", " & _
                              " " & CDbl(ru!NoHours) & ")"
            ru.MoveNext
        Wend
        ru.Close
        
        ' ---- Earnings (Overtime)
        u = "SELECT dbo.tbl_Personnel_Hours_Overtime.NoHours, dbo.tbl_Personnel_Payroll_Earnings_Table.RefRate, " & _
            " dbo.tbl_Personnel_Hours_Overtime.EarningKey " & _
            " FROM  dbo.tbl_Personnel_Hours_Overtime LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Payroll_Earnings_Table ON dbo.tbl_Personnel_Hours_Overtime.EarningKey = dbo.tbl_Personnel_Payroll_Earnings_Table.PK " & _
            " WHERE (dbo.tbl_Personnel_Hours_Overtime.MasterKey = " & rt!PK & ")"
        If ru.State = adStateOpen Then ru.Close
        ru.Open u, ConnOmega
        While Not ru.EOF
            dRatePerHour = EarningsPerHour(rt!ActionMemoKey, ru!RefRate)
            dAmt = Format((CDbl(ru!NoHours) * CDbl(dRatePerHour)) * Overtime_Multiplier(rt!PayrollDate, ru!EarningKey, rt!CompensationRateKey), "#0.00")
            dTaxable = 0: dNonTaxable = 0
            v = "SELECT Tax, SSS, PHIC, PagIbig " & _
                " FROM tbl_Personnel_Payroll_Earnings_Table " & _
                " WHERE (PK = " & ru!EarningKey & ")"
            If rv.State = adStateOpen Then rv.Close
            rv.Open v, ConnOmega
            If rv.RecordCount > 0 Then
                'If rv!TAX = 1 Then gTaxable = gTaxable + CDbl(dAmt)
                'If rv!SSS = 1 Then gSSS = gSSS + CDbl(dAmt)
                'If rv!PHIC = 1 Then gPHIC = gPHIC + CDbl(dAmt)
                'If rv!PagIbig = 1 Then gPagIbig = gPagIbig + CDbl(dAmt)
                If rv!TAX = 1 Then dTaxable = dAmt: dNonTaxable = 0 Else dTaxable = 0: dNonTaxable = dAmt
            End If
            rv.Close
            
            gTotEarnings = gTotEarnings + CDbl(dAmt)
            gbl_TotalEarning = gbl_TotalEarning + CDbl(dAmt)
            
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Earnings " & _
                              " (MasterKey, EarningKey, Taxable, NonTaxable, Hours) " & _
                              " VALUES (" & iPayrollKey & ", " & ru!EarningKey & ", " & _
                              " " & CDbl(dTaxable) & ", " & CDbl(dNonTaxable) & ", " & _
                              " " & CDbl(ru!NoHours) & ")"
            ru.MoveNext
        Wend
        ru.Close
        
        ' ---- Earnings (Adjustment)
        If CDbl(rt!Adjustment) <> 0 Then
            dAmt = CDbl(rt!Adjustment): dTaxable = 0: dNonTaxable = 0
            v = "SELECT Tax, SSS, PHIC, PagIbig " & _
                " FROM tbl_Personnel_Payroll_Earnings_Table " & _
                " WHERE (PK = 7)"
            If rv.State = adStateOpen Then rv.Close
            rv.Open v, ConnOmega
            If rv.RecordCount > 0 Then
                'If rv!TAX = 1 Then gTaxable = gTaxable + CDbl(dAmt)
                'If rv!SSS = 1 Then gSSS = gSSS + CDbl(dAmt)
                'If rv!PHIC = 1 Then gPHIC = gPHIC + CDbl(dAmt)
                'If rv!PagIbig = 1 Then gPagIbig = gPagIbig + CDbl(dAmt)
                If rv!TAX = 1 Then dTaxable = dAmt: dNonTaxable = 0 Else dTaxable = 0: dNonTaxable = dAmt
            End If
            rv.Close
            
            gTotEarnings = gTotEarnings + CDbl(dAmt)
            gbl_TotalEarning = gbl_TotalEarning + CDbl(dAmt)
            
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Earnings " & _
                              " (MasterKey, EarningKey, Taxable, NonTaxable, Remarks) " & _
                              " VALUES (" & iPayrollKey & ", 7, " & _
                              " " & CDbl(dTaxable) & ", " & CDbl(dNonTaxable) & ", " & _
                              " '" & FORMATSQL(rt!AdjustmentRem) & "')"
        End If
        
        gbl_TotalEarningTmp = gbl_TotalEarning - gbl_MinTakeHomePay
        
        ' ---- Statutory Deduction
        If CDbl(rt!PeriodTerms) = CDbl(rt!GovtTerms) Then
            gTaxableF = 0: gSSSF = 0: gPHICF = 0: gPagIbigF = 0
            u = "SELECT (CASE dbo.tbl_Personnel_Payroll_Earnings_Table.SSS WHEN 1 THEN ROUND(ISNULL(SUM(dbo.tbl_Personnel_Payroll_Earnings.TotalAmount), 0), 2) ELSE 0 END) AS SSS, " & _
                " (CASE dbo.tbl_Personnel_Payroll_Earnings_Table.PHIC WHEN 1 THEN ROUND(ISNULL(SUM(dbo.tbl_Personnel_Payroll_Earnings.TotalAmount), 0), 2) ELSE 0 END) AS PHIC, " & _
                " (CASE dbo.tbl_Personnel_Payroll_Earnings_Table.PagIbig WHEN 1 THEN ROUND(ISNULL(SUM(dbo.tbl_Personnel_Payroll_Earnings.TotalAmount), 0), 2) ELSE 0 END) AS PagIbig, " & _
                " (CASE dbo.tbl_Personnel_Payroll_Earnings_Table.Tax WHEN 1 THEN ROUND(ISNULL(SUM(dbo.tbl_Personnel_Payroll_Earnings.TotalAmount), 0), 2) ELSE 0 END) AS TAX " & _
                " FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_ActionNew.PK = dbo.tbl_Personnel_Payroll.ActionMemoKey RIGHT OUTER JOIN " & _
                " dbo.tbl_Personnel_Payroll_Earnings LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Payroll_Earnings_Table ON dbo.tbl_Personnel_Payroll_Earnings.EarningKey = dbo.tbl_Personnel_Payroll_Earnings_Table.PK ON dbo.tbl_Personnel_Payroll.PK = dbo.tbl_Personnel_Payroll_Earnings.MasterKey LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
                " Where (dbo.tbl_Personnel_Payroll.EmployeeKey = " & rt!EmployeeKey & ") " & _
                " And (Year(dbo.tbl_Personnel_Compensation_Period.DateTo) = " & Format(rt!PayrollDate, "yyyy") & ") " & _
                " And (Month(dbo.tbl_Personnel_Compensation_Period.DateTo) = " & Format(rt!PayrollDate, "mm") & ") " & _
                " And (dbo.tbl_Personnel_ActionNew.DivisionKey = " & rt!DivisionKey & ") " & _
                " GROUP BY dbo.tbl_Personnel_Payroll_Earnings_Table.SSS, dbo.tbl_Personnel_Payroll_Earnings_Table.PHIC, dbo.tbl_Personnel_Payroll_Earnings_Table.PagIbig, dbo.tbl_Personnel_Payroll_Earnings_Table.Tax"
            If ru.State = adStateOpen Then ru.Close
            ru.Open u, ConnOmega
            If ru.RecordCount > 0 Then
                gTaxableF = ru!TAX
                gSSSF = ru!SSS
                gPHICF = ru!PHIC
                gPagIbigF = ru!PAGIBIG
            End If
            ru.Close
            dSSSEmp = 0: dSSSEmr = 0: dSSSEC = 0: dPHICEmp = 0: dPHICEmr = 0: dPagIbigEmp = 0: dPagIbigEmr = 0: dForTaxable = 0
            u = "SELECT tbl_Personnel_Payroll_Deductions_Table.* " & _
                " FROM tbl_Personnel_Payroll_Deductions_Table " & _
                " WHERE (DedSched = " & rt!GovtTerms & ") " & _
                " AND (GovtDed = 2) AND (GovtDedMain = 1) " & _
                " ORDER BY Sorting"
            If ru.State = adStateOpen Then ru.Close
            ru.Open u, ConnOmega
            While Not ru.EOF
                Select Case ru!PK
                    Case 1  'SSS
                        If CDbl(rt!Is_SSS) = 1 Then
                            Array1 = Split(SSS_Cont_Value(gSSSF, rt!PayrollDate), "|", -1, 1)
                            dSSSEmp = Array1(0)
                            If CDbl(dSSSEmp) <= gbl_TotalEarningTmp Then
                                gbl_TotalEarningTmp = gbl_TotalEarningTmp - CDbl(dSSSEmp)
                                ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Deductions " & _
                                                  " (MasterKey, DeductionKey, Amount) " & _
                                                  " VALUES (" & iPayrollKey & ", 1, " & CDbl(Array1(0)) & ")"
                                ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_EmployerShare " & _
                                                  " (MasterKey, DeductionKey, Amount) " & _
                                                  " VALUES (" & iPayrollKey & ", 2, " & CDbl(Array1(1)) & ")"
                                ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_EmployerShare " & _
                                                  " (MasterKey, DeductionKey, Amount) " & _
                                                  " VALUES (" & iPayrollKey & ", 3, " & CDbl(Array1(2)) & ")"
                            End If
                        End If
                    Case 4  'PHIC
                        If CDbl(rt!Is_PHIC) = 1 Then
                            Array1 = Split(PHIC_Cont_Value(gPHICF, rt!PayrollDate), "|", -1, 1)
                            dPHICEmp = Array1(0)
                            If CDbl(dPHICEmp) <= gbl_TotalEarningTmp Then
                                gbl_TotalEarningTmp = gbl_TotalEarningTmp - CDbl(dPHICEmp)
                                ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Deductions " & _
                                                  " (MasterKey, DeductionKey, Amount) " & _
                                                  " VALUES (" & iPayrollKey & ", 4, " & CDbl(Array1(0)) & ")"
                                ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_EmployerShare " & _
                                                  " (MasterKey, DeductionKey, Amount) " & _
                                                  " VALUES (" & iPayrollKey & ", 5, " & CDbl(Array1(1)) & ")"
                            End If
                        End If
                    Case 6  'PagIbig
                        If CDbl(rt!Is_PAGIBIG) = 1 Then
                            Array1 = Split(PagIbig_Cont_Value(gPagIbigF, rt!PayrollDate), "|", -1, 1)
                            
                            Dim dPagIbigAddCont
                            dPagIbigAddCont = 0
                            v = "SELECT TOP (1) tbl_Personnel_PagIbig_AddAmount.* " & _
                                " FROM tbl_Personnel_PagIbig_AddAmount " & _
                                " WHERE (EmployeeKey = " & rt!EmployeeKey & ") " & _
                                " AND (EffectDate <= '" & FormatDateTime(rt!PayrollDate, vbShortDate) & "') " & _
                                " ORDER BY EffectDate DESC"
                            If rv.State = adStateOpen Then rv.Close
                            rv.Open v, ConnOmega
                            If rv.RecordCount > 0 Then
                                dPagIbigAddCont = rv!AddAmount
                            End If
                            rv.Close
                            
                            dPagIbigEmp = CDbl(Array1(0)) ' + CDbl(dPagIbigAddCont)
                            If CDbl(dPagIbigEmp) <= gbl_TotalEarningTmp Then
                                gbl_TotalEarningTmp = gbl_TotalEarningTmp - CDbl(dPagIbigEmp)
                                
                                If CDbl(dPagIbigAddCont) <= gbl_TotalEarningTmp Then
                                    gbl_TotalEarningTmp = gbl_TotalEarningTmp - CDbl(dPagIbigAddCont)
                                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Deductions " & _
                                                      " (MasterKey, DeductionKey, Amount) " & _
                                                      " VALUES (" & iPayrollKey & ", 6, " & CDbl(Array1(0)) + CDbl(dPagIbigAddCont) & ")"
                                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_EmployerShare " & _
                                                      " (MasterKey, DeductionKey, Amount) " & _
                                                      " VALUES (" & iPayrollKey & ", 7, " & CDbl(Array1(1)) & ")"
                                Else
                                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Deductions " & _
                                                      " (MasterKey, DeductionKey, Amount) " & _
                                                      " VALUES (" & iPayrollKey & ", 6, " & CDbl(Array1(0)) & ")"
                                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_EmployerShare " & _
                                                      " (MasterKey, DeductionKey, Amount) " & _
                                                      " VALUES (" & iPayrollKey & ", 7, " & CDbl(Array1(1)) & ")"
                                End If
                            End If
                        End If
                    Case 8  'Withholding
                        If CDbl(rt!Is_TIN) = 1 Then
                            dForTaxable = gTaxableF - CDbl(dSSSEmp) - CDbl(dPHICEmp) - CDbl(dPagIbigEmp)
                            dTaxDue = WithHolding_Cont_Value(dForTaxable, rt!TaxCategoryKey, rt!TaxStatusKey, rt!PayrollDate)
                            If CDbl(dTaxDue) <> 0 Then
                                If CDbl(dTaxDue) <= gbl_TotalEarningTmp Then
                                    gbl_TotalEarningTmp = gbl_TotalEarningTmp - CDbl(dTaxDue)
                                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Deductions " & _
                                                      " (MasterKey, DeductionKey, Amount) " & _
                                                      " VALUES (" & iPayrollKey & ", 8, " & CDbl(dTaxDue) & ")"
                                End If
                            End If
                        End If
                End Select
                ru.MoveNext
            Wend
            ru.Close
        End If
        
        ' ---- Loan/s Deductions
        If CDbl(rt!PeriodTerms) = CDbl(rt!LoanTerms) Then
            u = "SELECT tbl_Personnel_Payroll_Deductions_Table.* " & _
                " FROM tbl_Personnel_Payroll_Deductions_Table " & _
                " WHERE (DedSched = " & rt!LoanTerms & ") " & _
                " AND (GovtDed = 1) " & _
                " ORDER BY Sorting"
            If ru.State = adStateOpen Then ru.Close
            ru.Open u, ConnOmega
            While Not ru.EOF
                dLoanBal = 0
                v = "SELECT EmpPK, LoanType, ZeroOut, PK, Amortization, " & _
                    " (SELECT ROUND(SUM(Balance), 2) AS Bal " & _
                    " From dbo.tbl_Personnel_Loans_SL " & _
                    " WHERE (LoanKey = dbo.tbl_Personnel_Loans.PK)) AS Balance, DateFrom " & _
                    " From dbo.tbl_Personnel_Loans " & _
                    " WHERE ((SELECT ROUND(SUM(Balance), 2) AS Bal " & _
                    " FROM  dbo.tbl_Personnel_Loans_SL AS tbl_Personnel_Loans_SL_1 " & _
                    " WHERE (LoanKey = dbo.tbl_Personnel_Loans.PK)) > 0) " & _
                    " AND (EmpPK = " & rt!EmployeeKey & ") " & _
                    " AND (LoanType = " & ru!PK & ") " & _
                    " AND (ZeroOut = 0) " & _
                    " AND (DateFrom <= '" & FormatDateTime(rt!PayrollDate, vbShortDate) & "')"
                If rv.State = adStateOpen Then rv.Close
                rv.Open v, ConnOmega
                If rv.RecordCount > 0 Then
                    dLoanBal = rv!Balance
                    If CDbl(rv!Balance) >= CDbl(rv!Amortization) Then
                        dAmt = Format(rv!Amortization, "#0.00")
                    Else
                        dAmt = Format(rv!Balance, "#0.00")
                    End If
                    
                    If CDbl(dAmt) > gbl_TotalEarningTmp Then
                        dAmt = gbl_TotalEarningTmp
                    End If
                    
                    If CDbl(dAmt) > 0 Then
                        dLoanBal = dLoanBal - CDbl(dAmt)
                        gbl_TotalEarningTmp = gbl_TotalEarningTmp - CDbl(dAmt)
                        ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Deductions " & _
                                          " (MasterKey, DeductionKey, Amount, LoanKey, LoanBalance) " & _
                                          " VALUES (" & iPayrollKey & ", " & ru!PK & ", " & _
                                          " " & CDbl(dAmt) & ", " & rv!PK & ", " & CDbl(dLoanBal) & ")"
                        If CDbl(ru!withSL) = 1 Then
                            ConnOmega.Execute "INSERT INTO tbl_Personnel_Loans_SL " & _
                                              " (EmpPK, LoanKey, LoanType, InOut, TransactionDate, PayrollKey, Remarks, Credit) " & _
                                              " VALUES (" & rt!EmployeeKey & ", " & rv!PK & ", " & ru!PK & ", 'O', " & _
                                              " '" & FormatDateTime(rt!PayrollDate, vbShortDate) & "', " & iPayrollKey & ", 'Payroll Deduction', " & _
                                              " " & CDbl(dAmt) & ")"
                        End If
                    End If
                End If
                rv.Close
                
                ru.MoveNext
            Wend
            ru.Close
        End If
        
        ' ---- Employee Deductions
        u = "SELECT DeductionKey, SourceKey, Amount " & _
            " From dbo.tbl_Personnel_Deduction_Payroll " & _
            " WHERE (DivisionKey = " & rt!DivisionKey & ") " & _
            " AND (PayrollPeriodKey = " & rt!PayrollPeriodKey & ") " & _
            " AND (EmployeeKey = " & rt!EmployeeKey & ") " & _
            " AND (Amount <> 0)"
        If ru.State = adStateOpen Then ru.Close
        ru.Open u, ConnOmega
        While Not ru.EOF
        
            If CDbl(ru!Amount) <= gbl_TotalEarningTmp Then
                dAmt = ru!Amount
            Else
                dAmt = gbl_TotalEarningTmp
            End If
            
            If CDbl(dAmt) <> 0 Then
                gbl_TotalEarningTmp = gbl_TotalEarningTmp - CDbl(dAmt)
                v = "SELECT tbl_Personnel_Payroll_Deductions_Table.* " & _
                    " FROM tbl_Personnel_Payroll_Deductions_Table " & _
                    " WHERE (PK = " & ru!DeductionKey & ")"
                If rv.State = adStateOpen Then rv.Close
                rv.Open v, ConnOmega
                If rv.RecordCount > 0 Then
                    If CDbl(rv!withSL) = 1 Then
                        ConnOmega.Execute "INSERT INTO tbl_Personnel_Deduction_SL " & _
                                          " (SourceKey, EmployeeKey, DeductionKey, TransactionDate, Remarks, InOut, Credit, " & _
                                          " PayrollKey, TransactionType) " & _
                                          " VALUES (" & ru!SourceKey & ", " & rt!EmployeeKey & ", " & ru!DeductionKey & ", " & _
                                          " '" & FormatDateTime(rt!PayrollDate, vbShortDate) & "', 'Payroll Deduction', " & _
                                          " 'O', " & CDbl(dAmt) & ", " & iPayrollKey & ", 2)"
                    End If
                End If
                rv.Close
                
                v = "SELECT tbl_Personnel_Payroll_Deductions.* " & _
                    " FROM tbl_Personnel_Payroll_Deductions " & _
                    " WHERE (MasterKey = " & iPayrollKey & ") " & _
                    " AND (DeductionKey = " & ru!DeductionKey & ")"
                If rv.State = adStateOpen Then rv.Close
                rv.Open v, ConnOmega
                If rv.RecordCount = 0 Then
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Payroll_Deductions " & _
                                      " (MasterKey, DeductionKey, Amount) " & _
                                      " VALUES (" & iPayrollKey & ", " & ru!DeductionKey & ", " & CDbl(dAmt) & ")"
                Else
                    ConnOmega.Execute "UPDATE tbl_Personnel_Payroll_Deductions " & _
                                      " SET Amount = Amount + " & CDbl(dAmt) & " " & _
                                      " WHERE (MasterKey = " & iPayrollKey & ") " & _
                                      " AND (DeductionKey = " & ru!DeductionKey & ")"
                End If
                rv.Close
            End If
            ru.MoveNext
        Wend
        ru.Close
        
        ConnOmega.Execute "UPDATE tbl_Personnel_Payroll " & _
                          " SET ActionMemoKey = " & rt!ActionMemoKey & ", " & _
                          " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                          " WHERE (PK = " & iPayrollKey & ")"
        
        ConnOmega.Execute "UPDATE tbl_Personnel_ActionNew SET Locked = 1 WHERE (PK = " & rt!ActionMemoKey & ")"
        
        ConnOmega.Execute "UPDATE tbl_Personnel_Hours SET PayrollKey = " & iPayrollKey & ", Posted = 1 WHERE (PK = " & rt!PK & ")"
        
    End If
End If
rt.Close
End Sub

Private Function Loan_Ded_Value(EmpKey, DedKey) As Double
Loan_Ded_Value = 0
'a = "SELECT TOP (1) Multiplier " & _
'    " From dbo.tbl_Personnel_Overtime_Multiplier " & _
'    " WHERE (EffectDate <= '" & FormatDateTime(dDateTo, vbShortDate) & "') " & _
'    " AND (EarningKey = " & EarningKey & ") " & _
'    " ORDER BY EffectDate DESC"
'If ra.State = adStateOpen Then ra.Close
'ra.Open a, ConnOmega
'If ra.RecordCount > 0 Then
'    Loan_Ded_Value = ra!Multiplier
'End If
ra.Close
End Function

