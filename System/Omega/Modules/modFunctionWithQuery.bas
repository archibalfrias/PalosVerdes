Attribute VB_Name = "modFunctionWithQuery"
Option Explicit

Public Function CheckIfPaid(dDate) As Boolean
CheckIfPaid = True
a = "SELECT TOP (1) tbl_System_PerfectHours. * " & _
    " FROM tbl_System_PerfectHours " & _
    " ORDER BY PK"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    If IsNull(ra!Payment) = True Then
        CheckIfPaid = True
    Else
        'MsgBox dDate & " : " & FormatDateTime(ra!Payment, vbShortDate)
        If DateValue(ra!Payment) < DateValue(dDate) Then
            CheckIfPaid = False
        Else
            CheckIfPaid = True
        End If
    End If
End If
ra.Close
End Function

Public Function Get_Period_Terms(iPeriodKey) As Long
Get_Period_Terms = 0
a = "SELECT Terms " & _
    " From dbo.tbl_Personnel_Compensation_Period " & _
    " WHERE (PK = " & iPeriodKey & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    Get_Period_Terms = ra!Terms
End If
ra.Close
End Function

Public Function Get_Loan_Paid(iLoanKey) As Double
Get_Loan_Paid = 0
a = "SELECT ROUND(SUM(Credit), 2) AS Amt " & _
    " From dbo.tbl_Personnel_Loans_SL " & _
    " WHERE (LoanKey = " & iLoanKey & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    Get_Loan_Paid = IIf(IsNull(ra!Amt), 0, ra!Amt)
End If
ra.Close
End Function

Public Function Earning_Multiplier(dDateTo, EarningKey, CompKey) As Double
Earning_Multiplier = 0
a = "SELECT TOP (1) Multiplier " & _
    " From dbo.tbl_Personnel_Earning_Multiplier " & _
    " WHERE (EffectDate <= '" & FormatDateTime(dDateTo, vbShortDate) & "') " & _
    " AND (EarningKey = " & EarningKey & ") " & _
    " AND (CompKey = " & CompKey & ") " & _
    " ORDER BY EffectDate DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    Earning_Multiplier = ra!Multiplier
End If
ra.Close
End Function

Public Function Overtime_Multiplier(dDateTo, EarningKey, CompKey) As Double
Overtime_Multiplier = 0
a = "SELECT TOP (1) Multiplier " & _
    " From dbo.tbl_Personnel_Overtime_Multiplier " & _
    " WHERE (EffectDate <= '" & FormatDateTime(dDateTo, vbShortDate) & "') " & _
    " AND (EarningKey = " & EarningKey & ") " & _
    " AND (CompKey = " & CompKey & ") " & _
    " ORDER BY EffectDate DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    Overtime_Multiplier = ra!Multiplier
End If
ra.Close
End Function

Public Function SSS_Cont_Value(dTotalAmt, dDateTo) As String
Dim iMasterKey
SSS_Cont_Value = "0|0|0"
iMasterKey = 0

a = "SELECT TOP (1) dbo.tbl_Govt_SSSTable.* " & _
    " FROM dbo.tbl_Govt_SSSTable " & _
    " WHERE (EffectDate <= '" & FormatDateTime(dDateTo, vbShortDate) & "') " & _
    " ORDER BY EffectDate DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    iMasterKey = ra!PK
End If
ra.Close

'a = "SELECT TOP (1) dbo.tbl_Govt_SSSTable_Details.Employee as EmployeeShare, " & _
    " dbo.tbl_Govt_SSSTable_Details.Employer as EmployerShare, " & _
    " dbo.tbl_Govt_SSSTable_Details.EC " & _
    " FROM dbo.tbl_Govt_SSSTable_Details LEFT OUTER JOIN " & _
    " dbo.tbl_Govt_SSSTable ON dbo.tbl_Govt_SSSTable_Details.MasterKey = dbo.tbl_Govt_SSSTable.PK " & _
    " WHERE (dbo.tbl_Govt_SSSTable.EffectDate <= '" & FormatDateTime(dDateTo, vbShortDate) & "') " & _
    " AND (dbo.tbl_Govt_SSSTable_Details.RangeFrom <= " & CDbl(dTotalAmt) & ") " & _
    " ORDER BY dbo.tbl_Govt_SSSTable.EffectDate DESC, dbo.tbl_Govt_SSSTable_Details.RangeFrom DESC"

a = "SELECT TOP (1) dbo.tbl_Govt_SSSTable_Details.* " & _
    " FROM dbo.tbl_Govt_SSSTable_Details " & _
    " WHERE (MasterKey = " & iMasterKey & ") " & _
    " AND (RangeFrom <= " & CDbl(dTotalAmt) & ") " & _
    " ORDER BY RangeFrom DESC "
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    SSS_Cont_Value = ra!Employee & "|" & _
                     ra!Employer & "|" & _
                     ra!EC
End If
ra.Close
End Function

Public Function PHIC_Cont_Value(dTotalAmt, dDateTo) As String
Dim dPagIbigContV2, dEmployee, dEmployer, iMasterKey
PHIC_Cont_Value = "0|0"
iMasterKey = 0
a = "SELECT TOP (1) tbl_Govt_PhilHealthTable.* " & _
    " FROM tbl_Govt_PhilHealthTable " & _
    " WHERE (EffectDate <= '" & FormatDateTime(dDateTo, vbShortDate) & "') " & _
    " ORDER BY EffectDate DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    iMasterKey = ra!PK
End If
ra.Close

a = "SELECT TOP (1) dbo.tbl_Govt_PhilHealthTable_Details.* " & _
    " FROM  dbo.tbl_Govt_PhilHealthTable_Details " & _
    " Where (MasterKey = " & iMasterKey & ") " & _
    " And (RangeFrom <= " & CDbl(dTotalAmt) & ") " & _
    " ORDER BY RangeFrom DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    If CDbl(ra!wPercent) = 1 Then
        dPagIbigContV2 = Format(CDbl(dTotalAmt) * (CDbl(ra!Percentage) / 100), "#0.00")
        dEmployee = Format(CDbl(dPagIbigContV2) / 2, "#0.00")
        dEmployer = Format(CDbl(dPagIbigContV2) / 2, "#0.00") 'CDbl(dPagIbigContV2) - CDbl(dEmployee)
        PHIC_Cont_Value = dEmployee & "|" & _
                          dEmployer & "|"
    Else
        PHIC_Cont_Value = ra!Employee & "|" & _
                          ra!Employer & "|"
    End If
End If
ra.Close
End Function


Public Function PagIbig_Cont_Value(dTotalAmt, dDateTo) As String
PagIbig_Cont_Value = "0|0"
a = "SELECT TOP (1) dbo.tbl_Govt_PagIbigTable_Details.Employee as EmployeeShare, " & _
    " dbo.tbl_Govt_PagIbigTable_Details.Employer as EmployerShare " & _
    " FROM dbo.tbl_Govt_PagIbigTable_Details LEFT OUTER JOIN " & _
    " dbo.tbl_Govt_PagIbigTable ON dbo.tbl_Govt_PagIbigTable_Details.MasterKey = dbo.tbl_Govt_PagIbigTable.PK " & _
    " WHERE (dbo.tbl_Govt_PagIbigTable.EffectDate <= '" & FormatDateTime(dDateTo, vbShortDate) & "') " & _
    " AND (dbo.tbl_Govt_PagIbigTable_Details.RangeFrom <= " & CDbl(dTotalAmt) & ") " & _
    " ORDER BY dbo.tbl_Govt_PagIbigTable.EffectDate DESC, dbo.tbl_Govt_PagIbigTable_Details.RangeFrom DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    PagIbig_Cont_Value = ra!EmployeeShare & "|" & _
                         ra!EmployerShare & "|"
End If
ra.Close
End Function

Public Function WithHolding_Cont_Value(Taxable, TaxCategory, TaxStatus, dDateTo) As Double
Dim dTaxDueTmp
WithHolding_Cont_Value = 0
a = "SELECT TOP (1) dbo.tbl_Govt_TaxTable_Det_Det_Det.CompLevel, " & _
    " dbo.tbl_Govt_TaxTable_Det_Det_Det.Constant, " & _
    " dbo.tbl_Govt_TaxTable_Det_Det_Det.Percentage " & _
    " FROM  dbo.tbl_Govt_TaxTable_Det_Det_Det LEFT OUTER JOIN " & _
    " dbo.tbl_Govt_TaxTable ON dbo.tbl_Govt_TaxTable_Det_Det_Det.MasterKey = dbo.tbl_Govt_TaxTable.PK " & _
    " WHERE (dbo.tbl_Govt_TaxTable_Det_Det_Det.TaxCategoryKey = " & TaxCategory & ") " & _
    " AND (dbo.tbl_Govt_TaxTable_Det_Det_Det.TaxStatusKey = " & TaxStatus & ") " & _
    " AND (dbo.tbl_Govt_TaxTable.EffectDate <= '" & FormatDateTime(dDateTo, vbShortDate) & "') " & _
    " AND (dbo.tbl_Govt_TaxTable_Det_Det_Det.CompLevel <= " & CDbl(Taxable) & ") " & _
    " ORDER BY dbo.tbl_Govt_TaxTable.EffectDate DESC, dbo.tbl_Govt_TaxTable_Det_Det_Det.CompLevel DESC, dbo.tbl_Govt_TaxTable_Det_Det_Det.Constant DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    dTaxDueTmp = CDbl(Taxable) - CDbl(ra!CompLevel)
    dTaxDueTmp = CDbl(dTaxDueTmp) * (CDbl(ra!Percentage) / 100)
    dTaxDueTmp = Format(CDbl(dTaxDueTmp) + CDbl(ra!Constant), "#,##0.00")
    WithHolding_Cont_Value = CDbl(dTaxDueTmp)
End If
ra.Close
End Function


Public Function CheckDeductToReference(iEarnKey) As Long
CheckDeductToReference = 0
a = "SELECT DedtoRef " & _
    " From dbo.tbl_Personnel_Payroll_Earnings_Table " & _
    " WHERE (PK = " & iEarnKey & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    CheckDeductToReference = ra!DedtoRef
End If
ra.Close
End Function

Public Function GetReferenceEarning(iEarnKey) As Long
GetReferenceEarning = 0
a = "SELECT RefRate " & _
    " From dbo.tbl_Personnel_Payroll_Earnings_Table " & _
    " WHERE (PK = " & iEarnKey & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    GetReferenceEarning = ra!RefRate
End If
ra.Close
End Function

Public Function CheckPerfectDays(iEmp, dEffectDate) As Long
CheckPerfectDays = 0
a = "SELECT TOP (1) dbo.tbl_Personnel_CompensationRate.CheckPerfectDays " & _
    " FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_CompensationRate ON dbo.tbl_Personnel_ActionNew.CompensationRateKey = dbo.tbl_Personnel_CompensationRate.PK " & _
    " WHERE (dbo.tbl_Personnel_ActionNew.EmpPK = " & iEmp & ") " & _
    " AND (dbo.tbl_Personnel_ActionNew.EffectivityDate <= '" & FormatDateTime(dEffectDate, vbShortDate) & "') " & _
    " ORDER BY dbo.tbl_Personnel_ActionNew.EffectivityDate DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    CheckPerfectDays = ra!CheckPerfectDays
End If
ra.Close
End Function

Public Function NET_OF_VAT(dEffectDate, dVATable, Optional iItemKey As Double = 0) As Double
a = "SELECT TOP 1 tbl_System_VAT.* " & _
    " FROM tbl_System_VAT " & _
    " WHERE (EffectDate <= '" & FormatDateTime(dEffectDate, vbShortDate) & "') " & _
    " ORDER BY EffectDate DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    If CDbl(iItemKey) <> 0 Then
        b = "SELECT tbl_Inv_Items.* " & _
            " FROM tbl_Inv_Items " & _
            " WHERE (PK = " & iItemKey & ")"
        If rb.State = adStateOpen Then rb.Close
        rb.Open b, ConnOmega
        If rb.RecordCount > 0 Then
            If rb!NonVAT = 1 Then
                NET_OF_VAT = dVATable
            Else
                NET_OF_VAT = CDbl(dVATable) / (1 + (ra!Vat / 100))
            End If
        Else
            NET_OF_VAT = CDbl(dVATable) / (1 + (ra!Vat / 100))
        End If
        rb.Close
    Else
        NET_OF_VAT = CDbl(dVATable) / (1 + (ra!Vat / 100))
    End If
Else
    NET_OF_VAT = dVATable
End If
ra.Close
End Function

Public Function DEDUCTION_TABLE(intPK) As String
a = "SELECT PK, Description" & _
    " From tbl_Personnel_ActionNew_DedTable  " & _
    " WHERE (PK = " & intPK & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If Not ra.EOF Then
    DEDUCTION_TABLE = ra!PK & ";" & _
                      ra!Description
Else
    DEDUCTION_TABLE = ""
End If
ra.Close
End Function

Public Function TAX_CATEGORY(intPK) As String
a = "SELECT PK, TaxCategory" & _
    " From tbl_Govt_TaxCategory  " & _
    " WHERE (PK = " & intPK & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If Not ra.EOF Then
    TAX_CATEGORY = ra!PK & ";" & _
                   ra!TaxCategory
Else
    TAX_CATEGORY = ""
End If
ra.Close
End Function

Public Function COMPENSATION_RATE(intPK) As String
a = "SELECT PK, Description" & _
    " From tbl_Personnel_CompensationRate  " & _
    " WHERE (PK = " & intPK & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If Not ra.EOF Then
    COMPENSATION_RATE = ra!PK & ";" & _
                        ra!Description
Else
    COMPENSATION_RATE = ""
End If
ra.Close
End Function

Public Function DIV_NAME(intPK) As String
a = "SELECT PK, Description" & _
    " From tbl_Personnel_Division  " & _
    " WHERE (PK = " & intPK & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If Not ra.EOF Then
    DIV_NAME = ra!PK & ";" & _
               ra!Description
Else
    DIV_NAME = ""
End If
ra.Close
End Function

Public Function DEPT_NAME(intPK) As String
a = "SELECT DepartmentCode, DepartmentName" & _
    " From tbl_Personnel_Department  " & _
    " WHERE (PK = " & intPK & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If Not ra.EOF Then
    DEPT_NAME = ra!DepartmentCode & ";" & _
               ra!DepartmentName
Else
    DEPT_NAME = ""
End If
ra.Close
End Function

Public Function EMP_STATUS(intPK) As String
a = "SELECT StatusCode, StatusName" & _
    " From tbl_Personnel_EmploymentStatus  " & _
    " WHERE (PK = " & intPK & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If Not ra.EOF Then
    EMP_STATUS = ra!StatusCode & ";" & _
                 ra!StatusName
Else
    EMP_STATUS = ""
End If
ra.Close
End Function

Public Function POSITION_NAME(intPK) As String
a = "SELECT PositionCode, PositionName" & _
    " From tbl_Personnel_Position  " & _
    " WHERE (PK = " & intPK & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If Not ra.EOF Then
    POSITION_NAME = ra!PositionCode & ";" & _
                    ra!PositionName
Else
    POSITION_NAME = ""
End If
ra.Close
End Function

Public Function TAX_STATUS_NAME(intPK) As String
a = "SELECT PK, TaxStatus" & _
    " From tbl_Govt_TaxStatus " & _
    " WHERE (PK=" & intPK & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If Not ra.EOF Then
    TAX_STATUS_NAME = ra!PK & ";" & ra!TaxStatus
Else
    TAX_STATUS_NAME = "0" & ";" & ""
End If
ra.Close
End Function

Public Function GET_LAST_ACTION_EFFECTIVITY(strEmpNo) As Date
a = "SELECT Max(EffectivityDate) AS MaxOfEffectivityDate" & _
    " From tbl_Personnel_Action  " & _
    " WHERE (EmpPK = " & strEmpNo & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    GET_LAST_ACTION_EFFECTIVITY = IIf(IsNull(ra!MaxOfEffectivityDate), "01/01/1900", ra!MaxOfEffectivityDate)
Else
    GET_LAST_ACTION_EFFECTIVITY = CDate("01/01/1900")
End If
ra.Close
End Function

Public Function GET_LAST_ACTION_EFFECTIVITY_NEW(strEmpNo) As Date
a = "SELECT Max(EffectivityDate) AS MaxOfEffectivityDate" & _
    " From tbl_Personnel_ActionNew  " & _
    " WHERE (EmpPK = " & strEmpNo & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    GET_LAST_ACTION_EFFECTIVITY_NEW = IIf(IsNull(ra!MaxOfEffectivityDate), "01/01/1900", ra!MaxOfEffectivityDate)
Else
    GET_LAST_ACTION_EFFECTIVITY_NEW = CDate("01/01/1900")
End If
ra.Close
End Function

Public Function CHECK_TOURNAMENT_STATUS(TourKey) As Long
CHECK_TOURNAMENT_STATUS = 0
a = "SELECT tbl_Scoring_TournamentInfo.* " & _
    " FROM tbl_Scoring_TournamentInfo " & _
    " WHERE (PK = " & TourKey & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    CHECK_TOURNAMENT_STATUS = ra!Locked
End If
ra.Close
End Function

Public Function OVERTIME_RATE() As Double
OVERTIME_RATE = 0
a = "SELECT TOP 1 OverTime" & _
    " From tbl_Personnel_OverTimeTable " & _
    " ORDER BY PK DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If Not ra.EOF Then
    OVERTIME_RATE = 1 + (ra!OverTime / 100)
End If
ra.Close
End Function

Public Function RESTDAY_RATE() As Double
RESTDAY_RATE = 0
a = "SELECT TOP 1 RestDay" & _
    " From tbl_Personnel_OverTimeTable " & _
    " ORDER BY PK DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If Not ra.EOF Then
    RESTDAY_RATE = 1 + (ra!RestDay / 100)
End If
ra.Close
End Function

Public Function GET_PREVIOUS_GROSS(intPeriod, intDivision, intEmployee) As Double
Dim PrePeriod
GET_PREVIOUS_GROSS = 0
PrePeriod = 0
a = "SELECT TOP 1 PK " & _
    " From tbl_Personnel_Compensation_Period " & _
    " WHERE (PK < " & intPeriod & ") " & _
    " AND (Type = " & intDivision & ") " & _
    " ORDER BY DateFrom DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    PrePeriod = ra!PK
End If
ra.Close

a = "SELECT TOP 1 TotalEarning" & _
    " From tbl_Personnel_Compensation " & _
    " Where (EmpPK = " & intEmployee & ") " & _
    " And (Period = " & PrePeriod & ") " & _
    " ORDER BY Period DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    GET_PREVIOUS_GROSS = ra!TotalEarning
End If
ra.Close
End Function

Public Function IS_HAVE_SSS(strNo, dtmEffectivityDate) As Boolean
a = "SELECT TOP 1 Is_SSS" & _
    " From tbl_Personnel_Action " & _
    " Where (EmpPK = " & strNo & ") " & _
    " And (EffectivityDate <= '" & FormatDateTime(dtmEffectivityDate, vbShortDate) & "') " & _
    " ORDER BY EffectivityDate DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If Not ra.EOF Then
    If ra!Is_SSS = 1 Then
        IS_HAVE_SSS = True
    Else
        IS_HAVE_SSS = False
    End If
Else
    IS_HAVE_SSS = False
End If
ra.Close
End Function

Public Function IS_HAVE_PHIC(strNo, dtmEffectivityDate) As Boolean
s = "SELECT TOP 1 Is_PHIC" & _
    " From tbl_Personnel_Action " & _
    " Where (EmpPK = " & strNo & ") " & _
    " And (EffectivityDate <= '" & FormatDateTime(dtmEffectivityDate, vbShortDate) & "') " & _
    " ORDER BY EffectivityDate DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    If ra!Is_PHIC = 1 Then
        IS_HAVE_PHIC = True
    Else
        IS_HAVE_PHIC = False
    End If
Else
    IS_HAVE_PHIC = False
End If
ra.Close
End Function

Public Function IS_HAVE_PagIbig(strNo, dtmEffectivityDate) As Boolean
s = "SELECT TOP 1 Is_PAGIBIG" & _
    " From tbl_Personnel_Action " & _
    " Where (EmpPK = " & strNo & ") " & _
    " And (EffectivityDate <= '" & FormatDateTime(dtmEffectivityDate, vbShortDate) & "') " & _
    " ORDER BY EffectivityDate DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    If ra!Is_PAGIBIG = 1 Then
        IS_HAVE_PagIbig = True
    Else
        IS_HAVE_PagIbig = False
    End If
Else
    IS_HAVE_PagIbig = False
End If
ra.Close
End Function

Public Function IS_HAVE_TIN(strNo, dtmEffectivityDate) As Boolean
s = "SELECT TOP 1 Is_TIN" & _
    " From tbl_Personnel_Action " & _
    " Where (EmpPK = " & strNo & ") " & _
    " And (EffectivityDate <= '" & FormatDateTime(dtmEffectivityDate, vbShortDate) & "') " & _
    " ORDER BY EffectivityDate DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    If ra!Is_TIN = 1 Then
        IS_HAVE_TIN = True
    Else
        IS_HAVE_TIN = False
    End If
Else
    IS_HAVE_TIN = False
End If
ra.Close
End Function

Public Function CHECK_CONT_CUTOFF(strCont, intDayFrom) As Boolean
CHECK_CONT_CUTOFF = False
s = "SELECT CutOff, Day, Division" & _
    " From tbl_Personnel_CutOff" & _
    " WHERE (CutOff = '" & strCont & "') " & _
    " AND (Day = " & intDayFrom & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    CHECK_CONT_CUTOFF = True
End If
ra.Close
End Function

Public Function GET_SSS_CONTRIBUTION_EMPLOYER(lngBasic, dEffectDate As Date) As Double
GET_SSS_CONTRIBUTION_EMPLOYER = 0
's = "SELECT EmployerShare" & _
    " From tbl_Personnel_SSSTable " & _
    " WHERE ([From] <= " & lngBasic & ") " & _
    " AND ([To] >= " & lngBasic & ")"
's = "SELECT TOP 1 EmployerShare, EmployeeShare, EC " & _
    " From dbo.tbl_Personnel_SSSTable " & _
    " Where ([From] <= " & CDbl(lngBasic) & ") " & _
    " ORDER BY [From] DESC"
s = "SELECT TOP 1 dbo.tbl_Govt_SSSTable_Details.Employee as EmployeeShare, " & _
    " dbo.tbl_Govt_SSSTable_Details.Employer as EmployerShare, " & _
    " dbo.tbl_Govt_SSSTable_Details.EC " & _
    " FROM dbo.tbl_Govt_SSSTable_Details LEFT OUTER JOIN " & _
    " dbo.tbl_Govt_SSSTable ON dbo.tbl_Govt_SSSTable_Details.MasterKey = dbo.tbl_Govt_SSSTable.PK " & _
    " WHERE (dbo.tbl_Govt_SSSTable.EffectDate <= '" & FormatDateTime(dEffectDate, vbShortDate) & "') " & _
    " AND (dbo.tbl_Govt_SSSTable_Details.RangeFrom <= " & CDbl(lngBasic) & ") " & _
    " ORDER BY dbo.tbl_Govt_SSSTable.EffectDate DESC, dbo.tbl_Govt_SSSTable_Details.RangeFrom DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    GET_SSS_CONTRIBUTION_EMPLOYER = ra!EmployerShare
End If
ra.Close
End Function

Public Function GET_SSS_CONTRIBUTION_EMPLOYEE(lngBasic, dEffectDate As Date) As Double
GET_SSS_CONTRIBUTION_EMPLOYEE = 0
's = "SELECT EmployeeShare" & _
    " From tbl_Personnel_SSSTable " & _
    " WHERE ([From] <= " & lngBasic & ") " & _
    " AND ([To] >= " & lngBasic & ")"
's = "SELECT TOP 1 EmployerShare, EmployeeShare, EC " & _
    " From dbo.tbl_Personnel_SSSTable " & _
    " Where ([From] <= " & CDbl(lngBasic) & ") " & _
    " ORDER BY [From] DESC"
s = "SELECT TOP 1 dbo.tbl_Govt_SSSTable_Details.Employee as EmployeeShare, " & _
    " dbo.tbl_Govt_SSSTable_Details.Employer as EmployerShare, " & _
    " dbo.tbl_Govt_SSSTable_Details.EC " & _
    " FROM dbo.tbl_Govt_SSSTable_Details LEFT OUTER JOIN " & _
    " dbo.tbl_Govt_SSSTable ON dbo.tbl_Govt_SSSTable_Details.MasterKey = dbo.tbl_Govt_SSSTable.PK " & _
    " WHERE (dbo.tbl_Govt_SSSTable.EffectDate <= '" & FormatDateTime(dEffectDate, vbShortDate) & "') " & _
    " AND (dbo.tbl_Govt_SSSTable_Details.RangeFrom <= " & CDbl(lngBasic) & ") " & _
    " ORDER BY dbo.tbl_Govt_SSSTable.EffectDate DESC, dbo.tbl_Govt_SSSTable_Details.RangeFrom DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    GET_SSS_CONTRIBUTION_EMPLOYEE = ra!EmployeeShare
End If
ra.Close
End Function

Public Function GET_SSS_CONTRIBUTION_EC(lngBasic, dEffectDate As Date) As Double
GET_SSS_CONTRIBUTION_EC = 0
's = "SELECT EC" & _
    " From tbl_Personnel_SSSTable " & _
    " WHERE ([From] <= " & lngBasic & ") " & _
    " AND ([To] >= " & lngBasic & ")"
's = "SELECT TOP 1 EmployerShare, EmployeeShare, EC " & _
    " From dbo.tbl_Personnel_SSSTable " & _
    " Where ([From] <= " & CDbl(lngBasic) & ") " & _
    " ORDER BY [From] DESC"
s = "SELECT TOP 1 dbo.tbl_Govt_SSSTable_Details.Employee as EmployeeShare, " & _
    " dbo.tbl_Govt_SSSTable_Details.Employer as EmployerShare, " & _
    " dbo.tbl_Govt_SSSTable_Details.EC " & _
    " FROM dbo.tbl_Govt_SSSTable_Details LEFT OUTER JOIN " & _
    " dbo.tbl_Govt_SSSTable ON dbo.tbl_Govt_SSSTable_Details.MasterKey = dbo.tbl_Govt_SSSTable.PK " & _
    " WHERE (dbo.tbl_Govt_SSSTable.EffectDate <= '" & FormatDateTime(dEffectDate, vbShortDate) & "') " & _
    " AND (dbo.tbl_Govt_SSSTable_Details.RangeFrom <= " & CDbl(lngBasic) & ") " & _
    " ORDER BY dbo.tbl_Govt_SSSTable.EffectDate DESC, dbo.tbl_Govt_SSSTable_Details.RangeFrom DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    GET_SSS_CONTRIBUTION_EC = ra!EC
End If
ra.Close
End Function

Public Function GET_PHIC_CONTRIBUTION_EMPLOYER(lngBasic, dEffectDate As Date) As Double
GET_PHIC_CONTRIBUTION_EMPLOYER = 0
's = "SELECT TOP 1 EmployerShare, ComputeBy " & _
    " From tbl_Personnel_PhilHealthTable " & _
    " WHERE ([From] <= " & lngBasic & ") " & _
    " AND ([To] >= " & lngBasic & ") " & _
    " AND (EffectDate <= '" & FormatDateTime(dEffectDate, vbShortDate) & "') " & _
    " ORDER BY EffectDate DESC"
s = "SELECT TOP 1 dbo.tbl_Govt_PhilHealthTable_Details.Employee as EmployeeShare, " & _
    " dbo.tbl_Govt_PhilHealthTable_Details.Employer as EmployerShare" & _
    " FROM dbo.tbl_Govt_PhilHealthTable_Details LEFT OUTER JOIN " & _
    " dbo.tbl_Govt_PhilHealthTable ON dbo.tbl_Govt_PhilHealthTable_Details.MasterKey = dbo.tbl_Govt_PhilHealthTable.PK " & _
    " WHERE (dbo.tbl_Govt_PhilHealthTable.EffectDate <= '" & FormatDateTime(dEffectDate, vbShortDate) & "') " & _
    " AND (dbo.tbl_Govt_PhilHealthTable_Details.RangeFrom <= " & CDbl(lngBasic) & ") " & _
    " ORDER BY dbo.tbl_Govt_PhilHealthTable.EffectDate DESC, dbo.tbl_Govt_PhilHealthTable_Details.RangeFrom DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
'    GET_PHIC_CONTRIBUTION_EMPLOYER = 0
'    If ra!ComputeBy = 1 Then
        GET_PHIC_CONTRIBUTION_EMPLOYER = CDbl(ra!EmployerShare)
'    ElseIf ra!ComputeBy = 2 Then
'        GET_PHIC_CONTRIBUTION_EMPLOYER = CDbl(Format((ra!EmployerShare / 100) * CDbl(lngBasic), "#,##0.00"))
'    End If
End If
ra.Close
End Function

Public Function GET_PHIC_CONTRIBUTION_EMPLOYEE(lngBasic, dEffectDate As Date) As Double
GET_PHIC_CONTRIBUTION_EMPLOYEE = 0
's = "SELECT TOP 1 EmployeeShare, ComputeBy " & _
    " From tbl_Personnel_PhilHealthTable " & _
    " WHERE ([From] <= " & lngBasic & ") " & _
    " AND ([To] >= " & lngBasic & ") " & _
    " AND (EffectDate <= '" & FormatDateTime(dEffectDate, vbShortDate) & "') " & _
    " ORDER BY EffectDate DESC"
s = "SELECT TOP 1 dbo.tbl_Govt_PhilHealthTable_Details.Employee as EmployeeShare, " & _
    " dbo.tbl_Govt_PhilHealthTable_Details.Employer as EmployerShare" & _
    " FROM dbo.tbl_Govt_PhilHealthTable_Details LEFT OUTER JOIN " & _
    " dbo.tbl_Govt_PhilHealthTable ON dbo.tbl_Govt_PhilHealthTable_Details.MasterKey = dbo.tbl_Govt_PhilHealthTable.PK " & _
    " WHERE (dbo.tbl_Govt_PhilHealthTable.EffectDate <= '" & FormatDateTime(dEffectDate, vbShortDate) & "') " & _
    " AND (dbo.tbl_Govt_PhilHealthTable_Details.RangeFrom <= " & CDbl(lngBasic) & ") " & _
    " ORDER BY dbo.tbl_Govt_PhilHealthTable.EffectDate DESC, dbo.tbl_Govt_PhilHealthTable_Details.RangeFrom DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
'   GET_PHIC_CONTRIBUTION_EMPLOYEE = 0
'    If ra!ComputeBy = 1 Then
        GET_PHIC_CONTRIBUTION_EMPLOYEE = CDbl(ra!EmployeeShare)
'    ElseIf ra!ComputeBy = 2 Then
'        GET_PHIC_CONTRIBUTION_EMPLOYEE = CDbl(Format((ra!EmployeeShare / 100) * CDbl(lngBasic), "#,##0.00"))
'    End If
End If
ra.Close
End Function

Public Function GET_PAGIBIG_CONTRIBUTION_EMPLOYER(lngBasic) As Double
Dim varPagibigMaximum As Double
Dim varPercentage As Double
Dim varPagibigEmployer As Double
Dim varPagibigEmployee As Double
Dim varNewPagEmployer As Double
Dim varNewPagEmployee As Double
s = "SELECT Maximum, Percentage, EmployerShare, EmployeeShare" & _
    " FROM tbl_Personnel_PagIbigTable"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    varPagibigMaximum = ra!Maximum
    varPercentage = ra!Percentage
    varPagibigEmployer = ra!EmployerShare
    varPagibigEmployee = ra!EmployeeShare
End If
If lngBasic < varPagibigMaximum Then
    varNewPagEmployer = lngBasic * (varPercentage / 100)
    varNewPagEmployee = lngBasic * (varPercentage / 100)
Else
    varNewPagEmployer = varPagibigEmployer
    varNewPagEmployee = varPagibigEmployee
End If
GET_PAGIBIG_CONTRIBUTION_EMPLOYER = varNewPagEmployer
ra.Close
End Function

Public Function GET_PAGIBIG_CONTRIBUTION_EMPLOYEE(lngBasic) As Double
Dim varPagibigMaximum As Double
Dim varPercentage As Double
Dim varPagibigEmployer As Double
Dim varPagibigEmployee As Double
Dim varNewPagEmployer As Double
Dim varNewPagEmployee As Double
s = "SELECT Maximum, Percentage, EmployerShare, EmployeeShare" & _
    " FROM tbl_Personnel_PagIbigTable"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    varPagibigMaximum = ra!Maximum
    varPercentage = ra!Percentage
    varPagibigEmployer = ra!EmployerShare
    varPagibigEmployee = ra!EmployeeShare
End If
If lngBasic < varPagibigMaximum Then
    varNewPagEmployer = lngBasic * (varPercentage / 100)
    varNewPagEmployee = lngBasic * (varPercentage / 100)
Else
    varNewPagEmployer = varPagibigEmployer
    varNewPagEmployee = varPagibigEmployee
End If
GET_PAGIBIG_CONTRIBUTION_EMPLOYEE = varNewPagEmployee
ra.Close
End Function

Public Function GET_CURRENT_STATUS(lngIDNo, dtmDate) As Long
GET_CURRENT_STATUS = 0
s = "SELECT TOP 1 TaxStatus" & _
    " From tbl_Personnel_Action " & _
    " Where (EmpPK = " & lngIDNo & ") " & _
    " And (EffectivityDate <= '" & FormatDateTime(dtmDate, vbShortDate) & "')" & _
    " ORDER BY EffectivityDate DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    GET_CURRENT_STATUS = ra!TaxStatus
End If
ra.Close
End Function

Public Function COMPUTE_MONTHLY_TAX_EXEMP(lngTaxStatus, dblBracket, dtmDate) As Double
COMPUTE_MONTHLY_TAX_EXEMP = 0
s = "SELECT TOP 1 TaxExemption" & _
    " From tbl_Personnel_TaxTableMonthly  " & _
    " Where (TaxStatus = " & lngTaxStatus & ") " & _
    " And (Effectivity <= '" & dtmDate & "') " & _
    " And (BracketAmount <= " & CDbl(dblBracket) & ") " & _
    " ORDER BY Effectivity DESC, BracketAmount DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    COMPUTE_MONTHLY_TAX_EXEMP = ra!TaxExemption
End If
ra.Close
End Function

Public Function COMPUTE_MONTHLY_TAX_PERCENT(lngTaxStatus, dblBracket, dtmDate) As Double
COMPUTE_MONTHLY_TAX_PERCENT = 0
s = "SELECT TOP 1 PPercent" & _
    " From tbl_Personnel_TaxTableMonthly " & _
    " Where (TaxStatus = " & lngTaxStatus & ") " & _
    " And (Effectivity <= '" & dtmDate & "') " & _
    " And (BracketAmount <= " & CDbl(dblBracket) & ") " & _
    " ORDER BY Effectivity DESC, BracketAmount DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    COMPUTE_MONTHLY_TAX_PERCENT = ra!PPercent
End If
ra.Close
End Function

Public Function COMPUTE_MONTHLY_TAX_CONSTANT(lngTaxStatus, dblBracket, dtmDate) As Double
COMPUTE_MONTHLY_TAX_CONSTANT = 0
s = "SELECT TOP 1 CConstant" & _
    " From tbl_Personnel_TaxTableMonthly " & _
    " Where (TaxStatus = " & lngTaxStatus & ") " & _
    " And (Effectivity <= '" & dtmDate & "') " & _
    " And (BracketAmount <= " & CDbl(dblBracket) & ") " & _
    " ORDER BY Effectivity DESC, BracketAmount DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    COMPUTE_MONTHLY_TAX_CONSTANT = ra!CConstant
End If
ra.Close
End Function

Public Function COMPUTE_MONTHLY_TAX_BRACKET_AMOUNT(lngTaxStatus, dblBracket, dtmDate) As Double
COMPUTE_MONTHLY_TAX_BRACKET_AMOUNT = 0
s = "SELECT TOP 1 BracketAmount" & _
    " From tbl_Personnel_TaxTableMonthly " & _
    " Where (TaxStatus = " & lngTaxStatus & ") " & _
    " And (Effectivity <= '" & dtmDate & "') " & _
    " And (BracketAmount <= " & CDbl(dblBracket) & ")" & _
    " ORDER BY Effectivity DESC, BracketAmount DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    COMPUTE_MONTHLY_TAX_BRACKET_AMOUNT = ra!BracketAmount
End If
ra.Close
End Function

Public Function IS_HAVE_SSS_LOAN(varID, varDate) As Boolean
IS_HAVE_SSS_LOAN = False
s = "SELECT EmpPK, DateFrom, " & _
    " DateTo , LoanType " & _
    " From tbl_Personnel_Loans " & _
    " WHERE (EmpPK = " & varID & ") " & _
    " AND (DateFrom <= '" & FormatDateTime(varDate, vbShortDate) & "') " & _
    " AND (DateTo >= '" & FormatDateTime(varDate, vbShortDate) & "') " & _
    " AND (LoanType = 1)" & _
    " AND (ZeroOut = 0)"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    IS_HAVE_SSS_LOAN = True
End If
ra.Close
End Function

Public Function IS_HAVE_PAGIBIG_LOAN(varID, varDate) As Boolean
IS_HAVE_PAGIBIG_LOAN = False
s = "SELECT EmpPK, DateFrom, " & _
    " DateTo , LoanType " & _
    " From tbl_Personnel_Loans " & _
    " WHERE (EmpPK = " & varID & ") " & _
    " AND (DateFrom <= '" & FormatDateTime(varDate, vbShortDate) & "') " & _
    " AND (DateTo >= '" & FormatDateTime(varDate, vbShortDate) & "') " & _
    " AND (LoanType = 2)" & _
    " AND (ZeroOut = 0)"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    IS_HAVE_PAGIBIG_LOAN = True
End If
ra.Close
End Function

Public Function CHECK_LOAN_CUTOFF(strLoan, intDayFrom) As Boolean
s = "SELECT CutOff, Day, Division" & _
    " From tbl_Personnel_CutOff" & _
    " WHERE (CutOff = '" & strLoan & "') " & _
    " AND (Day = " & intDayFrom & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    CHECK_LOAN_CUTOFF = True
End If
ra.Close
End Function

Public Function GET_SSS_LOAN_NO(varIDNO, varDate) As Long
GET_SSS_LOAN_NO = 0
s = "SELECT PK" & _
    " From tbl_Personnel_Loans " & _
    " WHERE (LoanType = 1 ) " & _
    " AND (EmpPK = " & varIDNO & ") " & _
    " AND (DateFrom <= '" & FormatDateTime(varDate, vbShortDate) & "')" & _
    " AND (DateTo >= '" & FormatDateTime(varDate, vbShortDate) & "')" & _
    " AND (ZeroOut = 0)"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    GET_SSS_LOAN_NO = ra!PK
End If
ra.Close
End Function

Public Function GET_PAGIBIG_LOAN_NO(varIDNO, varDate) As Long
GET_PAGIBIG_LOAN_NO = 0
s = "SELECT PK" & _
    " From tbl_Personnel_Loans " & _
    " WHERE (LoanType = 2 ) " & _
    " AND (EmpPK = " & varIDNO & ") " & _
    " AND (DateFrom <= '" & FormatDateTime(varDate, vbShortDate) & "')" & _
    " AND (DateTo >= '" & FormatDateTime(varDate, vbShortDate) & "')" & _
    " AND (ZeroOut = 0)"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    GET_PAGIBIG_LOAN_NO = ra!PK
End If
ra.Close
End Function

Public Function GET_LOAN_INFO(lngNo) As String
s = "SELECT Amortization, DateFrom, DateTo, " & _
    " TotalAmount " & _
    " From tbl_Personnel_Loans " & _
    " WHERE (PK= " & lngNo & " )"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    GET_LOAN_INFO = ra!Amortization & ";" & ra!DateFrom & ";" & ra!DateTo & ";" & ra!TotalAmount
Else
    GET_LOAN_INFO = 0 & ";" & "1/1/1900" & ";" & "1/1/1900" & ";" & 0
End If
ra.Close
End Function

Public Function GET_TOTAL_PAID_SSS(lngNo, intPeriod) As Double
GET_TOTAL_PAID_SSS = 0
s = "SELECT Sum(SSSLoan) AS SumOfDSSSLoan " & _
    " From tbl_Personnel_Compensation " & _
    " WHERE (SSSLoan_No = " & lngNo & ") " & _
    " AND (Period < " & intPeriod & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
   GET_TOTAL_PAID_SSS = IIf(IsNull(ra!SumOfDSSSLoan), 0, ra!SumOfDSSSLoan)
End If
ra.Close
End Function

Public Function GET_TOTAL_PAID_PAGIBIG(lngNo, intPeriod) As Double
GET_TOTAL_PAID_PAGIBIG = 0
s = "SELECT Sum(PagIbigLoan) AS SumOfPagIbigLoan" & _
    " From tbl_Personnel_Compensation " & _
    " WHERE (PagIbigLoan_No = " & lngNo & ")" & _
    " AND (Period < " & intPeriod & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
   GET_TOTAL_PAID_PAGIBIG = IIf(IsNull(ra!SumOfPagIbigLoan), 0, ra!SumOfPagIbigLoan)
End If
ra.Close
End Function

Public Function CHECK_COMPENSATION_LOCKED(iDivision, iPeriod) As Boolean
CHECK_COMPENSATION_LOCKED = False
s = "SELECT TOP 1 Locked " & _
    " From tbl_Personnel_Compensation " & _
    " WHERE (Division = " & iDivision & ") " & _
    " AND (Period = " & iPeriod & ") " & _
    " AND (Locked = 1)"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If ra.RecordCount > 0 Then
    CHECK_COMPENSATION_LOCKED = True
End If
ra.Close
End Function

Public Function CHECK_IF_HAVE_ACTION(strEmpNo) As Long
CHECK_IF_HAVE_ACTION = 0
s = "SELECT Count(tbl_Personnel_Action.PK) AS CountOfPK" & _
    " From tbl_Personnel_Action  " & _
    " WHERE (tbl_Personnel_Action.EmpPK = " & strEmpNo & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    CHECK_IF_HAVE_ACTION = ra!CountOfPK
End If
ra.Close
End Function

Public Function GET_PERIOD(dtmFrom, dtmTo, intDiv) As Long
GET_PERIOD = 0
s = "SELECT PK" & _
    " From tbl_Personnel_Compensation_Period " & _
    " WHERE (DateFrom = '" & CDate(dtmFrom) & "') " & _
    " AND (DateTo = '" & CDate(dtmTo) & "') " & _
    " AND (Type = " & intDiv & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    GET_PERIOD = ra!PK
End If
ra.Close
End Function


Public Function GET_PERIOD_V2(dtmPayrollDate, intDiv) As Long
GET_PERIOD_V2 = 0
s = "SELECT PK" & _
    " From tbl_Personnel_Compensation_Period " & _
    " WHERE (PayrollDate = '" & FormatDateTime(dtmPayrollDate, vbShortDate) & "') " & _
    " AND (Type = " & intDiv & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    GET_PERIOD_V2 = ra!PK
End If
ra.Close
End Function

Public Function GET_PERIOD_CUTOFF(PeriodKey) As String
GET_PERIOD_CUTOFF = " - "
s = "SELECT DateFrom, DateTo " & _
    " From tbl_Personnel_Compensation_Period " & _
    " WHERE (PK = " & PeriodKey & ") "
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    GET_PERIOD_CUTOFF = Format(ra!DateFrom, "mm/dd/yyyy") & " - " & Format(ra!DateTo, "mm/dd/yyyy")
End If
ra.Close
End Function


Public Function GET_EMPLOYMENT_STATUS(iEmpPK, Effectdate As Date) As Long
GET_EMPLOYMENT_STATUS = 0
a = "SELECT TOP (1) dbo.tbl_Personnel_EmploymentStatus.Active " & _
    " FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_EmploymentStatus ON dbo.tbl_Personnel_ActionNew.EmpStatusKey = dbo.tbl_Personnel_EmploymentStatus.PK " & _
    " Where (dbo.tbl_Personnel_ActionNew.EmpPK = " & iEmpPK & ") " & _
    " AND (dbo.tbl_Personnel_ActionNew.EffectivityDate <= '" & FormatDateTime(Effectdate, vbShortDate) & "') " & _
    " ORDER BY dbo.tbl_Personnel_ActionNew.EffectivityDate DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    GET_EMPLOYMENT_STATUS = ra!Active
End If
ra.Close
End Function

Public Function GET_DIVISION(strEmpNo, Effectdate) As Long
GET_DIVISION = 0
's = "SELECT Division" & _
    " From tbl_PersonnelProfile " & _
    " WHERE (PK = " & strEmpNo & ")"
s = "SELECT Division " & _
    " FROM tbl_Personnel_Action " & _
    " WHERE (EmpPK = " & strEmpNo & ") " & _
    " AND (EffectivityDate <= '" & FormatDateTime(Effectdate, vbShortDate) & "') " & _
    " ORDER BY EffectivityDate DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    GET_DIVISION = ra!Division
End If
ra.Close
End Function

Public Function GET_DIVISION_V2(strEmpNo, Effectdate) As Long
GET_DIVISION_V2 = 0
's = "SELECT Division" & _
    " From tbl_PersonnelProfile " & _
    " WHERE (PK = " & strEmpNo & ")"
s = "SELECT DivisionKey " & _
    " FROM tbl_Personnel_ActionNew " & _
    " WHERE (EmpPK = " & strEmpNo & ") " & _
    " AND (EffectivityDate <= '" & FormatDateTime(Effectdate, vbShortDate) & "') " & _
    " ORDER BY EffectivityDate DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    GET_DIVISION_V2 = ra!DivisionKey
End If
ra.Close
End Function

Public Function GET_EMPLOYEE_INFO(dtmDate, strEmpNo) As String
s = "SELECT TOP 1 tbl_Personnel_Action.Division, tbl_Personnel_Action.Dept, " & _
    " tbl_Personnel_Department.DepartmentName, tbl_Personnel_Action.EmpStatus, " & _
    " tbl_Personnel_EmploymentStatus.StatusName, tbl_Personnel_Action.Positions, " & _
    " tbl_Personnel_Position.PositionName, tbl_Personnel_Action.RatePerHourBasic, " & _
    " tbl_Personnel_IDNumber.IDNumber, tbl_Personnel_IDNumber.ProfileKey, " & _
    " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS Name, " & _
    " tbl_Personnel_Action.PK, tbl_Personnel_Action.Basic, tbl_Personnel_Action.RatePerHourCola, " & _
    " tbl_Personnel_Action.RatePerHourAllow " & _
    " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
    " tbl_Personnel_EmploymentStatus ON tbl_Personnel_Action.EmpStatus = tbl_Personnel_EmploymentStatus.PK LEFT OUTER JOIN " & _
    " tbl_Personnel_Department ON tbl_Personnel_Action.Dept = tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
    " tbl_Personnel_Position ON tbl_Personnel_Action.Positions = tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
    " tbl_Personnel_IDNumber ON " & _
    " tbl_Personnel_Action.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
    " WHERE (tbl_Personnel_Action.EmpPK = " & strEmpNo & ") " & _
    " AND (tbl_Personnel_Action.EffectivityDate <= '" & FormatDateTime(dtmDate, vbShortDate) & "') " & _
    " ORDER BY tbl_Personnel_Action.EffectivityDate DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    GET_EMPLOYEE_INFO = ra!Division & ";" & ra!Dept & ";" & _
                        ra!DepartmentName & ";" & ra!EmpStatus & ";" & _
                        ra!Positions & ";" & ra!PositionName & ";" & _
                        ra!RatePerHourBasic & ";" & ra!IDNumber & ";" & _
                        ra!Name & ";" & ra!PK & ";" & ra!Basic & ";" & _
                        ra!RatePerHourCola & ";" & ra!RatePerHourAllow & ";" & _
                        ra!ProfileKey
End If
ra.Close
End Function

Public Function FIND_PAYROLL_PERIOD(dtmDate, intType) As String
'If intType = 1 Then
'    s = "SELECT TOP 1 qry_Period_Club.PK, qry_Period_Club.DateFrom, " & _
'        " qry_Period_Club.DateTo, qry_Period_Club.Terms" & _
'        " From qry_Period_Club " & _
'        " WHERE (((qry_Period_Club.DateFrom)<(SELECT qry_Period_Club.DateFrom FROM qry_Period_Club " & _
'        " WHERE (((qry_Period_Club.DateFrom)<='" & CDate(dtmDate) & "') " & _
'        " AND ((qry_Period_Club.DateTo)>='" & CDate(dtmDate) & "')))))" & _
'        " ORDER BY qry_Period_Club.DateFrom DESC"
'ElseIf intType = 2 Then
'    s = "SELECT TOP 1 qry_Period_Main.PK, qry_Period_Main.DateFrom, " & _
'        " qry_Period_Main.DateTo, qry_Period_Main.Terms" & _
'        " From qry_Period_Main " & _
'        " WHERE (((qry_Period_Main.DateFrom)<(SELECT qry_Period_Main.DateFrom FROM qry_Period_Main " & _
'        " WHERE (((qry_Period_Main.DateFrom)<='" & CDate(dtmDate) & "') " & _
'        " AND ((qry_Period_Main.DateTo)>='" & CDate(dtmDate) & "')))))" & _
'        " ORDER BY qry_Period_Main.DateFrom DESC"
'End If
s = "sp_Compensation_Period('" & FormatDateTime(dtmDate, vbShortDate) & "', " & intType & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    FIND_PAYROLL_PERIOD = ra!PK & ";" & ra!DateFrom & ";" & ra!DateTo & ";" & ra!Terms
End If
ra.Close
End Function

Public Function GET_TERMS(intPK) As Long
GET_TERMS = 0
s = "SELECT Terms" & _
    " From tbl_Personnel_Compensation_Period " & _
    " WHERE (PK = " & intPK & ")"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If Not ra.EOF Then
    GET_TERMS = ra!Terms
End If
ra.Close
End Function

Public Function IMAGEFILESIZE(dEffectDate As Date) As Double
IMAGEFILESIZE = 100
s = "SELECT TOP 1 tbl_ImageFileSize.* " & _
    " FROM tbl_ImageFileSize " & _
    " WHERE (EffectDate <= '" & FormatDateTime(dEffectDate, vbShortDate) & "') " & _
    " ORDER BY EffectDate DESC"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If ra.RecordCount > 0 Then
    IMAGEFILESIZE = ra!FileSize
End If
ra.Close
End Function

Public Function isWithMortuary(EmploymentKey) As Long
isWithMortuary = 0
s = "SELECT tbl_Personnel_EmploymentStatus.* " & _
    " FROM tbl_Personnel_EmploymentStatus " & _
    " WHERE (PK = " & EmploymentKey & ") "
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If ra.RecordCount > 0 Then
    isWithMortuary = ra!WithMortuary
End If
ra.Close
End Function

Public Function PositionLevel(PositionKey) As Long
PositionLevel = 0
s = "SELECT tbl_Personnel_Position.* " & _
    " FROM tbl_Personnel_Position " & _
    " WHERE (PK = " & PositionKey & ") "
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If ra.RecordCount > 0 Then
    PositionLevel = ra!PositionLevel
End If
ra.Close
End Function

Public Function POPULATE_COMBO(strNo, strName, strTable, strOrder, cmb As ComboBox)
'Dim s As String
'Dim rs As New ADODB.Recordset
a = "SELECT " & strNo & "," & strName & " FROM " & strTable & " ORDER BY " & strOrder & ""
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
With cmb
    .Clear
    While Not ra.EOF
        .AddItem UCase(ra(strName))
        .ItemData(.NewIndex) = ra(strNo)
        ra.MoveNext
    Wend
End With
ra.Close
End Function


Public Function POPULATE_COMBO_EXEMPTION(strNo, strName, strTable, strOrder, strField, strFilter, cmb As ComboBox)
'Dim s As String
'Dim rs As New ADODB.Recordset
a = "SELECT " & strNo & "," & strName & " FROM " & strTable & " WHERE (" & strField & " = " & strFilter & ") ORDER BY " & strOrder & ""
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
With cmb
    .Clear
    While Not ra.EOF
        .AddItem UCase(ra(strName))
        .ItemData(.NewIndex) = ra(strNo)
        ra.MoveNext
    Wend
End With
ra.Close
End Function
