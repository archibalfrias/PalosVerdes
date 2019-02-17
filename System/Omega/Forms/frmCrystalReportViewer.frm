VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCrystalReportViewer 
   BackColor       =   &H00C6B8A4&
   Caption         =   "Preview"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCrystalReportViewer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7515
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCrystalReportViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function PRINT_Employee_OutStandingBal(sUser)
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptEmployeeDedOutStanding.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(sUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_LOAN_Ledger(sUser)
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptLoanLedger.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(sUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_PAYROLL_Contribution(sUser)
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptContributions.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(sUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_PAYROLL_Loans(sUser)
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptLoans.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(sUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_PAYROLL_PAYSLIP_V3(sUser)
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptPayslipV3.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(sUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_PAYROLL_PAYSLIP(sUser)
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptPayslipV2.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(sUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_CHECK_VOUCHER_PREPRINTED(sUser)
Screen.MousePointer = vbHourglass
Dim Report As New rptCheckVoucher_PrePrinted
Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Report.DiscardSavedData
Report.ParameterFields(1).ClearCurrentValueAndRange
Report.ParameterFields(1).AddCurrentValue CStr(sUser)
CRViewer1.DisplayGroupTree = False
CRViewer1.EnableExportButton = True
CRViewer1.ReportSource = Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_CHECK_VOUCHER(sUser)
Screen.MousePointer = vbHourglass
Dim Report As New rptCheckVoucher
Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Report.DiscardSavedData
Report.ParameterFields(1).ClearCurrentValueAndRange
Report.ParameterFields(1).AddCurrentValue CStr(sUser)
CRViewer1.DisplayGroupTree = False
CRViewer1.EnableExportButton = True
CRViewer1.ReportSource = Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function


Public Function PRINT_PERSONNAL_DATASHEET(sUser)
Screen.MousePointer = vbHourglass
'Dim Report As New rptPersonnalDataSheet
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CStr(sUser)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport

C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptPersonalDataSheet.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(sUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport

Screen.MousePointer = vbDefault
End Function

Public Function PRINT_ALLOWANCE_REPORT(sCompany, iDivision, sPeriod, sUser)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptAllowanceReport
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CStr(sCompany)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CLng(iDivision)
'Report.ParameterFields(3).ClearCurrentValueAndRange
'Report.ParameterFields(3).AddCurrentValue CStr(sPeriod)
'Report.ParameterFields(4).ClearCurrentValueAndRange
'Report.ParameterFields(4).AddCurrentValue CStr(sUser)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptAllowanceReport.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(sCompany)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CLng(iDivision)
C_Report.ParameterFields(3).ClearCurrentValueAndRange
C_Report.ParameterFields(3).AddCurrentValue CStr(sPeriod)
C_Report.ParameterFields(4).ClearCurrentValueAndRange
C_Report.ParameterFields(4).AddCurrentValue CStr(sUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_ACTIVE_EMPLOYEE(strCompany, strUser)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptActiveEmployee
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CStr(strUser)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptActiveEmployee.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_INACTIVE_EMPLOYEE(strCompany, strUser)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptInactiveEmployee
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CStr(strUser)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptInactiveEmployee.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_EMPLOYEE_HEADCOUNT(strCompany, strUser)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptPersonnelHeadCount
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CStr(strUser)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptPersonnelHeadCount.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_13TH_MONTH_V2(strUser, iMonth)
'Screen.MousePointer = vbHourglass
'Dim Report As New rpt13thMonth
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CDbl(iDivision)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CDbl(iMonth)
'Report.ParameterFields(3).ClearCurrentValueAndRange
'Report.ParameterFields(3).AddCurrentValue CStr(strHeader)
'Report.ParameterFields(4).ClearCurrentValueAndRange
'Report.ParameterFields(4).AddCurrentValue CStr(strUser)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rpt13thMonthV2.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strUser)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CDbl(iMonth)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_13TH_MONTH(iDivision, iMonth, strHeader, strUser)
'Screen.MousePointer = vbHourglass
'Dim Report As New rpt13thMonth
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CDbl(iDivision)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CDbl(iMonth)
'Report.ParameterFields(3).ClearCurrentValueAndRange
'Report.ParameterFields(3).AddCurrentValue CStr(strHeader)
'Report.ParameterFields(4).ClearCurrentValueAndRange
'Report.ParameterFields(4).AddCurrentValue CStr(strUser)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rpt13thMonth.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CDbl(iDivision)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CDbl(iMonth)
C_Report.ParameterFields(3).ClearCurrentValueAndRange
C_Report.ParameterFields(3).AddCurrentValue CStr(strHeader)
C_Report.ParameterFields(4).ClearCurrentValueAndRange
C_Report.ParameterFields(4).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_SCORE_BEST_BALL(strUser, strHeader)
Screen.MousePointer = vbHourglass
Dim Report As New rptBestBall
Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Report.DiscardSavedData
Report.ParameterFields(1).ClearCurrentValueAndRange
Report.ParameterFields(1).AddCurrentValue CStr(strUser)
Report.ParameterFields(2).ClearCurrentValueAndRange
Report.ParameterFields(2).AddCurrentValue CStr(strHeader)
CRViewer1.DisplayGroupTree = False
CRViewer1.EnableExportButton = True
CRViewer1.ReportSource = Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_ACTION_MEMO()
'Screen.MousePointer = vbHourglass
'Dim Report As New rptActionMemo
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptActionMemo.rpt", 1)
C_Report.DiscardSavedData
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_ACTION_MEMO_V2(strUserName)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptActionMemo
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptActionMemoV2.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strUserName)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_SIGNATURE_LEDGER(strCompany, strUser)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptSignatureLedger
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CStr(strUser)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptSignatureLedger.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_SIGNATURE_LEDGER_V2(strUser)
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptSignatureLedgerV2.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_COLA_SUMMARY(strCompany, strUser)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptColaSummary
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CStr(strUser)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptColaSummary.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_ALLOWANCE_SUMMARY(strCompany, strUser)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptAllowanceSummary
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CStr(strUser)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptAllowanceSummary.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function



Public Function PRINT_PAYSLIP_SUPERVISORY(strCompany, strUser, intLevel)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptPaySlip_Supervisory
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CStr(strUser)
'Report.ParameterFields(3).ClearCurrentValueAndRange
'Report.ParameterFields(3).AddCurrentValue CInt(intLevel)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptPaySlip_Supervisory.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CStr(strUser)
C_Report.ParameterFields(3).ClearCurrentValueAndRange
C_Report.ParameterFields(3).AddCurrentValue CInt(intLevel)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_PAYSLIP_DEPT(strCompany, strUser, intDept)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptPaySlip_Dept
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CStr(strUser)
'Report.ParameterFields(3).ClearCurrentValueAndRange
'Report.ParameterFields(3).AddCurrentValue CInt(intDept)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptPaySlip_Dept.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CStr(strUser)
C_Report.ParameterFields(3).ClearCurrentValueAndRange
C_Report.ParameterFields(3).AddCurrentValue CInt(intDept)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_PAYSLIP_STATUS(strCompany, strUser, intStatus)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptPaySlip_Status
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CStr(strUser)
'Report.ParameterFields(3).ClearCurrentValueAndRange
'Report.ParameterFields(3).AddCurrentValue CInt(intStatus)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptPaySlip_Status.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CStr(strUser)
C_Report.ParameterFields(3).ClearCurrentValueAndRange
C_Report.ParameterFields(3).AddCurrentValue CInt(intStatus)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_PAYSLIP_POST(strCompany, strUser, intPost)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptPaySlip_Post
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CStr(strUser)
'Report.ParameterFields(3).ClearCurrentValueAndRange
'Report.ParameterFields(3).AddCurrentValue CInt(intPost)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptPaySlip_Post.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CStr(strUser)
C_Report.ParameterFields(3).ClearCurrentValueAndRange
C_Report.ParameterFields(3).AddCurrentValue CInt(intPost)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_PAYSLIP_EMPLOYEE(strCompany, strUser, intEmpPK)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptPaySlip_Employee
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CStr(strUser)
'Report.ParameterFields(3).ClearCurrentValueAndRange
'Report.ParameterFields(3).AddCurrentValue CDbl(intEmpPK)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptPaySlip_Employee.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CStr(strUser)
C_Report.ParameterFields(3).ClearCurrentValueAndRange
C_Report.ParameterFields(3).AddCurrentValue CDbl(intEmpPK)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_COMPENSATION_SUMMARY_TOP(strCompany, strUser)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptCompensationSummaryTop
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CStr(strUser)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptCompensationSummaryTop.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_COMPENSATION_SUMMARY_V3(strUser)
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
'Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptPayrollRegister.rpt", 1)
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptPayrollRegisterV3.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_COMPENSATION_SUMMARY_V2(strUser)
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
'Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptPayrollRegister.rpt", 1)
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptPayrollRegisterV2.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function


Public Function PRINT_COMPENSATION_SUMMARY(strCompany, strUser)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptCompensationSummary
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CStr(strUser)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptCompensationSummary.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_DEDUCTION_SUMMARY_TOP(strCompany, strUser)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptDeductionSummaryTop
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CStr(strUser)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptDeductionSummaryTop.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function


Public Function PRINT_DEDUCTION_SUMMARY_V4(strUser)
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
'Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptPayrollRegisterDed.rpt", 1)
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptPayrollRegisterDedV4.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_DEDUCTION_SUMMARY_V3(strUser)
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
'Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptPayrollRegisterDed.rpt", 1)
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptPayrollRegisterDedV3.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function


Public Function PRINT_DEDUCTION_SUMMARY_V2(strUser)
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
'Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptPayrollRegisterDed.rpt", 1)
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptPayrollRegisterDedV2.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_DEDUCTION_SUMMARY(strCompany, strUser)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptDeductionSummary
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CStr(strUser)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptDeductionSummary.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_SSS_LOAN(strCompany, strUser)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptSSSLoan
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CStr(strUser)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptSSSLoan.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_PAGIBIG_LOAN(strCompany, strUser)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptPagIbigLoan
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CStr(strUser)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptPagIbigLoan.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_SSS_COLLECTION(strCompany, strUser)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptSSSCollection
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CStr(strUser)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptSSSCollection.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_PHIC_COLLECTION(strCompany, strUser)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptPHICCollection
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CStr(strUser)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptPHICCollection.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Public Function PRINT_PAGIBIG_COLLECTION(strCompany, strCompanyTelNo, strCompanystrCompanySSSNo, strUser)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptHDMF
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CStr(strCompanyTelNo)
'Report.ParameterFields(3).ClearCurrentValueAndRange
'Report.ParameterFields(3).AddCurrentValue CStr(strCompanystrCompanySSSNo)
'Report.ParameterFields(4).ClearCurrentValueAndRange
'Report.ParameterFields(4).AddCurrentValue CStr(strUser)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.Zoom 100
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptHDMF.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CStr(strCompany)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CStr(strCompanyTelNo)
C_Report.ParameterFields(3).ClearCurrentValueAndRange
C_Report.ParameterFields(3).AddCurrentValue CStr(strCompanystrCompanySSSNo)
C_Report.ParameterFields(4).ClearCurrentValueAndRange
C_Report.ParameterFields(4).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.Zoom 100
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function


Public Function PRINT_TAX_COLLECTION(intDivision, strCompany, strUser)
'Screen.MousePointer = vbHourglass
'Dim Report As New rptTaxWithHeld
'Report.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
'Report.DiscardSavedData
'Report.ParameterFields(1).ClearCurrentValueAndRange
'Report.ParameterFields(1).AddCurrentValue CDbl(intDivision)
'Report.ParameterFields(2).ClearCurrentValueAndRange
'Report.ParameterFields(2).AddCurrentValue CStr(strCompany)
'Report.ParameterFields(3).ClearCurrentValueAndRange
'Report.ParameterFields(3).AddCurrentValue CStr(strUser)
'CRViewer1.DisplayGroupTree = False
'CRViewer1.EnableExportButton = True
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
Screen.MousePointer = vbHourglass
C_Application.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
Set C_Report = Nothing
Set C_Report = C_Application.OpenReport(App.Path & "\Reports\rptTaxWithHeld.rpt", 1)
C_Report.DiscardSavedData
C_Report.ParameterFields(1).ClearCurrentValueAndRange
C_Report.ParameterFields(1).AddCurrentValue CDbl(intDivision)
C_Report.ParameterFields(2).ClearCurrentValueAndRange
C_Report.ParameterFields(2).AddCurrentValue CStr(strCompany)
C_Report.ParameterFields(3).ClearCurrentValueAndRange
C_Report.ParameterFields(3).AddCurrentValue CStr(strUser)
CRViewer1.ReportSource = C_Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Function

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
End Sub

Private Sub Form_Load()
'Screen.MousePointer = vbHourglass
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
On Error Resume Next
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth
End Sub
