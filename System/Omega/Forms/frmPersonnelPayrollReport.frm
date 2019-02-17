VERSION 5.00
Begin VB.Form frmPersonnelPayrollReport 
   Appearance      =   0  'Flat
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8220
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
   ScaleHeight     =   5565
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimerForATM2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5760
      Top             =   0
   End
   Begin RPVGCC.b8Container pic13thMonth 
      Height          =   2535
      Left            =   2040
      TabIndex        =   20
      Top             =   1080
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4471
      BackColor       =   15396057
      Begin VB.ComboBox cmbDivision13th 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   600
         Width           =   2535
      End
      Begin VB.ComboBox cmbQuarter 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtYear13th 
         Height          =   315
         Left            =   1080
         TabIndex        =   24
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CommandButton cmdCancel13th 
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
         Picture         =   "frmPersonnelPayrollReport.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1800
         Width           =   1560
      End
      Begin VB.CommandButton cmdOK13th 
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
         Picture         =   "frmPersonnelPayrollReport.frx":075C
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1800
         Width           =   1560
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2640
         Width           =   3495
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   40
         TabIndex        =   25
         Top             =   40
         Width           =   3890
         _ExtentX        =   6853
         _ExtentY        =   609
         Caption         =   "13th Month"
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
         Icon            =   "frmPersonnelPayrollReport.frx":0DCE
         ShadowVisible   =   0   'False
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Quarter"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   1320
         Width           =   615
      End
   End
   Begin RPVGCC.b8Container picTaxWithHeldAlpha 
      Height          =   1815
      Left            =   2040
      TabIndex        =   10
      Top             =   1680
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3201
      BackColor       =   15396057
      Begin VB.ComboBox cmbDivisionAlpha 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2640
         Width           =   3495
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
         Picture         =   "frmPersonnelPayrollReport.frx":1368
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1080
         Width           =   1560
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
         Picture         =   "frmPersonnelPayrollReport.frx":19DA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1080
         Width           =   1560
      End
      Begin VB.TextBox txtTaxAlphaYear 
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin RPVGCC.b8TitleBar b8TitleBar4 
         Height          =   345
         Left            =   40
         TabIndex        =   15
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
         Icon            =   "frmPersonnelPayrollReport.frx":2136
         ShadowVisible   =   0   'False
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "SELECT DIVISION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1080
         TabIndex        =   16
         Top             =   600
         Width           =   615
      End
   End
   Begin RPVGCC.b8Container picProgress 
      Height          =   975
      Left            =   1080
      TabIndex        =   18
      Top             =   2160
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
         TabIndex        =   19
         Top             =   120
         Width           =   5295
      End
   End
   Begin VB.Timer TimerForATM 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5280
      Top             =   0
   End
   Begin VB.Timer Timer13Month 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4800
      Top             =   0
   End
   Begin VB.Timer TimerContri 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3840
      Top             =   0
   End
   Begin VB.Timer TimerLoans 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3360
      Top             =   0
   End
   Begin VB.Timer TimerDeductions 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2880
      Top             =   0
   End
   Begin VB.Timer TimerEarnings 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2400
      Top             =   0
   End
   Begin VB.Timer TimerPaySlip 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1920
      Top             =   0
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   240
      ScaleHeight     =   5055
      ScaleWidth      =   7695
      TabIndex        =   0
      Top             =   240
      Width           =   7695
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
         Left            =   4200
         Picture         =   "frmPersonnelPayrollReport.frx":26D0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4560
         Width           =   1560
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
         Left            =   5880
         Picture         =   "frmPersonnelPayrollReport.frx":2D42
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4560
         Width           =   1560
      End
      Begin VB.ListBox lstResultPrint 
         Height          =   4155
         Left            =   3960
         TabIndex        =   7
         Top             =   240
         Width           =   3735
      End
      Begin VB.ListBox lstReportType 
         Height          =   3375
         Left            =   0
         TabIndex        =   3
         Top             =   240
         Width           =   3855
      End
      Begin VB.ComboBox cmbDivision 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   3960
         Width           =   3855
      End
      Begin VB.ComboBox cmbPeriodPrint 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   4680
         Width           =   3855
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "REPORT TYPE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   3375
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "SELECT DIVISION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   3720
         Width           =   3855
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "PAYROLL PERIOD"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   4440
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmPersonnelPayrollReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public PostLevel As Long
Dim Filename As String
Dim WorkbookName As String
Dim iWorkSheet As Integer
Dim RowCnt, ColCnt, iCnt, strRange, i, j, l, k, x, strValue, iReset, strAmount, _
iPK, Arr, Arr1, iLocationKey, iFilterIndex, sLine, strPath, iRec, iTerms, iDivision, _
sTaxStatus, strRange1, strRange2, strGrossTot, staTaxTot, strSSSTot, strPHICTot, _
strHDMFTot, strColaTot, strAllowTot

Dim dNetPay

Private Sub b8TitleBar1_CLoseClick()
cmdCancel13th_Click
End Sub

Private Sub b8TitleBar4_CLoseClick()
cmdCancelTaxAlpha_Click
End Sub

Private Sub cmbDivision_Click()
If cmbDivision.ListIndex = -1 Then cmbPeriodPrint.Clear: lstResultPrint.Clear: Exit Sub
cmbPeriodPrint.Clear: lstResultPrint.Clear
t = "SELECT dbo.tbl_Personnel_Compensation_Period.PayrollDate, dbo.tbl_Personnel_Payroll.PayrollPeriodKey " & _
    " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK " & _
    " Where (dbo.tbl_Personnel_ActionNew.DivisionKey = " & cmbDivision.ItemData(cmbDivision.ListIndex) & ") " & _
    " GROUP BY dbo.tbl_Personnel_Compensation_Period.PayrollDate, dbo.tbl_Personnel_Payroll.PayrollPeriodKey " & _
    " ORDER BY dbo.tbl_Personnel_Compensation_Period.PayrollDate DESC"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
While Not rt.EOF
    cmbPeriodPrint.AddItem Format(FormatDateTime(rt!PayrollDate, vbShortDate), "mm/dd/yyyy")
    cmbPeriodPrint.ItemData(cmbPeriodPrint.NewIndex) = rt!PayrollPeriodKey
    rt.MoveNext
Wend
rt.Close
If cmbPeriodPrint.ListCount Then cmbPeriodPrint.ListIndex = 0
End Sub

Private Sub cmbPeriodPrint_Click()
If cmbPeriodPrint.ListIndex = -1 Then lstResultPrint.Clear: Exit Sub
lstResultPrint.Clear
Select Case lstReportType.ItemData(lstReportType.ListIndex)
    Case 5  'Loans
        'u = "SELECT dbo.tbl_Personnel_Payroll_Deductions.DeductionKey as PKey, " & _
            " dbo.tbl_Personnel_Payroll_Deductions_Table.Description as LabelName " & _
            " FROM  dbo.tbl_Personnel_Payroll_Deductions LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Deductions.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Payroll_Deductions.DeductionKey = dbo.tbl_Personnel_Payroll_Deductions_Table.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
            " Where (dbo.tbl_Personnel_Payroll.PayrollPeriodKey = " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ") " & _
            " And (dbo.tbl_Personnel_Payroll_Deductions_Table.GovtDed = 1) " & _
            " GROUP BY dbo.tbl_Personnel_Payroll_Deductions.DeductionKey, dbo.tbl_Personnel_Payroll_Deductions_Table.Description, dbo.tbl_Personnel_Payroll_Deductions_Table.Sorting " & _
            " ORDER BY dbo.tbl_Personnel_Payroll_Deductions_Table.Sorting"
        u = ""
    Case 6  'Contribution
        'u = "SELECT dbo.tbl_Personnel_Payroll_Deductions.DeductionKey as PKey, " & _
            " dbo.tbl_Personnel_Payroll_Deductions_Table.Description as LabelName " & _
            " FROM  dbo.tbl_Personnel_Payroll_Deductions LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Deductions.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Payroll_Deductions_Table ON dbo.tbl_Personnel_Payroll_Deductions.DeductionKey = dbo.tbl_Personnel_Payroll_Deductions_Table.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK " & _
            " Where (dbo.tbl_Personnel_Payroll.PayrollPeriodKey = " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ") " & _
            " And (dbo.tbl_Personnel_Payroll_Deductions_Table.GovtDed = 2) " & _
            " GROUP BY dbo.tbl_Personnel_Payroll_Deductions.DeductionKey, dbo.tbl_Personnel_Payroll_Deductions_Table.Description, dbo.tbl_Personnel_Payroll_Deductions_Table.Sorting " & _
            " ORDER BY dbo.tbl_Personnel_Payroll_Deductions_Table.Sorting"
        u = ""
    Case 7
        u = ""
    Case 8
        u = ""
    Case 9
        u = ""
    Case Else
        lstResultPrint.AddItem "ALL"
        lstResultPrint.ItemData(lstResultPrint.NewIndex) = 0
        u = "SELECT dbo.tbl_Personnel_Department.DepartmentName as  LabelName, dbo.tbl_Personnel_ActionNew.DeptKey as PKey " & _
            " FROM  dbo.tbl_Personnel_Department LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Department.PK = dbo.tbl_Personnel_ActionNew.DeptKey RIGHT OUTER JOIN " & _
            " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_ActionNew.PK = dbo.tbl_Personnel_Payroll.ActionMemoKey " & _
            " Where (dbo.tbl_Personnel_ActionNew.DivisionKey = " & cmbDivision.ItemData(cmbDivision.ListIndex) & ") " & _
            " And (dbo.tbl_Personnel_Payroll.PayrollPeriodKey = " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ") " & _
            " GROUP BY dbo.tbl_Personnel_Department.DepartmentName, dbo.tbl_Personnel_ActionNew.DeptKey " & _
            " ORDER BY dbo.tbl_Personnel_Department.DepartmentName"
End Select
If u = "" Then Exit Sub
If ru.State = adStateOpen Then ru.Close
ru.Open u, ConnOmega
While Not ru.EOF
    lstResultPrint.AddItem ru!LabelName
    lstResultPrint.ItemData(lstResultPrint.NewIndex) = ru!PKey
    ru.MoveNext
Wend
ru.Close
If lstResultPrint.ListCount Then lstResultPrint.ListIndex = 0
End Sub

Private Sub cmdCancel13th_Click()
pic13thMonth.Visible = False
picMain.Enabled = True
End Sub

Private Sub cmdCancelPrint_Click()
Unload Me
End Sub

Private Sub cmdCancelTaxAlpha_Click()
picTaxWithHeldAlpha.Visible = False
picMain.Enabled = True
End Sub

Private Sub cmdOK13th_Click()
If cmbDivision13th.ListIndex = -1 Then MsgBox "Please select division!                           ", vbCritical, "Error...": cmbDivision13th.SetFocus: Exit Sub
If cmbQuarter.ListIndex = -1 Then MsgBox "Please select quarter!                     ", vbCritical, "Error...": cmbQuarter.SetFocus: Exit Sub
If RETURNTEXTVALUE(txtYear13th) <= 0 Then MsgBox "Please supply year!                     ", vbCritical, "Error...": txtYear13th.SetFocus: Exit Sub
With MainFormPopupF
    For i = 1 To .mnuPayrollPrint1RnFSup.UBound
        Unload .mnuPayrollPrint1RnFSup(i)
    Next i
    l = 0
    t = "SELECT tbl_Personnel_Position_Level.* " & _
        " FROM tbl_Personnel_Position_Level " & _
        " ORDER BY PK"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        l = l + 1
        If l = 1 Then
            .mnuPayrollPrint1RnFSup(0).Caption = rt!LevelName
        Else
            Load .mnuPayrollPrint1RnFSup(l - 1)
            .mnuPayrollPrint1RnFSup(l - 1).Caption = rt!LevelName
        End If
        rt.MoveNext
    Wend
    rt.Close
    PopupMenu MainFormPopupF.mnuPayrollPrint1, , pic13thMonth.Left + cmdOK13th.Left + 200, pic13thMonth.Top + cmdOK13th.Top + 200
End With
End Sub

Private Sub cmdOKPrint_Click()
'MsgBox lstReportType.ItemData(lstReportType.ListIndex)
'Select Case lstReportType.ItemData(lstReportType.ListIndex)
'    Case 1: PopupMenu MainFormPopupF.mnuPayrollPrint, , picMain.Left + cmdOKPrint.Left + 200, picMain.Top + cmdOKPrint.Top + 200 'TimerPaySlip.Enabled = True     'Payslip
'    Case 2: PopupMenu MainFormPopupF.mnuPayrollPrint, , picMain.Left + cmdOKPrint.Left + 200, picMain.Top + cmdOKPrint.Top + 200 'TimerPaySlip.Enabled = True     'Signledger
'    Case 3: PopupMenu MainFormPopupF.mnuPayrollPrint, , picMain.Left + cmdOKPrint.Left + 200, picMain.Top + cmdOKPrint.Top + 200 ' TimerEarnings.Enabled = True    'Earnings
'    Case 4: PopupMenu MainFormPopupF.mnuPayrollPrint, , picMain.Left + cmdOKPrint.Left + 200, picMain.Top + cmdOKPrint.Top + 200 ' TimerDeductions.Enabled = True  'Deductions
'    Case 5: PopupMenu MainFormPopupF.mnuPayrollPrint, , picMain.Left + cmdOKPrint.Left + 200, picMain.Top + cmdOKPrint.Top + 200 'TimerLoans.Enabled = True       'Loans
'    Case 6: PopupMenu MainFormPopupF.mnuPayrollPrint, , picMain.Left + cmdOKPrint.Left + 200, picMain.Top + cmdOKPrint.Top + 200 'TimerContri.Enabled = True      'Contributions
'    Case 7: 'Alphalist
'    Case 8: PopupMenu MainFormPopupF.mnuPayrollPrint, , picMain.Left + cmdOKPrint.Left + 200, picMain.Top + cmdOKPrint.Top + 200 'Timer13Month.Enabled = True     '13th Month
'    Case 9: PopupMenu MainFormPopupF.mnuPayrollPrint, , picMain.Left + cmdOKPrint.Left + 200, picMain.Top + cmdOKPrint.Top + 200
'End Select



If lstReportType.ItemData(lstReportType.ListIndex) <> 7 And _
lstReportType.ItemData(lstReportType.ListIndex) <> 8 Then
    With MainFormPopupF
        For i = 1 To .mnuPayrollPrint1RnFSup.UBound
            Unload .mnuPayrollPrint1RnFSup(i)
        Next i
        l = 0
        t = "SELECT tbl_Personnel_Position_Level.* " & _
            " FROM tbl_Personnel_Position_Level " & _
            " ORDER BY PK"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        While Not rt.EOF
            l = l + 1
            If l = 1 Then
                .mnuPayrollPrint1RnFSup(0).Caption = rt!LevelName
            Else
                Load .mnuPayrollPrint1RnFSup(l - 1)
                .mnuPayrollPrint1RnFSup(l - 1).Caption = rt!LevelName
            End If
            rt.MoveNext
        Wend
        rt.Close
        PopupMenu MainFormPopupF.mnuPayrollPrint1, , picMain.Left + cmdOKPrint.Left + 200, picMain.Top + cmdOKPrint.Top + 200
    End With
ElseIf lstReportType.ItemData(lstReportType.ListIndex) = 8 Then
    '13th Month
    pic13thMonth.ZOrder 0
    cmbDivision13th.ListIndex = -1
    cmbQuarter.ListIndex = -1
    txtYear13th.Text = Format(Date, "yyyy")
    pic13thMonth.Visible = True
    cmbDivision13th.SetFocus
Else
    'Alphalist
    picTaxWithHeldAlpha.ZOrder 0
    txtTaxAlphaYear.Text = CDbl(Format(Date, "yyyy")) - 1
    picMain.Enabled = False
    picTaxWithHeldAlpha.Visible = True
    txtTaxAlphaYear.SetFocus
End If

End Sub

Private Sub cmdOKTaxAlpha_Click()
If RETURNTEXTVALUE(txtTaxAlphaYear) <= 0 Then Exit Sub

MainForm.CommonDialog1.CancelError = True
On Error GoTo ErrorHandler
MainForm.CommonDialog1.DialogTitle = "Save"
MainForm.CommonDialog1.Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls"
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

's = "SELECT tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
    " tbl_Personnel_Information.TIN " & _
    " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
    " tbl_Personnel_IDNumber ON tbl_Personnel_Compensation.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK LEFT OUTER JOIN " & _
    " tbl_Personnel_Compensation_Period ON tbl_Personnel_Compensation.Period = tbl_Personnel_Compensation_Period.PK " & _
    " Where (Year(tbl_Personnel_Compensation_Period.DateTo) = " & RETURNTEXTVALUE(txtTaxAlphaYear) & ") " & _
    " GROUP BY tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName, " & _
    " tbl_Personnel_Information.TIN " & _
    " ORDER BY tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName"
s = "SELECT dbo.tbl_Personnel_Information.LastName + ',  ' + dbo.tbl_Personnel_Information.FirstName + '  ' + dbo.tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
    " dbo.tbl_Personnel_Information.TIN, dbo.tbl_Personnel_IDNumber.ProfileKey " & _
    " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Payroll.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
    " Where (Year(dbo.tbl_Personnel_Compensation_Period.PayrollDate) = " & RETURNTEXTVALUE(txtTaxAlphaYear) & ") " & _
    " GROUP BY dbo.tbl_Personnel_Information.LastName + ',  ' + dbo.tbl_Personnel_Information.FirstName + '  ' + dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_Information.TIN, dbo.tbl_Personnel_IDNumber.ProfileKey " & _
    " ORDER BY EmployeeName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    DoEvents
    i = i + 1
    iDivision = 0
    't = "SELECT TOP 1 tbl_Personnel_Action.Division " & _
        " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
        " tbl_Personnel_IDNumber ON tbl_Personnel_Action.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
        " WHERE (YEAR(tbl_Personnel_Action.EffectivityDate) <= " & RETURNTEXTVALUE(txtTaxAlphaYear) & ") " & _
        " AND (tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName = '" & FORMATSQL(rs!EmployeeName) & "') " & _
        " ORDER BY tbl_Personnel_Action.EffectivityDate DESC"
    t = "SELECT TOP (1) dbo.tbl_Personnel_ActionNew.DivisionKey as Division " & _
        " FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_ActionNew.EmpPK = dbo.tbl_Personnel_IDNumber.PK " & _
        " Where (Year(dbo.tbl_Personnel_ActionNew.EffectivityDate) <= " & RETURNTEXTVALUE(txtTaxAlphaYear) & ") " & _
        " And (dbo.tbl_Personnel_IDNumber.ProfileKey = " & rs!ProfileKey & ") " & _
        " ORDER BY dbo.tbl_Personnel_ActionNew.EffectivityDate DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        iDivision = rt!Division
    End If
    rt.Close
    
    sTaxStatus = ""
    't = "SELECT TOP 1 tbl_Personnel_TaxStatus.TaxStatus " & _
        " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
        " tbl_Personnel_IDNumber ON tbl_Personnel_Action.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_TaxStatus ON tbl_Personnel_Action.TaxStatus = tbl_Personnel_TaxStatus.PK " & _
        " WHERE (YEAR(tbl_Personnel_Action.EffectivityDate) <= " & RETURNTEXTVALUE(txtTaxAlphaYear) & ") " & _
        " AND (tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName = '" & FORMATSQL(rs!EmployeeName) & "') " & _
        " ORDER BY tbl_Personnel_Action.EffectivityDate DESC"
    t = "SELECT TOP (1) dbo.tbl_Personnel_TaxStatus.TaxStatus " & _
        " FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_ActionNew.EmpPK = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_TaxStatus ON dbo.tbl_Personnel_ActionNew.TaxStatusKey = dbo.tbl_Personnel_TaxStatus.PK " & _
        " Where (Year(dbo.tbl_Personnel_ActionNew.EffectivityDate) <= " & RETURNTEXTVALUE(txtTaxAlphaYear) & ") " & _
        " And (dbo.tbl_Personnel_IDNumber.ProfileKey = " & rs!ProfileKey & ") " & _
        " ORDER BY dbo.tbl_Personnel_ActionNew.EffectivityDate DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        sTaxStatus = rt!TaxStatus
    End If
    rt.Close
    
    For j = 1 To 12
        't = "SELECT SUM(tbl_Personnel_Compensation.TotalEarning) AS Gross, " & _
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
        t = "SELECT YEAR(tbl_Personnel_Compensation_Period_1.PayrollDate) AS dYear, MONTH(tbl_Personnel_Compensation_Period_1.PayrollDate) AS dMonth, tbl_Personnel_IDNumber_1.ProfileKey, ISNULL((SELECT SUM(dbo.tbl_Personnel_Payroll_Earnings.Taxable) AS Taxable " & _
            " FROM  dbo.tbl_Personnel_Payroll_Earnings LEFT OUTER JOIN dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Earnings.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Payroll.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK WHERE (YEAR(dbo.tbl_Personnel_Compensation_Period.PayrollDate) = YEAR(tbl_Personnel_Compensation_Period_1.PayrollDate)) AND (MONTH(dbo.tbl_Personnel_Compensation_Period.PayrollDate) = MONTH(tbl_Personnel_Compensation_Period_1.PayrollDate)) AND (dbo.tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_IDNumber_1.ProfileKey)), 0) AS Taxable, " & _
            " IsNull((SELECT SUM(dbo.tbl_Personnel_Payroll_Deductions.Amount) AS Amount FROM  dbo.tbl_Personnel_Payroll_Deductions LEFT OUTER JOIN dbo.tbl_Personnel_Payroll AS tbl_Personnel_Payroll_2 ON dbo.tbl_Personnel_Payroll_Deductions.MasterKey = tbl_Personnel_Payroll_2.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period AS tbl_Personnel_Compensation_Period_2 ON tbl_Personnel_Payroll_2.PayrollPeriodKey = tbl_Personnel_Compensation_Period_2.PK LEFT OUTER JOIN dbo.tbl_Personnel_IDNumber AS tbl_Personnel_IDNumber_2 ON tbl_Personnel_Payroll_2.EmployeeKey = tbl_Personnel_IDNumber_2.PK " & _
            " WHERE (YEAR(tbl_Personnel_Compensation_Period_2.PayrollDate) = YEAR(tbl_Personnel_Compensation_Period_1.PayrollDate)) AND (MONTH(tbl_Personnel_Compensation_Period_2.PayrollDate) = MONTH(tbl_Personnel_Compensation_Period_1.PayrollDate)) AND (tbl_Personnel_IDNumber_2.ProfileKey = tbl_Personnel_IDNumber_1.ProfileKey) AND " & _
            " (dbo.tbl_Personnel_Payroll_Deductions.DeductionKey = 1)), 0) AS SSS, ISNULL ((SELECT SUM(tbl_Personnel_Payroll_Deductions_3.Amount) AS Amount FROM dbo.tbl_Personnel_Payroll_Deductions AS tbl_Personnel_Payroll_Deductions_3 LEFT OUTER JOIN dbo.tbl_Personnel_Payroll AS tbl_Personnel_Payroll_2 ON tbl_Personnel_Payroll_Deductions_3.MasterKey = tbl_Personnel_Payroll_2.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period AS tbl_Personnel_Compensation_Period_2 ON tbl_Personnel_Payroll_2.PayrollPeriodKey = tbl_Personnel_Compensation_Period_2.PK LEFT OUTER JOIN dbo.tbl_Personnel_IDNumber AS tbl_Personnel_IDNumber_2 ON tbl_Personnel_Payroll_2.EmployeeKey = tbl_Personnel_IDNumber_2.PK WHERE (YEAR(tbl_Personnel_Compensation_Period_2.PayrollDate) = YEAR(tbl_Personnel_Compensation_Period_1.PayrollDate)) " & _
            " AND (MONTH(tbl_Personnel_Compensation_Period_2.PayrollDate) = MONTH(tbl_Personnel_Compensation_Period_1.PayrollDate)) AND (tbl_Personnel_IDNumber_2.ProfileKey = tbl_Personnel_IDNumber_1.ProfileKey) AND (tbl_Personnel_Payroll_Deductions_3.DeductionKey = 4)), 0) AS PHIC, ISNULL((SELECT SUM(tbl_Personnel_Payroll_Deductions_2.Amount) AS Amount " & _
            " FROM  dbo.tbl_Personnel_Payroll_Deductions AS tbl_Personnel_Payroll_Deductions_2 LEFT OUTER JOIN dbo.tbl_Personnel_Payroll AS tbl_Personnel_Payroll_2 ON tbl_Personnel_Payroll_Deductions_2.MasterKey = tbl_Personnel_Payroll_2.PK LEFT OUTER JOIN dbo.tbl_Personnel_Compensation_Period AS tbl_Personnel_Compensation_Period_2 ON tbl_Personnel_Payroll_2.PayrollPeriodKey = tbl_Personnel_Compensation_Period_2.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_IDNumber AS tbl_Personnel_IDNumber_2 ON tbl_Personnel_Payroll_2.EmployeeKey = tbl_Personnel_IDNumber_2.PK WHERE (YEAR(tbl_Personnel_Compensation_Period_2.PayrollDate) = YEAR(tbl_Personnel_Compensation_Period_1.PayrollDate)) AND (MONTH(tbl_Personnel_Compensation_Period_2.PayrollDate) = MONTH(tbl_Personnel_Compensation_Period_1.PayrollDate)) AND (tbl_Personnel_IDNumber_2.ProfileKey = tbl_Personnel_IDNumber_1.ProfileKey) AND " & _
            " (tbl_Personnel_Payroll_Deductions_2.DeductionKey = 6)), 0) AS PagIbig, ISNULL((SELECT SUM(tbl_Personnel_Payroll_Deductions_1.Amount) AS Amount FROM  dbo.tbl_Personnel_Payroll_Deductions AS tbl_Personnel_Payroll_Deductions_1 LEFT OUTER JOIN dbo.tbl_Personnel_Payroll AS tbl_Personnel_Payroll_2 ON tbl_Personnel_Payroll_Deductions_1.MasterKey = tbl_Personnel_Payroll_2.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Compensation_Period AS tbl_Personnel_Compensation_Period_2 ON tbl_Personnel_Payroll_2.PayrollPeriodKey = tbl_Personnel_Compensation_Period_2.PK LEFT OUTER JOIN dbo.tbl_Personnel_IDNumber AS tbl_Personnel_IDNumber_2 ON tbl_Personnel_Payroll_2.EmployeeKey = tbl_Personnel_IDNumber_2.PK WHERE (YEAR(tbl_Personnel_Compensation_Period_2.PayrollDate) = YEAR(tbl_Personnel_Compensation_Period_1.PayrollDate)) AND " & _
            " (MONTH(tbl_Personnel_Compensation_Period_2.PayrollDate) = MONTH(tbl_Personnel_Compensation_Period_1.PayrollDate)) AND (tbl_Personnel_IDNumber_2.ProfileKey = tbl_Personnel_IDNumber_1.ProfileKey) AND (tbl_Personnel_Payroll_Deductions_1.DeductionKey = 8)), 0) AS WithHolding " & _
            " FROM  dbo.tbl_Personnel_Payroll AS tbl_Personnel_Payroll_1 LEFT OUTER JOIN dbo.tbl_Personnel_Compensation_Period AS tbl_Personnel_Compensation_Period_1 ON tbl_Personnel_Payroll_1.PayrollPeriodKey = tbl_Personnel_Compensation_Period_1.PK LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_IDNumber AS tbl_Personnel_IDNumber_1 ON tbl_Personnel_Payroll_1.EmployeeKey = tbl_Personnel_IDNumber_1.PK " & _
            " GROUP BY YEAR(tbl_Personnel_Compensation_Period_1.PayrollDate), MONTH(tbl_Personnel_Compensation_Period_1.PayrollDate), tbl_Personnel_IDNumber_1.ProfileKey " & _
            " HAVING (YEAR(tbl_Personnel_Compensation_Period_1.PayrollDate) = " & RETURNTEXTVALUE(txtTaxAlphaYear) & ") " & _
            " AND (MONTH(tbl_Personnel_Compensation_Period_1.PayrollDate) = " & j & ") " & _
            " AND (tbl_Personnel_IDNumber_1.ProfileKey = " & rs!ProfileKey & ")"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            u = "SELECT tbl_Personnel_Tax_Alphalist.* " & _
                " FROM tbl_Personnel_Tax_Alphalist " & _
                " WHERE (LogInName = '" & gbl_UserName & "') " & _
                " AND (ProfileKey = " & rs!ProfileKey & ")"
            If ru.State = adStateOpen Then ru.Close
            ru.Open u, ConnOmega
            If ru.RecordCount = 0 Then
                ConnOmega.Execute "INSERT INTO tbl_Personnel_Tax_Alphalist " & _
                                  " (LogInName, ProfileKey, EmployeeName, Tin, TaxStatus, Division) " & _
                                  " VALUES ('" & gbl_UserName & "', " & rs!ProfileKey & ", '" & FORMATSQL(rs!EmployeeName) & "', '" & rs!TIN & "', '" & FORMATSQL(CStr(sTaxStatus)) & "', " & iDivision & ")"
            End If
            ru.Close
            
            'ConnOmega.Execute "UPDATE tbl_Personnel_Tax_Alphalist " & _
                                      " SET " & "Gross" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!Taxable), 0, rt!Taxable)) & ", " & _
                                      " " & "Tax" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!Withholding), 0, rt!Withholding)) & ", " & _
                                      " " & "SSS" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!SSS), 0, rt!SSS)) & ", " & _
                                      " " & "PHIC" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!PHIC), 0, rt!PHIC)) & ", " & _
                                      " " & "HDMF" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!PAGIBIG), 0, rt!PAGIPAGIBIGBIG)) & ", " & _
                                      " " & "Cola" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!Cola), 0, rt!Cola)) & ", " & _
                                      " " & "Allow" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!Allowance), 0, rt!Allowance)) & " " & _
                                      " WHERE (LogInName = '" & gbl_UserName & "') " & _
                                      " AND (ProfileKey = " & rs!ProfileKey & ")"
            
            ConnOmega.Execute "UPDATE tbl_Personnel_Tax_Alphalist " & _
                                      " SET " & "Gross" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!Taxable), 0, rt!Taxable)) & ", " & _
                                      " " & "Tax" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!Withholding), 0, rt!Withholding)) & ", " & _
                                      " " & "SSS" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!SSS), 0, rt!SSS)) & ", " & _
                                      " " & "PHIC" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!PHIC), 0, rt!PHIC)) & ", " & _
                                      " " & "HDMF" & Format(j, "0#") & " = " & CDbl(IIf(IsNull(rt!PAGIBIG), 0, rt!PAGIBIG)) & ", " & _
                                      " " & "Cola" & Format(j, "0#") & " = 0, " & _
                                      " " & "Allow" & Format(j, "0#") & " = 0 " & _
                                      " WHERE (LogInName = '" & gbl_UserName & "') " & _
                                      " AND (ProfileKey = " & rs!ProfileKey & ")"
            
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
                'xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "#,##0.00"
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
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
        'xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "#,##0.00"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
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
strRange = EXCEL_RANGE(6, 6)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Select
xlsApp.ActiveWindow.FreezePanes = True

If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

xlsApp.Visible = True

picProgressBar.BackColor = &HFFFFFF
picProgress.Visible = False
picMain.Enabled = True

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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    If pic13thMonth.Visible = True Then cmdCancel13th_Click: Exit Sub
    If picTaxWithHeldAlpha.Visible = True Then cmdCancelTaxAlpha_Click: Exit Sub
    Unload Me
End If
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.Height - Me.Height) / 6
Me.Left = (MainForm.Width - Me.Width) / 3
POPULATE_COMBO "PK", "Description", "tbl_Personnel_Division", "Description", cmbDivision
POPULATE_COMBO "PK", "Description", "tbl_Personnel_Division", "Description", cmbDivision13th
With lstReportType
    .Clear
    .AddItem "PAYSLIP": .ItemData(.NewIndex) = 1
    .AddItem "SIGNATURE LEDGER": .ItemData(.NewIndex) = 2
    .AddItem "COMPENSATION SUMMARY": .ItemData(.NewIndex) = 3
    .AddItem "DEDUCTION SUMMARY": .ItemData(.NewIndex) = 4
    .AddItem "LOANS": .ItemData(.NewIndex) = 5
    .AddItem "GOVERNMENT CONTRIBUTION": .ItemData(.NewIndex) = 6
    .AddItem "WITHHOLDING TAX (Alpha List)": .ItemData(.NewIndex) = 7
    .AddItem "13th MONTH": .ItemData(.NewIndex) = 8
    .AddItem "FOR ATM": .ItemData(.NewIndex) = 9
    .ListIndex = 0
End With
With cmbQuarter
    .Clear
    .AddItem "1st Quarter": .ItemData(.NewIndex) = 3
    .AddItem "2nd Quarter": .ItemData(.NewIndex) = 6
    .AddItem "3rd Quarter": .ItemData(.NewIndex) = 9
    .AddItem "4th Quarter": .ItemData(.NewIndex) = 12
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
If pic13thMonth.Visible = True Then Cancel = -1
If picTaxWithHeldAlpha.Visible = True Then Cancel = -1
If picProgress.Visible = True Then Cancel = -1
End Sub

Private Sub Timer13Month_Timer()
Timer13Month.Enabled = False
'If cmbDivision.ListIndex = -1 Then MsgBox "Please select division!                  ", vbCritical, "Error...": cmbDivision.SetFocus: Exit Sub
'If cmbPeriodPrint.ListIndex = -1 Then MsgBox "Please select payroll date!                    ", vbCritical, "Error...": cmbPeriodPrint.SetFocus: Exit Sub
Screen.MousePointer = vbHourglass
'Generate13thMonth gbl_UserName, cmbPeriodPrint.List(cmbPeriodPrint.ListIndex), cmbDivision.ItemData(cmbDivision.ListIndex), cmbDivision.List(cmbDivision.ListIndex), PostLevel

Generate13thMonth gbl_UserName, cmbDivision13th.ItemData(cmbDivision13th.ListIndex), cmbDivision13th.List(cmbDivision13th.ListIndex), PostLevel, cmbQuarter.ItemData(cmbQuarter.ListIndex), RETURNTEXTVALUE(txtYear13th)

Screen.MousePointer = vbDefault
frmCrystalReportViewer.PRINT_13TH_MONTH_V2 gbl_UserName, cmbQuarter.ItemData(cmbQuarter.ListIndex) 'Month(cmbPeriodPrint.List(cmbPeriodPrint.ListIndex))
If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
End Sub

Private Sub TimerContri_Timer()
TimerContri.Enabled = False
If cmbDivision.ListIndex = -1 Then MsgBox "Please select division!                  ", vbCritical, "Error...": cmbDivision.SetFocus: Exit Sub
If cmbPeriodPrint.ListIndex = -1 Then MsgBox "Please select payroll date!                    ", vbCritical, "Error...": cmbPeriodPrint.SetFocus: Exit Sub
'If lstResultPrint.ListIndex = -1 Then MsgBox "Please select grouping!                    ", vbCritical, "Error...": lstResultPrint.SetFocus: Exit Sub
Dim dblContri
dblContri = 0
't = "SELECT ROUND(SUM(dbo.tbl_Personnel_Payroll_Deductions.Amount), 2) AS Amt " & _
    " FROM  dbo.tbl_Personnel_Payroll_Deductions LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Deductions.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
    " WHERE (dbo.tbl_Personnel_Payroll_Deductions.LoanKey IS NOT NULL) " & _
    " AND (dbo.tbl_Personnel_ActionNew.DivisionKey = " & cmbDivision.ItemData(cmbDivision.ListIndex) & ") " & _
    " AND (dbo.tbl_Personnel_Payroll.PayrollPeriodKey = " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ") " & _
    " AND (dbo.tbl_Personnel_Position.PositionLevel = " & PostLevel & ")"
t = "SELECT ROUND(SUM(dbo.tbl_Personnel_Payroll_Deductions.Amount), 2) AS Amt " & _
    " FROM  dbo.tbl_Personnel_Payroll_Deductions LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Deductions.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
    " WHERE (dbo.tbl_Personnel_ActionNew.DivisionKey = " & cmbDivision.ItemData(cmbDivision.ListIndex) & ") AND (dbo.tbl_Personnel_Position.PositionLevel = " & PostLevel & ") AND (dbo.tbl_Personnel_Payroll.PayrollPeriodKey = " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ") AND (dbo.tbl_Personnel_Payroll_Deductions.DeductionKey = 1) OR " & _
    " (dbo.tbl_Personnel_ActionNew.DivisionKey = " & cmbDivision.ItemData(cmbDivision.ListIndex) & ") AND (dbo.tbl_Personnel_Position.PositionLevel = " & PostLevel & ") AND (dbo.tbl_Personnel_Payroll.PayrollPeriodKey = " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ") AND (dbo.tbl_Personnel_Payroll_Deductions.DeductionKey = 4) OR " & _
    " (dbo.tbl_Personnel_ActionNew.DivisionKey = " & cmbDivision.ItemData(cmbDivision.ListIndex) & ") AND (dbo.tbl_Personnel_Position.PositionLevel = " & PostLevel & ") AND (dbo.tbl_Personnel_Payroll.PayrollPeriodKey = " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ") AND (dbo.tbl_Personnel_Payroll_Deductions.DeductionKey = 6) OR " & _
    " (dbo.tbl_Personnel_ActionNew.DivisionKey = " & cmbDivision.ItemData(cmbDivision.ListIndex) & ") AND (dbo.tbl_Personnel_Position.PositionLevel = " & PostLevel & ") AND (dbo.tbl_Personnel_Payroll.PayrollPeriodKey = " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ") AND (dbo.tbl_Personnel_Payroll_Deductions.DeductionKey = 8)"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then
    dblContri = IIf(IsNull(rt!Amt), 0, rt!Amt)
End If
rt.Close
If CDbl(dblContri) = 0 Then MsgBox "No contribution records!                     ", vbExclamation, "No data...": Exit Sub
Screen.MousePointer = vbHourglass
GenerateLoanContri gbl_UserName, cmbDivision.ItemData(cmbDivision.ListIndex), cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex), lstReportType.ItemData(lstReportType.ListIndex), cmbDivision.List(cmbDivision.ListIndex), PostLevel, Month(cmbPeriodPrint.List(cmbPeriodPrint.ListIndex)), Year(cmbPeriodPrint.List(cmbPeriodPrint.ListIndex))
Screen.MousePointer = vbDefault
frmCrystalReportViewer.PRINT_PAYROLL_Contribution gbl_UserName
If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
End Sub

Private Sub TimerDeductions_Timer()
TimerDeductions.Enabled = False
If cmbDivision.ListIndex = -1 Then MsgBox "Please select division!                  ", vbCritical, "Error...": cmbDivision.SetFocus: Exit Sub
If cmbPeriodPrint.ListIndex = -1 Then MsgBox "Please select payroll date!                    ", vbCritical, "Error...": cmbPeriodPrint.SetFocus: Exit Sub
If lstResultPrint.ListIndex = -1 Then MsgBox "Please select grouping!                    ", vbCritical, "Error...": lstResultPrint.SetFocus: Exit Sub
Screen.MousePointer = vbHourglass
t = "SELECT tbl_Personnel_Compensation_Period.* " & _
    " FROM tbl_Personnel_Compensation_Period " & _
    " WHERE (PayrollDate = '" & FormatDateTime(cmbPeriodPrint.List(cmbPeriodPrint.ListIndex), vbShortDate) & "')"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then
    iTerms = rt!Terms
End If
rt.Close
If lstResultPrint.ItemData(lstResultPrint.ListIndex) = 0 Then
    GenerateLedger gbl_UserName, 3, cmbDivision.ItemData(cmbDivision.ListIndex), cmbPeriodPrint.List(cmbPeriodPrint.ListIndex), 2, cmbDivision.ItemData(cmbDivision.ListIndex), PostLevel
Else
    GenerateLedger gbl_UserName, 2, lstResultPrint.ItemData(lstResultPrint.ListIndex), cmbPeriodPrint.List(cmbPeriodPrint.ListIndex), 2, cmbDivision.ItemData(cmbDivision.ListIndex), PostLevel
End If
Screen.MousePointer = vbDefault
If CDbl(iTerms) = 1 Then
    frmCrystalReportViewer.PRINT_DEDUCTION_SUMMARY_V3 gbl_UserName
ElseIf CDbl(iTerms) = 2 Then
    frmCrystalReportViewer.PRINT_DEDUCTION_SUMMARY_V4 gbl_UserName
End If
If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
End Sub

Private Sub TimerEarnings_Timer()
TimerEarnings.Enabled = False
If cmbDivision.ListIndex = -1 Then MsgBox "Please select division!                  ", vbCritical, "Error...": cmbDivision.SetFocus: Exit Sub
If cmbPeriodPrint.ListIndex = -1 Then MsgBox "Please select payroll date!                    ", vbCritical, "Error...": cmbPeriodPrint.SetFocus: Exit Sub
If lstResultPrint.ListIndex = -1 Then MsgBox "Please select grouping!                    ", vbCritical, "Error...": lstResultPrint.SetFocus: Exit Sub
Screen.MousePointer = vbHourglass
If lstResultPrint.ItemData(lstResultPrint.ListIndex) = 0 Then
    GenerateLedger gbl_UserName, 3, cmbDivision.ItemData(cmbDivision.ListIndex), cmbPeriodPrint.List(cmbPeriodPrint.ListIndex), 1, cmbDivision.ItemData(cmbDivision.ListIndex), PostLevel
Else
    GenerateLedger gbl_UserName, 2, lstResultPrint.ItemData(lstResultPrint.ListIndex), cmbPeriodPrint.List(cmbPeriodPrint.ListIndex), 1, cmbDivision.ItemData(cmbDivision.ListIndex), PostLevel
End If
Screen.MousePointer = vbDefault
frmCrystalReportViewer.PRINT_COMPENSATION_SUMMARY_V3 gbl_UserName
Screen.MousePointer = vbDefault
If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
End Sub

Private Sub TimerForATM_Timer()
TimerForATM.Enabled = False
If cmbDivision.ListIndex = -1 Then MsgBox "Please select division!                  ", vbCritical, "Error...": cmbDivision.SetFocus: Exit Sub
If cmbPeriodPrint.ListIndex = -1 Then MsgBox "Please select payroll date!                    ", vbCritical, "Error...": cmbPeriodPrint.SetFocus: Exit Sub
iRec = 0
i = 0
MainForm.picProgressBar.BackColor = &H8000000F
DoEvents
a = "SELECT dbo.tbl_Personnel_IDNumber.AccountNumber, dbo.tbl_Personnel_Information.LastName, " & _
    " dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName, " & _
    " ISNULL((SELECT SUM(TotalAmount) AS Amount From dbo.tbl_Personnel_Payroll_Earnings WHERE (MasterKey = dbo.tbl_Personnel_Payroll.PK)), 0) AS Earnings, " & _
    " ISNULL((SELECT SUM(Amount) AS Amount From dbo.tbl_Personnel_Payroll_Deductions WHERE (MasterKey = dbo.tbl_Personnel_Payroll.PK)), 0) AS Deductions " & _
    " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Payroll.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
    " Where (dbo.tbl_Personnel_Payroll.PayrollPeriodKey = " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ") " & _
    " And (dbo.tbl_Personnel_ActionNew.DivisionKey = " & cmbDivision.ItemData(cmbDivision.ListIndex) & ") " & _
    " And (dbo.tbl_Personnel_Position.PositionLevel = " & PostLevel & ") " & _
    " ORDER BY dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    
    MainForm.CommonDialog1.CancelError = True
    On Error GoTo ErrorHandler
    MainForm.CommonDialog1.DialogTitle = "Save"
    MainForm.CommonDialog1.Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls"
    MainForm.CommonDialog1.ShowSave
    Filename = Trim(MainForm.CommonDialog1.Filename)
    
    WorkbookName = CStr(Filename)
    
    On Error GoTo PG:
    
    Screen.MousePointer = vbHourglass
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
    xlsApp.Workbooks(1).Sheets(iWorkSheet).Name = "B P I"
    
    RowCnt = 0
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
    For k = 1 To 7
        Select Case k
            Case 1: strValue = "Account Number"
            Case 2: strValue = "First Name"
            Case 3: strValue = "Middle Name"
            Case 4: strValue = "Last Name"
            Case 5: strValue = "Total Earnings"
            Case 6: strValue = "Total Deduction"
            Case 7: strValue = "Amount" '"Allowance"
            'Case 9: strValue = "Amount"
        End Select
        ColCnt = k
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = strValue
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
        If k >= 5 And k <= 7 Then
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
        DoEvents
        iRec = iRec + 1
        i = i + 1
        strAmount = "="
        RowCnt = RowCnt + 1
        For k = 1 To 6
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
                'Case 7: strAmount = strAmount & "+" & strRange
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
        
        'UpdateProgress frmPersonnelCompensationReport.picProgressBar, i / ra.RecordCount
        UpdateProgress_No_Percent MainForm.picProgressBar, iRec / ra.RecordCount
        ra.MoveNext
    Wend
End If
ra.Close
MainForm.picProgressBar.BackColor = &H8000000F
DoEvents

SAVING:
On Error GoTo err_saving:
If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName
MainForm.picProgressBar.BackColor = &H8000000F
DoEvents
xlsApp.Visible = True
Screen.MousePointer = vbDefault

Exit Sub
ErrorHandler:
MainForm.picProgressBar.BackColor = &H8000000F
DoEvents
Screen.MousePointer = vbDefault
Exit Sub

Exit Sub
err_saving:
Screen.MousePointer = vbDefault
MainForm.picProgressBar.BackColor = &H8000000F
DoEvents
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING:

Exit Sub
PG:
MainForm.picProgressBar.BackColor = &H8000000F
DoEvents
Screen.MousePointer = vbDefault
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub TimerForATM2_Timer()
TimerForATM2.Enabled = False
If cmbDivision.ListIndex = -1 Then MsgBox "Please select division!                  ", vbCritical, "Error...": cmbDivision.SetFocus: Exit Sub
If cmbPeriodPrint.ListIndex = -1 Then MsgBox "Please select payroll date!                    ", vbCritical, "Error...": cmbPeriodPrint.SetFocus: Exit Sub
iRec = 0
i = 0
MainForm.picProgressBar.BackColor = &H8000000F
DoEvents
a = "SELECT dbo.tbl_Personnel_IDNumber.AccountNumber, dbo.tbl_Personnel_Information.LastName, " & _
    " dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName, " & _
    " ROUND(ISNULL((SELECT SUM(TotalAmount) AS Amount From dbo.tbl_Personnel_Payroll_Earnings WHERE (MasterKey = dbo.tbl_Personnel_Payroll.PK)), 0), 2) AS Earnings, " & _
    " ROUND(ISNULL((SELECT SUM(Amount) AS Amount From dbo.tbl_Personnel_Payroll_Deductions WHERE (MasterKey = dbo.tbl_Personnel_Payroll.PK)), 0), 2) AS Deductions " & _
    " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_IDNumber ON dbo.tbl_Personnel_Payroll.EmployeeKey = dbo.tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
    " Where (dbo.tbl_Personnel_Payroll.PayrollPeriodKey = " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ") " & _
    " And (dbo.tbl_Personnel_ActionNew.DivisionKey = " & cmbDivision.ItemData(cmbDivision.ListIndex) & ") " & _
    " And (dbo.tbl_Personnel_Position.PositionLevel = " & PostLevel & ") " & _
    " ORDER BY dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName"
If ra.State = adStateOpen Then ra.Close
ra.Open a, ConnOmega
If ra.RecordCount > 0 Then
    MainForm.CommonDialog1.CancelError = True
    On Error GoTo ErrorHandler
    MainForm.CommonDialog1.DialogTitle = "Save"
    MainForm.CommonDialog1.Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls"
    MainForm.CommonDialog1.ShowSave
    Filename = Trim(MainForm.CommonDialog1.Filename)
    
    WorkbookName = CStr(Filename)
    
    On Error GoTo PG:
    
    Screen.MousePointer = vbHourglass
    iWorkSheet = 1
    Set xlsApp = CreateObject("Excel.Application")
    xlsApp.Visible = False
    xlsApp.Workbooks.Add
    xlsApp.DisplayAlerts = False
    If xlsApp.Workbooks(1).Sheets.Count = 3 Then
        xlsApp.Workbooks(1).Sheets(2).Delete
        xlsApp.Workbooks(1).Sheets(2).Delete
    End If
    
    With xlsApp.Workbooks(1).Sheets(iWorkSheet)
        .Activate
        .Name = "Sheet1"
        
        RowCnt = 0
        RowCnt = RowCnt + 1
        ColCnt = 0
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = "H"
        .Range(strRange).Font.Name = "Calibri"
        .Range(strRange).Font.Size = 11
        .Range(strRange).Font.Bold = False
        .Range(strRange).HorizontalAlignment = 3
        
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = "Payroll Date"
        .Range(strRange).Font.Name = "Calibri"
        .Range(strRange).Font.Size = 11
        .Range(strRange).Font.Bold = True
        .Range(strRange).HorizontalAlignment = 3
        
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = cmbPeriodPrint.List(cmbPeriodPrint.ListIndex)
        .Range(strRange).Font.Name = "Calibri"
        .Range(strRange).Font.Size = 11
        .Range(strRange).Font.Bold = True
        .Range(strRange).HorizontalAlignment = 3
        
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = "Payroll Time"
        .Range(strRange).Font.Name = "Calibri"
        .Range(strRange).Font.Size = 11
        .Range(strRange).Font.Bold = True
        .Range(strRange).HorizontalAlignment = 3
        
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = ""
        .Range(strRange).Font.Name = "Calibri"
        .Range(strRange).Font.Size = 11
        .Range(strRange).Font.Bold = True
        .Range(strRange).HorizontalAlignment = 3
        
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = "Total Amount"
        .Range(strRange).Font.Name = "Calibri"
        .Range(strRange).Font.Size = 11
        .Range(strRange).Font.Bold = True
        .Range(strRange).HorizontalAlignment = 3
        
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = "=SUM(D3:D9994)"
        .Range(strRange).Font.Name = "Calibri"
        .Range(strRange).Font.Size = 11
        .Range(strRange).Font.Bold = True
        .Range(strRange).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        .Range(strRange).HorizontalAlignment = 3
        
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = "Total Amount"
        .Range(strRange).Font.Name = "Calibri"
        .Range(strRange).Font.Size = 11
        .Range(strRange).Font.Bold = True
        .Range(strRange).HorizontalAlignment = 3
        
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = "=COUNT(D3:D9994)"
        .Range(strRange).Font.Name = "Calibri"
        .Range(strRange).Font.Size = 11
        .Range(strRange).Font.Bold = True
        .Range(strRange).HorizontalAlignment = 3
        
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = "FUNDING ACCOUNT"
        .Range(strRange).Font.Name = "Calibri"
        .Range(strRange).Font.Size = 11
        .Range(strRange).Font.Bold = True
        .Range(strRange).HorizontalAlignment = 3
        
        '=====
        
        RowCnt = RowCnt + 1
        ColCnt = 0
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = "DETAIL CONSTANT"
        .Range(strRange).Font.Name = "Calibri"
        .Range(strRange).Font.Size = 11
        .Range(strRange).Font.Bold = True
        .Range(strRange).HorizontalAlignment = 3
        
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = "EMPLOYEE NAME"
        .Range(strRange).Font.Name = "Calibri"
        .Range(strRange).Font.Size = 11
        .Range(strRange).Font.Bold = True
        .Range(strRange).HorizontalAlignment = 3
        
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = "EMPLOYEE ACCOUNT"
        .Range(strRange).Font.Name = "Calibri"
        .Range(strRange).Font.Size = 11
        .Range(strRange).Font.Bold = True
        .Range(strRange).HorizontalAlignment = 3
        
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = "AMOUNT"
        .Range(strRange).Font.Name = "Calibri"
        .Range(strRange).Font.Size = 11
        .Range(strRange).Font.Bold = True
        .Range(strRange).HorizontalAlignment = 3
        
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = "REMARKS"
        .Range(strRange).Font.Name = "Calibri"
        .Range(strRange).Font.Size = 11
        .Range(strRange).Font.Bold = True
        .Range(strRange).HorizontalAlignment = 3
        
        While Not ra.EOF
            DoEvents
            iRec = iRec + 1
            dNetPay = CDbl(ra!Earnings) - CDbl(ra!Deductions)
            If CDbl(dNetPay) <> 0 Then
                RowCnt = RowCnt + 1
                ColCnt = 0
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = "D"
                .Range(strRange).Font.Name = "Calibri"
                .Range(strRange).Font.Size = 11
                .Range(strRange).Font.Bold = False
                .Range(strRange).HorizontalAlignment = 3
                
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = ra!FirstName & " " & ra!MiddleName & " " & ra!LastName
                .Range(strRange).Font.Name = "Calibri"
                .Range(strRange).Font.Size = 11
                .Range(strRange).Font.Bold = False
                .Range(strRange).HorizontalAlignment = 3
                
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = "'" & ra!AccountNumber
                .Range(strRange).Font.Name = "Calibri"
                .Range(strRange).Font.Size = 11
                .Range(strRange).Font.Bold = False
                .Range(strRange).HorizontalAlignment = 3
                
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = dNetPay 'CDbl(ra!Earnings) - CDbl(ra!Deductions)
                .Range(strRange).Font.Name = "Calibri"
                .Range(strRange).Font.Size = 11
                .Range(strRange).Font.Bold = False
                .Range(strRange).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
                .Range(strRange).HorizontalAlignment = 3
                
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = "payroll"
                .Range(strRange).Font.Name = "Calibri"
                .Range(strRange).Font.Size = 11
                .Range(strRange).Font.Bold = False
                .Range(strRange).HorizontalAlignment = 3
            End If
            UpdateProgress_No_Percent MainForm.picProgressBar, iRec / ra.RecordCount
            ra.MoveNext
        Wend
        
    End With
End If
ra.Close
MainForm.picProgressBar.BackColor = &H8000000F
DoEvents

SAVING:
On Error GoTo err_saving:
If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName
MainForm.picProgressBar.BackColor = &H8000000F
DoEvents
xlsApp.Visible = True
Screen.MousePointer = vbDefault

Exit Sub
ErrorHandler:
MainForm.picProgressBar.BackColor = &H8000000F
DoEvents
Screen.MousePointer = vbDefault
Exit Sub

Exit Sub
err_saving:
Screen.MousePointer = vbDefault
MainForm.picProgressBar.BackColor = &H8000000F
DoEvents
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING:

Exit Sub
PG:
MainForm.picProgressBar.BackColor = &H8000000F
DoEvents
Screen.MousePointer = vbDefault
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub TimerLoans_Timer()
TimerLoans.Enabled = False
If cmbDivision.ListIndex = -1 Then MsgBox "Please select division!                  ", vbCritical, "Error...": cmbDivision.SetFocus: Exit Sub
If cmbPeriodPrint.ListIndex = -1 Then MsgBox "Please select payroll date!                    ", vbCritical, "Error...": cmbPeriodPrint.SetFocus: Exit Sub
'If lstResultPrint.ListIndex = -1 Then MsgBox "Please select grouping!                    ", vbCritical, "Error...": lstResultPrint.SetFocus: Exit Sub
Dim dblLoans
dblLoans = 0
t = "SELECT ROUND(SUM(dbo.tbl_Personnel_Payroll_Deductions.Amount), 2) AS Amt " & _
    " FROM  dbo.tbl_Personnel_Payroll_Deductions LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Payroll ON dbo.tbl_Personnel_Payroll_Deductions.MasterKey = dbo.tbl_Personnel_Payroll.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK " & _
    " WHERE (dbo.tbl_Personnel_Payroll_Deductions.LoanKey IS NOT NULL) " & _
    " AND (dbo.tbl_Personnel_ActionNew.DivisionKey = " & cmbDivision.ItemData(cmbDivision.ListIndex) & ") " & _
    " AND (dbo.tbl_Personnel_Payroll.PayrollPeriodKey = " & cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex) & ") " & _
    " AND (dbo.tbl_Personnel_Position.PositionLevel = " & PostLevel & ")"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then
    dblLoans = IIf(IsNull(rt!Amt), 0, rt!Amt)
End If
rt.Close
If CDbl(dblLoans) = 0 Then MsgBox "No loan records!                     ", vbExclamation, "No data...": Exit Sub
Screen.MousePointer = vbHourglass
GenerateLoanContri gbl_UserName, cmbDivision.ItemData(cmbDivision.ListIndex), cmbPeriodPrint.ItemData(cmbPeriodPrint.ListIndex), lstReportType.ItemData(lstReportType.ListIndex), cmbDivision.List(cmbDivision.ListIndex), PostLevel, Month(cmbPeriodPrint.List(cmbPeriodPrint.ListIndex)), Year(cmbPeriodPrint.List(cmbPeriodPrint.ListIndex))
frmCrystalReportViewer.PRINT_PAYROLL_Loans gbl_UserName
If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
End Sub

Private Sub TimerPaySlip_Timer()
TimerPaySlip.Enabled = False
If cmbDivision.ListIndex = -1 Then MsgBox "Please select division!                  ", vbCritical, "Error...": cmbDivision.SetFocus: Exit Sub
If cmbPeriodPrint.ListIndex = -1 Then MsgBox "Please select payroll date!                    ", vbCritical, "Error...": cmbPeriodPrint.SetFocus: Exit Sub
If lstResultPrint.ListIndex = -1 Then MsgBox "Please select grouping!                    ", vbCritical, "Error...": lstResultPrint.SetFocus: Exit Sub
Screen.MousePointer = vbHourglass
If lstResultPrint.ItemData(lstResultPrint.ListIndex) = 0 Then
    GeneratePayslipSignLedger gbl_UserName, 3, cmbDivision.ItemData(cmbDivision.ListIndex), cmbPeriodPrint.List(cmbPeriodPrint.ListIndex), PostLevel, cmbDivision.ItemData(cmbDivision.ListIndex)
Else
    GeneratePayslipSignLedger gbl_UserName, 2, lstResultPrint.ItemData(lstResultPrint.ListIndex), cmbPeriodPrint.List(cmbPeriodPrint.ListIndex), PostLevel, cmbDivision.ItemData(cmbDivision.ListIndex)
End If
Screen.MousePointer = vbDefault
If lstReportType.ItemData(lstReportType.ListIndex) = 1 Then
    frmCrystalReportViewer.PRINT_PAYROLL_PAYSLIP_V3 gbl_UserName
ElseIf lstReportType.ItemData(lstReportType.ListIndex) = 2 Then
    frmCrystalReportViewer.PRINT_SIGNATURE_LEDGER_V2 gbl_UserName
End If
If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show
End Sub
