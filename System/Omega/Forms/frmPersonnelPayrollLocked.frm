VERSION 5.00
Begin VB.Form frmPersonnelPayrollLocked 
   Appearance      =   0  'Flat
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12855
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
   ScaleHeight     =   2010
   ScaleWidth      =   12855
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   240
      ScaleHeight     =   1725
      ScaleWidth      =   4455
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.ComboBox cmbPayrollDate 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   720
         Width           =   2775
      End
      Begin VB.ComboBox cmbLockedUnlocked 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   0
         Width           =   2775
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
         Picture         =   "frmPersonnelPayrollLocked.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1200
         Width           =   1560
      End
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
         Picture         =   "frmPersonnelPayrollLocked.frx":075C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   1560
      End
      Begin VB.ComboBox cmbDivision 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Locked / Unlocked"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   675
         Width           =   1575
      End
   End
   Begin VB.PictureBox picProgressBar 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   9435
      TabIndex        =   7
      Top             =   120
      Width           =   9495
   End
End
Attribute VB_Name = "frmPersonnelPayrollLocked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i, iLocked

Private Sub cmbDivision_Click()
If cmbDivision.ListIndex = -1 Then cmbPayrollDate.Clear: Exit Sub
If cmbLockedUnlocked.ListIndex = -1 Then cmbPayrollDate.Clear: Exit Sub

'MsgBox "Locked " & cmbLockedUnlocked.ItemData(cmbLockedUnlocked.ListIndex)
'MsgBox "Div " & cmbDivision.ItemData(cmbDivision.ListIndex)

cmbPayrollDate.Clear
s = "SELECT dbo.tbl_Personnel_Payroll.PayrollPeriodKey, dbo.tbl_Personnel_Compensation_Period.PayrollDate " & _
    " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK " & _
    " Where (dbo.tbl_Personnel_ActionNew.DivisionKey = " & cmbDivision.ItemData(cmbDivision.ListIndex) & ") " & _
    " And (dbo.tbl_Personnel_Payroll.Locked = " & cmbLockedUnlocked.ItemData(cmbLockedUnlocked.ListIndex) & ") " & _
    " GROUP BY dbo.tbl_Personnel_Payroll.PayrollPeriodKey, dbo.tbl_Personnel_Compensation_Period.PayrollDate " & _
    " ORDER BY dbo.tbl_Personnel_Compensation_Period.PayrollDate DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbPayrollDate.AddItem Format(rs!PayrollDate, "mm/dd/yyyy")
    cmbPayrollDate.ItemData(cmbPayrollDate.NewIndex) = rs!PayrollPeriodKey
    rs.MoveNext
Wend
rs.Close
If cmbPayrollDate.ListCount Then cmbPayrollDate.ListIndex = 0
End Sub

Private Sub cmbLockedUnlocked_Click()
If cmbDivision.ListIndex = -1 Then cmbPayrollDate.Clear: Exit Sub
If cmbLockedUnlocked.ListIndex = -1 Then cmbPayrollDate.Clear: Exit Sub


'MsgBox "Locked " & cmbLockedUnlocked.ItemData(cmbLockedUnlocked.ListIndex)
'MsgBox "Div " & cmbDivision.ItemData(cmbDivision.ListIndex)

cmbPayrollDate.Clear
s = "SELECT dbo.tbl_Personnel_Payroll.PayrollPeriodKey, dbo.tbl_Personnel_Compensation_Period.PayrollDate " & _
    " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Payroll.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK " & _
    " Where (dbo.tbl_Personnel_ActionNew.DivisionKey = " & cmbDivision.ItemData(cmbDivision.ListIndex) & ") " & _
    " And (dbo.tbl_Personnel_Payroll.Locked = " & cmbLockedUnlocked.ItemData(cmbLockedUnlocked.ListIndex) & ") " & _
    " GROUP BY dbo.tbl_Personnel_Payroll.PayrollPeriodKey, dbo.tbl_Personnel_Compensation_Period.PayrollDate " & _
    " ORDER BY dbo.tbl_Personnel_Compensation_Period.PayrollDate DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbPayrollDate.AddItem Format(rs!PayrollDate, "mm/dd/yyyy")
    cmbPayrollDate.ItemData(cmbPayrollDate.NewIndex) = rs!PayrollPeriodKey
    rs.MoveNext
Wend
rs.Close
If cmbPayrollDate.ListCount Then cmbPayrollDate.ListIndex = 0
End Sub

Private Sub cmdCancelAdd_Click()
Unload Me
End Sub

Private Sub cmdOKAdd_Click()
If cmbLockedUnlocked.ListIndex = -1 Then MsgBox "Please select transaction!                  ", vbCritical, "Error...": cmbLockedUnlocked.SetFocus: Exit Sub
If cmbDivision.ListIndex = -1 Then MsgBox "Please select division!                       ", vbCritical, "Error...": cmbDivision.SetFocus: Exit Sub
If cmbPayrollDate.ListIndex = -1 Then MsgBox "Please select payroll date!                    ", vbCritical, "Error...": cmbPayrollDate.SetFocus: Exit Sub
If MsgBox("CONTINUE TO " & UCase(cmbLockedUnlocked.List(cmbLockedUnlocked.ListIndex)) & " THOSE TRANSACTIONS IN " & UCase(cmbDivision.List(cmbDivision.ListIndex)) & " UNDER PAYROLL DATE " & UCase(cmbPayrollDate.List(cmbPayrollDate.ListIndex)), vbExclamation + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub

picProgressBar.BackColor = &HFFFFFF
picMain.Visible = False
picProgressBar.Visible = True
Me.Height = 1395
Me.Width = 9840
Me.Top = (MainForm.Height - Me.Height) / 5
Me.Left = (MainForm.Width - Me.Width) / 5
DoEvents
i = 0: iLocked = IIf(cmbLockedUnlocked.ItemData(cmbLockedUnlocked.ListIndex) = 0, 1, 0)
s = "SELECT dbo.tbl_Personnel_Payroll.PK, dbo.tbl_Personnel_ActionNew.DivisionKey, " & _
    " dbo.tbl_Personnel_Payroll.EmployeeKey, dbo.tbl_Personnel_Payroll.PayrollPeriodKey, " & _
    " dbo.tbl_Personnel_Payroll.ActionMemoKey, dbo.tbl_Personnel_Payroll.Locked " & _
    " FROM  dbo.tbl_Personnel_Payroll LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Payroll.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK " & _
    " WHERE (dbo.tbl_Personnel_ActionNew.DivisionKey = " & cmbDivision.ItemData(cmbDivision.ListIndex) & ") " & _
    " AND (dbo.tbl_Personnel_Payroll.PayrollPeriodKey = " & cmbPayrollDate.ItemData(cmbPayrollDate.ListIndex) & ") " & _
    " AND (dbo.tbl_Personnel_Payroll.Locked = " & cmbLockedUnlocked.ItemData(cmbLockedUnlocked.ListIndex) & ")"
If ra.State = adStateOpen Then ra.Close
rs.Open s, ConnOmega
While Not rs.EOF
    DoEvents
    i = i + 1
    
    t = "SELECT LoanKey, LoanBalance " & _
        " From dbo.tbl_Personnel_Payroll_Deductions " & _
        " WHERE (LoanKey IS NOT NULL) " & _
        " AND (LoanBalance IS NOT NULL) " & _
        " AND (MasterKey = " & rs!PK & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        If IIf(IsNull(rt!LoanBalance), 0, rt!LoanBalance) <= 0 Then
            ConnOmega.Execute "UPDATE tbl_Personnel_Loans " & _
                              " SET ZeroOut = 1 " & _
                              " WHERE (PK = " & rt!LoanKey & ")"
        End If
    End If
    rt.Close
    
    ConnOmega.Execute "UPDATE tbl_Personnel_Payroll " & _
                      " SET Locked = " & iLocked & " " & _
                      " WHERE (PK = " & rs!PK & ")"
                      
    UpdateProgress picProgressBar, i / rs.RecordCount
    rs.MoveNext
Wend
rs.Close
cmbLockedUnlocked_Click
picProgressBar.Visible = False
Me.Width = 5070
Me.Height = 2445 '2070
'Me.Caption = "Locked Payroll"
Me.Top = (MainForm.Height - Me.Height) / 3
Me.Left = (MainForm.Width - Me.Width) / 3
picMain.Visible = True
End Sub

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    If picProgressBar.Visible = True Then Exit Sub
    Unload Me
End If
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Width = 5070
Me.Height = 2445 '2070
Me.Top = (MainForm.Height - Me.Height) / 3
Me.Left = (MainForm.Width - Me.Width) / 3
With cmbLockedUnlocked
    .Clear
    .AddItem "Locked": .ItemData(.NewIndex) = 0
    .AddItem "Unlocked": .ItemData(.NewIndex) = 1
    .ListIndex = 0
End With
POPULATE_COMBO "PK", "Description", "tbl_Personnel_Division", "Description", cmbDivision
picMain.ZOrder 0
picMain.Visible = True
picProgressBar.Visible = False
End Sub
