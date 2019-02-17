VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAcctgChartOfAccounts 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "c"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAcctgChartOfAccounts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   12645
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   120
      ScaleHeight     =   6375
      ScaleWidth      =   12390
      TabIndex        =   0
      Top             =   120
      Width           =   12390
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Left            =   6480
         TabIndex        =   11
         Top             =   5640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtCurrCode 
         Height          =   285
         Left            =   4920
         TabIndex        =   10
         Top             =   5880
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtCurrPK 
         Height          =   285
         Left            =   3960
         TabIndex        =   9
         Top             =   5880
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   5640
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Timer TimerLoadChart 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   120
         Top             =   5880
      End
      Begin MSFlexGridLib.MSFlexGrid FGrid 
         Height          =   5540
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   12390
         _ExtentX        =   21855
         _ExtentY        =   9763
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   13023396
         ForeColorFixed  =   255
         BackColorSel    =   16777088
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         GridColor       =   0
         FocusRect       =   0
         HighLight       =   2
         GridLinesFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "BALANCE >>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   8880
         TabIndex        =   7
         Top             =   6120
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CREDIT >>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   8880
         TabIndex        =   6
         Top             =   5880
         Width           =   1095
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10080
         TabIndex        =   5
         Top             =   6120
         Width           =   2055
      End
      Begin VB.Label lblCredit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10080
         TabIndex        =   4
         Top             =   5880
         Width           =   2055
      End
      Begin VB.Label lblDebit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10080
         TabIndex        =   3
         Top             =   5640
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "DEBIT >>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   8880
         TabIndex        =   2
         Top             =   5640
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmAcctgChartOfAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i, j, dDebit, dCredit, dBalance, dTotDebit, dTotCredit, dTotBalance, CurrPK, _
HEADER1$

Private Sub CUSTOM_GRID_CASHIER()
With FGrid
    .Clear
    HEADER1$ = HEADER1$ & "|" & "CODE" & "|" & "NAME" & "|" & "DEBIT" & "|" & "CREDIT" & "|" & "BALANCE" & "|" & "PK"
    .FormatString = HEADER1$
    .ColWidth(1) = 1500      'Code
    .ColWidth(2) = 6000     'Name
    .ColWidth(3) = 1500     'Debit
    .ColWidth(4) = 1500     'Credit
    .ColWidth(5) = 1500     'Balance
    .ColWidth(6) = 0        'PK
    .ColAlignment(1) = 1 '3 'flexAlignRightCenter
    .ColAlignment(2) = 1 'flexAlignRightCenter
    .ColAlignment(3) = flexAlignRightCenter '3 '1
    .ColAlignment(4) = flexAlignRightCenter '3 '1
    .ColAlignment(5) = flexAlignRightCenter '3 '1
    .ColAlignment(6) = 1
    .Rows = 2
End With
End Sub

Private Sub LOAD_CHART_OF_ACCOUNT()
Screen.MousePointer = vbHourglass
dTotDebit = 0: dTotCredit = 0: dTotBalance = 0
With FGrid
    i = 0
    s = "SELECT PK, Code, DeptName " & _
        " From tbl_GL_Department " & _
        " ORDER BY Code"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        'DoEvents
        dDebit = 0: dCredit = 0: dBalance = 0
        i = i + 1
        .Rows = i + 1
        .TextMatrix(i, 1) = rs!Code
        .TextMatrix(i, 2) = UCase(rs!DeptName)
        .TextMatrix(i, 3) = ""
        .TextMatrix(i, 4) = ""
        .TextMatrix(i, 5) = ""
        .TextMatrix(i, 6) = "0"
        .ColWidth(1) = 1500      'Code
        .ColWidth(2) = 6000     'Name
        .ColWidth(3) = 1500     'Debit
        .ColWidth(4) = 1500     'Credit
        .ColWidth(5) = 1500     'Balance
        .ColWidth(6) = 0        'PK
        .ColAlignment(1) = 1 '3 'flexAlignRightCenter
        .ColAlignment(2) = 1 'flexAlignRightCenter
        .ColAlignment(3) = flexAlignRightCenter '3 '1
        .ColAlignment(4) = flexAlignRightCenter '3 '1
        .ColAlignment(5) = flexAlignRightCenter '3 '1
        .ColAlignment(6) = 1
        .ROW = i
        For j = 1 To 6
            .Col = j
            .CellFontBold = True
        Next j
        t = "SELECT PK, AccountCode, AccountName, Debit, Credit, Balance " & _
            " From tbl_GL_Accounts " & _
            " Where (Dept = " & rs!PK & ") " & _
            " ORDER BY AccountCode"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        While Not rt.EOF
            'DoEvents
            i = i + 1
            dDebit = dDebit + CDbl(rt!Debit)
            dCredit = dCredit + CDbl(rt!Credit)
            dBalance = dBalance + CDbl(rt!Balance)
            
            dTotDebit = dTotDebit + CDbl(rt!Debit)
            dTotCredit = dTotCredit + CDbl(rt!Credit)
            dTotBalance = dTotBalance + CDbl(rt!Balance)
            .Rows = i + 1
            .TextMatrix(i, 1) = "     " & rt!AccountCode
            .TextMatrix(i, 2) = "     " & rt!AccountName
            .TextMatrix(i, 3) = Format(rt!Debit, "#,##0.00")
            .TextMatrix(i, 4) = Format(rt!Credit, "#,##0.00")
            .TextMatrix(i, 5) = Format(rt!Balance, "#,##0.00")
            .TextMatrix(i, 6) = rt!PK
            .ColWidth(1) = 1500      'Code
            .ColWidth(2) = 6000     'Name
            .ColWidth(3) = 1500     'Debit
            .ColWidth(4) = 1500     'Credit
            .ColWidth(5) = 1500     'Balance
            .ColWidth(6) = 0        'PK
            .ColAlignment(1) = 1 '3 'flexAlignRightCenter
            .ColAlignment(2) = 1 'flexAlignRightCenter
            .ColAlignment(3) = flexAlignRightCenter '3 '1
            .ColAlignment(4) = flexAlignRightCenter '3 '1
            .ColAlignment(5) = flexAlignRightCenter '3 '1
            .ColAlignment(6) = 1
            .ROW = i
            For j = 1 To 6
                .Col = j
                .CellFontBold = False
            Next j
            rt.MoveNext
        Wend
        rt.Close
        
        'DoEvents
        i = i + 1
        .Rows = i + 1
        .TextMatrix(i, 1) = ""
        .TextMatrix(i, 2) = UCase(rs!DeptName) & " TOTAL >>"
        .TextMatrix(i, 3) = Format(dDebit, "#,##0.00")
        .TextMatrix(i, 4) = Format(dCredit, "#,##0.00")
        .TextMatrix(i, 5) = Format(dBalance, "#,##0.00")
        .TextMatrix(i, 6) = "0"
        .ColWidth(1) = 1500     'Code
        .ColWidth(2) = 6000     'Name
        .ColWidth(3) = 1500     'Debit
        .ColWidth(4) = 1500     'Credit
        .ColWidth(5) = 1500     'Balance
        .ColWidth(6) = 0        'PK
        .ColAlignment(1) = 1 '3 'flexAlignRightCenter
        .ColAlignment(2) = 1 'flexAlignRightCenter
        .ColAlignment(3) = flexAlignRightCenter '3 '1
        .ColAlignment(4) = flexAlignRightCenter '3 '1
        .ColAlignment(5) = flexAlignRightCenter '3 '1
        .ColAlignment(6) = 1
        .ROW = i
        For j = 1 To 6
            .Col = j
            .CellFontBold = True
        Next j
        rs.MoveNext
    Wend
    rs.Close
    
'    i = i + 1
'    .Rows = i + 1
'    .TextMatrix(i, 1) = ""
'    .TextMatrix(i, 2) = "GRAND TOTAL >>"
'    .TextMatrix(i, 3) = Format(dTotDebit, "#,##0.00")
'    .TextMatrix(i, 4) = Format(dTotCredit, "#,##0.00")
'    .TextMatrix(i, 5) = Format(dTotBalance, "#,##0.00")
'    .TextMatrix(i, 6) = "0"
'    .ColWidth(1) = 1500     'Code
'    .ColWidth(2) = 6000     'Name
'    .ColWidth(3) = 1500     'Debit
'    .ColWidth(4) = 1500     'Credit
'    .ColWidth(5) = 1500     'Balance
'    .ColWidth(6) = 0        'PK
'    .ColAlignment(1) = 1 '3 'flexAlignRightCenter
'    .ColAlignment(2) = 1 'flexAlignRightCenter
'    .ColAlignment(3) = flexAlignRightCenter '3 '1
'    .ColAlignment(4) = flexAlignRightCenter '3 '1
'    .ColAlignment(5) = flexAlignRightCenter '3 '1
'    .ColAlignment(6) = 1
'    .ROW = i
'    For j = 1 To 6
'        .Col = j
'        .CellFontBold = True
'    Next j
    
    lblDebit.Caption = Format(dTotDebit, "#,##0.00")
    lblCredit.Caption = Format(dTotCredit, "#,##0.00")
    lblBalance.Caption = Format(dTotBalance, "#,##0.00")
    
End With
Screen.MousePointer = vbDefault
End Sub

Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
s = "SELECT tbl_GL_Accounts.* " & _
    " FROM tbl_GL_Accounts"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    ConnOmega.Execute "UPDATE tbl_GL_Accounts " & _
                      " SET AccountName = '" & FORMATSQL(UCase(rs!AccountName)) & "' " & _
                      " WHERE (PK = " & rs!PK & ")"
    rs.MoveNext
Wend
rs.Close
Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Screen.MousePointer = vbHourglass
s = "SELECT SupplierSL.* " & _
    " FROM SupplierSL"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    t = "SELECT tbl_Inv_Supplier.* " & _
        " FROM tbl_Inv_Supplier " & _
        " WHERE (SupplierCode = '" & rs!SuppCode & "')"
    If rt.State = adStateOpen Then rs.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        ConnOmega.Execute "INSERT INTO tbl_Inv_Supplier_SL " & _
                          " (SupplierKey, GLCode, DocNumber, DocDate, Debit, Credit, iType, Reference) " & _
                          " VALUES (" & rt!PK & ", '" & rs!GLCode & "', 'END_BAL_SEP12', " & _
                          " '09/30/2012', " & CDbl(rs!Credit) & ", " & CDbl(rs!Debit) & ",  0, 'Ending Balance Sept 2012')"
    End If
    rt.Close
    rs.MoveNext
Wend
rs.Close
Screen.MousePointer = vbDefault
End Sub

Private Sub FGrid_DblClick()
'If RETURNTEXTVALUE(txtCurrPK) = 0 Then Exit Sub
'SaveSetting App.EXEName, "ChartOfAccount", "ChartOfAccount", txtCurrCode.Text
'If IsLoaded(frmAcctgChartOfAccountsCard) Then frmAcctgChartOfAccountsCard.ZOrder 0 Else frmAcctgChartOfAccountsCard.Show
'frmAcctgChartOfAccountsCard.BROWSER GetSetting(App.EXEName, "ChartOfAccount", "ChartOfAccount", ""), "is_LOAD"
End Sub

Private Sub FGrid_EnterCell()
txtCurrPK.Text = FGrid.TextMatrix(FGrid.ROW, 6)
txtCurrCode.Text = Trim(FGrid.TextMatrix(FGrid.ROW, 1))
End Sub

Private Sub FGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu MainFormPopupF.mnuChartOfAccounts
End If
End Sub

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
'LOAD_CHART_OF_ACCOUNT
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
'Me.Caption = "Chart of Accounts"
txtCurrPK.Text = ""
'DoEvents
CUSTOM_GRID_CASHIER
LOAD_CHART_OF_ACCOUNT
'TimerLoadChart.Enabled = True
End Sub

Private Sub TimerLoadChart_Timer()
TimerLoadChart.Enabled = False
LOAD_CHART_OF_ACCOUNT
End Sub
