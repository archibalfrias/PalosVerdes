VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAcctgChartOfAccountsSL 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14625
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAcctgChartOfAccountsSL.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   14625
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdQuery 
      Caption         =   "Query"
      Height          =   735
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtDateTo 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtDateFrom 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Timer TimerLoadSL 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   14040
      Top             =   120
   End
   Begin MSComctlLib.ListView lstSL 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   10398
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Book"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Doc Number"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Supplier Code"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Supplier Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Bank Code"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Bank"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "CheckNumber"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "CheckDate"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Particulars"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Inv Number"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Inv Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "Debit"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "Credit"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Text            =   "Balance"
         Object.Width           =   2293
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date To"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date From"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmAcctgChartOfAccountsSL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sFormCaption As String
Public sAccCode As String

Dim i, x, Arr, dRunningBal

Private Sub cmdQuery_Click()
txtDateFrom.SetFocus
TimerLoadSL.Enabled = True
End Sub

Private Sub Form_Activate()
Me.Caption = sFormCaption
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
Me.Caption = sFormCaption

End Sub

Private Sub Form_Resize()
On Error Resume Next
lstSL.Top = 960
lstSL.Left = 0 '120
lstSL.Width = Me.ScaleWidth - lstSL.Left
lstSL.Height = Me.ScaleHeight - lstSL.Top
End Sub

Private Sub TimerLoadSL_Timer()
Screen.MousePointer = vbHourglass
TimerLoadSL.Enabled = False
lstSL.ListItems.Clear: dRunningBal = 0
If IsDate(txtDateFrom.Text) = False And _
IsDate(txtDateTo.Text) = False Then
    s = "SELECT tbl_GL_Transaction.DocDate, tbl_Acctg_Book.Abb, tbl_GL_Transaction.DocNumber, " & _
        " tbl_GL_Transaction.SupplierCode, tbl_GL_Transaction.SupplierName, tbl_GL_Transaction.BankCode, " & _
        " tbl_GL_Transaction.BankName, tbl_GL_Transaction.CVNumber, tbl_GL_Transaction.CheckNumber, " & _
        " tbl_GL_Transaction.CheckDate, tbl_GL_Transaction.Particulars, tbl_GL_Transaction.InvoiceNumber, " & _
        " tbl_GL_Transaction.InvoiceDate, tbl_GL_Transaction.DeptCode, tbl_GL_Transaction.Debit, " & _
        " tbl_GL_Transaction.Credit, tbl_GL_Transaction.Balance " & _
        " FROM tbl_GL_Transaction LEFT OUTER JOIN " & _
        " tbl_Acctg_Book ON tbl_GL_Transaction.BookType = tbl_Acctg_Book.PK " & _
        " WHERE (tbl_GL_Transaction.GLCode = '" & sAccCode & "') " & _
        " ORDER BY tbl_GL_Transaction.DocDate"
ElseIf IsDate(txtDateFrom.Text) = False And _
IsDate(txtDateTo.Text) = True Then
    s = "SELECT tbl_GL_Transaction.DocDate, tbl_Acctg_Book.Abb, tbl_GL_Transaction.DocNumber, " & _
        " tbl_GL_Transaction.SupplierCode, tbl_GL_Transaction.SupplierName, tbl_GL_Transaction.BankCode, " & _
        " tbl_GL_Transaction.BankName, tbl_GL_Transaction.CVNumber, tbl_GL_Transaction.CheckNumber, " & _
        " tbl_GL_Transaction.CheckDate, tbl_GL_Transaction.Particulars, tbl_GL_Transaction.InvoiceNumber, " & _
        " tbl_GL_Transaction.InvoiceDate, tbl_GL_Transaction.DeptCode, tbl_GL_Transaction.Debit, " & _
        " tbl_GL_Transaction.Credit, tbl_GL_Transaction.Balance " & _
        " FROM tbl_GL_Transaction LEFT OUTER JOIN " & _
        " tbl_Acctg_Book ON tbl_GL_Transaction.BookType = tbl_Acctg_Book.PK " & _
        " WHERE (tbl_GL_Transaction.GLCode = '" & sAccCode & "') " & _
        " AND (tbl_GL_Transaction.DocDate <= '" & FormatDateTime(txtDateTo.Text, vbShortDate) & "') " & _
        " ORDER BY tbl_GL_Transaction.DocDate"
ElseIf IsDate(txtDateFrom.Text) = True And _
IsDate(txtDateTo.Text) = False Then
    s = "SELECT tbl_GL_Transaction.DocDate, tbl_Acctg_Book.Abb, tbl_GL_Transaction.DocNumber, " & _
        " tbl_GL_Transaction.SupplierCode, tbl_GL_Transaction.SupplierName, tbl_GL_Transaction.BankCode, " & _
        " tbl_GL_Transaction.BankName, tbl_GL_Transaction.CVNumber, tbl_GL_Transaction.CheckNumber, " & _
        " tbl_GL_Transaction.CheckDate, tbl_GL_Transaction.Particulars, tbl_GL_Transaction.InvoiceNumber, " & _
        " tbl_GL_Transaction.InvoiceDate, tbl_GL_Transaction.DeptCode, tbl_GL_Transaction.Debit, " & _
        " tbl_GL_Transaction.Credit, tbl_GL_Transaction.Balance " & _
        " FROM tbl_GL_Transaction LEFT OUTER JOIN " & _
        " tbl_Acctg_Book ON tbl_GL_Transaction.BookType = tbl_Acctg_Book.PK " & _
        " WHERE (tbl_GL_Transaction.GLCode = '" & sAccCode & "') " & _
        " AND (tbl_GL_Transaction.DocDate >= '" & FormatDateTime(txtDateFrom.Text, vbShortDate) & "') " & _
        " ORDER BY tbl_GL_Transaction.DocDate"
ElseIf IsDate(txtDateFrom.Text) = True And _
IsDate(txtDateTo.Text) = True Then
    s = "SELECT tbl_GL_Transaction.DocDate, tbl_Acctg_Book.Abb, tbl_GL_Transaction.DocNumber, " & _
        " tbl_GL_Transaction.SupplierCode, tbl_GL_Transaction.SupplierName, tbl_GL_Transaction.BankCode, " & _
        " tbl_GL_Transaction.BankName, tbl_GL_Transaction.CVNumber, tbl_GL_Transaction.CheckNumber, " & _
        " tbl_GL_Transaction.CheckDate, tbl_GL_Transaction.Particulars, tbl_GL_Transaction.InvoiceNumber, " & _
        " tbl_GL_Transaction.InvoiceDate, tbl_GL_Transaction.DeptCode, tbl_GL_Transaction.Debit, " & _
        " tbl_GL_Transaction.Credit, tbl_GL_Transaction.Balance " & _
        " FROM tbl_GL_Transaction LEFT OUTER JOIN " & _
        " tbl_Acctg_Book ON tbl_GL_Transaction.BookType = tbl_Acctg_Book.PK " & _
        " WHERE (tbl_GL_Transaction.GLCode = '" & sAccCode & "') " & _
        " AND (tbl_GL_Transaction.DocDate >= '" & FormatDateTime(txtDateFrom.Text, vbShortDate) & "') " & _
        " AND (tbl_GL_Transaction.DocDate <= '" & FormatDateTime(txtDateTo.Text, vbShortDate) & "') " & _
        " ORDER BY tbl_GL_Transaction.DocDate"
End If
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    Set x = lstSL.ListItems.Add()
    x.Text = ""
    x.SubItems(1) = Format(rs!DocDate, "mm/dd/yyyy")
    x.SubItems(2) = rs!Abb
    x.SubItems(3) = rs!DocNumber
    x.SubItems(4) = IIf(IsNull(rs!SupplierCode), " ", rs!SupplierCode)
    x.SubItems(5) = IIf(IsNull(rs!SupplierName), " ", rs!SupplierName)
    x.SubItems(6) = IIf(IsNull(rs!BankCode), " ", rs!BankCode)
    x.SubItems(7) = IIf(IsNull(rs!BankName), " ", rs!BankName)
    x.SubItems(8) = IIf(IsNull(rs!CheckNumber), " ", rs!CheckNumber)
    If IsNull(rs!CheckDate) = True Then
        x.SubItems(9) = " "
    Else
        x.SubItems(9) = Format(rs!CheckDate, "mm/dd/yyyy")
    End If
    x.SubItems(10) = IIf(IsNull(rs!Particulars), " ", rs!Particulars)
    x.SubItems(11) = IIf(IsNull(rs!InvoiceNumber), " ", rs!InvoiceNumber)
    If IsNull(rs!InvoiceDate) = True Then
        x.SubItems(12) = " "
    Else
        x.SubItems(12) = Format(rs!InvoiceDate, "mm/dd/yyyy")
    End If
    x.SubItems(13) = IIf(CDbl(rs!Debit) = 0, " ", Format(rs!Debit, "#,##0.00"))
    x.SubItems(14) = IIf(CDbl(rs!Credit) = 0, " ", Format(rs!Credit, "#,##0.00"))
    dRunningBal = dRunningBal + CDbl(rs!Balance)
    x.SubItems(15) = Format(dRunningBal, "#,##0.00")
    rs.MoveNext
Wend
rs.Close
Screen.MousePointer = vbDefault
End Sub

Private Sub txtDateFrom_LostFocus()
If IsDate(txtDateFrom.Text) = True Then
    txtDateFrom.Text = Format(FormatDateTime(txtDateFrom.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub

Private Sub txtDateTo_LostFocus()
If IsDate(txtDateTo.Text) = True Then
    txtDateTo.Text = Format(FormatDateTime(txtDateTo.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub
