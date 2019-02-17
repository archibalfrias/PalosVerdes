VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInvSupplierSL 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvSupplierSL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   13710
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimerLoad 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1200
      Top             =   4440
   End
   Begin VB.TextBox txtSupplier 
      Height          =   315
      Left            =   960
      MaxLength       =   50
      TabIndex        =   1
      Top             =   120
      Width           =   6920
   End
   Begin MSComctlLib.ListView lstSL 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   7011
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Account Code"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Account Name"
         Object.Width           =   5380
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Reference"
         Object.Width           =   6880
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Debit"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Credit"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Balance"
         Object.Width           =   2646
      EndProperty
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   12000
      TabIndex        =   6
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label lblTotalDebit 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9480
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label lblTotalCredit 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   10800
      TabIndex        =   4
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL >>"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8520
      TabIndex        =   3
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmInvSupplierSL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iSupplier

Dim x, i, dDebit, dCredit, dBalance

Private Sub Form_Activate()
TimerLoad.Enabled = True
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Caption = "Subsidiary Ledger"
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
txtSupplier.Locked = True
TimerLoad.Enabled = True
End Sub

Private Sub TimerLoad_Timer()
TimerLoad.Enabled = False
lstSL.ListItems.Clear: i = 0
dDebit = 0: dCredit = 0: dBalance = 0
's = "SELECT tbl_Inv_Supplier_SL.* " & _
    " FROM tbl_Inv_Supplier_SL " & _
    " WHERE (SupplierKey = " & iSupplier & ")" & _
    " ORDER BY DocDate"
s = "SELECT tbl_Inv_Supplier_SL.DocDate, " & _
    " tbl_Inv_Supplier_SL.GLCode, " & _
    " tbl_GL_Accounts.AccountName, " & _
    " tbl_Inv_Supplier_SL.Reference, " & _
    " tbl_Inv_Supplier_SL.Debit, " & _
    " tbl_Inv_Supplier_SL.Credit, " & _
    " tbl_Inv_Supplier_SL.Balance " & _
    " FROM tbl_Inv_Supplier_SL LEFT OUTER JOIN " & _
    " tbl_GL_Accounts ON tbl_Inv_Supplier_SL.GLCode = tbl_GL_Accounts.AccountCode " & _
    " Where (tbl_Inv_Supplier_SL.SupplierKey = " & iSupplier & ") " & _
    " ORDER BY tbl_Inv_Supplier_SL.DocDate"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    i = i + 1
    Set x = lstSL.ListItems.Add()
    x.Text = ""
    dBalance = dBalance + CDbl(rs!Balance)
    x.SubItems(1) = Format(rs!DocDate, "mm/dd/yyyy")
    x.SubItems(2) = IIf(IsNull(rs!GLCode), " ", rs!GLCode)
    x.SubItems(3) = IIf(IsNull(rs!AccountName), " ", rs!AccountName)
    x.SubItems(4) = IIf(IsNull(rs!Reference), " ", rs!Reference)
    x.SubItems(5) = IIf(CDbl(rs!Debit) > 0, Format(rs!Debit, "#,##0.00"), " ")
    x.SubItems(6) = IIf(CDbl(rs!Credit) > 0, Format(rs!Credit, "#,##0.00"), " ")
    x.SubItems(7) = Format(dBalance, "#,##0.00")
    dDebit = dDebit + CDbl(rs!Debit)
    dCredit = dCredit + CDbl(rs!Credit)
    
'    If rs!iType = 1 Then
'        t = "SELECT tbl_GL_Accounts.* " & _
'            " FROM tbl_GL_Accounts " & _
'            " WHERE (AccountCode = '" & rs!GLCode & "')"
'        If rt.State = adStateOpen Then rt.Close
'        rt.Open t, ConnOmega
'        x.SubItems(2) = rt!AccountName
'        rt.Close
'    ElseIf rs!iType = 2 Then
'        t = "SELECT tbl_Inv_Supplier.* " & _
'            " FROM tbl_Inv_Supplier " & _
'            " WHERE (SupplierCode = '" & rs!GLCode & "')"
'        If rt.State = adStateOpen Then rt.Close
'        rt.Open t, ConnOmega
'        x.SubItems(2) = rt!SupplierName
'        rt.Close
'    End If
    
    rs.MoveNext
Wend
rs.Close
lblTotalDebit.Caption = Format(dDebit, "#,##0.00")
lblTotalCredit.Caption = Format(dCredit, "#,##0.00")
lblBalance.Caption = Format(dBalance, "#,##0.00")
End Sub
