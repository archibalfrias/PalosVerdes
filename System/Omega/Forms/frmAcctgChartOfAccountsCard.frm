VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAcctgChartOfAccountsCard 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAcctgChartOfAccountsCard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   240
      ScaleHeight     =   2895
      ScaleWidth      =   7215
      TabIndex        =   3
      Top             =   960
      Width           =   7215
      Begin VB.TextBox txtSupplierKey 
         Height          =   315
         Left            =   3720
         TabIndex        =   24
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtSuppCode 
         Height          =   315
         Left            =   2040
         TabIndex        =   23
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtSuppName 
         Height          =   315
         Left            =   3600
         TabIndex        =   22
         Top             =   1080
         Width           =   3615
      End
      Begin VB.PictureBox picInventory 
         BackColor       =   &H00C6B8A4&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1200
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   20
         Top             =   1440
         Width           =   495
         Begin VB.CheckBox chkInventory 
            BackColor       =   &H00C6B8A4&
            Height          =   255
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.TextBox txtBalance 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         TabIndex        =   18
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         TabIndex        =   17
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txtDebit 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         TabIndex        =   16
         Top             =   1800
         Width           =   2295
      End
      Begin VB.PictureBox picWSL 
         BackColor       =   &H00C6B8A4&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1200
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   11
         Top             =   1080
         Width           =   495
         Begin VB.CheckBox chkWSL 
            BackColor       =   &H00C6B8A4&
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.ComboBox cmbAccountGroup 
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   720
         Width           =   6015
      End
      Begin VB.TextBox txtAccountName 
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   6015
      End
      Begin VB.TextBox txtAccountCode 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   0
         Width           =   2295
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Debit"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "With Subsidiary"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Group"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Code"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7920
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483648
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgChartOfAccountsCard.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgChartOfAccountsCard.frx":09CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgChartOfAccountsCard.frx":0B50
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgChartOfAccountsCard.frx":0E6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgChartOfAccountsCard.frx":1223
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgChartOfAccountsCard.frx":1675
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgChartOfAccountsCard.frx":1AC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgChartOfAccountsCard.frx":1E7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgChartOfAccountsCard.frx":1F91
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgChartOfAccountsCard.frx":24D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgChartOfAccountsCard.frx":262D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctgChartOfAccountsCard.frx":2B6F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   770
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   15000
      TabIndex        =   0
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   570
         Left            =   0
         TabIndex        =   1
         Top             =   105
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   1005
         ButtonWidth     =   1058
         ButtonHeight    =   1005
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   20
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Add"
               Key             =   "Add"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Edit"
               Key             =   "Edit"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Delete"
               Key             =   "Delete"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "First"
               Key             =   "First"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Back"
               Key             =   "Back"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Next"
               Key             =   "Next"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Last"
               Key             =   "Last"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Find"
               Key             =   "Find"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Print"
               Key             =   "Print"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   15000
         Y1              =   690
         Y2              =   690
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   15000
         Y1              =   90
         Y2              =   90
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   15000
         Y1              =   750
         Y2              =   750
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   4065
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   26458
            MinWidth        =   26458
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAcctgChartOfAccountsCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public AccountPK    As Long

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim tmp As Long

Dim iAccountGroup, iSupplier

Private Sub PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
txtAccountCode.SetFocus
End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1) = "" Then Exit Sub
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_EDITTING
End Sub

Private Sub PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1) = "" Then Exit Sub

End Sub

Private Sub PRESS_F5()
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then

End If
If TRANSACTIONTYPE = is_EDITTING Then
    ConnOmega.Execute "UPDATE tbl_GL_Accounts " & _
                      " SET AccountName = '" & FORMATSQL(Trim(txtAccountName.Text)) & "', " & _
                      " Dept = " & iAccountGroup & ", " & _
                      " withSL = " & chkWSL.Value & ", " & _
                      " SupplierKey = " & RETURNTEXTVALUE(txtSupplierKey) & ", " & _
                      " Inventory = " & chkInventory.Value & " " & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
End If
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "ChartOfAccount", "ChartOfAccount", ""), "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    Unload Me
Else
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "ChartOfAccount", "ChartOfAccount", ""), "is_LOAD"
    If Trim(txtAccountCode.Text) = "" Then BROWSER GetSetting(App.EXEName, "ChartOfAccount", "ChartOfAccount", ""), "is_HOME"
End If
End Sub

Public Sub BROWSER(sAccCode, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If sAccCode <> "" Then
            s = "SELECT TOP 1 tbl_GL_Accounts.* " & _
                " FROM tbl_GL_Accounts " & _
                " WHERE (AccountCode = '" & sAccCode & "') " & _
                " ORDER BY AccountCode"
        Else
            s = "SELECT TOP 1 tbl_GL_Accounts.* " & _
                " FROM tbl_GL_Accounts " & _
                " ORDER BY AccountCode"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_GL_Accounts.* " & _
            " FROM tbl_GL_Accounts " & _
            " ORDER BY AccountCode"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_GL_Accounts.* " & _
            " FROM tbl_GL_Accounts " & _
            " WHERE (AccountCode < '" & sAccCode & "') " & _
            " ORDER BY AccountCode DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_GL_Accounts.* " & _
            " FROM tbl_GL_Accounts " & _
            " WHERE (AccountCode > '" & sAccCode & "') " & _
            " ORDER BY AccountCode"
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_GL_Accounts.* " & _
            " FROM tbl_GL_Accounts " & _
            " ORDER BY AccountCode DESC"
    Case "is_FIND"
        s = "SELECT TOP 1 tbl_GL_Accounts.* " & _
            " FROM tbl_GL_Accounts " & _
            " WHERE (PK = " & sAccCode & ") "
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iAccountGroup = rs!Dept
    iSupplier = rs!SupplierKey
    txtAccountCode.Text = rs!AccountCode
    txtAccountName.Text = rs!AccountName
    
    txtSuppCode.Text = ""
    txtSuppName.Text = ""
    t = "SELECT tbl_Inv_Supplier.* " & _
        " FROM tbl_Inv_Supplier " & _
        " WHERE (PK = " & rs!SupplierKey & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtSuppCode.Text = rt!SupplierCode
        txtSuppName.Text = rt!SupplierName
    End If
    rt.Close
    
    cmbAccountGroup.Text = ""
    t = "SELECT tbl_GL_Department.* " & _
        " FROM tbl_GL_Department " & _
        " WHERE (PK = " & rs!Dept & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        cmbAccountGroup.Text = UCase(rt!DeptName)
    End If
    rt.Close
    
    chkWSL.Value = rs!withSL
    chkInventory.Value = rs!Inventory
    txtDebit.Text = Format(rs!Debit, "#,##0.00")
    txtCredit.Text = Format(rs!Credit, "#,##0.00")
    txtBalance.Text = Format(rs!Balance, "#,##0.00")
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    
    SaveSetting App.EXEName, "ChartOfAccount", "ChartOfAccount", rs!AccountCode
    
End If
rs.Close
End Sub

Private Sub CLEARTEXT()
iAccountGroup = 0
iSupplier = 0
txtAccountCode.Text = ""
txtAccountName.Text = ""
txtSuppCode.Text = ""
txtSuppName.Text = ""
cmbAccountGroup.Text = ""
cmbAccountGroup.ListIndex = -1
chkWSL.Value = 0
chkInventory.Value = 0
txtDebit.Text = ""
txtCredit.Text = ""
txtBalance.Text = ""
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtAccountCode.Locked = bln
txtAccountName.Locked = bln
cmbAccountGroup.Locked = bln
picWSL.Enabled = IIf(bln = True, False, True)
picInventory.Enabled = IIf(bln = True, False, True)
txtDebit.Locked = True
txtCredit.Locked = True
txtBalance.Locked = True
End Sub

Public Sub TOOLBARFUNC(intTrans As Integer)
Set Toolbar1.ImageList = ImageList1
With Toolbar1
    Select Case intTrans
        Case 1  '==== REFRESH ====
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(1).Enabled = True
            .Buttons(3).Enabled = True
            .Buttons(5).Enabled = True
            .Buttons(7).Image = 4
            .Buttons(7).Caption = "First"
            .Buttons(9).Image = 5
            .Buttons(9).Caption = "Back"
            .Buttons(7).Enabled = True
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = True
            .Buttons(13).Enabled = True
            .Buttons(15).Enabled = True
            .Buttons(17).Enabled = True
            .Buttons(19).Enabled = True
            .Buttons(1).ToolTipText = "NEW (Ins)"
            .Buttons(3).ToolTipText = "EDIT (F2)"
            .Buttons(5).ToolTipText = "DELETE (Del)"
            .Buttons(7).ToolTipText = "FIRST (Home)"
            .Buttons(9).ToolTipText = "BACK (Up)"
            .Buttons(11).ToolTipText = "NEXT (Down)"
            .Buttons(13).ToolTipText = "LAST (End)"
            .Buttons(15).ToolTipText = "FIND (F6)"
            .Buttons(17).ToolTipText = "PRINT (F9)"
            .Buttons(19).ToolTipText = "CLOSE (Esc)"
        Case 2  '=== ADD/EDIT ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(1).Enabled = False
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = False
            .Buttons(7).Image = 11
            .Buttons(7).Caption = "Save"
            .Buttons(9).Image = 12
            .Buttons(9).Caption = "Undo"
            .Buttons(7).Enabled = True
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
            .Buttons(19).Enabled = False
            .Buttons(1).ToolTipText = ""
            .Buttons(3).ToolTipText = ""
            .Buttons(5).ToolTipText = ""
            .Buttons(7).ToolTipText = "SAVE (F5)"
            .Buttons(9).ToolTipText = "UNDO (Esc)"
            .Buttons(11).ToolTipText = ""
            .Buttons(13).ToolTipText = ""
            .Buttons(15).ToolTipText = ""
            .Buttons(17).ToolTipText = ""
            .Buttons(19).ToolTipText = ""
        Case 3  '=== FIND ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(1).Enabled = False
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = False
            .Buttons(7).Image = 4
            .Buttons(7).Caption = "First"
            .Buttons(9).Image = 12
            .Buttons(9).Caption = "Undo"
            .Buttons(7).Enabled = False
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
            .Buttons(19).Enabled = False
            .Buttons(1).ToolTipText = ""
            .Buttons(3).ToolTipText = ""
            .Buttons(5).ToolTipText = ""
            .Buttons(7).ToolTipText = ""
            .Buttons(9).ToolTipText = "UNDO (Esc)"
            .Buttons(11).ToolTipText = ""
            .Buttons(13).ToolTipText = ""
            .Buttons(15).ToolTipText = ""
            .Buttons(17).ToolTipText = ""
            .Buttons(19).ToolTipText = ""
    End Select
End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert
    Case vbKeyF2
    Case vbKeyDelete
    Case vbKeyF5
    Case vbKeyF6
    Case vbKeyEscape
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "ChartOfAccount", "ChartOfAccount", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "ChartOfAccount", "ChartOfAccount", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "ChartOfAccount", "ChartOfAccount", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "ChartOfAccount", "ChartOfAccount", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
'Me.Caption = "Chart Of Account"
cmbAccountGroup.Clear
s = "SELECT tbl_GL_Department.* " & _
    " FROM tbl_GL_Department " & _
    " ORDER BY DeptName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbAccountGroup.AddItem UCase(rs!DeptName)
    cmbAccountGroup.ItemData(cmbAccountGroup.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "ChartOfAccount", "ChartOfAccount", ""), "is_LOAD"
If Trim(txtAccountCode.Text) = "" Then BROWSER GetSetting(App.EXEName, "ChartOfAccount", "ChartOfAccount", ""), "is_HOME"

tmp = SetWindowLong(txtAccountCode.hwnd, GWL_STYLE, GetWindowLong(txtAccountCode.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtAccountName.hwnd, GWL_STYLE, GetWindowLong(txtAccountName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "ChartOfAccount", "ChartOfAccount", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "ChartOfAccount", "ChartOfAccount", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "ChartOfAccount", "ChartOfAccount", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "ChartOfAccount", "ChartOfAccount", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Close":   PRESS_ESCAPE
End Select
End Sub
