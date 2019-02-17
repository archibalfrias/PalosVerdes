VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmCompanyModal 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCompanyModal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTelNo 
      Height          =   315
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1680
      Width           =   4740
   End
   Begin VB.TextBox txtAddress2 
      Height          =   315
      Left            =   1200
      MaxLength       =   100
      TabIndex        =   2
      Top             =   1320
      Width           =   4740
   End
   Begin VB.TextBox txtAddress1 
      Height          =   315
      Left            =   1200
      MaxLength       =   100
      TabIndex        =   1
      Top             =   960
      Width           =   4740
   End
   Begin VB.TextBox txtCompanyName 
      Height          =   315
      Left            =   1200
      MaxLength       =   100
      TabIndex        =   0
      Top             =   600
      Width           =   4740
   End
   Begin VB.TextBox txtFaxNo 
      Height          =   315
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2040
      Width           =   4740
   End
   Begin VB.TextBox txtSSSNo 
      Height          =   315
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2400
      Width           =   4740
   End
   Begin VB.TextBox txtPHICNo 
      Height          =   315
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   6
      Top             =   2760
      Width           =   4740
   End
   Begin VB.TextBox txtTIN 
      Height          =   315
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   7
      Top             =   3120
      Width           =   4740
   End
   Begin VB.TextBox txtPK 
      Height          =   315
      Left            =   5520
      MaxLength       =   100
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   180
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   405
      Left            =   4470
      TabIndex        =   8
      Top             =   3600
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      Caption         =   "&Save"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   15396057
      Focus           =   0   'False
      cGradient       =   15396057
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   405
      Left            =   3000
      TabIndex        =   9
      Top             =   3600
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      Caption         =   "&Cancel"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   15396057
      Focus           =   0   'False
      cGradient       =   15396057
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TEL #"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   1725
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   1005
      Width           =   735
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   645
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   345
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   3150
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "FAX #"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   2085
      Width           =   1095
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SSS #"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   2445
      Width           =   1095
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PHIC #"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   2805
      Width           =   1095
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "T I N"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   3165
      Width           =   1095
   End
End
Attribute VB_Name = "frmCompanyModal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmp As Long

Private Sub cmdCancel_Click()
'Unload Me
End
End Sub

Private Sub cmdSave_Click()
If Trim(txtCompanyName.Text) = "" Then MsgBox "Please Supply Company Name!                ", vbCritical, "Error...": txtCompanyName.SetFocus: Exit Sub
If Trim(txtTelNo.Text) = "" Then MsgBox "Please Supply Telephone Number!                ", vbCritical, "Error...": txtTelNo.SetFocus: Exit Sub
If Trim(txtSSSNo.Text) = "" Then MsgBox "Please Supply SSS Number!                ", vbCritical, "Error...": txtSSSNo.SetFocus: Exit Sub
If Trim(txtPHICNo.Text) = "" Then MsgBox "Please Supply PhilHealth Number!                ", vbCritical, "Error...": txtPHICNo.SetFocus: Exit Sub
If Trim(txtTIN.Text) = "" Then MsgBox "Please Supply Tax Identification Number!                ", vbCritical, "Error...": txtTIN.SetFocus: Exit Sub
If RETURNTEXTVALUE(txtPK) = 0 Then
    ConnOmega.Execute "INSERT INTO tbl_Company " & _
                      " (PK, CompanyName, Address1, " & _
                      " Address2, TelNo, FaxNo, SSSNo, " & _
                      " PHICNo, TIN) " & _
                      " VALUES (1, '" & FORMATSQL(Trim(txtCompanyName.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtAddress1.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtAddress2.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtTelNo.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtFaxNo.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtSSSNo.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtPHICNo.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtTIN.Text)) & "')"
Else
    ConnOmega.Execute "UPDATE tbl_Company " & _
                      " SET CompanyName = '" & FORMATSQL(Trim(txtCompanyName.Text)) & "', " & _
                      " Address1 = '" & FORMATSQL(Trim(txtAddress1.Text)) & "', " & _
                      " Address2 = '" & FORMATSQL(Trim(txtAddress2.Text)) & "', " & _
                      " TelNo = '" & FORMATSQL(Trim(txtTelNo.Text)) & "', " & _
                      " FaxNo = '" & FORMATSQL(Trim(txtFaxNo.Text)) & "', " & _
                      " SSSNo = '" & FORMATSQL(Trim(txtSSSNo.Text)) & "', " & _
                      " PHICNo = '" & FORMATSQL(Trim(txtPHICNo.Text)) & "', " & _
                      " TIN = '" & FORMATSQL(Trim(txtTIN.Text)) & "' " & _
                      " WHERE (PK = 1)"
End If

gbl_CompanyName = FORMATSQL(Trim(txtCompanyName.Text))
gbl_CompanyAddress1 = FORMATSQL(Trim(txtAddress1.Text))
gbl_CompanyAddress2 = FORMATSQL(Trim(txtAddress2.Text))
gbl_CompanyTelNo = FORMATSQL(Trim(txtTelNo.Text))
gbl_CompanyFaxNo = FORMATSQL(Trim(txtFaxNo.Text))
gbl_CompanySSSNo = FORMATSQL(Trim(txtSSSNo.Text))
gbl_CompanyPHICNo = FORMATSQL(Trim(txtPHICNo.Text))
gbl_CompanyTIN = FORMATSQL(Trim(txtTIN.Text))

Unload Me
MainForm.Show
frmBackground.Quotes
frmBackground.picQuotes.Visible = True
frmBackground.picFreeMem.Visible = True
frmBackground.picDayTime.Visible = True
MainForm.Timer_CheckIdle.Enabled = True

End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Caption = "Company Information"

s = "SELECT PK, CompanyName, Address1, " & _
    " Address2, TelNo, FaxNo, SSSNo, PHICNo, TIN " & _
    " From tbl_Company " & _
    " WHERE (PK = 1)"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtPK.Text = rs!PK
    txtCompanyName.Text = rs!CompanyName
    txtAddress1.Text = rs!Address1
    txtAddress2.Text = rs!Address2
    txtTelNo.Text = rs!TelNo
    txtFaxNo.Text = rs!FaxNo
    txtSSSNo.Text = rs!SSSNo
    txtPHICNo.Text = rs!PHICNo
    txtTIN.Text = rs!TIN
End If
rs.Close


tmp = SetWindowLong(txtCompanyName.hwnd, GWL_STYLE, GetWindowLong(txtCompanyName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtAddress1.hwnd, GWL_STYLE, GetWindowLong(txtAddress1.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtAddress2.hwnd, GWL_STYLE, GetWindowLong(txtAddress2.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtTelNo.hwnd, GWL_STYLE, GetWindowLong(txtTelNo.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtFaxNo.hwnd, GWL_STYLE, GetWindowLong(txtFaxNo.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSSSNo.hwnd, GWL_STYLE, GetWindowLong(txtSSSNo.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPHICNo.hwnd, GWL_STYLE, GetWindowLong(txtPHICNo.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtTIN.hwnd, GWL_STYLE, GetWindowLong(txtTIN.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub txtAddress1_GotFocus()
HTEXT txtAddress1
End Sub

Private Sub txtAddress1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtAddress2.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtCompanyName.SetFocus
End If
End Sub

Private Sub txtAddress2_GotFocus()
HTEXT txtAddress2
End Sub

Private Sub txtAddress2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtTelNo.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtAddress1.SetFocus
End If
End Sub

Private Sub txtCompanyName_GotFocus()
HTEXT txtCompanyName
End Sub

Private Sub txtCompanyName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtAddress1.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtTIN.SetFocus
End If
End Sub

Private Sub txtFaxNo_GotFocus()
HTEXT txtFaxNo
End Sub

Private Sub txtFaxNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtSSSNo.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtTelNo.SetFocus
End If
End Sub

Private Sub txtPHICNo_GotFocus()
HTEXT txtPHICNo
End Sub

Private Sub txtPHICNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtTIN.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtSSSNo.SetFocus
End If
End Sub

Private Sub txtSSSNo_GotFocus()
HTEXT txtSSSNo
End Sub

Private Sub txtSSSNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtPHICNo.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtFaxNo.SetFocus
End If
End Sub

Private Sub txtTelNo_GotFocus()
HTEXT txtTelNo
End Sub

Private Sub txtTelNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtFaxNo.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtAddress2.SetFocus
End If
End Sub

Private Sub txtTIN_GotFocus()
HTEXT txtTIN
End Sub

Private Sub txtTIN_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtCompanyName.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtPHICNo.SetFocus
End If
End Sub


