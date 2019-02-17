VERSION 5.00
Begin VB.Form bChangePassword 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "bChangePassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
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
      Left            =   2520
      Picture         =   "bChangePassword.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1560
   End
   Begin VB.CommandButton cmdOK 
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
      Left            =   840
      Picture         =   "bChangePassword.frx":0A66
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1560
   End
   Begin VB.TextBox txtConfirmPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox txtNewPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox txtOldPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CONFIRM PASSWORD"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NEW PASSWORD"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "OLD PASSWORD"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "bChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If Trim(txtOldPassword.Text) = "" Then MsgBox "Please Supply Old Password!                      ", vbCritical, "Error...": txtOldPassword.SetFocus: Exit Sub
If Trim(txtNewPassword.Text) = "" Then MsgBox "Please Supply New Password!                        ", vbCritical, "Error...": txtNewPassword.SetFocus: Exit Sub
If Trim(txtConfirmPassword.Text) = "" Then MsgBox "Please Supply Confirm Password!                    ", vbCritical, "Error...": txtConfirmPassword.SetFocus: Exit Sub
If Trim(gbl_Password) <> Trim(txtOldPassword.Text) Then MsgBox "Invalid Password!                   ", vbCritical, "Error...": txtOldPassword.SetFocus: Exit Sub
If Trim(txtNewPassword.Text) <> Trim(txtConfirmPassword.Text) Then MsgBox "Confirm Password did not Match from your New Password!                    ", vbCritical, "Error...": txtConfirmPassword.SetFocus: Exit Sub

ConnOmega.Execute "UPDATE tbl_Users_Account SET Password = '" & EncryptDecrypt(FORMATSQL(Trim(txtNewPassword.Text))) & "' WHERE (UserName = '" & gbl_UserName & "')"

gbl_Password = Trim(txtNewPassword.Text)

MsgBox "Password Successfully Changed!                      ", vbInformation, "PW Change"

Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Caption = "Change Password"
Me.Top = (MainForm.Height - Me.Height) / 4
Me.Left = (MainForm.Width - Me.Width) / 5

Dim tmp As Long
tmp = SetWindowLong(txtOldPassword.hwnd, GWL_STYLE, GetWindowLong(txtOldPassword.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtNewPassword.hwnd, GWL_STYLE, GetWindowLong(txtNewPassword.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtConfirmPassword.hwnd, GWL_STYLE, GetWindowLong(txtConfirmPassword.hwnd, GWL_STYLE) Or ES_UPPERCASE)

End Sub

Private Sub txtConfirmPassword_GotFocus()
HTEXT txtConfirmPassword
End Sub

Private Sub txtConfirmPassword_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then txtNewPassword.SetFocus
If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub txtNewPassword_GotFocus()
HTEXT txtNewPassword
End Sub

Private Sub txtNewPassword_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtConfirmPassword.SetFocus
If KeyCode = vbKeyDown Then txtConfirmPassword.SetFocus
If KeyCode = vbKeyUp Then txtOldPassword.SetFocus
End Sub

Private Sub txtOldPassword_GotFocus()
HTEXT txtOldPassword
End Sub

Private Sub txtOldPassword_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtNewPassword.SetFocus
If KeyCode = vbKeyDown Then txtNewPassword.SetFocus
End Sub
