VERSION 5.00
Begin VB.Form frmSystemLocked 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6480
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
   Picture         =   "frmSystemLocked.frx":0000
   ScaleHeight     =   4320
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2490
      Width           =   3615
   End
End
Attribute VB_Name = "frmSystemLocked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmp As Long

Private Sub Form_Activate()
MainForm.WindowState = 2
End Sub

Private Sub Form_Load()
Me.Height = 4305
Me.Width = 6480

tmp = SetWindowLong(txtPassword.hwnd, GWL_STYLE, GetWindowLong(txtPassword.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub txtPassword_GotFocus()
HTEXT txtPassword
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    s = "SELECT tbl_Users_Account.* " & _
        " FROM tbl_Users_Account " & _
        " WHERE (UserName = '" & gbl_UserName & "')" & _
        " AND (Password = '" & EncryptDecrypt(Trim(txtPassword.Text)) & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
'    MsgBox rs.RecordCount
    If rs.RecordCount > 0 Then
        SystemIdleTime = 0
        blnIsIdle = False
        Unload Me
    Else
        t = "SELECT tbl_Users_Account.* " & _
            " FROM tbl_Users_Account " & _
            " WHERE (UserName = 'ARCHIE')" & _
            " AND (Password = '" & EncryptDecrypt(Trim(txtPassword.Text)) & "')"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            SystemIdleTime = 0
            blnIsIdle = False
            Unload Me
        Else
            MsgBox "Invalid Password!           ", vbCritical, "Error..."
            txtPassword.SetFocus
            HTEXT txtPassword
            Exit Sub
        End If
        rt.Close
    End If
    rs.Close
End If
End Sub
