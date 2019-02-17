VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form aLogIn 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log In"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11655
   ControlBox      =   0   'False
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
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2640
      Top             =   4320
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5520
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogIn.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogIn.frx":1992
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picUsers 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   3720
      ScaleHeight     =   3255
      ScaleWidth      =   4095
      TabIndex        =   0
      Top             =   360
      Width           =   4095
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   405
         Left            =   2160
         TabIndex        =   1
         Top             =   2640
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   714
         Caption         =   "Next >>"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         mPointer        =   99
         mIcon           =   "frmLogIn.frx":3324
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   405
         Left            =   600
         TabIndex        =   2
         Top             =   2640
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   714
         Caption         =   "Cancel"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         mPointer        =   99
         mIcon           =   "frmLogIn.frx":363E
      End
      Begin MSComctlLib.ListView lstUser 
         Height          =   2295
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   4048
         Arrange         =   2
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         OLEDragMode     =   1
         PictureAlignment=   4
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   -2147483643
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmLogIn.frx":3958
         OLEDragMode     =   1
         NumItems        =   0
      End
   End
   Begin VB.PictureBox picPassword 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   3720
      ScaleHeight     =   3255
      ScaleWidth      =   4095
      TabIndex        =   3
      Top             =   360
      Width           =   4095
      Begin VB.TextBox txtPW 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         IMEMode         =   3  'DISABLE
         Left            =   615
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1695
         Width           =   2985
      End
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   615
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   8
         Top             =   915
         Width           =   2865
      End
      Begin VB.CheckBox chkRemeberUserName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remember My Name"
         ForeColor       =   &H00808080&
         Height          =   270
         Left            =   615
         TabIndex        =   7
         Top             =   2040
         Value           =   1  'Checked
         Width           =   2190
      End
      Begin lvButton.lvButtons_H cmdLogIn 
         Default         =   -1  'True
         Height          =   405
         Left            =   2160
         TabIndex        =   4
         Top             =   2640
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   714
         Caption         =   "Log In"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         mPointer        =   99
         mIcon           =   "frmLogIn.frx":4632
      End
      Begin lvButton.lvButtons_H cmdBack2 
         Height          =   405
         Left            =   600
         TabIndex        =   5
         Top             =   2640
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   714
         Caption         =   "<< Back"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         mPointer        =   99
         mIcon           =   "frmLogIn.frx":494C
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   615
         TabIndex        =   11
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   600
         TabIndex        =   10
         Top             =   720
         Width           =   780
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00BCB39A&
      Caption         =   "Label1"
      Height          =   255
      Left            =   2520
      TabIndex        =   17
      Top             =   5280
      Width           =   3015
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[1]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   4140
      TabIndex        =   16
      Top             =   -40
      Width           =   360
   End
   Begin VB.Label lblStep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select User"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5040
      TabIndex        =   15
      Top             =   0
      Width           =   1170
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Step"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3720
      TabIndex        =   14
      Top             =   50
      Width           =   390
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[2]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   345
      Left            =   4560
      TabIndex        =   13
      Top             =   -40
      Width           =   360
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00BCB39A&
      Caption         =   "Invalid Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   3840
      TabIndex        =   12
      Top             =   3690
      Width           =   3975
   End
   Begin VB.Image Image2 
      Height          =   4020
      Left            =   0
      Picture         =   "frmLogIn.frx":4C66
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "aLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Arr1, Arr2

Private Sub cmdBack2_Click()
lbl1.ForeColor = &H0&
lbl2.ForeColor = &HC0C0C0
lblStep.Caption = "Select User"
picUsers.ZOrder 0
picUsers.Visible = True
'lstUser.TabIndex = 0
lstUser.SetFocus
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdLogIn_Click()
If Trim(txtPW.Text) = "" Then MsgBox "Please Supply Password!                 ", vbCritical, "Error...": txtPW.SetFocus: HTEXT txtPW: Exit Sub
s = "SELECT TOP 1 tbl_Users_Account.*" & _
    " From tbl_Users_Account " & _
    " WHERE (UserName = '" & FORMATSQL(Trim(txtUserName.Text)) & "') " & _
    " AND (Password = '" & EncryptDecrypt(FORMATSQL(Trim(txtPW.Text))) & "') " & _
    " ORDER BY UserName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount = 0 Then Timer1.Enabled = True: Exit Sub
If chkRemeberUserName.Value = 1 Then
    SaveSetting App.EXEName, "LastUser", "LUser", Trim(txtUserName.Text)
Else
    SaveSetting App.EXEName, "LastUser", "LUser", ""
End If
Timer1.Enabled = False

LOAD_HIDE_MENU IIf(rs!Admin = 1, True, False)

gbl_UserName = rs!UserName
gbl_Password = Trim(txtPW.Text)
gbl_CompleteName = rs!CompleteName

ConnOmega.Execute "UPDATE tbl_Users_Account SET Online = 1 WHERE (UserName = '" & FORMATSQL(CStr(gbl_UserName)) & "')"

SystemSetting = rs!UserSettings
Arr1 = Split(SystemSetting, "}", -1, 1)
Arr2 = Split(Arr1(0), "/", -1, 1)
gbl_LockWhenIdle = CLng(Arr2(0))
gbl_Idle_Time = CDbl(Arr2(1))
Arr2 = Split(Arr1(1), "/", -1, 1)
gbl_Slides_Background = CLng(Arr2(0))
gbl_Slides_Time = CDbl(Arr2(1))
gbl_Quotes_Time = CDbl(Arr1(2))

MainForm.mnuMainLogInOut.Caption = "Log &Out"
MainForm.mnuLockedSystem.Enabled = True
MainForm.Statusbar1.Panels(3).Text = rs!UserName
rs.Close
Unload Me

If AccessRights(gbl_MODULE, gbl_MODULE_Action) = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If

If LogInWithOutLoading = 0 Then Exit Sub

'If LogInWithOutLoading = 1 Then LogInWithOutLoading = 0: Exit Sub

On Error Resume Next
If gbl_FORM_Modal = 1 Then
    gbl_FORM.Show 1
Else
    gbl_FORM.Show
End If

LogInWithOutLoading = 0

End Sub

Private Sub cmdNext_Click()
lbl1.ForeColor = &HC0C0C0
lbl2.ForeColor = &H0&
lblStep.Caption = "Enter Password.."
picPassword.ZOrder 0
txtUserName.Text = lstUser.SelectedItem.Text
picPassword.Visible = True
txtPW.Text = ""
txtPW.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()

KeyPreview = True

Me.Height = 4395
Me.Width = 8025

Label2.Caption = ""

s = "SELECT tbl_Users_Account.* " & _
    " FROM tbl_Users_Account " & _
    " ORDER BY UserName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
lstUser.ListItems.Clear
Set lstUser.SmallIcons = ImageList1
Set lstUser.Icons = ImageList1
While Not rs.EOF
    If rs!Gender = 2 Then
        lstUser.ListItems.Add , rs!UserName, rs!UserName, 2
    Else
        lstUser.ListItems.Add , rs!UserName, rs!UserName, 1
    End If
    rs.MoveNext
Wend
rs.Close

gbl_Last_User = GetSetting(App.EXEName, "LastUser", "LUser", "")

If Trim(gbl_Last_User) = "" Then
    lbl1.ForeColor = &H0&
    lbl2.ForeColor = &HC0C0C0
    lblStep.Caption = "Select User"
    picUsers.ZOrder 0
    picUsers.Visible = True
    lstUser.TabIndex = 0
Else
    s = "SELECT tbl_Users_Account.* " & _
        " FROM tbl_Users_Account " & _
        " WHERE (UserName = '" & gbl_Last_User & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount = 0 Then
        lbl1.ForeColor = &H0&
        lbl2.ForeColor = &HC0C0C0
        lblStep.Caption = "Select User"
        picUsers.ZOrder 0
        picUsers.Visible = True
        lstUser.TabIndex = 0
    Else
        lbl1.ForeColor = &HC0C0C0
        lbl2.ForeColor = &H0&
        lblStep.Caption = "Enter Password.."
        picPassword.ZOrder 0
        picPassword.Visible = True
        txtUserName.Text = gbl_Last_User
        txtPW.TabIndex = 0
    End If
    rs.Close
End If
Dim tmp As Long
tmp = SetWindowLong(txtPW.hwnd, GWL_STYLE, GetWindowLong(txtPW.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub lstUser_DblClick()
If lstUser.ListItems.Count = 0 Then Exit Sub
cmdNext_Click
End Sub

Private Sub Timer1_Timer()
DoEvents
If Label2.Caption = "" Then Label2.Caption = "Invalid Password" Else Label2.Caption = ""
'If Label2.Visible = True Then Label2.Visible = False Else Label2.Visible = True
End Sub
