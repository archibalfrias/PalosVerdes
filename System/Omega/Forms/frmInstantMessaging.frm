VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmInstantMessaging 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8520
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
   MDIChild        =   -1  'True
   ScaleHeight     =   336
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   568
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer_Msg 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5160
      Top             =   120
   End
   Begin VB.Timer Timer_Online_Offline 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2280
      Top             =   2640
   End
   Begin VB.TextBox txtMessage 
      Height          =   615
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4200
      Width           =   4335
   End
   Begin MSComctlLib.ListView lstUsersOn 
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3201
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   4498
      EndProperty
   End
   Begin MSComctlLib.ListView lstUsersOff 
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3201
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   4498
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdSend 
      Height          =   645
      Left            =   7560
      TabIndex        =   1
      Top             =   4200
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   1138
      Caption         =   "Send"
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
      cBhover         =   12648384
      Focus           =   0   'False
      LockHover       =   1
      cGradient       =   13023396
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
      mPointer        =   99
      mIcon           =   "frmInstantMessaging.frx":0000
   End
   Begin RPVGCC.SOLO_RTBSyntax Msg 
      Height          =   3495
      Left            =   3120
      TabIndex        =   7
      Top             =   600
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6165
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmInstantMessaging.frx":031A
      Syntax_LogicalOperators=   "}"
      ColorOf_Comments=   16711680
      Syntax_CommentChar=   ";"
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Offline Users"
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Online Users"
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.Image imgTitleClose 
      Height          =   195
      Left            =   3360
      MouseIcon       =   "frmInstantMessaging.frx":0336
      MousePointer    =   99  'Custom
      Picture         =   "frmInstantMessaging.frx":0640
      Top             =   120
      Width           =   195
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   480
      TabIndex        =   2
      Top             =   0
      Width           =   435
   End
   Begin VB.Image imgTitleHelp 
      Height          =   195
      Left            =   1200
      Picture         =   "frmInstantMessaging.frx":088A
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgTitleMinimize 
      Height          =   195
      Left            =   1440
      Picture         =   "frmInstantMessaging.frx":0AD4
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgWindowLeft 
      Height          =   450
      Left            =   0
      Picture         =   "frmInstantMessaging.frx":0D1E
      Stretch         =   -1  'True
      Top             =   840
      Width           =   285
   End
   Begin VB.Image imgWindowBottom 
      Height          =   450
      Left            =   1800
      Picture         =   "frmInstantMessaging.frx":1468
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   1800
      Picture         =   "frmInstantMessaging.frx":1BB2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgWindowBottomLeft 
      Height          =   450
      Left            =   0
      Picture         =   "frmInstantMessaging.frx":22FC
      Top             =   4560
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   8160
      Picture         =   "frmInstantMessaging.frx":2A46
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   0
      Picture         =   "frmInstantMessaging.frx":3190
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgWindowBottomRight 
      Height          =   450
      Left            =   8160
      Picture         =   "frmInstantMessaging.frx":38DA
      Top             =   4560
      Width           =   285
   End
   Begin VB.Image imgWindowRight 
      Height          =   450
      Left            =   8160
      Picture         =   "frmInstantMessaging.frx":4024
      Stretch         =   -1  'True
      Top             =   840
      Width           =   285
   End
End
Attribute VB_Name = "frmInstantMessaging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Online_Offline As Long

Dim strUser, strUsers, strFormsArr, i
Dim strNames, strForms As String
Dim Loaded As Boolean
Dim Form As Form

Dim x, OnL


Private Sub cmdSend_Click()
If lstUsersOn.ListItems.Count = 0 Then Exit Sub
If FORMATENTER(Trim(txtMessage.Text)) <> "" Then
    If Len(Msg.Text) <> 0 Then
        Msg.Text = Msg.Text & vbCrLf
    End If

    Msg.Text = Msg.Text & gbl_UserName & "; " & FORMATENTER(Trim(txtMessage.Text))
    
    Msg.LineScroll (Msg.GetLineCount)
    
    For i = 1 To lstUsersOn.ListItems.Count
        If Trim(lstUsersOn.ListItems.Item(i).SubItems(1)) <> Trim(gbl_UserName) Then
            ConnOmega.Execute "INSERT INTO tbl_InstantMessaging" & _
                              " (Date_Time, Message, From_User, To_User, MsgType) " & _
                              " VALUES('" & Now & "', " & _
                              " '" & FORMATENTER(FORMATSQL(Trim(txtMessage.Text))) & "', " & _
                              " '" & gbl_UserName & "', " & _
                              " '" & lstUsersOn.ListItems.Item(i).SubItems(1) & "', 1)"
        End If
    Next i
    
    txtMessage.Refresh
    txtMessage.Text = ""
    txtMessage.Text = FORMATENTER(txtMessage.Text)
    txtMessage.SelStart = 0
    txtMessage.SetFocus
    
Else
    txtMessage.Refresh
    txtMessage.Text = ""
    txtMessage.Text = FORMATENTER(txtMessage.Text)
    txtMessage.SelStart = 0
    txtMessage.SetFocus
End If
End Sub

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
End Sub

Private Sub Form_Load()
KeyPreview = True
lblTitle.Caption = "Instant Messaging"
MakeWindowInstantMessaging Me
Me.Top = (MainForm.Height - Me.Height) / 4
Me.Left = (MainForm.Width - Me.Width) / 5

lstUsersOn.ListItems.Clear
lstUsersOff.ListItems.Clear
s = "SELECT UserName" & _
    " From tbl_Users_Account " & _
    " WHERE (UserName <> '" & gbl_UserName & "') " & _
    " AND (Online = 1) " & _
    " ORDER BY UserName "
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    Set x = lstUsersOn.ListItems.Add()
    x.Text = " "
    x.SubItems(1) = UCase(rs!UserName)
    rs.MoveNext
Wend
rs.Close

s = "SELECT UserName" & _
    " From tbl_Users_Account " & _
    " WHERE (UserName <> '" & gbl_UserName & "') " & _
    " AND (Online = 0) " & _
    " ORDER BY UserName "
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    Set x = lstUsersOff.ListItems.Add()
    x.Text = " "
    x.SubItems(1) = UCase(rs!UserName)
    rs.MoveNext
Wend
rs.Close

Timer_Online_Offline.Enabled = True
Timer_Msg.Enabled = True

End Sub

Private Sub imgTitleClose_Click()
Unload Me
End Sub

Private Sub imgTitleMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DoDrag Me
End Sub

Private Sub imgTitleRight_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DoDrag Me
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DoDrag Me
End Sub

Private Sub lstUsersOff_DblClick()
If lstUsersOff.ListItems.Count = 0 Then Exit Sub
strUser = lstUsersOff.ListItems.Item(lstUsersOff.SelectedItem.Index).SubItems(1)
strNames = ""
For Each Form In Forms
    strNames = strNames & ";" & Form.Name
Next Form
strForms = Replace(Replace(Replace(strNames, ";frmBackground", ""), ";MainFormPopupF", ""), ";Mainform", "")
strFormsArr = Split(strForms, ";", -1, 1)
strUsers = ""
Loaded = False
For Each Form In Forms
    strUsers = strUsers & ";" & Form.Caption
    If Trim(strUser) = Form.Caption Then
        Loaded = True
        Exit For
    End If
Next Form
If Loaded = True Then
    Form.ZOrder 0
Else
    Dim objForm As New frmInstantMessagingPM
    objForm.Caption = strUser
    objForm.lblTitle.Caption = strUser
    objForm.Show
End If
End Sub

Private Sub lstUsersOn_DblClick()
If lstUsersOn.ListItems.Count = 0 Then Exit Sub
strUser = lstUsersOn.ListItems.Item(lstUsersOn.SelectedItem.Index).SubItems(1)
strNames = ""
For Each Form In Forms
    strNames = strNames & ";" & Form.Name
Next Form
strForms = Replace(Replace(Replace(strNames, ";frmBackground", ""), ";MainFormPopupF", ""), ";Mainform", "")
strFormsArr = Split(strForms, ";", -1, 1)
strUsers = ""
Loaded = False
For Each Form In Forms
    strUsers = strUsers & ";" & Form.Caption
    If Trim(strUser) = Form.Caption Then
        Loaded = True
        Exit For
    End If
Next Form
If Loaded = True Then
    Form.ZOrder 0
Else
    Dim objForm As New frmInstantMessagingPM
    objForm.Caption = strUser
    objForm.lblTitle.Caption = strUser
    objForm.Show
End If
End Sub

Private Sub Timer_Msg_Timer()
Timer_Msg.Enabled = False
s = "SELECT PK, Date_Time, Message, From_User" & _
    " From tbl_InstantMessaging " & _
    " WHERE (Convert(datetime, Convert(char(6), Date_Time,12), 102) = '" & FormatDateTime(Date, vbShortDate) & "') " & _
    " AND (MsgType = 1) " & _
    " AND (To_User = '" & gbl_UserName & "')" & _
    " AND (Opened = 0)"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    DoEvents
    
    If Len(Msg.Text) <> 0 Then
        Msg.Text = Msg.Text & vbCrLf
    End If
    
    Msg.Text = Msg.Text & rs!From_User & ": " & rs!Message
    
    Msg.LineScroll (Msg.GetLineCount - 16)
    
    ConnOmega.Execute "UPDATE tbl_InstantMessaging" & _
                      " SET Opened = 1 " & _
                      " WHERE (PK = " & rs!PK & ")"
    rs.MoveNext
Wend
rs.Close
Timer_Msg.Enabled = True
End Sub

Private Sub Timer_Online_Offline_Timer()
Timer_Online_Offline.Enabled = False
s = "SELECT UserName" & _
    " From tbl_Users_Account " & _
    " WHERE (UserName <> '" & gbl_UserName & "') " & _
    " AND (Online = 1) " & _
    " ORDER BY UserName "
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    
    For i = 1 To lstUsersOff.ListItems.Count
        If UCase(rs!UserName) = lstUsersOff.ListItems.Item(i).SubItems(1) Then
            lstUsersOff.ListItems.Remove i
            Exit For
        End If
    Next i
    
    OnL = 0
    For i = 1 To lstUsersOn.ListItems.Count
        If UCase(rs!UserName) = lstUsersOn.ListItems.Item(i).SubItems(1) Then
            OnL = 1
            Exit For
        End If
    Next i
    
    If CDbl(OnL) = 0 Then
        Set x = lstUsersOn.ListItems.Add()
        x.Text = ""
        x.SubItems(1) = UCase(rs!UserName)
    End If
    
    rs.MoveNext
Wend
rs.Close

s = "SELECT UserName" & _
    " From tbl_Users_Account " & _
    " WHERE (UserName <> '" & gbl_UserName & "') " & _
    " AND (Online = 0) " & _
    " ORDER BY UserName "
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    
    For i = 1 To lstUsersOn.ListItems.Count
        If UCase(rs!UserName) = lstUsersOn.ListItems.Item(i).SubItems(1) Then
            lstUsersOn.ListItems.Remove i
            Exit For
        End If
    Next i
    
    OnL = 0
    For i = 1 To lstUsersOff.ListItems.Count
        If UCase(rs!UserName) = lstUsersOff.ListItems.Item(i).SubItems(1) Then
            OnL = 1
            Exit For
        End If
    Next i
    
    If CDbl(OnL) = 0 Then
        Set x = lstUsersOff.ListItems.Add()
        x.Text = ""
        x.SubItems(1) = UCase(rs!UserName)
    End If
    
    rs.MoveNext
Wend
rs.Close
Timer_Online_Offline.Enabled = True
End Sub
