VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmInstantMessagingPM 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5715
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
   ScaleHeight     =   375
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   381
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer_Msg 
      Interval        =   100
      Left            =   3600
      Top             =   120
   End
   Begin VB.TextBox txtMessage 
      Height          =   735
      Left            =   120
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4680
      Width           =   4695
   End
   Begin RPVGCC.SOLO_RTBSyntax Msg 
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   6800
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmInstantMessagingPM.frx":0000
      Syntax_LogicalOperators=   "}"
      ColorOf_Comments=   16711680
      Syntax_CommentChar=   ";"
   End
   Begin lvButton.lvButtons_H cmdSend 
      Height          =   735
      Left            =   4830
      TabIndex        =   1
      Top             =   4680
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   1296
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
      mIcon           =   "frmInstantMessagingPM.frx":001C
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
   Begin VB.Image imgTitleClose 
      Height          =   195
      Left            =   4920
      MouseIcon       =   "frmInstantMessagingPM.frx":0336
      MousePointer    =   99  'Custom
      Picture         =   "frmInstantMessagingPM.frx":0640
      Top             =   120
      Width           =   195
   End
   Begin VB.Image imgTitleMinimize 
      Height          =   195
      Left            =   1440
      Picture         =   "frmInstantMessagingPM.frx":088A
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgTitleHelp 
      Height          =   195
      Left            =   1200
      Picture         =   "frmInstantMessagingPM.frx":0AD4
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgWindowRight 
      Height          =   450
      Left            =   5400
      Picture         =   "frmInstantMessagingPM.frx":0D1E
      Stretch         =   -1  'True
      Top             =   840
      Width           =   285
   End
   Begin VB.Image imgWindowBottomRight 
      Height          =   450
      Left            =   5400
      Picture         =   "frmInstantMessagingPM.frx":1468
      Top             =   5160
      Width           =   285
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   0
      Picture         =   "frmInstantMessagingPM.frx":1BB2
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   5400
      Picture         =   "frmInstantMessagingPM.frx":22FC
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgWindowBottomLeft 
      Height          =   450
      Left            =   0
      Picture         =   "frmInstantMessagingPM.frx":2A46
      Top             =   5160
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   2760
      Picture         =   "frmInstantMessagingPM.frx":3190
      Stretch         =   -1  'True
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgWindowBottom 
      Height          =   450
      Left            =   3480
      Picture         =   "frmInstantMessagingPM.frx":38DA
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   285
   End
   Begin VB.Image imgWindowLeft 
      Height          =   450
      Left            =   0
      Picture         =   "frmInstantMessagingPM.frx":4024
      Stretch         =   -1  'True
      Top             =   840
      Width           =   285
   End
End
Attribute VB_Name = "frmInstantMessagingPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i


Private Sub cmdSend_Click()
If FORMATENTER(Trim(txtMessage.Text)) <> "" Then
    If Len(Msg.Text) <> 0 Then
        Msg.Text = Msg.Text & vbCrLf '& vbCrLf
    End If

    Msg.Text = Msg.Text & gbl_UserName & "; " & FORMATENTER(Trim(txtMessage.Text))
    
    ConnOmega.Execute "INSERT INTO tbl_InstantMessaging" & _
                      " (Date_Time, Message, From_User, To_User) " & _
                      " VALUES('" & Now & "', " & _
                      " '" & FORMATENTER(FORMATSQL(Trim(txtMessage.Text))) & "', " & _
                      " '" & gbl_UserName & "', '" & Me.Caption & "')"
    
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
KeyPreview = True
MakeWindowInstantMessaging Me
Me.Top = (MainForm.Height - Me.Height) / 4
Me.Left = (MainForm.Width - Me.Width) / 5
iMsgLoaded = iMsgLoaded + 1
End Sub


Private Sub Form_Unload(Cancel As Integer)
iMsgLoaded = iMsgLoaded - 1
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

Private Sub Timer_Msg_Timer()
Timer_Msg.Enabled = False

s = "SELECT PK, Date_Time, Message, From_User" & _
    " From tbl_InstantMessaging " & _
    " WHERE (Opened = 0) " & _
    " AND (To_User = '" & gbl_UserName & "')" & _
    " AND (From_User = '" & Me.Caption & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    
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

Private Sub txtMessage_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdSend_Click
End Sub
