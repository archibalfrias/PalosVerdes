VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQuotes 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmQuotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMessage 
      Height          =   1815
      Left            =   1080
      MaxLength       =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   960
      Width           =   6735
   End
   Begin VB.TextBox txtAuthor 
      Height          =   330
      Left            =   1080
      TabIndex        =   3
      Top             =   2880
      Width           =   6735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7920
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuotes.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuotes.frx":048C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuotes.frx":0610
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuotes.frx":092A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuotes.frx":0CE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuotes.frx":1135
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuotes.frx":1587
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuotes.frx":193F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuotes.frx":1D91
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuotes.frx":1EEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuotes.frx":2045
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuotes.frx":2587
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture12 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
               Object.ToolTipText     =   "ADD (Ins)"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Edit"
               Key             =   "Edit"
               Object.ToolTipText     =   "EDIT (F2)"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Delete"
               Key             =   "Delete"
               Object.ToolTipText     =   "DELETE (Del)"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "First"
               Key             =   "First"
               Object.ToolTipText     =   "MOVE FIRST (Home)"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Back"
               Key             =   "Back"
               Object.ToolTipText     =   "PREVIOUS (PgUp)"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Next"
               Key             =   "Next"
               Object.ToolTipText     =   "NEXT (PgDown)"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Last"
               Key             =   "Last"
               Object.ToolTipText     =   "MOVE LAST (End)"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Find"
               Key             =   "Find"
               Object.ToolTipText     =   "FIND (F6)"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Load"
               Key             =   "Load"
               Object.ToolTipText     =   "LOAD MESSAGE(F8)"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
               Object.ToolTipText     =   "CLOSE (Esc)"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   15000
         Y1              =   750
         Y2              =   750
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
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   3420
      Width           =   8055
      _ExtentX        =   14208
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "QUOTES"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "AUTHOR"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   735
   End
End
Attribute VB_Name = "frmQuotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim tmp As Long

Dim sQuotes

Private Sub BROWSER(sQuotes, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If sQuotes <> "" Then
            s = "SELECT TOP 1 tbl_Greeting.*" & _
                " From tbl_Greeting " & _
                " WHERE (GGreeting = '" & FORMATSQL(CStr(sQuotes)) & "')"
        Else
            s = "SELECT TOP 1 tbl_Greeting.*" & _
                " From tbl_Greeting " & _
                " ORDER BY GGreeting"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Greeting.*" & _
            " From tbl_Greeting " & _
            " ORDER BY GGreeting"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Greeting.*" & _
            " From tbl_Greeting " & _
            " WHERE (GGreeting < '" & FORMATSQL(CStr(sQuotes)) & "')" & _
            " ORDER BY GGreeting DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Greeting.*" & _
            " From tbl_Greeting " & _
            " WHERE (GGreeting > '" & FORMATSQL(CStr(sQuotes)) & "')" & _
            " ORDER BY GGreeting "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Greeting.*" & _
            " From tbl_Greeting " & _
            " ORDER BY GGreeting DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtMessage.Text = rs!GGreeting
    txtAuthor.Text = rs!GAuthor
    StatusBar1.Panels(1).Text = rs!PK
    StatusBar1.Panels(2).Text = ""
    
    SaveSetting App.EXEName, "Quotes", "Quote", CStr(rs!GGreeting)
    
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
TRANSACTIONTYPE = is_ADDING
'Me.Caption = "Quotes - New"
TOOLBARFUNC 2
LOCKTEXT False
CLEARTEXT
txtMessage.SetFocus
End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If StatusBar1.Panels(1).Text = "" Then Exit Sub
TRANSACTIONTYPE = is_EDITTING
LOCKTEXT False
TOOLBARFUNC 2
'Me.Caption = "Quotes - Edit"
txtMessage.SetFocus
End Sub

Private Sub PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If StatusBar1.Panels(1).Text = "" Then Exit Sub

End Sub

Private Sub PRESS_F5()
If Trim(txtMessage.Text) = "" Then MsgBox "Please Seupply Quotes!              ", vbCritical, "Error...": txtMessage.SetFocus: HTEXT txtMessage: Exit Sub
If Trim(txtAuthor.Text) = "" Then MsgBox "Please Supply Author!               ", vbCritical, "Error...": txtAuthor.SetFocus:    HTEXT txtAuthor: Exit Sub
sQuotes = FORMATENTER(FORMATSQL(Trim(txtMessage.Text)))
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    ConnOmega.Execute "INSERT INTO tbl_Greeting" & _
                      " (GGreeting, GAuthor, LastModifiedBy) " & _
                      " VALUES('" & FORMATENTER(FORMATSQL(Trim(txtMessage.Text))) & "', " & _
                      " '" & FORMATSQL(Trim(txtAuthor.Text)) & "', " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
End If
If TRANSACTIONTYPE = is_EDITTING Then
    ConnOmega.Execute "UPDATE tbl_Greeting" & _
                      " SET GGreeting = '" & FORMATENTER(FORMATSQL(Trim(txtMessage.Text))) & "', " & _
                      " GAuthor = '" & FORMATSQL(Trim(txtAuthor.Text)) & "', " & _
                      " LastModifiedBy = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & StatusBar1.Panels(1).Text & ")"
End If
CLEARTEXT
TOOLBARFUNC 1
LOCKTEXT True
TRANSACTIONTYPE = is_REFRESH
'Me.Caption = "Quotes - Browse"
BROWSER sQuotes, "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
End Sub

Private Sub PRESS_F8()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If IsLoaded(frmBackground) = True Then frmBackground.txtQuotes.Text = "     " & Trim(txtMessage.Text) & "  (" & Trim(txtAuthor.Text) & ")": Exit Sub
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    Unload Me
Else
    CLEARTEXT
    TOOLBARFUNC 1
    LOCKTEXT True
    TRANSACTIONTYPE = is_REFRESH
    'Me.Caption = "Quotes - Browse"
    BROWSER GetSetting(App.EXEName, "Quotes", "Quote", ""), "is_LOAD"
    If Trim(txtMessage.Text) = "" Then BROWSER GetSetting(App.EXEName, "Quotes", "Quote", ""), "is_HOME"
End If
End Sub

Private Sub CLEARTEXT()
txtMessage.Text = ""
txtAuthor.Text = ""
StatusBar1.Panels(1).Text = ""
StatusBar1.Panels(2).Text = ""
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtMessage.Locked = bln
txtAuthor.Locked = bln
End Sub

Private Function TOOLBARFUNC(intSel As Integer)
With Toolbar1
    Select Case intSel
        Case 1      'REFRESH
            .Buttons.Item(1).Enabled = True
            .Buttons.Item(3).Enabled = True
            .Buttons.Item(5).Enabled = True
            .Buttons.Item(7).Image = 4
            .Buttons.Item(7).Caption = "First"
            .Buttons.Item(7).Enabled = True
            .Buttons.Item(9).Image = 5
            .Buttons.Item(9).Caption = "Back"
            .Buttons.Item(9).Enabled = True
            .Buttons.Item(11).Enabled = True
            .Buttons.Item(13).Enabled = True
            .Buttons.Item(15).Enabled = True
            .Buttons.Item(17).Enabled = True
            .Buttons.Item(19).Enabled = True
            .Buttons.Item(1).ToolTipText = "NEW (Ins)"
            .Buttons.Item(3).ToolTipText = "EDIT (F2)"
            .Buttons.Item(5).ToolTipText = "DELETE (Del)"
            .Buttons.Item(7).ToolTipText = "HOME (Home)"
            .Buttons.Item(9).ToolTipText = "PREVIOUS (PgUp)"
            .Buttons.Item(11).ToolTipText = "NEXT (PgDown)"
            .Buttons.Item(13).ToolTipText = "LAST (End)"
            .Buttons.Item(15).ToolTipText = "FIND (F6)"
            .Buttons.Item(17).ToolTipText = "FIND SPECIAL (F7)"
            .Buttons.Item(19).ToolTipText = "CLOSE (Esc)"
        Case 2      'ADD/EDIT
            .Buttons.Item(1).Enabled = False
            .Buttons.Item(3).Enabled = False
            .Buttons.Item(5).Enabled = False
            .Buttons.Item(7).Image = 11
            .Buttons.Item(7).Caption = "Save"
            .Buttons.Item(9).Image = 12
            .Buttons.Item(9).Caption = "Undo"
            .Buttons.Item(11).Enabled = False
            .Buttons.Item(13).Enabled = False
            .Buttons.Item(15).Enabled = False
            .Buttons.Item(17).Enabled = False
            .Buttons.Item(19).Enabled = False
            .Buttons.Item(1).ToolTipText = ""
            .Buttons.Item(3).ToolTipText = ""
            .Buttons.Item(5).ToolTipText = ""
            .Buttons.Item(7).ToolTipText = "SAVE (F5)"
            .Buttons.Item(9).ToolTipText = "UNDO (Esc)"
            .Buttons.Item(11).ToolTipText = ""
            .Buttons.Item(13).ToolTipText = ""
            .Buttons.Item(15).ToolTipText = ""
            .Buttons.Item(17).ToolTipText = ""
            .Buttons.Item(19).ToolTipText = ""
        Case 3      'FIND
            .Buttons.Item(1).Enabled = False
            .Buttons.Item(3).Enabled = False
            .Buttons.Item(5).Enabled = False
            .Buttons.Item(7).Enabled = False
            .Buttons.Item(9).Image = 12
            .Buttons.Item(9).Caption = "Undo"
            .Buttons.Item(11).Enabled = False
            .Buttons.Item(13).Enabled = False
            .Buttons.Item(15).Enabled = False
            .Buttons.Item(17).Enabled = False
            .Buttons.Item(19).Enabled = False
            .Buttons.Item(1).ToolTipText = ""
            .Buttons.Item(3).ToolTipText = ""
            .Buttons.Item(5).ToolTipText = ""
            .Buttons.Item(7).ToolTipText = ""
            .Buttons.Item(9).ToolTipText = "UNDO (Esc)"
            .Buttons.Item(11).ToolTipText = ""
            .Buttons.Item(13).ToolTipText = ""
            .Buttons.Item(15).ToolTipText = ""
            .Buttons.Item(17).ToolTipText = ""
            .Buttons.Item(19).ToolTipText = ""
    End Select
End With
End Function

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert:   PRESS_INSERT
    Case vbKeyF2:       PRESS_F2
    Case vbKeyDelete:   PRESS_DELETE
    Case vbKeyF5:       PRESS_F5
    Case vbKeyF6:       PRESS_F6
    Case vbKeyF8:       PRESS_F8
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "Quotes", "Quote", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "Quotes", "Quote", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "Quotes", "Quote", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "Quotes", "Quote", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.Height - Me.Height) / 4
Me.Left = (MainForm.Width - Me.Width) / 5
'Me.Caption = "Quotes - Browse"
CLEARTEXT
TOOLBARFUNC 1
LOCKTEXT True
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "Quotes", "Quote", ""), "is_LOAD"
If Trim(txtMessage.Text) = "" Then BROWSER GetSetting(App.EXEName, "Quotes", "Quote", ""), "is_HOME"

tmp = SetWindowLong(txtAuthor.hwnd, GWL_STYLE, GetWindowLong(txtAuthor.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "Quotes", "Quote", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "Quotes", "Quote", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "Quotes", "Quote", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "Quotes", "Quote", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Load":    PRESS_F8
    Case "Close":   PRESS_ESCAPE
End Select
End Sub
