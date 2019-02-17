VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelOTRDRate 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   6675
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPersonnelOTRDRate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picBody 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1440
      ScaleHeight     =   855
      ScaleWidth      =   3615
      TabIndex        =   2
      Top             =   960
      Width           =   3615
      Begin VB.TextBox txtOT 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   0
         Width           =   1815
      End
      Begin VB.TextBox txtRD 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "OVERTIME (%)"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "REST DAY (%)"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
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
            NumButtons      =   18
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
               Caption         =   "Close"
               Key             =   "Close"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   15000
         Y1              =   90
         Y2              =   90
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   15000
         Y1              =   690
         Y2              =   690
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5640
      Top             =   1080
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
            Picture         =   "frmPersonnelOTRDRate.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelOTRDRate.frx":09CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelOTRDRate.frx":0B50
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelOTRDRate.frx":0E6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelOTRDRate.frx":1223
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelOTRDRate.frx":1675
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelOTRDRate.frx":1AC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelOTRDRate.frx":1E7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelOTRDRate.frx":1F91
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelOTRDRate.frx":24D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelOTRDRate.frx":262D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelOTRDRate.frx":2B6F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   7
      Top             =   2055
      Width           =   6675
      _ExtentX        =   11774
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
Attribute VB_Name = "frmPersonnelOTRDRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Private Function PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function

If AccessRights("Personnel Overtime/Restday Rate", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
TRANSACTIONTYPE = is_EDITTING
TOOLBARFUNC 2
LOCKTEXT False
'Me.Caption = "OVERTIME/REST DAY RATE - EDIT"
txtOT.SetFocus
End Function

Private Function UPDATE_PAGIBIG(intPK, dbOT, _
dblRD, strLastMod)
s = "UPDATE tbl_Personnel_OverTimeTable " & _
    " SET OverTime = " & CDbl(dbOT) & ", " & _
    " RestDay = " & CDbl(dblRD) & ", " & _
    " LastModified = '" & strLastMod & "'" & _
    " WHERE (PK = " & intPK & ")"
ConnOmega.Execute s, , -1
End Function

Private Function PRESS_F5()
If TRANSACTIONTYPE = is_EDITTING Then
    On Error GoTo PG:
    UPDATE_PAGIBIG StatusBar.Panels(1).Text, _
        Trim(txtOT.Text), Trim(txtRD.Text), _
        CStr(Now) & " - " & gbl_CompleteName
    SETFIELDSLOAD StatusBar.Panels(1).Text
    picBody.SetFocus
    TRANSACTIONTYPE = is_REFRESH
    TOOLBARFUNC 1
    LOCKTEXT True
    'Me.Caption = "OVERTIME/REST DAY RATE - BROWSE"
End If
Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error...."
Exit Function
End Function

Private Function PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    Unload Me
Else
    SETFIELDSLOAD GetSetting(App.EXEName, "PersonnelOTRDRate", "PerOTRDRate", "")
    TRANSACTIONTYPE = is_REFRESH
    TOOLBARFUNC 1
    LOCKTEXT True
    picBody.SetFocus
    'Me.Caption = "OVERTIME/REST DAY RATE - BROWSE"
End If
End Function

Private Function SETFIELDSLOAD(intPK)

If intPK <> "" Then
    s = "SELECT TOP 1 PK, OverTime, RestDay, " & _
        " LastModified" & _
        " From tbl_Personnel_OverTimeTable " & _
        " WHERE (PK = " & intPK & ")" & _
        " ORDER BY PK "
Else
    s = "SELECT TOP 1 PK, OverTime, RestDay, " & _
        " LastModified" & _
        " From tbl_Personnel_OverTimeTable " & _
        " ORDER BY PK "
End If
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtOT.Text = Format(rs!OverTime, "##,##0.00")
    txtRD.Text = Format(rs!RestDay, "##,##0.00")
    StatusBar.Panels(1).Text = rs!PK
    StatusBar.Panels(2).Text = IIf(IsNull(rs!LastModified), "", "LAST MODIFIED BY : " & rs!LastModified)
    SaveSetting App.EXEName, "PersonnelOTRDRate", "PerOTRDRate", rs!PK
End If
rs.Close
End Function

Private Function LOCKTEXT(bln As Boolean)
If bln Then
    txtOT.Locked = True
    txtRD.Locked = True
Else
    txtOT.Locked = False
    txtRD.Locked = False
End If
End Function

Private Function TOOLBARFUNC(intSel As Integer)
With Toolbar1.Buttons
    Select Case intSel
        Case 1      'REFRSEH
            .Item(1).Enabled = False
            .Item(3).Enabled = True
            .Item(5).Enabled = False
            .Item(7).Enabled = False
            .Item(9).Enabled = False
            .Item(11).Enabled = False
            .Item(13).Enabled = False
            .Item(15).Enabled = False
            .Item(17).Enabled = True
            .Item(7).Image = 4
            .Item(9).Image = 5
            .Item(7).Caption = "First"
            .Item(9).Caption = "Back"
            .Item(3).ToolTipText = "Edit (F2)"
            .Item(7).ToolTipText = ""
            .Item(9).ToolTipText = ""
            .Item(15).ToolTipText = ""
            .Item(17).ToolTipText = "Close (Esc)"
        Case 2
            .Item(1).Enabled = False
            .Item(3).Enabled = False
            .Item(5).Enabled = False
            .Item(7).Enabled = True
            .Item(9).Enabled = True
            .Item(11).Enabled = False
            .Item(13).Enabled = False
            .Item(15).Enabled = False
            .Item(17).Enabled = False
            .Item(7).Image = 11
            .Item(9).Image = 12
            .Item(7).Caption = "Save"
            .Item(9).Caption = "Undo"
            .Item(3).ToolTipText = ""
            .Item(7).ToolTipText = "Save (F5)"
            .Item(9).ToolTipText = "Undo (Esc)"
            .Item(15).ToolTipText = ""
            .Item(17).ToolTipText = ""
    End Select
End With
End Function

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF2:     PRESS_F2
    Case vbKeyF5:     PRESS_F5
    Case vbKeyEscape: PRESS_ESCAPE
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
SETFIELDSLOAD GetSetting(App.EXEName, "PersonnelOTRDRate", "PerOTRDRate", "")
TRANSACTIONTYPE = is_REFRESH
TOOLBARFUNC 1
LOCKTEXT True
'Me.Caption = "OVERTIME/REST DAY RATE - BROWSE"
'On Error Resume Next
'Me.Picture = LoadPicture(App.Path & "\images\new-6.jpg")
'picTab.Picture = LoadPicture(App.Path & "\images\new-6.jpg")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If TRANSACTIONTYPE <> is_REFRESH Then
    Cancel = -1
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Edit":         PRESS_F2
    Case "First"
        Select Case Toolbar1.Buttons(7).Caption
            Case "Save": PRESS_F5
        End Select
    Case "Back"
        Select Case Toolbar1.Buttons(9).Caption
            Case "Undo": PRESS_ESCAPE
        End Select
    Case "Close":        PRESS_ESCAPE
End Select
End Sub

Private Sub txtOT_GotFocus()
txtOT.Alignment = 0
If IsNumeric(txtOT.Text) Then
    txtOT.Text = CDbl(txtOT.Text)
End If
HTEXT txtOT
End Sub

Private Sub txtOT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtRD.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtRD.SetFocus
End If
End Sub

Private Sub txtOT_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtOT_LostFocus()
txtOT.Alignment = 1
If Trim(txtOT.Text) <> "" Then
    txtOT.Text = Format(txtOT.Text, "##,##0.00")
Else
    txtOT.Text = "0.00"
End If
End Sub

Private Sub txtRD_GotFocus()
txtRD.Alignment = 0
If IsNumeric(txtRD.Text) Then
    txtRD.Text = CDbl(txtRD.Text)
End If
HTEXT txtRD
End Sub

Private Sub txtRD_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtOT.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtOT.SetFocus
End If
End Sub

Private Sub txtRD_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtRD_LostFocus()
txtRD.Alignment = 1
If Trim(txtRD.Text) <> "" Then
    txtRD.Text = Format(txtRD.Text, "##,##0.00")
Else
    txtRD.Text = "0.00"
End If
End Sub





