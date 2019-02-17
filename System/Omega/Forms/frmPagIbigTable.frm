VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPagIbigTable 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPagIbigTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   1200
      ScaleHeight     =   1455
      ScaleWidth      =   3495
      TabIndex        =   3
      Top             =   1080
      Width           =   3495
      Begin VB.TextBox txtEmployeeShare 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtEmployerShare 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtPercentage 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtMaximum 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "EMPLOYEE SHARE"
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "EMPLOYER SHARE"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PERCENTAGE"
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "MAXIMUM"
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   0
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
            NumButtons      =   16
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
               Caption         =   "Close"
               Key             =   "Close"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5520
      Top             =   960
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
            Picture         =   "frmPagIbigTable.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagIbigTable.frx":09CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagIbigTable.frx":0B50
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagIbigTable.frx":0E6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagIbigTable.frx":1223
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagIbigTable.frx":1675
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagIbigTable.frx":1AC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagIbigTable.frx":1E7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagIbigTable.frx":1F91
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagIbigTable.frx":24D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagIbigTable.frx":262D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagIbigTable.frx":2B6F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   2985
      Width           =   5970
      _ExtentX        =   10530
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
Attribute VB_Name = "frmPagIbigTable"
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
'If AccessRights("Personnel Pag Ibig Table", "Edit") = False Then
'    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
'           "ACCESS DENIED!                                      ", vbCritical, "Alert"
'    Exit Function
'End If
TRANSACTIONTYPE = is_EDITTING
TOOLBARFUNC 2
LOCKTEXT False
'Me.Caption = "Pag Ibig Table - Edit"
txtMaximum.SetFocus
End Function

Private Function UPDATE_PAGIBIG(intPK, dblMax, dblPercent, _
dblEmployerShare, dblEmployeeShare, strLastMod)
Dim s As String
s = "UPDATE tbl_Personnel_PagIbigTable " & _
    " SET Maximum = " & CDbl(dblMax) & ", " & _
    " Percentage = " & CDbl(dblPercent) & ", " & _
    " EmployerShare = " & CDbl(dblEmployerShare) & ", " & _
    " EmployeeShare = " & CDbl(dblEmployeeShare) & ", " & _
    " LastModified = '" & strLastMod & "'" & _
    " WHERE (PK = " & intPK & ")"
ConnOmega.Execute s, , -1
End Function

Private Function PRESS_F5()
If TRANSACTIONTYPE = is_EDITTING Then
    On Error GoTo PG:
    UPDATE_PAGIBIG StatusBar.Panels(1).Text, _
        Trim(txtMaximum.Text), Trim(txtPercentage.Text), _
        Trim(txtEmployerShare.Text), Trim(txtEmployeeShare.Text), _
        CStr(Now) & " - " & gbl_CompleteName
    SETFIELDSLOAD GetSetting(App.EXEName, "PerPagIbigCtrl", "PerPagIbig", "")
    Picture1.SetFocus
    TRANSACTIONTYPE = is_REFRESH
    TOOLBARFUNC 1
    LOCKTEXT True
    'Me.Caption = "Pag Ibig Table - Browse"
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
    SETFIELDSLOAD GetSetting(App.EXEName, "PerPagIbigCtrl", "PerPagIbig", "")
    TRANSACTIONTYPE = is_REFRESH
    TOOLBARFUNC 1
    LOCKTEXT True
    Picture1.SetFocus
    'Me.Caption = "Pag Ibig Table - Browse"
End If
End Function

Private Function SETFIELDSLOAD(intPK)
Dim s As String
Dim rs As New ADODB.Recordset
If intPK <> "" Then
    s = "SELECT TOP 1 PK, Maximum, Percentage, " & _
        " EmployerShare, EmployeeShare, LastModified" & _
        " From tbl_Personnel_PagIbigTable " & _
        " WHERE (PK = " & intPK & ")" & _
        " ORDER BY PK "
Else
    s = "SELECT TOP 1 PK, Maximum, Percentage, " & _
        " EmployerShare, EmployeeShare, LastModified" & _
        " From tbl_Personnel_PagIbigTable " & _
        " ORDER BY PK "
End If
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega, adOpenForwardOnly, adLockOptimistic
If rs.RecordCount > 0 Then
    txtMaximum.Text = Format(rs!Maximum, "##,##0.00")
    txtPercentage.Text = Format(rs!Percentage, "##,##0.00")
    txtEmployerShare.Text = Format(rs!EmployerShare, "##,##0.00")
    txtEmployeeShare.Text = Format(rs!EmployeeShare, "##,##0.00")
    StatusBar.Panels(1).Text = rs!PK
    StatusBar.Panels(2).Text = IIf(IsNull(rs!LastModified), "", "LAST MODIFIED BY : " & rs!LastModified)
    SaveSetting App.EXEName, "PerPagIbigCtrl", "PerPagIbig", rs!PK
End If
rs.Close
End Function

Private Function LOCKTEXT(bln As Boolean)
If bln Then
    txtMaximum.Locked = True
    txtPercentage.Locked = True
    txtEmployerShare.Locked = True
    txtEmployeeShare.Locked = True
Else
    txtMaximum.Locked = False
    txtPercentage.Locked = False
    txtEmployerShare.Locked = False
    txtEmployeeShare.Locked = False
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
            .Item(15).Enabled = True
            .Item(7).Image = 4
            .Item(9).Image = 5
            .Item(7).Caption = "First"
            .Item(9).Caption = "Back"
            .Item(3).ToolTipText = "Edit (F2)"
            .Item(7).ToolTipText = ""
            .Item(9).ToolTipText = ""
            .Item(15).ToolTipText = "Close (Esc)"
        Case 2
            .Item(1).Enabled = False
            .Item(3).Enabled = False
            .Item(5).Enabled = False
            .Item(7).Enabled = True
            .Item(9).Enabled = True
            .Item(11).Enabled = False
            .Item(13).Enabled = False
            .Item(15).Enabled = False
            .Item(7).Image = 11
            .Item(9).Image = 12
            .Item(7).Caption = "Save"
            .Item(9).Caption = "Undo"
            .Item(3).ToolTipText = ""
            .Item(7).ToolTipText = "Save (F5)"
            .Item(9).ToolTipText = "Undo (Esc)"
            .Item(15).ToolTipText = ""
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
SETFIELDSLOAD GetSetting(App.EXEName, "PerPagIbigCtrl", "PerPagIbig", "")
TRANSACTIONTYPE = is_REFRESH
TOOLBARFUNC 1
LOCKTEXT True
'Me.Caption = "Pag Ibig Table - Browse"
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

Private Sub txtEmployeeShare_GotFocus()
txtEmployeeShare.Alignment = 0
If IsNumeric(txtEmployeeShare.Text) Then
    txtEmployeeShare.Text = CDbl(txtEmployeeShare.Text)
End If
HTEXT txtEmployeeShare
End Sub

Private Sub txtEmployeeShare_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtMaximum.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtEmployerShare.SetFocus
End If
End Sub

Private Sub txtEmployeeShare_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtEmployeeShare_LostFocus()
txtEmployeeShare.Alignment = 1
If Trim(txtEmployeeShare.Text) <> "" Then
    txtEmployeeShare.Text = Format(txtEmployeeShare.Text, "##,##0.00")
Else
    txtEmployeeShare.Text = "0.00"
End If
End Sub

Private Sub txtEmployerShare_GotFocus()
txtEmployerShare.Alignment = 0
If IsNumeric(txtEmployerShare.Text) Then
    txtEmployerShare.Text = CDbl(txtEmployerShare.Text)
End If
HTEXT txtEmployerShare
End Sub

Private Sub txtEmployerShare_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtEmployeeShare.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtPercentage.SetFocus
End If
End Sub

Private Sub txtEmployerShare_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtEmployerShare_LostFocus()
txtEmployerShare.Alignment = 1
If Trim(txtEmployerShare.Text) <> "" Then
    txtEmployerShare.Text = Format(txtEmployerShare.Text, "##,##0.00")
Else
    txtEmployerShare.Text = "0.00"
End If
End Sub

Private Sub txtMaximum_GotFocus()
txtMaximum.Alignment = 0
If IsNumeric(txtMaximum.Text) Then
    txtMaximum.Text = CDbl(txtMaximum.Text)
End If
HTEXT txtMaximum
End Sub

Private Sub txtMaximum_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtPercentage.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtEmployeeShare.SetFocus
End If
End Sub

Private Sub txtMaximum_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtMaximum_LostFocus()
txtMaximum.Alignment = 1
If Trim(txtMaximum.Text) <> "" Then
    txtMaximum.Text = Format(txtMaximum.Text, "##,##0.00")
Else
    txtMaximum.Text = "0.00"
End If
End Sub

Private Sub txtPercentage_GotFocus()
txtPercentage.Alignment = 0
If IsNumeric(txtPercentage.Text) Then
    txtPercentage.Text = CDbl(txtPercentage.Text)
End If
HTEXT txtPercentage
End Sub

Private Sub txtPercentage_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtEmployerShare.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtMaximum.SetFocus
End If
End Sub

Private Sub txtPercentage_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtPercentage_LostFocus()
txtPercentage.Alignment = 1
If Trim(txtPercentage.Text) <> "" Then
    txtPercentage.Text = Format(txtPercentage.Text, "##,##0.00")
Else
    txtPercentage.Text = "0.00"
End If
End Sub




