VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelPost 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPersonnelPost.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6480
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
            Picture         =   "frmPersonnelPost.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPost.frx":09CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPost.frx":0B50
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPost.frx":0E6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPost.frx":1223
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPost.frx":1675
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPost.frx":1AC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPost.frx":1E7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPost.frx":1F91
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPost.frx":24D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPost.frx":262D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPost.frx":2B6F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBody 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   360
      ScaleHeight     =   1095
      ScaleWidth      =   6015
      TabIndex        =   2
      Top             =   960
      Width           =   6015
      Begin VB.ComboBox cmbLevel 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   720
         Width           =   4455
      End
      Begin VB.TextBox txtPostName 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   4455
      End
      Begin VB.TextBox txtPostCode 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   0
         Width           =   975
      End
      Begin VB.TextBox txtPostCode_1 
         Height          =   315
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "LEVEL"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "POSITION NAME"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "POSITION CODE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1575
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
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   2295
      Width           =   6570
      _ExtentX        =   11589
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
Attribute VB_Name = "frmPersonnelPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2
Const is_FINDING = 3

Dim tmp As Long


Private Function BROWSER(strName, is_Action As String)

Select Case is_Action
    Case "is_LOAD"
        If strName <> "" Then
            s = "SELECT TOP 1 PK, PositionCode, PositionName, " & _
                " PositionLevel, LastModified" & _
                " From tbl_Personnel_Position " & _
                " WHERE (PositionName = '" & strName & "')" & _
                " ORDER BY PositionName "
        Else
            s = "SELECT TOP 1 PK, PositionCode, PositionName, " & _
                " PositionLevel, LastModified" & _
                " From tbl_Personnel_Position " & _
                " ORDER BY PositionName"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 PK, PositionCode, PositionName, " & _
            " PositionLevel, LastModified" & _
            " From tbl_Personnel_Position " & _
            " ORDER BY PositionName"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 PK, PositionCode, PositionName, " & _
            " PositionLevel, LastModified" & _
            " From tbl_Personnel_Position " & _
            " WHERE (PositionName <'" & strName & "')" & _
            " ORDER BY PositionName DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 PK, PositionCode, PositionName, " & _
            " PositionLevel, LastModified" & _
            " From tbl_Personnel_Position " & _
            " WHERE (PositionName >'" & strName & "')" & _
            " ORDER BY PositionName"
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 PK, PositionCode, PositionName, " & _
            " PositionLevel, LastModified" & _
            " From tbl_Personnel_Position " & _
            " ORDER BY PositionName DESC"
    Case "is_FIND"
        s = "SELECT TOP 1 PK, PositionCode, PositionName, " & _
            " PositionLevel, LastModified" & _
            " From tbl_Personnel_Position " & _
            " WHERE (PK = " & strName & ")" & _
            " ORDER BY PositionName DESC"
    Case Else: Exit Function
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtPostCode.Text = rs!PositionCode
    txtPostName.Text = rs!PositionName
    cmbLevel.ListIndex = rs!PositionLevel - 1
    StatusBar.Panels(1).Text = rs!PK
    StatusBar.Panels(2).Text = IIf(IsNull(rs!LastModified), "", "LAST MODIFIED BY : " & rs!LastModified)
    SaveSetting App.EXEName, "PersonnelPosition", "PersonnelPost", rs!PositionName
End If
rs.Close
End Function

Private Function AUTOCODE() As Long

s = "SELECT Max(PositionCode) AS Code" & _
    " FROM tbl_Personnel_Position"
rs.Open s, ConnOmega
AUTOCODE = CLng(IIf(IsNull(rs!Code), 0, rs!Code)) + 1
rs.Close
End Function

Private Function PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If AccessRights("Personnel Position", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
TRANSACTIONTYPE = is_ADDING
TOOLBARFUNC 2
CLEARTEXT
LOCKTEXT False
'Me.Caption = "POSITION - NEW"
txtPostCode.Text = Format(AUTOCODE, "00#")
txtPostName.SetFocus
End Function

Private Function PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If StatusBar.Panels(1).Text = "" Then Exit Function
    
If AccessRights("Personnel Position", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
        
TRANSACTIONTYPE = is_EDITTING
TOOLBARFUNC 2
LOCKTEXT False
'Me.Caption = "POSITION - EDIT"
txtPostName.SetFocus
        
End Function

Private Function PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If StatusBar.Panels(1).Text = "" Then Exit Function
If AccessRights("Personnel Position", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
If MsgBox("ARE YOU SURE TO DELETE THIS RECORD?      ", vbInformation + vbYesNo, "CONFIRMATION") = vbNo Then Exit Function
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_Personnel_Position WHERE (PK = " & StatusBar.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_PAGEDOWN"
If Trim(txtPostCode.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_HOME"
Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
Exit Function
End Function



Private Function PRESS_F5()
If TRANSACTIONTYPE = is_ADDING Then
    On Error GoTo PG:
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Position" & _
                      " (PositionCode, PositionName, " & _
                      " PositionLevel, LastModified)" & _
                      " VALUES('" & Trim(txtPostCode.Text) & "', " & _
                      " '" & FORMATSQL(Trim(txtPostName.Text)) & "', " & _
                      " " & cmbLevel.ListIndex + 1 & ", " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
    BROWSER FORMATSQL(Trim(txtPostName.Text)), "is_LOAD"
    TRANSACTIONTYPE = is_REFRESH
    TOOLBARFUNC 1
    LOCKTEXT True
    'Me.Caption = "POSITION - BROWSE"
ElseIf TRANSACTIONTYPE = is_EDITTING Then
    On Error GoTo PG:
    'UPDATE_POST StatusBar.Panels(1).Text, _
        FORMATSQL(Trim(txtPostName.Text)), _
        CStr(Now) & " - " & gbl_CompleteName
    ConnOmega.Execute "UPDATE tbl_Personnel_Position" & _
                      " SET PositionName = '" & FORMATSQL(Trim(txtPostName.Text)) & "', " & _
                      " PositionLevel = " & cmbLevel.ListIndex + 1 & ", " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "'" & _
                      " WHERE (PK = " & StatusBar.Panels(1).Text & ")"
    BROWSER FORMATSQL(Trim(txtPostName.Text)), "is_LOAD"
    TRANSACTIONTYPE = is_REFRESH
    TOOLBARFUNC 1
    LOCKTEXT True
    'Me.Caption = "POSITION - BROWSE"
End If
Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
Exit Function
End Function

Private Function PRESS_F6()
If TRANSACTIONTYPE = is_REFRESH Then
'    PopupMenu mnuFind, , 5000, 400
End If
End Function

Private Function PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    Unload Me
Else
    BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_LOAD"
    If Trim(txtPostName.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_HOME"
    TRANSACTIONTYPE = is_REFRESH
    TOOLBARFUNC 1
    LOCKTEXT True
    txtPostCode_1.Visible = False
    'Me.Caption = "POSITION - BROWSE"
End If
End Function

Private Function FIND_CODE(strCode) As Long

s = "SELECT PK" & _
    " From tbl_Personnel_Position  " & _
    " WHERE (PositionCode='" & strCode & "')"
rs.Open s, ConnOmega
If Not rs.EOF Then
    FIND_CODE = IIf(IsNull(rs!PK), 0, rs!PK)
End If
rs.Close
End Function


Public Sub CLEARTEXT()
txtPostCode.Text = ""
txtPostName.Text = ""
cmbLevel.ListIndex = -1
StatusBar.Panels(1).Text = ""
StatusBar.Panels(2).Text = ""
End Sub

Private Function LOCKTEXT(bln As Boolean)
If bln Then
    txtPostCode.Locked = True
    txtPostName.Locked = True
    cmbLevel.Locked = True
Else
    txtPostCode.Locked = False
    txtPostName.Locked = False
    cmbLevel.Locked = False
End If
End Function


Public Function TOOLBARFUNC(intSel As Integer)
With Toolbar1
    Select Case intSel
        Case 1      'REFRESH
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 10
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
            .Buttons(1).ToolTipText = "NEW (Ins)"
            .Buttons(3).ToolTipText = "EDIT (F2)"
            .Buttons(5).ToolTipText = "DELETE (Del)"
            .Buttons(7).ToolTipText = "FIRST (Home)"
            .Buttons(9).ToolTipText = "BACK (PgUp)"
            .Buttons(11).ToolTipText = "NEXT (PgDown)"
            .Buttons(13).ToolTipText = "LAST (End)"
            .Buttons(15).ToolTipText = "FIND (F6)"
            .Buttons(17).ToolTipText = "CLOSE (Esc)"
        Case 2      'ADD/EDIT
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 10
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
            .Buttons(1).ToolTipText = ""
            .Buttons(3).ToolTipText = ""
            .Buttons(5).ToolTipText = ""
            .Buttons(7).ToolTipText = "SAVE (F5)"
            .Buttons(9).ToolTipText = "UNDO (Esc)"
            .Buttons(11).ToolTipText = ""
            .Buttons(13).ToolTipText = ""
            .Buttons(15).ToolTipText = ""
            .Buttons(17).ToolTipText = ""
        Case 3      'FIND
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 10
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
            .Buttons(1).ToolTipText = ""
            .Buttons(3).ToolTipText = ""
            .Buttons(5).ToolTipText = ""
            .Buttons(7).ToolTipText = ""
            .Buttons(9).ToolTipText = "UNDO (Esc)"
            .Buttons(11).ToolTipText = ""
            .Buttons(13).ToolTipText = ""
            .Buttons(15).ToolTipText = ""
            .Buttons(17).ToolTipText = ""
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
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_END"
    Case vbKeyEscape:   PRESS_ESCAPE
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
With cmbLevel
    .Clear
    .AddItem "RANK IN FILE" '1
    .AddItem "SUPERVISORY"  '2
End With
CLEARTEXT
LOCKTEXT True
'Me.Caption = "POSITION - BROWSE"
TRANSACTIONTYPE = is_REFRESH
TOOLBARFUNC 1
BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_LOAD"
If Trim(txtPostName.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_HOME"

tmp = SetWindowLong(txtPostCode.hwnd, GWL_STYLE, GetWindowLong(txtPostCode.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPostName.hwnd, GWL_STYLE, GetWindowLong(txtPostName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If TRANSACTIONTYPE <> is_REFRESH Then
    Cancel = -1
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":           PRESS_INSERT
    Case "Edit":          PRESS_F2
    Case "Delete":        PRESS_DELETE
    Case "First"
        Select Case Toolbar1.Buttons(7).Caption
            Case "Save":  PRESS_F5
            Case "First": BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_HOME"
        End Select
    Case "Back"
        Select Case Toolbar1.Buttons(9).Caption
            Case "Undo":  PRESS_ESCAPE
            Case "Back":  BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_PAGEUP"
        End Select
    Case "Next":          BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_PAGEDOWN"
    Case "Last":          BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_END"
    Case "Find":          PRESS_F6
    Case "Close":         PRESS_ESCAPE
End Select
End Sub

Private Sub txtPostCode_1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If TRANSACTIONTYPE = is_FINDING Then
        txtPostCode_1.Text = Format(txtPostCode_1.Text, "00#")
        If FIND_CODE(Format(txtPostCode_1.Text, "00#")) <> 0 Then
            BROWSER FIND_CODE(Format(txtPostCode_1.Text, "00#")), "is_FIND"
            TRANSACTIONTYPE = is_REFRESH
            TOOLBARFUNC 1
            txtPostCode_1.Visible = False
        Else
            MsgBox "UNABLE TO FIND '" & Format(txtPostCode_1.Text, "00#") & "' IN THE DATABASE!      ", vbCritical, "ERROR..."
            txtPostCode_1.SetFocus
            HTEXT txtPostCode_1
        End If
    End If
End If
End Sub




