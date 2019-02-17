VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelEmploymentStatus 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   6465
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPersonnelEmploymentStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5880
      Top             =   480
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
            Picture         =   "frmPersonnelEmploymentStatus.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelEmploymentStatus.frx":09CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelEmploymentStatus.frx":0B50
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelEmploymentStatus.frx":0E6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelEmploymentStatus.frx":1223
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelEmploymentStatus.frx":1675
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelEmploymentStatus.frx":1AC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelEmploymentStatus.frx":1E7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelEmploymentStatus.frx":1F91
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelEmploymentStatus.frx":24D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelEmploymentStatus.frx":262D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelEmploymentStatus.frx":2B6F
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
      ScaleWidth      =   5895
      TabIndex        =   3
      Top             =   960
      Width           =   5895
      Begin VB.TextBox txtStatusName 
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   360
         Width           =   4575
      End
      Begin VB.TextBox txtStatusCode 
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Top             =   0
         Width           =   975
      End
      Begin VB.TextBox txtStatusCode_1 
         Height          =   315
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C6B8A4&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1200
         ScaleHeight     =   375
         ScaleWidth      =   2535
         TabIndex        =   4
         Top             =   720
         Width           =   2535
         Begin VB.CheckBox chkActive 
            BackColor       =   &H00C6B8A4&
            Caption         =   "Active"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   975
         End
         Begin VB.CheckBox chkInActive 
            BackColor       =   &H00C6B8A4&
            Caption         =   "Inactive"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1080
            TabIndex        =   5
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "STATUS NAME"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "STATUS CODE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "STATUS TYPE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   720
         Width           =   1095
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
      TabIndex        =   2
      Top             =   2220
      Width           =   6465
      _ExtentX        =   11404
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
Attribute VB_Name = "frmPersonnelEmploymentStatus"
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

Private Function AUTOCODE() As Long
s = "SELECT Max(tbl_Personnel_EmploymentStatus.StatusCode) AS Code" & _
    " FROM tbl_Personnel_EmploymentStatus"
rs.Open s, ConnOmega
AUTOCODE = CLng(IIf(IsNull(rs!Code), 0, rs!Code)) + 1
rs.Close
End Function

Private Function PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If AccessRights("Personnel Employment Status", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
TRANSACTIONTYPE = is_ADDING
TOOLBARFUNC 2
LOCKTEXT False
CLEARTEXT
'Me.Caption = "EMPLOYMENT STATUS - NEW"
txtStatusCode.Text = Format(AUTOCODE, "00#")
txtStatusName.SetFocus
End Function

Private Function PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If StatusBar.Panels(1).Text = "" Then Exit Function
If AccessRights("Personnel Employment Status", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
TRANSACTIONTYPE = is_EDITTING
TOOLBARFUNC 2
LOCKTEXT False
'Me.Caption = "EMPLOYMENT STATUS - EDIT"
txtStatusName.SetFocus
End Function

Private Function PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If StatusBar.Panels(1).Text = "" Then Exit Function
If AccessRights("Personnel Employment Status", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
If MsgBox("ARE YOU SURE TO DELETE THIS RECORD?      ", vbInformation + vbYesNo, "CONFIRMATION") = vbNo Then Exit Function
On Error GoTo PG:
DELETE_STATUS StatusBar.Panels(1).Text
CLEARTEXT
BROWSER GetSetting(App.EXEName, "PersonnelEmploymentStatus", "PerEmpStat", ""), "is_PAGEDOWN"
If Trim(txtStatusCode.Text) = "" Then
    BROWSER GetSetting(App.EXEName, "PersonnelEmploymentStatus", "PerEmpStat", ""), "is_HOME"
End If

Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
Exit Function
End Function

Private Function DELETE_STATUS(intPK)
s = "DELETE FROM tbl_Personnel_EmploymentStatus" & _
    " WHERE (PK = " & intPK & ")"
ConnOmega.Execute s, , -1
End Function

Private Function INSERT_STATUS(strCode, strName, strLastMod)
s = "INSERT INTO tbl_Personnel_EmploymentStatus" & _
    " (StatusCode, StatusName, " & _
    " LastModified)" & _
    " VALUES('" & strCode & "', '" & strName & "', " & _
    " '" & strLastMod & "')"
ConnOmega.Execute s, , -1
End Function

Private Function UPDATE_STATUS(intPK, strName, strLastMod)
s = "UPDATE tbl_Personnel_EmploymentStatus" & _
    " SET StatusName = '" & strName & "', " & _
    " LastModified = '" & strLastMod & "'" & _
    " WHERE (PK =" & intPK & ")"
ConnOmega.Execute s, , -1
End Function

Private Function PRESS_F5()
If Trim(txtStatusCode.Text) = "" Then
    MsgBox "Please Suply Code!              ", vbCritical, "Error..."
    txtStatusCode.SetFocus
    HTEXT txtStatusCode
    Exit Function
End If
If Trim(txtStatusName.Text) = "" Then
    MsgBox "Please Supply Status Name!          ", vbCritical, "Error..."
    txtStatusName.SetFocus
    HTEXT txtStatusName
    Exit Function
End If
If chkActive.Value = 0 And chkInActive.Value = 0 Then
    MsgBox "Please Select Status Type!          ", vbCritical, "Error..."
    Exit Function
End If
If TRANSACTIONTYPE = is_ADDING Then
    On Error GoTo PG:
    s = "INSERT INTO tbl_Personnel_EmploymentStatus" & _
        " (StatusCode, StatusName, " & _
        " LastModified, Active)" & _
        " VALUES('" & Trim(txtStatusCode.Text) & "', " & _
        " '" & FORMATSQL(Trim(txtStatusName.Text)) & "', " & _
        " '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
        " " & IIf(chkActive.Value = 1, 1, 2) & ")"
    ConnOmega.Execute s, , -1
    BROWSER FORMATSQL(Trim(txtStatusName.Text)), "is_LOAD"
    TRANSACTIONTYPE = is_REFRESH
    TOOLBARFUNC 1
    LOCKTEXT True
    'Me.Caption = "EMPLOYMENT STATUS - BROWSE"
ElseIf TRANSACTIONTYPE = is_EDITTING Then
    On Error GoTo PG:
    s = "UPDATE tbl_Personnel_EmploymentStatus" & _
        " SET StatusName = '" & FORMATSQL(Trim(txtStatusName.Text)) & "', " & _
        " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
        " Active = " & IIf(chkActive.Value = 1, 1, 2) & " " & _
        " WHERE (PK =" & StatusBar.Panels(1).Text & ")"
    ConnOmega.Execute s, , -1
    BROWSER FORMATSQL(Trim(txtStatusName.Text)), "is_LOAD"
    TRANSACTIONTYPE = is_REFRESH
    TOOLBARFUNC 1
    LOCKTEXT True
    'Me.Caption = "EMPLOYMENT STATUS - BROWSE"
End If
Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
Exit Function
End Function

Private Function PRESS_F6()
If TRANSACTIONTYPE = is_REFRESH Then
'    PopupMenu frmPopUpMenu.mnuFindEmpStatus, , 5000, 400
End If
End Function

Private Function PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    Unload Me
Else
    BROWSER GetSetting(App.EXEName, "PersonnelEmploymentStatus", "PerEmpStat", ""), "is_LOAD"
    TRANSACTIONTYPE = is_REFRESH
    TOOLBARFUNC 1
    txtStatusCode_1.Visible = False
    LOCKTEXT True
    'Me.Caption = "EMPLOYMENT STATUS - BROWSE"
End If
End Function

Private Function FIND_CODE(strCode) As Long
s = "SELECT PK" & _
    " From tbl_Personnel_EmploymentStatus  " & _
    " WHERE (StatusCode='" & strCode & "')"
rs.Open s, ConnOmega
If Not rs.EOF Then
    FIND_CODE = IIf(IsNull(rs!PK), 0, rs!PK)
End If
rs.Close
End Function

Public Function BROWSER(strName, isWant As String)
Select Case isWant
    Case "is_LOAD"
        If strName <> "" Then
            s = "SELECT TOP 1 tbl_Personnel_EmploymentStatus.*" & _
                " From tbl_Personnel_EmploymentStatus " & _
                " WHERE (StatusName = '" & strName & "')" & _
                " ORDER BY StatusName"
        Else
            s = "SELECT TOP 1 tbl_Personnel_EmploymentStatus.*" & _
                " From tbl_Personnel_EmploymentStatus " & _
                " ORDER BY StatusName"
        End If
    Case "is_FIND"
        s = "SELECT TOP 1 tbl_Personnel_EmploymentStatus.*" & _
            " From tbl_Personnel_EmploymentStatus " & _
            " WHERE (PK = " & strName & ")" & _
            " ORDER BY StatusName DESC"
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Personnel_EmploymentStatus.*" & _
            " From tbl_Personnel_EmploymentStatus " & _
            " ORDER BY StatusName"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Personnel_EmploymentStatus.*" & _
            " From tbl_Personnel_EmploymentStatus " & _
            " WHERE (StatusName <'" & strName & "')" & _
            " ORDER BY StatusName DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Personnel_EmploymentStatus.*" & _
            " From tbl_Personnel_EmploymentStatus " & _
            " WHERE (StatusName >'" & strName & "')" & _
            " ORDER BY StatusName "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Personnel_EmploymentStatus.*" & _
            " From tbl_Personnel_EmploymentStatus " & _
            " ORDER BY StatusName DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtStatusCode.Text = rs!StatusCode
    txtStatusName.Text = rs!StatusName
    If rs!Active = 1 Then
        chkActive.Value = 1
        chkInActive.Value = 0
    Else
        chkActive.Value = 0
        chkInActive.Value = 1
    End If
    StatusBar.Panels(1).Text = rs!PK
    StatusBar.Panels(2).Text = IIf(IsNull(rs!LastModified), "", "LAST MODIFIED BY : " & rs!LastModified)
    SaveSetting App.EXEName, "PersonnelEmploymentStatus", "PerEmpStat", rs!StatusName
End If
rs.Close
End Function

Public Sub CLEARTEXT()
txtStatusCode.Text = ""
txtStatusName.Text = ""
chkActive.Value = 0
chkInActive.Value = 0
StatusBar.Panels(1).Text = ""
StatusBar.Panels(2).Text = ""
End Sub

Private Function LOCKTEXT(bln As Boolean)
If bln Then
    txtStatusCode.Locked = True
    txtStatusName.Locked = True
    Picture2.Enabled = False
Else
    txtStatusCode.Locked = False
    txtStatusName.Locked = False
    Picture2.Enabled = True
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

Private Sub chkActive_Click()
If chkActive.Value = 1 Then
    chkInActive.Value = 0
End If
End Sub

Private Sub chkInActive_Click()
If chkInActive.Value = 1 Then
    chkActive.Value = 0
End If
End Sub

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
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "PersonnelEmploymentStatus", "PerEmpStat", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "PersonnelEmploymentStatus", "PerEmpStat", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "PersonnelEmploymentStatus", "PerEmpStat", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "PersonnelEmploymentStatus", "PerEmpStat", ""), "is_END"
    Case vbKeyEscape:   PRESS_ESCAPE
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
TRANSACTIONTYPE = is_REFRESH
TOOLBARFUNC 1
LOCKTEXT True
'Me.Caption = "EMPLOYMENT STATUS - BROWSE"
BROWSER GetSetting(App.EXEName, "PersonnelEmploymentStatus", "PerEmpStat", ""), "is_LOAD"
If Trim(txtStatusName.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelEmploymentStatus", "PerEmpStat", ""), "is_HOME"

tmp = SetWindowLong(txtStatusCode.hwnd, GWL_STYLE, GetWindowLong(txtStatusCode.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtStatusName.hwnd, GWL_STYLE, GetWindowLong(txtStatusName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
'On Error Resume Next
'Me.Picture = LoadPicture(App.Path & "\images\new-6.jpg")
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
            Case "First": BROWSER GetSetting(App.EXEName, "PersonnelEmploymentStatus", "PerEmpStat", ""), "is_HOME"
        End Select
    Case "Back"
        Select Case Toolbar1.Buttons(9).Caption
            Case "Undo":  PRESS_ESCAPE
            Case "Back":  BROWSER GetSetting(App.EXEName, "PersonnelEmploymentStatus", "PerEmpStat", ""), "is_PAGEUP"
        End Select
    Case "Next":          BROWSER GetSetting(App.EXEName, "PersonnelEmploymentStatus", "PerEmpStat", ""), "is_PAGEDOWN"
    Case "Last":          BROWSER GetSetting(App.EXEName, "PersonnelEmploymentStatus", "PerEmpStat", ""), "is_END"
    Case "Find":          PRESS_F6
    Case "Close":         PRESS_ESCAPE
End Select
End Sub

Private Sub txtStatusCode_1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If TRANSACTIONTYPE = is_FINDING Then
        txtStatusCode_1.Text = Format(txtStatusCode_1.Text, "00#")
        If FIND_CODE(Format(txtStatusCode_1.Text, "00#")) <> 0 Then
            BROWSER FIND_CODE(Format(txtStatusCode_1.Text, "00#")), "is_FIND"
            TRANSACTIONTYPE = is_REFRESH
            TOOLBARFUNC 1
            txtStatusCode_1.Visible = False
            'Me.Caption = "EMPLOYMENT STATUS - BROWSE"
        Else
            MsgBox "UNABLE TO FIND '" & Format(txtStatusCode_1.Text, "00#") & "' IN THE DATABASE!      ", vbCritical, "ERROR..."
            txtStatusCode_1.SetFocus
            HTEXT txtStatusCode_1
        End If
    End If
End If
End Sub



