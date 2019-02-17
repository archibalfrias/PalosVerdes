VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInvClass 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   6630
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvClass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   600
      ScaleHeight     =   1455
      ScaleWidth      =   5295
      TabIndex        =   7
      Top             =   960
      Width           =   5295
      Begin VB.TextBox txtClassCode 
         Height          =   315
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   0
         Top             =   0
         Width           =   975
      End
      Begin VB.TextBox txtClassName 
         Height          =   315
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   1
         Top             =   360
         Width           =   3975
      End
      Begin VB.TextBox txtSectCode 
         Height          =   315
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtSectName 
         Height          =   315
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1080
         Width           =   3975
      End
      Begin VB.CommandButton Command1 
         Caption         =   ".."
         Height          =   315
         Left            =   2160
         TabIndex        =   9
         Top             =   720
         Width           =   315
      End
      Begin VB.TextBox txtSectKey 
         Height          =   315
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CLASSCODE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "CLASSNAME"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "SECTION CODE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "SECTION NAME"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   770
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   15000
      TabIndex        =   5
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   570
         Left            =   0
         TabIndex        =   6
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
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   2745
      Width           =   6630
      _ExtentX        =   11695
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8640
      Top             =   2520
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
            Picture         =   "frmInvClass.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvClass.frx":09CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvClass.frx":0B50
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvClass.frx":0E6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvClass.frx":1223
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvClass.frx":1675
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvClass.frx":1AC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvClass.frx":1E7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvClass.frx":1F91
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvClass.frx":24D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvClass.frx":262D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvClass.frx":2B6F
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInvClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim tmp As Long

Dim sCode, sNameTmp

Private Sub BROWSER(sName, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If sCode <> "" Then
            s = "SELECT TOP 1 tbl_Inv_Class.PK, tbl_Inv_Class.ClassCode, " & _
                " tbl_Inv_Class.ClassName, tbl_Inv_Class.SectKey, " & _
                " tbl_Inv_Class.LastModified, tbl_Inv_Section.SectCode, " & _
                " tbl_Inv_Section.SectName " & _
                " FROM tbl_Inv_Class LEFT OUTER JOIN " & _
                " tbl_Inv_Section ON tbl_Inv_Class.SectKey = tbl_Inv_Section.PK " & _
                " WHERE (tbl_Inv_Class.ClassName = '" & sName & "')" & _
                " ORDER BY tbl_Inv_Class.ClassName"
        Else
            s = "SELECT TOP 1 tbl_Inv_Class.PK, tbl_Inv_Class.ClassCode, " & _
                " tbl_Inv_Class.ClassName, tbl_Inv_Class.SectKey, " & _
                " tbl_Inv_Class.LastModified, tbl_Inv_Section.SectCode, " & _
                " tbl_Inv_Section.SectName " & _
                " FROM tbl_Inv_Class LEFT OUTER JOIN " & _
                " tbl_Inv_Section ON tbl_Inv_Class.SectKey = tbl_Inv_Section.PK " & _
                " ORDER BY tbl_Inv_Class.ClassName"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_Class.PK, tbl_Inv_Class.ClassCode, " & _
            " tbl_Inv_Class.ClassName, tbl_Inv_Class.SectKey, " & _
            " tbl_Inv_Class.LastModified, tbl_Inv_Section.SectCode, " & _
            " tbl_Inv_Section.SectName " & _
            " FROM tbl_Inv_Class LEFT OUTER JOIN " & _
            " tbl_Inv_Section ON tbl_Inv_Class.SectKey = tbl_Inv_Section.PK " & _
            " ORDER BY tbl_Inv_Class.ClassName"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_Class.PK, tbl_Inv_Class.ClassCode, " & _
            " tbl_Inv_Class.ClassName, tbl_Inv_Class.SectKey, " & _
            " tbl_Inv_Class.LastModified, tbl_Inv_Section.SectCode, " & _
            " tbl_Inv_Section.SectName " & _
            " FROM tbl_Inv_Class LEFT OUTER JOIN " & _
            " tbl_Inv_Section ON tbl_Inv_Class.SectKey = tbl_Inv_Section.PK " & _
            " WHERE (tbl_Inv_Class.ClassName < '" & sName & "')" & _
            " ORDER BY tbl_Inv_Class.ClassName DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_Class.PK, tbl_Inv_Class.ClassCode, " & _
            " tbl_Inv_Class.ClassName, tbl_Inv_Class.SectKey, " & _
            " tbl_Inv_Class.LastModified, tbl_Inv_Section.SectCode, " & _
            " tbl_Inv_Section.SectName " & _
            " FROM tbl_Inv_Class LEFT OUTER JOIN " & _
            " tbl_Inv_Section ON tbl_Inv_Class.SectKey = tbl_Inv_Section.PK " & _
            " WHERE (tbl_Inv_Class.ClassName > '" & sName & "')" & _
            " ORDER BY tbl_Inv_Class.ClassName "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_Class.PK, tbl_Inv_Class.ClassCode, " & _
            " tbl_Inv_Class.ClassName, tbl_Inv_Class.SectKey, " & _
            " tbl_Inv_Class.LastModified, tbl_Inv_Section.SectCode, " & _
            " tbl_Inv_Section.SectName " & _
            " FROM tbl_Inv_Class LEFT OUTER JOIN " & _
            " tbl_Inv_Section ON tbl_Inv_Class.SectKey = tbl_Inv_Section.PK " & _
            " ORDER BY tbl_Inv_Class.ClassName DESC"
    Case Else: Exit Sub
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtClassCode.Text = rs!ClassCode
    txtClassName.Text = rs!ClassName
    txtSectKey.Text = rs!SectKey
    txtSectCode.Text = rs!SectCode
    txtSectName.Text = rs!SectName
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    SaveSetting App.EXEName, "ClassName", "ClsName", rs!ClassName
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If AccessRights("Inventory Classification", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
s = "SELECT TOP 1 ClassCode" & _
    " From tbl_Inv_Class " & _
    " ORDER BY ClassCode DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtClassCode.Text = Format(CDbl(rs!ClassCode) + 1, "00#")
Else
    txtClassCode.Text = "001"
End If
rs.Close
TRANSACTIONTYPE = is_ADDING
'Me.Caption = "Classification - New"
txtClassName.SetFocus
End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("Inventory Classification", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_EDITTING
'Me.Caption = "Classification - Edit"
End Sub

Private Sub PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("Inventory Classification", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If MsgBox("ARE YOU SURE TO DELETE THIS RECORD?          ", vbCritical + vbYesNo + vbDefaultButton2, "CONFIRM") = vbNo Then Exit Sub
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_Inv_Class WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
CLEARTEXT

Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If Trim(txtClassName.Text) = "" Then MsgBox "Please Supply Class Name!                ", vbCritical, "Error...": txtClassName.SetFocus: Exit Sub
If RETURNTEXTVALUE(txtSectKey) = 0 Then MsgBox "Plase Supply Section!                 ", vbCritical, "Error...": txtSectCode.SetFocus: Exit Sub
On Error GoTo PG:
sCode = Trim(txtClassCode.Text)
sNameTmp = FORMATSQL(Trim(txtClassName.Text))
If TRANSACTIONTYPE = is_ADDING Then
    Do
        s = "SELECT tbl_Inv_Class.* " & _
            " FROM tbl_Inv_Class " & _
            " WHERE (ClassCode = '" & sCode & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If
        rs.Close
        sCode = Format(CDbl(sCode) + 1, "00#")
    Loop
    ConnOmega.Execute "INSERT INTO tbl_Inv_Class" & _
                      " (ClassCode, ClassName, SectKey, LastModified) " & _
                      " VALUES('" & sCode & "', " & _
                      " '" & sNameTmp & "', " & _
                      " " & RETURNTEXTVALUE(txtSectKey) & ", " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
End If
If TRANSACTIONTYPE = is_EDITTING Then
    ConnOmega.Execute "UPDATE tbl_Inv_Class" & _
                      " SET ClassName = '" & sNameTmp & "', " & _
                      " SectKey = " & RETURNTEXTVALUE(txtSectKey) & ", " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
End If
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
'Me.Caption = "Classification - Browse"
BROWSER sNameTmp, "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub

End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    Unload Me
Else
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    'Me.Caption = "Classification - Browse"
    BROWSER GetSetting(App.EXEName, "ClassName", "ClsName", ""), "is_LOAD"
    If Trim(txtClassCode.Text) = "" Then BROWSER GetSetting(App.EXEName, "ClassName", "ClsName", ""), "is_HOME"
End If
End Sub

Private Sub CLEARTEXT()
txtClassCode.Text = ""
txtClassName.Text = ""
txtSectKey.Text = "0"
txtSectCode.Text = ""
txtSectName.Text = ""
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtClassCode.Locked = True
txtClassName.Locked = bln
txtSectCode.Locked = bln
txtSectName.Locked = True
End Sub


Private Sub TOOLBARFUNC(intTrans As Integer)
Set Toolbar1.ImageList = ImageList1
With Toolbar1
    Select Case intTrans
        Case 1  '==== REFRESH ====
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
            .Buttons(9).ToolTipText = "BACK (Up)"
            .Buttons(11).ToolTipText = "NEXT (Down)"
            .Buttons(13).ToolTipText = "LAST (End)"
            .Buttons(15).ToolTipText = "FIND (F6)"
            .Buttons(17).ToolTipText = "CLOSE (Esc)"
        Case 2  '=== ADD/EDIT ===
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
        Case 3  '=== FIND ===
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
End Sub

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
'BROWSER GetSetting(App.EXEName, "ClassName", "ClsName", ""), "is_LOAD"
'If Trim(txtClassCode.Text) = "" Then BROWSER GetSetting(App.EXEName, "ClassName", "ClsName", ""), "is_HOME"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert:   PRESS_INSERT
    Case vbKeyF2:       PRESS_F2
    Case vbKeyDelete:   PRESS_DELETE
    Case vbKeyF5:       PRESS_F5
    Case vbKeyF6:       PRESS_F6
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "ClassName", "ClsName", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "ClassName", "ClsName", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "ClassName", "ClsName", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "ClassName", "ClsName", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
'Me.Caption = "Classification - Browse"
BROWSER GetSetting(App.EXEName, "ClassName", "ClsName", ""), "is_LOAD"
If Trim(txtClassCode.Text) = "" Then BROWSER GetSetting(App.EXEName, "ClassName", "ClsName", ""), "is_HOME"

tmp = SetWindowLong(txtClassCode.hwnd, GWL_STYLE, GetWindowLong(txtClassCode.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtClassName.hwnd, GWL_STYLE, GetWindowLong(txtClassName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSectCode.hwnd, GWL_STYLE, GetWindowLong(txtSectCode.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSectName.hwnd, GWL_STYLE, GetWindowLong(txtSectName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "ClassName", "ClsName", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "ClassName", "ClsName", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "ClassName", "ClsName", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "ClassName", "ClsName", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtClassCode_GotFocus()
HTEXT txtClassCode
End Sub

Private Sub txtClassName_GotFocus()
HTEXT txtClassName
End Sub

Private Sub txtSectCode_GotFocus()
HTEXT txtSectCode
End Sub

Private Sub txtSectCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtSectName.SetFocus
End Sub

Private Sub txtSectCode_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtSectCode_LostFocus()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
s = "SELECT PK, SectCode, SectName" & _
    " From tbl_Inv_Section " & _
    " WHERE (SectCode = '" & Format(RETURNTEXTVALUE(txtSectCode), "00#") & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtSectKey.Text = rs!PK
    txtSectCode.Text = rs!SectCode
    txtSectName.Text = rs!SectName
End If
rs.Close
End Sub

Private Sub txtSectName_GotFocus()
HTEXT txtSectName
End Sub
