VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelDept 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPersonnelDept.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picSLine 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   840
      ScaleHeight     =   855
      ScaleWidth      =   8175
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   8175
      Begin RPVGCC.b8Container picADSLine1 
         Height          =   855
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   7815
         _extentx        =   13785
         _extenty        =   1508
         backcolor       =   8438015
         Begin VB.CommandButton cmdCOA 
            Caption         =   "..."
            Height          =   315
            Left            =   7320
            TabIndex        =   23
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Left            =   3720
            TabIndex        =   21
            Top             =   360
            Width           =   3555
         End
         Begin VB.TextBox txtAccountNo 
            Height          =   315
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   3435
         End
         Begin VB.TextBox txtAccountNo1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtAccountName1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtDebit1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtCredit1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Chart Of Accounts"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3720
            TabIndex        =   22
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "Payroll Account"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   120
            Width           =   1455
         End
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   9
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   10
         Top             =   105
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   1429
         ButtonWidth     =   1217
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   22
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
               Caption         =   "Print"
               Key             =   "Print"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               Key             =   "Refresh"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         MousePointer    =   99
         MouseIcon       =   "frmPersonnelDept.frx":08CA
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   15000
         Y1              =   1005
         Y2              =   1005
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
         Y1              =   910
         Y2              =   910
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   960
      ScaleHeight     =   3735
      ScaleWidth      =   7455
      TabIndex        =   0
      Top             =   1320
      Width           =   7455
      Begin MSComctlLib.ListView lstPayrollAcounts 
         Height          =   2295
         Left            =   0
         TabIndex        =   11
         Top             =   1440
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PayrollAccKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Payroll Account"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "COAKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Chart of Accounts"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   3240
         MaxLength       =   50
         TabIndex        =   8
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtDeptCode 
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
      Begin VB.TextBox txtDeptName 
         Height          =   315
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   2
         Top             =   360
         Width           =   5535
      End
      Begin VB.TextBox txtDeptCode_1 
         Height          =   315
         Left            =   3600
         MaxLength       =   3
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Related Accounts"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Chart of Accounts"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Code"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9000
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDept.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDept.frx":18BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDept.frx":2598
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDept.frx":3272
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDept.frx":3F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDept.frx":4C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDept.frx":5900
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDept.frx":65DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDept.frx":72B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDept.frx":7B8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDept.frx":8868
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDept.frx":9542
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDept.frx":A21C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDept.frx":AEF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDept.frx":BBD0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   24
      Top             =   5490
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   2469
            MinWidth        =   2469
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
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
End
Attribute VB_Name = "frmPersonnelDept"
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

Private Function BROWSER(strCode, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If strCode <> "" Then
            s = "SELECT TOP 1 PK, DepartmentCode, DepartmentName, " & _
                " LastModified" & _
                " From tbl_Personnel_Department " & _
                " WHERE (DepartmentName = '" & strCode & "')" & _
                " ORDER BY DepartmentName"
        Else
            s = "SELECT TOP 1 PK, DepartmentCode, DepartmentName, " & _
                " LastModified" & _
                " From tbl_Personnel_Department " & _
                " ORDER BY DepartmentName"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 PK, DepartmentCode, DepartmentName, " & _
            " LastModified" & _
            " From tbl_Personnel_Department " & _
            " ORDER BY DepartmentName"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 PK, DepartmentCode, DepartmentName, " & _
            " LastModified" & _
            " From tbl_Personnel_Department " & _
            " WHERE (DepartmentName < '" & strCode & "')" & _
            " ORDER BY DepartmentName DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 PK, DepartmentCode, DepartmentName, " & _
            " LastModified" & _
            " From tbl_Personnel_Department " & _
            " WHERE (DepartmentName > '" & strCode & "')" & _
            " ORDER BY DepartmentName "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 PK, DepartmentCode, DepartmentName, " & _
            " LastModified" & _
            " From tbl_Personnel_Department " & _
            " ORDER BY DepartmentName DESC"
    Case "is_FIND"
        s = "SELECT TOP 1 PK, DepartmentCode, DepartmentName, " & _
            " LastModified" & _
            " From tbl_Personnel_Department " & _
            " WHERE (PK = " & strCode & ") " & _
            " ORDER BY DepartmentName"
    Case Else: Exit Function
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtDeptCode.Text = rs!DepartmentCode
    txtDeptName.Text = rs!DepartmentName
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", "LAST MODIFIED BY : " & rs!LastModified)
    SaveSetting App.EXEName, "PersonnelDepartment", "PersonnelDept", rs!DepartmentName
End If
rs.Close
End Function

Private Function AUTOCODE() As Long
s = "SELECT Max(DepartmentCode) AS Code" & _
    " FROM tbl_Personnel_Department"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
AUTOCODE = CLng(IIf(IsNull(rs!Code), 0, rs!Code)) + 1
rs.Close
End Function

Private Function PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function

If AccessRights("Personnel Department", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If

TRANSACTIONTYPE = is_ADDING
TOOLBARFUNC 2
CLEARTEXT
LOCKTEXT False
'Me.Caption = "DEPARTMENT - NEW"
txtDeptCode.Text = Format(AUTOCODE, "00#")
txtDeptName.SetFocus
    
End Function

Private Function PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If Statusbar1.Panels(1).Text = "" Then Exit Function
If AccessRights("Personnel Department", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
TRANSACTIONTYPE = is_EDITTING
TOOLBARFUNC 2
LOCKTEXT False
'Me.Caption = "DEPARTMENT - EDIT"
txtDeptName.SetFocus
End Function

Private Function PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If Statusbar1.Panels(1).Text = "" Then Exit Function
If AccessRights("Personnel Department", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
If MsgBox("ARE YOU SURE TO DELETE THIS RECORD?      ", vbInformation + vbYesNo, "CONFIRMATION") = vbNo Then Exit Function
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_Personnel_Department" & _
                  " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "PersonnelDepartment", "PersonnelDept", ""), "is_PAGEDOWN"
If Trim(txtDeptCode.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelDepartment", "PersonnelDept", ""), "is_HOME"
Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
Exit Function
End Function

Private Function PRESS_F5()
If TRANSACTIONTYPE = is_ADDING Then
    On Error GoTo PG:
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Department" & _
                      " (DepartmentCode, DepartmentName, " & _
                      " LastModified)" & _
                      " VALUES('" & Trim(txtDeptCode.Text) & "', " & _
                      " '" & FORMATSQL(Trim(txtDeptName.Text)) & "', " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
    TRANSACTIONTYPE = is_REFRESH
    TOOLBARFUNC 1
    LOCKTEXT True
    BROWSER Trim(txtDeptName.Text), "is_LOAD"
    'Me.Caption = "DEPARTMENT - BROWSE"
ElseIf TRANSACTIONTYPE = is_EDITTING Then
    On Error GoTo PG:
    ConnOmega.Execute "UPDATE tbl_Personnel_Department" & _
                      " SET DepartmentName = '" & FORMATSQL(Trim(txtDeptName.Text)) & "', " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "'" & _
                      " WHERE (PK =" & Statusbar1.Panels(1).Text & ")"
    TRANSACTIONTYPE = is_REFRESH
    TOOLBARFUNC 1
    LOCKTEXT True
    BROWSER Trim(txtDeptName.Text), "is_LOAD"
    'Me.Caption = "DEPARTMENT - BROWSE"
End If
Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
Exit Function
End Function

Private Function PRESS_F6()
If TRANSACTIONTYPE = is_REFRESH Then
    'PopupMenu frmPopUpMenu.mnuFindDept, , 5000, 400
End If
End Function

Private Function PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    Unload Me
Else
    TRANSACTIONTYPE = is_REFRESH
    TOOLBARFUNC 1
    txtDeptCode_1.Visible = False
    LOCKTEXT True
    BROWSER GetSetting(App.EXEName, "PersonnelDepartment", "PersonnelDept", ""), "is_LOAD"
    'Me.Caption = "DEPARTMENT - BROWSE"
End If
End Function

Private Function FIND_CODE(strCode) As Long
s = "SELECT PK " & _
    " From tbl_Personnel_Department  " & _
    " WHERE (DepartmentCode ='" & strCode & "')"
rs.Open s, ConnOmega
If Not rs.EOF Then
    FIND_CODE = IIf(IsNull(rs!PK), 0, rs!PK)
End If
rs.Close
End Function

Public Sub CLEARTEXT()
txtDeptCode.Text = ""
txtDeptName.Text = ""
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
End Sub

Private Function LOCKTEXT(bln As Boolean)
If bln Then
    txtDeptCode.Locked = True
    txtDeptName.Locked = True
Else
    txtDeptCode.Locked = False
    txtDeptName.Locked = False
End If
End Function

Private Sub TOOLBARFUNC(intSelect As Integer)
With Toolbar1
    Select Case intSelect
        Case 1      '=== REFRESH ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 4
            .Buttons(9).Image = 5
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 12
            .Buttons(21).Image = 13
            '.Buttons(23).Image = 13
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "First"
            .Buttons(9).Caption = "Back"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            '.Buttons(19).Caption = "Post"
            .Buttons(19).Caption = "Refresh"
            .Buttons(21).Caption = "Close"
            .Buttons(1).Enabled = True
            .Buttons(3).Enabled = True
            .Buttons(5).Enabled = True
            .Buttons(7).Enabled = True
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = True
            .Buttons(13).Enabled = True
            .Buttons(15).Enabled = True
            .Buttons(17).Enabled = True
            .Buttons(19).Enabled = True
            .Buttons(21).Enabled = True
            '.Buttons(23).Enabled = True
            .Buttons(1).ToolTipText = "NEW (Ins)"
            .Buttons(3).ToolTipText = "EDIT (F2)"
            .Buttons(5).ToolTipText = "DELETE (Del)"
            .Buttons(7).ToolTipText = "FIRST (Home)"
            .Buttons(9).ToolTipText = "BACK (PgUp)"
            .Buttons(11).ToolTipText = "NEXT (PgDown)"
            .Buttons(13).ToolTipText = "LAST (End)"
            .Buttons(15).ToolTipText = "FIND (F6)"
            .Buttons(17).ToolTipText = "PRINT (F9)"
            '.Buttons(19).ToolTipText = "POST (F8)"
            .Buttons(19).ToolTipText = "REFRESH (F11)"
            .Buttons(21).ToolTipText = "CLOSE (Esc)"
        Case 2      '=== ADD/EDIT ====
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 14
            .Buttons(9).Image = 15
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 12
            .Buttons(21).Image = 13
            '.Buttons(23).Image = 13
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "Save"
            .Buttons(9).Caption = "Undo"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            '.Buttons(19).Caption = "Post"
            .Buttons(19).Caption = "Refresh"
            .Buttons(21).Caption = "Close"
            .Buttons(1).Enabled = False
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = False
            .Buttons(7).Enabled = True
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
            .Buttons(19).Enabled = False
            .Buttons(21).Enabled = False
            '.Buttons(23).Enabled = False
            .Buttons(1).ToolTipText = ""
            .Buttons(3).ToolTipText = ""
            .Buttons(5).ToolTipText = ""
            .Buttons(7).ToolTipText = "SAVE (F5)"
            .Buttons(9).ToolTipText = "UNDO (Esc)"
            .Buttons(11).ToolTipText = ""
            .Buttons(13).ToolTipText = ""
            .Buttons(15).ToolTipText = ""
            .Buttons(17).ToolTipText = ""
            .Buttons(19).ToolTipText = ""
            .Buttons(21).ToolTipText = ""
            '.Buttons(23).ToolTipText = ""
        Case 3      '=== FIND ===
           .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 4
            .Buttons(9).Image = 15
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 12
            .Buttons(21).Image = 13
            '.Buttons(23).Image = 13
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "First"
            .Buttons(9).Caption = "Undo"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            '.Buttons(19).Caption = "Post"
            .Buttons(19).Caption = "Refresh"
            .Buttons(21).Caption = "Close"
            .Buttons(1).Enabled = False
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = False
            .Buttons(7).Enabled = False
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
            .Buttons(19).Enabled = False
            .Buttons(21).Enabled = False
            '.Buttons(23).Enabled = False
            .Buttons(1).ToolTipText = ""
            .Buttons(3).ToolTipText = ""
            .Buttons(5).ToolTipText = ""
            .Buttons(7).ToolTipText = ""
            .Buttons(9).ToolTipText = "UNDO (Esc)"
            .Buttons(11).ToolTipText = ""
            .Buttons(13).ToolTipText = ""
            .Buttons(15).ToolTipText = ""
            .Buttons(17).ToolTipText = ""
            .Buttons(19).ToolTipText = ""
            .Buttons(21).ToolTipText = ""
            '.Buttons(23).ToolTipText = ""
        Case 4      '=== EMPTY DETAIL ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 14
            .Buttons(9).Image = 15
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 12
            .Buttons(21).Image = 13
            '.Buttons(23).Image = 13
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "Save"
            .Buttons(9).Caption = "Undo"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            '.Buttons(19).Caption = "Post"
            .Buttons(19).Caption = "Refresh"
            .Buttons(21).Caption = "Close"
            .Buttons(1).Enabled = True
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = False
            .Buttons(7).Enabled = True
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
            .Buttons(19).Enabled = False
            .Buttons(21).Enabled = False
            '.Buttons(23).Enabled = False
            .Buttons(1).ToolTipText = "NEW (Ins)"
            .Buttons(3).ToolTipText = ""
            .Buttons(5).ToolTipText = ""
            .Buttons(7).ToolTipText = "SAVE (F5)"
            .Buttons(9).ToolTipText = "UNDO (Esc)"
            .Buttons(11).ToolTipText = ""
            .Buttons(13).ToolTipText = ""
            .Buttons(15).ToolTipText = ""
            .Buttons(17).ToolTipText = ""
            .Buttons(19).ToolTipText = ""
            .Buttons(21).ToolTipText = ""
            '.Buttons(23).ToolTipText = ""
        Case 5      '=== NOT EMPTY DETAIL ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 14
            .Buttons(9).Image = 15
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 12
            .Buttons(21).Image = 13
            '.Buttons(23).Image = 13
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "Save"
            .Buttons(9).Caption = "Undo"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            '.Buttons(19).Caption = "Post"
            .Buttons(19).Caption = "Refresh"
            .Buttons(21).Caption = "Close"
            .Buttons(1).Enabled = True
            .Buttons(3).Enabled = True
            .Buttons(5).Enabled = True
            .Buttons(7).Enabled = True
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
            .Buttons(19).Enabled = False
            .Buttons(21).Enabled = False
            '.Buttons(23).Enabled = False
            .Buttons(1).ToolTipText = "NEW (Ins)"
            .Buttons(3).ToolTipText = "EDIT (F2)"
            .Buttons(5).ToolTipText = "DELET (Del)"
            .Buttons(7).ToolTipText = "SAVE (F5)"
            .Buttons(9).ToolTipText = "UNDO (Esc)"
            .Buttons(11).ToolTipText = ""
            .Buttons(13).ToolTipText = ""
            .Buttons(15).ToolTipText = ""
            .Buttons(17).ToolTipText = ""
            .Buttons(19).ToolTipText = ""
            .Buttons(21).ToolTipText = ""
            '.Buttons(23).ToolTipText = ""
    End Select
End With
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
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "PersonnelDepartment", "PersonnelDept", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "PersonnelDepartment", "PersonnelDept", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "PersonnelDepartment", "PersonnelDept", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "PersonnelDepartment", "PersonnelDept", ""), "is_END"
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
BROWSER GetSetting(App.EXEName, "PersonnelDepartment", "PersonnelDept", ""), "is_LOAD"
If Trim(txtDeptCode.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelDepartment", "PersonnelDept", ""), "is_HOME"
LOCKTEXT True
'Me.Caption = "DEPARTMENT - BROWSE"

tmp = SetWindowLong(txtDeptCode.hwnd, GWL_STYLE, GetWindowLong(txtDeptCode.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtDeptName.hwnd, GWL_STYLE, GetWindowLong(txtDeptName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
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
            Case "First": BROWSER GetSetting(App.EXEName, "PersonnelDepartment", "PersonnelDept", ""), "is_HOME"
        End Select
    Case "Back"
        Select Case Toolbar1.Buttons(9).Caption
            Case "Undo":  PRESS_ESCAPE
            Case "Back":  BROWSER GetSetting(App.EXEName, "PersonnelDepartment", "PersonnelDept", ""), "is_PAGEUP"
        End Select
    Case "Next":          BROWSER GetSetting(App.EXEName, "PersonnelDepartment", "PersonnelDept", ""), "is_PAGEDOWN"
    Case "Last":          BROWSER GetSetting(App.EXEName, "PersonnelDepartment", "PersonnelDept", ""), "is_END"
    Case "Find":          PRESS_F6
    Case "Close":         PRESS_ESCAPE
End Select
End Sub

Private Sub txtDeptCode_1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If TRANSACTIONTYPE = is_FINDING Then
        txtDeptCode_1.Text = Format(txtDeptCode_1.Text, "00#")
        If FIND_CODE(Format(txtDeptCode_1.Text, "00#")) <> 0 Then
            BROWSER FIND_CODE(Format(txtDeptCode_1.Text, "00#")), "is_FIND"
            TRANSACTIONTYPE = is_REFRESH
            TOOLBARFUNC 1
            txtDeptCode_1.Visible = False
        Else
            MsgBox "UNABLE TO FIND '" & Format(txtDeptCode_1.Text, "00#") & "' IN THE DATABASE!      ", vbCritical, "ERROR..."
            txtDeptCode_1.SetFocus
            HTEXT txtDeptCode_1
        End If
    End If
End If
End Sub

Private Sub txtDeptCode_1_LostFocus()
'If TRANSACTIONTYPE = is_FINDING Then
'    If FIND_CODE(CONCATINATE_CODE(txtDeptCode_1.Text)) <> 0 Then
'        SETFIELDS FIND_CODE(CONCATINATE_CODE(txtDeptCode_1.Text))
'        TRANSACTIONTYPE = is_REFRESH
'        TOOLBARFUNC 1
'        txtDeptCode_1.Visible = False
'    Else
'        MsgBox "UNABLE TO FIND '" & CONCATINATE_CODE(txtDeptCode_1.Text) & "' IN THE DATABASE!      ", vbCritical, "ERROR..."
'        txtDeptCode_1.SetFocus
'        htext txtDeptCode_1
'    End If
'End If
End Sub




