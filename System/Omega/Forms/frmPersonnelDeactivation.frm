VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelDeactivation 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPersonnelDeactivation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picSearch 
      Height          =   4335
      Left            =   2400
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   7646
      BackColor       =   15266266
      Begin VB.TextBox txtSearchSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   3375
      End
      Begin VB.ListBox lstResultSearch 
         Height          =   2205
         Left            =   120
         TabIndex        =   34
         Top             =   885
         Width           =   3375
      End
      Begin VB.CommandButton cmdCancelSearch 
         Height          =   480
         Left            =   1920
         Picture         =   "frmPersonnelDeactivation.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3600
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKSearch 
         Height          =   480
         Left            =   120
         Picture         =   "frmPersonnelDeactivation.frx":1426
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3600
         Width           =   1560
      End
      Begin VB.ComboBox cmbEffectivityDate 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   3240
         Width           =   2055
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   40
         TabIndex        =   36
         Top             =   40
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   609
         Caption         =   "Search"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   8.25
         AutoFunction    =   0   'False
         Icon            =   "frmPersonnelDeactivation.frx":1A98
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Effectivity Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   3240
         Width           =   1335
      End
   End
   Begin RPVGCC.b8Container picAdd 
      Height          =   4335
      Left            =   2400
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   7646
      BackColor       =   15266266
      Begin VB.CommandButton cmdOK 
         Height          =   480
         Left            =   120
         Picture         =   "frmPersonnelDeactivation.frx":2032
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   3600
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancel 
         Height          =   480
         Left            =   1920
         Picture         =   "frmPersonnelDeactivation.frx":26A4
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3600
         Width           =   1560
      End
      Begin VB.ListBox lstResult 
         Height          =   2595
         Left            =   120
         TabIndex        =   26
         Top             =   890
         Width           =   3375
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   3375
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   40
         TabIndex        =   29
         Top             =   40
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   609
         Caption         =   "Search"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   8.25
         AutoFunction    =   0   'False
         Icon            =   "frmPersonnelDeactivation.frx":2E00
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   240
      ScaleHeight     =   3255
      ScaleWidth      =   7335
      TabIndex        =   1
      Top             =   840
      Width           =   7335
      Begin VB.ComboBox cmbEmploymentStatus 
         Height          =   315
         Left            =   1800
         TabIndex        =   21
         Text            =   "cmbEmploymentStatus"
         Top             =   2160
         Width           =   5535
      End
      Begin VB.TextBox txtRemarks 
         Height          =   315
         Left            =   1800
         TabIndex        =   19
         Top             =   2520
         Width           =   5505
      End
      Begin VB.TextBox txtEffectDate 
         Height          =   315
         Left            =   1800
         TabIndex        =   17
         Top             =   2880
         Width           =   1665
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1800
         TabIndex        =   15
         Top             =   360
         Width           =   5505
      End
      Begin VB.TextBox txtCtrl 
         Height          =   315
         Left            =   1800
         TabIndex        =   14
         Top             =   0
         Width           =   5505
      End
      Begin VB.TextBox txtCompRate 
         Height          =   315
         Left            =   5400
         TabIndex        =   13
         Top             =   2880
         Width           =   1905
      End
      Begin VB.TextBox txtPosition 
         Height          =   315
         Left            =   1800
         TabIndex        =   12
         Top             =   1800
         Width           =   5505
      End
      Begin VB.TextBox txtTaxStatus 
         Height          =   315
         Left            =   1800
         TabIndex        =   11
         Top             =   1440
         Width           =   5505
      End
      Begin VB.TextBox txtDept 
         Height          =   315
         Left            =   1800
         TabIndex        =   10
         Top             =   1080
         Width           =   5505
      End
      Begin VB.TextBox txtDivision 
         Height          =   315
         Left            =   1800
         TabIndex        =   9
         Top             =   720
         Width           =   5505
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "EMPLOYMENT STATUS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "REMARKS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   2565
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "EFFECTIVITY DATE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "POSITION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "DEPARTMENT"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "RATE COMPENSATION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   6
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "TAX STATUS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "DIVISION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NAME"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "CTRL NO."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   770
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   15000
      TabIndex        =   22
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   570
         Left            =   0
         TabIndex        =   23
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
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Edit"
               Key             =   "Edit"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Delete"
               Key             =   "Delete"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Save"
               Key             =   "Save"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Undo"
               Key             =   "Undo"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Find"
               Key             =   "Find"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Print"
               Key             =   "Print"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
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
      Left            =   7320
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483648
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeactivation.frx":339A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeactivation.frx":349C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeactivation.frx":3620
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeactivation.frx":393A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeactivation.frx":3A4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeactivation.frx":3F8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeactivation.frx":40E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelDeactivation.frx":462A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   4290
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10936
            MinWidth        =   10936
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPersonnelDeactivation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim EmployeeNo          As Double
Dim Division            As Double
Dim Department          As Double
Dim TaxStatus           As Double
Dim Positions           As Double
Dim CompRate            As Double
Dim EmploymentStatus    As Double
Dim Basic               As Double
Dim Cola                As Double
Dim Allowance           As Double
Dim WithSSS             As Long
Dim WithPagIbig         As Long
Dim WithPHIC            As Long
Dim WithTIN             As Long
Dim SSSNum              As String
Dim PagIbigNum          As String
Dim PHICNum             As String
Dim TINNum              As String
Dim tmp                 As Long

Dim TRANSACTIONTYPE     As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim Arr, strCtrl, dblRatePerHour, dblAllowanceRate, dblColaPerHour, _
strDeptFrom, strPostFrom, strStatusFrom, strTaxStatus, dblBasic, dblAllowance, Compensation

Private Function BROWSER(strCtrl)

s = "sp_Personnel_Action_Browse('" & strCtrl & "',0) "
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    EmployeeNo = rs!EmpPK
    Division = rs!Division
    Department = rs!Dept
    TaxStatus = rs!TaxStatus
    Positions = rs!Positions
    CompRate = rs!CompensationRate
    EmploymentStatus = rs!EmpStatus
    Basic = rs!Basic
    Cola = rs!Cola
    Allowance = rs!Allowance
    WithSSS = rs!Is_SSS
    WithPagIbig = rs!Is_PAGIBIG
    WithPHIC = rs!Is_PHIC
    WithTIN = rs!Is_TIN
    SSSNum = rs!SSS
    PagIbigNum = rs!PAGIBIG
    PHICNum = rs!PHIC
    TINNum = rs!TIN
    
    txtCtrl.Text = rs!CntrlNo
    txtName.Text = rs!IDNumber & " - " & rs!EmployeeName
    txtDivision.Text = IIf(IsNull(rs!Division), "", IIf(rs!Division = 1, "CLUB HOUSE", IIf(rs!Division = 2, "MAINTENANCE", "")))
    Arr = Split(DEPT_NAME(rs!Dept), ";", -1, 1)
    txtDept.Text = CStr(Arr(1))
    Arr = Split(TAX_STATUS_NAME(rs!TaxStatus), ";", -1, 1)
    txtTaxStatus.Text = CStr(Arr(1))
    Arr = Split(POSITION_NAME(rs!Positions), ";", -1, 1)
    txtPosition.Text = CStr(Arr(1))
    txtRemarks.Text = rs!Remarks
    txtEffectDate.Text = Format(rs!EffectivityDate, "mm/dd/yyyy")
    txtCompRate.Text = IIf(rs!CompensationRate = 1, "MONTHLY", "DAILY")
    Arr = Split(EMP_STATUS(rs!EmpStatus), ";", -1, 1)
    cmbEmploymentStatus.Text = CStr(Arr(1))
    StatusBar.Panels(1).Text = rs!PK
    StatusBar.Panels(2).Text = IIf(IsNull(rs!LastModified), "", "LAST MODIFIED BY : " & rs!LastModified)
    
    SaveSetting App.EXEName, "PersonnelDeactivationCtrl", "PerDectCtrl", rs!CntrlNo
    
End If
rs.Close
End Function


Private Function PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If picAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
If AccessRights("Personnel Action Memo", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
picMain.Enabled = False
picToolbar.Enabled = False
picAdd.ZOrder 0
txtSearch.Text = ""
picAdd.Visible = True
txtSearch.SetFocus
End Function

Private Function PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If StatusBar.Panels(1).Text = "" Then Exit Function
If picAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
If AccessRights("Personnel Action Memo", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If

End Function

Private Function PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If StatusBar.Panels(1).Text = "" Then Exit Function
If picAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
If AccessRights("Personnel Action Memo", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If

s = "SELECT ActionMemo " & _
    " From tbl_Personnel_Compensation " & _
    " WHERE (ActionMemo = " & StatusBar.Panels(1).Text & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    MsgBox "CANNOT BE DELETED!          " & vbCrLf & _
           "" & vbCrLf & _
           "Action used by the Payroll Transaction.   " & vbCrLf & _
           "Any changes have an effect on Payroll Computation.  ", vbCritical, "Error..."
    Exit Function
Else
    If MsgBox("ARE YOU SURE TO DELETE THIS RECORD?          ", vbCritical + vbYesNo, "CONFIRMATION") = vbNo Then Exit Function
    
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Action " & _
                      " WHERE (PK = " & StatusBar.Panels(1).Text & ")"
        
    CLEARTEXT
    TOOLBAR_FUNC 0
End If
rs.Close

End Function

Private Function PRESS_F5()
If picAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
If IsDate(txtEffectDate.Text) = False Then MsgBox "Please Supply a Valid Date!            ", vbCritical, "Error...": txtEffectDate.SetFocus: HTEXT txtEffectDate: Exit Function
If EmploymentStatus = 0 Then MsgBox "Please Select Status!                  ", vbCritical, "Error...": cmbEmploymentStatus.SetFocus: Exit Function
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    If DateValue(GET_LAST_ACTION_EFFECTIVITY(EmployeeNo)) > DateValue(FormatDateTime(txtEffectDate.Text, vbShortDate)) Then
        MsgBox "EFFECTIVITY DATE MUST BE HIGHER THAN THE LAST ACTION MEMO!          ", vbInformation, "Error..."
        Exit Function
    End If
    strCtrl = ""
    s = "SELECT TOP 1 CntrlNo" & _
        " From tbl_Personnel_Action " & _
        " ORDER BY CntrlNo DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        strCtrl = Format(CLng(rs!CntrlNo) + 1, "0000000#")
    Else
        strCtrl = Format(1, "0000000#")
    End If
    rs.Close
    
    Do
        s = "SELECT tbl_Personnel_Action.* " & _
            " From tbl_Personnel_Action " & _
            " WHERE (CntrlNo = '" & strCtrl & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If
        rs.Close
        strCtrl = Format(CDbl(strCtrl) + 1, "0000000#")
    Loop

    dblRatePerHour = 0: dblAllowanceRate = 0: dblColaPerHour = 0
    If CompRate = 1 Then
        dblRatePerHour = ((CDbl(Basic) / 2) / 13.08333) / 8
        dblAllowanceRate = ((CDbl(Allowance) / 2) / 13.08333) / 8
        dblColaPerHour = ((CDbl(Cola) / 2) / 13.08333) / 8
    ElseIf CompRate = 2 Then
        dblRatePerHour = CDbl(Basic) / 8
        dblAllowanceRate = CDbl(Allowance) / 8
        dblColaPerHour = CDbl(Cola) / 8
    End If

    ConnOmega.Execute "INSERT INTO tbl_Personnel_Action" & _
                        " (CntrlNo, EmpPK, Division, Dept, EmpStatus, TaxStatus," & _
                        " Positions, CompensationRate, Is_PAGIBIG, Is_PHIC, Is_SSS, " & _
                        " Is_TIN, SSS, PAGIBIG, PHIC, TIN, Remarks, EffectivityDate, " & _
                        " Basic, RatePerHourBasic, Allowance, RatePerHourAllow, LastModified, " & _
                        " Cola, RatePerHourCola)" & _
                        " VALUES ('" & strCtrl & "', " & EmployeeNo & ", " & Division & ", " & _
                        " " & Department & ", " & EmploymentStatus & ", " & TaxStatus & ", " & Positions & ", " & _
                        " " & CompRate & ", " & WithPagIbig & ", " & WithPHIC & ", " & WithSSS & ", " & _
                        " " & WithTIN & ", '" & SSSNum & "', '" & PagIbigNum & "', '" & PHICNum & "', " & _
                        " '" & TINNum & "', '" & FORMATSQL(Trim(txtRemarks.Text)) & "', '" & FormatDateTime(txtEffectDate.Text, vbShortDate) & "', " & _
                        " " & CDbl(Basic) & ", " & CDbl(dblRatePerHour) & "," & CDbl(Allowance) & ", " & _
                        " " & CDbl(dblAllowanceRate) & ",'" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                        " " & CDbl(Cola) & ", " & CDbl(dblColaPerHour) & ")"
    
    LOCKTEXT True
    TOOLBAR_FUNC 1
    TRANSACTIONTYPE = is_REFRESH
    
    BROWSER strCtrl
    
End If
If TRANSACTIONTYPE = is_EDITTING Then
    
    dblRatePerHour = 0: dblAllowanceRate = 0: dblColaPerHour = 0
    If CompRate = 1 Then
        dblRatePerHour = ((CDbl(Basic) / 2) / 13.08333) / 8
        dblAllowanceRate = ((CDbl(Allowance) / 2) / 13.08333) / 8
        dblColaPerHour = ((CDbl(Cola) / 2) / 13.08333) / 8
    ElseIf CompRate = 2 Then
        dblRatePerHour = CDbl(Basic) / 8
        dblAllowanceRate = CDbl(Allowance) / 8
        dblColaPerHour = CDbl(Cola) / 8
    End If

    ConnOmega.Execute "UPDATE tbl_Personnel_Action" & _
                        " SET Remarks = '" & FORMATSQL(Trim(txtRemarks.Text)) & "', " & _
                        " EffectivityDate = '" & FormatDateTime(txtEffectDate.Text, vbShortDate) & "', " & _
                        " Basic = " & CDbl(Basic) & ", " & _
                        " RatePerHourBasic = " & CDbl(dblRatePerHour) & ", " & _
                        " Allowance = " & CDbl(Allowance) & ", " & _
                        " RatePerHourAllow = " & CDbl(dblAllowanceRate) & ", " & _
                        " Cola = " & CDbl(Cola) & ", " & _
                        " RatePerHourCola = " & CDbl(dblColaPerHour) & ", " & _
                        " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                        " WHERE (PK = " & StatusBar.Panels(1).Text & ")"
    
    LOCKTEXT True
    TOOLBAR_FUNC 1
    TRANSACTIONTYPE = is_REFRESH
    
    BROWSER GetSetting(App.EXEName, "PersonnelDeactivationCtrl", "PerDectCtrl", "")
    
End If
Exit Function
PG:
MsgBox Err.Number & Err.Description, vbCritical, "Error..."
Exit Function
End Function

Private Function PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If picAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
picToolbar.Enabled = False
picMain.Enabled = False
picSearch.ZOrder 0
txtSearchSearch.Text = ""
picSearch.Visible = True
txtSearchSearch.SetFocus
End Function

Private Function PRESS_F9()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If StatusBar.Panels(1).Text = "" Then Exit Function
If picAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function

If AccessRights("Personnel Action Memo", "Print") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If

s = "sp_Personnel_Action_Print(" & EmployeeNo & ", '" & FormatDateTime(txtEffectDate.Text, vbShortDate) & "', 0)"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    
    ConnOmega.Execute "DELETE FROM tbl_PersonnelAction_Tmp WHERE (LogInName = '" & gbl_UserName & "')"
    
    If rs!CompensationRate = 1 Then
        Compensation = "MONTHLY"
    ElseIf rs!CompensationRate = 2 Then
        Compensation = "DAILY"
    End If
    
    ConnOmega.Execute "INSERT INTO  tbl_PersonnelAction_Tmp " & _
                      " (CrtlNo, IDNo, Name, DateHired, EffectDate, SSS, " & _
                      " PHIC, PagIbig, TIN, Remarks, StatusTo, " & _
                      " PostTo, DeptTo, CompTo, " & _
                      " BasicTo, AllowanceTo, LogInName ) " & _
                      " VALUES('" & rs!CntrlNo & "', '" & rs!IDNumber & "', '" & rs!EmployeeName & "', " & _
                      " '" & CStr(UCase(Format(rs!DHired, "mmmm dd, yyyy"))) & "', '" & CStr(UCase(Format(rs!EffectivityDate, "mmmm dd, yyyy"))) & "', " & _
                      " '" & rs!SSS & "', " & _
                      " '" & rs!PHIC & "', '" & rs!PAGIBIG & "', '" & rs!TIN & "', " & _
                      " '" & FORMATSQL(rs!Remarks) & "', '" & rs!StatusName & "', " & _
                      " '" & rs!PositionName & "', '" & rs!DepartmentName & "', " & _
                      " '" & Compensation & "', " & _
                      " " & CDbl(rs!Basic) & ", " & _
                      " " & CDbl(rs!Allowance) & ", '" & gbl_UserName & "')"

    t = "sp_Personnel_Action_Print(" & EmployeeNo & ", '" & FormatDateTime(rs!EffectivityDate, vbShortDate) & "',1)"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        If rt!CompensationRate = 1 Then
            Compensation = "MONTHLY"
        ElseIf rt!CompensationRate = 2 Then
            Compensation = "DAILY"
        End If
        
        ConnOmega.Execute "UPDATE tbl_PersonnelAction_Tmp " & _
                          " SET StatusFrom = '" & rt!StatusName & "', " & _
                          " PostFrom = '" & rt!PositionName & "', " & _
                          " DeptFrom = '" & rt!DepartmentName & "', " & _
                          " CompFrom = '" & Compensation & "', " & _
                          " BasicFrom = " & CDbl(rt!Basic) & ", " & _
                          " AllowanceFrom = " & CDbl(rt!Allowance) & "" & _
                          " WHERE (CrtlNo = '" & rs!CntrlNo & "') " & _
                          " AND (LogInName ='" & gbl_UserName & "')"
        
    Else
    
        ConnOmega.Execute "UPDATE tbl_PersonnelAction_Tmp " & _
                          " SET StatusFrom = '', " & _
                          " PostFrom = '', " & _
                          " DeptFrom = '', " & _
                          " CompFrom = '', " & _
                          " BasicFrom = 0, " & _
                          " AllowanceFrom = 0 " & _
                          " WHERE (CrtlNo = '" & rs!CntrlNo & "')" & _
                          " AND (LogInName ='" & gbl_UserName & "')"
    End If
    rt.Close
End If
rs.Close

s = "SELECT tbl_PersonnelAction_Tmp.*" & _
    " FROM tbl_PersonnelAction_Tmp"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
rs.Requery
rs.Close

frmCrystalReportViewer.PRINT_ACTION_MEMO
If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show

End Function

Private Function PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picAdd.Visible = True Then cmdCancel_Click: Exit Function
    If picSearch.Visible = True Then cmdCancelSearch_Click: Exit Function
    Unload Me
Else
    CLEARTEXT
    LOCKTEXT True
    If StatusBar.Panels(1).Text = "" Then
        TOOLBAR_FUNC 0
    Else
        TOOLBAR_FUNC 1
        BROWSER GetSetting(App.EXEName, "PersonnelDeactivationCtrl", "PerDectCtrl", "")
    End If
    TRANSACTIONTYPE = is_REFRESH
End If
End Function

Private Function CLEARTEXT()
EmployeeNo = 0
Division = 0
Department = 0
TaxStatus = 0
Positions = 0
CompRate = 0
EmploymentStatus = 0
Basic = 0
Cola = 0
Allowance = 0
WithSSS = 0
WithPagIbig = 0
WithPHIC = 0
WithTIN = 0
SSSNum = ""
PagIbigNum = ""
PHICNum = ""
TINNum = ""
txtCtrl.Text = ""
txtName.Text = ""
txtDivision.Text = ""
txtDept.Text = ""
txtTaxStatus.Text = ""
txtPosition.Text = ""
txtRemarks.Text = ""
txtEffectDate.Text = ""
txtCompRate.Text = ""
cmbEmploymentStatus.Text = ""
cmbEmploymentStatus.ListIndex = -1
StatusBar.Panels(1).Text = ""
StatusBar.Panels(2).Text = ""
End Function

Private Function LOCKTEXT(bln As Boolean)
If bln Then
    txtCtrl.Locked = True
    txtName.Locked = True
    txtDivision.Locked = True
    txtDept.Locked = True
    txtTaxStatus.Locked = True
    txtPosition.Locked = True
    txtRemarks.Locked = True
    txtEffectDate.Locked = True
    txtCompRate.Locked = True
    cmbEmploymentStatus.Locked = True
Else
    txtEffectDate.Locked = False
    cmbEmploymentStatus.Locked = False
    txtRemarks.Locked = False
End If
End Function

Private Function TOOLBAR_FUNC(isSelect As Integer)
With Toolbar1
    Set .ImageList = ImageList1
    .Buttons(1).Image = 1
    .Buttons(3).Image = 2
    .Buttons(5).Image = 3
    .Buttons(7).Image = 7
    .Buttons(9).Image = 8
    .Buttons(11).Image = 4
    .Buttons(13).Image = 5
    .Buttons(15).Image = 6
    Select Case isSelect
        Case 0  'Empty Fields
            .Buttons(1).Enabled = True
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = False
            .Buttons(7).Enabled = False
            .Buttons(9).Enabled = False
            .Buttons(11).Enabled = True
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = True
            .Buttons(1).ToolTipText = "NEW (Ins)"
            .Buttons(3).ToolTipText = ""
            .Buttons(5).ToolTipText = ""
            .Buttons(7).ToolTipText = ""
            .Buttons(9).ToolTipText = ""
            .Buttons(11).ToolTipText = "FIND (F6)"
            .Buttons(13).ToolTipText = ""
            .Buttons(15).ToolTipText = "CLOSE (Esc)"
        Case 1
            .Buttons(1).Enabled = True
            .Buttons(3).Enabled = True
            .Buttons(5).Enabled = True
            .Buttons(7).Enabled = False
            .Buttons(9).Enabled = False
            .Buttons(11).Enabled = True
            .Buttons(13).Enabled = True
            .Buttons(15).Enabled = True
            .Buttons(1).ToolTipText = "NEW (Ins)"
            .Buttons(3).ToolTipText = "EDIT (F2)"
            .Buttons(5).ToolTipText = "DELETE (Del)"
            .Buttons(7).ToolTipText = ""
            .Buttons(9).ToolTipText = ""
            .Buttons(11).ToolTipText = "FIND (F6)"
            .Buttons(13).ToolTipText = "PRINT (F9)"
            .Buttons(15).ToolTipText = "CLOSE (Esc)"
        Case 2
            .Buttons(1).Enabled = False
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = False
            .Buttons(7).Enabled = True
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(1).ToolTipText = ""
            .Buttons(3).ToolTipText = ""
            .Buttons(5).ToolTipText = ""
            .Buttons(7).ToolTipText = "Save (F5)"
            .Buttons(9).ToolTipText = "Undo (Esc)"
            .Buttons(11).ToolTipText = ""
            .Buttons(13).ToolTipText = ""
            .Buttons(15).ToolTipText = ""
    End Select
End With
End Function

Private Sub b8TitleBar1_CLoseClick()
cmdCancel_Click
End Sub

Private Sub b8TitleBar2_CLoseClick()
cmdCancelSearch_Click
End Sub

Private Sub cmbEffectivityDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKSearch_Click
End Sub

Private Sub cmbEmploymentStatus_Click()
If cmbEmploymentStatus.ListIndex = -1 Then Exit Sub
EmploymentStatus = cmbEmploymentStatus.ItemData(cmbEmploymentStatus.ListIndex)
End Sub

Private Sub cmbEmploymentStatus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtRemarks.SetFocus
End Sub

Private Sub cmdCancel_Click()
picToolbar.Enabled = True
picMain.Enabled = True
picAdd.Visible = False
End Sub

Private Sub cmdCancelSearch_Click()
picToolbar.Enabled = True
picMain.Enabled = True
picSearch.Visible = False
End Sub

Private Sub cmdOK_Click()
If lstResult.ListIndex = -1 Then Exit Sub
CLEARTEXT
LOCKTEXT False
TOOLBAR_FUNC 2
EmployeeNo = lstResult.ItemData(lstResult.ListIndex)
Arr = Split(lstResult.List(lstResult.ListIndex), " - ", -1, 1)
s = "SELECT TOP 1 tbl_Personnel_Action.Division, tbl_Personnel_Action.Dept, " & _
    " tbl_Personnel_Action.TaxStatus, tbl_Personnel_Action.Positions, " & _
    " tbl_Personnel_Action.CompensationRate, tbl_Personnel_Action.Basic, " & _
    " tbl_Personnel_Action.Cola, tbl_Personnel_Action.Allowance, " & _
    " tbl_Personnel_Action.Is_SSS, tbl_Personnel_Action.Is_PHIC, " & _
    " tbl_Personnel_Action.Is_PAGIBIG, tbl_Personnel_Action.Is_TIN, " & _
    " tbl_Personnel_Department.DepartmentName, tbl_Personnel_TaxStatus.TaxStatus AS TaxStatusName, " & _
    " tbl_Personnel_Position.PositionName, tbl_Personnel_Action.SSS, " & _
    " tbl_Personnel_Action.PHIC, tbl_Personnel_Action.PAGIBIG, " & _
    " tbl_Personnel_Action.TIN " & _
    " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
    " tbl_Personnel_TaxStatus ON tbl_Personnel_Action.TaxStatus = tbl_Personnel_TaxStatus.PK LEFT OUTER JOIN " & _
    " tbl_Personnel_Position ON tbl_Personnel_Action.Positions = tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
    " tbl_Personnel_Department ON tbl_Personnel_Action.Dept = tbl_Personnel_Department.PK " & _
    " WHERE (tbl_Personnel_Action.EmpPK = " & EmployeeNo & ") " & _
    " AND (tbl_Personnel_Action.EffectivityDate <= CONVERT(DATETIME, CONVERT(char(6), GETDATE(), 12), 102)) " & _
    " ORDER BY tbl_Personnel_Action.EffectivityDate DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    Division = rs!Division
    Department = rs!Dept
    TaxStatus = rs!TaxStatus
    Positions = rs!Positions
    CompRate = rs!CompensationRate
    Basic = rs!Basic
    Cola = rs!Cola
    Allowance = rs!Allowance
    WithSSS = rs!Is_SSS
    WithPagIbig = rs!Is_PAGIBIG
    WithPHIC = rs!Is_PHIC
    WithTIN = rs!Is_TIN
    SSSNum = rs!SSS
    PagIbigNum = rs!PAGIBIG
    PHICNum = rs!PHIC
    TINNum = rs!TIN

    txtName.Text = lstResult.List(lstResult.ListIndex)
    txtDivision.Text = IIf(IsNull(rs!Division), "", IIf(rs!Division = 1, "CLUB HOUSE", IIf(rs!Division = 2, "MAINTENANCE", "")))
    txtDept.Text = rs!DepartmentName
    txtTaxStatus.Text = rs!TaxStatusName
    txtPosition.Text = rs!PositionName
    txtCompRate.Text = IIf(rs!CompensationRate = 1, "MONTHLY", "DAILY")
End If
rs.Close
TRANSACTIONTYPE = is_ADDING
cmdCancel_Click
cmbEmploymentStatus.SetFocus
End Sub

Private Sub cmdOKSearch_Click()
If cmbEffectivityDate.ListIndex = -1 Then Exit Sub
s = "SELECT CntrlNo " & _
    " From tbl_Personnel_Action " & _
    " WHERE (EmpPK = " & lstResultSearch.ItemData(lstResultSearch.ListIndex) & ") " & _
    " AND (EffectivityDate = '" & FormatDateTime(cmbEffectivityDate.List(cmbEffectivityDate.ListIndex), vbShortDate) & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    BROWSER rs!CntrlNo
    TOOLBAR_FUNC 1
End If
If rs.State = adStateOpen Then rs.Close
cmdCancelSearch_Click
End Sub

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert:   PRESS_INSERT
    Case vbKey2:        PRESS_F2
    Case vbKeyDelete:   PRESS_DELETE
    Case vbKeyF5:       PRESS_F5
    Case vbKeyF6:       PRESS_F6
    Case vbKeyF9:       PRESS_F9
    Case vbKeyEscape:   PRESS_ESCAPE
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 2
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
cmbEmploymentStatus.Clear
s = "SELECT PK, StatusName " & _
    " From tbl_Personnel_EmploymentStatus " & _
    " Where (Active = 2) " & _
    " ORDER BY StatusName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbEmploymentStatus.AddItem rs!StatusName
    cmbEmploymentStatus.ItemData(cmbEmploymentStatus.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
CLEARTEXT
LOCKTEXT True
TOOLBAR_FUNC 0
TRANSACTIONTYPE = is_REFRESH
'Me.Caption = "Deactivation Memo"

tmp = SetWindowLong(txtSearchSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearchSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtRemarks.hwnd, GWL_STYLE, GetWindowLong(txtRemarks.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picAdd.Visible = True Then Cancel = -1
If picSearch.Visible = True Then Cancel = 1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub


Private Sub lstResult_KeyDown(KeyCode As Integer, Shift As Integer)
If lstResult.ListIndex = -1 Then Exit Sub
If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub lstResultSearch_Click()
If lstResultSearch.ListIndex = -1 Then cmbEffectivityDate.Clear: Exit Sub
cmbEffectivityDate.Clear
s = "SELECT tbl_Personnel_Action.EffectivityDate " & _
    " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
    " tbl_Personnel_EmploymentStatus ON tbl_Personnel_Action.EmpStatus = tbl_Personnel_EmploymentStatus.PK " & _
    " Where (tbl_Personnel_EmploymentStatus.Active = 2) " & _
    " And (tbl_Personnel_Action.EmpPK = " & lstResultSearch.ItemData(lstResultSearch.ListIndex) & ") " & _
    " ORDER BY tbl_Personnel_Action.EffectivityDate DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbEffectivityDate.AddItem Format(rs!EffectivityDate, "mm/dd/yyyy")
    rs.MoveNext
Wend
rs.Close
If cmbEffectivityDate.ListCount Then cmbEffectivityDate.ListIndex = 0
End Sub

Private Sub lstResultSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbEffectivityDate.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "Save":    PRESS_F5
    Case "Undo":    PRESS_ESCAPE
    Case "Find":    PRESS_F6
    Case "Print":   PRESS_F9
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtRemarks_GotFocus()
HTEXT txtRemarks
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtEffectDate.SetFocus
ElseIf KeyCode = vbKeyUp Then
    cmbEmploymentStatus.SetFocus
End If
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstResult.Clear: Exit Sub
lstResult.Clear
s = "SELECT tbl_Personnel_IDNumber.PK, " & _
    " tbl_Personnel_IDNumber.IDNumber, " & _
    " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName " & _
    " FROM tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
    " WHERE (tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
    " AND (ISNULL((SELECT TOP 1 tbl_Personnel_EmploymentStatus.Active " & _
    " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
    " tbl_Personnel_EmploymentStatus ON tbl_Personnel_Action.EmpStatus = tbl_Personnel_EmploymentStatus.PK " & _
    " WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_IDNumber.PK) " & _
    " AND (tbl_Personnel_Action.EffectivityDate <= CONVERT(DATETIME, CONVERT(char(6), getdate(), 12), 102)) ORDER BY tbl_Personnel_Action.EffectivityDate DESC), 0) = 1) " & _
    " ORDER BY tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResult.AddItem rs!IDNumber & " - " & rs!EmployeeName
    lstResult.ItemData(lstResult.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResult.ListCount Then lstResult.ListIndex = 0
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then lstResult.SetFocus
End Sub

Private Sub txtSearchSearch_Change()
If Trim(txtSearchSearch.Text) = "" Then lstResultSearch.Clear: cmbEffectivityDate.Clear: Exit Sub
lstResultSearch.Clear
s = "SELECT tbl_Personnel_IDNumber.PK, " & _
    " tbl_Personnel_IDNumber.IDNumber, " & _
    " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName " & _
    " FROM tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
    " WHERE (tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearchSearch.Text)) & "%') " & _
    " AND (ISNULL((SELECT TOP 1 tbl_Personnel_EmploymentStatus.Active " & _
    " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
    " tbl_Personnel_EmploymentStatus ON tbl_Personnel_Action.EmpStatus = tbl_Personnel_EmploymentStatus.PK " & _
    " WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_IDNumber.PK) " & _
    " AND (tbl_Personnel_Action.EffectivityDate <= CONVERT(DATETIME, CONVERT(char(6), getdate(), 12), 102)) ORDER BY tbl_Personnel_Action.EffectivityDate DESC), 0) = 2) " & _
    " ORDER BY tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultSearch.AddItem rs!IDNumber & " - " & rs!EmployeeName
    lstResultSearch.ItemData(lstResultSearch.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResultSearch.ListCount Then lstResultSearch.ListIndex = 0
End Sub

Private Sub txtSearchSearch_GotFocus()
HTEXT txtSearchSearch
End Sub

Private Sub txtSearchSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultSearch.SetFocus
End Sub
