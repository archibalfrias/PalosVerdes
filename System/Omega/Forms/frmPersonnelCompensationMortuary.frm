VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelCompensationMortuary 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPersonnelCompensationMortuary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   1320
      ScaleHeight     =   1455
      ScaleWidth      =   4455
      TabIndex        =   3
      Top             =   1080
      Width           =   4455
      Begin VB.TextBox txtNoMortuary 
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtPeriodTo 
         Height          =   315
         Left            =   3120
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtPeriodFrom 
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cmbDivision 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtCtrl 
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Mortuary"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Period"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   675
         Width           =   1575
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "Control Number"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1575
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7680
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
            Picture         =   "frmPersonnelCompensationMortuary.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationMortuary.frx":09CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationMortuary.frx":0B50
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationMortuary.frx":0E6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationMortuary.frx":1223
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationMortuary.frx":1675
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationMortuary.frx":1AC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationMortuary.frx":1E7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationMortuary.frx":1F91
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationMortuary.frx":24D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationMortuary.frx":262D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationMortuary.frx":2B6F
            Key             =   ""
         EndProperty
      EndProperty
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
            NumButtons      =   20
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
               Caption         =   "Close"
               Key             =   "Close"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      TabIndex        =   2
      Top             =   3000
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8643
            MinWidth        =   8643
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPersonnelCompensationMortuary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim iDivision, iPeriod, sCtrl

Private Sub BROWSER(Ctrl, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If Ctrl <> "" Then
            s = "SELECT TOP 1 tbl_Personnel_Compensation_Mortuary.PK, tbl_Personnel_Compensation_Mortuary.Ctrl, " & _
                " tbl_Personnel_Compensation_Mortuary.Division, tbl_Personnel_Compensation_Mortuary.Period, " & _
                " tbl_Personnel_Compensation_Period.DateFrom, tbl_Personnel_Compensation_Period.DateTo, " & _
                " tbl_Personnel_Compensation_Mortuary.NoOfMortuary, tbl_Personnel_Compensation_Mortuary.LastModified, " & _
                " tbl_Personnel_Compensation_Mortuary.Locked " & _
                " FROM tbl_Personnel_Compensation_Mortuary LEFT OUTER JOIN " & _
                " tbl_Personnel_Compensation_Period ON tbl_Personnel_Compensation_Mortuary.Period = tbl_Personnel_Compensation_Period.PK " & _
                " WHERE (tbl_Personnel_Compensation_Mortuary.Ctrl = '" & Ctrl & "') " & _
                " ORDER BY tbl_Personnel_Compensation_Mortuary.Ctrl"
        Else
            s = "SELECT TOP 1 tbl_Personnel_Compensation_Mortuary.PK, tbl_Personnel_Compensation_Mortuary.Ctrl, " & _
                " tbl_Personnel_Compensation_Mortuary.Division, tbl_Personnel_Compensation_Mortuary.Period, " & _
                " tbl_Personnel_Compensation_Period.DateFrom, tbl_Personnel_Compensation_Period.DateTo, " & _
                " tbl_Personnel_Compensation_Mortuary.NoOfMortuary, tbl_Personnel_Compensation_Mortuary.LastModified, " & _
                " tbl_Personnel_Compensation_Mortuary.Locked " & _
                " FROM tbl_Personnel_Compensation_Mortuary LEFT OUTER JOIN " & _
                " tbl_Personnel_Compensation_Period ON tbl_Personnel_Compensation_Mortuary.Period = tbl_Personnel_Compensation_Period.PK " & _
                " ORDER BY tbl_Personnel_Compensation_Mortuary.Ctrl"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Personnel_Compensation_Mortuary.PK, tbl_Personnel_Compensation_Mortuary.Ctrl, " & _
            " tbl_Personnel_Compensation_Mortuary.Division, tbl_Personnel_Compensation_Mortuary.Period, " & _
            " tbl_Personnel_Compensation_Period.DateFrom, tbl_Personnel_Compensation_Period.DateTo, " & _
            " tbl_Personnel_Compensation_Mortuary.NoOfMortuary, tbl_Personnel_Compensation_Mortuary.LastModified, " & _
            " tbl_Personnel_Compensation_Mortuary.Locked " & _
            " FROM tbl_Personnel_Compensation_Mortuary LEFT OUTER JOIN " & _
            " tbl_Personnel_Compensation_Period ON tbl_Personnel_Compensation_Mortuary.Period = tbl_Personnel_Compensation_Period.PK " & _
            " ORDER BY tbl_Personnel_Compensation_Mortuary.Ctrl"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Personnel_Compensation_Mortuary.PK, tbl_Personnel_Compensation_Mortuary.Ctrl, " & _
            " tbl_Personnel_Compensation_Mortuary.Division, tbl_Personnel_Compensation_Mortuary.Period, " & _
            " tbl_Personnel_Compensation_Period.DateFrom, tbl_Personnel_Compensation_Period.DateTo, " & _
            " tbl_Personnel_Compensation_Mortuary.NoOfMortuary, tbl_Personnel_Compensation_Mortuary.LastModified, " & _
            " tbl_Personnel_Compensation_Mortuary.Locked " & _
            " FROM tbl_Personnel_Compensation_Mortuary LEFT OUTER JOIN " & _
            " tbl_Personnel_Compensation_Period ON tbl_Personnel_Compensation_Mortuary.Period = tbl_Personnel_Compensation_Period.PK " & _
            " WHERE (tbl_Personnel_Compensation_Mortuary.Ctrl < '" & Ctrl & "') " & _
            " ORDER BY tbl_Personnel_Compensation_Mortuary.Ctrl DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Personnel_Compensation_Mortuary.PK, tbl_Personnel_Compensation_Mortuary.Ctrl, " & _
            " tbl_Personnel_Compensation_Mortuary.Division, tbl_Personnel_Compensation_Mortuary.Period, " & _
            " tbl_Personnel_Compensation_Period.DateFrom, tbl_Personnel_Compensation_Period.DateTo, " & _
            " tbl_Personnel_Compensation_Mortuary.NoOfMortuary, tbl_Personnel_Compensation_Mortuary.LastModified, " & _
            " tbl_Personnel_Compensation_Mortuary.Locked " & _
            " FROM tbl_Personnel_Compensation_Mortuary LEFT OUTER JOIN " & _
            " tbl_Personnel_Compensation_Period ON tbl_Personnel_Compensation_Mortuary.Period = tbl_Personnel_Compensation_Period.PK " & _
            " WHERE (tbl_Personnel_Compensation_Mortuary.Ctrl > '" & Ctrl & "') " & _
            " ORDER BY tbl_Personnel_Compensation_Mortuary.Ctrl"
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Personnel_Compensation_Mortuary.PK, tbl_Personnel_Compensation_Mortuary.Ctrl, " & _
            " tbl_Personnel_Compensation_Mortuary.Division, tbl_Personnel_Compensation_Mortuary.Period, " & _
            " tbl_Personnel_Compensation_Period.DateFrom, tbl_Personnel_Compensation_Period.DateTo, " & _
            " tbl_Personnel_Compensation_Mortuary.NoOfMortuary, tbl_Personnel_Compensation_Mortuary.LastModified, " & _
            " tbl_Personnel_Compensation_Mortuary.Locked " & _
            " FROM tbl_Personnel_Compensation_Mortuary LEFT OUTER JOIN " & _
            " tbl_Personnel_Compensation_Period ON tbl_Personnel_Compensation_Mortuary.Period = tbl_Personnel_Compensation_Period.PK " & _
            " ORDER BY tbl_Personnel_Compensation_Mortuary.Ctrl DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iDivision = rs!Division
    iPeriod = rs!Period
    txtCtrl.Text = rs!Ctrl
    txtPeriodFrom.Text = Format(rs!DateFrom, "mm/dd/yyyy")
    txtPeriodTo.Text = Format(rs!DateTo, "mm/dd/yyyy")
    txtNoMortuary.Text = rs!NoOfMortuary
    cmbDivision.ListIndex = rs!Division - 1
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    Statusbar1.Panels(3).Text = IIf(rs!Locked = 1, "LOCKED", "UNLOCKED")
    SaveSetting App.EXEName, "MortuaryCtrl", "MortuaryCtrl", rs!Ctrl
    
End If
rs.Close
End Sub


Private Sub PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If AccessRights("Mortuary", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If

CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
cmbDivision.SetFocus
End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("Mortuary", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
End Sub

Private Sub PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("Mortuary", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
On Error GoTo PG:
CLEARTEXT
BROWSER GetSetting(App.EXEName, "MortuaryCtrl", "MortuaryCtrl", ""), "is_PAGEDOWN"
If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "MortuaryCtrl", "MortuaryCtrl", ""), "is_HOME"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If cmbDivision.ListIndex = -1 Then MsgBox "Please Select Division!                          ", vbCritical, "Error...": cmbDivision.SetFocus: Exit Sub
If IsDate(txtPeriodFrom.Text) = False Then MsgBox "Please Supply a Valid Date!                            ", vbCritical, "Error...": txtPeriodFrom.SetFocus: Exit Sub
If IsDate(txtPeriodTo.Text) = False Then MsgBox "Please Supply a Valid Date!                                  ", vbCritical, "Error...": txtPeriodTo.SetFocus: Exit Sub
iPeriod = GET_PERIOD(FormatDateTime(txtPeriodFrom.Text, vbShortDate), FormatDateTime(txtPeriodTo.Text, vbShortDate), iDivision)
If iPeriod = 0 Then MsgBox "Invalid Payroll Period!                             ", vbCritical, "Error...": Exit Sub
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    sCtrl = ""
    s = "SELECT TOP 1 tbl_Personnel_Compensation_Mortuary.Ctrl " & _
        " FROM tbl_Personnel_Compensation_Mortuary LEFT OUTER JOIN " & _
        " tbl_Personnel_Compensation_Period ON tbl_Personnel_Compensation_Mortuary.Period = tbl_Personnel_Compensation_Period.PK " & _
        " Where (Year(tbl_Personnel_Compensation_Period.DateTo) = " & Format(FormatDateTime(txtPeriodTo.Text, vbShortDate), "yyyy") & ") " & _
        " ORDER BY tbl_Personnel_Compensation_Mortuary.Ctrl DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        sCtrl = Format(CDbl(rs!Ctrl) + 1, "0000000#")
    Else
        sCtrl = Format(FormatDateTime(txtPeriodTo.Text, vbShortDate), "yyyy") & "0000"
    End If
    rs.Close
    
    Do
        s = "SELECT tbl_Personnel_Compensation_Mortuary.* " & _
            " FROM tbl_Personnel_Compensation_Mortuary " & _
            " WHERE (Ctrl = '" & sCtrl & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If
        rs.Close
        sCtrl = Format(CDbl(sCtrl) + 1, "0000000#")
    Loop
    
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Compensation_Mortuary " & _
                      " (Ctrl, Division, Period, NoOfMortuary, LastModified) " & _
                      " VALUES ('" & sCtrl & "', " & iDivision & ", " & iPeriod & ", " & _
                      " " & RETURNTEXTVALUE(txtNoMortuary) & ", " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
    
End If
If TRANSACTIONTYPE = is_EDITTING Then
    sCtrl = Trim(txtCtrl.Text)
    ConnOmega.Execute "UPDATE tbl_Personnel_Compensation_Mortuary " & _
                      " SET Division = " & iDivision & ", " & _
                      " Period = " & iPeriod & ", " & _
                      " NoOfMortuary = " & RETURNTEXTVALUE(txtNoMortuary) & ", " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
End If
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER sCtrl, "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    Unload Me
Else
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "MortuaryCtrl", "MortuaryCtrl", ""), "is_LOAD"
    If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "MortuaryCtrl", "MortuaryCtrl", ""), "is_HOME"
End If
End Sub


Private Sub CLEARTEXT()
iDivision = 0
iPeriod = 0
txtCtrl.Text = ""
txtPeriodFrom.Text = ""
txtPeriodTo.Text = ""
txtNoMortuary.Text = ""
cmbDivision.ListIndex = -1
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
Statusbar1.Panels(3).Text = ""
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtCtrl.Locked = True
txtPeriodFrom.Locked = bln
txtPeriodTo.Locked = bln
txtNoMortuary.Locked = bln
cmbDivision.Locked = bln
End Sub

Private Sub TOOLBARFUNC(intSelect As Integer)
With Toolbar1
    Select Case intSelect
        Case 1      '=== REFRESH ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
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
            .Buttons(19).Enabled = True
            '.Buttons(21).Enabled = True
            .Buttons(1).ToolTipText = "NEW (Ins)"
            .Buttons(3).ToolTipText = "EDIT (F2)"
            .Buttons(5).ToolTipText = "DELETE (Del)"
            .Buttons(7).ToolTipText = "FIRST (Home)"
            .Buttons(9).ToolTipText = "BACK (PgUp)"
            .Buttons(11).ToolTipText = "NEXT (PgDown)"
            .Buttons(13).ToolTipText = "LAST (End)"
            .Buttons(15).ToolTipText = "FIND (F6)"
            .Buttons(17).ToolTipText = "PRINT (F9)"
            .Buttons(19).ToolTipText = "CLOSE (Esc)"
            '.Buttons(21).ToolTipText = "CLOSE (Esc)"
        Case 2      '=== ADD/EDIT ====
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
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
            .Buttons(19).Enabled = False
'            .Buttons(21).Enabled = False
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
'            .Buttons(21).ToolTipText = ""
        Case 3      '=== FIND ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(1).Enabled = False
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = False
            .Buttons(7).Image = 4
            .Buttons(7).Caption = "First"
            .Buttons(9).Image = 13
            .Buttons(9).Caption = "Undo"
            .Buttons(7).Enabled = False
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
            .Buttons(19).Enabled = False
'            .Buttons(21).Enabled = False
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
'            .Buttons(21).ToolTipText = ""
        Case 4      '=== EMPTY DETAIL ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(1).Enabled = True
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = False
            .Buttons(7).Image = 12
            .Buttons(7).Caption = "Save"
            .Buttons(9).Image = 13
            .Buttons(9).Caption = "Undo"
            .Buttons(7).Enabled = True
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
            .Buttons(19).Enabled = False
'            .Buttons(21).Enabled = False
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
'            .Buttons(21).ToolTipText = ""
        Case 5      '=== NOT EMPTY DETAIL ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(1).Enabled = True
            .Buttons(3).Enabled = True
            .Buttons(5).Enabled = True
            .Buttons(7).Image = 12
            .Buttons(7).Caption = "Save"
            .Buttons(9).Image = 13
            .Buttons(9).Caption = "Undo"
            .Buttons(7).Enabled = True
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
            .Buttons(19).Enabled = False
'            .Buttons(21).Enabled = False
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
'            .Buttons(21).ToolTipText = ""
    End Select
End With
End Sub

Private Sub cmbDivision_Click()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    iDivision = cmbDivision.ListIndex + 1
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
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "MortuaryCtrl", "MortuaryCtrl", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "MortuaryCtrl", "MortuaryCtrl", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "MortuaryCtrl", "MortuaryCtrl", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "MortuaryCtrl", "MortuaryCtrl", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.Height - Me.Height) / 5
Me.Left = (MainForm.Width - Me.Width) / 5
With cmbDivision
    .Clear
    .AddItem "CLUB HOUSE"
    .AddItem "MAINTENANCE"
End With
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "MortuaryCtrl", "MortuaryCtrl", ""), "is_LOAD"
If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "MortuaryCtrl", "MortuaryCtrl", ""), "is_HOME"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "MortuaryCtrl", "MortuaryCtrl", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "MortuaryCtrl", "MortuaryCtrl", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "MortuaryCtrl", "MortuaryCtrl", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "MortuaryCtrl", "MortuaryCtrl", ""), "is_END"
    Case "Find":
    Case "Print":
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtNoMortuary_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtPeriodFrom_LostFocus()
If IsDate(txtPeriodFrom.Text) = True Then
    txtPeriodFrom.Text = Format(FormatDateTime(txtPeriodFrom.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub

Private Sub txtPeriodTo_LostFocus()
If IsDate(txtPeriodTo.Text) = True Then
    txtPeriodTo.Text = Format(FormatDateTime(txtPeriodTo.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub
