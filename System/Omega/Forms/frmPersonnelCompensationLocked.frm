VERSION 5.00
Begin VB.Form frmPersonnelCompensationLocked 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5220
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPersonnelCompensationLocked.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   240
      ScaleHeight     =   1695
      ScaleWidth      =   4455
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.OptionButton optUnlocked 
         BackColor       =   &H00C6B8A4&
         Caption         =   "UNLOCKED"
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   0
         Width           =   1215
      End
      Begin VB.OptionButton optLocked 
         BackColor       =   &H00C6B8A4&
         Caption         =   "LOCKED"
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   0
         Width           =   1215
      End
      Begin VB.ComboBox cmbDivision 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtPeriodFrom 
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtPeriodTo 
         Height          =   315
         Left            =   3120
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdOKAdd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   600
         Picture         =   "frmPersonnelCompensationLocked.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelAdd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2280
         Picture         =   "frmPersonnelCompensationLocked.frx":0AB4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1200
         Width           =   1560
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Period"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   675
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.PictureBox picProgressBar 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   9435
      TabIndex        =   8
      Top             =   120
      Width           =   9495
   End
End
Attribute VB_Name = "frmPersonnelCompensationLocked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iPeriod, i, iLocked

Private Sub cmdCancelAdd_Click()
Unload Me
End Sub

Private Sub cmdOKAdd_Click()
If cmbDivision.ListIndex = -1 Then MsgBox "Please Select Division!                       ", vbCritical, "Error...": cmbDivision.SetFocus: Exit Sub
If IsDate(txtPeriodFrom.Text) = False Then MsgBox "Please Supply a Valid Date!                    ", vbCritical, "Error...": txtPeriodFrom.SetFocus: Exit Sub
If IsDate(txtPeriodTo.Text) = False Then MsgBox "Please Supply a Valid Date!                      ", vbCritical, "Error...": txtPeriodTo.SetFocus: Exit Sub
iPeriod = GET_PERIOD(FormatDateTime(txtPeriodFrom.Text, vbShortDate), FormatDateTime(txtPeriodTo.Text, vbShortDate), cmbDivision.ListIndex + 1)
If CDbl(iPeriod) = 0 Then MsgBox "Invalid Cut-Off!                      ", vbCritical, "Error...": txtPeriodFrom.SetFocus: Exit Sub
picProgressBar.BackColor = &HFFFFFF
picMain.Visible = False
picProgressBar.Visible = True
Me.Height = 1395
Me.Width = 9840
Me.Top = (MainForm.Height - Me.Height) / 5
Me.Left = (MainForm.Width - Me.Width) / 5
DoEvents
iLocked = IIf(optLocked.Value = True, 1, IIf(optUnlocked.Value = True, 2, 0))
Me.Caption = IIf(iLocked = 1, "Locking", "Unlocking") & " Payroll . . . . "
i = 0
s = "SELECT tbl_Personnel_Compensation.PK, tbl_Personnel_Compensation.ActionMemo " & _
    " FROM tbl_Personnel_Compensation LEFT OUTER JOIN " & _
    " tbl_Personnel_Compensation_Period ON tbl_Personnel_Compensation.Period = tbl_Personnel_Compensation_Period.PK " & _
    " WHERE (tbl_Personnel_Compensation.Division = " & cmbDivision.ListIndex + 1 & ") " & _
    " AND (tbl_Personnel_Compensation_Period.DateFrom >= '" & FormatDateTime(txtPeriodFrom.Text, vbShortDate) & "') " & _
    " AND (tbl_Personnel_Compensation_Period.DateTo <= '" & FormatDateTime(txtPeriodTo.Text, vbShortDate) & "')"
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
While Not ra.EOF
    DoEvents
    i = i + 1
    ConnOmega.Execute "UPDATE tbl_Personnel_Compensation " & _
                      " SET Locked = " & IIf(iLocked = 1, 1, 0) & " " & _
                      " WHERE (PK = " & ra!PK & ")"
                      
    ConnOmega.Execute "UPDATE tbl_Personnel_Action  " & _
                      " SET Locked = 1 " & _
                      " WHERE (PK = " & ra!ActionMemo & ")"
                      
    UpdateProgress picProgressBar, i / ra.RecordCount
    ra.MoveNext
Wend
ra.Close

s = "SELECT tbl_Personnel_Compensation_Mortuary.PK " & _
    " FROM tbl_Personnel_Compensation_Mortuary LEFT OUTER JOIN " & _
    " tbl_Personnel_Compensation_Period ON tbl_Personnel_Compensation_Mortuary.Period = tbl_Personnel_Compensation_Period.PK " & _
    " WHERE (tbl_Personnel_Compensation_Mortuary.Division = " & cmbDivision.ListIndex + 1 & ") " & _
    " AND (tbl_Personnel_Compensation_Period.DateFrom >= '" & FormatDateTime(txtPeriodFrom.Text, vbShortDate) & "') " & _
    " AND (tbl_Personnel_Compensation_Period.DateTo <= '" & FormatDateTime(txtPeriodTo.Text, vbShortDate) & "') "
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
While Not ra.EOF
    ConnOmega.Execute "UPDATE tbl_Personnel_Compensation_Mortuary " & _
                      " SET Locked = " & IIf(iLocked = 1, 1, 0) & " " & _
                      " WHERE (PK = " & ra!PK & ")"
    rs.MoveNext
Wend
ra.Close

MsgBox "Successfully " & IIf(iLocked = 1, "Locked!", "Unlocked!") & "                            ", vbCritical, "Info"

picProgressBar.Visible = False
Me.Width = 5070
Me.Height = 2445 '2070
Me.Caption = "Locked Payroll"
Me.Top = (MainForm.Height - Me.Height) / 5
Me.Left = (MainForm.Width - Me.Width) / 5
picMain.Visible = True
End Sub

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Width = 5070
Me.Height = 2445 '2070
Me.Top = (MainForm.Height - Me.Height) / 5
Me.Left = (MainForm.Width - Me.Width) / 5
optLocked.Value = True
With cmbDivision
    .Clear
    .AddItem "CLUB HOUSE"
    .AddItem "MAINTENANCE"
End With
picMain.ZOrder 0
picMain.Visible = True
picProgressBar.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picProgressBar.Visible = True Then Cancel = -1
End Sub

Private Sub txtPeriodFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtPeriodTo.SetFocus
End Sub

Private Sub txtPeriodFrom_LostFocus()
If IsDate(txtPeriodFrom.Text) = True Then
    txtPeriodFrom.Text = Format(FormatDateTime(txtPeriodFrom.Text, vbShortDate), "mm/dd/yyyy")
    t = "SELECT Type, DateTo " & _
        " From tbl_Personnel_Compensation_Period " & _
        " WHERE (Type = " & cmbDivision.ListIndex + 1 & ") " & _
        " AND (DateFrom = '" & FormatDateTime(txtPeriodFrom.Text, vbShortDate) & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtPeriodTo.Text = Format(rt!DateTo, "mm/dd/yyyy")
    End If
    rt.Close
End If
End Sub

Private Sub txtPeriodTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAdd_Click
End Sub

Private Sub txtPeriodTo_LostFocus()
If IsDate(txtPeriodTo.Text) = True Then
    txtPeriodTo.Text = Format(FormatDateTime(txtPeriodTo.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub

