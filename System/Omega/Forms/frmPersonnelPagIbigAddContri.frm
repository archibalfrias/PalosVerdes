VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelPagIbigAddContri 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9825
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   1080
      ScaleHeight     =   1455
      ScaleWidth      =   7455
      TabIndex        =   3
      Top             =   1440
      Width           =   7455
      Begin VB.TextBox txtCtrl 
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtRemarks 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1080
         Width           =   5895
      End
      Begin VB.TextBox txtAddAmount 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtEffectDate 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "12/15/2017"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   0
         Width           =   5895
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Additional Amount"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Effect"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1215
      End
   End
   Begin RPVGCC.b8Container picAdd 
      Height          =   3135
      Left            =   2640
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5530
      BackColor       =   15396057
      Begin VB.ListBox lstResultAdd 
         Height          =   1620
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox txtSearchAdd 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   3735
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
         Left            =   2040
         Picture         =   "frmPersonnelPagIbigAddContri.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2520
         Width           =   1560
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
         Left            =   360
         Picture         =   "frmPersonnelPagIbigAddContri.frx":075C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2520
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   18
         Top             =   45
         Width           =   3885
         _ExtentX        =   6853
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
         Icon            =   "frmPersonnelPagIbigAddContri.frx":0DCE
         ShadowVisible   =   0   'False
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   0
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   1
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
               ImageKey        =   "IMG1"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Edit"
               Key             =   "Edit"
               ImageKey        =   "IMG2"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Delete"
               Key             =   "Delete"
               ImageKey        =   "IMG3"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "First"
               Key             =   "First"
               ImageKey        =   "IMG4"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Back"
               Key             =   "Back"
               ImageKey        =   "IMG5"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Next"
               Key             =   "Next"
               ImageKey        =   "IMG6"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Last"
               Key             =   "Last"
               ImageKey        =   "IMG7"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Find"
               Key             =   "Find"
               ImageKey        =   "IMG8"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Print"
               Key             =   "Print"
               ImageKey        =   "IMG9"
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               Key             =   "Refresh"
               ImageKey        =   "IMG12"
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
               ImageKey        =   "IMG13"
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         MousePointer    =   99
         MouseIcon       =   "frmPersonnelPagIbigAddContri.frx":1368
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9000
      Top             =   1320
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
            Picture         =   "frmPersonnelPagIbigAddContri.frx":1682
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPagIbigAddContri.frx":235C
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPagIbigAddContri.frx":3036
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPagIbigAddContri.frx":3D10
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPagIbigAddContri.frx":49EA
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPagIbigAddContri.frx":56C4
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPagIbigAddContri.frx":639E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPagIbigAddContri.frx":7078
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPagIbigAddContri.frx":7D52
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPagIbigAddContri.frx":862C
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPagIbigAddContri.frx":9306
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPagIbigAddContri.frx":9FE0
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPagIbigAddContri.frx":ACBA
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPagIbigAddContri.frx":B994
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPagIbigAddContri.frx":C66E
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   3495
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
            MinWidth        =   2469
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
End
Attribute VB_Name = "frmPersonnelPagIbigAddContri"
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

Dim sCtrl, iEmployeeKey

Private Sub BROWSER(Ctrl, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If Ctrl <> "" Then
            s = "SELECT TOP (1) tbl_Personnel_PagIbig_AddAmount.* " & _
                " FROM tbl_Personnel_PagIbig_AddAmount " & _
                " WHERE (Ctrl = '" & Ctrl & "') " & _
                " ORDER BY Ctrl"
        Else
            s = "SELECT TOP (1) tbl_Personnel_PagIbig_AddAmount.* " & _
                " FROM tbl_Personnel_PagIbig_AddAmount " & _
                " ORDER BY Ctrl"
        End If
    Case "is_HOME"
        If picAdd.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) tbl_Personnel_PagIbig_AddAmount.* " & _
            " FROM tbl_Personnel_PagIbig_AddAmount " & _
            " ORDER BY Ctrl"
    Case "is_PAGEUP"
        If picAdd.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) tbl_Personnel_PagIbig_AddAmount.* " & _
            " FROM tbl_Personnel_PagIbig_AddAmount " & _
            " WHERE (Ctrl < '" & Ctrl & "') " & _
            " ORDER BY Ctrl DESC"
    Case "is_PAGEDOWN"
        If picAdd.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) tbl_Personnel_PagIbig_AddAmount.* " & _
            " FROM tbl_Personnel_PagIbig_AddAmount " & _
            " WHERE (Ctrl > '" & Ctrl & "') " & _
            " ORDER BY Ctrl"
    Case "is_END"
        If picAdd.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) tbl_Personnel_PagIbig_AddAmount.* " & _
            " FROM tbl_Personnel_PagIbig_AddAmount " & _
            " ORDER BY Ctrl DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iEmployeeKey = rs!EmployeeKey
    txtCtrl.Text = rs!Ctrl
    txtName.Text = ""
    t = "SELECT dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName " & _
        " FROM  dbo.tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
        " WHERE (dbo.tbl_Personnel_IDNumber.PK = " & rs!EmployeeKey & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtName.Text = rt!IDNumber & " - " & rt!LastName & ",  " & rt!FirstName & "  " & rt!MiddleName
    End If
    rt.Close
    txtEffectDate.Text = Format(rs!Effectdate, "mm/dd/yyyy")
    txtAddAmount.Text = Format(rs!AddAmount, "#,##0.00")
    txtRemarks.Text = rs!Remarks
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    
    SaveSetting App.EXEName, "PersonnelPagIbigAdd", "PersonnelPagIbigAdd", rs!Ctrl
    
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If picAdd.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If AccessRights("PagIbig Additional Contribution", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
picAdd.ZOrder 0
picMain.Enabled = False
picToolbar.Enabled = False
txtSearchAdd.Text = ""
picAdd.Visible = True
txtSearchAdd.SetFocus
'CLEARTEXT
'LOCKTEXT False
'TOOLBARFUNC 2
'TRANSACTIONTYPE = is_ADDING
End Sub

Private Sub PRESS_F2()
If picAdd.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("PagIbig Additional Contribution", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_EDITTING
End Sub

Private Sub PRESS_DELETE()
If picAdd.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("PagIbig Additional Contribution", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If MsgBox("ARE YOU SURE IN DELETING THIS TRANSACTION?                   ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_Personnel_PagIbig_AddAmount WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "PersonnelPagIbigAdd", "PersonnelPagIbigAdd", ""), "is_PAGEDOWN"
If Trim(txtName.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelPagIbigAdd", "PersonnelPagIbigAdd", ""), "is_HOME"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If picAdd.Visible = True Then Exit Sub
If iEmployeeKey = 0 Then MsgBox "Please supply employee!                    ", vbCritical, "Error...": Exit Sub
If IsDate(txtEffectDate.Text) = False Then MsgBox "Please supply a valid date!                ", vbCritical, "Error...": txtEffectDate.SetFocus: Exit Sub
If RETURNTEXTVALUE(txtAddAmount) < 0 Then MsgBox "Invalid amount!                    ", vbCritical, "Error...": txtAddAmount.SetFocus: Exit Sub
If Trim(txtRemarks.Text) = "" Then MsgBox "Please supply remarks!                         ", vbCritical, "Error...": txtRemarks.SetFocus: Exit Sub
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    sCtrl = ""
    s = "SELECT TOP (1) tbl_Personnel_PagIbig_AddAmount.* " & _
        " FROM tbl_Personnel_PagIbig_AddAmount " & _
        " WHERE (Year(EffectDate) = " & Format(FormatDateTime(txtEffectDate.Text, vbShortDate), "yyyy") & ") " & _
        " ORDER BY Ctrl DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        sCtrl = Format(CDbl(rs!Ctrl) + 1, "0000000#")
    Else
        sCtrl = Format(FormatDateTime(txtEffectDate.Text, vbShortDate), "yyyy") & "0000"
    End If
    rs.Close
    Do
        s = "SELECT tbl_Personnel_PagIbig_AddAmount.* " & _
            " FROM tbl_Personnel_PagIbig_AddAmount " & _
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
    ConnOmega.Execute "INSERT INTO tbl_Personnel_PagIbig_AddAmount " & _
                      " (Ctrl, EmployeeKey, EffectDate, AddAmount, Remarks, LastModified) " & _
                      " VALUES ('" & sCtrl & "', " & iEmployeeKey & ", " & _
                      " '" & FormatDateTime(txtEffectDate.Text, vbShortDate) & "', " & _
                      " " & RETURNTEXTVALUE(txtAddAmount) & ", '" & FORMATSQL(Trim(txtRemarks.Text)) & "', " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
End If
If TRANSACTIONTYPE = is_EDITTING Then
    sCtrl = txtCtrl.Text
    ConnOmega.Execute "UPDATE tbl_Personnel_PagIbig_AddAmount " & _
                      " SET EffectDate = '" & FormatDateTime(txtEffectDate.Text, vbShortDate) & "', " & _
                      " AddAmount = " & RETURNTEXTVALUE(txtAddAmount) & ", " & _
                      " Remarks = '" & FORMATSQL(Trim(txtRemarks.Text)) & "', " & _
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

Private Sub PRESS_F6()
If picAdd.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picAdd.Visible = True Then cmdCancelAdd_Click: Exit Sub
    Unload Me
Else
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "PersonnelPagIbigAdd", "PersonnelPagIbigAdd", ""), "is_LOAD"
    If Trim(txtName.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelPagIbigAdd", "PersonnelPagIbigAdd", ""), "is_HOME"
End If
End Sub

Private Sub CLEARTEXT()
iEmployeeKey = 0
txtCtrl.Text = ""
txtName.Text = ""
txtEffectDate.Text = ""
txtAddAmount.Text = ""
txtRemarks.Text = ""
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtName.Locked = True
txtEffectDate.Locked = bln
txtAddAmount.Locked = bln
txtRemarks.Locked = bln
End Sub


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

Private Sub b8TitleBar2_CLoseClick()
cmdCancelAdd_Click
End Sub

Private Sub cmdCancelAdd_Click()
picAdd.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True
End Sub

Private Sub cmdOKAdd_Click()
If lstResultAdd.ListIndex = -1 Then MsgBox "Please select employee!                  ", vbCritical, "Error...": Exit Sub
CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
txtName.Text = lstResultAdd.List(lstResultAdd.ListIndex)
iEmployeeKey = lstResultAdd.ItemData(lstResultAdd.ListIndex)
txtEffectDate.Text = Format(Date, "mm/dd/yyyy")
TRANSACTIONTYPE = is_ADDING
cmdCancelAdd_Click
txtEffectDate.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert:   PRESS_INSERT
    Case vbKeyF2:       PRESS_F2
    Case vbKeyDelete:   PRESS_DELETE
    Case vbKeyF5:       PRESS_F5
    Case vbKeyF6:       PRESS_F6
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "PersonnelPagIbigAdd", "PersonnelPagIbigAdd", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "PersonnelPagIbigAdd", "PersonnelPagIbigAdd", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "PersonnelPagIbigAdd", "PersonnelPagIbigAdd", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "PersonnelPagIbigAdd", "PersonnelPagIbigAdd", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 3
Me.Left = (MainForm.ScaleWidth - Me.Width) / 3
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "PersonnelPagIbigAdd", "PersonnelPagIbigAdd", ""), "is_LOAD"
If Trim(txtName.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelPagIbigAdd", "PersonnelPagIbigAdd", ""), "is_HOME"
tmp = SetWindowLong(txtSearchAdd.hwnd, GWL_STYLE, GetWindowLong(txtSearchAdd.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picAdd.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "PersonnelPagIbigAdd", "PersonnelPagIbigAdd", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "PersonnelPagIbigAdd", "PersonnelPagIbigAdd", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "PersonnelPagIbigAdd", "PersonnelPagIbigAdd", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "PersonnelPagIbigAdd", "PersonnelPagIbigAdd", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtSearchAdd_Change()
If Trim(txtSearchAdd.Text) = "" Then lstResultAdd.Clear: Exit Sub
lstResultAdd.Clear
t = "SELECT dbo.tbl_Personnel_IDNumber.PK, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, " & _
    " dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName " & _
    " FROM  dbo.tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
    " WHERE (ISNULL((SELECT TOP (1) tbl_Personnel_EmploymentStatus_1.Active " & _
    " FROM  dbo.tbl_Personnel_ActionNew AS tbl_Personnel_ActionNew_1 LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_EmploymentStatus AS tbl_Personnel_EmploymentStatus_1 ON tbl_Personnel_ActionNew_1.EmpStatusKey = tbl_Personnel_EmploymentStatus_1.PK " & _
    " WHERE (tbl_Personnel_ActionNew_1.EmpPK = dbo.tbl_Personnel_IDNumber.PK) " & _
    " AND (tbl_Personnel_ActionNew_1.EffectivityDate <= '" & FormatDateTime(Date, vbShortDate) & "') " & _
    " ORDER BY tbl_Personnel_ActionNew_1.EffectivityDate DESC), 0) = 1) " & _
    " AND (dbo.tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%') " & _
    " ORDER BY dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName, dbo.tbl_Personnel_IDNumber.IDNumber"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
While Not rt.EOF
    lstResultAdd.AddItem rt!IDNumber & " - " & rt!LastName & ",  " & rt!FirstName & "  " & rt!MiddleName
    lstResultAdd.ItemData(lstResultAdd.NewIndex) = rt!PK
    rt.MoveNext
Wend
rt.Close
If lstResultAdd.ListCount Then lstResultAdd.ListIndex = 0
End Sub

Private Sub txtSearchAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultAdd.SetFocus
End Sub
