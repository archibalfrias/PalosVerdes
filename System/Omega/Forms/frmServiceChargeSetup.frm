VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServiceChargeSetup 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServiceChargeSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   16
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   17
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
            NumButtons      =   24
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
               Caption         =   " Post   "
               Key             =   "Post"
               ImageKey        =   "IMG10"
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               Key             =   "Refresh"
               ImageKey        =   "IMG12"
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
               ImageKey        =   "IMG13"
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         MousePointer    =   99
         MouseIcon       =   "frmServiceChargeSetup.frx":C472
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   9900
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   18
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmServiceChargeSetup.frx":C78C
               Top             =   120
               Visible         =   0   'False
               Width           =   1395
            End
         End
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
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   4275
      Width           =   9870
      _ExtentX        =   17410
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
   Begin RPVGCC.b8Container picSLine 
      Height          =   855
      Left            =   2640
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox txtRate1 
         Height          =   315
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   15
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtForCompany1 
         Height          =   315
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtForCompany 
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   10
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rate (%)"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3120
         TabIndex        =   13
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "For Company"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   2640
      ScaleHeight     =   2655
      ScaleWidth      =   4095
      TabIndex        =   4
      Top             =   1200
      Width           =   4095
      Begin MSComctlLib.ListView lstDetail 
         Height          =   1455
         Left            =   0
         TabIndex        =   8
         Top             =   1200
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   2566
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "For Company"
            Object.Width           =   4798
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Rate (%)"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.TextBox txtRankNFile 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtSupervisory 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtEffectDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   0
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Rank In File"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Supervisory"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Effectivity Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   0
         Width           =   1575
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9360
      Top             =   1200
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
            Picture         =   "frmServiceChargeSetup.frx":CE9F
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSetup.frx":DB79
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSetup.frx":E853
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSetup.frx":F52D
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSetup.frx":10207
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSetup.frx":10EE1
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSetup.frx":11BBB
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSetup.frx":12895
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSetup.frx":1356F
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSetup.frx":13E49
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSetup.frx":14B23
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSetup.frx":157FD
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSetup.frx":164D7
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSetup.frx":171B1
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSetup.frx":17E8B
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmServiceChargeSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim TRANS_DETAIL As Long
Const is_DET_REFRESH = 0
Const is_DET_ADDING = 1
Const is_DET_EDITTING = 2


Dim ListFocus   As Long
Dim ROW         As Long
Dim tmp         As Long

Dim x, i, RowLine, MasterKey, strEffectDate

Private Function BROWSER(dtmEffect, isAction As String)
'If IsDate(dtmEffect) = False Then Exit Function
Select Case isAction
    Case "is_LOAD"
        'If dtmEffect <> "" Then
        If IsDate(dtmEffect) = True Then
            s = "SELECT TOP 1 tbl_Service_Charge_Setup.* " & _
                " FROM tbl_Service_Charge_Setup " & _
                " WHERE (EffectDate = '" & FormatDateTime(dtmEffect, vbShortDate) & "') " & _
                " ORDER BY EffectDate"
        Else
            s = "SELECT TOP 1 tbl_Service_Charge_Setup.* " & _
                " FROM tbl_Service_Charge_Setup " & _
                " ORDER BY EffectDate"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Service_Charge_Setup.* " & _
            " FROM tbl_Service_Charge_Setup " & _
            " ORDER BY EffectDate"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Service_Charge_Setup.* " & _
            " FROM tbl_Service_Charge_Setup " & _
            " WHERE (EffectDate < '" & FormatDateTime(dtmEffect, vbShortDate) & "') " & _
            " ORDER BY EffectDate DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Service_Charge_Setup.* " & _
            " FROM tbl_Service_Charge_Setup " & _
            " WHERE (EffectDate > '" & FormatDateTime(dtmEffect, vbShortDate) & "') " & _
            " ORDER BY EffectDate "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Service_Charge_Setup.* " & _
            " FROM tbl_Service_Charge_Setup " & _
            " ORDER BY EffectDate DESC"
    Case Else: Exit Function
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then

    txtEffectDate.Text = Format(rs!Effectdate, "mm/dd/yyyy")
    txtSupervisory.Text = Format(rs!Supervisory, "#,##0.00")
    txtRankNFile.Text = Format(rs!RankInFile, "#,##0.00")
    StatusBar1.Panels(1).Text = rs!PK
    StatusBar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    
    lstDetail.ListItems.Clear
    Set x = lstDetail.ListItems.Add()
    x.Text = ""
    x.SubItems(1) = " "
    x.SubItems(2) = " "
    
    t = "SELECT tbl_Service_Charge_SetupDetail.* " & _
        " FROM tbl_Service_Charge_SetupDetail " & _
        " WHERE (MasterKey = " & rs!PK & ") " & _
        " ORDER BY Line"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstDetail.ListItems.Clear
        While Not rt.EOF
            Set x = lstDetail.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rt!ForCompany
            x.SubItems(2) = Format(rt!Rate, "#,##0.00")
            rt.MoveNext
        Wend
    End If
    rt.Close
    
    SaveSetting App.EXEName, "ServiceSetupEff", "SrvcSetupEff", Format(rs!Effectdate, "mm/dd/yyyy")
    
End If
rs.Close
End Function

Private Function PRESS_INSERT()
If TRANSACTIONTYPE = is_REFRESH Then
    If AccessRights("Service Charge Setup", "Add") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Function
    End If
    CLEARTEXT
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_ADDING
    'Me.Caption = "Service Charge - New"
    txtEffectDate.SetFocus
Else
    If ListFocus = 0 Then Exit Function
    If picSLine.Visible = True Then Exit Function
    With lstDetail.ListItems
        If Trim(.Item(.Count).SubItems(1)) = "" Then
            ROW = .Count
            txtForCompany.Text = ""
            txtRate.Text = ""
            picToolbar.Enabled = False
            picMain.Enabled = False
            picSLine.ZOrder 0
            picSLine.Visible = True
            TRANS_DETAIL = is_DET_ADDING
            txtForCompany.SetFocus
        Else
            Set x = .Add()
            x.Text = ""
            x.SubItems(1) = " "
            x.SubItems(2) = " "
            ROW = .Count
            txtForCompany.Text = ""
            txtRate.Text = ""
            picToolbar.Enabled = False
            picMain.Enabled = False
            picSLine.ZOrder 0
            picSLine.Visible = True
            TRANS_DETAIL = is_DET_ADDING
            txtForCompany.SetFocus
        End If
    End With
End If
End Function

Private Function PRESS_F2()
If TRANSACTIONTYPE = is_REFRESH Then
    If StatusBar1.Panels(1).Text = "" Then Exit Function
    If AccessRights("Service Charge Setup", "Edit") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Function
    End If
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_EDITTING
    'Me.Caption = "Service Charge - Edit"
Else
    If ListFocus = 0 Then Exit Function
    If picSLine.Visible = True Then Exit Function
    With lstDetail.ListItems
        txtForCompany.Text = .Item(ROW).SubItems(1)
        txtRate.Text = .Item(ROW).SubItems(2)
        txtForCompany1.Text = .Item(ROW).SubItems(1)
        txtRate1.Text = .Item(ROW).SubItems(2)
        picToolbar.Enabled = False
        picMain.Enabled = False
        picSLine.ZOrder 0
        picSLine.Visible = True
        TRANS_DETAIL = is_DET_EDITTING
        txtForCompany.SetFocus
    End With
End If
End Function

Private Function PRESS_DELETE()
If TRANSACTIONTYPE = is_REFRESH Then
    If StatusBar1.Panels(1).Text = "" Then Exit Function
    If AccessRights("Service Charge Setup", "Delete") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Function
    End If
    If MsgBox("ARE YOU SURE IN DELETING THIS RECORD?                        ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Function
    On Error GoTo PG:
    ConnOmega.Execute "DELETE FROM tbl_Service_Charge_Setup WHERE (PK = " & StatusBar1.Panels(1).Text & ")"
    CLEARTEXT
    BROWSER GetSetting(App.EXEName, "ServiceSetupEff", "SrvcSetupEff", ""), "is_PAGEDOWN"
    If Trim(txtEffectDate.Text) = "" Then BROWSER GetSetting(App.EXEName, "ServiceSetupEff", "SrvcSetupEff", ""), "is_HOME"
Else
    If ListFocus = 0 Then Exit Function
    If picSLine.Visible = True Then Exit Function
    With lstDetail.ListItems
        If .Count > 1 Then
            .Remove ROW
            If CDbl(ROW) > .Count Then
                ROW = .Count
            End If
        Else
            .Item(.Count).SubItems(1) = " "
            .Item(.Count).SubItems(2) = " "
            ROW = .Count
        End If
        lstDetail.ListItems(ROW).EnsureVisible
        lstDetail.ListItems(ROW).Selected = True
        lstDetail.SetFocus
    End With
End If
Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Function
End Function

Private Function PRESS_F5()
If IsDate(txtEffectDate.Text) = False Then MsgBox "Please Supply a Valid Date!                ", vbCritical, "Error...": txtEffectDate.SetFocus: Exit Function
txtEffectDate.Text = Format(FormatDateTime(txtEffectDate.Text, vbShortDate), "mm/dd/yyyy")
strEffectDate = txtEffectDate.Text
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then

    ConnOmega.Execute "INSERT INTO tbl_Service_Charge_Setup " & _
                      " (EffectDate, Supervisory, RankInFile, LastModified) " & _
                      " VALUES ('" & FormatDateTime(txtEffectDate.Text, vbShortDate) & "', " & _
                      " " & RETURNTEXTVALUE(txtSupervisory) & ", " & _
                      " " & RETURNTEXTVALUE(txtRankNFile) & ", " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
    MasterKey = 0
    s = "SELECT PK " & _
        " FROM tbl_Service_Charge_Setup " & _
        " WHERE (EffectDate = '" & FormatDateTime(txtEffectDate.Text, vbShortDate) & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        MasterKey = rs!PK
    End If
    rs.Close
    
    If CDbl(MasterKey) > 0 Then
        With lstDetail.ListItems
            RowLine = 0
            For i = 1 To .Count
                If Trim(.Item(i).SubItems(1)) <> "" Then
                    RowLine = RowLine + 1
                    ConnOmega.Execute "INSERT INTO tbl_Service_Charge_SetupDetail " & _
                                      " (MasterKey, Line, ForCompany, Rate) " & _
                                      " VALUES (" & MasterKey & ", " & RowLine & ", " & _
                                      " '" & Trim(.Item(i).SubItems(1)) & "', " & _
                                      " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2))) & ")"
                End If
            Next i
        End With
    End If
    
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    'Me.Caption = "Service Charge - Browse"
    txtEffectDate.SetFocus
    BROWSER strEffectDate, "is_LOAD"
    
End If
If TRANSACTIONTYPE = is_EDITTING Then

    MasterKey = StatusBar1.Panels(1).Text
    
    ConnOmega.Execute "UPDATE tbl_Service_Charge_Setup " & _
                      " SET EffectDate = '" & FormatDateTime(txtEffectDate.Text, vbShortDate) & "', " & _
                      " Supervisory = " & RETURNTEXTVALUE(txtSupervisory) & ", " & _
                      " RankInFile = " & RETURNTEXTVALUE(txtRankNFile) & ", " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & MasterKey & ")"
    
    If CDbl(MasterKey) > 0 Then
        ConnOmega.Execute "DELETE FROM tbl_Service_Charge_SetupDetail WHERE (MasterKey = " & MasterKey & ")"
        With lstDetail.ListItems
            RowLine = 0
            For i = 1 To .Count
                If Trim(.Item(i).SubItems(1)) <> "" Then
                    RowLine = RowLine + 1
                    ConnOmega.Execute "INSERT INTO tbl_Service_Charge_SetupDetail " & _
                                      " (MasterKey, Line, ForCompany, Rate) " & _
                                      " VALUES (" & MasterKey & ", " & RowLine & ", " & _
                                      " '" & Trim(.Item(i).SubItems(1)) & "', " & _
                                      " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2))) & ")"
                End If
            Next i
        End With
    End If
    
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    'Me.Caption = "Service Charge - Browse"
    txtEffectDate.SetFocus
    BROWSER strEffectDate, "is_LOAD"
    
End If
Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Function
End Function

Private Function PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    Unload Me
Else
    If picSLine.Visible = True Then
        If TRANS_DETAIL = is_DET_ADDING Then
            With lstDetail.ListItems
                If .Count > 1 Then
                    .Remove ROW
                    ROW = .Count
                Else
                    .Item(ROW).SubItems(1) = " "
                    .Item(ROW).SubItems(2) = " "
                End If
            End With
        End If
        If TRANS_DETAIL = is_DET_EDITTING Then
            With lstDetail.ListItems
                .Item(ROW).SubItems(1) = txtForCompany1.Text
                .Item(ROW).SubItems(2) = txtRate1.Text
            End With
        End If
        picMain.Enabled = True
        picToolbar.Enabled = True
        picSLine.Visible = False
        lstDetail.ListItems(ROW).EnsureVisible
        lstDetail.ListItems(ROW).Selected = True
        lstDetail.SetFocus
        Exit Function
    End If
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    'Me.Caption = "Service Charge - Browse"
    BROWSER GetSetting(App.EXEName, "ServiceSetupEff", "SrvcSetupEff", ""), "is_LOAD"
End If
End Function


Private Function CLEARTEXT()
txtEffectDate.Text = ""
txtSupervisory.Text = ""
txtRankNFile.Text = ""
StatusBar1.Panels(1).Text = ""
StatusBar1.Panels(2).Text = ""
lstDetail.ListItems.Clear
Set x = lstDetail.ListItems.Add()
x.Text = ""
x.SubItems(1) = " "
x.SubItems(2) = " "
End Function

Private Function LOCKTEXT(bln As Boolean)
If bln Then
    txtEffectDate.Locked = True
    txtSupervisory.Locked = True
    txtRankNFile.Locked = True
Else
    txtEffectDate.Locked = False
    txtSupervisory.Locked = False
    txtRankNFile.Locked = False
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
            .Buttons(19).Image = 10
            .Buttons(21).Image = 12
            .Buttons(23).Image = 13
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "First"
            .Buttons(9).Caption = "Back"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            .Buttons(19).Caption = "Post"
            .Buttons(21).Caption = "Refresh"
            .Buttons(23).Caption = "Close"
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
            .Buttons(23).Enabled = True
            .Buttons(1).ToolTipText = "NEW (Ins)"
            .Buttons(3).ToolTipText = "EDIT (F2)"
            .Buttons(5).ToolTipText = "DELETE (Del)"
            .Buttons(7).ToolTipText = "FIRST (Home)"
            .Buttons(9).ToolTipText = "BACK (PgUp)"
            .Buttons(11).ToolTipText = "NEXT (PgDown)"
            .Buttons(13).ToolTipText = "LAST (End)"
            .Buttons(15).ToolTipText = "FIND (F6)"
            .Buttons(17).ToolTipText = "PRINT (F9)"
            .Buttons(19).ToolTipText = "POST (F8)"
            .Buttons(21).ToolTipText = "REFRESH (F11)"
            .Buttons(23).ToolTipText = "CLOSE (Esc)"
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
            .Buttons(19).Image = 10
            .Buttons(21).Image = 12
            .Buttons(23).Image = 13
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "Save"
            .Buttons(9).Caption = "Undo"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            .Buttons(19).Caption = "Post"
            .Buttons(21).Caption = "Refresh"
            .Buttons(23).Caption = "Close"
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
            .Buttons(23).Enabled = False
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
            .Buttons(23).ToolTipText = ""
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
            .Buttons(19).Image = 10
            .Buttons(21).Image = 12
            .Buttons(23).Image = 13
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "First"
            .Buttons(9).Caption = "Undo"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            .Buttons(19).Caption = "Post"
            .Buttons(21).Caption = "Refresh"
            .Buttons(23).Caption = "Close"
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
            .Buttons(23).Enabled = False
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
            .Buttons(23).ToolTipText = ""
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
            .Buttons(19).Image = 10
            .Buttons(21).Image = 12
            .Buttons(23).Image = 13
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "Save"
            .Buttons(9).Caption = "Undo"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            .Buttons(19).Caption = "Post"
            .Buttons(21).Caption = "Refresh"
            .Buttons(23).Caption = "Close"
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
            .Buttons(23).Enabled = False
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
            .Buttons(23).ToolTipText = ""
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
            .Buttons(19).Image = 10
            .Buttons(21).Image = 12
            .Buttons(23).Image = 13
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "Save"
            .Buttons(9).Caption = "Undo"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            .Buttons(19).Caption = "Post"
            .Buttons(21).Caption = "Refresh"
            .Buttons(23).Caption = "Close"
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
            .Buttons(23).Enabled = False
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
            .Buttons(23).ToolTipText = ""
        Case 6      '=== NOT EMPTY DETAIL NAME ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 14
            .Buttons(9).Image = 15
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(21).Image = 12
            .Buttons(23).Image = 13
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "Save"
            .Buttons(9).Caption = "Undo"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            .Buttons(19).Caption = "Post"
            .Buttons(21).Caption = "Refresh"
            .Buttons(23).Caption = "Close"
            .Buttons(1).Enabled = True
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = True
            .Buttons(7).Enabled = True
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
            .Buttons(19).Enabled = False
            .Buttons(21).Enabled = False
            .Buttons(23).Enabled = False
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
            .Buttons(23).ToolTipText = ""
    End Select
End With
End Sub


Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
If TRANSACTIONTYPE = is_REFRESH Then
    BROWSER GetSetting(App.EXEName, "ServiceSetupEff", "SrvcSetupEff", ""), "is_LOAD"
    If Trim(txtEffectDate.Text) = "" Then BROWSER GetSetting(App.EXEName, "ServiceSetupEff", "SrvcSetupEff", ""), "is_HOME"
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert:   PRESS_INSERT
    Case vbKeyF2:       PRESS_F2
    Case vbKeyDelete:   PRESS_DELETE
    Case vbKeyF5:       PRESS_F5
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "ServiceSetupEff", "SrvcSetupEff", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "ServiceSetupEff", "SrvcSetupEff", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "ServiceSetupEff", "SrvcSetupEff", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "ServiceSetupEff", "SrvcSetupEff", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
'Me.Caption = "Service Charge - Browse"
ListFocus = 0
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANS_DETAIL = is_DET_REFRESH
'MsgBox GetSetting(App.EXEName, "ServiceSetupEff", "SrvcSetupEff", "")
BROWSER GetSetting(App.EXEName, "ServiceSetupEff", "SrvcSetupEff", ""), "is_LOAD"
If Trim(txtEffectDate.Text) = "" Then BROWSER GetSetting(App.EXEName, "ServiceSetupEff", "SrvcSetupEff", ""), "is_HOME"

tmp = SetWindowLong(txtForCompany.hwnd, GWL_STYLE, GetWindowLong(txtForCompany.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub lstDetail_GotFocus()
TRANS_DETAIL = is_DET_REFRESH
ROW = lstDetail.SelectedItem.Index
ListFocus = 1
End Sub

Private Sub lstDetail_ItemClick(ByVal Item As MSComctlLib.ListItem)
ROW = lstDetail.SelectedItem.Index
End Sub

Private Sub lstDetail_LostFocus()
ListFocus = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
        Case "Refresh"
            'ToDo: Add 'Refresh' button code.
            MsgBox "Add 'Refresh' button code."
        Case "Post"
            'ToDo: Add 'Post' button code.
            MsgBox "Add 'Post' button code."
        Case "Print"
            'ToDo: Add 'Print' button code.
            MsgBox "Add 'Print' button code."
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "ServiceSetupEff", "SrvcSetupEff", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "ServiceSetupEff", "SrvcSetupEff", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "ServiceSetupEff", "SrvcSetupEff", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "ServiceSetupEff", "SrvcSetupEff", ""), "is_END"
    Case "Find":
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtEffectDate_GotFocus()
HTEXT txtEffectDate
End Sub

Private Sub txtEffectDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtSupervisory.SetFocus
End If
End Sub

Private Sub txtEffectDate_LostFocus()
If IsDate(txtEffectDate.Text) = True Then
    txtEffectDate.Text = Format(FormatDateTime(txtEffectDate.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub

Private Sub txtForCompany_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(ROW).SubItems(1) = Trim(txtForCompany.Text)
    End With
End If
End Sub

Private Sub txtForCompany_GotFocus()
HTEXT txtForCompany
End Sub

Private Sub txtForCompany_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtRate.SetFocus
End Sub

Private Sub txtRankNFile_GotFocus()
HTEXT txtRankNFile
End Sub

Private Sub txtRankNFile_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    lstDetail.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtSupervisory.SetFocus
End If
End Sub

Private Sub txtRankNFile_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtRate_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(ROW).SubItems(2) = Format(RETURNTEXTVALUE(txtRate), "#,##0.00")
    End With
End If
End Sub

Private Sub txtRate_GotFocus()
HTEXT txtRate
End Sub

Private Sub txtRate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then txtForCompany.SetFocus: Exit Sub
If KeyCode = vbKeyReturn Then
    picSLine.Visible = False
    picMain.Enabled = True
    picToolbar.Enabled = True
    lstDetail.ListItems(ROW).EnsureVisible
    lstDetail.ListItems(ROW).Selected = True
    lstDetail.SetFocus
End If
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtSupervisory_GotFocus()
HTEXT txtSupervisory
End Sub

Private Sub txtSupervisory_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtRankNFile.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtEffectDate.SetFocus
End If
End Sub

Private Sub txtSupervisory_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub
