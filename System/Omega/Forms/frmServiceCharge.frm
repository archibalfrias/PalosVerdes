VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmServiceCharge 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11430
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServiceCharge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   14
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   15
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
         MouseIcon       =   "frmServiceCharge.frx":08CA
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   9900
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   16
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmServiceCharge.frx":0BE4
               Top             =   120
               Visible         =   0   'False
               Width           =   1395
            End
         End
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   15000
         Y1              =   910
         Y2              =   910
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
         Y1              =   1005
         Y2              =   1005
      End
   End
   Begin RPVGCC.b8Container picAdd 
      Height          =   1935
      Left            =   4080
      TabIndex        =   8
      Top             =   2280
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3413
      BackColor       =   15396057
      Begin VB.TextBox txtYear 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox cmbMonth 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   720
         Width           =   1815
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   45
         TabIndex        =   11
         Top             =   45
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   609
         Caption         =   "Select Month / Year"
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
         Icon            =   "frmServiceCharge.frx":12F7
      End
      Begin VB.CommandButton cmdCancelAdd 
         Height          =   480
         Left            =   2040
         Picture         =   "frmServiceCharge.frx":1891
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1200
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKAdd 
         Height          =   480
         Left            =   240
         Picture         =   "frmServiceCharge.frx":1FED
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1200
         Width           =   1560
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   6375
      Width           =   11430
      _ExtentX        =   20161
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
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   2640
      ScaleHeight     =   4695
      ScaleWidth      =   6150
      TabIndex        =   3
      Top             =   1320
      Width           =   6150
      Begin VB.TextBox txtMonthYear 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3240
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   1
         Top             =   0
         Width           =   2175
      End
      Begin VB.TextBox txtCtrl 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   0
         Top             =   0
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid FGrid01 
         Height          =   4200
         Left            =   0
         TabIndex        =   6
         Top             =   480
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   7408
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   13023396
         ForeColorFixed  =   255
         BackColorSel    =   8388608
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         FocusRect       =   0
      End
      Begin MSFlexGridLib.MSFlexGrid FGrid02 
         Height          =   4200
         Left            =   3120
         TabIndex        =   7
         Top             =   480
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   7408
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   13023396
         ForeColorFixed  =   255
         BackColorSel    =   8388608
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         FocusRect       =   0
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Month / Year"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl #"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   0
         Width           =   495
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10440
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
            Picture         =   "frmServiceCharge.frx":265F
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceCharge.frx":3339
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceCharge.frx":4013
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceCharge.frx":4CED
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceCharge.frx":59C7
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceCharge.frx":66A1
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceCharge.frx":737B
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceCharge.frx":8055
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceCharge.frx":8D2F
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceCharge.frx":9609
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceCharge.frx":A2E3
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceCharge.frx":AFBD
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceCharge.frx":BC97
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceCharge.frx":C971
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceCharge.frx":D64B
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmServiceCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iMonth As Double
Dim iYear As Double

Dim PressCount As Double

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim i, dblSC, LTEXT, strCtrl, iPK, iLine, HEADER1$, HEADER2$

Private Sub BROWSER(sCtrl, is_Action As String)
Select Case is_Action
    Case "is_LOAD"
        If sCtrl <> "" Then
            s = "SELECT TOP 1 tbl_Service_Charge.* " & _
                " FROM tbl_Service_Charge " & _
                " WHERE (Ctrl = '" & sCtrl & "') " & _
                " ORDER BY Ctrl"
        Else
            s = "SELECT TOP 1 tbl_Service_Charge.* " & _
                " FROM tbl_Service_Charge " & _
                " ORDER BY Ctrl"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Service_Charge.* " & _
            " FROM tbl_Service_Charge " & _
            " ORDER BY Ctrl"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Service_Charge.* " & _
            " FROM tbl_Service_Charge " & _
            " WHERE (Ctrl < '" & sCtrl & "') " & _
            " ORDER BY Ctrl DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Service_Charge.* " & _
            " FROM tbl_Service_Charge " & _
            " WHERE (Ctrl > '" & sCtrl & "') " & _
            " ORDER BY Ctrl"
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Service_Charge.* " & _
            " FROM tbl_Service_Charge " & _
            " ORDER BY Ctrl DESC"
    Case Else: Exit Sub
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iMonth = rs!sMonth
    iYear = rs!sYear
    txtCtrl.Text = rs!Ctrl
    txtMonthYear.Text = rs!MonthYear
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    
    imgPosted.Visible = IIf(rs!Posted = 1, True, False)
    Toolbar1.Buttons(19).Caption = IIf(rs!Posted = 1, "UnPost", " Post ")
    Toolbar1.Buttons(19).Image = IIf(rs!Posted = 1, 11, 10)
    
    SaveSetting App.EXEName, "ServiceChargeCtrl", "SCCtrl", rs!Ctrl
    
    CUSTOM_GRID
    i = 0
    t = "SELECT tbl_Service_Charge_Detail.* " & _
        " FROM tbl_Service_Charge_Detail " & _
        " WHERE (MasterKey = " & rs!PK & ") " & _
        " ORDER BY sDate"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        i = i + 1
        If i >= 1 And i <= 16 Then
            FGrid01.TextMatrix(i, 1) = Format(rt!sDate, "dd-mmm-yyyy")
            FGrid01.TextMatrix(i, 2) = Format(rt!ServiceCharge, "#,##0.00")
        End If
        If i >= 17 Then
            FGrid02.TextMatrix(i - 16, 1) = Format(rt!sDate, "dd-mmm-yyyy")
            FGrid02.TextMatrix(i - 16, 2) = Format(rt!ServiceCharge, "#,##0.00")
        End If
        rt.MoveNext
    Wend
    rt.Close
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If AccessRights("Service Charge", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
picMain.Enabled = False
picToolbar.Enabled = False
picAdd.ZOrder 0
cmbMonth.ListIndex = Month(Date) - 1
txtYear.Text = Format(Date, "yyyy")
picAdd.Visible = True
cmbMonth.SetFocus
End Sub

Private Sub PRESS_F2()
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If imgPosted.Visible = True Then MsgBox "Already Posted!             ", vbCritical, "Error...": Exit Sub
If AccessRights("Service Charge", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
TOOLBARFUNC 2
TRANSACTIONTYPE = is_EDITTING
End Sub

Private Sub PRESS_DELETE()
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If imgPosted.Visible = True Then MsgBox "Already Posted!             ", vbCritical, "Error...": Exit Sub
If AccessRights("Service Charge", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
On Error GoTo PG:
If MsgBox("ARE YOU SURE IN DELETING THIS RECORD?                    ", vbCritical + vbYesNo + vbDefaultButton2, "Error...") = vbNo Then Exit Sub
ConnOmega.Execute "DELETE FROM tbl_Service_Charge WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "ServiceChargeCtrl", "SCCtrl", ""), "is_PAGEDOWN"
If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "ServiceChargeCtrl", "SCCtrl", ""), "is_HOME"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    strCtrl = ""
    s = "SELECT TOP 1 tbl_Service_Charge.* " & _
        " FROM tbl_Service_Charge " & _
        " WHERE (sYear = " & iYear & ") " & _
        " ORDER BY Ctrl DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        strCtrl = Format(CDbl(rs!Ctrl) + 1, "0000000#")
    Else
        strCtrl = CStr(iYear) & "0000"
    End If
    rs.Close
    Do
        s = "SELECT tbl_Service_Charge.* " & _
            " FROM tbl_Service_Charge " & _
            " WHERE (Ctrl = '" & strCtrl & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If
        rs.Close
        strCtrl = Format(CDbl(strCtrl) + 1, "0000000#")
    Loop
    
    ConnOmega.Execute "INSERT INTO tbl_Service_Charge " & _
                      " (Ctrl, sMonth, sYear, MonthYear, LastModified) " & _
                      " VALUES ('" & strCtrl & "', " & iMonth & ", " & _
                      " " & iYear & ", '" & txtMonthYear.Text & "', " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
    
    iPK = 0
    s = "SELECT PK " & _
        " FROM tbl_Service_Charge " & _
        " WHERE (Ctrl= '" & strCtrl & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        iPK = rs!PK
    End If
    rs.Close
    
    If CDbl(iPK) > 0 Then
        With FGrid01
            For i = 1 To .Rows - 1
                If IsDate(.TextMatrix(i, 1)) = True Then
                    ConnOmega.Execute "INSERT INTO tbl_Service_Charge_Detail " & _
                                      " (MasterKey, sDate, ServiceCharge) " & _
                                      " VALUES (" & iPK & ", '" & FormatDateTime(.TextMatrix(i, 1), vbShortDate) & "', " & _
                                      " " & CDbl(IIf(IsNumeric(.TextMatrix(i, 2)) = False, 0, .TextMatrix(i, 2))) & ")"
                End If
            Next i
        End With
        With FGrid02
            For i = 1 To .Rows - 1
                If IsDate(.TextMatrix(i, 1)) = True Then
                    ConnOmega.Execute "INSERT INTO tbl_Service_Charge_Detail " & _
                                      " (MasterKey, sDate, ServiceCharge) " & _
                                      " VALUES (" & iPK & ", '" & FormatDateTime(.TextMatrix(i, 1), vbShortDate) & "', " & _
                                      " " & CDbl(IIf(IsNumeric(.TextMatrix(i, 2)) = False, 0, .TextMatrix(i, 2))) & ")"
                End If
            Next i
        End With
    End If
    CLEARTEXT
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER strCtrl, "is_LOAD"
    txtCtrl.SetFocus
    
End If
If TRANSACTIONTYPE = is_EDITTING Then
    iPK = Statusbar1.Panels(1).Text
    
    ConnOmega.Execute "UPDATE tbl_Service_Charge " & _
                      " SET sMonth = " & iMonth & ", " & _
                      " sYear = " & iYear & ", " & _
                      " MonthYear = '" & txtMonthYear.Text & "', " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & iPK & ")"
    
    
    
    If CDbl(iPK) > 0 Then
        ConnOmega.Execute "DELETE FROM tbl_Service_Charge_Detail WHERE (MasterKey = " & iPK & ")"
        With FGrid01
            For i = 1 To .Rows - 1
                If IsDate(.TextMatrix(i, 1)) = True Then
                    ConnOmega.Execute "INSERT INTO tbl_Service_Charge_Detail " & _
                                      " (MasterKey, sDate, ServiceCharge) " & _
                                      " VALUES (" & iPK & ", '" & FormatDateTime(.TextMatrix(i, 1), vbShortDate) & "', " & _
                                      " " & CDbl(IIf(IsNumeric(.TextMatrix(i, 2)) = False, 0, .TextMatrix(i, 2))) & ")"
                End If
            Next i
        End With
        With FGrid02
            For i = 1 To .Rows - 1
                If IsDate(.TextMatrix(i, 1)) = True Then
                    ConnOmega.Execute "INSERT INTO tbl_Service_Charge_Detail " & _
                                      " (MasterKey, sDate, ServiceCharge) " & _
                                      " VALUES (" & iPK & ", '" & FormatDateTime(.TextMatrix(i, 1), vbShortDate) & "', " & _
                                      " " & CDbl(IIf(IsNumeric(.TextMatrix(i, 2)) = False, 0, .TextMatrix(i, 2))) & ")"
                End If
            Next i
        End With
    End If
    CLEARTEXT
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "ServiceChargeCtrl", "SCCtrl", ""), "is_LOAD"
    txtCtrl.SetFocus
End If
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
End Sub

Private Sub PRESS_F8()
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If imgPosted.Visible = True Then MsgBox "Already Posted!             ", vbCritical, "Error...": Exit Sub
If AccessRights("Service Charge", "Post") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If MsgBox("CONTINUE POSTING THIS TRANSACTION?                   ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
On Error GoTo PG:
ConnOmega.Execute "UPDATE tbl_Service_Charge SET Posted = 1 WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
BROWSER GetSetting(App.EXEName, "ServiceChargeCtrl", "SCCtrl", ""), "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picAdd.Visible = True Then cmdCancelAdd_Click: Exit Sub
    Unload Me
Else
    CLEARTEXT
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "ServiceChargeCtrl", "SCCtrl", ""), "is_LOAD"
    If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "ServiceChargeCtrl", "SCCtrl", ""), "is_HOME"
End If
End Sub

Private Function CLEARTEXT()
iMonth = 0
iYear = 0
txtCtrl.Text = ""
txtMonthYear.Text = ""
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
imgPosted.Visible = False
For i = 1 To 16
    FGrid01.TextMatrix(i, 1) = ""
    FGrid01.TextMatrix(i, 2) = ""
    FGrid02.TextMatrix(i, 1) = ""
    FGrid02.TextMatrix(i, 2) = ""
Next i
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

Private Function CUSTOM_GRID()
HEADER1$ = "": HEADER2$ = ""
With FGrid01
    HEADER1$ = HEADER1$ & "|" & _
               "Date" & "|" & _
               "S. C."
    .FormatString = HEADER1$
    .ColWidth(1) = 1300      'Date
    .ColWidth(2) = 1300      'S.C.
    .ColAlignment(1) = 3
    .ColAlignment(2) = flexAlignRightCenter
End With
With FGrid02
    HEADER2$ = HEADER2$ & "|" & _
               "Date" & "|" & _
               "S. C."
    .FormatString = HEADER2$
    .ColWidth(1) = 1300      'Date
    .ColWidth(2) = 1300      'S.C.
    .ColAlignment(1) = 3
    .ColAlignment(2) = flexAlignRightCenter
End With
For i = 1 To 16
    FGrid01.Rows = i + 1
    FGrid01.TextMatrix(i, 1) = ""
    FGrid01.TextMatrix(i, 2) = ""
    FGrid02.Rows = i + 1
    FGrid02.TextMatrix(i, 1) = ""
    FGrid02.TextMatrix(i, 2) = ""
Next i
End Function

Private Sub b8TitleBar1_CLoseClick()
cmdCancelAdd_Click
End Sub

Private Sub cmbMonth_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtYear.SetFocus
End Sub

Private Sub cmdCancelAdd_Click()
picMain.Enabled = True
picToolbar.Enabled = True
picAdd.Visible = False
End Sub

Private Sub cmdOKAdd_Click()
If cmbMonth.ListIndex = -1 Then Exit Sub
If RETURNTEXTVALUE(txtYear) = 0 Then Exit Sub
'CUSTOM_GRID
CLEARTEXT
With FGrid01
    For i = 1 To 16
        .TextMatrix(i, 1) = Format(DateSerial(RETURNTEXTVALUE(txtYear), cmbMonth.ListIndex + 1, i), "dd-mmm-yyyy")
        .TextMatrix(i, 2) = "0.00"
    Next i
End With
With FGrid02
    For i = 17 To Day(DateSerial(RETURNTEXTVALUE(txtYear), CDbl(cmbMonth.ListIndex + 1) + 1, 0))
        .TextMatrix(i - 16, 1) = Format(DateSerial(RETURNTEXTVALUE(txtYear), cmbMonth.ListIndex + 1, i), "dd-mmm-yyyy")
        .TextMatrix(i - 16, 2) = "0.00"
    Next i
End With
cmdCancelAdd_Click
iMonth = cmbMonth.ListIndex + 1
iYear = RETURNTEXTVALUE(txtYear)
txtMonthYear.Text = cmbMonth.List(cmbMonth.ListIndex) & " " & txtYear.Text
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
FGrid01.SetFocus
End Sub

Private Sub FGrid01_EnterCell()
With FGrid01
    If .Col = 2 Then
        dblSC = IIf(IsNumeric(.TextMatrix(.ROW, .Col)) = False, 0, .TextMatrix(.ROW, .Col))
        PressCount = 0
    End If
End With
End Sub

Private Sub FGrid01_KeyPress(KeyAscii As Integer)
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
With FGrid01
    If .Col = 2 Then
        If KeyAscii = 8 Then
            LTEXT = IIf(Len(.TextMatrix(.ROW, .Col)) > 0, Len(.TextMatrix(.ROW, .Col)) - 1, 0)
            .TextMatrix(.ROW, .Col) = Mid(.TextMatrix(.ROW, .Col), 1, LTEXT)
        ElseIf KeyAscii >= 1 And KeyAscii <= 7 Then
        ElseIf KeyAscii >= 9 And KeyAscii <= 13 Then
        ElseIf KeyAscii >= 14 And KeyAscii <= 44 Then
        ElseIf KeyAscii = 47 Then
        ElseIf KeyAscii >= 58 And KeyAscii <= 126 Then
        Else
            PressCount = PressCount + 1
            If PressCount > 1 Then
                .TextMatrix(.ROW, .Col) = .TextMatrix(.ROW, .Col) & Chr(KeyAscii)
            Else
                .TextMatrix(.ROW, .Col) = Chr(KeyAscii)
            End If
        End If
    End If
End With
End Sub

Private Sub FGrid01_LeaveCell()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
With FGrid01
    If .Col = 2 Then
        dblSC = IIf(IsNumeric(.TextMatrix(.ROW, .Col)) = False, 0, .TextMatrix(.ROW, .Col))
        .TextMatrix(.ROW, .Col) = Format(dblSC, "#,##0.00")
    End If
End With
End Sub

Private Sub FGrid02_EnterCell()
With FGrid02
    If .Col = 2 Then
        dblSC = IIf(IsNumeric(.TextMatrix(.ROW, .Col)) = False, 0, .TextMatrix(.ROW, .Col))
        PressCount = 0
    End If
End With
End Sub

Private Sub FGrid02_KeyPress(KeyAscii As Integer)
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
With FGrid02
    If .Col = 2 Then
        If KeyAscii = 8 Then
            LTEXT = IIf(Len(.TextMatrix(.ROW, .Col)) > 0, Len(.TextMatrix(.ROW, .Col)) - 1, 0)
            .TextMatrix(.ROW, .Col) = Mid(.TextMatrix(.ROW, .Col), 1, LTEXT)
        ElseIf KeyAscii >= 1 And KeyAscii <= 7 Then
        ElseIf KeyAscii >= 9 And KeyAscii <= 13 Then
        ElseIf KeyAscii >= 14 And KeyAscii <= 44 Then
        ElseIf KeyAscii = 47 Then
        ElseIf KeyAscii >= 58 And KeyAscii <= 126 Then
        Else
            PressCount = PressCount + 1
            If PressCount > 1 Then
                .TextMatrix(.ROW, .Col) = .TextMatrix(.ROW, .Col) & Chr(KeyAscii)
            Else
                .TextMatrix(.ROW, .Col) = Chr(KeyAscii)
            End If
        End If
    End If
End With
End Sub

Private Sub FGrid02_LeaveCell()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
With FGrid02
    If .Col = 2 Then
        dblSC = IIf(IsNumeric(.TextMatrix(.ROW, .Col)) = False, 0, .TextMatrix(.ROW, .Col))
        .TextMatrix(.ROW, .Col) = Format(dblSC, "#,##0.00")
    End If
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
    Case vbKeyF8:       PRESS_F8
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "ServiceChargeCtrl", "SCCtrl", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "ServiceChargeCtrl", "SCCtrl", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "ServiceChargeCtrl", "SCCtrl", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "ServiceChargeCtrl", "SCCtrl", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
CUSTOM_GRID
With cmbMonth
    .Clear
    .AddItem "January"
    .AddItem "February"
    .AddItem "March"
    .AddItem "April"
    .AddItem "May"
    .AddItem "June"
    .AddItem "July"
    .AddItem "August"
    .AddItem "September"
    .AddItem "October"
    .AddItem "November"
    .AddItem "December"
End With
CLEARTEXT
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "ServiceChargeCtrl", "SCCtrl", ""), "is_LOAD"
If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "ServiceChargeCtrl", "SCCtrl", ""), "is_HOME"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "ServiceChargeCtrl", "SCCtrl", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "ServiceChargeCtrl", "SCCtrl", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "ServiceChargeCtrl", "SCCtrl", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "ServiceChargeCtrl", "SCCtrl", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":
    Case "Post":    PRESS_F8
    Case "Refresh":
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtCtrl_GotFocus()
HTEXT txtCtrl
End Sub

Private Sub txtMonthYear_GotFocus()
HTEXT txtMonthYear
End Sub

Private Sub txtYear_GotFocus()
HTEXT txtYear
End Sub

Private Sub txtYear_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAdd_Click
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub
