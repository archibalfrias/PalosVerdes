VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMembershipShareID 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShareID.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picSearch 
      Height          =   3855
      Left            =   2280
      TabIndex        =   20
      Top             =   480
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6800
      BackColor       =   15396057
      Begin VB.CommandButton cmdOK 
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
         Picture         =   "frmShareID.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3195
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancel 
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
         Picture         =   "frmShareID.frx":133C
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3195
         Width           =   1560
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   4215
      End
      Begin VB.ListBox lstResult 
         Height          =   2205
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   4215
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   40
         TabIndex        =   25
         Top             =   40
         Width           =   4365
         _ExtentX        =   7699
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
         Icon            =   "frmShareID.frx":1A98
         ShadowVisible   =   0   'False
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   18
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   19
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
         MouseIcon       =   "frmShareID.frx":2032
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
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   2040
      ScaleHeight     =   3135
      ScaleWidth      =   4935
      TabIndex        =   4
      Top             =   1200
      Width           =   4935
      Begin VB.TextBox txtShareCertNumber 
         Height          =   315
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   16
         Top             =   2400
         Width           =   3735
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C6B8A4&
         Height          =   1215
         Left            =   0
         TabIndex        =   8
         Top             =   960
         Width           =   4935
         Begin VB.PictureBox picWithLotDetails 
            BackColor       =   &H00C6B8A4&
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   360
            ScaleHeight     =   735
            ScaleWidth      =   4455
            TabIndex        =   11
            Top             =   360
            Width           =   4455
            Begin VB.TextBox txtLotTitleNumber 
               Height          =   315
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   14
               Top             =   360
               Width           =   3255
            End
            Begin VB.TextBox txtLotArea 
               Height          =   315
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   12
               Top             =   0
               Width           =   3255
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Title Number"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   15
               Top             =   390
               Width           =   1095
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Area (sq. mtr)"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   13
               Top             =   30
               Width           =   1095
            End
         End
         Begin VB.PictureBox picWithLot 
            BackColor       =   &H00C6B8A4&
            BorderStyle     =   0  'None
            Height          =   350
            Left            =   120
            ScaleHeight     =   345
            ScaleWidth      =   1575
            TabIndex        =   9
            Top             =   0
            Width           =   1575
            Begin VB.CheckBox chkWithLot 
               BackColor       =   &H00C6B8A4&
               Caption         =   "With Fairway Lot"
               Height          =   255
               Left            =   0
               TabIndex        =   10
               Top             =   0
               Width           =   1575
            End
         End
      End
      Begin VB.TextBox txtRemarks 
         Height          =   315
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   2
         Top             =   2760
         Width           =   3735
      End
      Begin VB.ComboBox cmbShareType 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox txtIDNumber 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   0
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Share Cert. No."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Share Type"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   390
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Number"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   30
         Width           =   1095
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   4590
      Width           =   8985
      _ExtentX        =   15849
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8400
      Top             =   960
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
            Picture         =   "frmShareID.frx":234C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShareID.frx":3026
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShareID.frx":3D00
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShareID.frx":49DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShareID.frx":56B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShareID.frx":638E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShareID.frx":7068
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShareID.frx":7D42
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShareID.frx":8A1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShareID.frx":92F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShareID.frx":9FD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShareID.frx":ACAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShareID.frx":B984
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShareID.frx":C65E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShareID.frx":D338
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMembershipShareID"
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

Dim sIDNumber

Private Sub BROWSER(strID, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If strID <> "" Then
            s = "SELECT TOP 1 tbl_Share_IDNumber.* " & _
                " FROM tbl_Share_IDNumber " & _
                " WHERE (IDNumber = '" & strID & "') " & _
                " ORDER BY IDNumber"
        Else
            s = "SELECT TOP 1 tbl_Share_IDNumber.* " & _
                " FROM tbl_Share_IDNumber " & _
                " ORDER BY IDNumber"
        End If
    Case "is_HOME"
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Share_IDNumber.* " & _
            " FROM tbl_Share_IDNumber " & _
            " ORDER BY IDNumber"
    Case "is_PAGEUP"
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Share_IDNumber.* " & _
            " FROM tbl_Share_IDNumber " & _
            " WHERE (IDNumber < '" & strID & "') " & _
            " ORDER BY IDNumber DESC"
    Case "is_PAGEDOWN"
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Share_IDNumber.* " & _
            " FROM tbl_Share_IDNumber " & _
            " WHERE (IDNumber > '" & strID & "') " & _
            " ORDER BY IDNumber "
    Case "is_END"
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Share_IDNumber.* " & _
            " FROM tbl_Share_IDNumber " & _
            " ORDER BY IDNumber DESC"
    Case Else: Exit Sub
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtIDNumber.Text = rs!IDNumber
    cmbShareType.ListIndex = rs!ShareType
    chkWithLot.Value = rs!WithLot
    txtLotArea.Text = rs!LotArea
    txtLotTitleNumber.Text = rs!LotTitleNumber
    txtShareCertNumber.Text = rs!ShareCertNo
    txtRemarks.Text = rs!Remarks
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    SaveSetting App.EXEName, "ShareIDNumber", "ShareID", rs!IDNumber
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If AccessRights("Share ID Number", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
txtIDNumber.SetFocus
End Sub

Private Sub PRESS_F2()
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("Share ID Number", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_EDITTING
End Sub

Private Sub PRESS_DELETE()
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("Share ID Number", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If MsgBox("ARE YOU SURE IN DELETING THIS RECORD?                           ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_Share_IDNumber WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "ShareIDNumber", "ShareID", ""), "is_PAGEDOWN"
If Trim(txtIDNumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "ShareIDNumber", "ShareID", ""), "is_HOME"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()

If picSearch.Visible = True Then Exit Sub
If Trim(txtIDNumber.Text) = "" Then MsgBox "Please Supply ID Number!                  ", vbCritical, "Error...": txtIDNumber.SetFocus: Exit Sub
If cmbShareType.ListIndex <= 0 Then MsgBox "Please Select Share Type!                 ", vbCritical, "Error...": cmbShareType.SetFocus: Exit Sub
If chkWithLot.Value = 0 Then
    txtLotArea.Text = "": txtLotTitleNumber.Text = ""
End If
On Error GoTo PG:
sIDNumber = Trim(txtIDNumber.Text)
If TRANSACTIONTYPE = is_ADDING Then
    ConnOmega.Execute "INSERT INTO tbl_Share_IDNumber " & _
                      " (IDNumber, ShareType, Remarks, LastModified, " & _
                      " WithLot, LotArea, LotTitleNumber, ShareCertNo) " & _
                      " VALUES ('" & Trim(txtIDNumber.Text) & "', " & _
                      " " & cmbShareType.ListIndex & ", " & _
                      " '" & FORMATSQL(txtRemarks.Text) & "', " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                      " " & chkWithLot.Value & ", '" & txtLotArea.Text & "', " & _
                      " '" & FORMATSQL(Trim(txtLotTitleNumber.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtShareCertNumber.Text)) & "')"
End If
If TRANSACTIONTYPE = is_EDITTING Then
    ConnOmega.Execute "UPDATE tbl_Share_IDNumber " & _
                      " SET IDNumber = '" & Trim(txtIDNumber.Text) & "', " & _
                      " ShareType = " & cmbShareType.ListIndex & ", " & _
                      " Remarks = '" & FORMATSQL(txtRemarks.Text) & "', " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                      " WithLot = " & chkWithLot.Value & ", " & _
                      " LotArea = '" & txtLotArea.Text & "', " & _
                      " LotTitleNumber = '" & FORMATSQL(Trim(txtLotTitleNumber.Text)) & "', " & _
                      " ShareCertNo = '" & FORMATSQL(Trim(txtShareCertNumber.Text)) & "' " & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
End If
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER sIDNumber, "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
picMain.Enabled = False
picToolbar.Enabled = False
picSearch.ZOrder 0
txtSearch.Text = ""
picSearch.Visible = True
txtSearch.SetFocus
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSearch.Visible = True Then cmdCancel_Click: Exit Sub
    Unload Me
Else
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "ShareIDNumber", "ShareID", ""), "is_LOAD"
End If
End Sub

Private Sub CLEARTEXT()
txtIDNumber.Text = ""
cmbShareType.ListIndex = 0
chkWithLot.Value = 0
txtLotArea.Text = ""
txtLotTitleNumber.Text = ""
txtShareCertNumber.Text = ""
txtRemarks.Text = ""
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtIDNumber.Locked = bln
cmbShareType.Locked = bln
picWithLot.Enabled = IIf(bln = True, False, True)
'If picWithLot.Enabled = True Then
'    If chkWithLot.Value = 0 Then
'        picWithLotDetails.Enabled = False
'    Else
'        picWithLotDetails.Enabled = True
'    End If
'End If
txtShareCertNumber.Locked = bln
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

Private Sub b8TitleBar1_CLoseClick()
cmdCancel_Click
End Sub

Private Sub chkWithLot_Click()
If chkWithLot.Value = 0 Then
    picWithLotDetails.Enabled = False
Else
    picWithLotDetails.Enabled = True
End If
End Sub

Private Sub cmdCancel_Click()
picSearch.Visible = False
picToolbar.Enabled = True
picMain.Enabled = True
End Sub

Private Sub cmdOK_Click()
If lstResult.ListIndex = -1 Then Exit Sub
BROWSER lstResult.List(lstResult.ListIndex), "is_LOAD"
cmdCancel_Click
End Sub

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert:       PRESS_INSERT
    Case vbKeyF2:           PRESS_F2
    Case vbKeyDelete:       PRESS_DELETE
    Case vbKeyF5:           PRESS_F5
    Case vbKeyF6:           PRESS_F6
    Case vbKeyEscape:       PRESS_ESCAPE
    Case vbKeyHome:         BROWSER GetSetting(App.EXEName, "ShareIDNumber", "ShareID", ""), "is_HOME"
    Case vbKeyPageUp:       BROWSER GetSetting(App.EXEName, "ShareIDNumber", "ShareID", ""), "is_PAGEUP"
    Case vbKeyPageDown:     BROWSER GetSetting(App.EXEName, "ShareIDNumber", "ShareID", ""), "is_PAGEDOWN"
    Case vbKeyEnd:          BROWSER GetSetting(App.EXEName, "ShareIDNumber", "ShareID", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.Height - Me.Height) / 3
Me.Left = (MainForm.Width - Me.Width) / 5
With cmbShareType
    .Clear
    .AddItem "--Select--"
    .ItemData(.NewIndex) = 0
    s = "SELECT tbl_Share_Type.* " & _
        " FROM tbl_Share_Type " & _
        " ORDER BY ShareTypeName "
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        .AddItem rs!ShareTypeName
        .ItemData(.NewIndex) = rs!PK
        rs.MoveNext
    Wend
    rs.Close
End With
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "ShareIDNumber", "ShareID", ""), "is_LOAD"
If Trim(txtIDNumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "ShareIDNumber", "ShareID", ""), "is_HOME"

tmp = SetWindowLong(txtIDNumber.hwnd, GWL_STYLE, GetWindowLong(txtIDNumber.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtRemarks.hwnd, GWL_STYLE, GetWindowLong(txtRemarks.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtLotTitleNumber.hwnd, GWL_STYLE, GetWindowLong(txtLotTitleNumber.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtShareCertNumber.hwnd, GWL_STYLE, GetWindowLong(txtShareCertNumber.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picSearch.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub


Private Sub lstResult_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "ShareIDNumber", "ShareID", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "ShareIDNumber", "ShareID", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "ShareIDNumber", "ShareID", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "ShareIDNumber", "ShareID", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtIDNumber_GotFocus()
HTEXT txtIDNumber
End Sub

Private Sub txtLotArea_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtRemarks_GotFocus()
HTEXT txtRemarks
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstResult.Clear: Exit Sub
lstResult.Clear
t = "SELECT tbl_Share_IDNumber.* " & _
    " FROM tbl_Share_IDNumber " & _
    " WHERE (IDNumber LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
    " ORDER BY IDNumber"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
While Not rt.EOF
    lstResult.AddItem rt!IDNumber
    lstResult.ItemData(lstResult.NewIndex) = rt!PK
    rt.MoveNext
Wend
rt.Close
If lstResult.ListCount Then lstResult.ListIndex = 0
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResult.SetFocus
End Sub
