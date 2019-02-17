VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGolfCartInformation 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGolfCartInformation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   9015
   Begin RPVGCC.b8Container picSearch 
      Height          =   3615
      Left            =   2040
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6376
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
         Picture         =   "frmGolfCartInformation.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2980
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
         Picture         =   "frmGolfCartInformation.frx":1FF4
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2980
         Width           =   1560
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   4215
      End
      Begin VB.ListBox lstResult 
         Height          =   2010
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   4215
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   21
         Top             =   45
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
         Icon            =   "frmGolfCartInformation.frx":2750
         ShadowVisible   =   0   'False
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   40
         TabIndex        =   22
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
         Icon            =   "frmGolfCartInformation.frx":2CEA
         ShadowVisible   =   0   'False
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   720
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
            Picture         =   "frmGolfCartInformation.frx":3284
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGolfCartInformation.frx":3F5E
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGolfCartInformation.frx":4C38
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGolfCartInformation.frx":5912
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGolfCartInformation.frx":65EC
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGolfCartInformation.frx":72C6
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGolfCartInformation.frx":7FA0
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGolfCartInformation.frx":8C7A
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGolfCartInformation.frx":9954
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGolfCartInformation.frx":A22E
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGolfCartInformation.frx":AF08
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGolfCartInformation.frx":BBE2
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGolfCartInformation.frx":C8BC
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGolfCartInformation.frx":D596
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGolfCartInformation.frx":E270
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   23
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   24
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
         MouseIcon       =   "frmGolfCartInformation.frx":EF4A
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   11460
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   25
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmGolfCartInformation.frx":F264
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
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   1560
      ScaleHeight     =   2535
      ScaleWidth      =   5535
      TabIndex        =   8
      Top             =   1200
      Width           =   5535
      Begin VB.TextBox txtDescription 
         Height          =   315
         Left            =   960
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2160
         Width           =   4575
      End
      Begin VB.ComboBox cmbOwnerType 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1080
         Width           =   4575
      End
      Begin VB.TextBox txtCoOwner 
         Height          =   315
         Left            =   960
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1800
         Width           =   4575
      End
      Begin VB.TextBox txtOwner 
         Height          =   315
         Left            =   960
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1440
         Width           =   4575
      End
      Begin VB.TextBox txtEngineNo 
         Height          =   315
         Left            =   960
         MaxLength       =   50
         TabIndex        =   2
         Top             =   720
         Width           =   4575
      End
      Begin VB.TextBox txtChasisNo 
         Height          =   315
         Left            =   960
         MaxLength       =   50
         TabIndex        =   1
         Top             =   360
         Width           =   4575
      End
      Begin VB.TextBox txtGolfCartNo 
         Height          =   315
         Left            =   960
         MaxLength       =   50
         TabIndex        =   0
         Top             =   0
         Width           =   4575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Owner Type"
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Co-Owner"
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Owner"
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Engine No"
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Chasis No"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "GolfCart No"
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   975
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   7
      Top             =   3930
      Width           =   9015
      _ExtentX        =   15901
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
Attribute VB_Name = "frmGolfCartInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim sGolfCart   As String
Dim tmp         As Long


Private Sub BROWSER(sGolf, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If sGolf <> "" Then
            s = "SELECT TOP 1 tbl_GolfCart_Info.* " & _
                " FROM tbl_GolfCart_Info " & _
                " WHERE (GolfCartNo = '" & FORMATSQL(CStr(sGolf)) & "') " & _
                " ORDER BY GolfCartNo"
        Else
            s = "SELECT TOP 1 tbl_GolfCart_Info.* " & _
                " FROM tbl_GolfCart_Info " & _
                " ORDER BY GolfCartNo"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_GolfCart_Info.* " & _
            " FROM tbl_GolfCart_Info " & _
            " ORDER BY GolfCartNo"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_GolfCart_Info.* " & _
            " FROM tbl_GolfCart_Info " & _
            " WHERE (GolfCartNo < '" & FORMATSQL(CStr(sGolf)) & "') " & _
            " ORDER BY GolfCartNo DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_GolfCart_Info.* " & _
            " FROM tbl_GolfCart_Info " & _
            " WHERE (GolfCartNo > '" & FORMATSQL(CStr(sGolf)) & "') " & _
            " ORDER BY GolfCartNo "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_GolfCart_Info.* " & _
            " FROM tbl_GolfCart_Info " & _
            " ORDER BY GolfCartNo DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtGolfCartNo.Text = rs!GolfCartNo
    txtChasisNo.Text = rs!ChasisNo
    txtEngineNo.Text = rs!EngineNo
    cmbOwnerType.ListIndex = rs!OwnerType - 1
'    If IsNull(rs!Owner) = True Then
'        txtOwner.Text = ""
'    Else
'        txtOwner.Text = ""
'    End If
'    If IsNull(rs!CoOwner) = True Then
'        txtCoOwner.Text = ""
'    Else
'        txtCoOwner.Text = ""
'    End If
    txtOwner.Text = IIf(IsNull(rs!Owner), "", rs!Owner)
    txtCoOwner.Text = IIf(IsNull(rs!CoOwner), "", rs!CoOwner)
    txtDescription.Text = rs!Description
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    
    SaveSetting App.EXEName, "GolfCartInformation", "GolfCartInfo", rs!GolfCartNo
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If AccessRights("Golf Cart Information", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING

End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If AccessRights("Golf Cart Information", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_EDITTING
End Sub

Private Sub PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If AccessRights("Golf Cart Information", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If MsgBox("ARE YOU SURE IN DELETING THIS RECORD?                        ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_GolfCart_Info WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "GolfCartInformation", "GolfCartInfo", ""), "is_PAGEDOWN"
If Trim(txtGolfCartNo.Text) = "" Then BROWSER GetSetting(App.EXEName, "GolfCartInformation", "GolfCartInfo", ""), "is_HOME"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If Trim(txtGolfCartNo.Text) = "" Then MsgBox "Please Supply Golf Cart Number!                 ", vbCritical, "Error...": txtGolfCartNo.SetFocus: Exit Sub
If cmbOwnerType.ListIndex = -1 Then MsgBox "Please Select Owner Type!                ", vbCritical, "Error...": cmbOwnerType.SetFocus: Exit Sub
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    ConnOmega.Execute "INSERT INTO tbl_GolfCart_Info " & _
                      " (GolfCartNo, ChasisNo, EngineNo, OwnerType, Owner, CoOwner, Description, LastModified) " & _
                      " VALUES ('" & FORMATSQL(txtGolfCartNo.Text) & "', '" & FORMATSQL(txtChasisNo.Text) & "', " & _
                      " '" & FORMATSQL(txtEngineNo.Text) & "', " & cmbOwnerType.ListIndex + 1 & ", " & _
                      " '" & FORMATSQL(txtOwner.Text) & "', '" & FORMATSQL(txtCoOwner.Text) & "', " & _
                      " '" & FORMATSQL(txtDescription.Text) & "', '" & CStr(Now) & " - " & gbl_CompleteName & "')"
End If
If TRANSACTIONTYPE = is_EDITTING Then
    ConnOmega.Execute "UPDATE tbl_GolfCart_Info " & _
                      " SET GolfCartNo = '" & FORMATSQL(txtGolfCartNo.Text) & "', " & _
                      " ChasisNo = '" & FORMATSQL(txtChasisNo.Text) & "', " & _
                      " EngineNo = '" & FORMATSQL(txtEngineNo.Text) & "', " & _
                      " OwnerType = " & cmbOwnerType.ListIndex + 1 & ", " & _
                      " Owner = '" & FORMATSQL(txtOwner.Text) & "', " & _
                      " CoOwner = '" & FORMATSQL(txtCoOwner.Text) & "', " & _
                      " Description = '" & FORMATSQL(txtDescription.Text) & "', " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
End If
sGolfCart = Trim(txtGolfCartNo.Text)
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER sGolfCart, "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSearch.Visible = True Then Exit Sub
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
    BROWSER GetSetting(App.EXEName, "GolfCartInformation", "GolfCartInfo", ""), "is_LOAD"
End If
End Sub

Private Sub CLEARTEXT()
txtGolfCartNo.Text = ""
txtChasisNo.Text = ""
txtEngineNo.Text = ""
cmbOwnerType.ListIndex = -1
txtOwner.Text = ""
txtCoOwner.Text = ""
txtDescription.Text = ""
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtGolfCartNo.Locked = bln
txtChasisNo.Locked = bln
txtEngineNo.Locked = bln
cmbOwnerType.Locked = bln
txtOwner.Locked = bln
txtCoOwner.Locked = bln
txtDescription.Locked = bln
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

Private Sub b8TitleBar2_CLoseClick()
cmdCancel_Click
End Sub

Private Sub cmdCancel_Click()
picMain.Enabled = True
picToolbar.Enabled = True
picSearch.Visible = False
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
    Case vbKeyInsert:   PRESS_INSERT
    Case vbKeyF2:       PRESS_F2
    Case vbKeyDelete:   PRESS_DELETE
    Case vbKeyF5:       PRESS_F5
    Case vbKeyF6:       PRESS_F6
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "GolfCartInformation", "GolfCartInfo", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "GolfCartInformation", "GolfCartInfo", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "GolfCartInformation", "GolfCartInfo", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "GolfCartInformation", "GolfCartInfo", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption

cmbOwnerType.Clear
s = "SELECT tbl_GolfCart_Owner_Type.* " & _
    " FROM tbl_GolfCart_Owner_Type " & _
    " ORDER BY PK"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbOwnerType.AddItem rs!sName
    rs.MoveNext
Wend
rs.Close

CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "GolfCartInformation", "GolfCartInfo", ""), "is_LOAD"
If Trim(txtGolfCartNo.Text) = "" Then BROWSER GetSetting(App.EXEName, "GolfCartInformation", "GolfCartInfo", ""), "is_HOME"


tmp = SetWindowLong(txtGolfCartNo.hwnd, GWL_STYLE, GetWindowLong(txtGolfCartNo.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtChasisNo.hwnd, GWL_STYLE, GetWindowLong(txtChasisNo.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtEngineNo.hwnd, GWL_STYLE, GetWindowLong(txtEngineNo.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtOwner.hwnd, GWL_STYLE, GetWindowLong(txtOwner.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtCoOwner.hwnd, GWL_STYLE, GetWindowLong(txtCoOwner.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtDescription.hwnd, GWL_STYLE, GetWindowLong(txtDescription.hwnd, GWL_STYLE) Or ES_UPPERCASE)
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
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "GolfCartInformation", "GolfCartInfo", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "GolfCartInformation", "GolfCartInfo", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "GolfCartInformation", "GolfCartInfo", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "GolfCartInformation", "GolfCartInfo", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstResult.Clear: Exit Sub
lstResult.Clear
s = "SELECT tbl_GolfCart_Info.* " & _
    " FROM tbl_GolfCart_Info " & _
    " WHERE (GolfCartNo LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
    " ORDER BY GolfCartNo"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResult.AddItem rs!GolfCartNo
    rs.MoveNext
Wend
rs.Close
If lstResult.ListCount Then lstResult.ListIndex = 0
End Sub

Private Sub txtSearch_GotFocus()
HTEXT txtSearch
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then lstResult.SetFocus
End Sub
