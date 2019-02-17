VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCaddyInformation 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   10845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCaddyInformation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   30
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   31
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
         MouseIcon       =   "frmCaddyInformation.frx":0CCA
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   11460
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   32
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmCaddyInformation.frx":0FE4
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
   Begin RPVGCC.b8Container picSearch 
      Height          =   3615
      Left            =   3120
      TabIndex        =   16
      Top             =   840
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6376
      BackColor       =   15396057
      Begin VB.ListBox lstResult 
         Height          =   2010
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   4215
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
         Picture         =   "frmCaddyInformation.frx":16F7
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2980
         Width           =   1560
      End
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
         Picture         =   "frmCaddyInformation.frx":1E53
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2980
         Width           =   1560
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
         Icon            =   "frmCaddyInformation.frx":24C5
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
         Icon            =   "frmCaddyInformation.frx":2A5F
         ShadowVisible   =   0   'False
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1320
      TabIndex        =   27
      Top             =   5280
      Width           =   1455
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   360
      ScaleHeight     =   3375
      ScaleWidth      =   10335
      TabIndex        =   8
      Top             =   1200
      Width           =   10335
      Begin VB.TextBox txtPicturePath 
         Height          =   315
         Left            =   7200
         TabIndex        =   29
         Top             =   2520
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   6960
         ScaleHeight     =   3225
         ScaleWidth      =   3225
         TabIndex        =   28
         Top             =   0
         Width           =   3255
         Begin VB.Image imgPicture 
            Height          =   3225
            Left            =   0
            Stretch         =   -1  'True
            Top             =   -120
            Width           =   3225
         End
      End
      Begin VB.TextBox txtRelation 
         Height          =   315
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   25
         Top             =   2520
         Width           =   5415
      End
      Begin VB.TextBox txtNickName 
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   23
         Top             =   1440
         Width           =   5415
      End
      Begin VB.TextBox txtContactPersonNo 
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2880
         Width           =   5415
      End
      Begin VB.TextBox txtCaddyNo 
         Height          =   315
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   0
         Top             =   0
         Width           =   5415
      End
      Begin VB.TextBox txtLastName 
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   1
         Top             =   360
         Width           =   5415
      End
      Begin VB.TextBox txtFirstName 
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   2
         Top             =   720
         Width           =   5415
      End
      Begin VB.TextBox txtMiddleName 
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1080
         Width           =   5415
      End
      Begin VB.TextBox txtContactNo 
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1800
         Width           =   5415
      End
      Begin VB.TextBox txtContactPerson 
         Height          =   315
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   5
         Top             =   2160
         Width           =   5415
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         Height          =   3255
         Left            =   7035
         Top             =   75
         Width           =   3255
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Relation"
         Height          =   255
         Left            =   0
         TabIndex        =   26
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "NickName"
         Height          =   255
         Left            =   0
         TabIndex        =   24
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Person No"
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Caddy No"
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Person"
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   2160
         Width           =   1335
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   7
      Top             =   4695
      Width           =   10845
      _ExtentX        =   19129
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
      Left            =   0
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
            Picture         =   "frmCaddyInformation.frx":2FF9
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaddyInformation.frx":3CD3
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaddyInformation.frx":49AD
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaddyInformation.frx":5687
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaddyInformation.frx":6361
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaddyInformation.frx":703B
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaddyInformation.frx":7D15
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaddyInformation.frx":89EF
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaddyInformation.frx":96C9
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaddyInformation.frx":9FA3
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaddyInformation.frx":AC7D
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaddyInformation.frx":B957
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaddyInformation.frx":C631
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaddyInformation.frx":D30B
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaddyInformation.frx":DFE5
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCaddyInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim cn          As ADODB.Connection
Dim strPath     As String
Dim sCaddyName  As String
Dim tmp         As Long

Dim Arr, iPK, Filename

Private Sub BROWSER(cName, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If cName <> "" Then
            s = "SELECT TOP 1 tbl_Caddy_Information.* " & _
                " From tbl_Caddy_Information " & _
                " WHERE (CaddyLName + ',  ' + CaddyFName + '  ' + CaddyMName = '" & FORMATSQL(CStr(cName)) & "') " & _
                " ORDER BY CaddyLName + ',  ' + CaddyFName + '  ' + CaddyMName"
        Else
            s = "SELECT TOP 1 tbl_Caddy_Information.* " & _
                " From tbl_Caddy_Information " & _
                " ORDER BY CaddyLName + ',  ' + CaddyFName + '  ' + CaddyMName"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Caddy_Information.* " & _
            " From tbl_Caddy_Information " & _
            " ORDER BY CaddyLName + ',  ' + CaddyFName + '  ' + CaddyMName"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Caddy_Information.* " & _
            " From tbl_Caddy_Information " & _
            " WHERE (CaddyLName + ',  ' + CaddyFName + '  ' + CaddyMName < '" & FORMATSQL(CStr(cName)) & "') " & _
            " ORDER BY CaddyLName + ',  ' + CaddyFName + '  ' + CaddyMName DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Caddy_Information.* " & _
            " From tbl_Caddy_Information " & _
            " WHERE (CaddyLName + ',  ' + CaddyFName + '  ' + CaddyMName > '" & FORMATSQL(CStr(cName)) & "') " & _
            " ORDER BY CaddyLName + ',  ' + CaddyFName + '  ' + CaddyMName"
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Caddy_Information.* " & _
            " From tbl_Caddy_Information " & _
            " ORDER BY CaddyLName + ',  ' + CaddyFName + '  ' + CaddyMName DESC"
    Case Else: Exit Sub
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtCaddyNo.Text = rs!CaddyNo
    txtLastName.Text = rs!CaddyLName
    txtFirstName.Text = rs!CaddyFName
    txtMiddleName.Text = rs!CaddyMName
    txtContactNo.Text = rs!CaddyContactNo
    txtContactPerson.Text = rs!CaddyContactPerson
    txtContactPersonNo.Text = rs!CaddyContactPersonNo
    txtNickName.Text = rs!CaddyNickName
    txtRelation.Text = rs!CaddyContactPersonRelation
    imgPicture.Picture = LoadPicture(SHOW_IMAGES(rs!PK, 0, "Caddy Information"))
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    
    SaveSetting App.EXEName, "CaddyInformation", "CaddyInfo", rs!CaddyLName & ",  " & rs!CaddyFName & "  " & rs!CaddyMName
    
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If AccessRights("Caddy Information", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
txtCaddyNo.SetFocus
End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If AccessRights("Caddy Information", "Edit") = False Then
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
If AccessRights("Caddy Information", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If MsgBox("ARE YOU SURE IN DELETING THIS RECORD?                        ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_Member_Company_Corporate_Keys WHERE (KeyType = 1) AND (PrimaryKeys = " & Statusbar1.Panels(1).Text & ")"
ConnOmega.Execute "DELETE FROM tbl_Caddy_Information WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "CaddyInformation", "CaddyInfo", ""), "is_PAGEDOWN"
If Trim(txtCaddyNo.Text) = "" Then BROWSER GetSetting(App.EXEName, "CaddyInformation", "CaddyInfo", ""), "is_HOME"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If Trim(txtCaddyNo.Text) = "" Then MsgBox "Please Supply Caddy Number!                    ", vbCritical, "Error...": txtCaddyNo.SetFocus: Exit Sub
If Trim(txtLastName.Text) = "" Then MsgBox "Please Supply Last Name!                  ", vbCritical, "Error...": txtLastName.SetFocus: Exit Sub
If Trim(txtFirstName.Text) = "" Then MsgBox "Please Supply First Name!                  ", vbCritical, "Error...": txtFirstName.SetFocus: Exit Sub
If Trim(txtMiddleName.Text) = "" Then MsgBox "Please Supply Middle Name!                  ", vbCritical, "Error...": txtMiddleName.SetFocus: Exit Sub
If Trim(txtContactPerson.Text) = "" Then MsgBox "Please Supply Contact Person!                  ", vbCritical, "Error...": txtContactPerson.SetFocus: Exit Sub

On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    ConnOmega.Execute "INSERT INTO tbl_Caddy_Information " & _
                      " (CaddyNo, CaddyLName, CaddyFName, CaddyMName, CaddyContactNo, CaddyContactPerson, CaddyContactPersonNo, LastModified, CaddyNickName, " & _
                      " CaddyContactPersonRelation) " & _
                      " VALUES ('" & FORMATSQL(txtCaddyNo.Text) & "', '" & FORMATSQL(txtLastName.Text) & "', '" & FORMATSQL(txtFirstName.Text) & "', " & _
                      " '" & FORMATSQL(txtMiddleName.Text) & "', '" & FORMATSQL(txtContactNo.Text) & "', '" & FORMATSQL(txtContactPerson.Text) & "', " & _
                      " '" & FORMATSQL(txtContactPersonNo.Text) & "', '" & CStr(Now) & " - " & gbl_CompleteName & "', '" & FORMATSQL(Trim(txtNickName.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtRelation.Text)) & "')"
    If Trim(txtPicturePath.Text) <> "" Then
        iPK = 0
        s = "SELECT PK " & _
            " FROM tbl_Caddy_Information " & _
            " WHERE (CaddyLName = '" & FORMATSQL(txtLastName.Text) & "') " & _
            " AND (CaddyFName = '" & FORMATSQL(txtFirstName.Text) & "') " & _
            " AND (CaddyMName = '" & FORMATSQL(txtMiddleName.Text) & "') "
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            iPK = rs!PK
        End If
        rs.Close
        If CDbl(iPK) <> 0 Then
            SAVE_IMAGES iPK, 0, Trim(txtPicturePath.Text), "Caddy Information"
        End If
    End If
End If
If TRANSACTIONTYPE = is_EDITTING Then
    ConnOmega.Execute "UPDATE tbl_Caddy_Information " & _
                      " SET CaddyNo = '" & FORMATSQL(txtCaddyNo.Text) & "', " & _
                      " CaddyLName = '" & FORMATSQL(txtLastName.Text) & "', " & _
                      " CaddyFName = '" & FORMATSQL(txtFirstName.Text) & "', " & _
                      " CaddyMName = '" & FORMATSQL(txtMiddleName.Text) & "', " & _
                      " CaddyContactNo = '" & FORMATSQL(txtContactNo.Text) & "', " & _
                      " CaddyContactPerson = '" & FORMATSQL(txtContactPerson.Text) & "', " & _
                      " CaddyContactPersonNo = '" & FORMATSQL(txtContactPersonNo.Text) & "', " & _
                      " CaddyNickName = '" & FORMATSQL(Trim(txtNickName.Text)) & "', " & _
                      " CaddyContactPersonRelation = '" & FORMATSQL(Trim(txtRelation.Text)) & "', " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    If Trim(txtPicturePath.Text) <> "" Then
        iPK = Statusbar1.Panels(1).Text
        SAVE_IMAGES iPK, 0, Trim(txtPicturePath.Text), "Caddy Information"
    End If
End If

sCaddyName = txtLastName.Text & ",  " & txtFirstName.Text & "  " & txtMiddleName.Text
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER sCaddyName, "is_LOAD"

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
    BROWSER GetSetting(App.EXEName, "CaddyInformation", "CaddyInfo", ""), "is_LOAD"
    If Trim(txtCaddyNo.Text) = "" Then BROWSER GetSetting(App.EXEName, "CaddyInformation", "CaddyInfo", ""), "is_HOME"
End If
End Sub

Private Sub CLEARTEXT()
txtCaddyNo.Text = ""
txtLastName.Text = ""
txtFirstName.Text = ""
txtMiddleName.Text = ""
txtContactNo.Text = ""
txtContactPerson.Text = ""
txtContactPersonNo.Text = ""
txtNickName.Text = ""
txtRelation.Text = ""
imgPicture.Picture = LoadPicture("")
txtPicturePath.Text = ""
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtCaddyNo.Locked = bln
txtLastName.Locked = bln
txtFirstName.Locked = bln
txtMiddleName.Locked = bln
txtContactNo.Locked = bln
txtContactPerson.Locked = bln
txtContactPersonNo.Locked = bln
txtNickName.Locked = bln
txtRelation.Locked = bln
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
cmdCancel_Click
End Sub

Private Sub cmdCancel_Click()
picMain.Enabled = True
picToolbar.Enabled = True
picSearch.Visible = False
End Sub

Private Sub cmdOK_Click()
If lstResult.ListIndex = -1 Then Exit Sub
Arr = Split(lstResult.List(lstResult.ListIndex), " - ", -1, 1)
BROWSER CStr(Arr(1)), "is_LOAD"
cmdCancel_Click
End Sub

Private Sub Command1_Click()

CommonDialog1.DialogTitle = "OPEN FILE"
CommonDialog1.Filename = ""
CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowOpen
strPath = CommonDialog1.Filename

Set cn = New ADODB.Connection
cn.Provider = "Microsoft.Jet.OLEDB.4.0"
cn.ConnectionString = _
    "Data Source= " & Trim(strPath) & ";" & _
    "Extended Properties=Excel 8.0;"
cn.CursorLocation = adUseClient
If cn.State = adStateOpen Then cn.Close
cn.Open

Set rs = New ADODB.Recordset
If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT * FROM [CaddyInfo$] ", cn, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    If Trim(IIf(IsNull(rs!CaddyLName), "", rs!CaddyLName)) <> "" Then
        ConnOmega.Execute "INSERT INTO tbl_Caddy_Information " & _
                          " (CaddyNo, CaddyLName, CaddyFName, CaddyMName, CaddyNickName, CaddyContactNo, " & _
                          " CaddyContactPerson, CaddyContactPersonRelation, CaddyContactPersonNo) " & _
                          " VALUES (" & rs!CaddyNo & ", '" & FORMATSQL(rs!CaddyLName) & "', '" & FORMATSQL(IIf(IsNull(rs!CaddyFName), "", rs!CaddyFName)) & "', " & _
                          " '" & FORMATSQL(IIf(IsNull(rs!CaddyMName), "", rs!CaddyMName)) & "', '" & FORMATSQL(IIf(IsNull(rs!CaddyNickName), "", rs!CaddyNickName)) & "', " & _
                          " '" & FORMATSQL(IIf(IsNull(rs!CaddyContactNo), "", rs!CaddyContactNo)) & "', '" & FORMATSQL(IIf(IsNull(rs!CaddyContactPerson), "", rs!CaddyContactPerson)) & "', " & _
                          " '" & FORMATSQL(IIf(IsNull(rs!CaddyContactPersonRelation), "", rs!CaddyContactPersonRelation)) & "', '" & FORMATSQL(Replace(IIf(IsNull(rs!CaddyContactPersonNo), "", rs!CaddyContactPersonNo), "'", "")) & "')"
    End If
    rs.MoveNext
Wend
rs.Close
cn.Close
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
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "CaddyInformation", "CaddyInfo", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "CaddyInformation", "CaddyInfo", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "CaddyInformation", "CaddyInfo", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "CaddyInformation", "CaddyInfo", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
'Me.Caption = "Caddy Information"

CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "CaddyInformation", "CaddyInfo", ""), "is_LOAD"
If Trim(txtCaddyNo.Text) = "" Then BROWSER GetSetting(App.EXEName, "CaddyInformation", "CaddyInfo", ""), "is_HOME"


tmp = SetWindowLong(txtCaddyNo.hwnd, GWL_STYLE, GetWindowLong(txtCaddyNo.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtLastName.hwnd, GWL_STYLE, GetWindowLong(txtLastName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtFirstName.hwnd, GWL_STYLE, GetWindowLong(txtFirstName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtMiddleName.hwnd, GWL_STYLE, GetWindowLong(txtMiddleName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtContactNo.hwnd, GWL_STYLE, GetWindowLong(txtContactNo.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtContactPerson.hwnd, GWL_STYLE, GetWindowLong(txtContactPerson.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtContactPersonNo.hwnd, GWL_STYLE, GetWindowLong(txtContactPersonNo.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picSearch.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub imgPicture_DblClick()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    MainForm.CommonDialog1.CancelError = True
    On Error GoTo ErrorHandler
'    Mainform.CommonDialog1.Filter = "Image Files (*.jpg)|*.jpg"
'    Mainform.CommonDialog1.Filter = "JPG|*.JPG;*.JPEG;*.JPE|BMP|*.BMP;*.RLE;*.DIB|GIF|*.GIF|PNG|*.PNG|TIFF|*.TIF;*.TIFF"
    MainForm.CommonDialog1.Filter = "Image Files|*.JPG;*.JPEG;*.JPE;*.BMP;*.RLE;*.DIB;*.GIF;*.PNG;*.TIF;*.TIFF"
    MainForm.CommonDialog1.ShowOpen
    Filename = Trim(MainForm.CommonDialog1.Filename)
'    If ((FileLen(Filename) \ 1024) + 1) > 50 Then
'        MsgBox "Image is too large please reduce the size to 50kb or below!          ", vbCritical, "Error..."
'        Exit Sub
'    End If
'    MsgBox CDbl(IMAGEFILESIZE(Date))
    If ((FileLen(Filename) \ 1024) + 1) > CDbl(IMAGEFILESIZE(Date)) Then
        MsgBox "Image is too large please reduce the size to " & IMAGEFILESIZE(Date) & "kb or below!          ", vbCritical, "Error..."
        Exit Sub
    End If
    txtPicturePath.Text = Filename
    imgPicture.Picture = LoadPicture(Filename)
End If
Exit Sub
ErrorHandler:
Exit Sub
End Sub

Private Sub lstResult_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
        Case "Refresh"
            'ToDo: Add 'Refresh' button code.
            MsgBox "Add 'Refresh' button code."
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "CaddyInformation", "CaddyInfo", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "CaddyInformation", "CaddyInfo", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "CaddyInformation", "CaddyInfo", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "CaddyInformation", "CaddyInfo", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstResult.Clear: Exit Sub
lstResult.Clear
s = "SELECT tbl_Caddy_Information.* " & _
    " From tbl_Caddy_Information " & _
    " WHERE (CaddyLName LIKE '" & FORMATSQL(txtSearch.Text) & "%') " & _
    " ORDER BY CaddyLName, CaddyFName, CaddyMName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResult.AddItem rs!CaddyNo & " - " & rs!CaddyLName & ",  " & rs!CaddyFName & "  " & rs!CaddyMName
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
