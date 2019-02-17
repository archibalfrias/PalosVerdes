VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelPosition 
   Appearance      =   0  'Flat
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9780
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
   ScaleHeight     =   3510
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picSearch 
      Height          =   3135
      Left            =   2520
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5530
      BackColor       =   15396057
      Begin VB.CommandButton cmdOKSearch 
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
         Left            =   960
         Picture         =   "frmPersonnelPosition.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2520
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelSearch 
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
         Left            =   2640
         Picture         =   "frmPersonnelPosition.frx":0672
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2520
         Width           =   1560
      End
      Begin VB.TextBox txtSearchSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   4815
      End
      Begin VB.ListBox lstResultSearch 
         Height          =   1620
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   4815
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   45
         TabIndex        =   17
         Top             =   45
         Width           =   4965
         _ExtentX        =   8758
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
         Icon            =   "frmPersonnelPosition.frx":0DCE
         ShadowVisible   =   0   'False
      End
   End
   Begin VB.PictureBox picBody 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   1200
      ScaleHeight     =   1095
      ScaleWidth      =   7215
      TabIndex        =   4
      Top             =   1440
      Width           =   7215
      Begin VB.TextBox txtPostCode_1 
         Height          =   315
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtPostCode 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   0
         Width           =   975
      End
      Begin VB.TextBox txtPostName 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   5775
      End
      Begin VB.ComboBox cmbLevel 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   720
         Width           =   5775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "POSITION CODE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "POSITION NAME"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "LEVEL"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   720
         Width           =   1575
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
         MouseIcon       =   "frmPersonnelPosition.frx":1368
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   9900
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   2
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmPersonnelPosition.frx":1682
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8880
      Top             =   1680
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
            Picture         =   "frmPersonnelPosition.frx":1D95
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPosition.frx":2A6F
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPosition.frx":3749
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPosition.frx":4423
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPosition.frx":50FD
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPosition.frx":5DD7
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPosition.frx":6AB1
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPosition.frx":778B
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPosition.frx":8465
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPosition.frx":8D3F
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPosition.frx":9A19
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPosition.frx":A6F3
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPosition.frx":B3CD
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPosition.frx":C0A7
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelPosition.frx":CD81
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   3210
      Width           =   9780
      _ExtentX        =   17251
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
Attribute VB_Name = "frmPersonnelPosition"
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

Dim iLevelKey As Long


Private Sub BROWSER(strName, is_Action As String)

Select Case is_Action
    Case "is_LOAD"
        If strName <> "" Then
            s = "SELECT TOP (1) dbo.tbl_Personnel_Position.PK, dbo.tbl_Personnel_Position.PositionCode, " & _
                " dbo.tbl_Personnel_Position.PositionName, dbo.tbl_Personnel_Position.PositionLevel, " & _
                " dbo.tbl_Personnel_Position_Level.LevelName, dbo.tbl_Personnel_Position.LastModified " & _
                " FROM  dbo.tbl_Personnel_Position LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Position_Level ON dbo.tbl_Personnel_Position.PositionLevel = dbo.tbl_Personnel_Position_Level.PK " & _
                " WHERE (dbo.tbl_Personnel_Position.PositionName = '" & FORMATSQL(CStr(strName)) & "') " & _
                " ORDER BY dbo.tbl_Personnel_Position.PositionName"
        Else
            s = "SELECT TOP (1) dbo.tbl_Personnel_Position.PK, dbo.tbl_Personnel_Position.PositionCode, " & _
                " dbo.tbl_Personnel_Position.PositionName, dbo.tbl_Personnel_Position.PositionLevel, " & _
                " dbo.tbl_Personnel_Position_Level.LevelName, dbo.tbl_Personnel_Position.LastModified " & _
                " FROM  dbo.tbl_Personnel_Position LEFT OUTER JOIN " & _
                " dbo.tbl_Personnel_Position_Level ON dbo.tbl_Personnel_Position.PositionLevel = dbo.tbl_Personnel_Position_Level.PK " & _
                " ORDER BY dbo.tbl_Personnel_Position.PositionName"
        End If
    Case "is_HOME"
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Position.PK, dbo.tbl_Personnel_Position.PositionCode, " & _
            " dbo.tbl_Personnel_Position.PositionName, dbo.tbl_Personnel_Position.PositionLevel, " & _
            " dbo.tbl_Personnel_Position_Level.LevelName, dbo.tbl_Personnel_Position.LastModified " & _
            " FROM  dbo.tbl_Personnel_Position LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Position_Level ON dbo.tbl_Personnel_Position.PositionLevel = dbo.tbl_Personnel_Position_Level.PK " & _
            " ORDER BY dbo.tbl_Personnel_Position.PositionName"
    Case "is_PAGEUP"
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Position.PK, dbo.tbl_Personnel_Position.PositionCode, " & _
            " dbo.tbl_Personnel_Position.PositionName, dbo.tbl_Personnel_Position.PositionLevel, " & _
            " dbo.tbl_Personnel_Position_Level.LevelName, dbo.tbl_Personnel_Position.LastModified " & _
            " FROM  dbo.tbl_Personnel_Position LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Position_Level ON dbo.tbl_Personnel_Position.PositionLevel = dbo.tbl_Personnel_Position_Level.PK " & _
            " WHERE (dbo.tbl_Personnel_Position.PositionName < '" & FORMATSQL(CStr(strName)) & "') " & _
            " ORDER BY dbo.tbl_Personnel_Position.PositionName DESC"
    Case "is_PAGEDOWN"
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Position.PK, dbo.tbl_Personnel_Position.PositionCode, " & _
            " dbo.tbl_Personnel_Position.PositionName, dbo.tbl_Personnel_Position.PositionLevel, " & _
            " dbo.tbl_Personnel_Position_Level.LevelName, dbo.tbl_Personnel_Position.LastModified " & _
            " FROM  dbo.tbl_Personnel_Position LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Position_Level ON dbo.tbl_Personnel_Position.PositionLevel = dbo.tbl_Personnel_Position_Level.PK " & _
            " WHERE (dbo.tbl_Personnel_Position.PositionName > '" & FORMATSQL(CStr(strName)) & "') " & _
            " ORDER BY dbo.tbl_Personnel_Position.PositionName"
    Case "is_END"
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) dbo.tbl_Personnel_Position.PK, dbo.tbl_Personnel_Position.PositionCode, " & _
            " dbo.tbl_Personnel_Position.PositionName, dbo.tbl_Personnel_Position.PositionLevel, " & _
            " dbo.tbl_Personnel_Position_Level.LevelName, dbo.tbl_Personnel_Position.LastModified " & _
            " FROM  dbo.tbl_Personnel_Position LEFT OUTER JOIN " & _
            " dbo.tbl_Personnel_Position_Level ON dbo.tbl_Personnel_Position.PositionLevel = dbo.tbl_Personnel_Position_Level.PK " & _
            " ORDER BY dbo.tbl_Personnel_Position.PositionName DESC"
    Case "is_FIND"
        s = "SELECT TOP (1) dbo.tbl_Personnel_Position.PK, dbo.tbl_Personnel_Position.PositionCode, " & _
        " dbo.tbl_Personnel_Position.PositionName, dbo.tbl_Personnel_Position.PositionLevel, " & _
        " dbo.tbl_Personnel_Position_Level.LevelName, dbo.tbl_Personnel_Position.LastModified " & _
        " FROM  dbo.tbl_Personnel_Position LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Position_Level ON dbo.tbl_Personnel_Position.PositionLevel = dbo.tbl_Personnel_Position_Level.PK " & _
        " WHERE (dbo.tbl_Personnel_Position.PK = " & strName & ") " & _
        " ORDER BY dbo.tbl_Personnel_Position.PositionName DESC"
    Case Else: Exit Sub
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iLevelKey = rs!PositionLevel
    txtPostCode.Text = rs!PositionCode
    txtPostName.Text = rs!PositionName
    cmbLevel.Text = rs!LevelName
    StatusBar1.Panels(1).Text = rs!PK
    StatusBar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", "LAST MODIFIED BY : " & rs!LastModified)
    SaveSetting App.EXEName, "PersonnelPosition", "PersonnelPost", rs!PositionName
End If
rs.Close
End Sub

Private Function AUTOCODE() As Long

s = "SELECT Max(PositionCode) AS Code" & _
    " FROM tbl_Personnel_Position"
rs.Open s, ConnOmega
AUTOCODE = CLng(IIf(IsNull(rs!Code), 0, rs!Code)) + 1
rs.Close
End Function

Private Sub PRESS_INSERT()
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If AccessRights("Personnel Position", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
TRANSACTIONTYPE = is_ADDING
TOOLBARFUNC 2
CLEARTEXT
LOCKTEXT False
'Me.Caption = "POSITION - NEW"
txtPostCode.Text = Format(AUTOCODE, "00#")
txtPostName.SetFocus
End Sub

Private Sub PRESS_F2()
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If StatusBar1.Panels(1).Text = "" Then Exit Sub
    
If AccessRights("Personnel Position", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
        
TRANSACTIONTYPE = is_EDITTING
TOOLBARFUNC 2
LOCKTEXT False
'Me.Caption = "POSITION - EDIT"
txtPostName.SetFocus
        
End Sub

Private Sub PRESS_DELETE()
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If StatusBar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("Personnel Position", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If MsgBox("ARE YOU SURE TO DELETE THIS RECORD?      ", vbInformation + vbYesNo, "CONFIRMATION") = vbNo Then Exit Sub
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_Personnel_Position WHERE (PK = " & StatusBar1.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_PAGEDOWN"
If Trim(txtPostCode.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_HOME"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
Exit Sub
End Sub

Private Sub PRESS_F5()
If picSearch.Visible = True Then Exit Sub
If Trim(txtPostName.Text) = "" Then MsgBox "Please supply position name!                    ", vbCritical, "Error...": txtPostName.SetFocus: Exit Sub
If iLevelKey = 0 Then MsgBox "Please select position level!                     ", vbCritical, "Error...": cmbLevel.SetFocus: Exit Sub
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Position" & _
                      " (PositionCode, PositionName, " & _
                      " PositionLevel, LastModified)" & _
                      " VALUES('" & Trim(txtPostCode.Text) & "', " & _
                      " '" & FORMATSQL(Trim(txtPostName.Text)) & "', " & _
                      " " & iLevelKey & ", " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
    BROWSER FORMATSQL(Trim(txtPostName.Text)), "is_LOAD"
    TRANSACTIONTYPE = is_REFRESH
    TOOLBARFUNC 1
    LOCKTEXT True
ElseIf TRANSACTIONTYPE = is_EDITTING Then
    ConnOmega.Execute "UPDATE tbl_Personnel_Position" & _
                      " SET PositionName = '" & FORMATSQL(Trim(txtPostName.Text)) & "', " & _
                      " PositionLevel = " & iLevelKey & ", " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "'" & _
                      " WHERE (PK = " & StatusBar1.Panels(1).Text & ")"
    BROWSER FORMATSQL(Trim(txtPostName.Text)), "is_LOAD"
    TRANSACTIONTYPE = is_REFRESH
    TOOLBARFUNC 1
    LOCKTEXT True
End If
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
Exit Sub
End Sub

Private Sub PRESS_F6()
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
picSearch.ZOrder 0
txtSearchSearch.Text = ""
picBody.Enabled = False
picToolbar.Enabled = False
picSearch.Visible = True
txtSearchSearch.SetFocus
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSearch.Visible = True Then cmdCancelSearch_Click: Exit Sub
    Unload Me
Else
    BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_LOAD"
    If Trim(txtPostName.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_HOME"
    TRANSACTIONTYPE = is_REFRESH
    TOOLBARFUNC 1
    LOCKTEXT True
    txtPostCode_1.Visible = False
    'Me.Caption = "POSITION - BROWSE"
End If
End Sub

Private Function FIND_CODE(strCode) As Long
s = "SELECT PK" & _
    " From tbl_Personnel_Position  " & _
    " WHERE (PositionCode='" & strCode & "')"
rs.Open s, ConnOmega
If Not rs.EOF Then
    FIND_CODE = IIf(IsNull(rs!PK), 0, rs!PK)
End If
rs.Close
End Function


Public Sub CLEARTEXT()
iLevelKey = 0
txtPostCode.Text = ""
txtPostName.Text = ""
cmbLevel.Text = ""
cmbLevel.ListIndex = -1
StatusBar1.Panels(1).Text = ""
StatusBar1.Panels(2).Text = ""
End Sub

Private Function LOCKTEXT(bln As Boolean)
txtPostCode.Locked = True
txtPostName.Locked = bln
cmbLevel.Locked = bln
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


Private Sub b8TitleBar1_CLoseClick()
cmdCancelSearch_Click
End Sub

Private Sub cmbLevel_Click()
If cmbLevel.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    iLevelKey = cmbLevel.ItemData(cmbLevel.ListIndex)
End If
End Sub

Private Sub cmdCancelSearch_Click()
picBody.Enabled = True
picToolbar.Enabled = True
picSearch.Visible = False
End Sub

Private Sub cmdOKSearch_Click()
If lstResultSearch.ListIndex = -1 Then Exit Sub
BROWSER lstResultSearch.ItemData(lstResultSearch.ListIndex), "is_FIND"
cmdCancelSearch_Click
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
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_END"
    Case vbKeyEscape:   PRESS_ESCAPE
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
POPULATE_COMBO "PK", "LevelName", "tbl_Personnel_Position_Level", "PK", cmbLevel
CLEARTEXT
LOCKTEXT True
TRANSACTIONTYPE = is_REFRESH
TOOLBARFUNC 1
BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_LOAD"
If Trim(txtPostName.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_HOME"

tmp = SetWindowLong(txtPostCode.hwnd, GWL_STYLE, GetWindowLong(txtPostCode.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPostName.hwnd, GWL_STYLE, GetWindowLong(txtPostName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearchSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearchSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picSearch.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub


Private Sub lstResultSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKSearch_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":           PRESS_INSERT
    Case "Edit":          PRESS_F2
    Case "Delete":        PRESS_DELETE
    Case "First"
        Select Case Toolbar1.Buttons(7).Caption
            Case "Save":  PRESS_F5
            Case "First": BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_HOME"
        End Select
    Case "Back"
        Select Case Toolbar1.Buttons(9).Caption
            Case "Undo":  PRESS_ESCAPE
            Case "Back":  BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_PAGEUP"
        End Select
    Case "Next":          BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_PAGEDOWN"
    Case "Last":          BROWSER GetSetting(App.EXEName, "PersonnelPosition", "PersonnelPost", ""), "is_END"
    Case "Find":          PRESS_F6
    Case "Close":         PRESS_ESCAPE
End Select
End Sub

Private Sub txtPostCode_1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If TRANSACTIONTYPE = is_FINDING Then
        txtPostCode_1.Text = Format(txtPostCode_1.Text, "00#")
        If FIND_CODE(Format(txtPostCode_1.Text, "00#")) <> 0 Then
            BROWSER FIND_CODE(Format(txtPostCode_1.Text, "00#")), "is_FIND"
            TRANSACTIONTYPE = is_REFRESH
            TOOLBARFUNC 1
            txtPostCode_1.Visible = False
        Else
            MsgBox "UNABLE TO FIND '" & Format(txtPostCode_1.Text, "00#") & "' IN THE DATABASE!      ", vbCritical, "ERROR..."
            txtPostCode_1.SetFocus
            HTEXT txtPostCode_1
        End If
    End If
End If
End Sub

Private Sub txtSearchSearch_Change()
If Trim(txtSearchSearch.Text) = "" Then lstResultSearch.Clear: Exit Sub
lstResultSearch.Clear
s = "SELECT PK, PositionName " & _
    " From dbo.tbl_Personnel_Position " & _
    " WHERE (PositionName LIKE '" & FORMATSQL(Trim(txtSearchSearch.Text)) & "%') " & _
    " ORDER BY PositionName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultSearch.AddItem rs!PositionName
    lstResultSearch.ItemData(lstResultSearch.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResultSearch.ListCount Then lstResultSearch.ListIndex = 0
End Sub

Private Sub txtSearchSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultSearch.SetFocus
End Sub
