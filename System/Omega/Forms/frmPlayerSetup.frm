VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlayerSetup 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
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
   Icon            =   "frmPlayerSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   33
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   34
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
         MouseIcon       =   "frmPlayerSetup.frx":08CA
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
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   960
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   1560
      ScaleHeight     =   2295
      ScaleWidth      =   5535
      TabIndex        =   6
      Top             =   1920
      Width           =   5535
      Begin VB.TextBox txtClassIndex 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2280
         TabIndex        =   32
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtIndex 
         Height          =   315
         Left            =   1200
         TabIndex        =   30
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtAllowedTeam 
         Height          =   315
         Left            =   1200
         TabIndex        =   28
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtClass 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2280
         TabIndex        =   12
         Top             =   1200
         Width           =   495
      End
      Begin VB.ComboBox cmbGender 
         Height          =   315
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtHandicap 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtMName 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtFName 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox txtLName 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   120
         Width           =   4215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Index"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Allowed Team"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Handicap"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1695
      Left            =   4920
      TabIndex        =   29
      Top             =   6120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2990
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
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
         Text            =   "HDCP"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "PlayerKey"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Height          =   495
      Left            =   840
      ScaleHeight     =   435
      ScaleWidth      =   3195
      TabIndex        =   20
      Top             =   7080
      Width           =   3255
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   720
      TabIndex        =   19
      Top             =   6600
      Width           =   3855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1920
      Top             =   6120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   840
      TabIndex        =   18
      Top             =   6120
      Width           =   975
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   4545
      Width           =   9015
      _ExtentX        =   15901
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2640
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   735
      Left            =   1560
      ScaleHeight     =   735
      ScaleWidth      =   5535
      TabIndex        =   13
      Top             =   1200
      Width           =   5535
      Begin VB.TextBox txtDate 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1200
         TabIndex        =   15
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox txtTournament 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1200
         TabIndex        =   14
         Top             =   0
         Width           =   4215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tournament"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   0
         Width           =   1335
      End
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
            Picture         =   "frmPlayerSetup.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlayerSetup.frx":18BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlayerSetup.frx":2598
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlayerSetup.frx":3272
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlayerSetup.frx":3F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlayerSetup.frx":4C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlayerSetup.frx":5900
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlayerSetup.frx":65DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlayerSetup.frx":72B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlayerSetup.frx":7B8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlayerSetup.frx":8868
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlayerSetup.frx":9542
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlayerSetup.frx":A21C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlayerSetup.frx":AEF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlayerSetup.frx":BBD0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RPVGCC.b8Container picSearch 
      Height          =   3735
      Left            =   2160
      TabIndex        =   21
      Top             =   240
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6588
      BackColor       =   15396057
      Begin VB.ListBox lstResult 
         Height          =   2205
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   4095
      End
      Begin VB.CommandButton cmdCancel 
         Height          =   480
         Left            =   2235
         Picture         =   "frmPlayerSetup.frx":C8AA
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3120
         Width           =   1560
      End
      Begin VB.CommandButton cmdOK 
         Height          =   480
         Left            =   480
         Picture         =   "frmPlayerSetup.frx":D006
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3120
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   40
         TabIndex        =   24
         Top             =   40
         Width           =   4245
         _ExtentX        =   7488
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
         Icon            =   "frmPlayerSetup.frx":D678
      End
   End
End
Attribute VB_Name = "frmPlayerSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public TournamentKey As Double

Public TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim tmp As Long

Dim cn As ADODB.Connection

Dim dblrpv, dblmatina, dblapo, dblOthers, strTeamName, strTeamHDCP, strTeamArr, j, strTeamID, _
iPK, strName, dblTeamHDCP, dblTeamHDCPInx, intTeamKey, strPath, i, Array1, dblHandicap, strClass, x, _
dblTeamHDCPTot, dblTeamHDCPIndx, strTeamClass, intPlayerKey


Private Function BROWSER(strName, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If strName <> "" Then
            s = "SELECT TOP 1 tbl_Scoring_PlayerName.* " & _
                " From tbl_Scoring_PlayerName " & _
                " WHERE (TournamentKey = " & TournamentKey & ") " & _
                " AND (LastName + ',  ' + FirstName + '  ' + MiddleName = '" & FORMATSQL(CStr(strName)) & "') " & _
                " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
        Else
            s = "SELECT TOP 1 tbl_Scoring_PlayerName.* " & _
                " From tbl_Scoring_PlayerName " & _
                " WHERE (TournamentKey = " & TournamentKey & ") " & _
                " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
         s = "SELECT TOP 1 tbl_Scoring_PlayerName.* " & _
            " From tbl_Scoring_PlayerName " & _
            " WHERE (TournamentKey = " & TournamentKey & ") " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Scoring_PlayerName.* " & _
            " From tbl_Scoring_PlayerName " & _
            " WHERE (TournamentKey = " & TournamentKey & ") " & _
            " AND (LastName + ',  ' + FirstName + '  ' + MiddleName < '" & FORMATSQL(CStr(strName)) & "') " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Scoring_PlayerName.* " & _
            " From tbl_Scoring_PlayerName " & _
            " WHERE (TournamentKey = " & TournamentKey & ") " & _
            " AND (LastName + ',  ' + FirstName + '  ' + MiddleName > '" & FORMATSQL(CStr(strName)) & "') " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Scoring_PlayerName.* " & _
            " From tbl_Scoring_PlayerName " & _
            " WHERE (TournamentKey = " & TournamentKey & ") " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName DESC"
    Case "is_FIND"
        s = "SELECT TOP 1 tbl_Scoring_PlayerName.* " & _
            " From tbl_Scoring_PlayerName " & _
            " WHERE (PK = " & strName & ") " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName DESC"
    Case Else: Exit Function
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtLName.Text = Trim(rs!LastName)
    txtFName.Text = Trim(rs!FirstName)
    txtMName.Text = Trim(rs!MiddleName)
    txtHandicap.Text = IIf(IsNull(rs!Handicap), 0, rs!Handicap)
    txtIndex.Text = IIf(IsNull(rs!iIndex), 0, rs!iIndex) 'rs!iIndex
    txtClass.Text = ""
    'If Trim(IIf(IsNull(rs!Class), "", rs!Class)) = "" Then
    t = "SELECT Class " & _
        " From tbl_Scoring_TournamentInfo_Class " & _
        " WHERE (TournamentKey = " & TournamentKey & ") " & _
        " AND (HFrom <= " & IIf(IsNull(rs!Handicap), 0, rs!Handicap) & ") " & _
        " AND (HTo >= " & IIf(IsNull(rs!Handicap), 0, rs!Handicap) & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtClass.Text = rt!Class
    Else
        txtClass.Text = ""
    End If
    rt.Close
    'Else
    'txtClass.Text = rs!Class
    'End If
    
    txtClassIndex.Text = ""
    t = "SELECT Class " & _
        " From tbl_Scoring_TournamentInfo_Index " & _
        " WHERE (TournamentKey = " & TournamentKey & ") " & _
        " AND (HFrom <= " & IIf(IsNull(rs!iIndex), 0, rs!iIndex) & ") " & _
        " AND (HTo >= " & IIf(IsNull(rs!iIndex), 0, rs!iIndex) & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtClassIndex.Text = rt!Class
    Else
        txtClassIndex.Text = ""
    End If
    rt.Close
        
    txtAllowedTeam.Text = rs!AllowedTeam
    cmbGender.ListIndex = rs!Gender - 1
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    
    SaveSetting App.EXEName, "PlayerInfo", "PlayerSetup", rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
    
End If
rs.Close
End Function

Private Function PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If picSearch.Visible = True Then Exit Function
If AccessRights("Scoring Player Information", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
If CHECK_TOURNAMENT_STATUS(TournamentKey) <> 0 Then MsgBox "Tournament was already locked!               ", vbCritical, "Error...": Exit Function
PopupMenu MainFormPopupF.mnuPlayerAdd, , Toolbar1.Buttons(1).Left, Toolbar1.Buttons(1).Top + Toolbar1.Buttons(1).Height

'CLEARTEXT
'LOCKTEXT False
'TOOLBARFUNC 2
'TRANSACTIONTYPE = is_ADDING
'txtLName.SetFocus
End Function

Private Function PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If Statusbar1.Panels(1).Text = "" Then Exit Function
If AccessRights("Scoring Player Information", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
If CHECK_TOURNAMENT_STATUS(TournamentKey) <> 0 Then MsgBox "Tournament was already locked!               ", vbCritical, "Error...": Exit Function
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_EDITTING
End Function

Private Function PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If Statusbar1.Panels(1).Text = "" Then Exit Function
If picSearch.Visible = True Then Exit Function
If AccessRights("Scoring Player Information", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
If CHECK_TOURNAMENT_STATUS(TournamentKey) <> 0 Then MsgBox "Tournament was already locked!               ", vbCritical, "Error...": Exit Function
If MsgBox("ARE YOU SURE IN DELETING THIS RECORD?                ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Function
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_Scoring_PlayerName WHERE (PK =" & Statusbar1.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "PlayerInfo", "PlayerSetup", ""), "is_PAGEDOWN"
If Trim(txtLName.Text) = "" Then BROWSER GetSetting(App.EXEName, "PlayerInfo", "PlayerSetup", ""), "is_HOME"
Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Function
End Function

Private Function PRESS_F5()
If picSearch.Visible = True Then Exit Function
If Trim(txtLName.Text) = "" Then MsgBox "Please Supply Last Name!             ", vbCritical, "Error...": txtLName.SetFocus: Exit Function
If Trim(txtFName.Text) = "" Then MsgBox "Please Supply First Name!             ", vbCritical, "Error...": txtFName.SetFocus: Exit Function
If Trim(txtMName.Text) = "" Then
    If MsgBox("Continue Saving without Middle Name?                     ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Function
End If
If cmbGender.ListIndex = -1 Then MsgBox "Please Supply Gender!                   ", vbCritical, "Error...": cmbGender.SetFocus: Exit Function
If ScoringType <> 3 Then
    If RETURNTEXTVALUE(txtHandicap) < 0 Then MsgBox "Please Supply Handicap!               ", vbCritical, "Error...": txtHandicap.SetFocus: Exit Function
    If Trim(txtClass.Text) = "" Then MsgBox "Handicap has no Class Category!               ", vbCritical, "Error...": txtHandicap.SetFocus: Exit Function
    If RETURNTEXTVALUE(txtHandicap) > TopHandicap Then MsgBox "Invalid Handicap!                 ", vbCritical, "Error...": txtHandicap.SetFocus: HTEXT txtHandicap: Exit Function
End If

On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    ConnOmega.Execute "INSERT INTO tbl_Scoring_PlayerName " & _
                      " (TournamentKey, LastName, FirstName, MiddleName, " & _
                      " HandiCap, Class, Gender, LastModified, AllowedTeam, iIndex) " & _
                      " VALUES (" & TournamentKey & ", '" & FORMATSQL(Trim(txtLName.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtFName.Text)) & "', '" & FORMATSQL(Trim(txtMName.Text)) & "', " & _
                      " " & RETURNTEXTVALUE(txtHandicap) & ", '" & Trim(txtClass.Text) & "', " & _
                      " " & cmbGender.ListIndex + 1 & ", '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                      " " & RETURNTEXTVALUE(txtAllowedTeam) & ", " & RoundOffIndex(RETURNTEXTVALUE(txtIndex)) & ")"
    strName = Trim(txtLName.Text) & ",  " & Trim(txtFName.Text) & "  " & Trim(txtMName.Text)
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER strName, "is_LOAD"
    iPK = Statusbar1.Panels(1).Text
ElseIf TRANSACTIONTYPE = is_EDITTING Then
    iPK = Statusbar1.Panels(1).Text
    ConnOmega.Execute "UPDATE tbl_Scoring_PlayerName " & _
                      " SET TournamentKey = " & TournamentKey & ", " & _
                      " LastName = '" & FORMATSQL(Trim(txtLName.Text)) & "', " & _
                      " FirstName = '" & FORMATSQL(Trim(txtFName.Text)) & "', " & _
                      " MiddleName = '" & FORMATSQL(Trim(txtMName.Text)) & "', " & _
                      " HandiCap = " & RETURNTEXTVALUE(txtHandicap) & ", " & _
                      " Class = '" & Trim(txtClass.Text) & "', " & _
                      " Gender = " & cmbGender.ListIndex + 1 & ", " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                      " AllowedTeam = " & RETURNTEXTVALUE(txtAllowedTeam) & ", " & _
                      " iIndex = " & RoundOffIndex(RETURNTEXTVALUE(txtIndex)) & "  " & _
                      " WHERE (PK = " & iPK & ")"
    
    intTeamKey = 0
    s = "SELECT TeamKey " & _
        " From tbl_Scoring_Team_Detail " & _
        " WHERE (PlayerKey = " & Statusbar1.Panels(1).Text & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        intTeamKey = rs!TeamKey
    End If
    rs.Close
    
    If CDbl(intTeamKey) > 0 Then
        dblTeamHDCP = 0: dblTeamHDCPInx = 0
        If TeamAverage = 2 Then
            If TeamDivisorOrder = 0 Then
                s = "SELECT TOP " & HandicapDivisor & " tbl_Scoring_PlayerName.iIndex as HandiCap " & _
                    " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                    " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                    " Where (tbl_Scoring_Team_Detail.TeamKey = " & intTeamKey & ") " & _
                    " ORDER BY tbl_Scoring_PlayerName.iIndex "
            ElseIf TeamDivisorOrder = 1 Then
                s = "SELECT TOP " & HandicapDivisor & " tbl_Scoring_PlayerName.iIndex as HandiCap " & _
                    " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                    " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                    " Where (tbl_Scoring_Team_Detail.TeamKey = " & intTeamKey & ") " & _
                    " ORDER BY tbl_Scoring_PlayerName.iIndex DESC"
            End If
            If rs.State = adStateOpen Then rs.Close
            rs.Open s, ConnOmega
            While Not rs.EOF
                dblTeamHDCP = dblTeamHDCP + CDbl(rs!Handicap)
                rs.MoveNext
            Wend
            rs.Close
            
            dblTeamHDCPInx = CDbl(dblTeamHDCP) / CDbl(HandicapDivisor)
            dblTeamHDCPInx = RoundOffIndex(CDbl(dblTeamHDCPInx))
            ConnOmega.Execute "UPDATE tbl_Scoring_Team " & _
                              " SET TeamIndex = " & CDbl(dblTeamHDCPInx) & " " & _
                              " Where (PK = " & intTeamKey & ")"
        Else
            If TeamDivisorOrder = 0 Then
                s = "SELECT TOP " & HandicapDivisor & " tbl_Scoring_PlayerName.HandiCap " & _
                    " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                    " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                    " Where (tbl_Scoring_Team_Detail.TeamKey = " & intTeamKey & ") " & _
                    " ORDER BY tbl_Scoring_PlayerName.HandiCap "
            ElseIf TeamDivisorOrder = 1 Then
                s = "SELECT TOP " & HandicapDivisor & " tbl_Scoring_PlayerName.HandiCap " & _
                    " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                    " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                    " Where (tbl_Scoring_Team_Detail.TeamKey = " & intTeamKey & ") " & _
                    " ORDER BY tbl_Scoring_PlayerName.HandiCap DESC"
            End If
            If rs.State = adStateOpen Then rs.Close
            rs.Open s, ConnOmega
            While Not rs.EOF
                dblTeamHDCP = dblTeamHDCP + CDbl(rs!Handicap)
                rs.MoveNext
            Wend
            rs.Close
            
            dblTeamHDCPInx = Format(CDbl(dblTeamHDCP) / CDbl(HandicapDivisor), "#0.0")
            
            ConnOmega.Execute "UPDATE tbl_Scoring_Team " & _
                              " SET TeamHDCP = " & CDbl(dblTeamHDCPInx) & " " & _
                              " Where (PK = " & intTeamKey & ")"
        End If
    End If
    
    strName = Trim(txtLName.Text) & ",  " & Trim(txtFName.Text) & "  " & Trim(txtMName.Text)
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER strName, "is_LOAD"
End If

Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Function
End Function

Private Function PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If picSearch.Visible = True Then Exit Function
picSearch.ZOrder 0
txtSearch.Text = ""
picToolbar.Enabled = False
picMain.Enabled = False
picSearch.Visible = True
txtSearch.SetFocus
End Function

Private Function PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSearch.Visible = True Then cmdCancel_Click: Exit Function
    Unload Me
Else
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "PlayerInfo", "PlayerSetup", ""), "is_LOAD"
End If
End Function

Public Sub CLEARTEXT()
txtLName.Text = ""
txtFName.Text = ""
txtMName.Text = ""
txtHandicap.Text = ""
txtClass.Text = ""
txtAllowedTeam.Text = ""
txtIndex.Text = ""
txtClassIndex.Text = ""
cmbGender.ListIndex = -1
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
End Sub

Public Sub LOCKTEXT(bln As Boolean)
txtLName.Locked = bln
txtFName.Locked = bln
txtMName.Locked = bln
txtHandicap.Locked = bln
cmbGender.Locked = bln
txtClass.Locked = True
txtAllowedTeam.Locked = bln
txtIndex.Locked = bln
txtClassIndex.Locked = True
'If bln Then
'    txtLName.Locked = True
'    txtFName.Locked = True
'    txtMName.Locked = True
'    txtHandicap.Locked = True
'    cmbGender.Locked = True
'    txtClass.Locked = True
'    txtAllowedTeam.Locked = True
'Else
'    txtLName.Locked = False
'    txtFName.Locked = False
'    txtMName.Locked = False
'    txtHandicap.Locked = False
'    cmbGender.Locked = False
'    txtClass.Locked = True
'    txtAllowedTeam.Locked = False
'End If
End Sub


Public Sub TOOLBARFUNC(intSelect As Integer)
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

Private Sub cmdCancel_Click()
picSearch.Visible = False
picToolbar.Enabled = True
picMain.Enabled = True
txtLName.SetFocus
End Sub

Private Sub cmdOK_Click()
If lstResult.ListIndex = -1 Then Exit Sub
BROWSER lstResult.ItemData(lstResult.ListIndex), "is_FIND"
cmdCancel_Click
End Sub

Private Sub Command1_Click()
CommonDialog1.DialogTitle = "OPEN FILE"
CommonDialog1.Filename = ""
CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowOpen
strPath = CommonDialog1.Filename
If Trim(strPath) = "" Then Exit Sub
txtPath.Text = strPath
txtPath.SetFocus

'Timer1.Enabled = True
Timer2.Enabled = True

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
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "PlayerInfo", "PlayerSetup", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "PlayerInfo", "PlayerSetup", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "PlayerInfo", "PlayerSetup", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "PlayerInfo", "PlayerSetup", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption

s = "SELECT tbl_Scoring_TournamentInfo.* " & _
    " FROM tbl_Scoring_TournamentInfo " & _
    " WHERE (PK = " & TournamentKey & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtTournament.Text = rs!TournamentName
    txtDate.Text = Format(rs!TournamentStart, "mm/dd/yyyy") & " - " & Format(rs!TournamentEnd, "mm/dd/yyyy")
End If
rs.Close

cmbGender.Clear
cmbGender.AddItem "MALE"
cmbGender.AddItem "FEMALE"
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "PlayerInfo", "PlayerSetup", ""), "is_LOAD"
If Trim(txtLName.Text) = "" Then BROWSER GetSetting(App.EXEName, "PlayerInfo", "PlayerSetup", ""), "is_HOME"

tmp = SetWindowLong(txtLName.hwnd, GWL_STYLE, GetWindowLong(txtLName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtFName.hwnd, GWL_STYLE, GetWindowLong(txtFName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtMName.hwnd, GWL_STYLE, GetWindowLong(txtMName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picSearch.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstResult_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False

Set cn = New ADODB.Connection
cn.Provider = "Microsoft.Jet.OLEDB.4.0"
cn.ConnectionString = _
    "Data Source= " & Trim(txtPath.Text) & ";" & _
    "Extended Properties=Excel 8.0;"
cn.CursorLocation = adUseClient
If cn.State = adStateOpen Then cn.Close
cn.Open
i = 0: strTeamName = "": strTeamHDCP = ""
Set rs = New ADODB.Recordset
If rs.State = adStateOpen Then rs.Close
'rs.Open "SELECT * FROM [withTeam$] ", cn, adOpenDynamic, adLockOptimistic
rs.Open "SELECT * FROM [NoTeam$] ", cn, adOpenDynamic, adLockOptimistic
'NoTeam
While Not rs.EOF
    i = i + 1
    If IsNull(rs!PlayerName) = False Then
        Array1 = Split(rs!PlayerName, ",", -1, 1)
        
        dblrpv = CDbl(IIf(IsNull(rs!RPV), 0, IIf(IsNumeric(rs!RPV) = False, 0, rs!RPV)))
        dblmatina = CDbl(IIf(IsNull(rs!MATINA), 0, IIf(IsNumeric(rs!MATINA) = False, 0, rs!MATINA)))
        dblapo = CDbl(IIf(IsNull(rs!APO), 0, IIf(IsNumeric(rs!APO) = False, 0, rs!APO)))
        dblOthers = CDbl(IIf(IsNull(rs!Others), 0, IIf(IsNumeric(rs!Others) = False, 0, rs!Others)))
        
        If CDbl(dblrpv) < CDbl(dblmatina) Then
            If CDbl(dblrpv) = 0 Then
                dblHandicap = dblmatina
            Else
                dblHandicap = dblrpv
            End If
        Else
            If CDbl(dblmatina) = 0 Then
                dblHandicap = dblrpv
            Else
                dblHandicap = dblmatina
            End If
        End If
                
        If CDbl(dblHandicap) < CDbl(dblapo) Then
            If CDbl(dblHandicap) = 0 Then
                dblHandicap = dblapo
            Else
                dblHandicap = dblHandicap
            End If
        Else
            If CDbl(dblapo) = 0 Then
                dblHandicap = dblHandicap
            Else
                dblHandicap = dblapo
            End If
        End If
        
        If CDbl(dblHandicap) < CDbl(dblOthers) Then
            If CDbl(dblHandicap) = 0 Then
                dblHandicap = dblOthers
            Else
                dblHandicap = dblHandicap
            End If
        Else
            If CDbl(dblOthers) = 0 Then
                dblHandicap = dblHandicap
            Else
                dblHandicap = dblOthers
            End If
        End If
        
        If CDbl(dblHandicap) > 24 Then
            dblHandicap = 24
        Else
            dblHandicap = dblHandicap
        End If
        
        strClass = ""
        
        't = "SELECT Class " & _
            " From tbl_Scoring_TournamentInfo_Class " & _
            " WHERE (TournamentKey = " & TournamentKey & ") " & _
            " AND (HFrom <= " & CDbl(dblHandicap) & ") " & _
            " AND (HTo >= " & CDbl(dblHandicap) & ")"
        t = "SELECT Class " & _
            " From tbl_Scoring_TournamentInfo_Class " & _
            " Where (TournamentKey = " & TournamentKey & ") " & _
            " And (HFrom <= " & CDbl(dblHandicap) & ") " & _
            " And (HTo >= " & CDbl(dblHandicap) & ")"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            strClass = rt!Class
        End If
        rt.Close
'        MsgBox dblHandicap & " ; " & strClass
        
        t = "SELECT tbl_Scoring_PlayerName.* " & _
            " FROM tbl_Scoring_PlayerName " & _
            " WHERE (LastName = '" & Trim(FORMATSQL(CStr(Array1(0)))) & "') " & _
            " AND (FirstName = '" & Trim(FORMATSQL(CStr(Array1(1)))) & "')"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount = 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Scoring_PlayerName " & _
                              " (TournamentKey, LastName, FirstName, MiddleName, Gender, HandiCap, Class, AllowedTeam) " & _
                              " VALUES (" & TournamentKey & ", '" & Trim(FORMATSQL(CStr(Array1(0)))) & "', '" & Trim(FORMATSQL(CStr(Array1(1)))) & "', " & _
                              " '',1, " & dblHandicap & ", '" & strClass & "', " & AllowedTeam & ")"
        End If
        rt.Close
'        intPlayerKey = 0
'
'        t = "SELECT PK " & _
'            " FROM tbl_Scoring_PlayerName " & _
'            " WHERE (LastName = '" & Trim(FORMATSQL(CStr(Array1(0)))) & "') " & _
'            " AND (FirstName = '" & Trim(FORMATSQL(CStr(Array1(1)))) & "')"
'        If rt.State = adStateOpen Then rt.Close
'        rt.Open t, ConnOmega
'        If rt.RecordCount > 0 Then
'            intPlayerKey = rt!PK
'        End If
'        rt.Close
'
'        If Trim(strTeamName) <> Trim(rs!Team) Then
'            If Trim(strTeamName) <> "" Then
'                dblTeamHDCPTot = 0: dblTeamHDCPIndx = 0: strTeamClass = ""
'                For j = 1 To 3
'                    dblTeamHDCPTot = dblTeamHDCPTot + CDbl(ListView1.ListItems.Item(j).SubItems(1))
'                Next j
'                dblTeamHDCPIndx = CDbl(dblTeamHDCPTot) / 3
'
'                dblTeamHDCPIndx = Format(dblTeamHDCPIndx, "#0.0")
'
'                strTeamID = ""
'                t = "SELECT TOP 1 TeamID " & _
'                    " FROM tbl_Scoring_Team " & _
'                    " WHERE (TournamentKey = " & TournamentKey & ")" & _
'                    " ORDER BY TeamID DESC"
'                If rt.State = adStateOpen Then rt.Close
'                rt.Open t, ConnOmega
'                If rt.RecordCount > 0 Then
'                    strTeamID = Format(CDbl(rt!TeamID) + 1, "000#")
'                Else
'                    strTeamID = "0001"
'                End If
'                rt.Close
'
''                MsgBox strTeamID
'
'                Do
'                    t = "SELECT tbl_Scoring_Team.* " & _
'                        " FROM tbl_Scoring_Team " & _
'                        " WHERE (TournamentKey = " & TournamentKey & ")" & _
'                        " AND (TeamID = '" & strTeamID & "')"
'                    If rt.State = adStateOpen Then rt.Close
'                    rt.Open t, ConnOmega
'                    If rt.RecordCount = 0 Then
'                        rt.Close
'                        Exit Do
'                    End If
'                    rt.Close
'                    strTeamID = Format(CDbl(strTeamID) + 1, "000#")
''                    MsgBox "do end"
'                Loop
'
''                MsgBox "pass" & strTeamID
'
'                ConnOmega.Execute "INSERT INTO tbl_Scoring_Team " & _
'                                  " (TournamentKey, TeamID, LastModified, TeamName, TeamHDCP) " & _
'                                  " VALUES (" & TournamentKey & ", " & _
'                                  " '" & strTeamID & "', '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
'                                  " '" & FORMATSQL(Trim(CStr(strTeamName))) & "', " & _
'                                  " " & CDbl(dblTeamHDCPIndx) & ")"
'                intTeamKey = 0
'                t = "SELECT PK " & _
'                    " FROM tbl_Scoring_Team " & _
'                    " WHERE (TeamID = " & strTeamID & ") " & _
'                    " AND (TournamentKey = " & TournamentKey & ")"
'                If rt.State = adStateOpen Then rt.Close
'                rt.Open t, ConnOmega
'                If rt.RecordCount > 0 Then
'                    intTeamKey = rt!PK
'                End If
'                rt.Close
'                For j = 1 To ListView1.ListItems.Count
''                    MsgBox ListView1.ListItems.Item(j).SubItems(2)
'                    ConnOmega.Execute "INSERT INTO tbl_Scoring_Team_Detail " & _
'                                      " (TeamKey, Line, PlayerKey) " & _
'                                      " VALUES (" & intTeamKey & ", " & j & ", " & _
'                                      " " & ListView1.ListItems.Item(j).SubItems(2) & ")"
'                Next j
'            End If
'            strTeamName = rs!Team
'            ListView1.ListItems.Clear
'        End If
'
'        Set x = ListView1.ListItems.Add()
'        x.Text = ""
'        x.SubItems(1) = dblHandicap
'        x.SubItems(2) = intPlayerKey
        
    End If
    UpdateProgress Picture2, i / rs.RecordCount
    rs.MoveNext
Wend
rs.Close
MsgBox "Update Done!                    ", vbInformation, "Done"


End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
Screen.MousePointer = vbHourglass
Set cn = New ADODB.Connection
cn.Provider = "Microsoft.Jet.OLEDB.4.0"
cn.ConnectionString = _
    "Data Source= " & Trim(txtPath.Text) & ";" & _
    "Extended Properties=Excel 8.0;"
cn.CursorLocation = adUseClient
If cn.State = adStateOpen Then cn.Close
cn.Open

i = 0: strTeamName = "": strTeamHDCP = ""
Set rs = New ADODB.Recordset
If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT * FROM [PlayerName$] ", cn, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    i = i + 1
    If IsNull(rs!LastName) = False Then
        strClass = ""
        t = "SELECT Class " & _
            " From tbl_Scoring_TournamentInfo_Class " & _
            " WHERE (TournamentKey = " & TournamentKey & ") " & _
            " AND (HFrom <= " & IIf(IsNumeric(IIf(IsNull(rs!Handicap), 0, rs!Handicap)) = False, 0, rs!Handicap) & ") " & _
            " AND (HTo >= " & IIf(IsNumeric(IIf(IsNull(rs!Handicap), 0, rs!Handicap)) = False, 0, rs!Handicap) & ")"
        If rt.State = adStateOpen Then rt.Close
        rt.Open s, ConnOmega
        If rt.RecordCount > 0 Then
            strClass = rt!Class
        End If
        rt.Close
        t = "SELECT tbl_Scoring_PlayerName.* " & _
            " FROM tbl_Scoring_PlayerName " & _
            " WHERE (TournamentKey = " & TournamentKey & ") " & _
            " AND (LastName = '" & Trim(FORMATSQL(CStr(IIf(IsNull(rs!LastName), "", rs!LastName)))) & "') " & _
            " AND (FirstName = '" & Trim(FORMATSQL(CStr(IIf(IsNull(rs!FirstName), "", rs!FirstName)))) & "') " & _
            " AND (MiddleName = '" & Trim(FORMATSQL(CStr(IIf(IsNull(rs!MiddleName), "", rs!MiddleName)))) & "')"
        If rt.State = adStateOpen Then rt.Close
        If rt.RecordCount = 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Scoring_PlayerName " & _
                              " (TournamentKey, LastName, FirstName, MiddleName, Gender, HandiCap, Class, AllowedTeam, iIndex, LastModified) " & _
                              " VALUES (" & TournamentKey & ", " & _
                              " '" & Trim(FORMATSQL(CStr(IIf(IsNull(rs!LastName), "", rs!LastName)))) & "', " & _
                              " '" & Trim(FORMATSQL(CStr(IIf(IsNull(rs!FirstName), "", rs!FirstName)))) & "', " & _
                              " '" & Trim(FORMATSQL(CStr(IIf(IsNull(rs!MiddleName), "", rs!MiddleName)))) & "', " & _
                              " " & IIf(IsNumeric(IIf(IsNull(rs!Gender), 1, rs!Gender)) = False, 1, rs!Gender) & ", " & _
                              " " & IIf(IsNumeric(IIf(IsNull(rs!Handicap), 0, rs!Handicap)) = False, 0, rs!Handicap) & ", " & _
                              " '" & FORMATSQL(CStr(strClass)) & "', " & AllowedTeam & ", " & rs!Index & ", '" & CStr(Now) & " - " & gbl_CompleteName & "')"
        Else
            ConnOmega.Execute "UPDATE tbl_Scoring_PlayerName " & _
                              " SET Gender = " & IIf(IsNumeric(IIf(IsNull(rs!Gender), 1, rs!Gender)) = False, 1, rs!Gender) & ", " & _
                              " HandiCap = " & IIf(IsNumeric(IIf(IsNull(rs!Handicap), 0, rs!Handicap)) = False, 0, rs!Handicap) & ", " & _
                              " Class = '" & FORMATSQL(CStr(strClass)) & "', " & _
                              " AllowedTeam = " & AllowedTeam & ", " & _
                              " iIndex = " & rs!Index & ", " & _
                              " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                              " WHERE (PK " & rt!PK & ")"
        End If
        rt.Close
    End If
    UpdateProgress Picture2, i / rs.RecordCount
    rs.MoveNext
Wend
rs.Close
Screen.MousePointer = vbDefault
MsgBox "Update Done!                    ", vbInformation, "Done"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "PlayerInfo", "PlayerSetup", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "PlayerInfo", "PlayerSetup", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "PlayerInfo", "PlayerSetup", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "PlayerInfo", "PlayerSetup", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Close":   PRESS_ESCAPE
    Case Else:      Exit Sub
End Select
End Sub


Private Sub txtFName_GotFocus()
HTEXT txtFName
End Sub

Private Sub txtFName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtMName.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtLName.SetFocus
End If
End Sub

Private Sub txtHandicap_Change()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If RETURNTEXTVALUE(txtHandicap) > TopHandicap Then MsgBox "Invalid Handicap!                 ", vbCritical, "Error...": txtHandicap.SetFocus: HTEXT txtHandicap: Exit Sub
    s = "SELECT Class " & _
        " From tbl_Scoring_TournamentInfo_Class " & _
        " WHERE (TournamentKey = " & TournamentKey & ") " & _
        " AND (HFrom <= " & RETURNTEXTVALUE(txtHandicap) & ") " & _
        " AND (HTo >= " & RETURNTEXTVALUE(txtHandicap) & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        txtClass.Text = rs!Class
    Else
        txtClass.Text = ""
    End If
    rs.Close
End If
End Sub

Private Sub txtHandicap_GotFocus()
HTEXT txtHandicap
End Sub

Private Sub txtHandicap_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtAllowedTeam.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtMName.SetFocus
End If
End Sub

Private Sub txtHandicap_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtIndex_Change()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If RETURNTEXTVALUE(txtIndex) > TopIndex Then MsgBox "Invalid Index!                 ", vbCritical, "Error...": txtIndex.SetFocus: HTEXT txtHandicap: Exit Sub
    s = "SELECT Class " & _
        " From tbl_Scoring_TournamentInfo_Index " & _
        " WHERE (TournamentKey = " & TournamentKey & ") " & _
        " AND (HFrom <= " & RETURNTEXTVALUE(txtIndex) & ") " & _
        " AND (HTo >= " & RETURNTEXTVALUE(txtIndex) & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        txtClassIndex.Text = rs!Class
    Else
        txtClassIndex.Text = ""
    End If
    rs.Close
End If
End Sub

Private Sub txtIndex_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtIndex_LostFocus()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    txtIndex.Text = RoundOffIndex(RETURNTEXTVALUE(txtIndex))
End If
End Sub

Private Sub txtLName_GotFocus()
HTEXT txtLName
End Sub

Private Sub txtLName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtFName.SetFocus
ElseIf KeyCode = vbKeyUp Then

End If
End Sub


Private Sub txtMName_GotFocus()
HTEXT txtMName
End Sub

Private Sub txtMName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtHandicap.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtFName.SetFocus
End If
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstResult.Clear: Exit Sub
lstResult.Clear
s = "SELECT PK, LastName, FirstName, MiddleName " & _
    " From tbl_Scoring_PlayerName " & _
    " WHERE (LastName LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
    " AND (TournamentKey = " & TournamentKey & ") " & _
    " ORDER BY LastName, FirstName, MiddleName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResult.AddItem rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
    lstResult.ItemData(lstResult.NewIndex) = rs!PK
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
