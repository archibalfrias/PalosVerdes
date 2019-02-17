VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTeamSetup 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTeamSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picSearch 
      Height          =   4695
      Left            =   2280
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   8281
      BackColor       =   15396057
      Begin VB.ListBox lstTeam 
         Height          =   1425
         Left            =   120
         TabIndex        =   25
         Top             =   2520
         Width           =   4095
      End
      Begin VB.CommandButton cmdOK 
         Height          =   480
         Left            =   480
         Picture         =   "frmTeamSetup.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   4080
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancel 
         Height          =   480
         Left            =   2235
         Picture         =   "frmTeamSetup.frx":0F3C
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   4080
         Width           =   1560
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   4095
      End
      Begin VB.ListBox lstResult 
         Height          =   1620
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   4095
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   40
         TabIndex        =   24
         Top             =   45
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
         Icon            =   "frmTeamSetup.frx":1698
      End
   End
   Begin RPVGCC.b8Container picSLine 
      Height          =   855
      Left            =   1800
      TabIndex        =   11
      Top             =   1800
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox txtHandicap1 
         Height          =   315
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtClass1 
         Height          =   315
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.CommandButton btnSupplier 
         BackColor       =   &H00D4D4D4&
         Height          =   270
         Left            =   4920
         Picture         =   "frmTeamSetup.frx":1C32
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   380
         Width           =   240
      End
      Begin VB.TextBox txtPlayerKey 
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtPlayer 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   5055
      End
      Begin VB.TextBox txtPlayerName1 
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtPlayerKey1 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.ComboBox cmbPlayer 
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   0
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Player"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   36
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   37
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
         MouseIcon       =   "frmTeamSetup.frx":1DDC
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
      Height          =   2775
      Left            =   1440
      ScaleHeight     =   2775
      ScaleWidth      =   6015
      TabIndex        =   7
      Top             =   2040
      Width           =   6015
      Begin VB.TextBox txtTeamName 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox txtClass 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   0
         Width           =   855
      End
      Begin VB.TextBox txtHandicap 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   0
         Width           =   1095
      End
      Begin MSComctlLib.ListView lstMember 
         Height          =   1815
         Left            =   0
         TabIndex        =   9
         Top             =   960
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3201
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "#"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "PlayerKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Player Name"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Handicap"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Class"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Gross Pts"
            Object.Width           =   1587
         EndProperty
      End
      Begin VB.TextBox txtTeamID 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Team Name"
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Handicap"
         Height          =   255
         Left            =   2760
         TabIndex        =   17
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Team Members"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   5775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Team ID"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ListView lstHandicap 
      Height          =   1455
      Left            =   120
      TabIndex        =   34
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2566
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Handicap"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.ListBox lstPlayerSearch 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   1800
      TabIndex        =   28
      Top             =   2475
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.PictureBox picTour 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   735
      Left            =   1680
      ScaleHeight     =   735
      ScaleWidth      =   5535
      TabIndex        =   1
      Top             =   1200
      Width           =   5535
      Begin VB.TextBox txtTournament 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   0
         Width           =   4215
      End
      Begin VB.TextBox txtDate 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tournament"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Top             =   5085
      Width           =   9075
      _ExtentX        =   16007
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
            Picture         =   "frmTeamSetup.frx":20F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamSetup.frx":2DD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamSetup.frx":3AAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamSetup.frx":4784
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamSetup.frx":545E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamSetup.frx":6138
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamSetup.frx":6E12
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamSetup.frx":7AEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamSetup.frx":87C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamSetup.frx":90A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamSetup.frx":9D7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamSetup.frx":AA54
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamSetup.frx":B72E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamSetup.frx":C408
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTeamSetup.frx":D0E2
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTeamSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public TournamentKey        As Double
'Public NoofPlayerPerTeam    As Double

Dim TRANSACTIONTYPE         As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim TRANSACTIONTYPE_DET     As Long
Const is_DET_REFRESH = 0
Const is_DET_ADDING = 1
Const is_DET_EDITTING = 2


Dim ROW As Long
Dim PlayerCount As Double

Dim ListFocus As Long

Dim SearchFocus As Long
Dim tmp As Long

Dim x, dblHandicap, i, cnt, strTeamID, TeamKey, j, a, dblTeamHDCP, dblTeamHDCPInx, Arr, _
dblTeamHDCPInxtmp, iIndexRange, iHDCPIndex, iTeamPerPlayer, iTeamIndex, Array1


Private Function BROWSER(Team, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If Team <> "" Then
            s = "SELECT tbl_Scoring_Team.* " & _
                " From tbl_Scoring_Team " & _
                " WHERE (TournamentKey = " & TournamentKey & ") " & _
                " AND (TeamID = '" & Team & "') " & _
                " ORDER BY TeamID"
        Else
            s = "SELECT tbl_Scoring_Team.* " & _
                " From tbl_Scoring_Team " & _
                " WHERE (TournamentKey = " & TournamentKey & ") " & _
                " ORDER BY TeamID"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT tbl_Scoring_Team.* " & _
            " From tbl_Scoring_Team " & _
            " WHERE (TournamentKey = " & TournamentKey & ") " & _
            " ORDER BY TeamID"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT tbl_Scoring_Team.* " & _
            " From tbl_Scoring_Team " & _
            " WHERE (TournamentKey = " & TournamentKey & ") " & _
            " AND (TeamID < '" & Team & "') " & _
            " ORDER BY TeamID DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT tbl_Scoring_Team.* " & _
            " From tbl_Scoring_Team " & _
            " WHERE (TournamentKey = " & TournamentKey & ") " & _
            " AND (TeamID > '" & Team & "') " & _
            " ORDER BY TeamID"
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT tbl_Scoring_Team.* " & _
            " From tbl_Scoring_Team " & _
            " WHERE (TournamentKey = " & TournamentKey & ") " & _
            " ORDER BY TeamID DESC"
    Case Else: Exit Function
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtTeamID.Text = rs!TeamID
    txtTeamName.Text = rs!TeamName
    If TeamAverage = 1 Then
        txtHandicap.Text = rs!TeamHDCP
        txtClass.Text = ""
        t = "SELECT Class " & _
            " From tbl_Scoring_TournamentInfo_Class " & _
            " Where (TournamentKey = " & TournamentKey & ") " & _
            " And (HFrom <= " & CDbl(rs!TeamHDCP) & ") " & _
            " And (HTo >= " & CDbl(rs!TeamHDCP) & ")"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
             txtClass.Text = rt!Class
        End If
        rt.Close
    ElseIf TeamAverage = 2 Then
        txtHandicap.Text = rs!TeamIndex
        txtClass.Text = ""
        t = "SELECT Class " & _
            " From tbl_Scoring_TournamentInfo_Index " & _
            " Where (TournamentKey = " & TournamentKey & ") " & _
            " And (HFrom <= " & CDbl(rs!TeamIndex) & ") " & _
            " And (HTo >= " & CDbl(rs!TeamIndex) & ")"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
             txtClass.Text = rt!Class
        End If
        rt.Close
    End If
    
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", "Last Modified : " & rs!LastModified)
    
    SaveSetting App.EXEName, "TeamControl", "TeamCtrl", rs!TeamID
    
    'MsgBox ScoringType
    'MsgBox TeamAverage
    
    If ScoringType = 5 Then
        If TeamAverage = 2 Then
            t = "SELECT tbl_Scoring_Team_Detail.Line, tbl_Scoring_Team_Detail.PlayerKey, " & _
                " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
                " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.iIndex, tbl_Scoring_PlayerName.Class, " & _
                " ISNULL((SELECT tbl_Scoring_ScoreCard_ModMolave.GrossPoints as GrossPoints " & _
                " From tbl_Scoring_ScoreCard_ModMolave " & _
                " Where (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) as GrossPts " & _
                " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                " ORDER BY tbl_Scoring_PlayerName.iIndex"
        Else
            t = "SELECT tbl_Scoring_Team_Detail.Line, tbl_Scoring_Team_Detail.PlayerKey, " & _
                " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
                " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.iIndex, tbl_Scoring_PlayerName.Class, " & _
                " ISNULL((SELECT tbl_Scoring_ScoreCard_ModMolave.GrossPoints as GrossPoints " & _
                " From tbl_Scoring_ScoreCard_ModMolave " & _
                " Where (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) as GrossPts " & _
                " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                " ORDER BY tbl_Scoring_Team_Detail.Line"
        End If
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        i = 0
        If rt.RecordCount > 0 Then
            With lstMember.ListItems
                .Clear
                While Not rt.EOF
                    dblHandicap = dblHandicap + CDbl(rt!Handicap)
                    i = i + 1
                    Set x = .Add()
                    x.Text = " "
                    x.SubItems(1) = Format(i, "0#")
                    x.SubItems(2) = rt!PlayerKey
                    x.SubItems(3) = rt!PlayerName
                    If TeamAverage = 2 Then
                        x.SubItems(4) = rt!iIndex
                        u = "SELECT Class " & _
                            " From tbl_Scoring_TournamentInfo_Index " & _
                            " Where (TournamentKey = " & TournamentKey & ") " & _
                            " And (HFrom <= " & CDbl(rt!iIndex) & ") " & _
                            " And (HTo >= " & CDbl(rt!iIndex) & ")"
                        If ru.State = adStateOpen Then ru.Close
                        ru.Open u, ConnOmega
                        If ru.RecordCount > 0 Then
                            x.SubItems(5) = ru!Class
                        Else
                            x.SubItems(5) = " "
                        End If
                    Else
                        x.SubItems(4) = rt!Handicap
                        u = "SELECT Class " & _
                            " From tbl_Scoring_TournamentInfo_Class " & _
                            " Where (TournamentKey = " & TournamentKey & ") " & _
                            " And (HFrom <= " & CDbl(rt!Handicap) & ") " & _
                            " And (HTo >= " & CDbl(rt!Handicap) & ")"
                        If ru.State = adStateOpen Then ru.Close
                        ru.Open u, ConnOmega
                        If ru.RecordCount > 0 Then
                            x.SubItems(5) = ru!Class
                        Else
                            x.SubItems(5) = " "
                        End If
                    End If
                    x.SubItems(6) = rt!GrossPts
                    ru.Close
                    rt.MoveNext
                Wend
            End With
        Else
            lstMember.ListItems.Clear
            Set x = lstMember.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = " "
            x.SubItems(2) = " "
            x.SubItems(3) = " "
            x.SubItems(4) = " "
            x.SubItems(5) = " "
        End If
        rt.Close
    ElseIf ScoringType = 3 Then
'        t = "SELECT tbl_Scoring_Team_Detail.Line, tbl_Scoring_Team_Detail.PlayerKey, " & _
'            " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
'            " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.iIndex, tbl_Scoring_PlayerName.Class, " & _
'            " ISNULL((SELECT tbl_Scoring_ScoreCard_System36.NetPoints as GrossPoints " & _
'            " From tbl_Scoring_ScoreCard_System36 " & _
'            " Where (tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) as GrossPts " & _
'            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
'            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
'            " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
'            " ORDER BY tbl_Scoring_Team_Detail.Line"
        If TeamAverage = 1 Then
            t = "SELECT tbl_Scoring_Team_Detail.Line, tbl_Scoring_Team_Detail.PlayerKey, " & _
                " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
                " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.iIndex, tbl_Scoring_PlayerName.Class, " & _
                " ISNULL((SELECT tbl_Scoring_ScoreCard_ModMolave.GrossPoints as GrossPoints " & _
                " From tbl_Scoring_ScoreCard_ModMolave " & _
                " Where (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) as GrossPts " & _
                " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                " ORDER BY tbl_Scoring_Team_Detail.Line"
        ElseIf TeamAverage = 2 Then
            t = "SELECT tbl_Scoring_Team_Detail.Line, tbl_Scoring_Team_Detail.PlayerKey, " & _
                " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
                " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.iIndex, tbl_Scoring_PlayerName.Class, " & _
                " ISNULL((SELECT tbl_Scoring_ScoreCard_ModMolave.GrossPoints as GrossPoints " & _
                " From tbl_Scoring_ScoreCard_ModMolave " & _
                " Where (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) as GrossPts " & _
                " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                " ORDER BY tbl_Scoring_PlayerName.iIndex"
        End If
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        i = 0
        If rt.RecordCount > 0 Then
            With lstMember.ListItems
                .Clear
                While Not rt.EOF
                    dblHandicap = dblHandicap + CDbl(rt!Handicap)
                    i = i + 1
                    Set x = .Add()
                    x.Text = " "
                    x.SubItems(1) = Format(i, "0#")
                    x.SubItems(2) = rt!PlayerKey
                    x.SubItems(3) = rt!PlayerName
                    If TeamAverage = 1 Then
                        x.SubItems(4) = rt!Handicap
                        u = "SELECT Class " & _
                            " From tbl_Scoring_TournamentInfo_Class " & _
                            " Where (TournamentKey = " & TournamentKey & ") " & _
                            " And (HFrom <= " & CDbl(rt!Handicap) & ") " & _
                            " And (HTo >= " & CDbl(rt!Handicap) & ")"
                        If ru.State = adStateOpen Then ru.Close
                        ru.Open u, ConnOmega
                        If ru.RecordCount > 0 Then
                            x.SubItems(5) = ru!Class
                        Else
                            x.SubItems(5) = " "
                        End If
                    ElseIf TeamAverage = 2 Then
                        x.SubItems(4) = rt!iIndex
                        u = "SELECT Class " & _
                            " From tbl_Scoring_TournamentInfo_Index " & _
                            " Where (TournamentKey = " & TournamentKey & ") " & _
                            " And (HFrom <= " & CDbl(rt!iIndex) & ") " & _
                            " And (HTo >= " & CDbl(rt!iIndex) & ")"
                        If ru.State = adStateOpen Then ru.Close
                        ru.Open u, ConnOmega
                        If ru.RecordCount > 0 Then
                            x.SubItems(5) = ru!Class
                        Else
                            x.SubItems(5) = " "
                        End If
                    End If
                    x.SubItems(6) = rt!GrossPts
                    ru.Close
                    rt.MoveNext
                Wend
            End With
        Else
            lstMember.ListItems.Clear
            Set x = lstMember.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = " "
            x.SubItems(2) = " "
            x.SubItems(3) = " "
            x.SubItems(4) = " "
            x.SubItems(5) = " "
        End If
        rt.Close
    ElseIf ScoringType = 2 Then
        't = "SELECT tbl_Scoring_Team_Detail.Line, tbl_Scoring_Team_Detail.PlayerKey, " & _
            " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
            " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.Class, " & _
            " ISNULL((SELECT tbl_Scoring_ScoreCard_ModStableFord.GrossPoints as GrossPoints " & _
            " From tbl_Scoring_ScoreCard_ModStableFord " & _
            " Where (tbl_Scoring_ScoreCard_ModStableFord.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) as GrossPts " & _
            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
            " ORDER BY tbl_Scoring_Team_Detail.Line"
        't = "SELECT tbl_Scoring_Team_Detail.Line, tbl_Scoring_Team_Detail.PlayerKey, " & _
            " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
            " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.iIndex, tbl_Scoring_PlayerName.Class, " & _
            " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard_ModStableFord.GrossPoints) as GrossPoints " & _
            " From tbl_Scoring_ScoreCard_ModStableFord " & _
            " Where (tbl_Scoring_ScoreCard_ModStableFord.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) as GrossPts " & _
            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
            " ORDER BY tbl_Scoring_Team_Detail.Line"
        If TeamAverage = 1 Then
            t = "SELECT tbl_Scoring_Team_Detail.Line, tbl_Scoring_Team_Detail.PlayerKey, " & _
                " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
                " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.iIndex, tbl_Scoring_PlayerName.Class, " & _
                " ISNULL((SELECT tbl_Scoring_ScoreCard_ModMolave.GrossPoints as GrossPoints " & _
                " From tbl_Scoring_ScoreCard_ModMolave " & _
                " Where (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) as GrossPts " & _
                " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                " ORDER BY tbl_Scoring_Team_Detail.Line"
        ElseIf TeamAverage = 2 Then
            t = "SELECT tbl_Scoring_Team_Detail.Line, tbl_Scoring_Team_Detail.PlayerKey, " & _
                " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
                " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.iIndex, tbl_Scoring_PlayerName.Class, " & _
                " ISNULL((SELECT tbl_Scoring_ScoreCard_ModMolave.GrossPoints as GrossPoints " & _
                " From tbl_Scoring_ScoreCard_ModMolave " & _
                " Where (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) as GrossPts " & _
                " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                " ORDER BY tbl_Scoring_PlayerName.iIndex"
        End If
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        i = 0
        If rt.RecordCount > 0 Then
            With lstMember.ListItems
                .Clear
                While Not rt.EOF
                    dblHandicap = dblHandicap + CDbl(rt!Handicap)
                    i = i + 1
                    Set x = .Add()
                    x.Text = " "
                    x.SubItems(1) = Format(i, "0#")
                    x.SubItems(2) = rt!PlayerKey
                    x.SubItems(3) = rt!PlayerName
                    x.SubItems(4) = rt!Handicap
                    If TeamAverage = 1 Then
                        x.SubItems(4) = rt!Handicap
                        u = "SELECT Class " & _
                            " From tbl_Scoring_TournamentInfo_Class " & _
                            " Where (TournamentKey = " & TournamentKey & ") " & _
                            " And (HFrom <= " & CDbl(rt!Handicap) & ") " & _
                            " And (HTo >= " & CDbl(rt!Handicap) & ")"
                        If ru.State = adStateOpen Then ru.Close
                        ru.Open u, ConnOmega
                        If ru.RecordCount > 0 Then
                            x.SubItems(5) = ru!Class
                        Else
                            x.SubItems(5) = " "
                        End If
                    ElseIf TeamAverage = 2 Then
                        x.SubItems(4) = rt!iIndex
                        u = "SELECT Class " & _
                            " From tbl_Scoring_TournamentInfo_Index " & _
                            " Where (TournamentKey = " & TournamentKey & ") " & _
                            " And (HFrom <= " & CDbl(rt!iIndex) & ") " & _
                            " And (HTo >= " & CDbl(rt!iIndex) & ")"
                        If ru.State = adStateOpen Then ru.Close
                        ru.Open u, ConnOmega
                        If ru.RecordCount > 0 Then
                            x.SubItems(5) = ru!Class
                        Else
                            x.SubItems(5) = " "
                        End If
                    End If
                    
                    x.SubItems(6) = rt!GrossPts
                    ru.Close
                    rt.MoveNext
                Wend
            End With
        Else
            lstMember.ListItems.Clear
            Set x = lstMember.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = " "
            x.SubItems(2) = " "
            x.SubItems(3) = " "
            x.SubItems(4) = " "
            x.SubItems(5) = " "
        End If
        rt.Close
    ElseIf ScoringType = 1 Then
        't = "SELECT tbl_Scoring_Team_Detail.TeamKey, tbl_Scoring_Team_Detail.Line, " & _
            " tbl_Scoring_Team_Detail.PlayerKey, tbl_Scoring_PlayerName.LastName +',  ' + " & _
            " tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName as PlayerName, " & _
            " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.iIndex, tbl_Scoring_PlayerName.Class, " & _
            " IsNull((SELECT SUM(tbl_Scoring_ScoreCard.GrossPoints) AS GrossPoints " & _
            " From tbl_Scoring_ScoreCard " & _
            " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") " & _
            " AND (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)), 0) AS GrossPts " & _
            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
            " ORDER BY tbl_Scoring_Team_Detail.Line"
        If TeamAverage = 1 Then
            t = "SELECT tbl_Scoring_Team_Detail.Line, tbl_Scoring_Team_Detail.PlayerKey, " & _
                " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
                " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.iIndex, tbl_Scoring_PlayerName.Class, " & _
                " ISNULL((SELECT tbl_Scoring_ScoreCard_ModMolave.GrossPoints as GrossPoints " & _
                " From tbl_Scoring_ScoreCard_ModMolave " & _
                " Where (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) as GrossPts " & _
                " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                " ORDER BY tbl_Scoring_Team_Detail.Line"
        ElseIf TeamAverage = 2 Then
            t = "SELECT tbl_Scoring_Team_Detail.Line, tbl_Scoring_Team_Detail.PlayerKey, " & _
                " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
                " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.iIndex, tbl_Scoring_PlayerName.Class, " & _
                " ISNULL((SELECT tbl_Scoring_ScoreCard_ModMolave.GrossPoints as GrossPoints " & _
                " From tbl_Scoring_ScoreCard_ModMolave " & _
                " Where (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) as GrossPts " & _
                " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                " ORDER BY tbl_Scoring_PlayerName.iIndex"
        End If
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        i = 0
        If rt.RecordCount > 0 Then
            With lstMember.ListItems
                .Clear
                While Not rt.EOF
                    dblHandicap = dblHandicap + CDbl(rt!Handicap)
                    i = i + 1
                    Set x = .Add()
                    x.Text = " "
                    x.SubItems(1) = Format(i, "0#")
                    x.SubItems(2) = rt!PlayerKey
                    x.SubItems(3) = rt!PlayerName
                    x.SubItems(4) = rt!Handicap
                    If TeamAverage = 1 Then
                        x.SubItems(4) = rt!Handicap
                        u = "SELECT Class " & _
                            " From tbl_Scoring_TournamentInfo_Class " & _
                            " Where (TournamentKey = " & TournamentKey & ") " & _
                            " And (HFrom <= " & CDbl(rt!Handicap) & ") " & _
                            " And (HTo >= " & CDbl(rt!Handicap) & ")"
                        If ru.State = adStateOpen Then ru.Close
                        ru.Open u, ConnOmega
                        If ru.RecordCount > 0 Then
                            x.SubItems(5) = ru!Class
                        Else
                            x.SubItems(5) = " "
                        End If
                    ElseIf TeamAverage = 2 Then
                        x.SubItems(4) = rt!iIndex
                        u = "SELECT Class " & _
                            " From tbl_Scoring_TournamentInfo_Index " & _
                            " Where (TournamentKey = " & TournamentKey & ") " & _
                            " And (HFrom <= " & CDbl(rt!iIndex) & ") " & _
                            " And (HTo >= " & CDbl(rt!iIndex) & ")"
                        If ru.State = adStateOpen Then ru.Close
                        ru.Open u, ConnOmega
                        If ru.RecordCount > 0 Then
                            x.SubItems(5) = ru!Class
                        Else
                            x.SubItems(5) = " "
                        End If
                    End If
                    
                    x.SubItems(6) = rt!GrossPts
                    
                    rt.MoveNext
                Wend
            End With
        Else
            lstMember.ListItems.Clear
            Set x = lstMember.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = " "
            x.SubItems(2) = " "
            x.SubItems(3) = " "
            x.SubItems(4) = " "
            x.SubItems(5) = " "
        End If
        rt.Close
    End If
    
    
'    SaveSetting App.EXEName, "TeamControl", "TeamCtrl", rs!TeamID
    
End If
rs.Close
If rs.State = adStateOpen Then rs.Close
End Function

Private Function PRESS_INSERT()

If TRANSACTIONTYPE = is_REFRESH Then
    If picSearch.Visible = True Then Exit Function
    If AccessRights("Scoring Team Information", "Add") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Function
    End If
    txtTeamID.SetFocus
    txtTeamName.Locked = False
    CLEARTEXT
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_ADDING
    txtTeamName.SetFocus
ElseIf TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If picSLine.Visible = True Then Exit Function
    If ListFocus = 0 Then Exit Function
    PlayerCount = 0
    For i = 1 To lstMember.ListItems.Count
        If Trim(lstMember.ListItems.Item(i).SubItems(2)) <> "" Then
            PlayerCount = PlayerCount + 1
        End If
    Next i
    If CDbl(NoofPlayerPerTeam) < (CDbl(PlayerCount)) + 1 Then MsgBox "Team Member Already Exceeded!              ", vbCritical, "Error...": Exit Function
    With lstMember.ListItems
        If Trim(.Item(ROW).SubItems(2)) <> "" Then
            Set x = .Add()
            x.Text = " "
            x.SubItems(1) = Format(.Count, "0#")
            x.SubItems(2) = " "
            x.SubItems(3) = " "
            x.SubItems(4) = " "
            x.SubItems(5) = " "
            ROW = .Count
        Else
            .Item(1).SubItems(1) = "01"
            ROW = 1
        End If
    End With
    TRANSACTIONTYPE_DET = is_DET_ADDING
    txtPlayer.Text = ""
    picMain.Enabled = False
    picToolbar.Enabled = False
    cmbPlayer.ListIndex = -1
    picSLine.ZOrder 0
    picSLine.Visible = True
    txtPlayer.SetFocus
'    cmbPlayer.SetFocus
End If
End Function

Private Function PRESS_F2()
If TRANSACTIONTYPE = is_REFRESH Then
    If Statusbar1.Panels(1).Text = "" Then Exit Function
    If picSearch.Visible = True Then Exit Function
    If AccessRights("Scoring Team Information", "Edit") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Function
    End If
    txtTeamName.Locked = False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_EDITTING
ElseIf TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If picSLine.Visible = True Then Exit Function
    If ListFocus = 0 Then Exit Function
    With lstMember.ListItems
        txtPlayerKey1.Text = .Item(ROW).SubItems(2)
        txtPlayerName1.Text = .Item(ROW).SubItems(3)
        txtHandicap1.Text = .Item(ROW).SubItems(4)
        txtClass1.Text = .Item(ROW).SubItems(5)
        'cmbPlayer.Text = .Item(ROW).SubItems(3)
    End With
    TRANSACTIONTYPE_DET = is_DET_EDITTING
    picSLine.ZOrder 0
    picSLine.Visible = True
'    cmbPlayer.SetFocus
End If
End Function

Private Function PRESS_DELETE()
If TRANSACTIONTYPE = is_REFRESH Then
    If Statusbar1.Panels(1).Text = "" Then Exit Function
    If picSearch.Visible = True Then Exit Function
    If AccessRights("Scoring Team Information", "Delete") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Function
    End If
    If MsgBox("ARE YOU SURE IN DELETING THIS TEAM?                  ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Function
    ConnOmega.Execute "DELETE FROM tbl_Scoring_Team WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    CLEARTEXT
    BROWSER GetSetting(App.EXEName, "TeamControl", "TeamCtrl", ""), "is_PAGEDOWN"
    If Trim(txtTeamID.Text) = "" Then BROWSER GetSetting(App.EXEName, "TeamControl", "TeamCtrl", ""), "is_HOME"
ElseIf TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If picSLine.Visible = True Then Exit Function
    If ListFocus = 0 Then Exit Function
    With lstMember.ListItems
        If .Count = 1 Then
            .Item(1).SubItems(1) = " "
            .Item(1).SubItems(2) = " "
            .Item(1).SubItems(3) = " "
            .Item(1).SubItems(4) = " "
            .Item(1).SubItems(5) = " "
            ROW = 1
            lstMember.ListItems(ROW).EnsureVisible
            lstMember.ListItems(ROW).Selected = True
        Else
            .Remove ROW
            If CDbl(ROW) > CDbl(.Count) Then
                ROW = .Count
            End If
            lstMember.ListItems(ROW).EnsureVisible
            lstMember.ListItems(ROW).Selected = True
        End If
    End With
    
    
End If
End Function

Private Function PRESS_F5()
If picSLine.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
cnt = 0
With lstMember.ListItems
    For i = 1 To .Count
        If CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2))) <> 0 Then
            cnt = cnt + 1
        End If
    Next i
    If CDbl(cnt) = 0 Then MsgBox "Please Supply Players!                  ", vbCritical, "Error...": Exit Function
    PlayerCount = 0
    For i = 1 To lstMember.ListItems.Count
        If Trim(lstMember.ListItems.Item(i).SubItems(2)) <> "" Then
            PlayerCount = PlayerCount + 1
        End If
    Next i
        
    If CDbl(NoofPlayerPerTeam) < CDbl(PlayerCount) Then MsgBox "Team Member Already Exceeded!              ", vbCritical, "Error...": Exit Function
    
    On Error GoTo PG:
    
    If TRANSACTIONTYPE = is_ADDING Then
        
        If CDbl(AllowedTeam) = 1 Then
            For i = 1 To .Count
                If CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2))) <> 0 Then
                    '== Check Teammate
                    For j = 1 To .Count
                        If CDbl(i) <> CDbl(j) Then
                            s = "SELECT TeamKey " & _
                                " From tbl_Scoring_Team_Detail " & _
                                " WHERE (PlayerKey = " & .Item(i).SubItems(2) & ") " & _
                                " AND (TeamKey <> 0)"
                            If rs.State = adStateOpen Then rs.Close
                            rs.Open s, ConnOmega
                            If rs.RecordCount > 0 Then
                                MsgBox "'" & .Item(i).SubItems(3) & "' is not valid for this Team!                  ", vbCritical, "Error..."
                                rs.Close
                                Exit Function
                            End If
                            rs.Close
                        End If
                    Next j
                End If
            Next i
        Else
            For i = 1 To .Count
                If Trim(.Item(i).SubItems(2)) <> "" Then
                    iTeamPerPlayer = 1
                    t = "SELECT tbl_Scoring_Team_Detail.PlayerKey, tbl_Scoring_Team.TournamentKey " & _
                        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                        " tbl_Scoring_Team ON tbl_Scoring_Team_Detail.TeamKey = tbl_Scoring_Team.PK " & _
                        " WHERE (tbl_Scoring_Team.TournamentKey = " & TournamentKey & ") " & _
                        " AND (tbl_Scoring_Team_Detail.PlayerKey = " & .Item(i).SubItems(2) & ")"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        iTeamPerPlayer = iTeamPerPlayer + 1
                        rt.MoveNext
                    Wend
                    rt.Close
                End If
                If CDbl(AllowedTeam) < CDbl(iTeamPerPlayer) Then MsgBox "Number of Team per Player Already Exceed!                  ", vbCritical, "Error...": Exit Function
                              
            Next i
        End If
        
        strTeamID = ""
        s = "SELECT TOP 1 TeamID " & _
            " FROM tbl_Scoring_Team " & _
            " WHERE (TournamentKey = " & TournamentKey & ")" & _
            " ORDER BY TeamID DESC"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            strTeamID = Format(CDbl(rs!TeamID) + 1, "000#")
        Else
            strTeamID = "0001"
        End If
        rs.Close
        Do
            s = "SELECT tbl_Scoring_Team.* " & _
                " FROM tbl_Scoring_Team " & _
                " WHERE (TournamentKey = " & TournamentKey & ")" & _
                " AND (TeamID = '" & strTeamID & "')"
            If rs.State = adStateOpen Then rs.Close
            rs.Open s, ConnOmega
            If rs.RecordCount = 0 Then
                rs.Close
                Exit Do
            End If
            rs.Close
            strTeamID = Format(CDbl(strTeamID) + 1, "000#")
        Loop
        
        If Trim(txtTeamName.Text) = "" Then txtTeamName.Text = strTeamID
        
        ConnOmega.Execute "INSERT INTO tbl_Scoring_Team " & _
                          " (TournamentKey, TeamID, LastModified, TeamName, TeamHDCP) " & _
                          " VALUES (" & TournamentKey & ", " & _
                          " '" & strTeamID & "', '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                          " '" & FORMATSQL(Trim(txtTeamName.Text)) & "', 0)"
        
        TeamKey = 0
        s = "SELECT PK " & _
            " FROM tbl_Scoring_Team " & _
            " WHERE (TournamentKey = " & TournamentKey & ")" & _
            " AND (TeamID = '" & strTeamID & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            TeamKey = rs!PK
        End If
        rs.Close
        a = 0
        If CDbl(TeamKey) <> 0 Then
            For i = 1 To .Count
                If CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2))) <> 0 Then
                    a = a + 1
                    ConnOmega.Execute "INSERT INTO tbl_Scoring_Team_Detail " & _
                                      " (TeamKey, Line, PlayerKey) " & _
                                      " VALUES (" & TeamKey & ", " & a & ", " & _
                                      " " & .Item(i).SubItems(2) & ")"
                End If
            Next i
        End If
        
        dblTeamHDCP = 0: dblTeamHDCPInx = 0
        s = "SELECT TOP " & HandicapDivisor & " tbl_Scoring_PlayerName.HandiCap " & _
            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " Where (tbl_Scoring_Team_Detail.TeamKey = " & TeamKey & ") " & _
            " ORDER BY tbl_Scoring_PlayerName.HandiCap "
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        While Not rs.EOF
            dblTeamHDCP = dblTeamHDCP + CDbl(rs!Handicap)
            rs.MoveNext
        Wend
        rs.Close
        
        
        
        If CDbl(HandicapDivisor) > 0 Then
            dblTeamHDCPInx = Format(CDbl(dblTeamHDCP) / CDbl(HandicapDivisor), "#0.00")
            
            Arr = Split(dblTeamHDCPInx, ".", -1, 1)
            
            dblTeamHDCPInxtmp = Arr(1)
            
            iIndexRange = Mid(dblTeamHDCPInxtmp, 2, 1)
            
            If CDbl(iIndexRange) >= 0 And CDbl(iIndexRange) <= 4 Then
                iHDCPIndex = Arr(0) & "." & Mid(dblTeamHDCPInxtmp, 1, 1)
            Else
                iHDCPIndex = Arr(0) & "." & CStr(CDbl(Mid(dblTeamHDCPInxtmp, 1, 1)) + 1)
            End If
            
            ConnOmega.Execute "UPDATE tbl_Scoring_Team " & _
                              " SET TeamHDCP = " & CDbl(iHDCPIndex) & " " & _
                              " Where (PK = " & TeamKey & ")"
                              
            'ConnOmega.Execute "UPDATE tbl_Scoring_Team " & _
                              " SET TeamHDCP = " & CDbl(dblTeamHDCPInx) & " " & _
                              " Where (PK = " & TeamKey & ")"
        End If
        
        If TeamAverage = 2 Then
            iTeamIndex = 0
            If TeamDivisorOrder = 0 Then
                t = "SELECT TOP " & HandicapDivisor & " tbl_Scoring_PlayerName.HandiCap, " & _
                    " tbl_Scoring_PlayerName.iIndex " & _
                    " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                    " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                    " Where (tbl_Scoring_Team_Detail.TeamKey = " & TeamKey & ") " & _
                    " ORDER BY tbl_Scoring_PlayerName.iIndex"
            Else
                t = "SELECT TOP " & HandicapDivisor & " tbl_Scoring_PlayerName.HandiCap, " & _
                    " tbl_Scoring_PlayerName.iIndex " & _
                    " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                    " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                    " Where (tbl_Scoring_Team_Detail.TeamKey = " & TeamKey & ") " & _
                    " ORDER BY tbl_Scoring_PlayerName.iIndex DESC"
            End If
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            While Not rt.EOF
                iTeamIndex = iTeamIndex + CDbl(rt!iIndex)
                rt.MoveNext
            Wend
            rt.Close
            iTeamIndex = CDbl(iTeamIndex) / HandicapDivisor
            iTeamIndex = RoundOffIndex(CDbl(iTeamIndex))
            ConnOmega.Execute "UPDATE tbl_Scoring_Team " & _
                              " SET TeamIndex = " & CDbl(iTeamIndex) & " " & _
                              " WHERE (PK = " & TeamKey & ")"
        End If
        
    End If
    If TRANSACTIONTYPE = is_EDITTING Then
        TeamKey = Statusbar1.Panels(1).Text
        strTeamID = GetSetting(App.EXEName, "TeamControl", "TeamCtrl", "")
        If CDbl(AllowedTeam) = 1 Then
            For i = 1 To .Count
                If CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2))) <> 0 Then
                    
                    '== Check Teammate
                    For j = 1 To .Count
                        If CDbl(i) <> CDbl(j) Then
                            s = "SELECT TeamKey " & _
                                " From tbl_Scoring_Team_Detail " & _
                                " WHERE (PlayerKey = " & .Item(i).SubItems(2) & ") " & _
                                " AND (TeamKey <> " & TeamKey & ")"
                            If rs.State = adStateOpen Then rs.Close
                            rs.Open s, ConnOmega
                            If rs.RecordCount > 0 Then
                                MsgBox "'" & .Item(i).SubItems(3) & "' is not valid for this Team!                  ", vbCritical, "Error..."
                                rs.Close
                                Exit Function
                            End If
                            rs.Close
                        End If
                    Next j
                End If
            Next i
        Else
            For i = 1 To .Count
                If Trim(.Item(i).SubItems(2)) <> "" Then
                    iTeamPerPlayer = 1
                    t = "SELECT tbl_Scoring_Team_Detail.PlayerKey, tbl_Scoring_Team.TournamentKey " & _
                        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                        " tbl_Scoring_Team ON tbl_Scoring_Team_Detail.TeamKey = tbl_Scoring_Team.PK " & _
                        " WHERE (tbl_Scoring_Team.TournamentKey = " & TournamentKey & ") " & _
                        " AND (tbl_Scoring_Team.PK <> " & TeamKey & ") " & _
                        " AND (tbl_Scoring_Team_Detail.PlayerKey = " & .Item(i).SubItems(2) & ")"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        iTeamPerPlayer = iTeamPerPlayer + 1
                        rt.MoveNext
                    Wend
                    rt.Close
                End If
                If CDbl(AllowedTeam) < CDbl(iTeamPerPlayer) Then MsgBox "Number of Team per Player Already Exceed!                  ", vbCritical, "Error...": Exit Function
            Next i
        End If
        
        ConnOmega.Execute "UPDATE tbl_Scoring_Team " & _
                          " SET TeamName = '" & FORMATSQL(Trim(txtTeamName.Text)) & "', " & _
                          " TeamHDCP = 0, " & _
                          " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                          " Where (PK = " & TeamKey & ")"
        
        ConnOmega.Execute "DELETE FROM tbl_Scoring_Team_Detail WHERE (TeamKey = " & TeamKey & ")"
        For i = 1 To .Count
            If CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2))) <> 0 Then
                a = a + 1
                ConnOmega.Execute "INSERT INTO tbl_Scoring_Team_Detail " & _
                                  " (TeamKey, Line, PlayerKey) " & _
                                  " VALUES (" & TeamKey & ", " & a & ", " & _
                                  " " & .Item(i).SubItems(2) & ")"
            End If
        Next i
        
        dblTeamHDCP = 0: dblTeamHDCPInx = 0
        s = "SELECT TOP " & HandicapDivisor & " tbl_Scoring_PlayerName.HandiCap " & _
            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " Where (tbl_Scoring_Team_Detail.TeamKey = " & TeamKey & ") " & _
            " ORDER BY tbl_Scoring_PlayerName.HandiCap "
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        While Not rs.EOF
            dblTeamHDCP = dblTeamHDCP + CDbl(rs!Handicap)
            rs.MoveNext
        Wend
        rs.Close
        
        'If CDbl(HandicapDivisor) > 0 Then
        '    dblTeamHDCPInx = Format(CDbl(dblTeamHDCP) / CDbl(HandicapDivisor), "#0.0")
        '
        '    ConnOmega.Execute "UPDATE tbl_Scoring_Team " & _
        '                      " SET TeamHDCP = " & CDbl(dblTeamHDCPInx) & " " & _
        '                      " Where (PK = " & TeamKey & ")"
        'End If
        
        If CDbl(HandicapDivisor) > 0 Then
            dblTeamHDCPInx = Format(CDbl(dblTeamHDCP) / CDbl(HandicapDivisor), "#0.00")
            
            Arr = Split(dblTeamHDCPInx, ".", -1, 1)
            
            dblTeamHDCPInxtmp = Arr(1)
            
            iIndexRange = Mid(dblTeamHDCPInxtmp, 2, 1)
            
            If CDbl(iIndexRange) >= 0 And CDbl(iIndexRange) <= 4 Then
                iHDCPIndex = Arr(0) & "." & Mid(dblTeamHDCPInxtmp, 1, 1)
            Else
                iHDCPIndex = Arr(0) & "." & CStr(CDbl(Mid(dblTeamHDCPInxtmp, 1, 1)) + 1)
            End If
            
            ConnOmega.Execute "UPDATE tbl_Scoring_Team " & _
                              " SET TeamHDCP = " & CDbl(iHDCPIndex) & " " & _
                              " Where (PK = " & TeamKey & ")"
                              
            'ConnOmega.Execute "UPDATE tbl_Scoring_Team " & _
                              " SET TeamHDCP = " & CDbl(dblTeamHDCPInx) & " " & _
                              " Where (PK = " & TeamKey & ")"
        End If
        
'        txtTeamName.Locked = True
'        TOOLBARFUNC 1
'        TRANSACTIONTYPE = is_REFRESH
'        BROWSER GetSetting(App.EXEName, "TeamControl", "TeamCtrl", ""), "is_LOAD"
    End If
    
    'If TeamAverage = 1 Then
        
    If TeamAverage = 2 Then
        'MsgBox HandicapDivisor
        'MsgBox "pass 2"
        iTeamIndex = 0
        If TeamDivisorOrder = 0 Then
            'MsgBox "pass 0"
            t = "SELECT TOP " & HandicapDivisor & " tbl_Scoring_PlayerName.HandiCap, " & _
                " tbl_Scoring_PlayerName.iIndex " & _
                " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                " Where (tbl_Scoring_Team_Detail.TeamKey = " & TeamKey & ") " & _
                " ORDER BY tbl_Scoring_PlayerName.iIndex"
        Else
            t = "SELECT TOP " & HandicapDivisor & " tbl_Scoring_PlayerName.HandiCap, " & _
                " tbl_Scoring_PlayerName.iIndex " & _
                " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                " Where (tbl_Scoring_Team_Detail.TeamKey = " & TeamKey & ") " & _
                " ORDER BY tbl_Scoring_PlayerName.iIndex DESC"
        End If
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        While Not rt.EOF
            iTeamIndex = iTeamIndex + CDbl(rt!iIndex)
            rt.MoveNext
        Wend
        rt.Close
        iTeamIndex = CDbl(iTeamIndex) / HandicapDivisor
        iTeamIndex = RoundOffIndex(CDbl(iTeamIndex))
        ConnOmega.Execute "UPDATE tbl_Scoring_Team " & _
                          " SET TeamIndex = " & CDbl(iTeamIndex) & " " & _
                          " WHERE (PK = " & TeamKey & ")"
    End If
    txtTeamName.Locked = True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER strTeamID, "is_LOAD"
End With
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
    If picSLine.Visible = True Then
        If TRANSACTIONTYPE_DET = is_DET_ADDING Then
            With lstMember.ListItems
                If .Count = 1 Then
                    .Item(.Count).SubItems(1) = " "
                    .Item(.Count).SubItems(2) = " "
                    .Item(.Count).SubItems(3) = " "
                    .Item(.Count).SubItems(4) = " "
                    .Item(.Count).SubItems(5) = " "
                Else
                    .Remove .Count
                End If
            End With
            lstPlayerSearch.Visible = False
        ElseIf TRANSACTIONTYPE_DET = is_DET_EDITTING Then
            With lstMember.ListItems
                .Item(ROW).SubItems(2) = txtPlayerKey1.Text
                .Item(ROW).SubItems(3) = txtPlayerName1.Text
                .Item(ROW).SubItems(4) = txtHandicap1.Text
                .Item(ROW).SubItems(5) = txtClass1.Text
            End With
            lstPlayerSearch.Visible = False
        End If
        picMain.Enabled = True
        picToolbar.Enabled = True
        picSLine.Visible = False
        lstMember.SetFocus
    Else
        txtTeamName.Locked = True
        CLEARTEXT
        TOOLBARFUNC 1
        TRANSACTIONTYPE = is_REFRESH
        BROWSER GetSetting(App.EXEName, "TeamControl", "TeamCtrl", ""), "is_LOAD"
        If Trim(txtTeamID.Text) = "" Then BROWSER GetSetting(App.EXEName, "TeamControl", "TeamCtrl", ""), "is_HOME"
    End If
End If
End Function

Private Function CLEARTEXT()
txtTeamID.Text = ""
txtHandicap.Text = ""
txtClass.Text = ""
txtTeamName.Text = ""
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
If TeamAverage = 1 Then
    Label4.Caption = "HANDICAP"
    lstMember.ColumnHeaders(5).Text = "Handicap"
ElseIf TeamAverage = 2 Then
    Label4.Caption = "INDEX"
    lstMember.ColumnHeaders(5).Text = "Index"
End If
lstMember.ListItems.Clear
Set x = lstMember.ListItems.Add()
x.Text = ""
x.SubItems(1) = " "
x.SubItems(2) = " "
x.SubItems(3) = " "
x.SubItems(4) = " "
x.SubItems(5) = " "
x.SubItems(6) = " "
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

Private Sub btnSupplier_Click()
txtPlayer.SetFocus
End Sub

Private Sub cmbPlayer_Click()
If cmbPlayer.ListIndex = -1 Then Exit Sub
lstMember.ListItems.Item(ROW).SubItems(2) = cmbPlayer.ItemData(cmbPlayer.ListIndex)
lstMember.ListItems.Item(ROW).SubItems(3) = cmbPlayer.List(cmbPlayer.ListIndex)
s = "SELECT HandiCap, Class " & _
    " From tbl_Scoring_PlayerName " & _
    " WHERE (PK = " & cmbPlayer.ItemData(cmbPlayer.ListIndex) & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    lstMember.ListItems.Item(ROW).SubItems(4) = rs!Handicap
    lstMember.ListItems.Item(ROW).SubItems(5) = rs!Class
End If
rs.Close
dblHandicap = 0
With lstMember.ListItems
    For i = 1 To .Count
        If IsNumeric(.Item(i).SubItems(4)) = True Then
            a = a + 1
            dblHandicap = dblHandicap + CDbl(.Item(i).SubItems(4))
        End If
    Next i
End With

txtHandicap.Text = Mid(Format((CDbl(dblHandicap) / CDbl(a)), "#,##0.00"), 1, Len(Format((CDbl(dblHandicap) / CDbl(a)), "#,##0.00")) - 3)

End Sub

Private Sub cmbPlayer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If cmbPlayer.ListIndex = -1 Then Exit Sub
    With lstMember.ListItems
        For i = 1 To .Count
            If CDbl(i) <> CDbl(ROW) Then
                If CDbl(cmbPlayer.ItemData(cmbPlayer.ListIndex)) = CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2))) Then
                    MsgBox "Found Duplicate!                ", vbCritical, "Error..."
                    Exit Sub
                End If
            End If
        Next i
    End With
    picMain.Enabled = True
    picToolbar.Enabled = True
    picSLine.Visible = False
    lstMember.SetFocus
End If
End Sub

Private Sub cmdCancel_Click()
picSearch.Visible = False
picToolbar.Enabled = True
picMain.Enabled = True
End Sub

Private Sub cmdOK_Click()
If lstTeam.ListIndex = -1 Then Exit Sub
Array1 = Split(lstTeam.List(lstTeam.ListIndex), " - ", -1, 1)
BROWSER CStr(Array1(0)), "is_LOAD"
cmdCancel_Click
End Sub

Private Sub Command1_Click()
If TeamAverage = 2 Then
    Screen.MousePointer = vbHourglass
    s = "SELECT tbl_Scoring_Team.* " & _
        " FROM tbl_Scoring_Team " & _
        " WHERE (TournamentKey = " & TournamentKey & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        iTeamIndex = 0
        t = "SELECT TOP " & HandicapDivisor & " tbl_Scoring_PlayerName.HandiCap, " & _
            " tbl_Scoring_PlayerName.iIndex " & _
            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
            " ORDER BY tbl_Scoring_PlayerName.iIndex"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        While Not rt.EOF
            iTeamIndex = iTeamIndex + CDbl(rt!iIndex)
            rt.MoveNext
        Wend
        rt.Close
        
        iTeamIndex = CDbl(iTeamIndex) / HandicapDivisor
        iTeamIndex = RoundOffIndex(CDbl(iTeamIndex))
        ConnOmega.Execute "UPDATE tbl_Scoring_Team " & _
                          " SET TeamIndex = " & CDbl(iTeamIndex) & " " & _
                          " WHERE (PK = " & rs!PK & ")"
        rs.MoveNext
    Wend
    rs.Close
    Screen.MousePointer = vbDefault
End If
End Sub

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
If TRANSACTIONTYPE = is_REFRESH Then
    BROWSER GetSetting(App.EXEName, "TeamControl", "TeamCtrl", ""), "is_LOAD"
    If Trim(txtTeamID.Text) = "" Then BROWSER GetSetting(App.EXEName, "TeamControl", "TeamCtrl", ""), "is_HOME"
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert:   PRESS_INSERT
    Case vbKeyF2:       PRESS_F2
    Case vbKeyDelete:   PRESS_DELETE
    Case vbKeyF5:       PRESS_F5
    Case vbKeyF6:       PRESS_F6
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "TeamControl", "TeamCtrl", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "TeamControl", "TeamCtrl", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "TeamControl", "TeamCtrl", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "TeamControl", "TeamCtrl", ""), "is_END"
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

Select Case PointsToCnt
    Case 1: lstMember.ColumnHeaders(7).Text = "Gross Pts"
    Case 2: lstMember.ColumnHeaders(7).Text = "Net Pts"
    Case 3: If ScoringType = 2 Then lstMember.ColumnHeaders(7).Text = "Gross Pts" Else lstMember.ColumnHeaders(7).Text = ""
End Select

txtTeamName.Locked = True
'Me.Caption = "Team Setup"
CLEARTEXT
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANSACTIONTYPE_DET = is_DET_REFRESH
ListFocus = 0

'BROWSER GetSetting(App.EXEName, "TeamControl", "TeamCtrl", ""), "is_LOAD"
'    If Trim(txtTeamID.Text) = "" Then BROWSER GetSetting(App.EXEName, "TeamControl", "TeamCtrl", ""), "is_HOME"
    

tmp = SetWindowLong(txtSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPlayer.hwnd, GWL_STYLE, GetWindowLong(txtPlayer.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtTeamName.hwnd, GWL_STYLE, GetWindowLong(txtTeamName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub lstMember_GotFocus()
ListFocus = 1
ROW = lstMember.SelectedItem.Index
TRANSACTIONTYPE_DET = is_DET_REFRESH
End Sub

Private Sub lstMember_ItemClick(ByVal Item As MSComctlLib.ListItem)
ROW = lstMember.SelectedItem.Index
End Sub

Private Sub lstMember_LostFocus()
ListFocus = 0
End Sub

Private Sub lstPlayerSearch_Click()
If lstPlayerSearch.ListIndex = -1 Then Exit Sub
If SearchFocus = 1 Then Exit Sub
txtPlayerKey.Text = lstPlayerSearch.ItemData(lstPlayerSearch.ListIndex)
txtPlayer.Text = lstPlayerSearch.List(lstPlayerSearch.ListIndex)
End Sub

Private Sub lstPlayerSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtPlayerKey.Text = lstPlayerSearch.ItemData(lstPlayerSearch.ListIndex)
    txtPlayer.Text = lstPlayerSearch.List(lstPlayerSearch.ListIndex)
    txtPlayer.SetFocus
    lstPlayerSearch.Visible = False
End If
End Sub

Private Sub lstResult_Click()
If lstResult.ListIndex = -1 Then lstTeam.Clear: Exit Sub
lstTeam.Clear
s = "SELECT tbl_Scoring_Team.PK " & _
    " FROM tbl_Scoring_Team LEFT OUTER JOIN " & _
    " tbl_Scoring_Team_Detail ON tbl_Scoring_Team.PK = tbl_Scoring_Team_Detail.TeamKey LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_Team_Detail.PlayerKey = " & lstResult.ItemData(lstResult.ListIndex) & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    t = "SELECT tbl_Scoring_Team.PK, " & _
        " tbl_Scoring_Team.TeamID, " & _
        " tbl_Scoring_PlayerName.LastName, " & _
        " tbl_Scoring_PlayerName.FirstName " & _
        " FROM tbl_Scoring_Team LEFT OUTER JOIN " & _
        " tbl_Scoring_Team_Detail ON tbl_Scoring_Team.PK = tbl_Scoring_Team_Detail.TeamKey LEFT OUTER JOIN " & _
        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
        " WHERE (tbl_Scoring_Team.PK = " & rs!PK & ") " & _
        " AND (tbl_Scoring_Team_Detail.PlayerKey <> " & lstResult.ItemData(lstResult.ListIndex) & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        lstTeam.AddItem rt!TeamID & " - " & rt!LastName & ", " & rt!FirstName
        lstTeam.ItemData(lstTeam.NewIndex) = rt!PK
        rt.MoveNext
    Wend
    rt.Close
    rs.MoveNext
Wend
rs.Close
If lstTeam.ListCount Then lstTeam.ListIndex = 0
End Sub

Private Sub lstResult_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstTeam.SetFocus
End Sub

Private Sub lstTeam_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "TeamControl", "TeamCtrl", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "TeamControl", "TeamCtrl", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "TeamControl", "TeamCtrl", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "TeamControl", "TeamCtrl", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Close":   PRESS_ESCAPE
    Case Else:      Exit Sub
End Select
End Sub

Private Sub txtHandicap_Change()
's = "SELECT Class " & _
'    " From tbl_Scoring_TournamentInfo_Class " & _
'    " WHERE (TournamentKey = " & TournamentKey & ") " & _
'    " AND (HFrom <= " & RETURNTEXTVALUE(txtHandicap) & ") " & _
'    " AND (HTo >= " & RETURNTEXTVALUE(txtHandicap) & ")"
'If rs.State = adStateOpen Then rs.Close
'rs.Open s, ConnOmega
'If rs.RecordCount > 0 Then
'    txtClass.Text = rs!Class
'Else
'    txtClass.Text = ""
'End If
'rs.Close
End Sub

Private Sub txtPlayer_Change()
If SearchFocus = 1 Then
    If Trim(txtPlayer.Text) = "" Then lstPlayerSearch.Visible = False: lstPlayerSearch.Clear: Exit Sub
    lstPlayerSearch.ZOrder 0
    lstPlayerSearch.Visible = True
    lstPlayerSearch.Clear
    s = "SELECT PK, LTRIM(RTRIM(LastName)) + ',  ' + LTRIM(RTRIM(FirstName)) + '  ' + LTRIM(RTRIM(MiddleName)) AS PlayerName " & _
        " From tbl_Scoring_PlayerName " & _
        " WHERE (LTRIM(RTRIM(LastName)) LIKE '" & FORMATSQL(Trim(txtPlayer.Text)) & "%') " & _
        " AND (TournamentKey = " & TournamentKey & ") " & _
        " ORDER BY LTRIM(RTRIM(LastName)) + ',  ' + LTRIM(RTRIM(FirstName)) + '  ' + LTRIM(RTRIM(MiddleName))"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        lstPlayerSearch.AddItem rs!PlayerName
        lstPlayerSearch.ItemData(lstPlayerSearch.NewIndex) = rs!PK
        rs.MoveNext
    Wend
    rs.Close
    If lstPlayerSearch.ListCount Then lstPlayerSearch.ListIndex = 0
End If
End Sub

Private Sub txtPlayer_GotFocus()
SearchFocus = 1
HTEXT txtPlayer
End Sub

Private Sub txtPlayer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
    lstPlayerSearch.SetFocus
End If
If KeyCode = vbKeyReturn Then
    If lstPlayerSearch.Visible = True Then
        lstPlayerSearch.SetFocus
        Exit Sub
    End If
    If RETURNTEXTVALUE(txtPlayerKey) > 0 Then
        lstMember.ListItems.Item(ROW).SubItems(2) = RETURNTEXTVALUE(txtPlayerKey)
        lstMember.ListItems.Item(ROW).SubItems(3) = txtPlayer.Text
        t = "SELECT HandiCap " & _
            " From tbl_Scoring_PlayerName " & _
            " WHERE (TournamentKey = " & TournamentKey & ") " & _
            " AND (PK = " & RETURNTEXTVALUE(txtPlayerKey) & ")"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            lstMember.ListItems.Item(ROW).SubItems(4) = rt!Handicap
            u = "SELECT Class " & _
                " From tbl_Scoring_TournamentInfo_Class " & _
                " Where (TournamentKey = " & TournamentKey & ") " & _
                " And (HFrom <= " & CDbl(rt!Handicap) & ") " & _
                " And (HTo >= " & CDbl(rt!Handicap) & ")"
            If ru.State = adStateOpen Then ru.Close
            ru.Open u, ConnOmega
            If ru.RecordCount > 0 Then
                lstMember.ListItems.Item(ROW).SubItems(5) = ru!Class
            Else
                lstMember.ListItems.Item(ROW).SubItems(5) = " "
            End If
        End If
        rt.Close
        picMain.Enabled = True
        picToolbar.Enabled = True
        picSLine.Visible = False
        lstMember.SetFocus
    End If
End If
End Sub

Private Sub txtPlayer_LostFocus()
SearchFocus = 0
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstResult.Clear: lstTeam.Clear: Exit Sub
lstResult.Clear: lstTeam.Clear
s = "SELECT tbl_Scoring_Team_Detail.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
    " tbl_Scoring_PlayerName.MiddleName" & _
    " FROM tbl_Scoring_Team LEFT OUTER JOIN " & _
    " tbl_Scoring_Team_Detail ON tbl_Scoring_Team.PK = tbl_Scoring_Team_Detail.TeamKey LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " Where (tbl_Scoring_PlayerName.TournamentKey = " & TournamentKey & ") " & _
    " GROUP BY tbl_Scoring_Team_Detail.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
    " tbl_Scoring_PlayerName.MiddleName " & _
    " HAVING (tbl_Scoring_PlayerName.LastName LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
    " ORDER BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResult.AddItem rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
    lstResult.ItemData(lstResult.NewIndex) = rs!PlayerKey
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
