VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmScoreCardTeamOnly 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12105
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmScoreCardTeamOnly.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picProgress 
      Height          =   975
      Left            =   3840
      TabIndex        =   142
      Top             =   3000
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1720
      BackColor       =   13023396
      Begin VB.PictureBox picProgressBar 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5235
         TabIndex        =   143
         Top             =   120
         Width           =   5295
      End
   End
   Begin RPVGCC.b8Container picPrint 
      Height          =   1695
      Left            =   4440
      TabIndex        =   125
      Top             =   480
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2990
      BackColor       =   15396057
      Begin VB.CommandButton cmdOKPrint 
         Height          =   480
         Left            =   480
         Picture         =   "frmScoreCardTeamOnly.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   128
         Top             =   960
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelPrint 
         Height          =   480
         Left            =   2280
         Picture         =   "frmScoreCardTeamOnly.frx":0F3C
         Style           =   1  'Graphical
         TabIndex        =   127
         Top             =   960
         Width           =   1560
      End
      Begin VB.PictureBox picElse 
         BackColor       =   &H00EAECD9&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   4095
         TabIndex        =   135
         Top             =   840
         Width           =   4095
         Begin VB.TextBox txtDateTo 
            Height          =   315
            Left            =   2400
            TabIndex        =   138
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtDateFrom 
            Height          =   315
            Left            =   840
            TabIndex        =   137
            Top             =   1320
            Width           =   1335
         End
         Begin VB.ComboBox cmbSortAll 
            Height          =   315
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   136
            Top             =   0
            Width           =   4095
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            Height          =   255
            Left            =   2160
            TabIndex        =   140
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   255
            Left            =   360
            TabIndex        =   139
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.PictureBox picScoreCard 
         BackColor       =   &H00EAECD9&
         BorderStyle     =   0  'None
         Height          =   3495
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   4095
         TabIndex        =   130
         Top             =   840
         Width           =   4095
         Begin VB.ComboBox cmbDateScore 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   134
            Top             =   3120
            Width           =   1695
         End
         Begin VB.ListBox lstResultScore 
            Height          =   1425
            Left            =   0
            TabIndex        =   133
            Top             =   360
            Width           =   4095
         End
         Begin VB.ListBox lstResultScoreTeam 
            Height          =   1230
            Left            =   0
            TabIndex        =   132
            Top             =   1830
            Width           =   4095
         End
         Begin VB.TextBox txtSearchScore 
            Height          =   315
            Left            =   0
            TabIndex        =   131
            Top             =   0
            Width           =   4095
         End
      End
      Begin VB.ComboBox cmbPrintType 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   129
         Top             =   480
         Width           =   4095
      End
      Begin RPVGCC.b8TitleBar b8TitleBar3 
         Height          =   345
         Left            =   40
         TabIndex        =   126
         Top             =   40
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   609
         Caption         =   "Print"
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
         Icon            =   "frmScoreCardTeamOnly.frx":1698
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2760
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483648
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardTeamOnly.frx":1C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardTeamOnly.frx":1D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardTeamOnly.frx":1EB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardTeamOnly.frx":21D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardTeamOnly.frx":258B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardTeamOnly.frx":29DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardTeamOnly.frx":2E2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardTeamOnly.frx":31E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardTeamOnly.frx":32F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardTeamOnly.frx":383B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardTeamOnly.frx":3995
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardTeamOnly.frx":3ED7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   109
      Top             =   6510
      Width           =   12105
      _ExtentX        =   21352
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
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   0
      Picture         =   "frmScoreCardTeamOnly.frx":40FB
      ScaleHeight     =   6255
      ScaleWidth      =   12135
      TabIndex        =   0
      Top             =   0
      Width           =   12135
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   9600
         Top             =   2040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   615
         Left            =   10080
         TabIndex        =   141
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox picToolbar 
         BorderStyle     =   0  'None
         Height          =   770
         Left            =   0
         ScaleHeight     =   765
         ScaleWidth      =   15000
         TabIndex        =   2
         Top             =   0
         Width           =   15000
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   570
            Left            =   0
            TabIndex        =   3
            Top             =   105
            Width           =   15000
            _ExtentX        =   26458
            _ExtentY        =   1005
            ButtonWidth     =   1058
            ButtonHeight    =   1005
            Style           =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   20
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
                  Caption         =   "Close"
                  Key             =   "Close"
                  ImageIndex      =   10
               EndProperty
               BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
            EndProperty
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            X1              =   0
            X2              =   15000
            Y1              =   690
            Y2              =   690
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00C0C0C0&
            X1              =   0
            X2              =   15000
            Y1              =   750
            Y2              =   750
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
            Y1              =   650
            Y2              =   650
         End
      End
      Begin VB.TextBox txtCtrl 
         Height          =   285
         Left            =   10200
         TabIndex        =   1
         Top             =   1800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin RPVGCC.b8Container b8Container3 
         Height          =   1095
         Left            =   5880
         TabIndex        =   4
         Top             =   840
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   1931
         BackColor       =   49152
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   855
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   5895
            TabIndex        =   5
            Top             =   120
            Width           =   5895
            Begin VB.TextBox txtTourDate 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1095
               TabIndex        =   7
               Text            =   "06/01/2010 - 06/04/2010"
               Top             =   480
               Width           =   4695
            End
            Begin VB.TextBox txtTournament 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1095
               TabIndex        =   6
               Top             =   120
               Width           =   4695
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Date Range"
               Height          =   255
               Left            =   120
               TabIndex        =   9
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Tournament"
               Height          =   255
               Left            =   120
               TabIndex        =   8
               Top             =   120
               Width           =   1335
            End
         End
      End
      Begin RPVGCC.b8Container b8Container5 
         Height          =   2535
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   4471
         BackColor       =   49152
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   2295
            Left            =   120
            ScaleHeight     =   2295
            ScaleWidth      =   5415
            TabIndex        =   11
            Top             =   120
            Width           =   5415
            Begin VB.TextBox txtTeamHandicap 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   4560
               TabIndex        =   119
               Top             =   480
               Width           =   735
            End
            Begin MSComctlLib.ListView lstPlayer 
               Height          =   1335
               Left            =   120
               TabIndex        =   118
               Top             =   840
               Width           =   5175
               _ExtentX        =   9128
               _ExtentY        =   2355
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   0   'False
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   14737632
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   4
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Player Name"
                  Object.Width           =   5292
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   2
                  Text            =   "Handicap"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   3
                  Text            =   "Class"
                  Object.Width           =   1411
               EndProperty
            End
            Begin VB.TextBox txtPlayer 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   720
               TabIndex        =   14
               Top             =   480
               Width           =   1575
            End
            Begin VB.TextBox txtDay 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   4560
               TabIndex        =   13
               Top             =   120
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox txtDate 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   720
               TabIndex        =   12
               Top             =   120
               Width           =   1575
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Handicap"
               Height          =   255
               Left            =   3720
               TabIndex        =   120
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Team"
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Day"
               Height          =   255
               Left            =   3720
               TabIndex        =   16
               Top             =   120
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   120
               Width           =   495
            End
         End
      End
      Begin RPVGCC.b8Container b8Container2 
         Height          =   1335
         Left            =   5880
         TabIndex        =   18
         Top             =   2040
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   2355
         BackColor       =   49152
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1095
            Left            =   120
            ScaleHeight     =   1095
            ScaleWidth      =   3375
            TabIndex        =   19
            Top             =   120
            Width           =   3375
            Begin VB.TextBox txtNetScore 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   2400
               TabIndex        =   121
               Top             =   600
               Width           =   855
            End
            Begin VB.TextBox txtSGrossTot 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1440
               TabIndex        =   22
               Top             =   600
               Width           =   855
            End
            Begin VB.TextBox txtScoreGrossB 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   720
               TabIndex        =   21
               Top             =   600
               Width           =   495
            End
            Begin VB.TextBox txtScoreGrossF 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   120
               TabIndex        =   20
               Top             =   600
               Width           =   495
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Net Score"
               Height          =   255
               Left            =   2400
               TabIndex        =   122
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "B - 9"
               Height          =   255
               Left            =   720
               TabIndex        =   26
               Top             =   360
               Width           =   495
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "F - 9"
               Height          =   255
               Left            =   120
               TabIndex        =   25
               Top             =   360
               Width           =   495
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Scores"
               BeginProperty Font 
                  Name            =   "Garamond"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   120
               TabIndex        =   24
               Top             =   0
               Width           =   975
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Gross Score"
               Height          =   255
               Left            =   1440
               TabIndex        =   23
               Top             =   360
               Width           =   855
            End
         End
      End
      Begin RPVGCC.b8Container b8Container1 
         Height          =   2490
         Left            =   120
         TabIndex        =   27
         Top             =   3480
         Width           =   11865
         _ExtentX        =   20929
         _ExtentY        =   4392
         BackColor       =   16185592
         ShadowColor1    =   49152
         ShadowColor2    =   8454016
         Begin VB.PictureBox picScoreMain 
            Appearance      =   0  'Flat
            BackColor       =   &H00C6B8A4&
            ForeColor       =   &H80000008&
            Height          =   2400
            Left            =   50
            ScaleHeight     =   2370
            ScaleWidth      =   11745
            TabIndex        =   28
            Top             =   50
            Width           =   11780
            Begin VB.PictureBox picScoreDis 
               Appearance      =   0  'Flat
               BackColor       =   &H00C6B8A4&
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   1680
               Left            =   -10
               ScaleHeight     =   1650
               ScaleWidth      =   12300
               TabIndex        =   99
               Top             =   -10
               Width           =   12330
               Begin MSFlexGridLib.MSFlexGrid FGrid 
                  Height          =   2025
                  Left            =   -105
                  TabIndex        =   100
                  Top             =   -30
                  Width           =   12450
                  _ExtentX        =   21960
                  _ExtentY        =   3572
                  _Version        =   393216
                  BackColor       =   13023396
                  ForeColor       =   0
                  BackColorFixed  =   13023396
                  ForeColorFixed  =   0
                  BackColorSel    =   16777215
                  ForeColorSel    =   0
                  BackColorBkg    =   13023396
                  FocusRect       =   0
                  GridLinesFixed  =   1
                  Appearance      =   0
               End
            End
            Begin VB.PictureBox picScoreEn 
               Appearance      =   0  'Flat
               BackColor       =   &H00C6B8A4&
               ForeColor       =   &H80000008&
               Height          =   975
               Left            =   -10
               ScaleHeight     =   945
               ScaleWidth      =   12300
               TabIndex        =   29
               Top             =   1650
               Width           =   12330
               Begin VB.PictureBox Picture2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   495
                  Left            =   1980
                  ScaleHeight     =   465
                  ScaleWidth      =   10305
                  TabIndex        =   53
                  Top             =   230
                  Width           =   10335
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFF80&
                     Height          =   255
                     Index           =   0
                     Left            =   -10
                     TabIndex        =   95
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0C000&
                     Height          =   255
                     Index           =   0
                     Left            =   -10
                     TabIndex        =   94
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFF80&
                     Height          =   255
                     Index           =   1
                     Left            =   440
                     TabIndex        =   93
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0C000&
                     Height          =   255
                     Index           =   1
                     Left            =   440
                     TabIndex        =   92
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFF80&
                     Height          =   255
                     Index           =   2
                     Left            =   890
                     TabIndex        =   91
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0C000&
                     Height          =   255
                     Index           =   2
                     Left            =   890
                     TabIndex        =   90
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFF80&
                     Height          =   255
                     Index           =   3
                     Left            =   1340
                     TabIndex        =   89
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0C000&
                     Height          =   255
                     Index           =   3
                     Left            =   1340
                     TabIndex        =   88
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFF80&
                     Height          =   255
                     Index           =   4
                     Left            =   1780
                     TabIndex        =   87
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0C000&
                     Height          =   255
                     Index           =   4
                     Left            =   1780
                     TabIndex        =   86
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFF80&
                     Height          =   255
                     Index           =   5
                     Left            =   2240
                     TabIndex        =   85
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0C000&
                     Height          =   255
                     Index           =   5
                     Left            =   2240
                     TabIndex        =   84
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFF80&
                     Height          =   255
                     Index           =   6
                     Left            =   2680
                     TabIndex        =   83
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0C000&
                     Height          =   255
                     Index           =   6
                     Left            =   2680
                     TabIndex        =   82
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFF80&
                     Height          =   255
                     Index           =   7
                     Left            =   3130
                     TabIndex        =   81
                     Text            =   "0"
                     Top             =   -15
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0C000&
                     Height          =   255
                     Index           =   7
                     Left            =   3130
                     TabIndex        =   80
                     Text            =   "0"
                     Top             =   225
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFF80&
                     Height          =   255
                     Index           =   8
                     Left            =   3580
                     TabIndex        =   79
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0C000&
                     Height          =   255
                     Index           =   8
                     Left            =   3580
                     TabIndex        =   78
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFF80&
                     Height          =   255
                     Index           =   9
                     Left            =   4590
                     TabIndex        =   77
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0C000&
                     Height          =   255
                     Index           =   9
                     Left            =   4590
                     TabIndex        =   76
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFF80&
                     Height          =   255
                     Index           =   10
                     Left            =   5040
                     TabIndex        =   75
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0C000&
                     Height          =   255
                     Index           =   10
                     Left            =   5040
                     TabIndex        =   74
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFF80&
                     Height          =   255
                     Index           =   11
                     Left            =   5490
                     TabIndex        =   73
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0C000&
                     Height          =   255
                     Index           =   11
                     Left            =   5490
                     TabIndex        =   72
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFF80&
                     Height          =   255
                     Index           =   12
                     Left            =   5940
                     TabIndex        =   71
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0C000&
                     Height          =   255
                     Index           =   12
                     Left            =   5940
                     TabIndex        =   70
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFF80&
                     Height          =   255
                     Index           =   13
                     Left            =   6390
                     TabIndex        =   69
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0C000&
                     Height          =   255
                     Index           =   13
                     Left            =   6390
                     TabIndex        =   68
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFF80&
                     Height          =   255
                     Index           =   14
                     Left            =   6840
                     TabIndex        =   67
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0C000&
                     Height          =   255
                     Index           =   14
                     Left            =   6840
                     TabIndex        =   66
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFF80&
                     Height          =   255
                     Index           =   15
                     Left            =   7290
                     TabIndex        =   65
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0C000&
                     Height          =   255
                     Index           =   15
                     Left            =   7290
                     TabIndex        =   64
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFF80&
                     Height          =   255
                     Index           =   16
                     Left            =   7740
                     TabIndex        =   63
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0C000&
                     Height          =   255
                     Index           =   16
                     Left            =   7740
                     TabIndex        =   62
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFF80&
                     Height          =   255
                     Index           =   17
                     Left            =   8190
                     TabIndex        =   61
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0C000&
                     Height          =   255
                     Index           =   17
                     Left            =   8190
                     TabIndex        =   60
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPtsF 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Left            =   4040
                     TabIndex        =   59
                     Text            =   "0"
                     Top             =   -10
                     Width           =   570
                  End
                  Begin VB.TextBox txtNetPtsF 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Left            =   4040
                     TabIndex        =   58
                     Text            =   "0"
                     Top             =   230
                     Width           =   570
                  End
                  Begin VB.TextBox txtGrossPtsB 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
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
                     Left            =   8640
                     TabIndex        =   57
                     Text            =   "0"
                     Top             =   -10
                     Width           =   570
                  End
                  Begin VB.TextBox txtNetPtsB 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
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
                     Left            =   8640
                     TabIndex        =   56
                     Text            =   "0"
                     Top             =   230
                     Width           =   570
                  End
                  Begin VB.TextBox txtGrossPtsTot 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
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
                     Left            =   9200
                     TabIndex        =   55
                     Text            =   "0"
                     Top             =   -10
                     Width           =   570
                  End
                  Begin VB.TextBox txtNetPtsTot 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
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
                     Left            =   9200
                     TabIndex        =   54
                     Text            =   "0"
                     Top             =   230
                     Width           =   570
                  End
               End
               Begin VB.TextBox txtGrossScore 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Index           =   0
                  Left            =   1980
                  TabIndex        =   52
                  Text            =   "0"
                  Top             =   0
                  Width           =   460
               End
               Begin VB.TextBox txtGrossScore 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Index           =   1
                  Left            =   2430
                  TabIndex        =   51
                  Text            =   "0"
                  Top             =   -10
                  Width           =   460
               End
               Begin VB.TextBox txtGrossScore 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Index           =   2
                  Left            =   2880
                  TabIndex        =   50
                  Text            =   "0"
                  Top             =   -10
                  Width           =   460
               End
               Begin VB.TextBox txtGrossScore 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Index           =   3
                  Left            =   3330
                  TabIndex        =   49
                  Text            =   "0"
                  Top             =   -10
                  Width           =   460
               End
               Begin VB.TextBox txtGrossScore 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Index           =   4
                  Left            =   3780
                  TabIndex        =   48
                  Text            =   "0"
                  Top             =   -10
                  Width           =   460
               End
               Begin VB.TextBox txtGrossScore 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Index           =   5
                  Left            =   4230
                  TabIndex        =   47
                  Text            =   "0"
                  Top             =   -10
                  Width           =   460
               End
               Begin VB.TextBox txtGrossScore 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Index           =   6
                  Left            =   4680
                  TabIndex        =   46
                  Text            =   "0"
                  Top             =   -10
                  Width           =   460
               End
               Begin VB.TextBox txtGrossScore 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Index           =   7
                  Left            =   5130
                  TabIndex        =   45
                  Text            =   "0"
                  Top             =   -10
                  Width           =   460
               End
               Begin VB.TextBox txtGrossScore 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Index           =   8
                  Left            =   5580
                  TabIndex        =   44
                  Text            =   "0"
                  Top             =   -10
                  Width           =   460
               End
               Begin VB.TextBox txtGrossScore 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Index           =   17
                  Left            =   10180
                  TabIndex        =   43
                  Text            =   "0"
                  Top             =   -10
                  Width           =   460
               End
               Begin VB.TextBox txtGrossScore 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Index           =   16
                  Left            =   9730
                  TabIndex        =   42
                  Text            =   "0"
                  Top             =   -10
                  Width           =   460
               End
               Begin VB.TextBox txtGrossScore 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Index           =   15
                  Left            =   9280
                  TabIndex        =   41
                  Text            =   "0"
                  Top             =   -10
                  Width           =   460
               End
               Begin VB.TextBox txtGrossScore 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Index           =   14
                  Left            =   8830
                  TabIndex        =   40
                  Text            =   "0"
                  Top             =   -10
                  Width           =   460
               End
               Begin VB.TextBox txtGrossScore 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Index           =   13
                  Left            =   8380
                  TabIndex        =   39
                  Text            =   "0"
                  Top             =   -10
                  Width           =   460
               End
               Begin VB.TextBox txtGrossScore 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Index           =   12
                  Left            =   7930
                  TabIndex        =   38
                  Text            =   "0"
                  Top             =   -10
                  Width           =   460
               End
               Begin VB.TextBox txtGrossScore 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Index           =   11
                  Left            =   7480
                  TabIndex        =   37
                  Text            =   "0"
                  Top             =   -10
                  Width           =   460
               End
               Begin VB.TextBox txtGrossScore 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Index           =   10
                  Left            =   7030
                  TabIndex        =   36
                  Text            =   "0"
                  Top             =   -10
                  Width           =   460
               End
               Begin VB.TextBox txtGrossScore 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Index           =   9
                  Left            =   6580
                  TabIndex        =   35
                  Text            =   "0"
                  Top             =   -10
                  Width           =   460
               End
               Begin VB.PictureBox Picture3 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   6030
                  ScaleHeight     =   225
                  ScaleWidth      =   540
                  TabIndex        =   33
                  Top             =   -10
                  Width           =   570
                  Begin VB.TextBox txtGrossScoreF 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
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
                     Left            =   -10
                     TabIndex        =   34
                     Text            =   "0"
                     Top             =   -10
                     Width           =   570
                  End
               End
               Begin VB.PictureBox Picture4 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   10640
                  ScaleHeight     =   225
                  ScaleWidth      =   1860
                  TabIndex        =   30
                  Top             =   -10
                  Width           =   1890
                  Begin VB.TextBox txtGrossScoreB 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
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
                     Left            =   -10
                     TabIndex        =   32
                     Text            =   "0"
                     Top             =   -10
                     Width           =   570
                  End
                  Begin VB.TextBox txtGrossScoreTot 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
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
                     Left            =   540
                     TabIndex        =   31
                     Text            =   "0"
                     Top             =   -10
                     Width           =   570
                  End
               End
               Begin VB.Label Label1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " GROSS SCORE"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   -15
                  TabIndex        =   98
                  Top             =   -15
                  Width           =   2010
               End
               Begin VB.Label Label2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " GROSS POINTS"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   -10
                  TabIndex        =   97
                  Top             =   230
                  Width           =   2010
               End
               Begin VB.Label Label3 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " NET POINTS"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   -10
                  TabIndex        =   96
                  Top             =   470
                  Width           =   2010
               End
            End
         End
      End
   End
   Begin RPVGCC.b8Container picSearch 
      Height          =   4935
      Left            =   4440
      TabIndex        =   101
      Top             =   480
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   8705
      BackColor       =   15396057
      Begin VB.ListBox lstTeam 
         Height          =   1230
         Left            =   120
         TabIndex        =   123
         Top             =   2400
         Width           =   4095
      End
      Begin VB.CommandButton cmdCancelSearch 
         Height          =   480
         Left            =   2280
         Picture         =   "frmScoreCardTeamOnly.frx":3B92E
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   4200
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKSearch 
         Height          =   480
         Left            =   480
         Picture         =   "frmScoreCardTeamOnly.frx":3C08A
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   4200
         Width           =   1560
      End
      Begin VB.ListBox lstResult 
         Height          =   1425
         Left            =   120
         TabIndex        =   104
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   103
         Top             =   480
         Width           =   4095
      End
      Begin VB.ComboBox cmbDate 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   102
         Top             =   3720
         Width           =   1695
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   107
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
         Icon            =   "frmScoreCardTeamOnly.frx":3C6FC
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   255
         Left            =   1920
         TabIndex        =   108
         Top             =   3720
         Width           =   495
      End
   End
   Begin RPVGCC.b8Container picSearchAdd 
      Height          =   4935
      Left            =   4440
      TabIndex        =   110
      Top             =   480
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   8705
      BackColor       =   15396057
      Begin VB.ListBox lstTeamAdd 
         Height          =   1230
         Left            =   120
         TabIndex        =   124
         Top             =   2400
         Width           =   4095
      End
      Begin VB.TextBox txtSearchAdd 
         Height          =   315
         Left            =   120
         TabIndex        =   115
         Top             =   480
         Width           =   4095
      End
      Begin VB.ListBox lstResultAdd 
         Height          =   1425
         Left            =   120
         TabIndex        =   114
         Top             =   840
         Width           =   4095
      End
      Begin VB.CommandButton cmdOKAdd 
         Height          =   480
         Left            =   480
         Picture         =   "frmScoreCardTeamOnly.frx":3CC96
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   4200
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelAdd 
         Height          =   480
         Left            =   2280
         Picture         =   "frmScoreCardTeamOnly.frx":3D308
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   4200
         Width           =   1560
      End
      Begin VB.TextBox txtDateAdd 
         Height          =   315
         Left            =   1800
         TabIndex        =   111
         Top             =   3720
         Width           =   1215
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   45
         TabIndex        =   116
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
         Icon            =   "frmScoreCardTeamOnly.frx":3DA64
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   255
         Left            =   1320
         TabIndex        =   117
         Top             =   3720
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmScoreCardTeamOnly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public TournamentKey    As Double
Public TourNoOfPlays    As Double
'Dim PlayerKey           As Double
Dim TeamKey             As Double
Dim s                   As String
Dim rs                  As New ADODB.Recordset
Dim t                   As String
Dim rt                  As New ADODB.Recordset

Dim TRANSACTIONTYPE     As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim dDateEnd

Private Function BROWSER(strCtrl, isAction As String)
Dim i, x, dblHandicap
Select Case isAction
    Case "is_LOAD"
        If strCtrl <> "" Then
            s = "SELECT TOP 1 tbl_Scoring_ScoreCard_Team.PK, tbl_Scoring_ScoreCard_Team.CtrlNo, " & _
                " tbl_Scoring_ScoreCard_Team.TeamKey, tbl_Scoring_Team.TeamID AS PlayerName, " & _
                " tbl_Scoring_ScoreCard_Team.DDate, tbl_Scoring_ScoreCard_Team.Score, " & _
                " tbl_Scoring_ScoreCard_Team.Front9Gross, tbl_Scoring_ScoreCard_Team.Back9Gross, tbl_Scoring_ScoreCard_Team.GrossPoints, " & _
                " tbl_Scoring_ScoreCard_Team.Front9Net, tbl_Scoring_ScoreCard_Team.Back9Net, tbl_Scoring_ScoreCard_Team.NetPoints, " & _
                " tbl_Scoring_ScoreCard_Team.LastModified, tbl_Scoring_ScoreCard_Team.Front9Score, " & _
                " tbl_Scoring_ScoreCard_Team.Back9Score " & _
                " FROM tbl_Scoring_ScoreCard_Team LEFT OUTER JOIN " & _
                " tbl_Scoring_Team ON tbl_Scoring_ScoreCard_Team.TeamKey = tbl_Scoring_Team.PK " & _
                " WHERE (tbl_Scoring_ScoreCard_Team.TournamentKey = " & TournamentKey & ") " & _
                " AND (tbl_Scoring_ScoreCard_Team.CtrlNo = '" & strCtrl & "') " & _
                " ORDER BY tbl_Scoring_ScoreCard_Team.CtrlNo"
        Else
            s = "SELECT TOP 1 tbl_Scoring_ScoreCard_Team.PK, tbl_Scoring_ScoreCard_Team.CtrlNo, " & _
                " tbl_Scoring_ScoreCard_Team.TeamKey, tbl_Scoring_Team.TeamID AS PlayerName, " & _
                " tbl_Scoring_ScoreCard_Team.DDate, tbl_Scoring_ScoreCard_Team.Score, " & _
                " tbl_Scoring_ScoreCard_Team.Front9Gross, tbl_Scoring_ScoreCard_Team.Back9Gross, tbl_Scoring_ScoreCard_Team.GrossPoints, " & _
                " tbl_Scoring_ScoreCard_Team.Front9Net, tbl_Scoring_ScoreCard_Team.Back9Net, tbl_Scoring_ScoreCard_Team.NetPoints, " & _
                " tbl_Scoring_ScoreCard_Team.LastModified, tbl_Scoring_ScoreCard_Team.Front9Score, " & _
                " tbl_Scoring_ScoreCard_Team.Back9Score " & _
                " FROM tbl_Scoring_ScoreCard_Team LEFT OUTER JOIN " & _
                " tbl_Scoring_Team ON tbl_Scoring_ScoreCard_Team.TeamKey = tbl_Scoring_Team.PK " & _
                " WHERE (tbl_Scoring_ScoreCard_Team.TournamentKey = " & TournamentKey & ") " & _
                " ORDER BY tbl_Scoring_ScoreCard_Team.CtrlNo"
        End If
    Case "is_FIND"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Scoring_ScoreCard_Team.PK, tbl_Scoring_ScoreCard_Team.CtrlNo, " & _
            " tbl_Scoring_ScoreCard_Team.TeamKey, tbl_Scoring_Team.TeamID AS PlayerName, " & _
            " tbl_Scoring_ScoreCard_Team.DDate, tbl_Scoring_ScoreCard_Team.Score, " & _
            " tbl_Scoring_ScoreCard_Team.Front9Gross, tbl_Scoring_ScoreCard_Team.Back9Gross, tbl_Scoring_ScoreCard_Team.GrossPoints, " & _
            " tbl_Scoring_ScoreCard_Team.Front9Net, tbl_Scoring_ScoreCard_Team.Back9Net, tbl_Scoring_ScoreCard_Team.NetPoints, " & _
            " tbl_Scoring_ScoreCard_Team.LastModified, tbl_Scoring_ScoreCard_Team.Front9Score, " & _
            " tbl_Scoring_ScoreCard_Team.Back9Score " & _
            " FROM tbl_Scoring_ScoreCard_Team LEFT OUTER JOIN " & _
            " tbl_Scoring_Team ON tbl_Scoring_ScoreCard_Team.TeamKey = tbl_Scoring_Team.PK " & _
            " WHERE (tbl_Scoring_ScoreCard_Team.PK = " & strCtrl & ") " & _
            " ORDER BY tbl_Scoring_ScoreCard_Team.CtrlNo DESC"
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Scoring_ScoreCard_Team.PK, tbl_Scoring_ScoreCard_Team.CtrlNo, " & _
            " tbl_Scoring_ScoreCard_Team.TeamKey, tbl_Scoring_Team.TeamID AS PlayerName, " & _
            " tbl_Scoring_ScoreCard_Team.DDate, tbl_Scoring_ScoreCard_Team.Score, " & _
            " tbl_Scoring_ScoreCard_Team.Front9Gross, tbl_Scoring_ScoreCard_Team.Back9Gross, tbl_Scoring_ScoreCard_Team.GrossPoints, " & _
            " tbl_Scoring_ScoreCard_Team.Front9Net, tbl_Scoring_ScoreCard_Team.Back9Net, tbl_Scoring_ScoreCard_Team.NetPoints, " & _
            " tbl_Scoring_ScoreCard_Team.LastModified, tbl_Scoring_ScoreCard_Team.Front9Score, " & _
            " tbl_Scoring_ScoreCard_Team.Back9Score " & _
            " FROM tbl_Scoring_ScoreCard_Team LEFT OUTER JOIN " & _
            " tbl_Scoring_Team ON tbl_Scoring_ScoreCard_Team.TeamKey = tbl_Scoring_Team.PK " & _
            " WHERE (tbl_Scoring_ScoreCard_Team.TournamentKey = " & TournamentKey & ") " & _
            " ORDER BY tbl_Scoring_ScoreCard_Team.CtrlNo"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Scoring_ScoreCard_Team.PK, tbl_Scoring_ScoreCard_Team.CtrlNo, " & _
            " tbl_Scoring_ScoreCard_Team.TeamKey, tbl_Scoring_Team.TeamID AS PlayerName, " & _
            " tbl_Scoring_ScoreCard_Team.DDate, tbl_Scoring_ScoreCard_Team.Score, " & _
            " tbl_Scoring_ScoreCard_Team.Front9Gross, tbl_Scoring_ScoreCard_Team.Back9Gross, tbl_Scoring_ScoreCard_Team.GrossPoints, " & _
            " tbl_Scoring_ScoreCard_Team.Front9Net, tbl_Scoring_ScoreCard_Team.Back9Net, tbl_Scoring_ScoreCard_Team.NetPoints, " & _
            " tbl_Scoring_ScoreCard_Team.LastModified, tbl_Scoring_ScoreCard_Team.Front9Score, " & _
            " tbl_Scoring_ScoreCard_Team.Back9Score " & _
            " FROM tbl_Scoring_ScoreCard_Team LEFT OUTER JOIN " & _
            " tbl_Scoring_Team ON tbl_Scoring_ScoreCard_Team.TeamKey = tbl_Scoring_Team.PK " & _
            " WHERE (tbl_Scoring_ScoreCard_Team.TournamentKey = " & TournamentKey & ") " & _
            " AND (tbl_Scoring_ScoreCard_Team.CtrlNo < '" & strCtrl & "') " & _
            " ORDER BY tbl_Scoring_ScoreCard_Team.CtrlNo DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Scoring_ScoreCard_Team.PK, tbl_Scoring_ScoreCard_Team.CtrlNo, " & _
            " tbl_Scoring_ScoreCard_Team.TeamKey, tbl_Scoring_Team.TeamID AS PlayerName, " & _
            " tbl_Scoring_ScoreCard_Team.DDate, tbl_Scoring_ScoreCard_Team.Score, " & _
            " tbl_Scoring_ScoreCard_Team.Front9Gross, tbl_Scoring_ScoreCard_Team.Back9Gross, tbl_Scoring_ScoreCard_Team.GrossPoints, " & _
            " tbl_Scoring_ScoreCard_Team.Front9Net, tbl_Scoring_ScoreCard_Team.Back9Net, tbl_Scoring_ScoreCard_Team.NetPoints, " & _
            " tbl_Scoring_ScoreCard_Team.LastModified, tbl_Scoring_ScoreCard_Team.Front9Score, " & _
            " tbl_Scoring_ScoreCard_Team.Back9Score " & _
            " FROM tbl_Scoring_ScoreCard_Team LEFT OUTER JOIN " & _
            " tbl_Scoring_Team ON tbl_Scoring_ScoreCard_Team.TeamKey = tbl_Scoring_Team.PK " & _
            " WHERE (tbl_Scoring_ScoreCard_Team.TournamentKey = " & TournamentKey & ") " & _
            " AND (tbl_Scoring_ScoreCard_Team.CtrlNo > '" & strCtrl & "') " & _
            " ORDER BY tbl_Scoring_ScoreCard_Team.CtrlNo "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Scoring_ScoreCard_Team.PK, tbl_Scoring_ScoreCard_Team.CtrlNo, " & _
            " tbl_Scoring_ScoreCard_Team.TeamKey, tbl_Scoring_Team.TeamID AS PlayerName, " & _
            " tbl_Scoring_ScoreCard_Team.DDate, tbl_Scoring_ScoreCard_Team.Score, " & _
            " tbl_Scoring_ScoreCard_Team.Front9Gross, tbl_Scoring_ScoreCard_Team.Back9Gross, tbl_Scoring_ScoreCard_Team.GrossPoints, " & _
            " tbl_Scoring_ScoreCard_Team.Front9Net, tbl_Scoring_ScoreCard_Team.Back9Net, tbl_Scoring_ScoreCard_Team.NetPoints, " & _
            " tbl_Scoring_ScoreCard_Team.LastModified, tbl_Scoring_ScoreCard_Team.Front9Score, " & _
            " tbl_Scoring_ScoreCard_Team.Back9Score " & _
            " FROM tbl_Scoring_ScoreCard_Team LEFT OUTER JOIN " & _
            " tbl_Scoring_Team ON tbl_Scoring_ScoreCard_Team.TeamKey = tbl_Scoring_Team.PK " & _
            " WHERE (tbl_Scoring_ScoreCard_Team.TournamentKey = " & TournamentKey & ") " & _
            " ORDER BY tbl_Scoring_ScoreCard_Team.CtrlNo DESC"
    Case Else: Exit Function
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtCtrl.Text = rs!CtrlNo
    txtDate.Text = Format(rs!dDate, "mm/dd/yyyy")
    txtPlayer.Text = rs!PlayerName
    
    lstPlayer.ListItems.Clear
    dblHandicap = 0
    t = "SELECT tbl_Scoring_PlayerName.LastName, " & _
        " tbl_Scoring_PlayerName.FirstName, " & _
        " tbl_Scoring_PlayerName.HandiCap, " & _
        " tbl_Scoring_PlayerName.Class " & _
        " FROM tbl_Scoring_Team LEFT OUTER JOIN " & _
        " tbl_Scoring_Team_Detail ON tbl_Scoring_Team.PK = tbl_Scoring_Team_Detail.TeamKey LEFT OUTER JOIN " & _
        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
        " Where (tbl_Scoring_Team.PK = " & rs!TeamKey & ") " & _
        " ORDER BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    With lstPlayer.ListItems
    While Not rt.EOF
        dblHandicap = dblHandicap + CDbl(rt!HandiCap)
        Set x = .Add()
        x.Text = ""
        x.SubItems(1) = rt!LastName & ", " & rt!FirstName
        x.SubItems(2) = rt!HandiCap
        x.SubItems(3) = rt!Class
        rt.MoveNext
    Wend
    End With
    rt.Close
    
    txtTeamHandicap.Text = Mid(Format((CDbl(dblHandicap) / 2), "#,##0.00"), 1, Len(Format((CDbl(dblHandicap) / 2), "#,##0.00")) - 3)
    
    'txtHandicap.Text = rs!HandiCap
    'txtClass.Text = rs!Class
    
    'txtSGrossF.Text = rs!Front9Gross
    'txtSGrossB.Text = rs!Back9Gross
    
'    txtScoreGrossF.Text = rs!Front9Gross
'    txtScoreGrossB.Text = rs!Back9Gross
'    txtSGrossTot.Text = rs!GrossPoints
    'txtSNetF.Text = rs!Front9Net
    'txtSNetB.Text = rs!Back9Net
    'txtSNetTot.Text = rs!NetPoints

    txtGrossScoreF.Text = rs!Front9Score
    txtGrossScoreB.Text = rs!Back9Score
    txtGrossScoreTot.Text = rs!Score
    
    txtGrossPtsF.Text = rs!Front9Gross
    txtGrossPtsB.Text = rs!Back9Gross
    txtGrossPtsTot.Text = rs!GrossPoints
    txtNetPtsF.Text = rs!Front9Net
    txtNetPtsB.Text = rs!Back9Net
    txtNetPtsTot.Text = rs!NetPoints
    
    StatusBar1.Panels(1).Text = rs!PK
    StatusBar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", "Last Modified : " & rs!LastModified)
    
    i = -1
    t = "SELECT Par, Handicap, Score, Gross, Net " & _
        " From tbl_Scoring_ScoreCard_Team_Detail " & _
        " Where (ScoreCardKey = " & rs!PK & ") " & _
        " ORDER BY Hole"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        DoEvents
        i = i + 1
        txtGrossScore(i).Text = rt!Score
        txtGrossPts(i).Text = rt!Gross
        txtNetPts(i).Text = rt!Net
        rt.MoveNext
    Wend
    rt.Close
    
    SaveSetting App.EXEName, "ScoreCardControl", "ScoreCardCtrl", rs!CtrlNo
    
End If
rs.Close
End Function


Private Function PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If picSearchAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
If picPrint.Visible = True Then Exit Function
If AccessRights("Scoring Score Card", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
If CHECK_TOURNAMENT_STATUS(TournamentKey) <> 0 Then MsgBox "Tournament was already locked!               ", vbCritical, "Error...": Exit Function
picMain.Enabled = False
picToolbar.Enabled = False
picSearchAdd.ZOrder 0
txtSearchAdd.Text = ""
txtDateAdd.Text = ""
picSearchAdd.Visible = True
txtSearchAdd.SetFocus
End Function

Private Function PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If StatusBar1.Panels(1).Text = "" Then Exit Function
If picSearchAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
If picPrint.Visible = True Then Exit Function
If AccessRights("Scoring Score Card", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
If CHECK_TOURNAMENT_STATUS(TournamentKey) <> 0 Then MsgBox "Tournament was already locked!               ", vbCritical, "Error...": Exit Function
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_EDITTING
txtGrossScore(0).SetFocus
End Function

Private Function PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If StatusBar1.Panels(1).Text = "" Then Exit Function
If picSearchAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
If picPrint.Visible = True Then Exit Function
If AccessRights("Scoring Score Card", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
If CHECK_TOURNAMENT_STATUS(TournamentKey) <> 0 Then MsgBox "Tournament was already locked!               ", vbCritical, "Error...": Exit Function
If MsgBox("ARE YOU SURE IN DELETING THIS RECORD?                    ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Function
ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_Team WHERE (PK = " & StatusBar1.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "ScoreCardControl", "ScoreCardCtrl", ""), "is_PAGEDOWN"
If Trim(txtPlayer.Text) = "" Then BROWSER GetSetting(App.EXEName, "ScoreCardControl", "ScoreCardCtrl", ""), "is_HOME"
End Function

Private Function PRESS_F5()
If picSearchAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
If picPrint.Visible = True Then Exit Function
On Error GoTo PG:

If TRANSACTIONTYPE = is_ADDING Then
    Dim TourNoOfPlaysTmp, SCardKey, i, dblPar, dblHandicap, _
    dblScore, dblGross, dblNet, j, strCtrlNo
    s = "SELECT COUNT(*) AS NoofRec " & _
        " From tbl_Scoring_ScoreCard_Team " & _
        " WHERE (TournamentKey = " & TournamentKey & ") " & _
        " AND (TeamKey = " & TeamKey & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    TourNoOfPlaysTmp = rs!NoofRec
    rs.Close
    If CDbl(TourNoOfPlaysTmp) > CDbl(TourNoOfPlays) Then MsgBox "Number of Plays Exceeded!                  ", vbCritical, "Error...": Exit Function
    
    strCtrlNo = "00000001"
    s = "SELECT TOP 1 CtrlNo " & _
        " FROM tbl_Scoring_ScoreCard_Team " & _
        " WHERE (TournamentKey = " & TournamentKey & ") " & _
        " ORDER BY CtrlNo DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        strCtrlNo = Format(CDbl(rs!CtrlNo) + 1, "0000000#")
    End If
    rs.Close
    
    Do
        s = "SELECT tbl_Scoring_ScoreCard_Team.* " & _
            " FROM tbl_Scoring_ScoreCard_Team " & _
            " WHERE (TournamentKey = " & TournamentKey & ") " & _
            " AND (CtrlNo = '" & strCtrlNo & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If
        rs.Close
        strCtrlNo = Format(CDbl(strCtrlNo) + 1, "0000000#")
    Loop
    
    ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_Team " & _
                      " (TournamentKey, TeamKey, DDate, LastModified, CtrlNo) " & _
                      " VALUES (" & TournamentKey & ", " & TeamKey & ", " & _
                      " '" & FormatDateTime(txtDate.Text, vbShortDate) & "', " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "', '" & strCtrlNo & "')"
    
    SCardKey = 0
    s = "SELECT PK " & _
        " FROM tbl_Scoring_ScoreCard_Team " & _
        " WHERE (TournamentKey = " & TournamentKey & ") " & _
        " AND (TeamKey = " & TeamKey & ") " & _
        " AND (DDate = '" & FormatDateTime(txtDate.Text, vbShortDate) & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        SCardKey = rs!PK
    End If
    rs.Close
    
    If CDbl(SCardKey) <> 0 Then
        j = 0
        For i = 1 To 18
            With FGrid
                j = j + 1
                If j >= 1 And j <= 9 Then
                    dblPar = .TextMatrix(1, i + 1)
                    dblHandicap = .TextMatrix(2, i + 1)
                Else
                    dblPar = .TextMatrix(1, i + 2)
                    dblHandicap = .TextMatrix(2, i + 2)
                End If
            End With
            
            dblScore = RETURNTEXTVALUE(txtGrossScore(i - 1))
            dblGross = RETURNTEXTVALUE(txtGrossPts(i - 1))
            dblNet = RETURNTEXTVALUE(txtNetPts(i - 1))
            
            ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_Team_Detail " & _
                              " (ScoreCardKey, Hole, Par, Handicap, Score, Gross, Net) " & _
                              " VALUES (" & SCardKey & ", " & i & ", " & CDbl(dblPar) & ", " & _
                              " " & CDbl(dblHandicap) & ", " & CDbl(dblScore) & ", " & _
                              " " & CDbl(dblGross) & ", " & CDbl(dblNet) & ")"
                              
        Next i
    End If
    
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER strCtrlNo, "is_LOAD"
    
End If

If TRANSACTIONTYPE = is_EDITTING Then
    SCardKey = StatusBar1.Panels(1).Text
    ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_Team " & _
                      " SET LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & SCardKey & ")"
    
    ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_Team_Detail WHERE (ScoreCardKey = " & SCardKey & ")"
    j = 0
    For i = 1 To 18
        With FGrid
            j = j + 1
            If j >= 1 And j <= 9 Then
                dblPar = .TextMatrix(1, i + 1)
                dblHandicap = .TextMatrix(2, i + 1)
            Else
                dblPar = .TextMatrix(1, i + 2)
                dblHandicap = .TextMatrix(2, i + 2)
            End If
        End With
        
        dblScore = RETURNTEXTVALUE(txtGrossScore(i - 1))
        dblGross = RETURNTEXTVALUE(txtGrossPts(i - 1))
        dblNet = RETURNTEXTVALUE(txtNetPts(i - 1))
        
        ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_Team_Detail " & _
                          " (ScoreCardKey, Hole, Par, Handicap, Score, Gross, Net) " & _
                          " VALUES (" & SCardKey & ", " & i & ", " & CDbl(dblPar) & ", " & _
                          " " & CDbl(dblHandicap) & ", " & CDbl(dblScore) & ", " & _
                          " " & CDbl(dblGross) & ", " & CDbl(dblNet) & ")"
                          
    Next i
    
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER SCardKey, "is_FIND"
    
End If
Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Function
End Function

Private Function PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If picSearchAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
If picPrint.Visible = True Then Exit Function
picToolbar.Enabled = False
picMain.Enabled = False
txtSearch.Text = ""
picSearch.ZOrder 0
picSearch.Visible = True
txtSearch.SetFocus
End Function

Private Function PRESS_F9()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If picSearchAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
If picPrint.Visible = True Then Exit Function
cmdOKPrint.Top = 960
cmdCancelPrint.Top = 960
picElse.Visible = False
picScoreCard.Visible = False
picPrint.Height = 1695
picPrint.Width = 4335
picPrint.Top = (Me.ScaleHeight - picPrint.Height) / 4
picPrint.Left = (Me.ScaleWidth - picPrint.Width) / 2
cmbPrintType.ListIndex = -1
picPrint.ZOrder 0
picToolbar.Enabled = False
picMain.Enabled = False
picPrint.Visible = True
cmbPrintType.SetFocus
End Function

Private Function PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSearchAdd.Visible = True Then cmdCancelAdd_Click: Exit Function
    If picSearch.Visible = True Then cmdCancelSearch_Click: Exit Function
    If picPrint.Visible = True Then cmdCancelPrint_Click: Exit Function
    Unload Me
Else
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "ScoreCardControl", "ScoreCardCtrl", ""), "is_LOAD"
    If Trim(txtPlayer.Text) = "" Then BROWSER GetSetting(App.EXEName, "ScoreCardControl", "ScoreCardCtrl", ""), "is_HOME"
End If
End Function

Private Function CLEARTEXT()
Dim i, x
For i = 0 To 17
    txtGrossScore(i).Text = ""
    txtGrossPts(i).Text = "0"
    txtNetPts(i).Text = "0"
Next i
'txtTournament.Text = ""
'txtTourDate.Text = ""
txtCtrl.Text = ""
txtDate.Text = ""
txtPlayer.Text = ""
'txtHandicap.Text = ""
'txtClass.Text = ""
txtDay.Text = ""
txtDate.Text = ""

'txtScoreGrossF.Text = ""
'txtScoreGrossB.Text = ""
'txtSGrossTot.Text = ""
'txtSNetF.Text = ""
'txtSNetB.Text = ""
'txtSNetTot.Text = ""

txtGrossScoreF.Text = "0"
txtGrossScoreB.Text = "0"

StatusBar1.Panels(1).Text = ""
StatusBar1.Panels(2).Text = ""
lstPlayer.ListItems.Clear

End Function

Private Function LOCKTEXT(bln As Boolean)
Dim i
If bln Then
    For i = 0 To 17
        txtGrossScore(i).Locked = True
    Next i
Else
    For i = 0 To 17
        txtGrossScore(i).Locked = False
    Next i
End If
End Function

Public Function TOOLBARFUNC(intSel As Integer)
With Toolbar1
    Select Case intSel
        Case 1      'REFRESH
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(1).Enabled = True
            .Buttons(3).Enabled = True
            .Buttons(5).Enabled = True
            .Buttons(7).Image = 4
            .Buttons(7).Caption = "First"
            .Buttons(9).Image = 5
            .Buttons(9).Caption = "Back"
            .Buttons(7).Enabled = True
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = True
            .Buttons(13).Enabled = True
            .Buttons(15).Enabled = True
            .Buttons(17).Enabled = True
            .Buttons(19).Enabled = True
            .Buttons(1).ToolTipText = "NEW (Ins)"
            .Buttons(3).ToolTipText = "EDIT (F2)"
            .Buttons(5).ToolTipText = "DELETE (Del)"
            .Buttons(7).ToolTipText = "FIRST (Home)"
            .Buttons(9).ToolTipText = "BACK (PgUp)"
            .Buttons(11).ToolTipText = "NEXT (PgDown)"
            .Buttons(13).ToolTipText = "LAST (End)"
            .Buttons(15).ToolTipText = "FIND (F6)"
            .Buttons(17).ToolTipText = "PRINT (F9)"
            .Buttons(19).ToolTipText = "CLOSE (Esc)"
        Case 2      'ADD/EDIT
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 10
            .Buttons(1).Enabled = False
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = False
            .Buttons(7).Image = 11
            .Buttons(7).Caption = "Save"
            .Buttons(9).Image = 12
            .Buttons(9).Caption = "Undo"
            .Buttons(7).Enabled = True
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
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
        Case 3      'FIND
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(1).Enabled = False
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = False
            .Buttons(7).Image = 4
            .Buttons(7).Caption = "First"
            .Buttons(9).Image = 12
            .Buttons(9).Caption = "Undo"
            .Buttons(7).Enabled = False
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
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
    End Select
End With
End Function

'Private Sub LOAD_CARD()
'Dim GRow As Long
'Dim HEADER1$, i, j, Tot1, Tot2, a, Tot
't = "SELECT TOP 1 tbl_Scoring_Yardage_Par_HandicapIndex_Master.* " & _
'    " FROM tbl_Scoring_Yardage_Par_HandicapIndex_Master " & _
'    " WHERE (EffectDate <= '" & FormatDateTime(Date, vbShortDate) & "') " & _
'    " ORDER BY EffectDate DESC"
'If rt.State = adStateOpen Then rt.Close
'rt.Open t, ConnOmega
'If rt.RecordCount > 0 Then
'    GRow = 0
'    s = "SELECT Hole, Par, HandicapIndex, " & _
'        " Gold, Blue, White, Red " & _
'        " From tbl_Scoring_Yardage_Par_HandicapIndex " & _
'        " WHERE (MasterKey = " & rt!PK & ") " & _
'        " ORDER BY Hole"
'    If rs.State = adStateOpen Then rs.Close
'    rs.Open s, ConnOmega
'    If rs.RecordCount > 0 Then
'        With FGrid
'            .Clear
'            i = 0: j = 0
'            HEADER1$ = HEADER1$ & "|" & "HOLE"
'            rs.MoveFirst
'            While Not rs.EOF
'                If i = 9 Then
'                    i = 0
'                    HEADER1$ = HEADER1$ & "|" & "OUT"
'                    HEADER1$ = HEADER1$ & "|" & CStr(rs!Hole)
'                Else
'                    HEADER1$ = HEADER1$ & "|" & CStr(rs!Hole)
'                End If
'                i = i + 1
'                rs.MoveNext
'            Wend
'            HEADER1$ = HEADER1$ & "|" & "IN" & "|" & "TOT" ' & "|" & "TOT"
'            .FormatString = HEADER1$
'            For i = 1 To .Cols - 1
'                If i = 1 Then
'                    .ColWidth(i) = 2000
'                    .ColAlignment(i) = 1
'                ElseIf i = 11 Or _
'                i = 21 Or i = 22 Then
'                    .ColWidth(i) = 550
'                    .ColAlignment(i) = flexAlignRightCenter
'                Else
'                    .ColWidth(i) = 450
'                    .ColAlignment(i) = flexAlignRightCenter
'                End If
'            Next i
'        End With
'    End If
'    rs.Close
'
'    'Par
'    i = 1
'    GRow = GRow + 1
'    a = 0: Tot1 = 0: Tot2 = 0: Tot = 0
'    s = "SELECT Hole, Par, HandicapIndex, " & _
'        " Gold, Blue, White, Red " & _
'        " From tbl_Scoring_Yardage_Par_HandicapIndex " & _
'        " WHERE (MasterKey = " & rt!PK & ") " & _
'        " ORDER BY Hole"
'    If rs.State = adStateOpen Then rs.Close
'    rs.Open s, ConnOmega
'    If rs.RecordCount > 0 Then
'        With FGrid
'            rs.MoveFirst
'            .TextMatrix(GRow, i) = "PAR"
'            While Not rs.EOF
'                i = i + 1
'                If a = 9 Then
'                    .TextMatrix(GRow, i) = Tot
'                    i = i + 1
'                    .TextMatrix(GRow, i) = rs!Par
'                    Tot1 = Tot
'                    a = 0
'                    Tot = 0
'                Else
'                    .TextMatrix(GRow, i) = rs!Par
'                End If
'                Tot = Tot + CDbl(rs!Par)
'                a = a + 1
'                rs.MoveNext
'            Wend
'            Tot2 = CDbl(Tot) + CDbl(Tot1)
'            i = i + 1
'            .TextMatrix(GRow, i) = Tot
'            i = i + 1
'            .TextMatrix(GRow, i) = Tot2 'Tot1
'            'i = i + 1
'            '.TextMatrix(GRow, i) = Tot2
'        End With
'    End If
'    rs.Close
'
'    'Handicap
'    i = 1
'    GRow = GRow + 1
'    a = 0: Tot1 = 0: Tot2 = 0: Tot = 0
'    s = "SELECT Hole, Par, HandicapIndex, " & _
'        " Gold, Blue, White, Red " & _
'        " From tbl_Scoring_Yardage_Par_HandicapIndex " & _
'        " WHERE (MasterKey = " & rt!PK & ") " & _
'        " ORDER BY Hole"
'    If rs.State = adStateOpen Then rs.Close
'    rs.Open s, ConnOmega
'    If rs.RecordCount > 0 Then
'        With FGrid
'            rs.MoveFirst
'            .Rows = .Rows + 1
'            .TextMatrix(GRow, i) = "HANDICAP"
'            While Not rs.EOF
'                i = i + 1
'                If a = 9 Then
'                    .TextMatrix(GRow, i) = ""
'                    i = i + 1
'                    .TextMatrix(GRow, i) = rs!HandicapIndex
'                    Tot1 = Tot
'                    a = 0
'                    Tot = 0
'                Else
'                    .TextMatrix(GRow, i) = rs!HandicapIndex
'                End If
'                Tot = Tot + CDbl(rs!Par)
'                a = a + 1
'                rs.MoveNext
'            Wend
'            Tot2 = CDbl(Tot) + CDbl(Tot1)
'            i = i + 1
'            .TextMatrix(GRow, i) = ""
'            i = i + 1
'            .TextMatrix(GRow, i) = ""
'            'i = i + 1
'            '.TextMatrix(GRow, i) = ""
'        End With
'    End If
'    rs.Close
'
'    'Gold
'    i = 1
'    GRow = GRow + 1
'    a = 0: Tot1 = 0: Tot2 = 0: Tot = 0
'    s = "SELECT Hole, Par, HandicapIndex, " & _
'        " Gold, Blue, White, Red " & _
'        " From tbl_Scoring_Yardage_Par_HandicapIndex " & _
'        " WHERE (MasterKey = " & rt!PK & ") " & _
'        " ORDER BY Hole"
'    If rs.State = adStateOpen Then rs.Close
'    rs.Open s, ConnOmega
'    If rs.RecordCount > 0 Then
'        With FGrid
'            rs.MoveFirst
'            .Rows = .Rows + 1
'            .TextMatrix(GRow, i) = "GOLD"
'            While Not rs.EOF
'                i = i + 1
'                If a = 9 Then
'                    .TextMatrix(GRow, i) = Tot
'                    i = i + 1
'                    .TextMatrix(GRow, i) = rs!Gold
'                    Tot1 = Tot
'                    a = 0
'                    Tot = 0
'                Else
'                    .TextMatrix(GRow, i) = rs!Gold
'                End If
'                Tot = Tot + CDbl(rs!Gold)
'                a = a + 1
'                rs.MoveNext
'            Wend
'            Tot2 = CDbl(Tot) + CDbl(Tot1)
'            i = i + 1
'            .TextMatrix(GRow, i) = Tot
'            i = i + 1
'            .TextMatrix(GRow, i) = Tot2 'Tot1
'            'i = i + 1
'            '.TextMatrix(GRow, i) = Tot2
'        End With
'    End If
'    rs.Close
'
'    'Blue
'    i = 1
'    GRow = GRow + 1
'    a = 0: Tot1 = 0: Tot2 = 0: Tot = 0
'    s = "SELECT Hole, Par, HandicapIndex, " & _
'        " Gold, Blue, White, Red " & _
'        " From tbl_Scoring_Yardage_Par_HandicapIndex " & _
'        " WHERE (MasterKey = " & rt!PK & ") " & _
'        " ORDER BY Hole"
'    If rs.State = adStateOpen Then rs.Close
'    rs.Open s, ConnOmega
'    If rs.RecordCount > 0 Then
'        With FGrid
'            rs.MoveFirst
'            .Rows = .Rows + 1
'            .TextMatrix(GRow, i) = "BLUE"
'            While Not rs.EOF
'                i = i + 1
'                If a = 9 Then
'                    .TextMatrix(GRow, i) = Tot
'                    i = i + 1
'                    .TextMatrix(GRow, i) = rs!Blue
'                    Tot1 = Tot
'                    a = 0
'                    Tot = 0
'                Else
'                    .TextMatrix(GRow, i) = rs!Blue
'                End If
'                Tot = Tot + CDbl(rs!Blue)
'                a = a + 1
'                rs.MoveNext
'            Wend
'            Tot2 = CDbl(Tot) + CDbl(Tot1)
'            i = i + 1
'            .TextMatrix(GRow, i) = Tot
'            i = i + 1
'            .TextMatrix(GRow, i) = Tot2 'Tot1
'            'i = i + 1
'            '.TextMatrix(GRow, i) = Tot2
'        End With
'    End If
'    rs.Close
'
'    'White
'    i = 1
'    GRow = GRow + 1
'    a = 0: Tot1 = 0: Tot2 = 0: Tot = 0
'    s = "SELECT Hole, Par, HandicapIndex, " & _
'        " Gold, Blue, White, Red " & _
'        " From tbl_Scoring_Yardage_Par_HandicapIndex " & _
'        " WHERE (MasterKey = " & rt!PK & ") " & _
'        " ORDER BY Hole"
'    If rs.State = adStateOpen Then rs.Close
'    rs.Open s, ConnOmega
'    If rs.RecordCount > 0 Then
'        With FGrid
'            rs.MoveFirst
'            .Rows = .Rows + 1
'            .TextMatrix(GRow, i) = "WHITE"
'            While Not rs.EOF
'                i = i + 1
'                If a = 9 Then
'                    .TextMatrix(GRow, i) = Tot
'                    i = i + 1
'                    .TextMatrix(GRow, i) = rs!White
'                    Tot1 = Tot
'                    a = 0
'                    Tot = 0
'                Else
'                    .TextMatrix(GRow, i) = rs!White
'                End If
'                Tot = Tot + CDbl(rs!White)
'                a = a + 1
'                rs.MoveNext
'            Wend
'            Tot2 = CDbl(Tot) + CDbl(Tot1)
'            i = i + 1
'            .TextMatrix(GRow, i) = Tot
'            i = i + 1
'            .TextMatrix(GRow, i) = Tot2 'Tot1
'            'i = i + 1
'            '.TextMatrix(GRow, i) = Tot2
'        End With
'    End If
'    rs.Close
'
'    'Red
'    i = 1
'    GRow = GRow + 1
'    a = 0: Tot1 = 0: Tot2 = 0: Tot = 0
'    s = "SELECT Hole, Par, HandicapIndex, " & _
'        " Gold, Blue, White, Red " & _
'        " From tbl_Scoring_Yardage_Par_HandicapIndex " & _
'        " WHERE (MasterKey = " & rt!PK & ") " & _
'        " ORDER BY Hole"
'    If rs.State = adStateOpen Then rs.Close
'    rs.Open s, ConnOmega
'    If rs.RecordCount > 0 Then
'        With FGrid
'            rs.MoveFirst
'            .Rows = .Rows + 1
'            .TextMatrix(GRow, i) = "RED"
'            While Not rs.EOF
'                i = i + 1
'                If a = 9 Then
'                    .TextMatrix(GRow, i) = Tot
'                    i = i + 1
'                    .TextMatrix(GRow, i) = rs!Red
'                    Tot1 = Tot
'                    a = 0
'                    Tot = 0
'                Else
'                    .TextMatrix(GRow, i) = rs!Red
'                End If
'                Tot = Tot + CDbl(rs!Red)
'                a = a + 1
'                rs.MoveNext
'            Wend
'            Tot2 = CDbl(Tot) + CDbl(Tot1)
'            i = i + 1
'            .TextMatrix(GRow, i) = Tot
'            i = i + 1
'            .TextMatrix(GRow, i) = Tot2 'Tot1
'            'i = i + 1
'            '.TextMatrix(GRow, i) = Tot2
'        End With
'    End If
'    rs.Close
'End If
'rt.Close
'End Sub

'Private Function LOAD_CARD()
'Dim GRow As Long
'Dim HEADER1$, i, j, Tot1, Tot2, a, Tot
'GRow = 0
's = "SELECT Hole, Par, HandicapIndex, " & _
'    " Gold, Blue, White, Red " & _
'    " From tbl_Scoring_Yardage_Par_HandicapIndex " & _
'    " ORDER BY Hole"
'If rs.State = adStateOpen Then rs.Close
'rs.Open s, ConnOmega
'If rs.RecordCount > 0 Then
'    With FGrid
'        .Clear
'        i = 0: j = 0
'        HEADER1$ = HEADER1$ & "|" & "HOLE"
'        rs.MoveFirst
'        While Not rs.EOF
'            If i = 9 Then
'                i = 0
'                HEADER1$ = HEADER1$ & "|" & "OUT"
'                HEADER1$ = HEADER1$ & "|" & CStr(rs!Hole)
'            Else
'                HEADER1$ = HEADER1$ & "|" & CStr(rs!Hole)
'            End If
'            i = i + 1
'            rs.MoveNext
'        Wend
'        HEADER1$ = HEADER1$ & "|" & "IN" & "|" & "TOT" ' & "|" & "TOT"
'        .FormatString = HEADER1$
'        For i = 1 To .Cols - 1
'            If i = 1 Then
'                .ColWidth(i) = 2000
'                .ColAlignment(i) = 1
'            ElseIf i = 11 Or _
'            i = 21 Or i = 22 Then
'                .ColWidth(i) = 550
'                .ColAlignment(i) = flexAlignRightCenter
'            Else
'                .ColWidth(i) = 450
'                .ColAlignment(i) = flexAlignRightCenter
'            End If
'        Next i
'    End With
'End If
'rs.Close
'
''Par
'i = 1
'GRow = GRow + 1
'a = 0: Tot1 = 0: Tot2 = 0: Tot = 0
's = "SELECT Hole, Par, HandicapIndex, " & _
'    " Gold, Blue, White, Red " & _
'    " From tbl_Scoring_Yardage_Par_HandicapIndex " & _
'    " ORDER BY Hole"
'If rs.State = adStateOpen Then rs.Close
'rs.Open s, ConnOmega
'If rs.RecordCount > 0 Then
'    With FGrid
'        rs.MoveFirst
'        .TextMatrix(GRow, i) = "PAR"
'        While Not rs.EOF
'            i = i + 1
'            If a = 9 Then
'                .TextMatrix(GRow, i) = Tot
'                i = i + 1
'                .TextMatrix(GRow, i) = rs!Par
'                Tot1 = Tot
'                a = 0
'                Tot = 0
'            Else
'                .TextMatrix(GRow, i) = rs!Par
'            End If
'            Tot = Tot + CDbl(rs!Par)
'            a = a + 1
'            rs.MoveNext
'        Wend
'        Tot2 = CDbl(Tot) + CDbl(Tot1)
'        i = i + 1
'        .TextMatrix(GRow, i) = Tot
'        i = i + 1
'        .TextMatrix(GRow, i) = Tot2 'Tot1
'        'i = i + 1
'        '.TextMatrix(GRow, i) = Tot2
'    End With
'End If
'rs.Close
'
''Handicap
'i = 1
'GRow = GRow + 1
'a = 0: Tot1 = 0: Tot2 = 0: Tot = 0
's = "SELECT Hole, Par, HandicapIndex, " & _
'    " Gold, Blue, White, Red " & _
'    " From tbl_Scoring_Yardage_Par_HandicapIndex " & _
'    " ORDER BY Hole"
'If rs.State = adStateOpen Then rs.Close
'rs.Open s, ConnOmega
'If rs.RecordCount > 0 Then
'    With FGrid
'        rs.MoveFirst
'        .Rows = .Rows + 1
'        .TextMatrix(GRow, i) = "HANDICAP"
'        While Not rs.EOF
'            i = i + 1
'            If a = 9 Then
'                .TextMatrix(GRow, i) = ""
'                i = i + 1
'                .TextMatrix(GRow, i) = rs!HandicapIndex
'                Tot1 = Tot
'                a = 0
'                Tot = 0
'            Else
'                .TextMatrix(GRow, i) = rs!HandicapIndex
'            End If
'            Tot = Tot + CDbl(rs!Par)
'            a = a + 1
'            rs.MoveNext
'        Wend
'        Tot2 = CDbl(Tot) + CDbl(Tot1)
'        i = i + 1
'        .TextMatrix(GRow, i) = ""
'        i = i + 1
'        .TextMatrix(GRow, i) = ""
'        'i = i + 1
'        '.TextMatrix(GRow, i) = ""
'    End With
'End If
'rs.Close
'
''Gold
'i = 1
'GRow = GRow + 1
'a = 0: Tot1 = 0: Tot2 = 0: Tot = 0
's = "SELECT Hole, Par, HandicapIndex, " & _
'    " Gold, Blue, White, Red " & _
'    " From tbl_Scoring_Yardage_Par_HandicapIndex " & _
'    " ORDER BY Hole"
'If rs.State = adStateOpen Then rs.Close
'rs.Open s, ConnOmega
'If rs.RecordCount > 0 Then
'    With FGrid
'        rs.MoveFirst
'        .Rows = .Rows + 1
'        .TextMatrix(GRow, i) = "GOLD"
'        While Not rs.EOF
'            i = i + 1
'            If a = 9 Then
'                .TextMatrix(GRow, i) = Tot
'                i = i + 1
'                .TextMatrix(GRow, i) = rs!Gold
'                Tot1 = Tot
'                a = 0
'                Tot = 0
'            Else
'                .TextMatrix(GRow, i) = rs!Gold
'            End If
'            Tot = Tot + CDbl(rs!Gold)
'            a = a + 1
'            rs.MoveNext
'        Wend
'        Tot2 = CDbl(Tot) + CDbl(Tot1)
'        i = i + 1
'        .TextMatrix(GRow, i) = Tot
'        i = i + 1
'        .TextMatrix(GRow, i) = Tot2 'Tot1
'        'i = i + 1
'        '.TextMatrix(GRow, i) = Tot2
'    End With
'End If
'rs.Close
'
''Blue
'i = 1
'GRow = GRow + 1
'a = 0: Tot1 = 0: Tot2 = 0: Tot = 0
's = "SELECT Hole, Par, HandicapIndex, " & _
'    " Gold, Blue, White, Red " & _
'    " From tbl_Scoring_Yardage_Par_HandicapIndex " & _
'    " ORDER BY Hole"
'If rs.State = adStateOpen Then rs.Close
'rs.Open s, ConnOmega
'If rs.RecordCount > 0 Then
'    With FGrid
'        rs.MoveFirst
'        .Rows = .Rows + 1
'        .TextMatrix(GRow, i) = "BLUE"
'        While Not rs.EOF
'            i = i + 1
'            If a = 9 Then
'                .TextMatrix(GRow, i) = Tot
'                i = i + 1
'                .TextMatrix(GRow, i) = rs!Blue
'                Tot1 = Tot
'                a = 0
'                Tot = 0
'            Else
'                .TextMatrix(GRow, i) = rs!Blue
'            End If
'            Tot = Tot + CDbl(rs!Blue)
'            a = a + 1
'            rs.MoveNext
'        Wend
'        Tot2 = CDbl(Tot) + CDbl(Tot1)
'        i = i + 1
'        .TextMatrix(GRow, i) = Tot
'        i = i + 1
'        .TextMatrix(GRow, i) = Tot2 'Tot1
'        'i = i + 1
'        '.TextMatrix(GRow, i) = Tot2
'    End With
'End If
'rs.Close
'
''White
'i = 1
'GRow = GRow + 1
'a = 0: Tot1 = 0: Tot2 = 0: Tot = 0
's = "SELECT Hole, Par, HandicapIndex, " & _
'    " Gold, Blue, White, Red " & _
'    " From tbl_Scoring_Yardage_Par_HandicapIndex " & _
'    " ORDER BY Hole"
'If rs.State = adStateOpen Then rs.Close
'rs.Open s, ConnOmega
'If rs.RecordCount > 0 Then
'    With FGrid
'        rs.MoveFirst
'        .Rows = .Rows + 1
'        .TextMatrix(GRow, i) = "WHITE"
'        While Not rs.EOF
'            i = i + 1
'            If a = 9 Then
'                .TextMatrix(GRow, i) = Tot
'                i = i + 1
'                .TextMatrix(GRow, i) = rs!White
'                Tot1 = Tot
'                a = 0
'                Tot = 0
'            Else
'                .TextMatrix(GRow, i) = rs!White
'            End If
'            Tot = Tot + CDbl(rs!White)
'            a = a + 1
'            rs.MoveNext
'        Wend
'        Tot2 = CDbl(Tot) + CDbl(Tot1)
'        i = i + 1
'        .TextMatrix(GRow, i) = Tot
'        i = i + 1
'        .TextMatrix(GRow, i) = Tot2 'Tot1
'        'i = i + 1
'        '.TextMatrix(GRow, i) = Tot2
'    End With
'End If
'rs.Close
'
''Red
'i = 1
'GRow = GRow + 1
'a = 0: Tot1 = 0: Tot2 = 0: Tot = 0
's = "SELECT Hole, Par, HandicapIndex, " & _
'    " Gold, Blue, White, Red " & _
'    " From tbl_Scoring_Yardage_Par_HandicapIndex " & _
'    " ORDER BY Hole"
'If rs.State = adStateOpen Then rs.Close
'rs.Open s, ConnOmega
'If rs.RecordCount > 0 Then
'    With FGrid
'        rs.MoveFirst
'        .Rows = .Rows + 1
'        .TextMatrix(GRow, i) = "RED"
'        While Not rs.EOF
'            i = i + 1
'            If a = 9 Then
'                .TextMatrix(GRow, i) = Tot
'                i = i + 1
'                .TextMatrix(GRow, i) = rs!Red
'                Tot1 = Tot
'                a = 0
'                Tot = 0
'            Else
'                .TextMatrix(GRow, i) = rs!Red
'            End If
'            Tot = Tot + CDbl(rs!Red)
'            a = a + 1
'            rs.MoveNext
'        Wend
'        Tot2 = CDbl(Tot) + CDbl(Tot1)
'        i = i + 1
'        .TextMatrix(GRow, i) = Tot
'        i = i + 1
'        .TextMatrix(GRow, i) = Tot2 'Tot1
'        'i = i + 1
'        '.TextMatrix(GRow, i) = Tot2
'    End With
'End If
'rs.Close
'
'End Function

Private Sub b8TitleBar1_CLoseClick()
cmdCancelAdd_Click
End Sub

Private Sub b8TitleBar2_CLoseClick()
cmdCancelSearch_Click
End Sub

Private Sub b8TitleBar3_CLoseClick()
cmdCancelPrint_Click
End Sub

Private Sub cmbDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKSearch_Click
End Sub

Private Sub cmbPrintType_Click()
If cmbPrintType.ListIndex = -1 Then Exit Sub
Select Case cmbPrintType.ListIndex
    Case 0  '== Result
        '1695
        cmdOKPrint.Top = 960
        cmdCancelPrint.Top = 960
        picPrint.Height = 1695
        picPrint.Width = 4335
        picPrint.Top = (Me.ScaleHeight - picPrint.Height) / 4
        picPrint.Left = (Me.ScaleWidth - picPrint.Width) / 2
    Case 1  '== Score Card
        picElse.Visible = False
        txtSearchScore.Text = ""
        picScoreCard.Visible = True
        '4440
        cmdOKPrint.Top = 4440
        cmdCancelPrint.Top = 4440
        picPrint.Height = 5055
        picPrint.Width = 4335
        picPrint.Top = (Me.ScaleHeight - picPrint.Height) / 4
        picPrint.Left = (Me.ScaleWidth - picPrint.Width) / 2
        txtSearchScore.SetFocus
    Case Else
        picScoreCard.Visible = False
        picElse.ZOrder 0
        cmbSortAll.Clear
        cmbSortAll.AddItem "SORT BY TEAM HANDICAP"
        cmbSortAll.AddItem "SORT BY GROSS SCORE"
        cmbSortAll.AddItem "SORT BY NET SCORE"
        picPrint.Height = 2055 '2415 '1935
        picPrint.Width = 4335
        cmdOKPrint.Top = 1440 '1680
        cmdCancelPrint.Top = 1440 '1680
        picPrint.Top = (Me.ScaleHeight - picPrint.Height) / 4
        picPrint.Left = (Me.ScaleWidth - picPrint.Width) / 2
        cmbSortAll.ListIndex = -1
        picElse.Visible = True
End Select
End Sub

Private Sub cmbSortAll_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKPrint.SetFocus
End Sub

Private Sub cmdCancelAdd_Click()
picToolbar.Enabled = True
picMain.Enabled = True
picSearchAdd.Visible = False
End Sub

Private Sub cmdCancelPrint_Click()
picPrint.Visible = False
picToolbar.Enabled = True
picMain.Enabled = True
End Sub

Private Sub cmdCancelSearch_Click()
picToolbar.Enabled = True
picMain.Enabled = True
picSearch.Visible = False
End Sub

Private Sub cmdOKAdd_Click()
Dim Array1, dblHandicap, x
If lstTeamAdd.ListIndex = -1 Then Exit Sub
If IsDate(txtDateAdd.Text) = False Then Exit Sub
Array1 = Split(Trim(txtTourDate.Text), " - ", -1, 1)
txtDateAdd.Text = Format(FormatDateTime(txtDateAdd.Text, vbShortDate), "mm/dd/yyyy")
If DateValue(FormatDateTime(txtDateAdd.Text, vbShortDate)) < DateValue(FormatDateTime(Array1(0), vbShortDate)) Then MsgBox "Date Out of Range From the Tournament Date!                     ", vbCritical, "Error...": txtDateAdd.SetFocus: HTEXT txtDateAdd: Exit Sub
If DateValue(FormatDateTime(txtDateAdd.Text, vbShortDate)) > DateValue(FormatDateTime(Array1(1), vbShortDate)) Then MsgBox "Date Out of Range From the Tournament Date!                     ", vbCritical, "Error...": txtDateAdd.SetFocus: HTEXT txtDateAdd: Exit Sub
's = "SELECT tbl_Scoring_ScoreCard.* " & _
'    " FROM tbl_Scoring_ScoreCard " & _
'    " WHERE (TournamentKey = " & TournamentKey & ") " & _
'    " AND (PlayerKey = " & lstResultAdd.ItemData(lstResultAdd.ListIndex) & ") " & _
'    " AND (DDate = '" & FormatDateTime(txtDateAdd.Text, vbShortDate) & "')"
'If rs.State = adStateOpen Then rs.Close
'rs.Open s, ConnOmega
'If rs.RecordCount > 0 Then
'    MsgBox "Found Duplicate Entry!                  ", vbCritical, "Error..."
'    txtDateAdd.SetFocus
'    HTEXT txtDateAdd
'    Exit Sub
'End If
'rs.Close

CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
TeamKey = lstTeamAdd.ItemData(lstTeamAdd.ListIndex)
Array1 = Split(lstTeamAdd.List(lstTeamAdd.ListIndex), " - ", -1, 1)
txtPlayer.Text = Array1(0)
'txtPlayer.Text = lstResultAdd.List(lstResultAdd.ListIndex)
txtDate.Text = Format(FormatDateTime(txtDateAdd.Text, vbShortDate), "mm/dd/yyyy")

dblHandicap = 0

s = "SELECT tbl_Scoring_PlayerName.LastName, " & _
    " tbl_Scoring_PlayerName.FirstName, " & _
    " tbl_Scoring_PlayerName.HandiCap, " & _
    " tbl_Scoring_PlayerName.Class " & _
    " FROM tbl_Scoring_Team LEFT OUTER JOIN " & _
    " tbl_Scoring_Team_Detail ON tbl_Scoring_Team.PK = tbl_Scoring_Team_Detail.TeamKey LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " Where (tbl_Scoring_Team.PK = " & lstTeamAdd.ItemData(lstTeamAdd.ListIndex) & ") " & _
    " ORDER BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
With lstPlayer.ListItems
While Not rs.EOF
    dblHandicap = dblHandicap + CDbl(rs!HandiCap)
    Set x = .Add()
    x.Text = ""
    x.SubItems(1) = rs!LastName & ", " & rs!FirstName
    x.SubItems(2) = rs!HandiCap
    x.SubItems(3) = rs!Class
    rs.MoveNext
Wend
End With
rs.Close

txtTeamHandicap.Text = Mid(Format((CDbl(dblHandicap) / 2), "#,##0.00"), 1, Len(Format((CDbl(dblHandicap) / 2), "#,##0.00")) - 3)

's = "SELECT HandiCap, Class " & _
'    " From tbl_Scoring_PlayerName " & _
'    " WHERE (PK = " & lstResultAdd.ItemData(lstResultAdd.ListIndex) & ")"
'If rs.State = adStateOpen Then rs.Close
'rs.Open s, ConnOmega
'If rs.RecordCount > 0 Then
'    txtHandicap.Text = rs!HandiCap
'    txtClass.Text = rs!Class
'End If
'rs.Close
cmdCancelAdd_Click
txtGrossScore(0).SetFocus

End Sub

Private Sub cmdOKPrint_Click()
Dim i, j, strClass, TableName, DetailTableName, _
Columns, ColumnsDet, Clustered, sMasterFields, _
sDetailFields, Arr, Arr1, MasterKey, dblGrossScore, _
dblNetScore, Filename, cnt, HeaderRow, dblFront9, _
dblBack9

Dim sUserName As String

Dim WorkbookName    As String
Dim ColTop, RowTop, ColCount, RowCount, strRange, strRange1, _
ColCountDet, RowCountDet, RowFrom, RowTo

If cmbPrintType.ListIndex = -1 Then Exit Sub

Select Case cmbPrintType.ListIndex
    Case 0  ' result
    
    Case 1  ' score card
    
    Case Else
        If cmbSortAll.ListIndex = -1 Then Exit Sub
        cnt = 0
        With CommonDialog1
            .CancelError = True
            On Error GoTo ErrorHandler
            .DialogTitle = "Save"
            '.Filter = "Excel(*.xls)|*.xls"
            .Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
            .ShowSave
            Filename = Trim(.Filename)
        End With
        
        On Error GoTo PG:
        
        WorkbookName = Filename
        picProgressBar.BackColor = &HFFFFFF
        picProgress.ZOrder 0
        picPrint.Visible = False
        picProgress.Visible = True
        
        sUserName = Replace(CStr(gbl_UserName), " ", "")
        
        Screen.MousePointer = vbHourglass
        TableName = "tmp_" & gbl_UserName & "_Report"
        Columns = ""
        Columns = Columns & "|Sorting:int:NOT NULL:DEFAULT(0)"
        Columns = Columns & "|TeamKey:int:NOT NULL"
        Columns = Columns & "|TeamID:varchar:(50):NOT NULL:DEFAULT('')"
        Columns = Columns & "|TotalHDCP:float:NOT NULL:DEFAULT(0)"
        Columns = Columns & "|TeamHDCP:float:NOT NULL:DEFAULT(0)"
        Columns = Columns & "|TeamClass:varchar:(5):NOT NULL:DEFAULT('')"
        Columns = Columns & "|Front9:float:NOT NULL:DEFAULT(0)"
        Columns = Columns & "|Back9:float:NOT NULL:DEFAULT(0)"
        Columns = Columns & "|GrossScore:float:NOT NULL:DEFAULT(0)"
        Columns = Columns & "|NetScore:float:NOT NULL:DEFAULT(0)"
        
        Clustered = ""
        Clustered = Clustered & "|Sorting"
        
        ColumnsDet = ""
        ColumnsDet = ColumnsDet & "|PlayerName:varchar:(100):NOT NULL:DEFAULT('')"
        ColumnsDet = ColumnsDet & "|HDCP:float:NOT NULL:DEFAULT(0)"
        
        DetailTableName = TableName & "_Detail"
        CreateTable gbl_Database, TableName, Columns, CStr(Clustered), 1, CStr(DetailTableName), CStr(ColumnsDet)
        
        sMasterFields = ""
        Arr = Split(Columns, "|", -1, 1)
        For i = 1 To UBound(Arr)
            Arr1 = Split(Arr(i), ":", -1, 1)
            sMasterFields = sMasterFields & Arr1(0) & ", "
        Next i
        sMasterFields = Mid(Trim(CStr(sMasterFields)), 1, Len(Trim(CStr(sMasterFields))) - 1)
        
        sDetailFields = "MasterKey, Line, "
        Arr = Split(ColumnsDet, "|", -1, 1)
        For i = 1 To UBound(Arr)
            Arr1 = Split(Arr(i), ":", -1, 1)
            sDetailFields = sDetailFields & Arr1(0) & ", "
        Next i
        sDetailFields = Mid(Trim(CStr(sDetailFields)), 1, Len(Trim(CStr(sDetailFields))) - 1)
        
        s = "SELECT tbl_Scoring_Team.PK, tbl_Scoring_Team.TeamID, " & _
            " SUM(tbl_Scoring_PlayerName.HandiCap) AS TotalHDCP, " & _
            " CONVERT(int, SUM(tbl_Scoring_PlayerName.HandiCap) / COUNT(tbl_Scoring_Team_Detail.PlayerKey)) AS TeamHDCP, " & _
            " (SELECT tbl_Scoring_TournamentInfo_Class.Class " & _
            " From tbl_Scoring_TournamentInfo_Class " & _
            " WHERE (tbl_Scoring_TournamentInfo_Class.TournamentKey = " & TournamentKey & ") " & _
            " AND (tbl_Scoring_TournamentInfo_Class.HFrom <= CONVERT(int, SUM(tbl_Scoring_PlayerName.HandiCap) / COUNT(tbl_Scoring_Team_Detail.PlayerKey))) " & _
            " AND (tbl_Scoring_TournamentInfo_Class.HTo >= CONVERT(int, SUM(tbl_Scoring_PlayerName.HandiCap) / COUNT(tbl_Scoring_Team_Detail.PlayerKey)))) AS TeamClass " & _
            " FROM tbl_Scoring_Team LEFT OUTER JOIN " & _
            " tbl_Scoring_Team_Detail ON tbl_Scoring_Team.PK = tbl_Scoring_Team_Detail.TeamKey LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " Where (tbl_Scoring_Team.TournamentKey = " & TournamentKey & ") " & _
            " GROUP BY tbl_Scoring_Team.PK, tbl_Scoring_Team.TeamID"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        While Not rs.EOF
            DoEvents
            cnt = cnt + 1
            dblGrossScore = 0
            dblNetScore = 0
            dblFront9 = 0
            dblBack9 = 0
            
            t = "SELECT MIN(Score) AS GrossScore " & _
                " From tbl_Scoring_ScoreCard_Team " & _
                " WHERE (TournamentKey = " & TournamentKey & ") " & _
                " AND (TeamKey = " & rs!PK & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                dblGrossScore = IIf(IsNull(rt!GrossScore), 0, rt!GrossScore)
                dblNetScore = IIf(CDbl(dblGrossScore) = 0, 0, CDbl(dblGrossScore) - CDbl(rs!TeamHDCP))
            End If
            rt.Close
            
            t = "SELECT Front9Score, Back9Score " & _
                " FROM tbl_Scoring_ScoreCard_Team " & _
                " WHERE (TournamentKey = " & TournamentKey & ") " & _
                " AND (TeamKey = " & rs!PK & ") " & _
                " AND (Score = " & CDbl(dblGrossScore) & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                dblFront9 = rt!Front9Score
                dblBack9 = rt!Back9Score
            End If
            rt.Close
            
            ConnOmega.Execute "INSERT INTO " & TableName & " " & _
                              " (" & sMasterFields & ") " & _
                              " VALUES (0, " & rs!PK & ", '" & rs!TeamID & "', " & _
                              " " & rs!TotalHDCP & ", " & rs!TeamHDCP & ", " & _
                              " '" & rs!TeamClass & "', " & CDbl(dblFront9) & ", " & _
                              " " & CDbl(dblBack9) & ", " & CDbl(dblGrossScore) & ", " & _
                              " " & CDbl(dblNetScore) & ")"
            j = 0
            MasterKey = 0
            t = "SELECT TOP 1 PK " & _
                " FROM " & TableName & " " & _
                " ORDER BY PK DESC"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                MasterKey = rt!PK
            End If
            rt.Close
            
            t = "SELECT LTRIM(RTRIM(tbl_Scoring_PlayerName.LastName)) + ',  ' + LTRIM(RTRIM(tbl_Scoring_PlayerName.FirstName)) + '  ' + LTRIM(RTRIM(tbl_Scoring_PlayerName.MiddleName)) AS PlayerName, " & _
                " tbl_Scoring_PlayerName.HandiCap " & _
                " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                " ORDER BY LTRIM(RTRIM(tbl_Scoring_PlayerName.LastName)) + ',  ' + LTRIM(RTRIM(tbl_Scoring_PlayerName.FirstName)) " & _
                " + '  ' + LTRIM(RTRIM(tbl_Scoring_PlayerName.MiddleName))"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            While Not rt.EOF
                j = j + 1
                ConnOmega.Execute "INSERT INTO " & DetailTableName & " " & _
                                  " (" & sDetailFields & ") " & _
                                  " VALUES (" & MasterKey & ", " & j & ", " & _
                                  " '" & FORMATSQL(rt!PlayerName) & "', " & _
                                  " " & CDbl(rt!HandiCap) & ")"
                rt.MoveNext
            Wend
            rt.Close
            
            UpdateProgress picProgressBar, cnt / rs.RecordCount
            
            rs.MoveNext
        Wend
        rs.Close
        
        i = 0
        Select Case cmbSortAll.ListIndex
            Case 0  '== Handicap
                If cmbPrintType.ListIndex = 2 Then
                    s = "SELECT * FROM " & TableName & " " & _
                        " ORDER BY TeamHDCP, TotalHDCP, TeamID"
                Else
                    s = "SELECT * FROM " & TableName & " " & _
                        " WHERE (TeamClass = '" & Right(cmbPrintType.List(cmbPrintType.ListIndex), 1) & "') " & _
                        " ORDER BY TeamHDCP, TotalHDCP, TeamID"
                End If
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                While Not rs.EOF
                    i = i + 1
                    ConnOmega.Execute "UPDATE " & TableName & " " & _
                                      " SET Sorting = " & i & " " & _
                                      " WHERE (PK = " & rs!PK & ")"
                    rs.MoveNext
                Wend
                rs.Close
            Case 1  '== Gross
                If cmbPrintType.ListIndex = 2 Then
                    s = "SELECT * FROM " & TableName & " " & _
                        " WHERE (GrossScore <> 0) " & _
                        " ORDER BY GrossScore, Back9, Front9"
                Else
                    s = "SELECT * FROM " & TableName & " " & _
                        " WHERE (TeamClass = '" & Right(cmbPrintType.List(cmbPrintType.ListIndex), 1) & "') " & _
                        " AND (GrossScore <> 0) " & _
                        " ORDER BY GrossScore, Back9, Front9"
                End If
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                While Not rs.EOF
                    i = i + 1
                    ConnOmega.Execute "UPDATE " & TableName & " " & _
                                      " SET Sorting = " & i & " " & _
                                      " WHERE (PK = " & rs!PK & ")"
                    rs.MoveNext
                Wend
                rs.Close
                
                If cmbPrintType.ListIndex = 2 Then
                    s = "SELECT * FROM " & TableName & " " & _
                        " WHERE (GrossScore = 0) " & _
                        " ORDER BY TeamHDCP"
                Else
                    s = "SELECT * FROM " & TableName & " " & _
                        " WHERE (TeamClass = '" & Right(cmbPrintType.List(cmbPrintType.ListIndex), 1) & "') " & _
                        " AND (GrossScore = 0) " & _
                        " ORDER BY TeamHDCP"
                End If
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                While Not rs.EOF
                    i = i + 1
                    ConnOmega.Execute "UPDATE " & TableName & " " & _
                                      " SET Sorting = " & i & " " & _
                                      " WHERE (PK = " & rs!PK & ")"
                    rs.MoveNext
                Wend
                rs.Close
                
            Case 2  '== Net
                
                If cmbPrintType.ListIndex = 2 Then
                    s = "SELECT * FROM " & TableName & " " & _
                        " WHERE (NetScore <> 0) " & _
                        " ORDER BY NetScore, Back9, Front9"
                Else
                    s = "SELECT * FROM " & TableName & " " & _
                        " WHERE (TeamClass = '" & Right(cmbPrintType.List(cmbPrintType.ListIndex), 1) & "') " & _
                        " AND (NetScore <> 0) " & _
                        " ORDER BY NetScore, Back9, Front9"
                End If
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                While Not rs.EOF
                    i = i + 1
                    ConnOmega.Execute "UPDATE " & TableName & " " & _
                                      " SET Sorting = " & i & " " & _
                                      " WHERE (PK = " & rs!PK & ")"
                    rs.MoveNext
                Wend
                rs.Close
                
                If cmbPrintType.ListIndex = 2 Then
                    s = "SELECT * FROM " & TableName & " " & _
                        " WHERE (NetScore = 0) " & _
                        " ORDER BY TeamHDCP"
                Else
                    s = "SELECT * FROM " & TableName & " " & _
                        " WHERE (TeamClass = '" & Right(cmbPrintType.List(cmbPrintType.ListIndex), 1) & "') " & _
                        " AND (NetScore = 0) " & _
                        " ORDER BY TeamHDCP"
                End If
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                While Not rs.EOF
                    i = i + 1
                    ConnOmega.Execute "UPDATE " & TableName & " " & _
                                      " SET Sorting = " & i & " " & _
                                      " WHERE (PK = " & rs!PK & ")"
                    rs.MoveNext
                Wend
                rs.Close
        End Select
        
        ColTop = 0: RowTop = 0
        ColCount = 0: RowCount = 0
        cnt = 0
        Set xlsApp = CreateObject("Excel.Application")
        With xlsApp
            .Visible = False
            
            .Workbooks.Add
            .DisplayAlerts = False
            .Workbooks(1).Sheets(1).Activate
            .Workbooks(1).Sheets(1).Name = cmbPrintType.List(cmbPrintType.ListIndex) & " (" & Replace(cmbSortAll.List(cmbSortAll.ListIndex), "SORT BY ", "") & ")" '"Report"
            .Workbooks(1).Sheets(2).Delete
            .Workbooks(1).Sheets(2).Delete
            
            With xlsApp.ActiveWorkbook.Sheets(1)
                If cmbPrintType.ListIndex = 2 Then
                    s = "SELECT * FROM " & TableName & "" & _
                        " ORDER BY Sorting"
                Else
                    s = "SELECT * FROM " & TableName & "" & _
                        " WHERE (TeamClass = '" & Right(cmbPrintType.List(cmbPrintType.ListIndex), 1) & "') " & _
                        " ORDER BY Sorting"
                End If
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                '== Header
                RowCount = RowCount + 1
                ColCount = ColCount + 1
                strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 1) > 26, 64 + 1, 64) + rs.Fields.Count - 1)) & CStr(RowCount)
                .Range(strRange, strRange1).Select
                xlsApp.Selection.Merge
                
                strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                .Range(strRange).Value = gbl_CompanyName
                .Range(strRange).Font.Name = "Script MT Bold" '"Tahoma"
                .Range(strRange).Font.Size = 10
                .Range(strRange).Font.Bold = True
                .Range(strRange).HorizontalAlignment = 3
                .Range(strRange).VerticalAlignment = 2
                
                ColCount = 0
                RowCount = RowCount + 1
                ColCount = ColCount + 1
                strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 1) > 26, 64 + 1, 64) + rs.Fields.Count - 1)) & CStr(RowCount)
                .Range(strRange, strRange1).Select
                xlsApp.Selection.Merge
                
                strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                .Range(strRange).Value = gbl_CompanyAddress1
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                .Range(strRange).HorizontalAlignment = 3
                .Range(strRange).VerticalAlignment = 2
                
                ColCount = 0
                RowCount = RowCount + 1
                ColCount = ColCount + 1
                strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 1) > 26, 64 + 1, 64) + rs.Fields.Count - 1)) & CStr(RowCount)
                .Range(strRange, strRange1).Select
                xlsApp.Selection.Merge
                
                strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                .Range(strRange).Value = gbl_CompanyAddress2
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                .Range(strRange).HorizontalAlignment = 3
                .Range(strRange).VerticalAlignment = 2
                
                ColCount = 0
                RowCount = RowCount + 1
                ColCount = ColCount + 1
                strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 1) > 26, 64 + 1, 64) + rs.Fields.Count - 1)) & CStr(RowCount)
                .Range(strRange, strRange1).Select
                xlsApp.Selection.Merge
                
                strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                .Range(strRange).Value = ""
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                .Rows(RowCount).RowHeight = 9
                .Range(strRange).HorizontalAlignment = 3
                .Range(strRange).VerticalAlignment = 2
                
                ColCount = 0
                RowCount = RowCount + 1
                ColCount = ColCount + 1
                strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 1) > 26, 64 + 1, 64) + rs.Fields.Count - 1)) & CStr(RowCount)
                .Range(strRange, strRange1).Select
                xlsApp.Selection.Merge
                strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                .Range(strRange).Value = TournamentName
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 9
                .Range(strRange).Font.Bold = True
                .Range(strRange).HorizontalAlignment = 3
                .Range(strRange).VerticalAlignment = 2
                
                ColCount = 0
                RowCount = RowCount + 1
                ColCount = ColCount + 1
                strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 1) > 26, 64 + 1, 64) + rs.Fields.Count - 1)) & CStr(RowCount)
                .Range(strRange, strRange1).Select
                xlsApp.Selection.Merge
                strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                .Range(strRange).Value = TournamentRange
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                .Range(strRange).HorizontalAlignment = 3
                .Range(strRange).VerticalAlignment = 2
                                
                ColCount = 0
                RowCount = RowCount + 1
                ColCount = ColCount + 1
                strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 1) > 26, 64 + 1, 64) + rs.Fields.Count - 1)) & CStr(RowCount)
                .Range(strRange, strRange1).Select
                xlsApp.Selection.Merge
                
                strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                .Range(strRange).Value = ""
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                .Rows(RowCount).RowHeight = 9
                .Range(strRange).HorizontalAlignment = 3
                .Range(strRange).VerticalAlignment = 2
                
                '===
                ColCount = 0
                RowCount = RowCount + 1
                ColCount = ColCount + 1
                strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                .Range(strRange).Value = "Team ID"
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                .Range(strRange).HorizontalAlignment = 3
                .Range(strRange).Interior.ColorIndex = 15
                .Range(strRange).Interior.Pattern = 1 'xlSolid
                .Range(strRange).Select
                xlsApp.Selection.Borders.LineStyle = 1
                    
                ColCount = ColCount + 1
                strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                .Range(strRange).Value = "Player Name"
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                .Range(strRange).HorizontalAlignment = 1
                .Range(strRange).Interior.ColorIndex = 15
                .Range(strRange).Interior.Pattern = 1
                .Range(strRange).Select
                xlsApp.Selection.Borders.LineStyle = 1
                
                ColCount = ColCount + 1
                strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                .Range(strRange).Value = "H D C P"
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                .Range(strRange).HorizontalAlignment = 4
                .Range(strRange).Interior.ColorIndex = 15
                .Range(strRange).Interior.Pattern = 1 'xlSolid
                .Range(strRange).Select
                xlsApp.Selection.Borders.LineStyle = 1
                
                ColCount = ColCount + 1
                strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                .Range(strRange).Value = "Total HDCP"
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                .Range(strRange).HorizontalAlignment = 3
                .Range(strRange).Interior.ColorIndex = 15
                .Range(strRange).Interior.Pattern = 1 'xlSolid
                .Range(strRange).Select
                xlsApp.Selection.Borders.LineStyle = 1
                
                ColCount = ColCount + 1
                strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                .Range(strRange).Value = "Team HDCP"
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                .Range(strRange).HorizontalAlignment = 3
                .Range(strRange).Interior.ColorIndex = 15
                .Range(strRange).Interior.Pattern = 1 'xlSolid
                .Range(strRange).Select
                xlsApp.Selection.Borders.LineStyle = 1
                
                ColCount = ColCount + 1
                strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                .Range(strRange).Value = "Class"
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                .Range(strRange).HorizontalAlignment = 3
                .Range(strRange).Interior.ColorIndex = 15
                .Range(strRange).Interior.Pattern = 1 'xlSolid
                .Range(strRange).Select
                xlsApp.Selection.Borders.LineStyle = 1
                
                ColCount = ColCount + 1
                strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                .Range(strRange).Value = "Front 9"
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                .Range(strRange).HorizontalAlignment = 3
                .Range(strRange).Interior.ColorIndex = 15
                .Range(strRange).Interior.Pattern = 1 'xlSolid
                .Range(strRange).Select
                xlsApp.Selection.Borders.LineStyle = 1
                
                ColCount = ColCount + 1
                strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                .Range(strRange).Value = "Back 9"
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                .Range(strRange).HorizontalAlignment = 3
                .Range(strRange).Interior.ColorIndex = 15
                .Range(strRange).Interior.Pattern = 1 'xlSolid
                .Range(strRange).Select
                xlsApp.Selection.Borders.LineStyle = 1
                
                If cmbSortAll.ListIndex = 1 Then
                    
                    ColCount = ColCount + 1
                    strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                    .Range(strRange).Value = "Net"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    .Range(strRange).HorizontalAlignment = 3
                    .Range(strRange).Interior.ColorIndex = 15
                    .Range(strRange).Interior.Pattern = 1 'xlSolid
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    ColCount = ColCount + 1
                    strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                    .Range(strRange).Value = "Gross"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    .Range(strRange).HorizontalAlignment = 3
                    .Range(strRange).Interior.ColorIndex = 15
                    .Range(strRange).Interior.Pattern = 1 'xlSolid
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                Else
                
                    ColCount = ColCount + 1
                    strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                    .Range(strRange).Value = "Gross"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    .Range(strRange).HorizontalAlignment = 3
                    .Range(strRange).Interior.ColorIndex = 15
                    .Range(strRange).Interior.Pattern = 1 'xlSolid
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    ColCount = ColCount + 1
                    strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                    .Range(strRange).Value = "Net"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    .Range(strRange).HorizontalAlignment = 3
                    .Range(strRange).Interior.ColorIndex = 15
                    .Range(strRange).Interior.Pattern = 1 'xlSolid
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                
                End If
                
                HeaderRow = RowCount
                                
                While Not rs.EOF
                    cnt = cnt + 1
                    ColCount = 0
                    RowCount = RowCount + 1
                    ColCount = ColCount + 1
                    strRange = ""
                    strRange1 = ""
                    
                    RowFrom = RowCount
                    RowTo = RowFrom
                    
                    t = "SELECT * FROM " & DetailTableName & " " & _
                        " WHERE (MasterKey = " & rs!PK & ") " & _
                        " ORDER BY Line"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        RowTo = RowTo + 1
                        rt.MoveNext
                    Wend
                    rt.Close
                    RowTo = RowTo - 1
                    
                    strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
                    strRange1 = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowTo)
                    .Range(strRange, strRange1).Select
                    xlsApp.Selection.Merge
                    strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
                    .Range(strRange).Value = rs!TeamID
                    .Range(strRange).Font.Name = "Courier New"
                    .Range(strRange).Font.Size = 13
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCount).ColumnWidth = 7
                    .Range(strRange).HorizontalAlignment = 3
                    .Range(strRange).VerticalAlignment = 2
                    
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    ColCount = ColCount + 1
                    
                    '==
                    ColCountDet = ColCount
                    RowCountDet = RowFrom
                    t = "SELECT * FROM " & DetailTableName & " " & _
                        " WHERE (MasterKey = " & rs!PK & ") " & _
                        " ORDER BY Line"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    RowFrom = RowCount
                    RowTo = RowFrom
                    While Not rt.EOF
                        strRange = (Chr$(IIf(CDbl(ColCountDet) > 26, 64 + 1, 64) + ColCountDet)) & CStr(RowCountDet)
                        .Range(strRange).Value = rt!PlayerName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 10
                        .Range(strRange).Font.Color = vbBlue
                        .Columns(ColCountDet).ColumnWidth = 28
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        ColCountDet = ColCountDet + 1
                        strRange = (Chr$(IIf(CDbl(ColCountDet) > 26, 64 + 1, 64) + ColCountDet)) & CStr(RowCountDet)
                        .Range(strRange).Value = rt!HDCP
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 10
                        .Range(strRange).Font.Color = vbRed
                        .Columns(ColCountDet).ColumnWidth = 7
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        RowCountDet = RowCountDet + 1
                        ColCountDet = ColCount
                        RowTo = RowTo + 1
                        rt.MoveNext
                    Wend
                    ColCount = ColCount + (rt.Fields.Count) - 2
                    rt.Close
                    RowTo = RowTo - 1
                    
                    strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
                    strRange1 = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowTo)
                    .Range(strRange, strRange1).Select
                    xlsApp.Selection.Merge
                    strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
                    .Range(strRange).Value = rs!TotalHDCP
                    .Range(strRange).Font.Name = "Courier New"
                    .Range(strRange).Font.Size = 13
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 3
                    .Range(strRange).VerticalAlignment = 2
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                        
                    ColCount = ColCount + 1
                    strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
                    strRange1 = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowTo)
                    .Range(strRange, strRange1).Select
                    xlsApp.Selection.Merge
                    strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
                    .Range(strRange).Value = rs!TeamHDCP
                    .Range(strRange).Font.Name = "Courier New"
                    .Range(strRange).Font.Size = 13
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 3
                    .Range(strRange).VerticalAlignment = 2
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    ColCount = ColCount + 1
                    strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
                    strRange1 = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowTo)
                    .Range(strRange, strRange1).Select
                    xlsApp.Selection.Merge
                    strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
                    .Range(strRange).Value = rs!TeamClass
                    .Range(strRange).Font.Name = "Courier New"
                    .Range(strRange).Font.Size = 13
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 3
                    .Range(strRange).VerticalAlignment = 2
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    ColCount = ColCount + 1
                    strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
                    strRange1 = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowTo)
                    .Range(strRange, strRange1).Select
                    xlsApp.Selection.Merge
                    strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
                    .Range(strRange).Value = rs!Front9
                    .Range(strRange).Font.Name = "Courier New"
                    .Range(strRange).Font.Size = 13
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 3
                    .Range(strRange).VerticalAlignment = 2
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    ColCount = ColCount + 1
                    strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
                    strRange1 = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowTo)
                    .Range(strRange, strRange1).Select
                    xlsApp.Selection.Merge
                    strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
                    .Range(strRange).Value = rs!Back9
                    .Range(strRange).Font.Name = "Courier New"
                    .Range(strRange).Font.Size = 13
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 3
                    .Range(strRange).VerticalAlignment = 2
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    If cmbSortAll.ListIndex = 1 Then
                    
                        ColCount = ColCount + 1
                        strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
                        strRange1 = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowTo)
                        .Range(strRange, strRange1).Select
                        xlsApp.Selection.Merge
                        strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
                        .Range(strRange).Value = rs!NetScore
                        .Range(strRange).Font.Name = "Courier New"
                        .Range(strRange).Font.Size = 13
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).HorizontalAlignment = 3
                        .Range(strRange).VerticalAlignment = 2
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        ColCount = ColCount + 1
                        strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
                        strRange1 = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowTo)
                        .Range(strRange, strRange1).Select
                        xlsApp.Selection.Merge
                        strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
                        .Range(strRange).Value = rs!GrossScore
                        .Range(strRange).Font.Name = "Courier New"
                        .Range(strRange).Font.Size = 13
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).HorizontalAlignment = 3
                        .Range(strRange).VerticalAlignment = 2
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                    
                    Else
                    
                        ColCount = ColCount + 1
                        strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
                        strRange1 = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowTo)
                        .Range(strRange, strRange1).Select
                        xlsApp.Selection.Merge
                        strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
                        .Range(strRange).Value = rs!GrossScore
                        .Range(strRange).Font.Name = "Courier New"
                        .Range(strRange).Font.Size = 13
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).HorizontalAlignment = 3
                        .Range(strRange).VerticalAlignment = 2
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        ColCount = ColCount + 1
                        strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
                        strRange1 = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowTo)
                        .Range(strRange, strRange1).Select
                        xlsApp.Selection.Merge
                        strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
                        .Range(strRange).Value = rs!NetScore
                        .Range(strRange).Font.Name = "Courier New"
                        .Range(strRange).Font.Size = 13
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).HorizontalAlignment = 3
                        .Range(strRange).VerticalAlignment = 2
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                    
                    End If
                    
                    RowCount = RowTo
                    
                    UpdateProgress_Caption "Generating Excel Output", picProgressBar, cnt / rs.RecordCount
                    
                    rs.MoveNext
                Wend
                rs.Close
                
                .PageSetup.PaperSize = 1 'Letter
                .PageSetup.Orientation = 1 '2 'LandScape
                .PageSetup.TopMargin = 3
                .PageSetup.LeftMargin = 3
                .PageSetup.RightMargin = 3
                .PageSetup.BottomMargin = 3
'                .CenterHorizontally = 1 'True
'                .CenterVertically = True
                .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
            End With
            
'            .Locked = True
            
            If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
            .ActiveWorkbook.SaveAs Filename:=WorkbookName
            .Visible = True
            Set xlsApp = Nothing
        End With
        picProgress.Visible = False
        picMain.Enabled = True
        picToolbar.Enabled = True
        
        Screen.MousePointer = vbDefault
        
End Select

Exit Sub
ErrorHandler:
Exit Sub

Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub cmdOKSearch_Click()
If cmbDate.ListIndex = -1 Then Exit Sub
BROWSER cmbDate.ItemData(cmbDate.ListIndex), "is_FIND"
cmdCancelSearch_Click
End Sub

Private Sub Command1_Click()
Dim i
Dim WorkbookName    As String
Dim ColTop, RowTop, ColCount, RowCount, strRange, strRange1, ColCountDet, RowCountDet, _
RowFrom, RowTo
WorkbookName = "C:\Testing"

Set xlsApp = CreateObject("Excel.Application")
xlsApp.Visible = False

xlsApp.Workbooks.Add
xlsApp.DisplayAlerts = False
xlsApp.Workbooks(1).Sheets(1).Activate

ColTop = 0: RowTop = 0
ColCount = 0: RowCount = 0

s = "SELECT tbl_Scoring_Team.PK, tbl_Scoring_Team.TeamID, " & _
    " SUM(tbl_Scoring_PlayerName.HandiCap) AS TotalHDCP, " & _
    " CONVERT(int, SUM(tbl_Scoring_PlayerName.HandiCap) / COUNT(tbl_Scoring_Team_Detail.PlayerKey)) AS TeamHDCP, " & _
    " COUNT(tbl_Scoring_Team_Detail.PlayerKey) AS NoPlayer " & _
    " FROM tbl_Scoring_Team LEFT OUTER JOIN " & _
    " tbl_Scoring_Team_Detail ON tbl_Scoring_Team.PK = tbl_Scoring_Team_Detail.TeamKey LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " Where (tbl_Scoring_Team.TournamentKey = 2) " & _
    " GROUP BY tbl_Scoring_Team.PK, tbl_Scoring_Team.TeamID " & _
    " ORDER BY tbl_Scoring_Team.TeamID"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
RowCount = RowCount + 1
RowFrom = RowCount

ColCount = ColCount + 1
strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Value = "TEAM ID"
ColCount = ColCount + 1
strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
ColCount = ColCount + 1
strRange1 = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
xlsApp.ActiveWorkbook.Sheets(1).Range(strRange, strRange1).Select
xlsApp.Selection.Merge
ColCount = ColCount - 1
strRange1 = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
ColCount = ColCount + 1
xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Value = "PLAYER'S"
ColCount = ColCount + 1
strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Value = "TOTAL HDCP"
ColCount = ColCount + 1
strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Value = "TEAM HDCP"

While Not rs.EOF
    ColCount = 0
    RowCount = RowCount + 1
    ColCount = ColCount + 1
    strRange = ""
    strRange1 = ""
    
    t = "SELECT tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
        " tbl_Scoring_PlayerName.HandiCap " & _
        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
        " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
        " ORDER BY tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    RowFrom = RowCount
    RowTo = RowFrom
    While Not rt.EOF
        RowTo = RowTo + 1
        rt.MoveNext
    Wend
    rt.Close
    RowTo = RowTo - 1
    
    strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
    strRange1 = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowTo)
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange, strRange1).Select
    xlsApp.Selection.Merge
    strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Value = rs!TeamID
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Font.Name = "Courier New"
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Font.Size = 13
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Font.Bold = True
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).HorizontalAlignment = 3
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).VerticalAlignment = 2
    
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Select
    xlsApp.Selection.Borders.LineStyle = 1
    
    ColCount = ColCount + 1
    
    '==
    ColCountDet = ColCount
    RowCountDet = RowFrom
    t = "SELECT tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
        " tbl_Scoring_PlayerName.HandiCap " & _
        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
        " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
        " ORDER BY tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    RowFrom = RowCount
    RowTo = RowFrom
    While Not rt.EOF
        strRange = (Chr$(IIf(CDbl(ColCountDet) > 26, 64 + 1, 64) + ColCountDet)) & CStr(RowCountDet)
        xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Value = rt!PlayerName
        xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Font.Name = "Tahoma"
        xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Font.Size = 10
        xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Font.Color = vbBlue
        xlsApp.ActiveWorkbook.Sheets(1).Columns(ColCountDet).ColumnWidth = 35
        xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Select
        xlsApp.Selection.Borders.LineStyle = 1
        
        ColCountDet = ColCountDet + 1
        strRange = (Chr$(IIf(CDbl(ColCountDet) > 26, 64 + 1, 64) + ColCountDet)) & CStr(RowCountDet)
        xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Value = rt!HandiCap
        xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Font.Name = "Tahoma"
        xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Font.Size = 10
        xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Font.Color = vbRed
        xlsApp.ActiveWorkbook.Sheets(1).Columns(ColCountDet).ColumnWidth = 10
        xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Select
        xlsApp.Selection.Borders.LineStyle = 1
        
        RowCountDet = RowCountDet + 1
        ColCountDet = ColCount
        RowTo = RowTo + 1
        rt.MoveNext
    Wend
    ColCount = ColCount + rt.Fields.Count
    rt.Close
    RowTo = RowTo - 1
    
    strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
    strRange1 = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowTo)
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange, strRange1).Select
    xlsApp.Selection.Merge
    strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Value = rs!TotalHDCP
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Font.Name = "Courier New"
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Font.Size = 13
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Font.Bold = True
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).HorizontalAlignment = 3
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).VerticalAlignment = 2
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Select
    xlsApp.Selection.Borders.LineStyle = 1
        
    ColCount = ColCount + 1
    strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
    strRange1 = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowTo)
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange, strRange1).Select
    xlsApp.Selection.Merge
    strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowFrom)
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Value = rs!TeamHDCP
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Font.Name = "Courier New"
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Font.Size = 13
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Font.Bold = True
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).HorizontalAlignment = 3
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).VerticalAlignment = 2
    xlsApp.ActiveWorkbook.Sheets(1).Range(strRange).Select
    xlsApp.Selection.Borders.LineStyle = 1
    
'    t = ""
'    If rt.State = adStateOpen Then rt.Close
'    rt.Open t, ConnOmega
'    If rt.RecordCount > 0 Then
'
'    End If
'    rt.Close
    
    RowCount = RowTo
    
    rs.MoveNext
Wend
rs.Close

xlsApp.ActiveWorkbook.Sheets(1).PageSetup.PaperSize = 1 'Letter
xlsApp.ActiveWorkbook.Sheets(1).PageSetup.Orientation = 2 'LandScape
xlsApp.ActiveWorkbook.Sheets(1).PageSetup.TopMargin = 0.75
xlsApp.ActiveWorkbook.Sheets(1).PageSetup.LeftMargin = 0.75
xlsApp.ActiveWorkbook.Sheets(1).PageSetup.RightMargin = 0.75
xlsApp.ActiveWorkbook.Sheets(1).PageSetup.BottomMargin = 0.75
xlsApp.ActiveWorkbook.Sheets(1).PageSetup.PrintTitleRows = "$1" & ":$1"

'xlsApp.ActiveWorkbook.Sheets(1).Top = 0.75
'xlsApp.ActiveWorkbook.Sheets(1).PageSetup.Margins.Left = 0.75
'xlsApp.ActiveWorkbook.Sheets(1).PageSetup.Margins.Right = 0.75
'xlsApp.ActiveWorkbook.Sheets(1).PageSetup.Margins.Bottom = 0.75

'RowCount = RowCount + 1
'ColCount = ColCount + 1
'strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
'RowCount = RowCount + 3
'strRange1 = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
'
'xlsApp.ActiveWorkbook.Sheets(1).Range(strRange, strRange1).Select
'xlsApp.Selection.Merge

If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

xlsApp.Visible = True
End Sub

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
If TRANSACTIONTYPE = is_REFRESH Then
    BROWSER GetSetting(App.EXEName, "ScoreCardControl", "ScoreCardCtrl", ""), "is_LOAD"
    If Trim(txtPlayer.Text) = "" Then BROWSER GetSetting(App.EXEName, "ScoreCardControl", "ScoreCardCtrl", ""), "is_HOME"
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert:   PRESS_INSERT
    Case vbKeyF2:       PRESS_F2
    Case vbKeyDelete:   PRESS_DELETE
    Case vbKeyF5:       PRESS_F5
    Case vbKeyF6:       PRESS_F6
    Case vbKeyF9:       PRESS_F9
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "ScoreCardControl", "ScoreCardCtrl", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "ScoreCardControl", "ScoreCardCtrl", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "ScoreCardControl", "ScoreCardCtrl", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "ScoreCardControl", "ScoreCardCtrl", ""), "is_END"
    Case Else: Exit Sub
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Height = 6825
Me.Width = 12195
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2

With cmbPrintType
    .Clear
    .AddItem "RESULT"
    .AddItem "SCORE CARD"
    .AddItem "ALL SCORE"
    s = "SELECT Class " & _
        " From tbl_Scoring_TournamentInfo_Class " & _
        " Where (TournamentKey = " & TournamentKey & ") " & _
        " ORDER BY Class"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        .AddItem "CLASS " & rs!Class
        rs.MoveNext
    Wend
    rs.Close
End With

s = "SELECT tbl_Scoring_TournamentInfo.* " & _
    " FROM tbl_Scoring_TournamentInfo " & _
    " WHERE (PK = " & TournamentKey & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    TourNoOfPlays = rs!NoofPlays
    txtTournament.Text = rs!TournamentName
    txtTourDate.Text = Format(rs!TournamentStart, "mm/dd/yyyy") & " - " & Format(rs!TournamentEnd, "mm/dd/yyyy")
    dDateEnd = Format(rs!TournamentEnd, "mm/dd/yyyy")
End If
rs.Close

'Me.Caption = "Score Card (Team)"
With FGrid
    .BackColor = &HC6B8A4
    .BackColorBkg = &HC6B8A4
    .BackColorFixed = &HC6B8A4
    .BackColorSel = &HC6B8A4
    .ForeColor = &H80000012
    .ForeColorFixed = &H80000012
    .ForeColorSel = &H80000012
    .GridColor = &H80000012
    .GridColorFixed = &H80000012
End With
picMain.ZOrder 1
picScoreMain.Width = 11780 '12330
picScoreMain.Height = 2400

LOAD_CARD dDateEnd, FGrid

CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH

'BROWSER GetSetting(App.EXEName, "ScoreCardControl", "ScoreCardCtrl", ""), "is_LOAD"
'If Trim(txtPlayer.Text) = "" Then BROWSER GetSetting(App.EXEName, "ScoreCardControl", "ScoreCardCtrl", ""), "is_HOME"

Dim tmp As Long
tmp = SetWindowLong(txtSearchAdd.hwnd, GWL_STYLE, GetWindowLong(txtSearchAdd.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picSearchAdd.Visible = True Then Cancel = -1
If picSearch.Visible = True Then Cancel = -1
If picPrint.Visible = True Then Cancel = -1
If picProgress.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
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

Private Sub lstResultAdd_Click()
If lstResultAdd.ListIndex = -1 Then lstTeamAdd.Clear: Exit Sub
lstTeamAdd.Clear
s = "SELECT tbl_Scoring_Team.PK " & _
    " FROM tbl_Scoring_Team LEFT OUTER JOIN " & _
    " tbl_Scoring_Team_Detail ON tbl_Scoring_Team.PK = tbl_Scoring_Team_Detail.TeamKey LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_Team_Detail.PlayerKey = " & lstResultAdd.ItemData(lstResultAdd.ListIndex) & ")"
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
        " AND (tbl_Scoring_Team_Detail.PlayerKey <> " & lstResultAdd.ItemData(lstResultAdd.ListIndex) & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        lstTeamAdd.AddItem rt!TeamID & " - " & rt!LastName & ", " & rt!FirstName
        lstTeamAdd.ItemData(lstTeamAdd.NewIndex) = rt!PK
        rt.MoveNext
    Wend
    rt.Close
    rs.MoveNext
Wend
rs.Close
If lstTeamAdd.ListCount Then lstTeamAdd.ListIndex = 0
End Sub

Private Sub lstResultAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstTeamAdd.SetFocus
End Sub

Private Sub lstTeam_Click()
If lstTeam.ListIndex = -1 Then cmbDate.Clear: Exit Sub
cmbDate.Clear
s = "SELECT DDate, PK " & _
    " From tbl_Scoring_ScoreCard_Team " & _
    " Where (TeamKey = " & lstTeam.ItemData(lstTeam.ListIndex) & ") " & _
    " ORDER BY DDate"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbDate.AddItem Format(rs!dDate, "mm/dd/yyyy")
    cmbDate.ItemData(cmbDate.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If cmbDate.ListCount Then cmbDate.ListIndex = 0
End Sub

Private Sub lstTeam_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbDate.SetFocus
End Sub

Private Sub lstTeamAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtDateAdd.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "ScoreCardControl", "ScoreCardCtrl", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "ScoreCardControl", "ScoreCardCtrl", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "ScoreCardControl", "ScoreCardCtrl", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "ScoreCardControl", "ScoreCardCtrl", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":   PRESS_F9
    Case "Close":   PRESS_ESCAPE
    Case Else:  Exit Sub
End Select
End Sub

Private Sub txtDateAdd_GotFocus()
HTEXT txtDateAdd
End Sub

Private Sub txtDateAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAdd_Click
End Sub

Private Sub txtDateAdd_LostFocus()
If IsDate(txtDateAdd.Text) = True Then
    txtDateAdd.Text = Format(FormatDateTime(txtDateAdd.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub

Private Sub txtDateFrom_GotFocus()
HTEXT txtDateFrom
End Sub

Private Sub txtDateFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtDateTo.SetFocus
End Sub

Private Sub txtDateFrom_LostFocus()
If IsDate(txtDateFrom.Text) = True Then
    txtDateFrom.Text = Format(FormatDateTime(txtDateFrom.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub

Private Sub txtDateTo_GotFocus()
HTEXT txtDateTo
End Sub

Private Sub txtDateTo_LostFocus()
If IsDate(txtDateTo.Text) = True Then
    txtDateTo.Text = Format(FormatDateTime(txtDateTo.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub

Private Sub txtGrossPts_Change(Index As Integer)
If Index >= 0 And Index <= 8 Then
    txtGrossPtsF.Text = RETURNTEXTVALUE(txtGrossPts(0)) + _
                          RETURNTEXTVALUE(txtGrossPts(1)) + _
                          RETURNTEXTVALUE(txtGrossPts(2)) + _
                          RETURNTEXTVALUE(txtGrossPts(3)) + _
                          RETURNTEXTVALUE(txtGrossPts(4)) + _
                          RETURNTEXTVALUE(txtGrossPts(5)) + _
                          RETURNTEXTVALUE(txtGrossPts(6)) + _
                          RETURNTEXTVALUE(txtGrossPts(7)) + _
                          RETURNTEXTVALUE(txtGrossPts(8))
ElseIf Index >= 9 And Index <= 17 Then
    txtGrossPtsB.Text = RETURNTEXTVALUE(txtGrossPts(9)) + _
                          RETURNTEXTVALUE(txtGrossPts(10)) + _
                          RETURNTEXTVALUE(txtGrossPts(11)) + _
                          RETURNTEXTVALUE(txtGrossPts(12)) + _
                          RETURNTEXTVALUE(txtGrossPts(13)) + _
                          RETURNTEXTVALUE(txtGrossPts(14)) + _
                          RETURNTEXTVALUE(txtGrossPts(15)) + _
                          RETURNTEXTVALUE(txtGrossPts(16)) + _
                          RETURNTEXTVALUE(txtGrossPts(17))
End If
End Sub

Private Sub txtGrossPtsB_Change()
txtGrossPtsTot.Text = RETURNTEXTVALUE(txtGrossPtsF) + _
                      RETURNTEXTVALUE(txtGrossPtsB)
'txtSGrossB.Text = RETURNTEXTVALUE(txtGrossPtsB)
End Sub

Private Sub txtGrossPtsF_Change()
txtGrossPtsTot.Text = RETURNTEXTVALUE(txtGrossPtsF) + _
                      RETURNTEXTVALUE(txtGrossPtsB)
'txtSGrossF.Text = RETURNTEXTVALUE(txtGrossPtsF)
End Sub

Private Sub txtGrossScore_Change(Index As Integer)

If RETURNTEXTVALUE(txtGrossScore(Index)) <= 0 Then txtGrossPts(Index).Text = "0": txtNetPts(Index).Text = "0": Exit Sub

Dim dblPar, dblHandicap
With FGrid
    Select Case Index
        Case 0
            dblPar = .TextMatrix(1, 2)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 2)
            'txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 1
            dblPar = .TextMatrix(1, 3)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 3)
            'txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 2
            dblPar = .TextMatrix(1, 4)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 4)
            'txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 3
            dblPar = .TextMatrix(1, 5)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 5)
            'txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 4
            dblPar = .TextMatrix(1, 6)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 6)
            'txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 5
            dblPar = .TextMatrix(1, 7)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 7)
            'txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 6
            dblPar = .TextMatrix(1, 8)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 8)
            'txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 7
            dblPar = .TextMatrix(1, 9)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 9)
            'txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 8
            dblPar = .TextMatrix(1, 10)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 10)
            'txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 9
            dblPar = .TextMatrix(1, 12)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 12)
            'txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 10
            dblPar = .TextMatrix(1, 13)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 13)
            'txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 11
            dblPar = .TextMatrix(1, 14)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 14)
            'txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 12
            dblPar = .TextMatrix(1, 15)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 15)
            'txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 13
            dblPar = .TextMatrix(1, 16)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 16)
            'txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 14
            dblPar = .TextMatrix(1, 17)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 17)
            'txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 15
            dblPar = .TextMatrix(1, 18)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 18)
            'txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 16
            dblPar = .TextMatrix(1, 19)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 19)
            'txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 17
            dblPar = .TextMatrix(1, 20)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 20)
            'txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
    End Select
End With
If Index >= 0 And Index <= 8 Then
    txtGrossScoreF.Text = RETURNTEXTVALUE(txtGrossScore(0)) + _
                          RETURNTEXTVALUE(txtGrossScore(1)) + _
                          RETURNTEXTVALUE(txtGrossScore(2)) + _
                          RETURNTEXTVALUE(txtGrossScore(3)) + _
                          RETURNTEXTVALUE(txtGrossScore(4)) + _
                          RETURNTEXTVALUE(txtGrossScore(5)) + _
                          RETURNTEXTVALUE(txtGrossScore(6)) + _
                          RETURNTEXTVALUE(txtGrossScore(7)) + _
                          RETURNTEXTVALUE(txtGrossScore(8))
ElseIf Index >= 9 And Index <= 17 Then
    txtGrossScoreB.Text = RETURNTEXTVALUE(txtGrossScore(9)) + _
                          RETURNTEXTVALUE(txtGrossScore(10)) + _
                          RETURNTEXTVALUE(txtGrossScore(11)) + _
                          RETURNTEXTVALUE(txtGrossScore(12)) + _
                          RETURNTEXTVALUE(txtGrossScore(13)) + _
                          RETURNTEXTVALUE(txtGrossScore(14)) + _
                          RETURNTEXTVALUE(txtGrossScore(15)) + _
                          RETURNTEXTVALUE(txtGrossScore(16)) + _
                          RETURNTEXTVALUE(txtGrossScore(17))
End If
End Sub

Private Sub txtGrossScore_GotFocus(Index As Integer)
txtGrossScore(Index).Text = IIf(RETURNTEXTVALUE(txtGrossScore(Index)) <= 0, "", txtGrossScore(Index).Text)
txtGrossScore(Index).Alignment = 0
HTEXT txtGrossScore(Index)
End Sub

Private Sub txtGrossScore_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    Select Case Index
        Case 0: txtGrossScore(1).SetFocus
        Case 1: txtGrossScore(2).SetFocus
        Case 2: txtGrossScore(3).SetFocus
        Case 3: txtGrossScore(4).SetFocus
        Case 4: txtGrossScore(5).SetFocus
        Case 5: txtGrossScore(6).SetFocus
        Case 6: txtGrossScore(7).SetFocus
        Case 7: txtGrossScore(8).SetFocus
        Case 8: txtGrossScore(9).SetFocus
        Case 9: txtGrossScore(10).SetFocus
        Case 10: txtGrossScore(11).SetFocus
        Case 11: txtGrossScore(12).SetFocus
        Case 12: txtGrossScore(13).SetFocus
        Case 13: txtGrossScore(14).SetFocus
        Case 14: txtGrossScore(15).SetFocus
        Case 15: txtGrossScore(16).SetFocus
        Case 16: txtGrossScore(17).SetFocus
        Case 17: txtGrossScore(0).SetFocus
    End Select
ElseIf KeyCode = vbKeyUp Then
    Select Case Index
        Case 0: txtGrossScore(17).SetFocus
        Case 1: txtGrossScore(0).SetFocus
        Case 2: txtGrossScore(1).SetFocus
        Case 3: txtGrossScore(2).SetFocus
        Case 4: txtGrossScore(3).SetFocus
        Case 5: txtGrossScore(4).SetFocus
        Case 6: txtGrossScore(5).SetFocus
        Case 7: txtGrossScore(6).SetFocus
        Case 8: txtGrossScore(7).SetFocus
        Case 9: txtGrossScore(8).SetFocus
        Case 10: txtGrossScore(9).SetFocus
        Case 11: txtGrossScore(10).SetFocus
        Case 12: txtGrossScore(11).SetFocus
        Case 13: txtGrossScore(12).SetFocus
        Case 14: txtGrossScore(13).SetFocus
        Case 15: txtGrossScore(14).SetFocus
        Case 16: txtGrossScore(15).SetFocus
        Case 17: txtGrossScore(16).SetFocus
    End Select
End If
End Sub

Private Sub txtGrossScore_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtGrossScore_LostFocus(Index As Integer)
txtGrossScore(Index).Text = IIf(RETURNTEXTVALUE(txtGrossScore(Index)) <= 0, "", txtGrossScore(Index).Text)
txtGrossScore(Index).Alignment = 1
End Sub

Private Sub txtGrossScoreB_Change()
txtGrossScoreTot.Text = RETURNTEXTVALUE(txtGrossScoreF) + _
                        RETURNTEXTVALUE(txtGrossScoreB)
txtScoreGrossB.Text = txtGrossScoreB.Text
End Sub

Private Sub txtGrossScoreF_Change()
txtGrossScoreTot.Text = RETURNTEXTVALUE(txtGrossScoreF) + _
                        RETURNTEXTVALUE(txtGrossScoreB)
txtScoreGrossF.Text = txtGrossScoreF.Text
End Sub

Private Sub txtNetPts_Change(Index As Integer)
If Index >= 0 And Index <= 8 Then
    txtNetPtsF.Text = RETURNTEXTVALUE(txtNetPts(0)) + _
                          RETURNTEXTVALUE(txtNetPts(1)) + _
                          RETURNTEXTVALUE(txtNetPts(2)) + _
                          RETURNTEXTVALUE(txtNetPts(3)) + _
                          RETURNTEXTVALUE(txtNetPts(4)) + _
                          RETURNTEXTVALUE(txtNetPts(5)) + _
                          RETURNTEXTVALUE(txtNetPts(6)) + _
                          RETURNTEXTVALUE(txtNetPts(7)) + _
                          RETURNTEXTVALUE(txtNetPts(8))
ElseIf Index >= 9 And Index <= 17 Then
    txtNetPtsB.Text = RETURNTEXTVALUE(txtNetPts(9)) + _
                          RETURNTEXTVALUE(txtNetPts(10)) + _
                          RETURNTEXTVALUE(txtNetPts(11)) + _
                          RETURNTEXTVALUE(txtNetPts(12)) + _
                          RETURNTEXTVALUE(txtNetPts(13)) + _
                          RETURNTEXTVALUE(txtNetPts(14)) + _
                          RETURNTEXTVALUE(txtNetPts(15)) + _
                          RETURNTEXTVALUE(txtNetPts(16)) + _
                          RETURNTEXTVALUE(txtNetPts(17))
End If
End Sub

Private Sub txtNetPtsB_Change()
txtNetPtsTot.Text = RETURNTEXTVALUE(txtNetPtsF) + _
                    RETURNTEXTVALUE(txtNetPtsB)
'txtSNetB.Text = RETURNTEXTVALUE(txtNetPtsB)
End Sub

Private Sub txtNetPtsF_Change()
txtNetPtsTot.Text = RETURNTEXTVALUE(txtNetPtsF) + _
                    RETURNTEXTVALUE(txtNetPtsB)
'txtSNetF.Text = RETURNTEXTVALUE(txtNetPtsF)
End Sub

Private Sub txtScoreGrossB_Change()
txtSGrossTot.Text = RETURNTEXTVALUE(txtScoreGrossF) + RETURNTEXTVALUE(txtScoreGrossB)
End Sub

Private Sub txtScoreGrossF_Change()
txtSGrossTot.Text = RETURNTEXTVALUE(txtScoreGrossF) + RETURNTEXTVALUE(txtScoreGrossB)
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstResult.Clear: lstTeam.Clear: Exit Sub
Dim s As String
Dim rs As New ADODB.Recordset
lstResult.Clear: lstTeam.Clear
s = "SELECT tbl_Scoring_Team_Detail.PlayerKey, " & _
    " tbl_Scoring_PlayerName.LastName, " & _
    " tbl_Scoring_PlayerName.FirstName " & _
    " FROM tbl_Scoring_Team LEFT OUTER JOIN " & _
    " tbl_Scoring_Team_Detail ON tbl_Scoring_Team.PK = tbl_Scoring_Team_Detail.TeamKey LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_Team.TournamentKey = " & TournamentKey & ") " & _
    " GROUP BY tbl_Scoring_Team_Detail.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName " & _
    " HAVING (tbl_Scoring_PlayerName.LastName LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
    " ORDER BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResult.AddItem rs!LastName & ", " & rs!FirstName
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
If KeyCode = vbKeyReturn Then lstResult.SetFocus
End Sub

Private Sub txtSearchAdd_Change()
If Trim(txtSearchAdd.Text) = "" Then lstResultAdd.Clear: lstTeamAdd.Clear: Exit Sub
Dim s As String
Dim rs As New ADODB.Recordset
lstResultAdd.Clear: lstTeamAdd.Clear
's = "SELECT PK, LastName + ',  ' + FirstName + '  ' + MiddleName AS PlayerName " & _
    " From tbl_Scoring_PlayerName " & _
    " WHERE (LastName LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%') " & _
    " AND (TournamentKey = " & TournamentKey & ") " & _
    " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
s = "SELECT tbl_Scoring_Team_Detail.PlayerKey, " & _
    " tbl_Scoring_PlayerName.LastName, " & _
    " tbl_Scoring_PlayerName.FirstName " & _
    " FROM tbl_Scoring_Team LEFT OUTER JOIN " & _
    " tbl_Scoring_Team_Detail ON tbl_Scoring_Team.PK = tbl_Scoring_Team_Detail.TeamKey LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_Team.TournamentKey = " & TournamentKey & ") " & _
    " GROUP BY tbl_Scoring_Team_Detail.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName " & _
    " HAVING (tbl_Scoring_PlayerName.LastName LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%') " & _
    " ORDER BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultAdd.AddItem rs!LastName & ", " & rs!FirstName
    lstResultAdd.ItemData(lstResultAdd.NewIndex) = rs!PlayerKey
    rs.MoveNext
Wend
rs.Close
If lstResultAdd.ListCount Then lstResultAdd.ListIndex = 0
End Sub

Private Sub txtSearchAdd_GotFocus()
HTEXT txtSearchAdd
End Sub

Private Sub txtSearchAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultAdd.SetFocus
End Sub

Private Sub txtSGrossB_Change()
'txtSGrossTot.Text = RETURNTEXTVALUE(txtSGrossF) + _
                    RETURNTEXTVALUE(txtSGrossB)
End Sub

Private Sub txtSGrossF_Change()
'txtSGrossTot.Text = RETURNTEXTVALUE(txtSGrossF) + _
                    RETURNTEXTVALUE(txtSGrossB)
End Sub

Private Sub txtSNetB_Change()
'txtSNetTot.Text = RETURNTEXTVALUE(txtSNetF) + _
                  RETURNTEXTVALUE(txtSNetB)
End Sub

Private Sub txtSNetF_Change()
'txtSNetTot.Text = RETURNTEXTVALUE(txtSNetF) + _
                  RETURNTEXTVALUE(txtSNetB)
End Sub

Private Sub txtSGrossTot_Change()
txtNetScore = RETURNTEXTVALUE(txtSGrossTot) - RETURNTEXTVALUE(txtTeamHandicap)
End Sub

Private Sub txtTeamHandicap_Change()
txtNetScore = RETURNTEXTVALUE(txtSGrossTot) - RETURNTEXTVALUE(txtTeamHandicap)
End Sub
