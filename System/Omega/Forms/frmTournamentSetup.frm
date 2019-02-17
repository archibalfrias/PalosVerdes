VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTournamentSetup 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13245
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTournamentSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   13245
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picSLine 
      Height          =   855
      Left            =   4800
      TabIndex        =   57
      Top             =   3720
      Visible         =   0   'False
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.ComboBox cmbLocation 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   360
         Width           =   7275
      End
      Begin VB.TextBox txtType 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtItemKey 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtItemKey1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Location Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   120
         Width           =   7215
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   65
      Top             =   0
      Width           =   15000
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   12360
         ScaleHeight     =   495
         ScaleWidth      =   615
         TabIndex        =   67
         Top             =   240
         Width           =   615
         Begin VB.Image imgLocked 
            Height          =   465
            Left            =   120
            Picture         =   "frmTournamentSetup.frx":08CA
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   600
         End
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   66
         Top             =   105
         Width           =   12240
         _ExtentX        =   21590
         _ExtentY        =   1429
         ButtonWidth     =   1323
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   28
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
               Caption         =   "Activate"
               Key             =   "Set"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Export"
               Key             =   "Export"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Lock"
               Key             =   "Lock"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               Key             =   "Refresh"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         MousePointer    =   99
         MouseIcon       =   "frmTournamentSetup.frx":1594
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
      Left            =   -240
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTournamentSetup.frx":18AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTournamentSetup.frx":2588
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTournamentSetup.frx":3262
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTournamentSetup.frx":3F3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTournamentSetup.frx":4C16
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTournamentSetup.frx":58F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTournamentSetup.frx":65CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTournamentSetup.frx":72A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTournamentSetup.frx":7F7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTournamentSetup.frx":8858
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTournamentSetup.frx":9532
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTournamentSetup.frx":A20C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTournamentSetup.frx":AEE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTournamentSetup.frx":BBC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTournamentSetup.frx":C89A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTournamentSetup.frx":D574
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   6690
      Width           =   13245
      _ExtentX        =   23363
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
   Begin RPVGCC.b8Container picPrint 
      Height          =   2175
      Left            =   4560
      TabIndex        =   51
      Top             =   2400
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3836
      BackColor       =   15396057
      Begin VB.Timer TimerScoreCard1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   3600
         Top             =   960
      End
      Begin VB.CheckBox chkWithNetPts 
         BackColor       =   &H00EAECD9&
         Caption         =   "with Net Points"
         Height          =   255
         Left            =   1200
         TabIndex        =   56
         Top             =   960
         Width           =   2415
      End
      Begin VB.CheckBox chkWithGrossPts 
         BackColor       =   &H00EAECD9&
         Caption         =   "with Gross Points"
         Height          =   255
         Left            =   1200
         TabIndex        =   55
         Top             =   600
         Width           =   2415
      End
      Begin VB.Timer TimerScoreCard 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   3600
         Top             =   480
      End
      Begin VB.CommandButton cmdCancelPrint 
         Height          =   480
         Left            =   2235
         Picture         =   "frmTournamentSetup.frx":E24E
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   1395
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKPrint 
         Height          =   480
         Left            =   555
         Picture         =   "frmTournamentSetup.frx":E9AA
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   1395
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar3 
         Height          =   345
         Left            =   40
         TabIndex        =   54
         Top             =   40
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   609
         Caption         =   "Print Score Card"
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
         Icon            =   "frmTournamentSetup.frx":F01C
      End
   End
   Begin RPVGCC.b8Container picProgress 
      Height          =   975
      Left            =   3840
      TabIndex        =   49
      Top             =   3000
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1720
      BackColor       =   15396057
      Begin VB.PictureBox picProgressBar 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5235
         TabIndex        =   50
         Top             =   120
         Width           =   5295
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   600
      ScaleHeight     =   5295
      ScaleWidth      =   12015
      TabIndex        =   5
      Top             =   1200
      Width           =   12015
      Begin VB.Frame Frame6 
         BackColor       =   &H00C6B8A4&
         Caption         =   "      PARTNER PLAY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3960
         TabIndex        =   38
         Top             =   1920
         Width           =   3975
         Begin VB.PictureBox picPartner 
            BackColor       =   &H00C6B8A4&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   120
            ScaleHeight     =   375
            ScaleWidth      =   255
            TabIndex        =   45
            Top             =   0
            Width           =   255
            Begin VB.CheckBox chkPartner 
               BackColor       =   &H00C6B8A4&
               Height          =   255
               Left            =   0
               TabIndex        =   46
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.TextBox txtAllowPartner 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2760
            TabIndex        =   40
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox cmbPointsToCountPartner 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Allowed Partner Per Player"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Points to Count"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   720
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C6B8A4&
         Caption         =   "Location/s"
         Height          =   2055
         Left            =   3960
         TabIndex        =   36
         Top             =   3240
         Width           =   8055
         Begin MSComctlLib.ListView lstLocation 
            Height          =   1695
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   2990
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "LocationKey"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Location"
               Object.Width           =   12877
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "HomeCourt"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.TextBox txtParGrossPts 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6840
         TabIndex        =   34
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C6B8A4&
         Caption         =   "      INDIVIDUAL PLAY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   8040
         TabIndex        =   27
         Top             =   1920
         Width           =   3975
         Begin VB.PictureBox picIndividual 
            BackColor       =   &H00C6B8A4&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   120
            ScaleHeight     =   375
            ScaleWidth      =   255
            TabIndex        =   47
            Top             =   0
            Width           =   255
            Begin VB.CheckBox chkIndividualPlay 
               BackColor       =   &H00C6B8A4&
               Height          =   255
               Left            =   0
               TabIndex        =   48
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.ComboBox cmbPointsToCountIndi 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   720
            Width           =   2175
         End
         Begin VB.TextBox txtAllowTeam 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2760
            TabIndex        =   28
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Points to Count"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Allowed Team Per Player"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C6B8A4&
         Caption         =   "      TEAM PLAY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   8040
         TabIndex        =   20
         Top             =   0
         Width           =   3975
         Begin VB.ComboBox cmbOrder 
            Height          =   315
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   600
            Width           =   975
         End
         Begin VB.PictureBox picTeam 
            BackColor       =   &H00C6B8A4&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   120
            ScaleHeight     =   375
            ScaleWidth      =   255
            TabIndex        =   43
            Top             =   0
            Width           =   255
            Begin VB.CheckBox chkTeamPlay 
               BackColor       =   &H00C6B8A4&
               Height          =   255
               Left            =   0
               TabIndex        =   44
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.ComboBox cmbTeamAverage 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1320
            Width           =   2175
         End
         Begin VB.TextBox txtPlayerToCount 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1680
            TabIndex        =   23
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtHDCPDivisor 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1680
            TabIndex        =   22
            Top             =   600
            Width           =   1095
         End
         Begin VB.ComboBox cmbPointsToCount 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Order"
            Height          =   255
            Left            =   2880
            TabIndex        =   64
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Team Average"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Player to Count"
            Height          =   255
            Left            =   480
            TabIndex        =   26
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Divisor"
            Height          =   255
            Left            =   480
            TabIndex        =   25
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Points to Count"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   960
            Width           =   1695
         End
      End
      Begin VB.PictureBox picIndexClass 
         BackColor       =   &H00C6B8A4&
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   0
         ScaleHeight     =   2295
         ScaleWidth      =   3855
         TabIndex        =   17
         Top             =   3120
         Width           =   3855
         Begin VB.Frame Frame2 
            BackColor       =   &H00C6B8A4&
            Caption         =   "Index Classification"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2175
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   3855
            Begin MSFlexGridLib.MSFlexGrid FGridIndex 
               Height          =   1785
               Left            =   120
               TabIndex        =   19
               Top             =   240
               Width           =   3585
               _ExtentX        =   6324
               _ExtentY        =   3149
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
         End
      End
      Begin VB.ComboBox cmbScoring 
         Height          =   315
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtNoPlayer 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6840
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.PictureBox picClass 
         BackColor       =   &H00C6B8A4&
         BorderStyle     =   0  'None
         Height          =   2175
         Left            =   0
         ScaleHeight     =   2175
         ScaleWidth      =   3855
         TabIndex        =   10
         Top             =   840
         Width           =   3855
         Begin VB.Frame Frame1 
            BackColor       =   &H00C6B8A4&
            Caption         =   "Handicap Classification"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2175
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   3855
            Begin MSFlexGridLib.MSFlexGrid FGrid 
               Height          =   1785
               Left            =   120
               TabIndex        =   12
               Top             =   240
               Width           =   3585
               _ExtentX        =   6324
               _ExtentY        =   3149
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
         End
      End
      Begin VB.TextBox txtNoDays 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6840
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtTo 
         Height          =   315
         Left            =   2760
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtFrom 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Text            =   "05/20/2010"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   0
         Width           =   6495
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Par Gross Points"
         Height          =   255
         Left            =   4080
         TabIndex        =   35
         Top             =   1500
         Width           =   2055
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Scoring System"
         Height          =   255
         Left            =   4080
         TabIndex        =   16
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Player Per Team"
         Height          =   255
         Left            =   4080
         TabIndex        =   14
         Top             =   720
         Width           =   2055
      End
      Begin VB.Image imgLogo 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3480
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Days Player Can Play / Location"
         Height          =   255
         Left            =   4080
         TabIndex        =   8
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Duration"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Event Name"
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmTournamentSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iLocationKey As Long

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim Focus_Class         As Long
Dim Focus_ClassIndex    As Long
Dim PressCount          As Long
Dim iRow                As Long
Dim iFocus              As Long

Dim iSet                As Long

Public sFileName        As String
Dim WorkbookName        As String
Dim iWorkSheet          As Integer

Dim tmp As Long

Dim RowCnt, ColCnt, strRange, i, j, l, k, Arr, x, iDay, _
dDate, sTotal, RowCntTmp, ColCntTmp, strRangeFrom, _
strRangeTo, ArrDate, iDateDiff, HeaderRow, iProgressValue, _
dTeamPoints, sMinRange, sMaxRange, iScoreGrossNet, TourKey, _
HEADER1$, LTEXT, sTournamentRange, sValue, iTotProgressValue


Private Function BROWSER(strEvent, is_Action As String)
Select Case is_Action
    Case "is_LOAD"
        If Trim(strEvent) <> "" Then
            s = "SELECT TOP 1 tbl_Scoring_TournamentInfo.* " & _
                " FROM tbl_Scoring_TournamentInfo " & _
                " WHERE (PK = " & strEvent & ")" & _
                " ORDER BY PK "
        Else
            s = "SELECT TOP 1 tbl_Scoring_TournamentInfo.* " & _
                " FROM tbl_Scoring_TournamentInfo " & _
                " ORDER BY PK "
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Scoring_TournamentInfo.* " & _
            " FROM tbl_Scoring_TournamentInfo " & _
            " ORDER BY PK "
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Scoring_TournamentInfo.* " & _
            " FROM tbl_Scoring_TournamentInfo " & _
            " WHERE (PK < " & strEvent & ")" & _
            " ORDER BY PK DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Scoring_TournamentInfo.* " & _
            " FROM tbl_Scoring_TournamentInfo " & _
            " WHERE (PK > " & strEvent & ")" & _
            " ORDER BY PK "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Scoring_TournamentInfo.* " & _
            " FROM tbl_Scoring_TournamentInfo " & _
            " ORDER BY PK DESC"
    Case Else:      Exit Function
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtName.Text = rs!TournamentName
    txtFrom.Text = Format(rs!TournamentStart, "mm/dd/yyyy")
    txtTo.Text = Format(rs!TournamentEnd, "mm/dd/yyyy")
    txtNoDays.Text = rs!NoofPlays
    txtNoPlayer.Text = rs!NoofPlayerPerTeam
    txtPlayerToCount.Text = rs!PlayerToCount
    txtAllowTeam.Text = rs!AllowTeamPerPlayer
    txtHDCPDivisor.Text = rs!HandicapDivisor
    cmbOrder.ListIndex = rs!TeamDivisorOrder
    txtAllowPartner.Text = rs!AllowPartnerPerPlayer
    chkTeamPlay.Value = rs!TeamPlay
    chkIndividualPlay.Value = rs!IndividualPlay
    chkPartner.Value = rs!PartnerPlay
    cmbScoring.ListIndex = rs!Scoring
    cmbPointsToCount.ListIndex = rs!PointsToCountTeam
    cmbPointsToCountIndi.ListIndex = rs!PointsToCountIndi
    cmbPointsToCountPartner.ListIndex = rs!PointsToCountPartner
    cmbTeamAverage.ListIndex = rs!TeamAverage
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", "Last Modified : " & rs!LastModified)
    imgLocked.Visible = IIf(CDbl(rs!Locked) = 1, True, False)
    iSet = rs!Activated
    Toolbar1.Buttons(19).Enabled = IIf(CDbl(rs!Activated) = 1, False, True)
    txtParGrossPts.Text = rs!ParGrossPoints
    
    imgLogo.Picture = LoadPicture("")
    CUSTOM_GRID
    i = 0
    t = "SELECT tbl_Scoring_TournamentInfo_Class.* " & _
        " FROM tbl_Scoring_TournamentInfo_Class " & _
        " WHERE (TournamentKey = " & rs!PK & ") " & _
        " ORDER BY Class"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        With FGrid
            .Rows = rt.RecordCount + 1
            While Not rt.EOF
                i = i + 1
                .TextMatrix(i, 1) = rt!Class
                .TextMatrix(i, 2) = rt!HFrom
                .TextMatrix(i, 3) = rt!HTo
                rt.MoveNext
            Wend
        End With
    End If
    rt.Close
    
    i = 0
    t = "SELECT tbl_Scoring_TournamentInfo_Index.* " & _
        " FROM tbl_Scoring_TournamentInfo_Index " & _
        " WHERE (TournamentKey = " & rs!PK & ") " & _
        " ORDER BY Class"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        With FGridIndex
            .Rows = rt.RecordCount + 1
            While Not rt.EOF
                i = i + 1
                .TextMatrix(i, 1) = rt!Class
                .TextMatrix(i, 2) = rt!HFrom
                .TextMatrix(i, 3) = rt!HTo
                rt.MoveNext
            Wend
        End With
    End If
    rt.Close
    
    i = 0
    lstLocation.ListItems.Clear
    t = "SELECT dbo.tbl_Scoring_TournamentInfo_Location.LocationKey, " & _
        " dbo.tbl_Scoring_Location.ScoringLocation, " & _
        " dbo.tbl_Scoring_TournamentInfo_Location.HomeCourt " & _
        " FROM dbo.tbl_Scoring_TournamentInfo_Location LEFT OUTER JOIN " & _
        " dbo.tbl_Scoring_Location ON dbo.tbl_Scoring_TournamentInfo_Location.LocationKey = dbo.tbl_Scoring_Location.PK " & _
        " Where (dbo.tbl_Scoring_TournamentInfo_Location.MasterKey = " & rs!PK & ") " & _
        " ORDER BY dbo.tbl_Scoring_Location.ScoringLocation"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        i = i + 1
        Set x = lstLocation.ListItems.Add()
        x.Text = ""
        x.SubItems(1) = rt!LocationKey
        x.SubItems(2) = rt!ScoringLocation
        x.SubItems(3) = rt!HomeCourt
        rt.MoveNext
    Wend
    rt.Close
    
    SaveSetting App.EXEName, "EventName", "NameEvent", rs!PK
End If
rs.Close
End Function

Private Function PRESS_INSERT()

If TRANSACTIONTYPE = is_REFRESH Then
    If AccessRights("Scoring Tournament Information", "Add") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Function
    End If
    CLEARTEXT
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_ADDING
    txtName.SetFocus
ElseIf TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If Focus_Class = 1 Then
        With FGrid
            If .TextMatrix(.Rows - 1, 1) <> "" Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = ""
                .TextMatrix(.Rows - 1, 2) = ""
                .TextMatrix(.Rows - 1, 3) = ""
                .TopRow = .Rows - 1
                .ROW = .Rows - 1
                .Col = 1
            End If
        End With
        Exit Function
    End If
    If Focus_ClassIndex = 1 Then
        With FGridIndex
            If .TextMatrix(.Rows - 1, 1) <> "" Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = ""
                .TextMatrix(.Rows - 1, 2) = ""
                .TextMatrix(.Rows - 1, 3) = ""
                .TopRow = .Rows - 1
                .ROW = .Rows - 1
                .Col = 1
            End If
        End With
        Exit Function
    End If
    If iFocus = 1 Then
        With lstLocation.ListItems
            Set x = .Add()
            x.Text = ""
            x.SubItems(1) = "0"
            x.SubItems(2) = " "
            x.SubItems(3) = "0"
            iRow = .Count
            picSLine.ZOrder 0
            picMain.Enabled = False
            cmbLocation.ListIndex = -1
            picSLine.Visible = True
            cmbLocation.SetFocus
        End With
    End If
End If
End Function

Private Function PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If Statusbar1.Panels(1).Text = "" Then Exit Function
If AccessRights("Scoring Tournament Information", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
If imgLocked.Visible = True Then MsgBox "Event Already Locked!                   ", vbCritical, "Error...": Exit Function
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_EDITTING
End Function

Private Function PRESS_DELETE()
If TRANSACTIONTYPE = is_REFRESH Then
    If Statusbar1.Panels(1).Text = "" Then Exit Function
    If AccessRights("Scoring Tournament Information", "Delete") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Function
    End If
    If imgLocked.Visible = True Then MsgBox "Event Already Locked!                   ", vbCritical, "Error...": Exit Function
    If MsgBox("ARE YOU SURE IN DELETING THIS EVENT?                 ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Function
    On Error GoTo PG:
    ConnOmega.Execute "DELETE FROM tbl_Scoring_TournamentInfo WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    CLEARTEXT
    BROWSER GetSetting(App.EXEName, "EventName", "NameEvent", ""), "is_PAGEDOWN"
    If Trim(txtName.Text) = "" Then BROWSER GetSetting(App.EXEName, "EventName", "NameEvent", ""), "is_HOME"
ElseIf TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If Focus_Class = 1 Then
        With FGrid
            If .Rows = 2 Then
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
                .TextMatrix(1, 3) = ""
            Else
                .RemoveItem .ROW
            End If
        End With
        Exit Function
    End If
    If Focus_ClassIndex = 1 Then
        With FGridIndex
            If .Rows = 2 Then
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
                .TextMatrix(1, 3) = ""
            Else
                .RemoveItem .ROW
            End If
        End With
        Exit Function
    End If
    If iFocus = 1 Then
        With lstLocation.ListItems
            If CDbl(.Item(iRow).SubItems(3)) = 0 Then
                .Remove iRow
            End If
        End With
    End If
End If
Exit Function
PG:
MsgBox Err.Number & Err.Description, vbCritical, "Error..."
Exit Function
End Function

Private Function PRESS_F5()
If Trim(txtName.Text) = "" Then MsgBox "Please Supply Event Name!                 ", vbCritical, "Error...": txtName.SetFocus: Exit Function
If IsDate(txtFrom.Text) = False Then MsgBox "Please Supply Date Start!                ", vbCritical, "Error...": txtFrom.SetFocus: Exit Function
If IsDate(txtTo.Text) = False Then MsgBox "Please Supply Date Finish!                 ", vbCritical, "Error...": txtTo.SetFocus: Exit Function
If RETURNTEXTVALUE(txtNoDays) <= 0 Then MsgBox "Please Supply Number of Plays!                ", vbCritical, "Error...": txtNoDays.SetFocus: Exit Function
If RETURNTEXTVALUE(txtNoPlayer) <= 0 Then MsgBox "Please Supply Number of Player Per Team!                        ", vbCritical, "Error...": txtNoPlayer.SetFocus: Exit Function
'if cmbPointsToCount.ListIndex = 0 then msgbox
If RETURNTEXTVALUE(txtPlayerToCount) <= 0 Then MsgBox "Please Supply Team Player to Count!                    ", vbCritical, "Error...": txtPlayerToCount.SetFocus: HTEXT txtPlayerToCount: Exit Function
If cmbScoring.ListIndex = 0 Then MsgBox "Please Select Scoring!                         ", vbCritical, "Error...": cmbScoring.SetFocus: Exit Function
If cmbOrder.ListIndex = -1 Then MsgBox "Please Select Order!                                 ", vbCritical, "Error...": cmbOrder.SetFocus: Exit Function
If cmbScoring.ListIndex <> 3 Then
    If RETURNTEXTVALUE(txtHDCPDivisor) <= 0 Then MsgBox "Please Supply Handicap Divisor!                    ", vbCritical, "Error...": txtHDCPDivisor.SetFocus: HTEXT txtHDCPDivisor: Exit Function
End If
If cmbScoring.ListIndex = 4 Then MsgBox "This Scoring System is not yet Activated!                      ", vbCritical, "Error...": cmbScoring.SetFocus: Exit Function

If RETURNTEXTVALUE(txtParGrossPts) <= 0 Then MsgBox "Invalid Gross Points!                               ", vbCritical, "Error...": txtParGrossPts.SetFocus: Exit Function

On Error GoTo PG:
TourKey = 0
If TRANSACTIONTYPE = is_ADDING Then
    ConnOmega.Execute "INSERT INTO tbl_Scoring_TournamentInfo " & _
                      " (TournamentName, TournamentStart, " & _
                      " TournamentEnd, NoofPlays, NoofPlayerPerTeam, " & _
                      " LastModified, TeamPlay, IndividualPlay, " & _
                      " PlayerToCount, AllowTeamPerPlayer, HandicapDivisor, " & _
                      " Scoring, PointsToCountTeam, PointsToCountIndi, TeamAverage, " & _
                      " ParGrossPoints, PartnerPlay, AllowPartnerPerPlayer, PointsToCountPartner, TeamDivisorOrder) " & _
                      " VALUES ('" & FORMATSQL(Trim(txtName.Text)) & "', " & _
                      " '" & FormatDateTime(txtFrom.Text, vbShortDate) & "', " & _
                      " '" & FormatDateTime(txtTo.Text, vbShortDate) & "', " & _
                      " " & RETURNTEXTVALUE(txtNoDays) & ", " & _
                      " " & RETURNTEXTVALUE(txtNoPlayer) & ", " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                      " " & chkTeamPlay.Value & ", " & _
                      " " & chkIndividualPlay.Value & ", " & _
                      " " & RETURNTEXTVALUE(txtPlayerToCount) & ", " & _
                      " " & RETURNTEXTVALUE(txtAllowTeam) & ", " & _
                      " " & RETURNTEXTVALUE(txtHDCPDivisor) & ", " & _
                      " " & cmbScoring.ListIndex & ", " & _
                      " " & cmbPointsToCount.ListIndex & ", " & _
                      " " & cmbPointsToCountIndi.ListIndex & ", " & _
                      " " & cmbTeamAverage.ListIndex & ", " & _
                      " " & RETURNTEXTVALUE(txtParGrossPts) & ", " & _
                      " " & chkPartner.Value & ", " & RETURNTEXTVALUE(txtAllowPartner) & ", " & cmbPointsToCountPartner.ListIndex & " , " & cmbOrder.ListIndex & ")"
    s = "SELECT TOP 1 PK " & _
        " FROM tbl_Scoring_TournamentInfo " & _
        " ORDER BY PK DESC "
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        TourKey = rs!PK
    End If
    rs.Close
    If CDbl(TourKey) <> 0 Then
        With FGrid
            For i = 1 To (.Rows - 1)
                If Trim(.TextMatrix(i, 1)) <> "" Then
                    ConnOmega.Execute "INSERT INTO tbl_Scoring_TournamentInfo_Class " & _
                                      " (TournamentKey, Class, HFrom, HTo) " & _
                                      " VALUES (" & TourKey & ", '" & FORMATSQL(Trim(.TextMatrix(i, 1))) & "', " & _
                                      " " & CDbl(.TextMatrix(i, 2)) & ", " & CDbl(.TextMatrix(i, 3)) & ")"
                End If
            Next i
        End With
        With FGridIndex
            For i = 1 To (.Rows - 1)
                If Trim(.TextMatrix(i, 1)) <> "" Then
                    ConnOmega.Execute "INSERT INTO tbl_Scoring_TournamentInfo_Index " & _
                                      " (TournamentKey, Class, HFrom, HTo) " & _
                                      " VALUES (" & TourKey & ", '" & FORMATSQL(Trim(.TextMatrix(i, 1))) & "', " & _
                                      " " & CDbl(.TextMatrix(i, 2)) & ", " & CDbl(.TextMatrix(i, 3)) & ")"
                End If
            Next i
        End With
        
        s = "SELECT tbl_Scoring_Location.* " & _
            " FROM tbl_Scoring_Location " & _
            " WHERE (DefaultLocation = 1)"
        If rs.State = adStateOpen Then rt.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Scoring_TournamentInfo_Location " & _
                              " (MasterKey, LocationKey, HomeCourt) " & _
                              " VALUES (" & TourKey & ", " & rs!PK & ", 1)"
        End If
        rs.Close
        
        With lstLocation.ListItems
            For i = 1 To .Count
                s = "SELECT tbl_Scoring_TournamentInfo_Location.* " & _
                    " FROM tbl_Scoring_TournamentInfo_Location " & _
                    " WHERE (MasterKey =  " & TourKey & ") " & _
                    " AND (LocationKey = " & .Item(i).SubItems(1) & ")"
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                If rs.RecordCount = 0 Then
                    ConnOmega.Execute "INSERT INTO tbl_Scoring_TournamentInfo_Location " & _
                              " (MasterKey, LocationKey) " & _
                              " VALUES (" & TourKey & ", " & .Item(i).SubItems(1) & ")"
                End If
                rs.Close
            Next i
        End With
    End If
        
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER TourKey, "is_LOAD"
    
ElseIf TRANSACTIONTYPE = is_EDITTING Then
    ConnOmega.Execute "UPDATE tbl_Scoring_TournamentInfo " & _
                      " SET TournamentName = '" & FORMATSQL(Trim(txtName.Text)) & "', " & _
                      " TournamentStart = '" & FormatDateTime(txtFrom.Text, vbShortDate) & "', " & _
                      " TournamentEnd = '" & FormatDateTime(txtTo.Text, vbShortDate) & "', " & _
                      " NoofPlays = " & RETURNTEXTVALUE(txtNoDays) & ", " & _
                      " NoofPlayerPerTeam = " & RETURNTEXTVALUE(txtNoPlayer) & ", " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                      " TeamPlay = " & chkTeamPlay.Value & ", " & _
                      " IndividualPlay = " & chkIndividualPlay.Value & ", " & _
                      " PlayerToCount = " & RETURNTEXTVALUE(txtPlayerToCount) & ", " & _
                      " AllowTeamPerPlayer = " & RETURNTEXTVALUE(txtAllowTeam) & ", " & _
                      " HandicapDivisor = " & RETURNTEXTVALUE(txtHDCPDivisor) & ", " & _
                      " Scoring = " & cmbScoring.ListIndex & ", " & _
                      " PointsToCountTeam = " & cmbPointsToCount.ListIndex & ", " & _
                      " PointsToCountIndi = " & cmbPointsToCountIndi.ListIndex & ", " & _
                      " TeamAverage = " & cmbTeamAverage.ListIndex & ", " & _
                      " ParGrossPoints = " & RETURNTEXTVALUE(txtParGrossPts) & ", " & _
                      " PartnerPlay = " & chkPartner.Value & ", " & _
                      " AllowPartnerPerPlayer = " & RETURNTEXTVALUE(txtAllowPartner) & ", " & _
                      " PointsToCountPartner = " & cmbPointsToCountPartner.ListIndex & ", " & _
                      " TeamDivisorOrder = " & cmbOrder.ListIndex & " " & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    TourKey = Statusbar1.Panels(1).Text
    ConnOmega.Execute "DELETE FROM tbl_Scoring_TournamentInfo_Class WHERE (TournamentKey = " & TourKey & ")"
    ConnOmega.Execute "DELETE FROM tbl_Scoring_TournamentInfo_Index WHERE (TournamentKey = " & TourKey & ")"
    ConnOmega.Execute "DELETE FROM tbl_Scoring_TournamentInfo_Location WHERE (MasterKey = " & TourKey & ")"
    If CDbl(TourKey) <> 0 Then
        With FGrid
            For i = 1 To (.Rows - 1)
                If Trim(.TextMatrix(i, 1)) <> "" Then
                    ConnOmega.Execute "INSERT INTO tbl_Scoring_TournamentInfo_Class " & _
                                      " (TournamentKey, Class, HFrom, HTo) " & _
                                      " VALUES (" & TourKey & ", '" & FORMATSQL(Trim(.TextMatrix(i, 1))) & "', " & _
                                      " " & CDbl(.TextMatrix(i, 2)) & ", " & CDbl(.TextMatrix(i, 3)) & ")"
                End If
            Next i
        End With
        
        With FGridIndex
            For i = 1 To (.Rows - 1)
                If Trim(.TextMatrix(i, 1)) <> "" Then
                    ConnOmega.Execute "INSERT INTO tbl_Scoring_TournamentInfo_Index " & _
                                      " (TournamentKey, Class, HFrom, HTo) " & _
                                      " VALUES (" & TourKey & ", '" & FORMATSQL(Trim(.TextMatrix(i, 1))) & "', " & _
                                      " " & CDbl(.TextMatrix(i, 2)) & ", " & CDbl(.TextMatrix(i, 3)) & ")"
                End If
            Next i
        End With
        
        s = "SELECT tbl_Scoring_Location.* " & _
            " FROM tbl_Scoring_Location " & _
            " WHERE (DefaultLocation = 1)"
        If rs.State = adStateOpen Then rt.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            ConnOmega.Execute "INSERT INTO tbl_Scoring_TournamentInfo_Location " & _
                              " (MasterKey, LocationKey, HomeCourt) " & _
                              " VALUES (" & TourKey & ", " & rs!PK & ", 1)"
        End If
        rs.Close
        
        With lstLocation.ListItems
            For i = 1 To .Count
                s = "SELECT tbl_Scoring_TournamentInfo_Location.* " & _
                    " FROM tbl_Scoring_TournamentInfo_Location " & _
                    " WHERE (MasterKey =  " & TourKey & ") " & _
                    " AND (LocationKey = " & .Item(i).SubItems(1) & ")"
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                If rs.RecordCount = 0 Then
                    ConnOmega.Execute "INSERT INTO tbl_Scoring_TournamentInfo_Location " & _
                              " (MasterKey, LocationKey) " & _
                              " VALUES (" & TourKey & ", " & .Item(i).SubItems(1) & ")"
                End If
                rs.Close
            Next i
        End With
    End If
    
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER TourKey, "is_LOAD"
    
End If
Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Function
End Function

Private Function PRESS_F7()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If Statusbar1.Panels(1).Text = "" Then Exit Function
If MsgBox("ACTIVATE THIS EVENT!                 ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Function
s = "SELECT tbl_Scoring_TournamentInfo.* " & _
    " FROM tbl_Scoring_TournamentInfo"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    ConnOmega.Execute "UPDATE tbl_Scoring_TournamentInfo " & _
                      " SET Activated = 0 " & _
                      " WHERE (PK = " & rs!PK & ")"
    rs.MoveNext
Wend
rs.Close

ConnOmega.Execute "UPDATE tbl_Scoring_TournamentInfo " & _
                  " SET Activated = 1 " & _
                  " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"

s = "SELECT tbl_Scoring_TournamentInfo.* " & _
    " FROM tbl_Scoring_TournamentInfo " & _
    " WHERE (Activated = 1)"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    
    TournamentKey = rs!PK
    WithTeamPlay = rs!TeamPlay
    WithIndividualPlay = rs!IndividualPlay
    TournamentName = rs!TournamentName
    TournamentRange = Format(rs!TournamentStart, "mm/dd/yyyy") & " - " & Format(rs!TournamentEnd, "mm/dd/yyyy")
    TeamPlayer2Cnt = rs!PlayerToCount
    AllowedTeam = rs!AllowTeamPerPlayer
    NoofPlayerPerTeam = rs!NoofPlayerPerTeam
    HandicapDivisor = rs!HandicapDivisor
    DaysPlayerToPlay = rs!NoofPlays
    ScoringType = rs!Scoring
    PointsToCnt = rs!PointsToCountTeam
    PointsToCntIndi = rs!PointsToCountIndi
    TeamAverage = rs!TeamAverage
    ParGrossPoints = rs!ParGrossPoints
    TeamDivisorOrder = rs!TeamDivisorOrder
    
    t = "SELECT TOP 1 HTo " & _
        " From tbl_Scoring_TournamentInfo_Class " & _
        " Where (TournamentKey = " & TournamentKey & ") " & _
        " ORDER BY HTo DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        TopHandicap = CDbl(rt!HTo)
    End If
    rt.Close
    
    t = "SELECT TOP 1 HTo " & _
        " From tbl_Scoring_TournamentInfo_Index " & _
        " Where (TournamentKey = " & TournamentKey & ") " & _
        " ORDER BY HTo DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        TopIndex = CDbl(rt!HTo)
    End If
    rt.Close
    
    t = "SELECT COUNT(*) AS LocCnt " & _
        " From dbo.tbl_Scoring_TournamentInfo_Location " & _
        " WHERE (MasterKey = " & TournamentKey & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        LocationCnt = CDbl(IIf(IsNull(rt!LocCnt), 0, rt!LocCnt))
    End If
    rt.Close
    
End If
rs.Close
            
BROWSER GetSetting(App.EXEName, "EventName", "NameEvent", ""), "is_LOAD"

End Function

Private Function PRESS_F8()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If Statusbar1.Panels(1).Text = "" Then Exit Function
If imgLocked.Visible = True Then MsgBox "Event Already Locked!                   ", vbCritical, "Error...": Exit Function
If MsgBox("CONTINUE LOCKING THIS EVENT!                 ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Function
On Error GoTo PG:
ConnOmega.Execute "UPDATE tbl_Scoring_TournamentInfo " & _
                  " SET Locked = 1 " & _
                  " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
                  
BROWSER GetSetting(App.EXEName, "EventName", "NameEvent", ""), "is_LOAD"

Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Function
End Function

Private Sub PRESS_F9()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub

PopupMenu MainFormPopupF.mnuTournamentInfoPrint, , Toolbar1.Buttons(17).Left, Toolbar1.Buttons(17).Top + Toolbar1.Buttons(17).Height

End Sub

Private Sub PRESS_F3()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If imgLocked.Visible = True Then MsgBox "Event Already Locked!                   ", vbCritical, "Error...": Exit Sub
If CLng(iSet) = 0 Then MsgBox "Please Activate this Tournament!                     ", vbCritical, "Error...": Exit Sub
For i = 1 To MainFormPopupF.mnuScoringLocationName.UBound
    Unload MainFormPopupF.mnuScoringLocationName(i)
Next i
i = -1
t = "SELECT dbo.tbl_Scoring_TournamentInfo_Location.LocationKey, " & _
    " dbo.tbl_Scoring_Location.ScoringLocation " & _
    " FROM dbo.tbl_Scoring_TournamentInfo_Location LEFT OUTER JOIN " & _
    " dbo.tbl_Scoring_Location ON dbo.tbl_Scoring_TournamentInfo_Location.LocationKey = dbo.tbl_Scoring_Location.PK " & _
    " Where (dbo.tbl_Scoring_TournamentInfo_Location.MasterKey = " & TournamentKey & ") " & _
    " And (dbo.tbl_Scoring_TournamentInfo_Location.HomeCourt = 0) " & _
    " ORDER BY dbo.tbl_Scoring_Location.ScoringLocation"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
While Not rt.EOF
    i = i + 1
    If i = 0 Then
        MainFormPopupF.mnuScoringLocationName(i).Caption = rt!ScoringLocation
    Else
        Load MainFormPopupF.mnuScoringLocationName(i)
        MainFormPopupF.mnuScoringLocationName(i).Caption = rt!ScoringLocation
    End If
    rt.MoveNext
Wend
rt.Close
PopupMenu MainFormPopupF.mnuScoringLocation, , Toolbar1.Buttons(21).Left, Toolbar1.Buttons(21).Top + Toolbar1.Buttons(21).Height
End Sub

Private Function PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    Unload Me
Else
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "EventName", "NameEvent", ""), "is_LOAD"
End If
End Function

Private Function LOCKTEXT(bln As Boolean)
txtName.Locked = bln
txtFrom.Locked = bln
txtTo.Locked = bln
txtNoDays.Locked = bln
txtNoPlayer.Locked = bln
txtPlayerToCount.Locked = bln
txtAllowTeam.Locked = bln
txtAllowPartner.Locked = bln
txtHDCPDivisor.Locked = bln
cmbScoring.Locked = bln
cmbPointsToCountIndi.Locked = bln
cmbPointsToCountPartner.Locked = bln
cmbPointsToCount.Locked = bln
cmbTeamAverage.Locked = bln
txtParGrossPts.Locked = True
cmbOrder.Locked = bln
picTeam.Enabled = IIf(bln = True, False, True)
picPartner.Enabled = IIf(bln = True, False, True)
picIndividual.Enabled = IIf(bln = True, False, True)

End Function


Private Function CLEARTEXT()
iRow = 0
iSet = 0
txtName.Text = ""
txtFrom.Text = ""
txtTo.Text = ""
txtNoDays.Text = ""
txtNoPlayer.Text = ""
txtPlayerToCount.Text = ""
txtHDCPDivisor.Text = ""
txtParGrossPts.Text = ""
txtAllowTeam.Text = ""
txtAllowPartner.Text = ""
chkTeamPlay.Value = 0
chkPartner.Value = 0
chkIndividualPlay.Value = 0
cmbScoring.ListIndex = 0
cmbPointsToCount.ListIndex = 0
cmbPointsToCountIndi.ListIndex = 0
cmbPointsToCountPartner.ListIndex = 0
cmbOrder.ListIndex = -1
cmbTeamAverage.ListIndex = 0
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
imgLogo.Picture = LoadPicture("")
imgLocked.Visible = False
lstLocation.ListItems.Clear
CUSTOM_GRID
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
            .Buttons(21).Image = 11
            .Buttons(23).Image = 12
            .Buttons(25).Image = 13
            .Buttons(27).Image = 14
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "First"
            .Buttons(9).Caption = "Back"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            .Buttons(19).Caption = "Activate"
            .Buttons(21).Caption = "Export"
            .Buttons(23).Caption = "Lock"
            .Buttons(25).Caption = "Refresh"
            .Buttons(27).Caption = "Close"
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
            .Buttons(25).Enabled = True
            .Buttons(27).Enabled = True
            .Buttons(1).ToolTipText = "NEW (Ins)"
            .Buttons(3).ToolTipText = "EDIT (F2)"
            .Buttons(5).ToolTipText = "DELETE (Del)"
            .Buttons(7).ToolTipText = "FIRST (Home)"
            .Buttons(9).ToolTipText = "BACK (PgUp)"
            .Buttons(11).ToolTipText = "NEXT (PgDown)"
            .Buttons(13).ToolTipText = "LAST (End)"
            .Buttons(15).ToolTipText = "FIND (F6)"
            .Buttons(17).ToolTipText = "PRINT (F9)"
            .Buttons(19).ToolTipText = "ACTIVATE (F7)"
            .Buttons(19).ToolTipText = "EXPORT (F3)"
            .Buttons(21).ToolTipText = "LOCKED (F8)"
            .Buttons(19).ToolTipText = "REFRESH (F11)"
            .Buttons(21).ToolTipText = "CLOSE (Esc)"
        Case 2      '=== ADD/EDIT ====
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 15
            .Buttons(9).Image = 16
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(21).Image = 11
            .Buttons(23).Image = 12
            .Buttons(25).Image = 13
            .Buttons(27).Image = 14
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "Save"
            .Buttons(9).Caption = "Undo"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            .Buttons(19).Caption = "Activate"
            .Buttons(21).Caption = "Export"
            .Buttons(23).Caption = "Lock"
            .Buttons(25).Caption = "Refresh"
            .Buttons(27).Caption = "Close"
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
            .Buttons(25).Enabled = False
            .Buttons(27).Enabled = False
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
            .Buttons(25).ToolTipText = ""
            .Buttons(27).ToolTipText = ""
        Case 3      '=== FIND ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 4
            .Buttons(9).Image = 16
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(21).Image = 11
            .Buttons(23).Image = 12
            .Buttons(25).Image = 13
            .Buttons(27).Image = 14
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "First"
            .Buttons(9).Caption = "Undo"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            .Buttons(19).Caption = "Activate"
            .Buttons(21).Caption = "Export"
            .Buttons(23).Caption = "Lock"
            .Buttons(25).Caption = "Refresh"
            .Buttons(27).Caption = "Close"
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
            .Buttons(25).Enabled = False
            .Buttons(27).Enabled = False
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
            .Buttons(25).ToolTipText = ""
            .Buttons(27).ToolTipText = ""
        Case 4      '=== EMPTY DETAIL ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 15
            .Buttons(9).Image = 16
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(21).Image = 11
            .Buttons(23).Image = 12
            .Buttons(25).Image = 13
            .Buttons(27).Image = 14
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "Save"
            .Buttons(9).Caption = "Undo"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            .Buttons(19).Caption = "Activate"
            .Buttons(21).Caption = "Export"
            .Buttons(23).Caption = "Lock"
            .Buttons(25).Caption = "Refresh"
            .Buttons(27).Caption = "Close"
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
            .Buttons(25).Enabled = False
            .Buttons(27).Enabled = False
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
            .Buttons(25).ToolTipText = ""
            .Buttons(27).ToolTipText = ""
        Case 5      '=== NOT EMPTY DETAIL ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 15
            .Buttons(9).Image = 16
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(21).Image = 11
            .Buttons(23).Image = 12
            .Buttons(25).Image = 13
            .Buttons(27).Image = 14
            .Buttons(1).Caption = "Add"
            .Buttons(3).Caption = "Edit"
            .Buttons(5).Caption = "Delete"
            .Buttons(7).Caption = "Save"
            .Buttons(9).Caption = "Undo"
            .Buttons(11).Caption = "Next"
            .Buttons(13).Caption = "Last"
            .Buttons(15).Caption = "Find"
            .Buttons(17).Caption = "Print"
            .Buttons(19).Caption = "Activate"
            .Buttons(21).Caption = "Export"
            .Buttons(23).Caption = "Lock"
            .Buttons(25).Caption = "Refresh"
            .Buttons(27).Caption = "Close"
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
            .Buttons(25).Enabled = False
            .Buttons(27).Enabled = False
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
            .Buttons(25).ToolTipText = ""
            .Buttons(27).ToolTipText = ""
    End Select
End With

'Set Toolbar1.ImageList = ImageList1
'With Toolbar1
'    Select Case isAction
'        Case 1
'            .Buttons(1).Image = 1
'            .Buttons(3).Image = 2
'            .Buttons(5).Image = 3
'            .Buttons(11).Image = 6
'            .Buttons(13).Image = 7
'            .Buttons(15).Image = 13
'            .Buttons(17).Image = 16
'            .Buttons(19).Image = 15
'            .Buttons(21).Image = 9
'            .Buttons(23).Image = 10
'            .Buttons(1).Enabled = True
'            .Buttons(3).Enabled = True
'            .Buttons(5).Enabled = True
'            .Buttons(7).Image = 4
'            .Buttons(7).Caption = "First"
'            .Buttons(9).Image = 5
'            .Buttons(9).Caption = "Back"
'            .Buttons(7).Enabled = True
'            .Buttons(9).Enabled = True
'            .Buttons(11).Enabled = True
'            .Buttons(13).Enabled = True
'            .Buttons(15).Enabled = True
'            .Buttons(17).Enabled = True
'            .Buttons(19).Enabled = True
'            .Buttons(21).Enabled = True
'            .Buttons(23).Enabled = True
'            .Buttons(1).ToolTipText = "NEW (Ins)"
'            .Buttons(3).ToolTipText = "EDIT (F2)"
'            .Buttons(5).ToolTipText = "DELETE (Del)"
'            .Buttons(7).ToolTipText = "FIRST (Home)"
'            .Buttons(9).ToolTipText = "BACK (Page Up)"
'            .Buttons(11).ToolTipText = "NEXT (Page Down)"
'            .Buttons(13).ToolTipText = "LAST (End)"
'            .Buttons(15).ToolTipText = "ACTIVATE (F7)"
'            .Buttons(17).ToolTipText = "EXPORT (F3)"
'            .Buttons(19).ToolTipText = "LOCKED (F8)"
'            .Buttons(21).ToolTipText = "PRINT (F9)"
'            .Buttons(23).ToolTipText = "CLOSE (Esc)"
'        Case 2
'            .Buttons(1).Image = 1
'            .Buttons(3).Image = 2
'            .Buttons(5).Image = 3
'            .Buttons(11).Image = 6
'            .Buttons(13).Image = 7
'            .Buttons(15).Image = 13
'            .Buttons(17).Image = 16
'            .Buttons(19).Image = 15
'            .Buttons(21).Image = 9
'            .Buttons(23).Image = 10
'            .Buttons(1).Enabled = False
'            .Buttons(3).Enabled = False
'            .Buttons(5).Enabled = False
'            .Buttons(7).Image = 11
'            .Buttons(7).Caption = "Save"
'            .Buttons(9).Image = 12
'            .Buttons(9).Caption = "Undo"
'            .Buttons(7).Enabled = True
'            .Buttons(9).Enabled = True
'            .Buttons(11).Enabled = False
'            .Buttons(13).Enabled = False
'            .Buttons(15).Enabled = False
'            .Buttons(17).Enabled = False
'            .Buttons(19).Enabled = False
'            .Buttons(21).Enabled = False
'            .Buttons(23).Enabled = False
'            .Buttons(1).ToolTipText = ""
'            .Buttons(3).ToolTipText = ""
'            .Buttons(5).ToolTipText = ""
'            .Buttons(7).ToolTipText = "SAVE (F5)"
'            .Buttons(9).ToolTipText = "UNDO (Esc)"
'            .Buttons(11).ToolTipText = ""
'            .Buttons(13).ToolTipText = ""
'            .Buttons(15).ToolTipText = ""
'            .Buttons(17).ToolTipText = ""
'            .Buttons(19).ToolTipText = ""
'            .Buttons(21).ToolTipText = ""
'            .Buttons(23).ToolTipText = ""
'    End Select
'End With
End Sub
Private Function CUSTOM_GRID()
HEADER1$ = ""
With FGrid
    .Clear
    HEADER1$ = HEADER1$ & "|" & _
               "Class" & "|" & _
               "From" & "|" & _
               "To"
    .FormatString = HEADER1$
    .ColWidth(1) = 800     'Class
    .ColWidth(2) = 1000     'From
    .ColWidth(3) = 1000     'To
    .ColAlignment(1) = 2
    .ColAlignment(2) = flexAlignRightCenter
    .ColAlignment(3) = flexAlignRightCenter
    .Rows = 2
End With

HEADER1$ = ""
With FGridIndex
    .Clear
    HEADER1$ = HEADER1$ & "|" & _
               "Class" & "|" & _
               "From" & "|" & _
               "To"
    .FormatString = HEADER1$
    .ColWidth(1) = 800     'Class
    .ColWidth(2) = 1000     'From
    .ColWidth(3) = 1000     'To
    .ColAlignment(1) = 2
    .ColAlignment(2) = flexAlignRightCenter
    .ColAlignment(3) = flexAlignRightCenter
    .Rows = 2
End With

End Function

Private Sub b8TitleBar3_CLoseClick()
cmdCancelPrint_Click
End Sub

Private Sub cmbLocation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    With lstLocation.ListItems
        .Item(iRow).SubItems(1) = cmbLocation.ItemData(cmbLocation.ListIndex)
        .Item(iRow).SubItems(2) = cmbLocation.List(cmbLocation.ListIndex)
    End With
    picSLine.Visible = False
    picMain.Enabled = True
    lstLocation.SetFocus
End If
End Sub

Private Sub cmbScoring_Click()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    txtParGrossPts.Text = "0"
    t = "SELECT tbl_Scoring_System.* " & _
        " FROM tbl_Scoring_System " & _
        " WHERE (PK = " & cmbScoring.ListIndex & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtParGrossPts.Text = rt!ParPoints
    End If
    rt.Close
End If
End Sub

Private Sub cmdCancelPrint_Click()
picMain.Enabled = True
picToolbar.Enabled = True
picPrint.Visible = False
End Sub

Private Sub cmdOKPrint_Click()
With MainForm.CommonDialog1
    .CancelError = True
    On Error GoTo ErrorHandler
    .DialogTitle = "Save"
    .Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
    .ShowSave
    sFileName = Trim(.Filename)
End With
picMain.Enabled = False
picToolbar.Enabled = False
'TimerScoreCard.Enabled = True
TimerScoreCard1.Enabled = True
Exit Sub
ErrorHandler:
Exit Sub
End Sub

Private Sub FGrid_EnterCell()
PressCount = 0
End Sub

Private Sub FGrid_GotFocus()
FGrid.ForeColorSel = &HFFFFFF
FGrid.BackColorSel = &H800000
Focus_Class = 1
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
With FGrid
    If TRANSACTIONTYPE = is_ADDING Or _
    TRANSACTIONTYPE = is_EDITTING Then
        If .Col = 1 Then
            If KeyAscii = 8 Then
                LTEXT = IIf(Len(.TextMatrix(.ROW, .Col)) > 0, Len(.TextMatrix(.ROW, .Col)) - 1, 0)
                .TextMatrix(.ROW, .Col) = Mid(.TextMatrix(.ROW, .Col), 1, LTEXT)
            ElseIf KeyAscii >= 1 And KeyAscii <= 7 Then
            ElseIf KeyAscii >= 9 And KeyAscii <= 13 Then
            ElseIf KeyAscii >= 14 And KeyAscii <= 44 Then
            ElseIf KeyAscii = 47 Then
            ElseIf KeyAscii >= 58 And KeyAscii <= 126 Then
                PressCount = PressCount + 1
                If PressCount > 1 Then
                    '.TextMatrix(.ROW, .Col) = .TextMatrix(.ROW, .Col) & UCase(CStr(Chr(KeyAscii)))
                    .TextMatrix(.ROW, .Col) = UCase(CStr(Chr(KeyAscii)))
                Else
                    .TextMatrix(.ROW, .Col) = UCase(CStr(Chr(KeyAscii)))
                End If
            End If
        ElseIf .Col = 2 Or .Col = 3 Then
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
    End If
End With
End Sub

Private Sub FGrid_LostFocus()
FGrid.ForeColorSel = &H0&
FGrid.BackColorSel = &HFFFFFF
Focus_Class = 0
End Sub

'Private Sub FGridIndex_Click()
'
'End Sub

Private Sub FGridIndex_EnterCell()
PressCount = 0
End Sub

Private Sub FGridIndex_GotFocus()
FGridIndex.ForeColorSel = &HFFFFFF
FGridIndex.BackColorSel = &H800000
Focus_ClassIndex = 1
End Sub

Private Sub FGridIndex_KeyPress(KeyAscii As Integer)
With FGridIndex
    If TRANSACTIONTYPE = is_ADDING Or _
    TRANSACTIONTYPE = is_EDITTING Then
        If .Col = 1 Then
            If KeyAscii = 8 Then
                LTEXT = IIf(Len(.TextMatrix(.ROW, .Col)) > 0, Len(.TextMatrix(.ROW, .Col)) - 1, 0)
                .TextMatrix(.ROW, .Col) = Mid(.TextMatrix(.ROW, .Col), 1, LTEXT)
            ElseIf KeyAscii >= 1 And KeyAscii <= 7 Then
            ElseIf KeyAscii >= 9 And KeyAscii <= 13 Then
            ElseIf KeyAscii >= 14 And KeyAscii <= 44 Then
            ElseIf KeyAscii = 47 Then
            ElseIf KeyAscii >= 58 And KeyAscii <= 126 Then
                PressCount = PressCount + 1
                If PressCount > 1 Then
                    '.TextMatrix(.ROW, .Col) = .TextMatrix(.ROW, .Col) & UCase(CStr(Chr(KeyAscii)))
                    .TextMatrix(.ROW, .Col) = UCase(CStr(Chr(KeyAscii)))
                Else
                    .TextMatrix(.ROW, .Col) = UCase(CStr(Chr(KeyAscii)))
                End If
            End If
        ElseIf .Col = 2 Or .Col = 3 Then
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
    End If
End With
End Sub

Private Sub FGridIndex_LostFocus()
FGridIndex.ForeColorSel = &H0&
FGridIndex.BackColorSel = &HFFFFFF
Focus_ClassIndex = 0
End Sub

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert:   PRESS_INSERT
    Case vbKeyF2:       PRESS_F2
    Case vbKeyF3:       PRESS_F3
    Case vbKeyDelete:   PRESS_DELETE
    Case vbKeyF5:       PRESS_F5
    Case vbKeyF7:       PRESS_F7
    Case vbKeyF8:       PRESS_F8
    Case vbKeyF9:       PRESS_F9
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "EventName", "NameEvent", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "EventName", "NameEvent", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "EventName", "NameEvent", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "EventName", "NameEvent", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
FGrid.ForeColorSel = &H0&
FGrid.BackColorSel = &HFFFFFF

FGridIndex.ForeColorSel = &H0&
FGridIndex.BackColorSel = &HFFFFFF

With cmbOrder
    .Clear
    .AddItem "ASC"
    .AddItem "DESC"
End With

With cmbPointsToCount
    .Clear
    .AddItem "--Select--"
    .AddItem "GROSS PTS"
    .AddItem "NET POINTS"
    .AddItem "GROSS & NET POINTS"
End With
With cmbPointsToCountPartner
    .Clear
    .AddItem "--Select--"
    .AddItem "GROSS PTS"
    .AddItem "NET POINTS"
    .AddItem "GROSS & NET POINTS"
End With
With cmbPointsToCountIndi
    .Clear
    .AddItem "--Select--"
    .AddItem "GROSS PTS"
    .AddItem "NET POINTS"
    .AddItem "GROSS & NET POINTS"
End With
With cmbTeamAverage
    .Clear
    .AddItem "--Select--"
    .AddItem "HANDICAP"
    .AddItem "INDEX"
End With
With cmbScoring
    .Clear
    .AddItem "--Select--"
    s = "SELECT tbl_Scoring_System.* " & _
        " FROM tbl_Scoring_System " & _
        " ORDER BY PK"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        .AddItem rs!ScoringSystem
        rs.MoveNext
    Wend
    rs.Close
    
End With

With cmbLocation
    .Clear
    s = "SELECT tbl_Scoring_Location.* " & _
        " FROM tbl_Scoring_Location " & _
        " WHERE (DefaultLocation = 0) " & _
        " ORDER BY ScoringLocation"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        .AddItem rs!ScoringLocation
        .ItemData(.NewIndex) = rs!PK
        rs.MoveNext
    Wend
    rs.Close
End With

CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "EventName", "NameEvent", ""), "is_LOAD"
If Trim(txtName.Text) = "" Then BROWSER GetSetting(App.EXEName, "EventName", "NameEvent", ""), "is_HOME"

tmp = SetWindowLong(txtName.hwnd, GWL_STYLE, GetWindowLong(txtName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picPrint.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub


Private Sub lstLocation_Click()
iFocus = 1
If lstLocation.ListItems.Count = 0 Then Exit Sub
iRow = lstLocation.SelectedItem.Index
End Sub

Private Sub lstLocation_GotFocus()
iFocus = 1
If lstLocation.ListItems.Count = 0 Then Exit Sub
iRow = lstLocation.SelectedItem.Index
End Sub

Private Sub lstLocation_ItemClick(ByVal Item As MSComctlLib.ListItem)
If lstLocation.ListItems.Count = 0 Then Exit Sub
iRow = lstLocation.SelectedItem.Index
End Sub

Private Sub lstLocation_LostFocus()
iFocus = 0
End Sub

Private Sub TimerScoreCard_Timer()
TimerScoreCard.Enabled = False

ArrDate = Split(TournamentRange, " - ", -1, 1)
iDateDiff = DateDiff("d", ArrDate(0), ArrDate(1))

j = 0
picPrint.Visible = False
picProgress.Visible = True
picProgressBar.BackColor = &HFFFFFF
DoEvents

iProgressValue = 0
DoEvents
WorkbookName = sFileName
iWorkSheet = 1: RowCnt = 0: HeaderRow = 0
Set xlsApp = CreateObject("Excel.Application")
xlsApp.Visible = False
xlsApp.Workbooks.Add
xlsApp.DisplayAlerts = False
xlsApp.Workbooks(1).Sheets(2).Delete
xlsApp.Workbooks(1).Sheets(2).Delete
With xlsApp.Workbooks(1).Sheets(iWorkSheet)
    .Activate
    .Name = "Scores"
End With
With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
    a = "SELECT dbo.tbl_Scoring_TournamentInfo.* " & _
        " From dbo.tbl_Scoring_TournamentInfo " & _
        " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    If ra.State = adStateOpen Then ra.Close
    ra.Open a, ConnOmega
    If ra.RecordCount > 0 Then
        RowCnt = RowCnt + 1
        HeaderRow = HeaderRow + 1
        ColCnt = 0
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = ra!TournamentName
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 10
        .Range(strRange).Font.Bold = True
    
        RowCnt = RowCnt + 1
        HeaderRow = HeaderRow + 1
        ColCnt = 0
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = "Range : " & Format(ra!TournamentStart, "mm/dd/yyyy") & " - " & Format(ra!TournamentEnd, "mm/dd/yyyy")
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 8
        .Range(strRange).Font.Bold = False
        
        s = "SELECT dbo.tbl_Scoring_System.ScoringSystem " & _
            " FROM dbo.tbl_Scoring_TournamentInfo LEFT OUTER JOIN " & _
            " dbo.tbl_Scoring_System ON dbo.tbl_Scoring_TournamentInfo.Scoring = dbo.tbl_Scoring_System.PK " & _
            " WHERE (dbo.tbl_Scoring_TournamentInfo.PK = " & ra!PK & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            RowCnt = RowCnt + 1
            HeaderRow = HeaderRow + 1
            ColCnt = 0
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = "Scoring : " & rs!ScoringSystem
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = True
        End If
        rs.Close
        
        RowCnt = RowCnt + 1
        HeaderRow = HeaderRow + 1
        ColCnt = 0
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = "PLAYERS SCORES"
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 8
        .Range(strRange).Font.Bold = True
                
        RowCnt = RowCnt + 1
        HeaderRow = HeaderRow + 1
        ColCnt = 3
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = "Hole"
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 8
        .Range(strRange).Font.Bold = False
        s = "SELECT TOP 1 PK " & _
            " From dbo.tbl_Scoring_Yardage_Par_HandicapIndex_Master " & _
            " WHERE (EffectDate <= '" & FormatDateTime(ra!TournamentEnd, vbShortDate) & "') " & _
            " ORDER BY EffectDate DESC"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            l = 0: k = 0
            t = "SELECT Hole " & _
                " From dbo.tbl_Scoring_Yardage_Par_HandicapIndex " & _
                " Where (MasterKey = " & rs!PK & ") " & _
                " ORDER BY Line"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            While Not rt.EOF
                l = l + 1
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = rt!Hole
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                .Range(strRange).HorizontalAlignment = 3
                .Columns(ColCnt).ColumnWidth = 3
                If l = 9 Then
                    l = 0: k = k + 1
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    If k = 1 Then
                        .Range(strRange).Value = "F-9"
                        .Columns(ColCnt).ColumnWidth = 6
                    Else
                        .Range(strRange).Value = "B-9"
                        .Columns(ColCnt).ColumnWidth = 6
                    End If
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    .Range(strRange).HorizontalAlignment = 3
                End If
                rt.MoveNext
            Wend
            rt.Close
        End If
        rs.Close
        
        RowCnt = RowCnt + 1
        HeaderRow = HeaderRow + 1
        ColCnt = 3
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = "Par"
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 8
        .Range(strRange).Font.Bold = False
        s = "SELECT TOP 1 PK " & _
            " From dbo.tbl_Scoring_Yardage_Par_HandicapIndex_Master " & _
            " WHERE (EffectDate <= '" & FormatDateTime(ra!TournamentEnd, vbShortDate) & "') " & _
            " ORDER BY EffectDate DESC"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            l = 0: k = 0: sMinRange = "": sMaxRange = ""
            t = "SELECT Par " & _
                " From dbo.tbl_Scoring_Yardage_Par_HandicapIndex " & _
                " Where (MasterKey = " & rs!PK & ") " & _
                " ORDER BY Line"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            While Not rt.EOF
                l = l + 1
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = rt!Par
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                .Range(strRange).HorizontalAlignment = 3
                sMaxRange = strRange
                If l = 1 Then
                    sMinRange = strRange
                End If
                If l = 9 Then
                    l = 0: k = k + 1
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "=SUM(" & sMinRange & ":" & sMaxRange & ")"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    .Range(strRange).HorizontalAlignment = 3
                End If
                rt.MoveNext
            Wend
            rt.Close
        End If
        rs.Close
        
        RowCnt = RowCnt + 1
        HeaderRow = HeaderRow + 1
        ColCnt = 0
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = "#"
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 8
        .Range(strRange).Font.Bold = False
        .Columns(ColCnt).ColumnWidth = 5
        
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = "Name"
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 8
        .Range(strRange).Font.Bold = False
        .Columns(ColCnt).ColumnWidth = 30
        
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = "Date"
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 8
        .Range(strRange).Font.Bold = False
        
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = "Handicap"
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 8
        .Range(strRange).Font.Bold = False
        s = "SELECT TOP 1 PK " & _
            " From dbo.tbl_Scoring_Yardage_Par_HandicapIndex_Master " & _
            " WHERE (EffectDate <= '" & FormatDateTime(ra!TournamentEnd, vbShortDate) & "') " & _
            " ORDER BY EffectDate DESC"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            l = 0: k = 0
            t = "SELECT HandicapIndex " & _
                " From dbo.tbl_Scoring_Yardage_Par_HandicapIndex " & _
                " Where (MasterKey = " & rs!PK & ") " & _
                " ORDER BY Line"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            While Not rt.EOF
                l = l + 1
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = rt!HandicapIndex
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                .Range(strRange).HorizontalAlignment = 3
                
                If l = 9 Then
                    l = 0: k = k + 1
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = ""
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    .Range(strRange).HorizontalAlignment = 3
                End If
                rt.MoveNext
            Wend
            rt.Close
        End If
        rs.Close
        
        
        strRange = EXCEL_RANGE(1, HeaderRow + 1)
        .Range(strRange).Select
        xlsApp.ActiveWindow.FreezePanes = True
        
        picProgressBar.Visible = True
        Select Case ra!Scoring
            Case 1
                'If ra!IndividualPlay = 0 And ra!TeamPlay = 1 Then
                '    s = ""
                'Else
                If ra!IndividualPlay = 1 Then
                    s = "SELECT dbo.tbl_Scoring_ScoreCard.PlayerKey, dbo.tbl_Scoring_PlayerName.LastName, " & _
                        " dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName " & _
                        " FROM dbo.tbl_Scoring_ScoreCard LEFT OUTER JOIN " & _
                        " dbo.tbl_Scoring_PlayerName ON dbo.tbl_Scoring_ScoreCard.PlayerKey = dbo.tbl_Scoring_PlayerName.PK " & _
                        " Where (dbo.tbl_Scoring_ScoreCard.TournamentKey = " & ra!PK & ") " & _
                        " GROUP BY dbo.tbl_Scoring_ScoreCard.PlayerKey, dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName " & _
                        " ORDER BY dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName"
                End If
            Case 2
                s = "SELECT dbo.tbl_Scoring_ScoreCard.PlayerKey, dbo.tbl_Scoring_PlayerName.LastName, " & _
                    " dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName " & _
                    " FROM dbo.tbl_Scoring_ScoreCard LEFT OUTER JOIN " & _
                    " dbo.tbl_Scoring_PlayerName ON dbo.tbl_Scoring_ScoreCard.PlayerKey = dbo.tbl_Scoring_PlayerName.PK " & _
                    " Where (dbo.tbl_Scoring_ScoreCard.TournamentKey = " & ra!PK & ") " & _
                    " GROUP BY dbo.tbl_Scoring_ScoreCard.PlayerKey, dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName " & _
                    " ORDER BY dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName"
            Case 3
                s = "SELECT dbo.tbl_Scoring_ScoreCard_System36.PlayerKey, dbo.tbl_Scoring_PlayerName.LastName, " & _
                    " dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName " & _
                    " FROM dbo.tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
                    " dbo.tbl_Scoring_PlayerName ON dbo.tbl_Scoring_ScoreCard_System36.PlayerKey = dbo.tbl_Scoring_PlayerName.PK " & _
                    " Where (dbo.tbl_Scoring_ScoreCard_System36.TournamentKey = " & ra!PK & ") " & _
                    " GROUP BY dbo.tbl_Scoring_ScoreCard_System36.PlayerKey, dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName " & _
                    " ORDER BY dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName"
            Case 4
            
            Case 5
                s = "SELECT dbo.tbl_Scoring_ScoreCard.PlayerKey, dbo.tbl_Scoring_PlayerName.LastName, " & _
                    " dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName " & _
                    " FROM dbo.tbl_Scoring_ScoreCard LEFT OUTER JOIN " & _
                    " dbo.tbl_Scoring_PlayerName ON dbo.tbl_Scoring_ScoreCard.PlayerKey = dbo.tbl_Scoring_PlayerName.PK " & _
                    " Where (dbo.tbl_Scoring_ScoreCard.TournamentKey = " & ra!PK & ") " & _
                    " GROUP BY dbo.tbl_Scoring_ScoreCard.PlayerKey, dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName " & _
                    " ORDER BY dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName"
        End Select
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        While Not rs.EOF
            iProgressValue = iProgressValue + 1
            j = j + 1
            
            RowCnt = RowCnt + 1
            ColCnt = 0
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = j
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = rs!LastName & ",  " & rs!FirstName
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = False
            Select Case ra!Scoring
                Case 1
                    If ra!IndividualPlay = 1 Then
                        t = "SELECT PK, PlayerKey, DDate " & _
                            " From dbo.tbl_Scoring_ScoreCard " & _
                            " Where (PlayerKey = " & rs!PlayerKey & ") " & _
                            " ORDER BY DDate"
                    End If
                Case 2
                    t = "SELECT PK, PlayerKey, DDate " & _
                        " From dbo.tbl_Scoring_ScoreCard_ModStableFord " & _
                        " Where (PlayerKey = " & rs!PlayerKey & ") " & _
                        " ORDER BY DDate"
                Case 3
                    t = "SELECT PK, PlayerKey, DDate " & _
                        " From dbo.tbl_Scoring_ScoreCard_System36 " & _
                        " Where (PlayerKey = " & rs!PlayerKey & ") " & _
                        " ORDER BY DDate"
                Case 4
                
                Case 5
                    t = "SELECT PK, DDate, PlayerKey " & _
                        " From dbo.tbl_Scoring_ScoreCard_ModMolave " & _
                        " Where (PlayerKey = " & rs!PlayerKey & ") " & _
                        " ORDER BY DDate"
            End Select
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount = 1 Then
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = Format(rt!dDate, "mm/dd/yyyy")
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                'If ra!Scoring = 3 Then
                '    iScoreGrossNet = 1
                'Else
                    iScoreGrossNet = 3
                'End If
                'For i = 1 To iScoreGrossNet
                    sMinRange = "": sMaxRange = "": l = 0
                    'Select Case i
                    '    Case 1
                            ColCnt = ColCnt + 1
                            strRange = EXCEL_RANGE(ColCnt, RowCnt)
                            .Range(strRange).Value = "Score"
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            Select Case ra!Scoring
                                Case 1
                                    If ra!IndividualPlay = 1 Then
                                        u = "SELECT Score " & _
                                            " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                            " Where (ScoreCardKey = " & rt!PK & ") " & _
                                            " ORDER BY Hole"
                                    End If
                                Case 2
                                    u = "SELECT Score " & _
                                        " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                        " Where (ScoreCardKey = " & rt!PK & ") " & _
                                        " ORDER BY Hole"
                                Case 3
                                   u = "SELECT Score " & _
                                        " From dbo.tbl_Scoring_ScoreCard_System36_Detail " & _
                                        " Where (ScoreCardKey = " & rt!PK & ") " & _
                                        " ORDER BY Hole"
                                Case 4
                                
                                Case 5
                                    u = "SELECT Score " & _
                                        " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                        " Where (ScoreCardKey = " & rt!PK & ") " & _
                                        " ORDER BY Hole"
                            End Select
                            If ru.State = adStateOpen Then ru.Close
                            ru.Open u, ConnOmega
                            While Not ru.EOF
                                l = l + 1
                                ColCnt = ColCnt + 1
                                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                .Range(strRange).Value = ru!Score
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                                .Range(strRange).HorizontalAlignment = 3
                                sMaxRange = strRange
                                If l = 1 Then
                                    sMinRange = strRange
                                End If
                                If l = 9 Then
                                    ColCnt = ColCnt + 1
                                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                    .Range(strRange).Value = "=SUM(" & sMinRange & ":" & sMaxRange & ")"
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                    .Range(strRange).HorizontalAlignment = 3
                                    l = 0
                                End If
                                ru.MoveNext
                            Wend
                            ru.Close
                        'Case 2
                        If chkWithGrossPts.Value = 1 Then
                            l = 0
                            RowCnt = RowCnt + 1
                            ColCnt = 3
                            ColCnt = ColCnt + 1
                            strRange = EXCEL_RANGE(ColCnt, RowCnt)
                            .Range(strRange).Value = "Gross"
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            Select Case ra!Scoring
                                Case 1
                                    If ra!IndividualPlay = 1 Then
                                        u = "SELECT Gross " & _
                                            " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                            " Where (ScoreCardKey = " & rt!PK & ") " & _
                                            " ORDER BY Hole"
                                    End If
                                Case 2
                                    u = "SELECT Gross " & _
                                        " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                        " Where (ScoreCardKey = " & rt!PK & ") " & _
                                        " ORDER BY Hole"
                                Case 3
                                    u = "SELECT Gross " & _
                                        " From dbo.tbl_Scoring_ScoreCard_System36_Detail " & _
                                        " Where (ScoreCardKey = " & rt!PK & ") " & _
                                        " ORDER BY Hole"
                                Case 4
                                
                                Case 5
                                    u = "SELECT Gross " & _
                                        " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                        " Where (ScoreCardKey = " & rt!PK & ") " & _
                                        " ORDER BY Hole"
                            End Select
                            If ru.State = adStateOpen Then ru.Close
                            ru.Open u, ConnOmega
                            While Not ru.EOF
                                l = l + 1
                                ColCnt = ColCnt + 1
                                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                .Range(strRange).Value = ru!Gross
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                                .Range(strRange).HorizontalAlignment = 3
                                sMaxRange = strRange
                                If l = 1 Then
                                    sMinRange = strRange
                                End If
                                If l = 9 Then
                                    ColCnt = ColCnt + 1
                                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                    .Range(strRange).Value = "=SUM(" & sMinRange & ":" & sMaxRange & ")"
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                    .Range(strRange).HorizontalAlignment = 3
                                    l = 0
                                End If
                                ru.MoveNext
                            Wend
                            ru.Close
                        End If
                        'Case 3
                        If chkWithNetPts.Value = 1 Then
                            l = 0
                            RowCnt = RowCnt + 1
                            ColCnt = 3
                            ColCnt = ColCnt + 1
                            strRange = EXCEL_RANGE(ColCnt, RowCnt)
                            .Range(strRange).Value = "Net"
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            Select Case ra!Scoring
                                Case 1
                                    If ra!IndividualPlay = 1 Then
                                        u = "SELECT Net " & _
                                            " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                            " Where (ScoreCardKey = " & rt!PK & ") " & _
                                            " ORDER BY Hole"
                                    End If
                                Case 2
                                    u = "SELECT Net " & _
                                        " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                        " Where (ScoreCardKey = " & rt!PK & ") " & _
                                        " ORDER BY Hole"
                                Case 3
                                    u = "SELECT Net " & _
                                        " From dbo.tbl_Scoring_ScoreCard_System36_Detail " & _
                                        " Where (ScoreCardKey = " & rt!PK & ") " & _
                                        " ORDER BY Hole"
                                Case 4
                                
                                Case 5
                                    u = "SELECT Net " & _
                                        " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                        " Where (ScoreCardKey = " & rt!PK & ") " & _
                                        " ORDER BY Hole"
                            End Select
                            If ru.State = adStateOpen Then ru.Close
                            ru.Open u, ConnOmega
                            While Not ru.EOF
                                l = l + 1
                                ColCnt = ColCnt + 1
                                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                .Range(strRange).Value = ru!Net
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                                .Range(strRange).HorizontalAlignment = 3
                                sMaxRange = strRange
                                If l = 1 Then
                                    sMinRange = strRange
                                End If
                                If l = 9 Then
                                    ColCnt = ColCnt + 1
                                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                    .Range(strRange).Value = "=SUM(" & sMinRange & ":" & sMaxRange & ")"
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                    .Range(strRange).HorizontalAlignment = 3
                                    l = 0
                                End If
                                ru.MoveNext
                            Wend
                            ru.Close
                        End If
                    'End Select
                'Next i
            ElseIf rt.RecordCount > 1 Then
                k = 0
                While Not rt.EOF
                    k = k + 1
                    If k > 1 Then
                        RowCnt = RowCnt + 1
                        ColCnt = 2
                    End If
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = Format(rt!dDate, "mm/dd/yyyy")
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    'If ra!Scoring = 3 Then
                    '    iScoreGrossNet = 1
                    'Else
                        iScoreGrossNet = 3
                    'End If
                    'For i = 1 To iScoreGrossNet
                        sMinRange = "": sMaxRange = "": l = 0
                        'Select Case i
                        '    Case 1
                                ColCnt = ColCnt + 1
                                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                .Range(strRange).Value = "Score"
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                                Select Case ra!Scoring
                                    Case 1
                                        If ra!IndividualPlay = 1 Then
                                            u = "SELECT Score " & _
                                                " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                                " Where (ScoreCardKey = " & rt!PK & ") " & _
                                                " ORDER BY Hole"
                                        End If
                                    Case 2
                                       u = "SELECT Score " & _
                                            " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                            " Where (ScoreCardKey = " & rt!PK & ") " & _
                                            " ORDER BY Hole"
                                    Case 3
                                         u = "SELECT Score " & _
                                            " From dbo.tbl_Scoring_ScoreCard_System36_Detail " & _
                                            " Where (ScoreCardKey = " & rt!PK & ") " & _
                                            " ORDER BY Hole"
                                    Case 4
                                    
                                    Case 5
                                        u = "SELECT Score " & _
                                            " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                            " Where (ScoreCardKey = " & rt!PK & ") " & _
                                            " ORDER BY Hole"
                                End Select
                                If ru.State = adStateOpen Then ru.Close
                                ru.Open u, ConnOmega
                                While Not ru.EOF
                                    l = l + 1
                                    ColCnt = ColCnt + 1
                                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                    .Range(strRange).Value = ru!Score
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                    .Range(strRange).HorizontalAlignment = 3
                                    sMaxRange = strRange
                                    If l = 1 Then
                                        sMinRange = strRange
                                    End If
                                    If l = 9 Then
                                        ColCnt = ColCnt + 1
                                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                        .Range(strRange).Value = "=SUM(" & sMinRange & ":" & sMaxRange & ")"
                                        .Range(strRange).Font.Name = "Tahoma"
                                        .Range(strRange).Font.Size = 8
                                        .Range(strRange).Font.Bold = False
                                        .Range(strRange).HorizontalAlignment = 3
                                        l = 0
                                    End If
                                    ru.MoveNext
                                Wend
                                ru.Close
                            'Case 2
                            If chkWithGrossPts.Value = 1 Then
                                l = 0
                                RowCnt = RowCnt + 1
                                ColCnt = 3
                                ColCnt = ColCnt + 1
                                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                .Range(strRange).Value = "Gross"
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                                Select Case ra!Scoring
                                    Case 1
                                        If ra!IndividualPlay = 1 Then
                                            u = "SELECT Gross " & _
                                                " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                                " Where (ScoreCardKey = " & rt!PK & ") " & _
                                                " ORDER BY Hole"
                                        End If
                                    Case 2
                                        u = "SELECT Gross " & _
                                            " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                            " Where (ScoreCardKey = " & rt!PK & ") " & _
                                            " ORDER BY Hole"
                                    Case 3
                                        u = "SELECT Gross " & _
                                            " From dbo.tbl_Scoring_ScoreCard_System36_Detail " & _
                                            " Where (ScoreCardKey = " & rt!PK & ") " & _
                                            " ORDER BY Hole"
                                    Case 4
                                    
                                    Case 5
                                        u = "SELECT Gross " & _
                                            " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                            " Where (ScoreCardKey = " & rt!PK & ") " & _
                                            " ORDER BY Hole"
                                End Select
                                If ru.State = adStateOpen Then ru.Close
                                ru.Open u, ConnOmega
                                While Not ru.EOF
                                    l = l + 1
                                    ColCnt = ColCnt + 1
                                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                    .Range(strRange).Value = ru!Gross
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                    .Range(strRange).HorizontalAlignment = 3
                                    sMaxRange = strRange
                                    If l = 1 Then
                                        sMinRange = strRange
                                    End If
                                    If l = 9 Then
                                        ColCnt = ColCnt + 1
                                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                        .Range(strRange).Value = "=SUM(" & sMinRange & ":" & sMaxRange & ")"
                                        .Range(strRange).Font.Name = "Tahoma"
                                        .Range(strRange).Font.Size = 8
                                        .Range(strRange).Font.Bold = False
                                        .Range(strRange).HorizontalAlignment = 3
                                        l = 0
                                    End If
                                    ru.MoveNext
                                Wend
                                ru.Close
                            End If
                            'Case 3
                            If chkWithNetPts.Value = 1 Then
                                l = 0
                                RowCnt = RowCnt + 1
                                ColCnt = 3
                                ColCnt = ColCnt + 1
                                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                .Range(strRange).Value = "Net"
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                                Select Case ra!Scoring
                                    Case 1
                                        If ra!IndividualPlay = 1 Then
                                            u = "SELECT Net " & _
                                                " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                                " Where (ScoreCardKey = " & rt!PK & ") " & _
                                                " ORDER BY Hole"
                                        End If
                                    Case 2
                                        u = "SELECT Net " & _
                                            " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                            " Where (ScoreCardKey = " & rt!PK & ") " & _
                                            " ORDER BY Hole"
                                    Case 3
                                        u = "SELECT Net " & _
                                            " From dbo.tbl_Scoring_ScoreCard_System36_Detail " & _
                                            " Where (ScoreCardKey = " & rt!PK & ") " & _
                                            " ORDER BY Hole"
                                    Case 4
                                    
                                    Case 5
                                        u = "SELECT Net " & _
                                            " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                            " Where (ScoreCardKey = " & rt!PK & ") " & _
                                            " ORDER BY Hole"
                                End Select
                                If ru.State = adStateOpen Then ru.Close
                                ru.Open u, ConnOmega
                                While Not ru.EOF
                                    l = l + 1
                                    ColCnt = ColCnt + 1
                                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                    .Range(strRange).Value = ru!Net
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                    .Range(strRange).HorizontalAlignment = 3
                                    sMaxRange = strRange
                                    If l = 1 Then
                                        sMinRange = strRange
                                    End If
                                    If l = 9 Then
                                        ColCnt = ColCnt + 1
                                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                        .Range(strRange).Value = "=SUM(" & sMinRange & ":" & sMaxRange & ")"
                                        .Range(strRange).Font.Name = "Tahoma"
                                        .Range(strRange).Font.Size = 8
                                        .Range(strRange).Font.Bold = False
                                        .Range(strRange).HorizontalAlignment = 3
                                        l = 0
                                    End If
                                    ru.MoveNext
                                Wend
                                ru.Close
                            End If
                        'End Select
                    'Next i
                    rt.MoveNext
                Wend
            End If
            rt.Close
            
            UpdateProgress_Caption "Exporting to Excel", picProgressBar, iProgressValue / rs.RecordCount
            rs.MoveNext
        Wend
        rs.Close
        
        .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
    End If
    ra.Close
End With

SAVING4:
    On Error GoTo err_saving4:
    If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
    xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

    xlsApp.Visible = True
    
picProgress.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True

Exit Sub

err_saving4:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING4:

Exit Sub
End Sub

Private Sub TimerScoreCard1_Timer()
TimerScoreCard1.Enabled = False
sTournamentRange = Format(FormatDateTime(txtFrom.Text, vbShortDate), "mm/dd/yyyy") & " - " & Format(FormatDateTime(txtTo.Text, vbShortDate), "mm/dd/yyyy")
ArrDate = Split(sTournamentRange, " - ", -1, 1)
iDateDiff = DateDiff("d", ArrDate(0), ArrDate(1))

picPrint.Visible = False
picProgress.Visible = True
picProgressBar.BackColor = &HFFFFFF
DoEvents

iProgressValue = 0
DoEvents
WorkbookName = sFileName
iWorkSheet = 0
Set xlsApp = CreateObject("Excel.Application")
xlsApp.Visible = False
xlsApp.Workbooks.Add
xlsApp.DisplayAlerts = False
 
v = "SELECT dbo.tbl_Scoring_Location.ScoringLocation, " & _
    " dbo.tbl_Scoring_Location.Abbvt, " & _
    " dbo.tbl_Scoring_TournamentInfo_Location.LocationKey, " & _
    " dbo.tbl_Scoring_TournamentInfo_Location.MasterKey " & _
    " FROM dbo.tbl_Scoring_TournamentInfo_Location LEFT OUTER JOIN " & _
    " dbo.tbl_Scoring_Location ON dbo.tbl_Scoring_TournamentInfo_Location.LocationKey = dbo.tbl_Scoring_Location.PK " & _
    " Where (dbo.tbl_Scoring_TournamentInfo_Location.MasterKey = " & Statusbar1.Panels(1).Text & ") " & _
    " ORDER BY dbo.tbl_Scoring_TournamentInfo_Location.LocationKey"
If rv.State = adStateOpen Then rv.Close
rv.Open v, ConnOmega
While Not rv.EOF
    iWorkSheet = iWorkSheet + 1
    RowCnt = 0: HeaderRow = 0
    With xlsApp.Workbooks(1).Sheets(iWorkSheet)
        .Activate
        .Name = rv!Abbvt
    
        a = "SELECT dbo.tbl_Scoring_TournamentInfo.* " & _
            " From dbo.tbl_Scoring_TournamentInfo " & _
            " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
        If ra.State = adStateOpen Then ra.Close
        ra.Open a, ConnOmega
        If ra.RecordCount > 0 Then
            RowCnt = RowCnt + 1
            HeaderRow = HeaderRow + 1
            ColCnt = 0
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = ra!TournamentName
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 10
            .Range(strRange).Font.Bold = True
        
            RowCnt = RowCnt + 1
            HeaderRow = HeaderRow + 1
            ColCnt = 0
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = "Range : " & Format(ra!TournamentStart, "mm/dd/yyyy") & " - " & Format(ra!TournamentEnd, "mm/dd/yyyy")
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = False
            
            t = "SELECT dbo.tbl_Scoring_System.ScoringSystem " & _
                " FROM dbo.tbl_Scoring_TournamentInfo LEFT OUTER JOIN " & _
                " dbo.tbl_Scoring_System ON dbo.tbl_Scoring_TournamentInfo.Scoring = dbo.tbl_Scoring_System.PK " & _
                " WHERE (dbo.tbl_Scoring_TournamentInfo.PK = " & ra!PK & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                RowCnt = RowCnt + 1
                HeaderRow = HeaderRow + 1
                ColCnt = 0
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = "Scoring : " & rt!ScoringSystem
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = True
            End If
            rt.Close
            
            RowCnt = RowCnt + 1
            HeaderRow = HeaderRow + 1
            ColCnt = 0
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = "PLAYERS SCORES"
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = True
                   
            t = "SELECT TOP 1 PK " & _
                " From dbo.tbl_Scoring_Location_Master " & _
                " WHERE (MasterKey = " & rv!LocationKey & ") " & _
                " AND (EffectDate <= '" & FormatDateTime(ArrDate(0), vbShortDate) & "') " & _
                " ORDER BY EffectDate DESC"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                j = 0
                u = "SELECT Description, H1, H2, H3, H4, H5, H6, H7, H8, H9, " & _
                    " H10, H11, H12, H13, H14, H15, H16, H17, H18 " & _
                    " From dbo.tbl_Scoring_Location_Details " & _
                    " Where (MasterKey = " & rt!PK & ") " & _
                    " ORDER BY Line"
                If ru.State = adStateOpen Then ru.Close
                ru.Open u, ConnOmega
                While Not ru.EOF
                    j = j + 1
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    If CDbl(j) = CDbl(ru.RecordCount) Then
                        ColCnt = 0
                        For i = 1 To 3
                            ColCnt = ColCnt + 1
                            strRange = EXCEL_RANGE(ColCnt, RowCnt)
                            Select Case i
                                Case 1: sValue = "#"
                                Case 2: sValue = "Name"
                                Case 3: sValue = "Date"
                            End Select
                            .Range(strRange).Value = sValue
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            If i = 1 Then
                                .Columns(ColCnt).ColumnWidth = 4
                                .Range(strRange).HorizontalAlignment = 4
                            ElseIf i = 2 Then
                                .Columns(ColCnt).ColumnWidth = 30
                                .Range(strRange).HorizontalAlignment = 2
                            Else
                                .Columns(ColCnt).ColumnWidth = 10
                                .Range(strRange).HorizontalAlignment = 3
                            End If
                        Next i
                    Else
                        ColCnt = 3
                    End If
                    For i = 0 To ru.Fields.Count - 1
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = ru.Fields(i).Value
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        If CDbl(i) > 0 Then
                            .Range(strRange).HorizontalAlignment = 3
                            .Columns(ColCnt).ColumnWidth = 3
                        Else
                            .Columns(ColCnt).ColumnWidth = 13
                        End If
                        
                        If CDbl(j) = 1 Then
                            If CDbl(i) = 9 Then
                                ColCnt = ColCnt + 1
                                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                
                                .Range(strRange).Value = "F-9"
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                                .Range(strRange).HorizontalAlignment = 3
                                .Columns(ColCnt).ColumnWidth = 3
                            ElseIf CDbl(i) = 18 Then
                                ColCnt = ColCnt + 1
                                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                
                                .Range(strRange).Value = "B-9"
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                                .Range(strRange).HorizontalAlignment = 3
                                .Columns(ColCnt).ColumnWidth = 3
                            End If
                        ElseIf CDbl(j) = 2 Then
                            If CDbl(i) = 9 Then
                                ColCnt = ColCnt + 1
                                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                .Range(strRange).Value = "36"
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                                .Range(strRange).HorizontalAlignment = 3
                                .Columns(ColCnt).ColumnWidth = 3
                            ElseIf CDbl(i) = 18 Then
                                ColCnt = ColCnt + 1
                                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                .Range(strRange).Value = "36"
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                                .Range(strRange).HorizontalAlignment = 3
                                .Columns(ColCnt).ColumnWidth = 3
                            End If
                        Else
                            If CDbl(i) = 9 Then
                                ColCnt = ColCnt + 1
                                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                .Range(strRange).Value = ""
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                                .Range(strRange).HorizontalAlignment = 3
                                .Columns(ColCnt).ColumnWidth = 3
                            ElseIf CDbl(i) = 18 Then
                                ColCnt = ColCnt + 1
                                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                .Range(strRange).Value = ""
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                                .Range(strRange).HorizontalAlignment = 3
                                .Columns(ColCnt).ColumnWidth = 3
                            End If
                        End If
                        
                    Next i
                    ru.MoveNext
                Wend
                ru.Close
            End If
            rt.Close
        
            strRange = EXCEL_RANGE(1, HeaderRow + 1)
            .Range(strRange).Select
            xlsApp.ActiveWindow.FreezePanes = True
        
            '----------------------------
            picProgressBar.Visible = True
            
            iTotProgressValue = 0
            t = "SELECT tbl_Scoring_ScoreCard.* " & _
                " FROM tbl_Scoring_ScoreCard " & _
                " WHERE (TournamentKey = " & ra!PK & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            iTotProgressValue = rt.RecordCount
            rt.Close
            Select Case ra!Scoring
                Case 1
                    'If ra!IndividualPlay = 0 And ra!TeamPlay = 1 Then
                    '    s = ""
                    'Else
                    If ra!IndividualPlay = 1 Then
                        s = "SELECT dbo.tbl_Scoring_ScoreCard.PlayerKey, dbo.tbl_Scoring_PlayerName.LastName, " & _
                            " dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName " & _
                            " FROM dbo.tbl_Scoring_ScoreCard LEFT OUTER JOIN " & _
                            " dbo.tbl_Scoring_PlayerName ON dbo.tbl_Scoring_ScoreCard.PlayerKey = dbo.tbl_Scoring_PlayerName.PK " & _
                            " Where (dbo.tbl_Scoring_ScoreCard.TournamentKey = " & ra!PK & ") " & _
                            " And (dbo.tbl_Scoring_ScoreCard.LocationKey = " & rv!LocationKey & ") " & _
                            " GROUP BY dbo.tbl_Scoring_ScoreCard.PlayerKey, dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName " & _
                            " ORDER BY dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName"
                    End If
                Case 2
                    s = "SELECT dbo.tbl_Scoring_ScoreCard.PlayerKey, dbo.tbl_Scoring_PlayerName.LastName, " & _
                        " dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName " & _
                        " FROM dbo.tbl_Scoring_ScoreCard LEFT OUTER JOIN " & _
                        " dbo.tbl_Scoring_PlayerName ON dbo.tbl_Scoring_ScoreCard.PlayerKey = dbo.tbl_Scoring_PlayerName.PK " & _
                        " Where (dbo.tbl_Scoring_ScoreCard.TournamentKey = " & ra!PK & ") " & _
                        " And (dbo.tbl_Scoring_ScoreCard.LocationKey = " & rv!LocationKey & ") " & _
                        " GROUP BY dbo.tbl_Scoring_ScoreCard.PlayerKey, dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName " & _
                        " ORDER BY dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName"
                Case 3
                    s = "SELECT dbo.tbl_Scoring_ScoreCard_System36.PlayerKey, dbo.tbl_Scoring_PlayerName.LastName, " & _
                        " dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName " & _
                        " FROM dbo.tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
                        " dbo.tbl_Scoring_PlayerName ON dbo.tbl_Scoring_ScoreCard_System36.PlayerKey = dbo.tbl_Scoring_PlayerName.PK " & _
                        " Where (dbo.tbl_Scoring_ScoreCard_System36.TournamentKey = " & ra!PK & ") " & _
                        " GROUP BY dbo.tbl_Scoring_ScoreCard_System36.PlayerKey, dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName " & _
                        " ORDER BY dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName"
                Case 4
                
                Case 5
                    s = "SELECT dbo.tbl_Scoring_ScoreCard.PlayerKey, dbo.tbl_Scoring_PlayerName.LastName, " & _
                        " dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName " & _
                        " FROM dbo.tbl_Scoring_ScoreCard LEFT OUTER JOIN " & _
                        " dbo.tbl_Scoring_PlayerName ON dbo.tbl_Scoring_ScoreCard.PlayerKey = dbo.tbl_Scoring_PlayerName.PK " & _
                        " Where (dbo.tbl_Scoring_ScoreCard.TournamentKey = " & ra!PK & ") " & _
                        " And (dbo.tbl_Scoring_ScoreCard.LocationKey = " & rv!LocationKey & ") " & _
                        " GROUP BY dbo.tbl_Scoring_ScoreCard.PlayerKey, dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName " & _
                        " ORDER BY dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName"
            End Select
            If rs.State = adStateOpen Then rs.Close
            rs.Open s, ConnOmega
            While Not rs.EOF
                iProgressValue = iProgressValue + 1
                j = j + 1
                
                RowCnt = RowCnt + 1
                ColCnt = 0
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = j
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = rs!LastName & ",  " & rs!FirstName
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                Select Case ra!Scoring
                    Case 1
                        If ra!IndividualPlay = 1 Then
                            t = "SELECT PK, PlayerKey, DDate " & _
                                " From dbo.tbl_Scoring_ScoreCard " & _
                                " Where (PlayerKey = " & rs!PlayerKey & ") " & _
                                " And (LocationKey = " & rv!LocationKey & ") " & _
                                " ORDER BY DDate"
                        End If
                    Case 2
                        t = "SELECT PK, PlayerKey, DDate " & _
                            " From dbo.tbl_Scoring_ScoreCard " & _
                            " Where (PlayerKey = " & rs!PlayerKey & ") " & _
                            " And (LocationKey = " & rv!LocationKey & ") " & _
                            " ORDER BY DDate"
                    Case 3
                        t = "SELECT PK, PlayerKey, DDate " & _
                            " From dbo.tbl_Scoring_ScoreCard_System36 " & _
                            " Where (PlayerKey = " & rs!PlayerKey & ") " & _
                            " ORDER BY DDate"
                    Case 4
                    
                    Case 5
                        t = "SELECT PK, PlayerKey, DDate " & _
                            " From dbo.tbl_Scoring_ScoreCard " & _
                            " Where (PlayerKey = " & rs!PlayerKey & ") " & _
                            " And (LocationKey = " & rv!LocationKey & ") " & _
                            " ORDER BY DDate"
                End Select
                If rt.State = adStateOpen Then rt.Close
                rt.Open t, ConnOmega
                If rt.RecordCount = 1 Then
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = Format(rt!dDate, "mm/dd/yyyy")
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    'If ra!Scoring = 3 Then
                    '    iScoreGrossNet = 1
                    'Else
                        iScoreGrossNet = 3
                    'End If
                    'For i = 1 To iScoreGrossNet
                        sMinRange = "": sMaxRange = "": l = 0
                        'Select Case i
                        '    Case 1
                                ColCnt = ColCnt + 1
                                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                .Range(strRange).Value = "Score"
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                                Select Case ra!Scoring
                                    Case 1
                                        If ra!IndividualPlay = 1 Then
                                            u = "SELECT Score " & _
                                                " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                                " Where (ScoreCardKey = " & rt!PK & ") " & _
                                                " ORDER BY Hole"
                                        End If
                                    Case 2
                                        u = "SELECT Score " & _
                                            " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                            " Where (ScoreCardKey = " & rt!PK & ") " & _
                                            " ORDER BY Hole"
                                    Case 3
                                       u = "SELECT Score " & _
                                            " From dbo.tbl_Scoring_ScoreCard_System36_Detail " & _
                                            " Where (ScoreCardKey = " & rt!PK & ") " & _
                                            " ORDER BY Hole"
                                    Case 4
                                    
                                    Case 5
                                        u = "SELECT Score " & _
                                            " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                            " Where (ScoreCardKey = " & rt!PK & ") " & _
                                            " ORDER BY Hole"
                                End Select
                                If ru.State = adStateOpen Then ru.Close
                                ru.Open u, ConnOmega
                                While Not ru.EOF
                                    l = l + 1
                                    ColCnt = ColCnt + 1
                                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                    .Range(strRange).Value = ru!Score
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                    .Range(strRange).HorizontalAlignment = 3
                                    sMaxRange = strRange
                                    If l = 1 Then
                                        sMinRange = strRange
                                    End If
                                    If l = 9 Then
                                        ColCnt = ColCnt + 1
                                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                        .Range(strRange).Value = "=SUM(" & sMinRange & ":" & sMaxRange & ")"
                                        .Range(strRange).Font.Name = "Tahoma"
                                        .Range(strRange).Font.Size = 8
                                        .Range(strRange).Font.Bold = False
                                        .Range(strRange).HorizontalAlignment = 3
                                        l = 0
                                    End If
                                    ru.MoveNext
                                Wend
                                ru.Close
                            'Case 2
                            If chkWithGrossPts.Value = 1 Then
                                l = 0
                                RowCnt = RowCnt + 1
                                ColCnt = 3
                                ColCnt = ColCnt + 1
                                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                .Range(strRange).Value = "Gross"
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                                Select Case ra!Scoring
                                    Case 1
                                        If ra!IndividualPlay = 1 Then
                                            u = "SELECT Gross " & _
                                                " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                                " Where (ScoreCardKey = " & rt!PK & ") " & _
                                                " ORDER BY Hole"
                                        End If
                                    Case 2
                                        u = "SELECT Gross " & _
                                            " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                            " Where (ScoreCardKey = " & rt!PK & ") " & _
                                            " ORDER BY Hole"
                                    Case 3
                                        u = "SELECT Gross " & _
                                            " From dbo.tbl_Scoring_ScoreCard_System36_Detail " & _
                                            " Where (ScoreCardKey = " & rt!PK & ") " & _
                                            " ORDER BY Hole"
                                    Case 4
                                    
                                    Case 5
                                        u = "SELECT Gross " & _
                                            " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                            " Where (ScoreCardKey = " & rt!PK & ") " & _
                                            " ORDER BY Hole"
                                End Select
                                If ru.State = adStateOpen Then ru.Close
                                ru.Open u, ConnOmega
                                While Not ru.EOF
                                    l = l + 1
                                    ColCnt = ColCnt + 1
                                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                    .Range(strRange).Value = ru!Gross
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                    .Range(strRange).HorizontalAlignment = 3
                                    sMaxRange = strRange
                                    If l = 1 Then
                                        sMinRange = strRange
                                    End If
                                    If l = 9 Then
                                        ColCnt = ColCnt + 1
                                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                        .Range(strRange).Value = "=SUM(" & sMinRange & ":" & sMaxRange & ")"
                                        .Range(strRange).Font.Name = "Tahoma"
                                        .Range(strRange).Font.Size = 8
                                        .Range(strRange).Font.Bold = False
                                        .Range(strRange).HorizontalAlignment = 3
                                        l = 0
                                    End If
                                    ru.MoveNext
                                Wend
                                ru.Close
                            End If
                            'Case 3
                            If chkWithNetPts.Value = 1 Then
                                l = 0
                                RowCnt = RowCnt + 1
                                ColCnt = 3
                                ColCnt = ColCnt + 1
                                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                .Range(strRange).Value = "Net"
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                                Select Case ra!Scoring
                                    Case 1
                                        If ra!IndividualPlay = 1 Then
                                            u = "SELECT Net " & _
                                                " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                                " Where (ScoreCardKey = " & rt!PK & ") " & _
                                                " ORDER BY Hole"
                                        End If
                                    Case 2
                                        u = "SELECT Net " & _
                                            " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                            " Where (ScoreCardKey = " & rt!PK & ") " & _
                                            " ORDER BY Hole"
                                    Case 3
                                        u = "SELECT Net " & _
                                            " From dbo.tbl_Scoring_ScoreCard_System36_Detail " & _
                                            " Where (ScoreCardKey = " & rt!PK & ") " & _
                                            " ORDER BY Hole"
                                    Case 4
                                    
                                    Case 5
                                        u = "SELECT Net " & _
                                            " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                            " Where (ScoreCardKey = " & rt!PK & ") " & _
                                            " ORDER BY Hole"
                                End Select
                                If ru.State = adStateOpen Then ru.Close
                                ru.Open u, ConnOmega
                                While Not ru.EOF
                                    l = l + 1
                                    ColCnt = ColCnt + 1
                                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                    .Range(strRange).Value = ru!Net
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                    .Range(strRange).HorizontalAlignment = 3
                                    sMaxRange = strRange
                                    If l = 1 Then
                                        sMinRange = strRange
                                    End If
                                    If l = 9 Then
                                        ColCnt = ColCnt + 1
                                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                        .Range(strRange).Value = "=SUM(" & sMinRange & ":" & sMaxRange & ")"
                                        .Range(strRange).Font.Name = "Tahoma"
                                        .Range(strRange).Font.Size = 8
                                        .Range(strRange).Font.Bold = False
                                        .Range(strRange).HorizontalAlignment = 3
                                        l = 0
                                    End If
                                    ru.MoveNext
                                Wend
                                ru.Close
                            End If
                        'End Select
                    'Next i
                ElseIf rt.RecordCount > 1 Then
                    k = 0
                    While Not rt.EOF
                        k = k + 1
                        If k > 1 Then
                            RowCnt = RowCnt + 1
                            ColCnt = 2
                        End If
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = Format(rt!dDate, "mm/dd/yyyy")
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        'If ra!Scoring = 3 Then
                        '    iScoreGrossNet = 1
                        'Else
                            iScoreGrossNet = 3
                        'End If
                        'For i = 1 To iScoreGrossNet
                            sMinRange = "": sMaxRange = "": l = 0
                            'Select Case i
                            '    Case 1
                                    ColCnt = ColCnt + 1
                                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                    .Range(strRange).Value = "Score"
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                    Select Case ra!Scoring
                                        Case 1
                                            If ra!IndividualPlay = 1 Then
                                                u = "SELECT Score " & _
                                                    " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                                    " Where (ScoreCardKey = " & rt!PK & ") " & _
                                                    " ORDER BY Hole"
                                            End If
                                        Case 2
                                           u = "SELECT Score " & _
                                                " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                                " Where (ScoreCardKey = " & rt!PK & ") " & _
                                                " ORDER BY Hole"
                                        Case 3
                                             u = "SELECT Score " & _
                                                " From dbo.tbl_Scoring_ScoreCard_System36_Detail " & _
                                                " Where (ScoreCardKey = " & rt!PK & ") " & _
                                                " ORDER BY Hole"
                                        Case 4
                                        
                                        Case 5
                                            u = "SELECT Score " & _
                                                " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                                " Where (ScoreCardKey = " & rt!PK & ") " & _
                                                " ORDER BY Hole"
                                    End Select
                                    If ru.State = adStateOpen Then ru.Close
                                    ru.Open u, ConnOmega
                                    While Not ru.EOF
                                        l = l + 1
                                        ColCnt = ColCnt + 1
                                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                        .Range(strRange).Value = ru!Score
                                        .Range(strRange).Font.Name = "Tahoma"
                                        .Range(strRange).Font.Size = 8
                                        .Range(strRange).Font.Bold = False
                                        .Range(strRange).HorizontalAlignment = 3
                                        sMaxRange = strRange
                                        If l = 1 Then
                                            sMinRange = strRange
                                        End If
                                        If l = 9 Then
                                            ColCnt = ColCnt + 1
                                            strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                            .Range(strRange).Value = "=SUM(" & sMinRange & ":" & sMaxRange & ")"
                                            .Range(strRange).Font.Name = "Tahoma"
                                            .Range(strRange).Font.Size = 8
                                            .Range(strRange).Font.Bold = False
                                            .Range(strRange).HorizontalAlignment = 3
                                            l = 0
                                        End If
                                        ru.MoveNext
                                    Wend
                                    ru.Close
                                'Case 2
                                If chkWithGrossPts.Value = 1 Then
                                    l = 0
                                    RowCnt = RowCnt + 1
                                    ColCnt = 3
                                    ColCnt = ColCnt + 1
                                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                    .Range(strRange).Value = "Gross"
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                    Select Case ra!Scoring
                                        Case 1
                                            If ra!IndividualPlay = 1 Then
                                                u = "SELECT Gross " & _
                                                    " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                                    " Where (ScoreCardKey = " & rt!PK & ") " & _
                                                    " ORDER BY Hole"
                                            End If
                                        Case 2
                                            u = "SELECT Gross " & _
                                                " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                                " Where (ScoreCardKey = " & rt!PK & ") " & _
                                                " ORDER BY Hole"
                                        Case 3
                                            u = "SELECT Gross " & _
                                                " From dbo.tbl_Scoring_ScoreCard_System36_Detail " & _
                                                " Where (ScoreCardKey = " & rt!PK & ") " & _
                                                " ORDER BY Hole"
                                        Case 4
                                        
                                        Case 5
                                            u = "SELECT Gross " & _
                                                " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                                " Where (ScoreCardKey = " & rt!PK & ") " & _
                                                " ORDER BY Hole"
                                    End Select
                                    If ru.State = adStateOpen Then ru.Close
                                    ru.Open u, ConnOmega
                                    While Not ru.EOF
                                        l = l + 1
                                        ColCnt = ColCnt + 1
                                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                        .Range(strRange).Value = ru!Gross
                                        .Range(strRange).Font.Name = "Tahoma"
                                        .Range(strRange).Font.Size = 8
                                        .Range(strRange).Font.Bold = False
                                        .Range(strRange).HorizontalAlignment = 3
                                        sMaxRange = strRange
                                        If l = 1 Then
                                            sMinRange = strRange
                                        End If
                                        If l = 9 Then
                                            ColCnt = ColCnt + 1
                                            strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                            .Range(strRange).Value = "=SUM(" & sMinRange & ":" & sMaxRange & ")"
                                            .Range(strRange).Font.Name = "Tahoma"
                                            .Range(strRange).Font.Size = 8
                                            .Range(strRange).Font.Bold = False
                                            .Range(strRange).HorizontalAlignment = 3
                                            l = 0
                                        End If
                                        ru.MoveNext
                                    Wend
                                    ru.Close
                                End If
                                'Case 3
                                If chkWithNetPts.Value = 1 Then
                                    l = 0
                                    RowCnt = RowCnt + 1
                                    ColCnt = 3
                                    ColCnt = ColCnt + 1
                                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                    .Range(strRange).Value = "Net"
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                    Select Case ra!Scoring
                                        Case 1
                                            If ra!IndividualPlay = 1 Then
                                                u = "SELECT Net " & _
                                                    " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                                    " Where (ScoreCardKey = " & rt!PK & ") " & _
                                                    " ORDER BY Hole"
                                            End If
                                        Case 2
                                            u = "SELECT Net " & _
                                                " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                                " Where (ScoreCardKey = " & rt!PK & ") " & _
                                                " ORDER BY Hole"
                                        Case 3
                                            u = "SELECT Net " & _
                                                " From dbo.tbl_Scoring_ScoreCard_System36_Detail " & _
                                                " Where (ScoreCardKey = " & rt!PK & ") " & _
                                                " ORDER BY Hole"
                                        Case 4
                                        
                                        Case 5
                                            u = "SELECT Net " & _
                                                " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                                " Where (ScoreCardKey = " & rt!PK & ") " & _
                                                " ORDER BY Hole"
                                    End Select
                                    If ru.State = adStateOpen Then ru.Close
                                    ru.Open u, ConnOmega
                                    While Not ru.EOF
                                        l = l + 1
                                        ColCnt = ColCnt + 1
                                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                        .Range(strRange).Value = ru!Net
                                        .Range(strRange).Font.Name = "Tahoma"
                                        .Range(strRange).Font.Size = 8
                                        .Range(strRange).Font.Bold = False
                                        .Range(strRange).HorizontalAlignment = 3
                                        sMaxRange = strRange
                                        If l = 1 Then
                                            sMinRange = strRange
                                        End If
                                        If l = 9 Then
                                            ColCnt = ColCnt + 1
                                            strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                            .Range(strRange).Value = "=SUM(" & sMinRange & ":" & sMaxRange & ")"
                                            .Range(strRange).Font.Name = "Tahoma"
                                            .Range(strRange).Font.Size = 8
                                            .Range(strRange).Font.Bold = False
                                            .Range(strRange).HorizontalAlignment = 3
                                            l = 0
                                        End If
                                        ru.MoveNext
                                    Wend
                                    ru.Close
                                End If
                            'End Select
                        'Next i
                        rt.MoveNext
                    Wend
                End If
                rt.Close
                
                UpdateProgress_Caption "Exporting to Excel", picProgressBar, iProgressValue / iTotProgressValue
                rs.MoveNext
            Wend
            rs.Close
            
            '----------------------------
        End If
        ra.Close
        
        .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
    End With
    rv.MoveNext
Wend
rv.Close

If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

xlsApp.Visible = True

picProgress.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "EventName", "NameEvent", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "EventName", "NameEvent", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "EventName", "NameEvent", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "EventName", "NameEvent", ""), "is_END"
    Case "Find"
    Case "Print":   PRESS_F9
    Case "Set":     PRESS_F7
    Case "Lock":    PRESS_F8
    Case "Export":  PRESS_F3
    Case "Refresh":
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtAllowPartner_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtAllowTeam_GotFocus()
HTEXT txtAllowTeam
End Sub

Private Sub txtAllowTeam_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtFrom_GotFocus()
HTEXT txtFrom
End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtTo.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtName.SetFocus
End If
End Sub

Private Sub txtFrom_LostFocus()
If IsDate(txtFrom.Text) = True Then
    txtFrom.Text = Format(FormatDateTime(txtFrom.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub


Private Sub txtHDCPDivisor_GotFocus()
HTEXT txtHDCPDivisor
End Sub

Private Sub txtHDCPDivisor_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtName_GotFocus()
HTEXT txtName
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtFrom.SetFocus
ElseIf KeyCode = vbKeyUp Then

End If
End Sub

Private Sub txtNoDays_GotFocus()
HTEXT txtNoDays
End Sub

Private Sub txtNoDays_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtNoPlayer.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtTo.SetFocus
End If
End Sub

Private Sub txtNoDays_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtNoPlayer_GotFocus()
HTEXT txtNoPlayer
End Sub

Private Sub txtNoPlayer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    
ElseIf KeyCode = vbKeyUp Then
    txtNoDays.SetFocus
End If
End Sub

Private Sub txtNoPlayer_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtParGrossPts_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtPlayerToCount_GotFocus()
HTEXT txtPlayerToCount
End Sub

Private Sub txtPlayerToCount_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtTo_GotFocus()
HTEXT txtTo
End Sub

Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtNoDays.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtFrom.SetFocus
End If
End Sub

Private Sub txtTo_LostFocus()
If IsDate(txtTo.Text) = True Then
    txtTo.Text = Format(FormatDateTime(txtTo.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub
