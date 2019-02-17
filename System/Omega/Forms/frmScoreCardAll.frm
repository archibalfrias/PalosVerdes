VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmScoreCardAll 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13890
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
   ScaleHeight     =   8985
   ScaleWidth      =   13890
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimerReportMolave 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   4200
   End
   Begin VB.Timer TimerReportStableford 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   3240
   End
   Begin VB.PictureBox picPrint 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   4680
      ScaleHeight     =   3015
      ScaleWidth      =   4335
      TabIndex        =   99
      Top             =   2400
      Visible         =   0   'False
      Width           =   4335
      Begin RPVGCC.b8Container picPrint1 
         Height          =   3015
         Left            =   0
         TabIndex        =   100
         Top             =   0
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   5318
         BackColor       =   15396057
         Begin VB.ComboBox cmbReportType 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   108
            Top             =   480
            Width           =   4095
         End
         Begin VB.ComboBox cmbDivision 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   107
            Top             =   1560
            Width           =   2535
         End
         Begin VB.ComboBox cmbGroup 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   106
            Top             =   1200
            Width           =   4095
         End
         Begin VB.CommandButton cmdCancelPrint 
            Height          =   480
            Left            =   2235
            Picture         =   "frmScoreCardAll.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   105
            Top             =   2355
            Width           =   1560
         End
         Begin VB.CommandButton cmdOKPrint 
            Height          =   480
            Left            =   555
            Picture         =   "frmScoreCardAll.frx":075C
            Style           =   1  'Graphical
            TabIndex        =   104
            Top             =   2355
            Width           =   1560
         End
         Begin VB.TextBox txtTop 
            Height          =   315
            Left            =   3240
            TabIndex        =   103
            Top             =   1560
            Width           =   975
         End
         Begin VB.ComboBox cmbGender 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   102
            Top             =   840
            Width           =   4095
         End
         Begin VB.TextBox txtDatePrint 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1440
            TabIndex        =   101
            Top             =   1920
            Width           =   1815
         End
         Begin RPVGCC.b8TitleBar b8TitleBar3 
            Height          =   345
            Left            =   40
            TabIndex        =   109
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
            Icon            =   "frmScoreCardAll.frx":0DCE
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Top"
            Height          =   255
            Left            =   2760
            TabIndex        =   111
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   255
            Left            =   960
            TabIndex        =   110
            Top             =   1920
            Width           =   495
         End
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "modStable"
      Height          =   375
      Left            =   12960
      TabIndex        =   161
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "modMolave"
      Height          =   375
      Left            =   12960
      TabIndex        =   160
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer TimerTeamAllStableFord 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1680
   End
   Begin VB.Timer TimerAddLocation 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1200
   End
   Begin RPVGCC.b8Container picProgress 
      Height          =   975
      Left            =   3960
      TabIndex        =   97
      Top             =   3360
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1720
      BackColor       =   13023396
      Begin VB.PictureBox picProgressBar 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5355
         TabIndex        =   98
         Top             =   120
         Width           =   5415
      End
   End
   Begin VB.Timer TimerReportModifiedStableford 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   3720
   End
   Begin VB.Timer TimerReportModifiedMolave 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   4680
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   1
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   2
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
         MouseIcon       =   "frmScoreCardAll.frx":1368
         Begin VB.PictureBox Picture10 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   9000
            ScaleHeight     =   375
            ScaleWidth      =   4695
            TabIndex        =   3
            Top             =   240
            Width           =   4695
            Begin VB.TextBox txtTeamName 
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   240
               TabIndex        =   4
               Text            =   "Text1"
               Top             =   0
               Width           =   4095
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
      TabIndex        =   0
      Top             =   8670
      Width           =   13890
      _ExtentX        =   24500
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
      Left            =   13200
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
            Picture         =   "frmScoreCardAll.frx":1682
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardAll.frx":235C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardAll.frx":3036
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardAll.frx":3D10
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardAll.frx":49EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardAll.frx":56C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardAll.frx":639E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardAll.frx":7078
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardAll.frx":7D52
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardAll.frx":862C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardAll.frx":9306
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardAll.frx":9FE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardAll.frx":ACBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardAll.frx":B994
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardAll.frx":C66E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RPVGCC.b8Container picPrintStableford 
      Height          =   2775
      Left            =   4680
      TabIndex        =   153
      Top             =   2520
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4895
      BackColor       =   15396057
      Begin VB.CommandButton cmdOKPrintStableford 
         Height          =   480
         Left            =   555
         Picture         =   "frmScoreCardAll.frx":D348
         Style           =   1  'Graphical
         TabIndex        =   159
         Top             =   2115
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelPrintStableford 
         Height          =   480
         Left            =   2235
         Picture         =   "frmScoreCardAll.frx":D9BA
         Style           =   1  'Graphical
         TabIndex        =   158
         Top             =   2115
         Width           =   1560
      End
      Begin VB.ComboBox cmbGroupStableford 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   157
         Top             =   915
         Width           =   4095
      End
      Begin VB.ComboBox cmbDivisionStableford 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   156
         Top             =   1320
         Width           =   4095
      End
      Begin VB.ComboBox cmbReportTypeStableford 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   155
         Top             =   480
         Width           =   4095
      End
      Begin VB.ComboBox cmbDay 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   154
         Top             =   1680
         Width           =   4095
      End
      Begin RPVGCC.b8TitleBar b8TitleBar4 
         Height          =   345
         Left            =   40
         TabIndex        =   152
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
         Icon            =   "frmScoreCardAll.frx":E116
      End
   End
   Begin RPVGCC.b8Container picSearch 
      Height          =   5055
      Left            =   4680
      TabIndex        =   89
      Top             =   1200
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   8916
      BackColor       =   15396057
      Begin VB.ComboBox cmbLocationSearch 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   162
         Top             =   480
         Width           =   4095
      End
      Begin VB.CommandButton cmdCancelSearch 
         Height          =   480
         Left            =   2280
         Picture         =   "frmScoreCardAll.frx":E6B0
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   4320
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKSearch 
         Height          =   480
         Left            =   480
         Picture         =   "frmScoreCardAll.frx":EE0C
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   4320
         Width           =   1560
      End
      Begin VB.ListBox lstResult 
         Height          =   2595
         Left            =   120
         TabIndex        =   92
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   91
         Top             =   840
         Width           =   4095
      End
      Begin VB.ComboBox cmbDate 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   90
         Top             =   3840
         Width           =   1695
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   95
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
         Icon            =   "frmScoreCardAll.frx":F47E
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   255
         Left            =   1200
         TabIndex        =   96
         Top             =   3840
         Width           =   495
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   840
      ScaleHeight     =   5655
      ScaleWidth      =   12105
      TabIndex        =   5
      Top             =   1160
      Width           =   12105
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   0
         TabIndex        =   151
         Top             =   480
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtCtrl 
         Height          =   285
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin RPVGCC.b8Container b8Container1 
         Height          =   2490
         Left            =   120
         TabIndex        =   7
         Top             =   3000
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
            TabIndex        =   8
            Top             =   50
            Width           =   11780
            Begin VB.PictureBox picScoreEn 
               Appearance      =   0  'Flat
               BackColor       =   &H00C6B8A4&
               ForeColor       =   &H80000008&
               Height          =   975
               Left            =   -10
               ScaleHeight     =   945
               ScaleWidth      =   12300
               TabIndex        =   11
               Top             =   1650
               Width           =   12330
               Begin VB.PictureBox Picture8 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   10640
                  ScaleHeight     =   225
                  ScaleWidth      =   1860
                  TabIndex        =   75
                  Top             =   -10
                  Width           =   1890
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
                     TabIndex        =   77
                     Text            =   "0"
                     Top             =   -10
                     Width           =   570
                  End
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
                     TabIndex        =   76
                     Text            =   "0"
                     Top             =   -10
                     Width           =   570
                  End
               End
               Begin VB.PictureBox Picture7 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   6030
                  ScaleHeight     =   225
                  ScaleWidth      =   540
                  TabIndex        =   73
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
                     TabIndex        =   74
                     Text            =   "0"
                     Top             =   -10
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
                  Index           =   9
                  Left            =   6580
                  TabIndex        =   72
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
                  TabIndex        =   71
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
                  TabIndex        =   70
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
                  TabIndex        =   69
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
                  TabIndex        =   68
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
                  TabIndex        =   67
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
                  TabIndex        =   66
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
                  TabIndex        =   65
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
                  TabIndex        =   64
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
                  TabIndex        =   63
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
                  TabIndex        =   62
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
                  TabIndex        =   61
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
                  TabIndex        =   60
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
                  TabIndex        =   59
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
                  TabIndex        =   58
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
                  TabIndex        =   57
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
                  Index           =   1
                  Left            =   2430
                  TabIndex        =   56
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
                  Index           =   0
                  Left            =   1980
                  TabIndex        =   55
                  Text            =   "0"
                  Top             =   -10
                  Width           =   460
               End
               Begin VB.PictureBox Picture4 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   495
                  Left            =   1980
                  ScaleHeight     =   465
                  ScaleWidth      =   10305
                  TabIndex        =   12
                  Top             =   230
                  Width           =   10335
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
                     TabIndex        =   53
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
                     TabIndex        =   52
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
                     TabIndex        =   51
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
                     TabIndex        =   50
                     Text            =   "0"
                     Top             =   230
                     Width           =   570
                  End
                  Begin VB.TextBox txtGrossPtsF 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Left            =   4040
                     TabIndex        =   49
                     Text            =   "0"
                     Top             =   -10
                     Width           =   570
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0C000&
                     Height          =   255
                     Index           =   17
                     Left            =   8190
                     TabIndex        =   48
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
                     TabIndex        =   47
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
                     TabIndex        =   46
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
                     TabIndex        =   45
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
                     TabIndex        =   44
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
                     TabIndex        =   43
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
                     TabIndex        =   42
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
                     TabIndex        =   41
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
                     TabIndex        =   40
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
                     TabIndex        =   39
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
                     TabIndex        =   38
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
                     TabIndex        =   37
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
                     TabIndex        =   36
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
                     TabIndex        =   35
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
                     TabIndex        =   34
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
                     TabIndex        =   33
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
                     TabIndex        =   32
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
                     TabIndex        =   31
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
                     TabIndex        =   30
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFF80&
                     Height          =   255
                     Index           =   8
                     Left            =   3580
                     TabIndex        =   29
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0C000&
                     Height          =   255
                     Index           =   7
                     Left            =   3130
                     TabIndex        =   28
                     Text            =   "0"
                     Top             =   225
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFF80&
                     Height          =   255
                     Index           =   7
                     Left            =   3130
                     TabIndex        =   27
                     Text            =   "0"
                     Top             =   -15
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0C000&
                     Height          =   255
                     Index           =   6
                     Left            =   2680
                     TabIndex        =   26
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
                     TabIndex        =   25
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
                     TabIndex        =   24
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
                     TabIndex        =   23
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
                     TabIndex        =   22
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
                     TabIndex        =   21
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
                     TabIndex        =   20
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
                     TabIndex        =   19
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
                     TabIndex        =   18
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
                     TabIndex        =   17
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
                     TabIndex        =   16
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
                     TabIndex        =   15
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
                     TabIndex        =   14
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFF80&
                     Height          =   255
                     Index           =   0
                     Left            =   -10
                     TabIndex        =   13
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
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
                  TabIndex        =   80
                  Top             =   470
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
                  TabIndex        =   79
                  Top             =   230
                  Width           =   2010
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
                  TabIndex        =   78
                  Top             =   -15
                  Width           =   2010
               End
            End
            Begin VB.PictureBox picScoreDis 
               Appearance      =   0  'Flat
               BackColor       =   &H00C6B8A4&
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   1680
               Left            =   -10
               ScaleHeight     =   1650
               ScaleWidth      =   12300
               TabIndex        =   9
               Top             =   -10
               Width           =   12330
               Begin MSFlexGridLib.MSFlexGrid FGrid 
                  Height          =   2025
                  Left            =   -105
                  TabIndex        =   10
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
         End
      End
      Begin RPVGCC.b8Container b8Container3 
         Height          =   1455
         Left            =   6360
         TabIndex        =   112
         Top             =   0
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2566
         BackColor       =   49152
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1215
            Left            =   120
            ScaleHeight     =   1215
            ScaleWidth      =   5415
            TabIndex        =   113
            Top             =   120
            Width           =   5415
            Begin VB.TextBox txtTourDate 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1095
               TabIndex        =   116
               Text            =   "06/01/2010 - 06/04/2010"
               Top             =   480
               Width           =   4215
            End
            Begin VB.TextBox txtTournament 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1095
               TabIndex        =   115
               Top             =   120
               Width           =   4215
            End
            Begin VB.TextBox txtLocation 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1080
               TabIndex        =   114
               Top             =   840
               Width           =   4215
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Date Range"
               Height          =   255
               Left            =   120
               TabIndex        =   119
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Tournament"
               Height          =   255
               Left            =   120
               TabIndex        =   118
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Location"
               Height          =   255
               Left            =   120
               TabIndex        =   117
               Top             =   840
               Width           =   975
            End
         End
      End
      Begin RPVGCC.b8Container b8Container2 
         Height          =   1335
         Left            =   8520
         TabIndex        =   120
         Top             =   1560
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   2355
         BackColor       =   49152
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1095
            Left            =   120
            ScaleHeight     =   1095
            ScaleWidth      =   3255
            TabIndex        =   121
            Top             =   120
            Width           =   3255
            Begin VB.TextBox txtSNetTot 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   2400
               TabIndex        =   127
               Top             =   720
               Width           =   735
            End
            Begin VB.TextBox txtSGrossTot 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   2400
               TabIndex        =   126
               Top             =   360
               Width           =   735
            End
            Begin VB.TextBox txtSNetB 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1680
               TabIndex        =   125
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtSGrossB 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1680
               TabIndex        =   124
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox txtSNetF 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1080
               TabIndex        =   123
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtSGrossF 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1080
               TabIndex        =   122
               Top             =   360
               Width           =   495
            End
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Total"
               Height          =   255
               Left            =   2400
               TabIndex        =   133
               Top             =   120
               Width           =   735
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "B - 9"
               Height          =   255
               Left            =   1680
               TabIndex        =   132
               Top             =   120
               Width           =   495
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "F - 9"
               Height          =   255
               Left            =   1080
               TabIndex        =   131
               Top             =   120
               Width           =   495
            End
            Begin VB.Label Label9 
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
               TabIndex        =   130
               Top             =   0
               Width           =   975
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Net Points"
               Height          =   255
               Left            =   120
               TabIndex        =   129
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Gross Points"
               Height          =   255
               Left            =   120
               TabIndex        =   128
               Top             =   360
               Width           =   975
            End
         End
      End
      Begin RPVGCC.b8Container b8Container4 
         Height          =   1335
         Left            =   6360
         TabIndex        =   134
         Top             =   1560
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   2355
         BackColor       =   49152
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1095
            Left            =   120
            ScaleHeight     =   1095
            ScaleWidth      =   1815
            TabIndex        =   135
            Top             =   120
            Width           =   1815
            Begin VB.TextBox txtHandicap 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1080
               TabIndex        =   137
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtClass 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1080
               TabIndex        =   136
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Handicap"
               Height          =   255
               Left            =   120
               TabIndex        =   139
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Class"
               Height          =   255
               Left            =   120
               TabIndex        =   138
               Top             =   600
               Width           =   975
            End
         End
      End
      Begin RPVGCC.b8Container b8Container5 
         Height          =   1215
         Left            =   120
         TabIndex        =   140
         Top             =   0
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2143
         BackColor       =   49152
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   975
            Left            =   120
            ScaleHeight     =   975
            ScaleWidth      =   5895
            TabIndex        =   141
            Top             =   120
            Width           =   5895
            Begin VB.TextBox txtPlayer 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   720
               TabIndex        =   144
               Top             =   120
               Width           =   5055
            End
            Begin VB.TextBox txtDay 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   5040
               TabIndex        =   143
               Top             =   480
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox txtDate 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   720
               TabIndex        =   142
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Player"
               Height          =   255
               Left            =   120
               TabIndex        =   147
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Day"
               Height          =   255
               Left            =   4200
               TabIndex        =   146
               Top             =   480
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               Height          =   255
               Left            =   120
               TabIndex        =   145
               Top             =   480
               Width           =   495
            End
         End
      End
      Begin RPVGCC.b8Container b8Container6 
         Height          =   1575
         Left            =   120
         TabIndex        =   148
         Top             =   1320
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2778
         BackColor       =   49152
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   120
            ScaleHeight     =   1335
            ScaleWidth      =   5895
            TabIndex        =   149
            Top             =   120
            Width           =   5895
            Begin MSComctlLib.ListView lstTeamMates 
               Height          =   1335
               Left            =   0
               TabIndex        =   150
               Top             =   0
               Width           =   5895
               _ExtentX        =   10398
               _ExtentY        =   2355
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   5
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "PlayerName"
                  Object.Width           =   4463
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   2
                  Text            =   "Gross Score"
                  Object.Width           =   1941
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   3
                  Text            =   "Gross Points"
                  Object.Width           =   1941
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   4
                  Text            =   "Net Points"
                  Object.Width           =   1941
               EndProperty
            End
         End
      End
   End
   Begin RPVGCC.b8Container picSearchAdd 
      Height          =   4695
      Left            =   4680
      TabIndex        =   81
      Top             =   1440
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   8281
      BackColor       =   15396057
      Begin VB.TextBox txtSearchAdd 
         Height          =   315
         Left            =   120
         TabIndex        =   86
         Top             =   480
         Width           =   4095
      End
      Begin VB.ListBox lstResultAdd 
         Height          =   2595
         Left            =   120
         TabIndex        =   85
         Top             =   840
         Width           =   4095
      End
      Begin VB.CommandButton cmdOKAdd 
         Height          =   480
         Left            =   480
         Picture         =   "frmScoreCardAll.frx":FA18
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   3960
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelAdd 
         Height          =   480
         Left            =   2280
         Picture         =   "frmScoreCardAll.frx":1008A
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   3960
         Width           =   1560
      End
      Begin VB.TextBox txtDateAdd 
         Height          =   315
         Left            =   1800
         TabIndex        =   82
         Top             =   3480
         Width           =   1215
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   45
         TabIndex        =   87
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
         Icon            =   "frmScoreCardAll.frx":107E6
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   255
         Left            =   1320
         TabIndex        =   88
         Top             =   3480
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmScoreCardAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public TourNoOfPlays    As Double
Dim PlayerKey           As Double
Dim sFileName           As String
Dim tmp                 As Long
Dim cn                  As ADODB.Connection

Dim TableName           As String
Dim DetailTableName     As String
Dim WorkbookName        As String
Dim iWorkSheet          As Integer

Dim TRANSACTIONTYPE     As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim iPlayCnt, dTeamTotal, dB9Total, dF9Total, iDLine, dLastMan, sFieldNum, dDateEnd, _
TeamTmp, RowCnt, ColCnt, strRange, x, i, j, l, k, Arr, iDay, dDate, sTotal, RowCntTmp, _
ColCntTmp, strRangeFrom, strRangeTo, ArrDate, iDateDiff, HeaderRow, iProgressValue, _
dTeamPoints, sMinRange, sMaxRange, iTeamCounter, dTotalTeam, Array1, TourNoOfPlaysTmp, _
strCtrlNo, Arr1, sFileNameMaster, sFileNameDetail, StrFile, sFileArrDet, sFileArr, _
SCardKey, dblPar, dblHandicap, dblScore, dblGross, dblNet, strClass, Columns, ColumnsDet, _
Clustered, sMasterFields, sDetailFields, Arr2, MasterKey, dblGrossPts, Filename, cnt, _
dblCntBackPlayer1, dblCntBackPlayer2, dblEagle, dblTotalHDCP, strPlayerName, dblGrossPts1, _
dblGrossPts2, dblGrossPtsTot, dblCntBackPlayerTot, sFreezePane, ColTop, RowTop, ColCount, _
RowCount, strRange1, ColCountDet, RowCountDet, RowFrom, RowTo, dblCntBackPlayer, sGrossPts, _
iScoreKey

Private Function BROWSER(strCtrl, isAction As String)
'MsgBox isAction
s = ""
Select Case isAction
    Case "is_LOAD"
        If strCtrl <> "" Then
            s = "SELECT TOP 1 tbl_Scoring_ScoreCard.PK, tbl_Scoring_ScoreCard.CtrlNo, " & _
                " tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
                " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.Class, tbl_Scoring_ScoreCard.DDate, tbl_Scoring_ScoreCard.Score, " & _
                " tbl_Scoring_ScoreCard.Front9Gross, tbl_Scoring_ScoreCard.Back9Gross, tbl_Scoring_ScoreCard.GrossPoints, " & _
                " tbl_Scoring_ScoreCard.Front9Net, tbl_Scoring_ScoreCard.Back9Net, tbl_Scoring_ScoreCard.NetPoints, " & _
                " tbl_Scoring_ScoreCard.LastModified, tbl_Scoring_ScoreCard.Front9Score, " & _
                " tbl_Scoring_ScoreCard.Back9Score, tbl_Scoring_ScoreCard.LocationKey " & _
                " FROM tbl_Scoring_ScoreCard LEFT OUTER JOIN " & _
                " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") " & _
                " AND (tbl_Scoring_ScoreCard.CtrlNo = '" & strCtrl & "') " & _
                " ORDER BY tbl_Scoring_ScoreCard.CtrlNo"
        Else
            s = "SELECT TOP 1 tbl_Scoring_ScoreCard.PK, tbl_Scoring_ScoreCard.CtrlNo, " & _
                " tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
                " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.Class, tbl_Scoring_ScoreCard.DDate, tbl_Scoring_ScoreCard.Score, " & _
                " tbl_Scoring_ScoreCard.Front9Gross, tbl_Scoring_ScoreCard.Back9Gross, tbl_Scoring_ScoreCard.GrossPoints, " & _
                " tbl_Scoring_ScoreCard.Front9Net, tbl_Scoring_ScoreCard.Back9Net, tbl_Scoring_ScoreCard.NetPoints, " & _
                " tbl_Scoring_ScoreCard.LastModified, tbl_Scoring_ScoreCard.Front9Score, " & _
                " tbl_Scoring_ScoreCard.Back9Score, tbl_Scoring_ScoreCard.LocationKey " & _
                " FROM tbl_Scoring_ScoreCard LEFT OUTER JOIN " & _
                " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") " & _
                " ORDER BY tbl_Scoring_ScoreCard.CtrlNo"
        End If
    Case "is_FIND"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Scoring_ScoreCard.PK, tbl_Scoring_ScoreCard.CtrlNo, " & _
            " tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
            " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.Class, tbl_Scoring_ScoreCard.DDate, tbl_Scoring_ScoreCard.Score, " & _
            " tbl_Scoring_ScoreCard.Front9Gross, tbl_Scoring_ScoreCard.Back9Gross, tbl_Scoring_ScoreCard.GrossPoints, " & _
            " tbl_Scoring_ScoreCard.Front9Net, tbl_Scoring_ScoreCard.Back9Net, tbl_Scoring_ScoreCard.NetPoints, " & _
            " tbl_Scoring_ScoreCard.LastModified, tbl_Scoring_ScoreCard.Front9Score, " & _
            " tbl_Scoring_ScoreCard.Back9Score, tbl_Scoring_ScoreCard.LocationKey " & _
            " FROM tbl_Scoring_ScoreCard LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " WHERE (tbl_Scoring_ScoreCard.PK = " & strCtrl & ") " & _
            " ORDER BY tbl_Scoring_ScoreCard.CtrlNo DESC"
    Case "is_HOME"
        If picPrintStableford.Visible = True Then Exit Function
        If picSearchAdd.Visible = True Then Exit Function
        If picSearch.Visible = True Then Exit Function
        If picPrint.Visible = True Then Exit Function
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Scoring_ScoreCard.PK, tbl_Scoring_ScoreCard.CtrlNo, " & _
            " tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
            " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.Class, tbl_Scoring_ScoreCard.DDate, tbl_Scoring_ScoreCard.Score, " & _
            " tbl_Scoring_ScoreCard.Front9Gross, tbl_Scoring_ScoreCard.Back9Gross, tbl_Scoring_ScoreCard.GrossPoints, " & _
            " tbl_Scoring_ScoreCard.Front9Net, tbl_Scoring_ScoreCard.Back9Net, tbl_Scoring_ScoreCard.NetPoints, " & _
            " tbl_Scoring_ScoreCard.LastModified, tbl_Scoring_ScoreCard.Front9Score, " & _
            " tbl_Scoring_ScoreCard.Back9Score, tbl_Scoring_ScoreCard.LocationKey " & _
            " FROM tbl_Scoring_ScoreCard LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") " & _
            " ORDER BY tbl_Scoring_ScoreCard.CtrlNo"
    Case "is_PAGEUP"
        If picPrintStableford.Visible = True Then Exit Function
        If picSearchAdd.Visible = True Then Exit Function
        If picSearch.Visible = True Then Exit Function
        If picPrint.Visible = True Then Exit Function
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Scoring_ScoreCard.PK, tbl_Scoring_ScoreCard.CtrlNo, " & _
            " tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
            " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.Class, tbl_Scoring_ScoreCard.DDate, tbl_Scoring_ScoreCard.Score, " & _
            " tbl_Scoring_ScoreCard.Front9Gross, tbl_Scoring_ScoreCard.Back9Gross, tbl_Scoring_ScoreCard.GrossPoints, " & _
            " tbl_Scoring_ScoreCard.Front9Net, tbl_Scoring_ScoreCard.Back9Net, tbl_Scoring_ScoreCard.NetPoints, " & _
            " tbl_Scoring_ScoreCard.LastModified, tbl_Scoring_ScoreCard.Front9Score, " & _
            " tbl_Scoring_ScoreCard.Back9Score, tbl_Scoring_ScoreCard.LocationKey " & _
            " FROM tbl_Scoring_ScoreCard LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") " & _
            " AND (tbl_Scoring_ScoreCard.CtrlNo < '" & strCtrl & "') " & _
            " ORDER BY tbl_Scoring_ScoreCard.CtrlNo DESC"
    Case "is_PAGEDOWN"
        If picPrintStableford.Visible = True Then Exit Function
        If picSearchAdd.Visible = True Then Exit Function
        If picSearch.Visible = True Then Exit Function
        If picPrint.Visible = True Then Exit Function
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Scoring_ScoreCard.PK, tbl_Scoring_ScoreCard.CtrlNo, " & _
            " tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
            " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.Class, tbl_Scoring_ScoreCard.DDate, tbl_Scoring_ScoreCard.Score, " & _
            " tbl_Scoring_ScoreCard.Front9Gross, tbl_Scoring_ScoreCard.Back9Gross, tbl_Scoring_ScoreCard.GrossPoints, " & _
            " tbl_Scoring_ScoreCard.Front9Net, tbl_Scoring_ScoreCard.Back9Net, tbl_Scoring_ScoreCard.NetPoints, " & _
            " tbl_Scoring_ScoreCard.LastModified, tbl_Scoring_ScoreCard.Front9Score, " & _
            " tbl_Scoring_ScoreCard.Back9Score, tbl_Scoring_ScoreCard.LocationKey " & _
            " FROM tbl_Scoring_ScoreCard LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") " & _
            " AND (tbl_Scoring_ScoreCard.CtrlNo > '" & strCtrl & "') " & _
            " ORDER BY tbl_Scoring_ScoreCard.CtrlNo "
    Case "is_END"
        If picPrintStableford.Visible = True Then Exit Function
        If picSearchAdd.Visible = True Then Exit Function
        If picSearch.Visible = True Then Exit Function
        If picPrint.Visible = True Then Exit Function
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Scoring_ScoreCard.PK, tbl_Scoring_ScoreCard.CtrlNo, " & _
            " tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
            " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.Class, tbl_Scoring_ScoreCard.DDate, tbl_Scoring_ScoreCard.Score, " & _
            " tbl_Scoring_ScoreCard.Front9Gross, tbl_Scoring_ScoreCard.Back9Gross, tbl_Scoring_ScoreCard.GrossPoints, " & _
            " tbl_Scoring_ScoreCard.Front9Net, tbl_Scoring_ScoreCard.Back9Net, tbl_Scoring_ScoreCard.NetPoints, " & _
            " tbl_Scoring_ScoreCard.LastModified, tbl_Scoring_ScoreCard.Front9Score, " & _
            " tbl_Scoring_ScoreCard.Back9Score, tbl_Scoring_ScoreCard.LocationKey " & _
            " FROM tbl_Scoring_ScoreCard LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") " & _
            " ORDER BY tbl_Scoring_ScoreCard.CtrlNo DESC"
    Case Else: Exit Function
End Select
'If CStr(s) = "" Then Exit Function
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    LocationKey = rs!LocationKey
    txtCtrl.Text = rs!CtrlNo
    txtDate.Text = Format(rs!dDate, "mm/dd/yyyy")
    txtPlayer.Text = rs!PlayerName
    If TeamAverage = 2 Then
        Label13.Caption = "INDEX"
        txtHandicap.Text = ""
        txtClass.Text = ""
        t = "SELECT tbl_Scoring_PlayerName.* " & _
            " FROM tbl_Scoring_PlayerName " & _
            " WHERE (PK = " & rs!PlayerKey & ")"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            txtHandicap.Text = rt!iIndex
            u = "SELECT Class " & _
                " From tbl_Scoring_TournamentInfo_Index " & _
                " WHERE (TournamentKey = " & TournamentKey & ") " & _
                " AND (HFrom <= " & rt!iIndex & ") " & _
                " AND (HTo >= " & rt!iIndex & ")"
            If ru.State = adStateOpen Then ru.Close
            ru.Open u, ConnOmega
            If ru.RecordCount > 0 Then
                txtClass.Text = ru!Class
            Else
                txtClass.Text = ""
            End If
            ru.Close
        End If
    Else
        Label13.Caption = "HANDICAP"
        txtHandicap.Text = rs!Handicap
        If Trim(rs!Class) = "" Then
            u = "SELECT Class " & _
                " From tbl_Scoring_TournamentInfo_Class " & _
                " WHERE (TournamentKey = " & TournamentKey & ") " & _
                " AND (HFrom <= " & rs!Handicap & ") " & _
                " AND (HTo >= " & rs!Handicap & ")"
            If ru.State = adStateOpen Then ru.Close
            ru.Open u, ConnOmega
            If ru.RecordCount > 0 Then
                txtClass.Text = ru!Class
            Else
                txtClass.Text = ""
            End If
            ru.Close
        Else
            txtClass.Text = rs!Class
        End If
    End If
    
    txtSGrossF.Text = rs!Front9Gross
    txtSGrossB.Text = rs!Back9Gross
    txtSGrossTot.Text = rs!GrossPoints
    txtSNetF.Text = rs!Front9Net
    txtSNetB.Text = rs!Back9Net
    txtSNetTot.Text = rs!NetPoints
    
    txtGrossScoreF.Text = rs!Front9Score
    txtGrossScoreB.Text = rs!Back9Score
    txtGrossScoreTot.Text = rs!Score
    txtGrossPtsF.Text = rs!Front9Gross
    txtGrossPtsB.Text = rs!Back9Gross
    txtGrossPtsTot.Text = rs!GrossPoints
    txtNetPtsF.Text = rs!Front9Net
    txtNetPtsB.Text = rs!Back9Net
    txtNetPtsTot.Text = rs!NetPoints
    
    TeamTmp = 0
    txtTeamName.Text = ""
    t = "SELECT tbl_Scoring_Team_Detail.TeamKey, tbl_Scoring_Team.TeamName " & _
        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
        " tbl_Scoring_Team ON tbl_Scoring_Team_Detail.TeamKey = tbl_Scoring_Team.PK " & _
        " WHERE (tbl_Scoring_Team_Detail.PlayerKey = " & rs!PlayerKey & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        TeamTmp = rt!TeamKey
        txtTeamName.Text = rt!TeamName
    End If
    rt.Close
    
    lstTeamMates.ListItems.Clear
    t = "SELECT TeamKey " & _
        " From tbl_Scoring_Team_Detail " & _
        " WHERE (PlayerKey = " & rs!PlayerKey & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        u = "SELECT tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
            " tbl_Scoring_PlayerName.MiddleName, " & _
            " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.Score) " & _
            " From tbl_Scoring_ScoreCard " & _
            " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS Score, " & _
            " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.GrossPoints) " & _
            " From tbl_Scoring_ScoreCard " & _
            " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS GrossPts, " & _
            " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.NetPoints) " & _
            " From tbl_Scoring_ScoreCard " & _
            " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS NetPoints " & _
            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " Where (tbl_Scoring_Team_Detail.TeamKey = " & rt!TeamKey & ") " & _
            " And (tbl_Scoring_Team_Detail.PlayerKey <> " & rs!PlayerKey & ") " & _
            " Order By ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.GrossPoints) " & _
            " From tbl_Scoring_ScoreCard " & _
            " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) DESC"
        If ru.State = adStateOpen Then ru.Close
        ru.Open u, ConnOmega
        While Not ru.EOF
            Set x = lstTeamMates.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = Trim(ru!LastName) & ",  " & Trim(ru!FirstName) & IIf(Trim(ru!MiddleName) = "", "", "  " & ru!MiddleName)
            x.SubItems(2) = ru!Score
            x.SubItems(3) = ru!GrossPts
            x.SubItems(4) = ru!NetPoints
            ru.MoveNext
        Wend
        ru.Close
        rt.MoveNext
    Wend
    rt.Close
        
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", "Last Modified : " & rs!LastModified)
    
    SaveSetting App.EXEName, "ScoreCardControlAll", "ScoreCardControlAll", rs!CtrlNo
    
    txtLocation.Text = ""
    t = "SELECT tbl_Scoring_Location.* " & _
        " FROM tbl_Scoring_Location " & _
        " WHERE (PK = " & rs!LocationKey & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtLocation.Text = rt!ScoringLocation
    End If
    rt.Close
    
    LOAD_CARD_LOCATION rs!LocationKey, rs!dDate, FGrid
    
    i = -1
    t = "SELECT Par, Handicap, Score, Gross, Net " & _
        " From tbl_Scoring_ScoreCard_Detail " & _
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
    
End If
'rs.Close
If rs.State = adStateOpen Then rs.Close
End Function

Private Function PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If picPrintStableford.Visible = True Then Exit Function
If picSearchAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
If picPrint.Visible = True Then Exit Function
If AccessRights("Scoring Score Card", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
If CHECK_TOURNAMENT_STATUS(TournamentKey) <> 0 Then MsgBox "Tournament was already locked!               ", vbCritical, "Error...": Exit Function
'picMain.Enabled = False
'picToolbar.Enabled = False
'picSearchAdd.ZOrder 0
'txtSearchAdd.Text = ""
'txtDateAdd.Text = Format(FormatDateTime(Date, vbShortDate), "mm/dd/yyyy")
'picSearchAdd.Visible = True
'txtSearchAdd.SetFocus
If CDbl(LocationCnt) = 1 Then
    picMain.Enabled = False
    picToolbar.Enabled = False
    picSearchAdd.ZOrder 0
    txtSearchAdd.Text = ""
    txtDateAdd.Text = Format(FormatDateTime(Date, vbShortDate), "mm/dd/yyyy")
    picSearchAdd.Visible = True
    txtSearchAdd.SetFocus
Else
    For i = 1 To MainFormPopupF.mnuScoringLocationNameAdd.UBound
        Unload MainFormPopupF.mnuScoringLocationNameAdd(i)
    Next i
    i = -1
    t = "SELECT dbo.tbl_Scoring_TournamentInfo_Location.LocationKey, " & _
        " dbo.tbl_Scoring_Location.ScoringLocation " & _
        " FROM dbo.tbl_Scoring_TournamentInfo_Location LEFT OUTER JOIN " & _
        " dbo.tbl_Scoring_Location ON dbo.tbl_Scoring_TournamentInfo_Location.LocationKey = dbo.tbl_Scoring_Location.PK " & _
        " Where (dbo.tbl_Scoring_TournamentInfo_Location.MasterKey = " & TournamentKey & ") " & _
        " ORDER BY dbo.tbl_Scoring_Location.ScoringLocation"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        i = i + 1
        If i = 0 Then
            MainFormPopupF.mnuScoringLocationNameAdd(i).Caption = rt!ScoringLocation
        Else
            Load MainFormPopupF.mnuScoringLocationNameAdd(i)
            MainFormPopupF.mnuScoringLocationNameAdd(i).Caption = rt!ScoringLocation
        End If
        rt.MoveNext
    Wend
    rt.Close
    PopupMenu MainFormPopupF.mnuScoringLocationAdd, , Toolbar1.Buttons(1).Left, Toolbar1.Buttons(1).Top + Toolbar1.Buttons(1).Height
End If
End Function

Private Function PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If picPrintStableford.Visible = True Then Exit Function
If Statusbar1.Panels(1).Text = "" Then Exit Function
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
For i = 0 To 17
    txtGrossScore_Change (i)
Next i
txtGrossScore(0).SetFocus
End Function

Private Function PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If Statusbar1.Panels(1).Text = "" Then Exit Function
If picPrintStableford.Visible = True Then Exit Function
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
ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
'Select Case ScoringType
'    Case 2  'Modified Stableford
'        ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_ModStableFord WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
'    Case 5  'Modified Molave
'        ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_ModMolave WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
'End Select
CLEARTEXT
BROWSER GetSetting(App.EXEName, "ScoreCardControlAll", "ScoreCardControlAll", ""), "is_PAGEDOWN"
If Trim(txtPlayer.Text) = "" Then BROWSER GetSetting(App.EXEName, "ScoreCardControlAll", "ScoreCardControlAll", ""), "is_HOME"
End Function

Private Function PRESS_F5()
If picPrintStableford.Visible = True Then Exit Function
If picSearchAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
If picPrint.Visible = True Then Exit Function
On Error GoTo PG:

If TRANSACTIONTYPE = is_ADDING Then
    
'    Select Case ScoringType
'        Case 1  'Stableford
            s = "SELECT COUNT(*) AS NoofRec " & _
                " From tbl_Scoring_ScoreCard " & _
                " WHERE (TournamentKey = " & TournamentKey & ") " & _
                " AND (PlayerKey = " & PlayerKey & ")"
'        Case 2  'Modified Stableford
'            s = "SELECT COUNT(*) AS NoofRec " & _
'                " From tbl_Scoring_ScoreCard_ModStableFord " & _
'                " WHERE (TournamentKey = " & TournamentKey & ") " & _
'                " AND (PlayerKey = " & PlayerKey & ")"
'        Case 5  'Modified Molave
'            s = "SELECT COUNT(*) AS NoofRec " & _
'                " From tbl_Scoring_ScoreCard_ModMolave " & _
'                " WHERE (TournamentKey = " & TournamentKey & ") " & _
'                " AND (PlayerKey = " & PlayerKey & ")"
'    End Select
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    TourNoOfPlaysTmp = rs!NoofRec
    rs.Close
    
    If CDbl(TourNoOfPlaysTmp) + 1 > CDbl(DaysPlayerToPlay) Then MsgBox "Number of Plays Exceeded!                  ", vbCritical, "Error...": Exit Function
    
    strCtrlNo = "00000001"
'    Select Case ScoringType
'        Case 1  'Stableford
            s = "SELECT TOP 1 CtrlNo " & _
                " FROM tbl_Scoring_ScoreCard " & _
                " WHERE (TournamentKey = " & TournamentKey & ") " & _
                " ORDER BY CtrlNo DESC"
'        Case 2  'Modified Stableford
'            s = "SELECT TOP 1 CtrlNo " & _
'                " FROM tbl_Scoring_ScoreCard_ModStableFord " & _
'                " WHERE (TournamentKey = " & TournamentKey & ") " & _
'                " ORDER BY CtrlNo DESC"
'        Case 5  'Modified Molave
'            s = "SELECT TOP 1 CtrlNo " & _
'                " FROM tbl_Scoring_ScoreCard_ModMolave " & _
'                " WHERE (TournamentKey = " & TournamentKey & ") " & _
'                " ORDER BY CtrlNo DESC"
'    End Select
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        strCtrlNo = Format(CDbl(rs!CtrlNo) + 1, "0000000#")
    End If
    rs.Close
    
    Do
'        Select Case ScoringType
'            Case 1  'Stableford
                s = "SELECT tbl_Scoring_ScoreCard.* " & _
                    " FROM tbl_Scoring_ScoreCard " & _
                    " WHERE (TournamentKey = " & TournamentKey & ") " & _
                    " AND (CtrlNo = '" & strCtrlNo & "')"
'            Case 2  'Modified Stableford
'                s = "SELECT tbl_Scoring_ScoreCard_ModStableFord.* " & _
'                    " FROM tbl_Scoring_ScoreCard_ModStableFord " & _
'                    " WHERE (TournamentKey = " & TournamentKey & ") " & _
'                    " AND (CtrlNo = '" & strCtrlNo & "')"
'            Case 5  'Modified Molave
'                s = "SELECT tbl_Scoring_ScoreCard_ModMolave.* " & _
'                    " FROM tbl_Scoring_ScoreCard_ModMolave " & _
'                    " WHERE (TournamentKey = " & TournamentKey & ") " & _
'                    " AND (CtrlNo = '" & strCtrlNo & "')"
'        End Select
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If
        rs.Close
        strCtrlNo = Format(CDbl(strCtrlNo) + 1, "0000000#")
    Loop
    
'    Select Case ScoringType
'        Case 1  'Stableford
            ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard " & _
                              " (TournamentKey, PlayerKey, DDate, LastModified, CtrlNo, LocationKey) " & _
                              " VALUES (" & TournamentKey & ", " & PlayerKey & ", " & _
                              " '" & FormatDateTime(txtDate.Text, vbShortDate) & "', " & _
                              " '" & CStr(Now) & " - " & gbl_CompleteName & "', '" & strCtrlNo & "', " & LocationKey & ")"
    
            SCardKey = 0
            s = "SELECT PK " & _
                " FROM tbl_Scoring_ScoreCard " & _
                " WHERE (TournamentKey = " & TournamentKey & ") " & _
                " AND (PlayerKey = " & PlayerKey & ") " & _
                " AND (DDate = '" & FormatDateTime(txtDate.Text, vbShortDate) & "')"
            If rs.State = adStateOpen Then rs.Close
            rs.Open s, ConnOmega
            If rs.RecordCount > 0 Then
                SCardKey = rs!PK
            End If
            rs.Close
'        Case 2  'Modified Stableford
'            ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_ModStableFord " & _
'                              " (TournamentKey, PlayerKey, DDate, LastModified, CtrlNo, LocationKey) " & _
'                              " VALUES (" & TournamentKey & ", " & PlayerKey & ", " & _
'                              " '" & FormatDateTime(txtDate.Text, vbShortDate) & "', " & _
'                              " '" & CStr(Now) & " - " & gbl_CompleteName & "', '" & strCtrlNo & "', " & LocationKey & ")"
'
'            SCardKey = 0
'            s = "SELECT PK " & _
'                " FROM tbl_Scoring_ScoreCard_ModStableFord " & _
'                " WHERE (TournamentKey = " & TournamentKey & ") " & _
'                " AND (PlayerKey = " & PlayerKey & ") " & _
'                " AND (DDate = '" & FormatDateTime(txtDate.Text, vbShortDate) & "')"
'            If rs.State = adStateOpen Then rs.Close
'            rs.Open s, ConnOmega
'            If rs.RecordCount > 0 Then
'                SCardKey = rs!PK
'            End If
'            rs.Close
'        Case 5  'Modified Molave
'            ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_ModMolave " & _
'                              " (TournamentKey, PlayerKey, DDate, LastModified, CtrlNo, LocationKey) " & _
'                              " VALUES (" & TournamentKey & ", " & PlayerKey & ", " & _
'                              " '" & FormatDateTime(txtDate.Text, vbShortDate) & "', " & _
'                              " '" & CStr(Now) & " - " & gbl_CompleteName & "', '" & strCtrlNo & "', " & LocationKey & ")"
'
'            SCardKey = 0
'            s = "SELECT PK " & _
'                " FROM tbl_Scoring_ScoreCard_ModMolave " & _
'                " WHERE (TournamentKey = " & TournamentKey & ") " & _
'                " AND (PlayerKey = " & PlayerKey & ") " & _
'                " AND (DDate = '" & FormatDateTime(txtDate.Text, vbShortDate) & "')"
'            If rs.State = adStateOpen Then rs.Close
'            rs.Open s, ConnOmega
'            If rs.RecordCount > 0 Then
'                SCardKey = rs!PK
'            End If
'            rs.Close
'    End Select
End If

If TRANSACTIONTYPE = is_EDITTING Then
    SCardKey = Statusbar1.Panels(1).Text
    strCtrlNo = Trim(txtCtrl.Text)
'    Select Case ScoringType
'        Case 1  'Stableford
            ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard " & _
                              " SET LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                              " WHERE (PK = " & SCardKey & ")"
'        Case 2  'Modified Stableford
'            ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_ModStableFord " & _
'                              " SET LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
'                              " WHERE (PK = " & SCardKey & ")"
'        Case 5  'Modified Molave
'            ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_ModMolave " & _
'                              " SET LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
'                              " WHERE (PK = " & SCardKey & ")"
'    End Select
End If

If CDbl(SCardKey) <> 0 Then
'    Select Case ScoringType
'        Case 1  'Stableford
            ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_Detail WHERE (ScoreCardKey = " & SCardKey & ")"
'        Case 2  'Modified Stableford
'            ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_ModStableFord_Detail WHERE (ScoreCardKey = " & SCardKey & ")"
'        Case 5  'Modified Molave
'            ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_ModMolave_Detail WHERE (ScoreCardKey = " & SCardKey & ")"
'    End Select
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
'        Select Case ScoringType
'            Case 1  'Stableford
                ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_Detail " & _
                                  " (ScoreCardKey, Hole, Par, Handicap, Score, Gross, Net) " & _
                                  " VALUES (" & SCardKey & ", " & i & ", " & CDbl(dblPar) & ", " & _
                                  " " & CDbl(dblHandicap) & ", " & CDbl(dblScore) & ", " & _
                                  " " & CDbl(dblGross) & ", " & CDbl(dblNet) & ")"
'            Case 2  'Modified Stableford
'                ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_ModStableFord_Detail " & _
'                                  " (ScoreCardKey, Hole, Par, Handicap, Score, Gross, Net) " & _
'                                  " VALUES (" & SCardKey & ", " & i & ", " & CDbl(dblPar) & ", " & _
'                                  " " & CDbl(dblHandicap) & ", " & CDbl(dblScore) & ", " & _
'                                  " " & CDbl(dblGross) & ", " & CDbl(dblNet) & ")"
'            Case 5  'Modified Molave
'                ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_ModMolave_Detail " & _
'                                  " (ScoreCardKey, Hole, Par, Handicap, Score, Gross, Net) " & _
'                                  " VALUES (" & SCardKey & ", " & i & ", " & CDbl(dblPar) & ", " & _
'                                  " " & CDbl(dblHandicap) & ", " & CDbl(dblScore) & ", " & _
'                                  " " & CDbl(dblGross) & ", " & CDbl(dblNet) & ")"
'        End Select
    Next i
End If

CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER strCtrlNo, "is_LOAD"

Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Function
End Function

Private Function PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If picPrintStableford.Visible = True Then Exit Function
If picSearchAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
If picPrint.Visible = True Then Exit Function
picToolbar.Enabled = False
picMain.Enabled = False
cmbLocationSearch.ListIndex = 0
txtSearch.Text = ""
picSearch.ZOrder 0
picSearch.Visible = True
txtSearch.SetFocus
End Function

Private Function PRESS_F9()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If picPrintStableford.Visible = True Then Exit Function
If picSearchAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
If picPrint.Visible = True Then Exit Function
picToolbar.Enabled = False
picMain.Enabled = False
If ScoringType = 1 Then
    cmbReportTypeStableford.ListIndex = -1
    cmbGroupStableford.ListIndex = -1
    cmbDivisionStableford.ListIndex = -1
    cmbDay.ListIndex = -1
    picPrintStableford.ZOrder 0
    picPrintStableford.Visible = True
    cmbReportTypeStableford.SetFocus
Else
    cmbReportType.ListIndex = -1
    cmbGroup.ListIndex = -1
    cmbDivision.ListIndex = -1
    txtTop.Text = "10"
    picPrint.ZOrder 0
    picPrint.Visible = True
    cmbReportType.SetFocus
End If
End Function

Private Function PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picPrintStableford.Visible = True Then cmdCancelPrintStableford_Click: Exit Function
    If picSearchAdd.Visible = True Then cmdCancelAdd_Click: Exit Function
    If picSearch.Visible = True Then cmdCancelSearch_Click: Exit Function
    If picPrint.Visible = True Then cmdCancelPrint_Click: Exit Function
    Unload Me
Else
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "ScoreCardControlAll", "ScoreCardControlAll", ""), "is_LOAD"
    If Trim(txtPlayer.Text) = "" Then BROWSER GetSetting(App.EXEName, "ScoreCardControlAll", "ScoreCardControlAll", ""), "is_HOME"
End If
End Function

Private Sub CLEARTEXT()

For i = 0 To 17
    txtGrossScore(i).Text = ""
    txtGrossPts(i).Text = "0"
    txtNetPts(i).Text = "0"
Next i
txtTeamName.Text = ""
txtCtrl.Text = ""
txtDate.Text = ""
txtPlayer.Text = ""
txtHandicap.Text = ""
txtClass.Text = ""
txtDay.Text = ""
txtDate.Text = ""
txtSGrossF.Text = ""
txtSGrossB.Text = ""
txtSGrossTot.Text = ""
txtSNetF.Text = ""
txtSNetB.Text = ""
txtSNetTot.Text = ""

txtGrossScoreF.Text = "0"
txtGrossPtsF.Text = "0"
txtNetPtsF.Text = "0"

txtGrossScoreB.Text = "0"

txtGrossPtsB.Text = "0"
txtNetPtsB.Text = "0"
txtGrossScoreTot.Text = "0"
txtGrossPtsTot.Text = "0"
txtNetPtsTot.Text = "0"

lstTeamMates.ListItems.Clear
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
End Sub

Private Sub LOCKTEXT(bln As Boolean)
For i = 0 To 17
    txtGrossScore(i).Locked = bln
Next i
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
cmdCancelAdd_Click
End Sub

Private Sub b8TitleBar2_CLoseClick()
cmdCancelSearch_Click
End Sub

Private Sub b8TitleBar3_CLoseClick()
cmdCancelPrint_Click
End Sub


Private Sub b8TitleBar4_CLoseClick()
cmdCancelPrintStableford_Click
End Sub

Private Sub cmbDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKSearch_Click
End Sub

Private Sub cmbDivision_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtTop.SetFocus
End Sub

Private Sub cmbGender_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbGroup.SetFocus
End Sub

Private Sub cmbGroup_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbDivision.SetFocus
End Sub

Private Sub cmbGroupStableford_Click()
If cmbGroupStableford.ListIndex = -1 Then Exit Sub
If cmbGroupStableford.ListIndex = 0 Then
    cmbDivisionStableford.Clear
    cmbDivisionStableford.AddItem "ALL"
    If CDbl(TeamAverage) = 2 Then
        u = "SELECT Class" & _
            " From tbl_Scoring_TournamentInfo_Index " & _
            " Where (TournamentKey = " & TournamentKey & ") " & _
            " ORDER BY Class"
    Else
        u = "SELECT Class" & _
            " From tbl_Scoring_TournamentInfo_Class " & _
            " Where (TournamentKey = " & TournamentKey & ") " & _
            " ORDER BY Class"
    End If
    If ru.State = adStateOpen Then ru.Close
    ru.Open u, ConnOmega
    While Not ru.EOF
        cmbDivisionStableford.AddItem ru!Class
        ru.MoveNext
    Wend
    ru.Close
    
    cmbDay.Clear
    cmbDay.AddItem "ALL"
    cmbDay.ItemData(cmbDay.NewIndex) = 0
    u = "SELECT dbo.tbl_Scoring_TournamentInfo_Location.LocationKey, " & _
        " dbo.tbl_Scoring_Location.ScoringLocation " & _
        " FROM  dbo.tbl_Scoring_TournamentInfo_Location LEFT OUTER JOIN " & _
        " dbo.tbl_Scoring_Location ON dbo.tbl_Scoring_TournamentInfo_Location.LocationKey = dbo.tbl_Scoring_Location.PK " & _
        " WHERE (dbo.tbl_Scoring_TournamentInfo_Location.MasterKey = " & TournamentKey & ")"
    If ru.State = adStateOpen Then ru.Close
    ru.Open u, ConnOmega
    While Not ru.EOF
        cmbDay.AddItem ru!ScoringLocation
        cmbDay.ItemData(cmbDay.NewIndex) = ru!LocationKey
        ru.MoveNext
    Wend
    ru.Close
Else
    cmbDivisionStableford.Clear
    cmbDivisionStableford.AddItem "ALL"
    u = "SELECT Class" & _
        " From tbl_Scoring_TournamentInfo_Class " & _
        " Where (TournamentKey = " & TournamentKey & ") " & _
        " ORDER BY Class"
    If ru.State = adStateOpen Then ru.Close
    ru.Open u, ConnOmega
    While Not ru.EOF
        cmbDivisionStableford.AddItem ru!Class
        ru.MoveNext
    Wend
    ru.Close
    
    cmbDay.Clear
    cmbDay.AddItem "1"
    cmbDay.AddItem "2"
End If
End Sub

Private Sub cmbReportType_Click()
If cmbReportType.ListIndex = -1 Then Exit Sub
If cmbReportType.ListIndex = 0 Then
    With cmbDivision
        .Clear
        .AddItem "-- ALL --"
        s = "SELECT Class " & _
            " From tbl_Scoring_TournamentInfo_Class " & _
            " Where (TournamentKey = " & TournamentKey & ") " & _
            " ORDER BY Class"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        While Not rs.EOF
            .AddItem rs!Class
            rs.MoveNext
        Wend
        rs.Close
    End With
Else
    With cmbDivision
        .Clear
        .AddItem "-- ALL --"
        If TeamAverage = 2 Then
            s = "SELECT Class " & _
                " From tbl_Scoring_TournamentInfo_Index " & _
                " Where (TournamentKey = " & TournamentKey & ") " & _
                " ORDER BY Class"
            If rs.State = adStateOpen Then rs.Close
            rs.Open s, ConnOmega
            While Not rs.EOF
                .AddItem rs!Class
                rs.MoveNext
            Wend
            rs.Close
        Else
            s = "SELECT Class " & _
                " From tbl_Scoring_TournamentInfo_Class " & _
                " Where (TournamentKey = " & TournamentKey & ") " & _
                " ORDER BY Class"
            If rs.State = adStateOpen Then rs.Close
            rs.Open s, ConnOmega
            While Not rs.EOF
                .AddItem rs!Class
                rs.MoveNext
            Wend
            rs.Close
        End If
    End With
End If
End Sub

Private Sub cmbReportType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbGender.SetFocus
End Sub

Private Sub cmbReportTypeStableford_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbGroupStableford.SetFocus
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

Private Sub cmdCancelPrintStableford_Click()
picPrintStableford.Visible = False
picToolbar.Enabled = True
picMain.Enabled = True
End Sub

Private Sub cmdCancelSearch_Click()
picToolbar.Enabled = True
picMain.Enabled = True
picSearch.Visible = False
End Sub

Private Sub cmdOKAdd_Click()
If lstResultAdd.ListIndex = -1 Then Exit Sub
If IsDate(txtDateAdd.Text) = False Then Exit Sub
Array1 = Split(Trim(txtTourDate.Text), " - ", -1, 1)
txtDateAdd.Text = Format(FormatDateTime(txtDateAdd.Text, vbShortDate), "mm/dd/yyyy")
If DateValue(FormatDateTime(txtDateAdd.Text, vbShortDate)) < DateValue(FormatDateTime(Array1(0), vbShortDate)) Then MsgBox "Date Out of Range From the Tournament Date!                     ", vbCritical, "Error...": txtDateAdd.SetFocus: HTEXT txtDateAdd: Exit Sub
If DateValue(FormatDateTime(txtDateAdd.Text, vbShortDate)) > DateValue(FormatDateTime(Array1(1), vbShortDate)) Then MsgBox "Date Out of Range From the Tournament Date!                     ", vbCritical, "Error...": txtDateAdd.SetFocus: HTEXT txtDateAdd: Exit Sub
s = "SELECT COUNT(*) AS NoofRec " & _
    " From tbl_Scoring_ScoreCard " & _
    " WHERE (TournamentKey = " & TournamentKey & ") " & _
    " AND (PlayerKey = " & lstResultAdd.ItemData(lstResultAdd.ListIndex) & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
TourNoOfPlaysTmp = rs!NoofRec
rs.Close

If CDbl(TourNoOfPlaysTmp) + 1 > CDbl(DaysPlayerToPlay) Then MsgBox "Number of Plays Exceeded!                  ", vbCritical, "Error...": Exit Sub

s = "SELECT tbl_Scoring_ScoreCard.* " & _
    " FROM tbl_Scoring_ScoreCard " & _
    " WHERE (TournamentKey = " & TournamentKey & ") " & _
    " AND (PlayerKey = " & lstResultAdd.ItemData(lstResultAdd.ListIndex) & ") " & _
    " AND (DDate = '" & FormatDateTime(txtDateAdd.Text, vbShortDate) & "') " & _
    " AND (LocationKey = " & LocationKey & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    MsgBox "Found Duplicate Entry!                          ", vbCritical, "Error..."
    rs.Close
    Exit Sub
End If
rs.Close

CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
PlayerKey = lstResultAdd.ItemData(lstResultAdd.ListIndex)
txtPlayer.Text = lstResultAdd.List(lstResultAdd.ListIndex)
txtDate.Text = Format(FormatDateTime(txtDateAdd.Text, vbShortDate), "mm/dd/yyyy")

TeamTmp = 0
txtTeamName.Text = ""
s = "SELECT tbl_Scoring_Team_Detail.TeamKey, tbl_Scoring_Team.TeamName " & _
    " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
    " tbl_Scoring_Team ON tbl_Scoring_Team_Detail.TeamKey = tbl_Scoring_Team.PK " & _
    " WHERE (tbl_Scoring_Team_Detail.PlayerKey = " & PlayerKey & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    TeamTmp = rs!TeamKey
    txtTeamName.Text = rs!TeamName
End If
rs.Close
If CDbl(TeamTmp) > 0 Then
    s = "SELECT tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
        " tbl_Scoring_PlayerName.MiddleName, " & _
        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.Score) " & _
        " From tbl_Scoring_ScoreCard " & _
        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS Score, " & _
        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.GrossPoints) " & _
        " From tbl_Scoring_ScoreCard " & _
        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS GrossPts, " & _
        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.NetPoints) " & _
        " From tbl_Scoring_ScoreCard " & _
        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS NetPoints " & _
        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
        " Where (tbl_Scoring_Team_Detail.TeamKey = " & TeamTmp & ") " & _
        " And (tbl_Scoring_Team_Detail.PlayerKey <> " & PlayerKey & ") " & _
        " Order By ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.GrossPoints) " & _
        " From tbl_Scoring_ScoreCard " & _
        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        Set x = lstTeamMates.ListItems.Add()
        x.Text = ""
        x.SubItems(1) = Trim(rs!LastName) & ",  " & Trim(rs!FirstName) & IIf(Trim(rs!MiddleName) = "", "", "  " & rs!MiddleName)
        x.SubItems(2) = rs!Score
        x.SubItems(3) = rs!GrossPts
        x.SubItems(4) = rs!NetPoints
        rs.MoveNext
    Wend
    rs.Close
End If

s = "SELECT HandiCap, Class " & _
    " From tbl_Scoring_PlayerName " & _
    " WHERE (PK = " & lstResultAdd.ItemData(lstResultAdd.ListIndex) & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtHandicap.Text = rs!Handicap
    If Trim(IIf(IsNull(rs!Class), "", rs!Class)) = "" Then
        u = "SELECT Class " & _
            " From tbl_Scoring_TournamentInfo_Class " & _
            " WHERE (TournamentKey = " & TournamentKey & ") " & _
            " AND (HFrom <= " & rs!Handicap & ") " & _
            " AND (HTo >= " & rs!Handicap & ")"
        If ru.State = adStateOpen Then ru.Close
        ru.Open u, ConnOmega
        If ru.RecordCount > 0 Then
            txtClass.Text = ru!Class
        Else
            txtClass.Text = ""
        End If
        ru.Close
    Else
        txtClass.Text = rs!Class
    End If
End If
rs.Close
cmdCancelAdd_Click
txtGrossScore(0).SetFocus

End Sub

Private Sub cmdOKPrint_Click()
If cmbReportType.ListIndex = -1 Then MsgBox "Please Select Report Type!                  ", vbCritical, "Error...": cmbReportType.SetFocus: Exit Sub
If cmbGender.ListIndex = -1 Then MsgBox "Please Select Gender!                          ", vbCritical, "Error...": cmbGender.SetFocus: Exit Sub
If cmbGroup.ListIndex = -1 Then MsgBox "Please Select Score!                     ", vbCritical, "Error...": cmbGroup.SetFocus: Exit Sub
If cmbDivision.ListIndex = -1 Then MsgBox "Please Select Division!                   ", vbCritical, "Error...": cmbDivision.SetFocus: Exit Sub
If RETURNTEXTVALUE(txtTop) <= 0 Then MsgBox "Please Supply a Value Higher than Zero!                  ", vbCritical, "Error...": txtTop.SetFocus: Exit Sub
If cmbReportType.ListIndex = 2 Then Exit Sub
If IsDate(txtDatePrint.Text) = True Then
    Arr = Split(TournamentRange, " - ", -1, 1)
    If DateValue(FormatDateTime(txtDatePrint.Text, vbShortDate)) < DateValue(Arr(0)) Then
        MsgBox "Date out of Range!                          ", vbCritical, "Error...": Exit Sub
    ElseIf DateValue(FormatDateTime(txtDatePrint.Text, vbShortDate)) > DateValue(Arr(1)) Then
        MsgBox "Date out of Range!                          ", vbCritical, "Error...": Exit Sub
    End If
End If

With MainForm.CommonDialog1
    .CancelError = True
    On Error GoTo ErrorHandler
    .DialogTitle = "Save"
    .Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
    .ShowSave
    sFileName = Trim(.Filename)
End With

Screen.MousePointer = vbHourglass
Select Case ScoringType
    Case 1  'Stableford
        TimerReportStableford.Enabled = True
    Case 2  'Modified Stableford
        TimerReportModifiedStableford.Enabled = True
    Case 4  'Molave
        TimerReportMolave.Enabled = True
    Case 5  'Modified Molave
        TimerReportModifiedMolave.Enabled = True
End Select
Screen.MousePointer = vbDefault

Exit Sub
ErrorHandler:
Exit Sub
End Sub

Private Sub cmdOKPrintStableford_Click()
If cmbReportTypeStableford.ListIndex = -1 Then Exit Sub
If cmbReportTypeStableford.ListIndex = 1 Then
    If cmbGroupStableford.ListIndex = -1 Then Exit Sub
    If cmbDivisionStableford.ListIndex = -1 Then Exit Sub
End If
If cmbDay.ListIndex = -1 Then Exit Sub
If cmbReportTypeStableford.ListIndex = 1 Then
    If cmbGroupStableford.ListIndex = 0 Then
        If cmbDay.ListIndex = 0 Then TimerTeamAllStableFord.Enabled = True: Exit Sub
    End If
End If

With MainForm.CommonDialog1
    .CancelError = True
    On Error GoTo ErrorHandler
    .DialogTitle = "Save"
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

Select Case cmbReportTypeStableford.ListIndex
    Case 0  'Result
    
    Case 1  'Summary
        
        Select Case cmbGroupStableford.ListIndex
            Case 0  'Team
                
                Screen.MousePointer = vbHourglass
                TableName = "tmp_" & gbl_UserName & "_Report"
                Columns = ""
                Columns = Columns & "|Sorting:int:NOT NULL:DEFAULT(0)"
                Columns = Columns & "|TeamKey:int:NOT NULL"
                Columns = Columns & "|TeamID:varchar:(50):NOT NULL:DEFAULT('')"
                Columns = Columns & "|TeamName:varchar:(50):NOT NULL:DEFAULT('')"
                Columns = Columns & "|TotalHDCP:float:NOT NULL:DEFAULT(0)"
                Columns = Columns & "|TeamHDCP:float:NOT NULL:DEFAULT(0)"
                Columns = Columns & "|TeamClass:varchar:(5):NOT NULL:DEFAULT('')"
                Columns = Columns & "|CntBackPlayer:float:NOT NULL:DEFAULT(0)"
                Columns = Columns & "|Eagle:float:NOT NULL:DEFAULT(0)"
                Columns = Columns & "|GrossPts:float:NOT NULL:DEFAULT(0)"
                
                Clustered = ""
                Clustered = Clustered & "|Sorting"
                
                ColumnsDet = ""
                ColumnsDet = ColumnsDet & "|PlayerName:varchar:(100):NOT NULL:DEFAULT('')"
                ColumnsDet = ColumnsDet & "|HDCP:float:NOT NULL:DEFAULT(0)"
                ColumnsDet = ColumnsDet & "|GrossPts:float:NOT NULL:DEFAULT(0)"
                
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
                
                cnt = 0
                If Trim(cmbDivisionStableford.List(cmbDivisionStableford.ListIndex)) = "ALL" Then
                    If TeamAverage = 1 Then
                        s = "SELECT PK, (CASE LTRIM(RTRIM(TeamName)) WHEN '' then TeamID ELSE LTRIM(RTRIM(TeamName)) END) as TeamName, TeamHDCP, " & _
                            " TeamID, (SELECT tbl_Scoring_TournamentInfo_Class.Class " & _
                            " From tbl_Scoring_TournamentInfo_Class " & _
                            " WHERE (tbl_Scoring_TournamentInfo_Class.TournamentKey = " & TournamentKey & ") " & _
                            " AND (tbl_Scoring_TournamentInfo_Class.HFrom <= tbl_Scoring_Team.TeamHDCP) " & _
                            " AND (tbl_Scoring_TournamentInfo_Class.HTo >= tbl_Scoring_Team.TeamHDCP)) AS Class " & _
                            " From tbl_Scoring_Team " & _
                            " WHERE (TournamentKey = " & TournamentKey & ") "
                    Else
                        s = "SELECT PK, (CASE LTRIM(RTRIM(TeamName)) WHEN '' then TeamID ELSE LTRIM(RTRIM(TeamName)) END) as TeamName, TeamIndex as TeamHDCP, " & _
                            " TeamID, (SELECT tbl_Scoring_TournamentInfo_Index.Class " & _
                            " From tbl_Scoring_TournamentInfo_Index " & _
                            " WHERE (tbl_Scoring_TournamentInfo_Index.TournamentKey = " & TournamentKey & ") " & _
                            " AND (tbl_Scoring_TournamentInfo_Index.HFrom <= tbl_Scoring_Team.TeamIndex) " & _
                            " AND (tbl_Scoring_TournamentInfo_Index.HTo >= tbl_Scoring_Team.TeamIndex)) AS Class " & _
                            " From tbl_Scoring_Team " & _
                            " WHERE (TournamentKey = " & TournamentKey & ") "
                    End If
                Else
                    If TeamAverage = 1 Then
                        s = "SELECT PK, (CASE LTRIM(RTRIM(TeamName)) WHEN '' then TeamID ELSE LTRIM(RTRIM(TeamName)) END) as TeamName, TeamHDCP, " & _
                            " TeamID, (SELECT tbl_Scoring_TournamentInfo_Class.Class " & _
                            " From tbl_Scoring_TournamentInfo_Class " & _
                            " WHERE (tbl_Scoring_TournamentInfo_Class.TournamentKey = " & TournamentKey & ") " & _
                            " AND (tbl_Scoring_TournamentInfo_Class.HFrom <= tbl_Scoring_Team.TeamHDCP) " & _
                            " AND (tbl_Scoring_TournamentInfo_Class.HTo >= tbl_Scoring_Team.TeamHDCP)) AS Class " & _
                            " From tbl_Scoring_Team " & _
                            " WHERE (TournamentKey = " & TournamentKey & ") " & _
                            " AND ((SELECT tbl_Scoring_TournamentInfo_Class.Class " & _
                            " From tbl_Scoring_TournamentInfo_Class " & _
                            " WHERE (tbl_Scoring_TournamentInfo_Class.TournamentKey = " & TournamentKey & ") " & _
                            " AND (tbl_Scoring_TournamentInfo_Class.HFrom <= tbl_Scoring_Team.TeamHDCP) " & _
                            " AND (tbl_Scoring_TournamentInfo_Class.HTo >= tbl_Scoring_Team.TeamHDCP)) = '" & cmbDivisionStableford.List(cmbDivisionStableford.ListIndex) & "')"
                    Else
                        s = "SELECT PK, (CASE LTRIM(RTRIM(TeamName)) WHEN '' then TeamID ELSE LTRIM(RTRIM(TeamName)) END) as TeamName, TeamIndex as TeamHDCP, " & _
                            " TeamID, (SELECT tbl_Scoring_TournamentInfo_Index.Class " & _
                            " From tbl_Scoring_TournamentInfo_Index " & _
                            " WHERE (tbl_Scoring_TournamentInfo_Index.TournamentKey = " & TournamentKey & ") " & _
                            " AND (tbl_Scoring_TournamentInfo_Index.HFrom <= tbl_Scoring_Team.TeamIndex) " & _
                            " AND (tbl_Scoring_TournamentInfo_Index.HTo >= tbl_Scoring_Team.TeamIndex)) AS Class " & _
                            " From tbl_Scoring_Team " & _
                            " WHERE (TournamentKey = " & TournamentKey & ") " & _
                            " AND ((SELECT tbl_Scoring_TournamentInfo_Index.Class " & _
                            " From tbl_Scoring_TournamentInfo_Index " & _
                            " WHERE (tbl_Scoring_TournamentInfo_Index.TournamentKey = " & TournamentKey & ") " & _
                            " AND (tbl_Scoring_TournamentInfo_Index.HFrom <= tbl_Scoring_Team.TeamIndex) " & _
                            " AND (tbl_Scoring_TournamentInfo_Index.HTo >= tbl_Scoring_Team.TeamIndex)) = '" & cmbDivisionStableford.List(cmbDivisionStableford.ListIndex) & "')"
                    End If
                End If
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                While Not rs.EOF
                    DoEvents
                    cnt = cnt + 1
                    ConnOmega.Execute "INSERT INTO " & TableName & " " & _
                                      " (" & sMasterFields & ") " & _
                                      " VALUES (0, " & rs!PK & ", " & _
                                      " '" & rs!TeamID & "', " & _
                                      " '" & FORMATSQL(rs!TeamName) & "', " & _
                                      " 0, " & rs!TeamHDCP & ", " & _
                                      " '" & rs!Class & "', 0, " & _
                                      " 0, 0)"
                    
                    MasterKey = 0
                    t = "SELECT PK " & _
                        " FROM " & TableName & " " & _
                        " WHERE (TeamID = '" & rs!TeamID & "')"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    If rt.RecordCount > 0 Then
                        MasterKey = rt!PK
                    End If
                    rt.Close
                    
                    j = 0
                    't = "SELECT tbl_Scoring_Team_Detail.PlayerKey, tbl_Scoring_PlayerName.LastName, " & _
                        " tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, " & _
                        " tbl_Scoring_PlayerName.HandiCap, IsNull(tbl_Scoring_ScoreCard.GrossPoints, 0) AS GrossPoints " & _
                        " FROM tbl_Scoring_PlayerName RIGHT OUTER JOIN " & _
                        " tbl_Scoring_Team_Detail ON tbl_Scoring_PlayerName.PK = tbl_Scoring_Team_Detail.PlayerKey LEFT OUTER JOIN " & _
                        " tbl_Scoring_ScoreCard ON tbl_Scoring_PlayerName.PK = tbl_Scoring_ScoreCard.PlayerKey " & _
                        " WHERE (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                        " ORDER BY ISNULL(tbl_Scoring_ScoreCard.GrossPoints, 0) DESC"
                    
                    t = "SELECT tbl_Scoring_Team_Detail.PlayerKey, tbl_Scoring_PlayerName.LastName, " & _
                        " tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, " & _
                        " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_Team_Detail.TeamKey " & _
                        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                        " WHERE (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ")"
                    
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        j = j + 1
                        
                        ConnOmega.Execute "INSERT INTO " & DetailTableName & " " & _
                                          " (" & sDetailFields & ") " & _
                                          " VALUES (" & MasterKey & ", " & j & ", " & _
                                          " '" & FORMATSQL(rt!LastName & ",  " & rt!FirstName & IIf(Trim(rt!MiddleName) = "", "", "  " & rt!MiddleName)) & "', " & _
                                          " " & CDbl(rt!Handicap) & ", " & _
                                          " 0)"
                        
                        u = "SELECT TOP 1 GrossPoints " & _
                            " FROM tbl_Scoring_ScoreCard " & _
                            " WHERE (PlayerKey = " & rt!PlayerKey & ") " & _
                            " AND (LocationKey = " & cmbDay.ItemData(cmbDay.ListIndex) & ")"
                        If ru.State = adStateOpen Then ru.Close
                        ru.Open u, ConnOmega
                        If ru.RecordCount > 0 Then
                            ConnOmega.Execute "UPDATE " & DetailTableName & " " & _
                                              " SET GrossPts = " & ru!GrossPoints & " " & _
                                              " WHERE (MasterKey = " & MasterKey & ") " & _
                                              " AND (Line = " & j & ")"
                        End If
                        ru.Close
'
'                        If cmbDay.ListIndex = 0 Then
'                            u = "SELECT TOP 1 GrossPoints " & _
'                                " FROM tbl_Scoring_ScoreCard " & _
'                                " WHERE (PlayerKey = " & rt!PlayerKey & ") " & _
'                                " ORDER BY DDate"
'                            If ru.State = adStateOpen Then ru.Close
'                            ru.Open u, ConnOmega
'                            If ru.RecordCount > 0 Then
'                                ConnOmega.Execute "UPDATE " & DetailTableName & " " & _
'                                                  " SET GrossPts = " & ru!GrossPoints & " " & _
'                                                  " WHERE (MasterKey = " & MasterKey & ") " & _
'                                                  " AND (Line = " & j & ")"
'                            End If
'                            ru.Close
'                        Else
'                            u = "SELECT TOP 1 GrossPoints " & _
'                                " FROM tbl_Scoring_ScoreCard " & _
'                                " WHERE (PlayerKey = " & rt!PlayerKey & ") " & _
'                                " ORDER BY DDate DESC"
'                            If ru.State = adStateOpen Then ru.Close
'                            ru.Open u, ConnOmega
'                            If ru.RecordCount > 0 Then
'                                ConnOmega.Execute "UPDATE " & DetailTableName & " " & _
'                                                  " SET GrossPts = " & ru!GrossPoints & " " & _
'                                                  " WHERE (MasterKey = " & MasterKey & ") " & _
'                                                  " AND (Line = " & j & ")"
'                            End If
'                            ru.Close
'                        End If
                        
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    dblTotalHDCP = 0
                    t = "SELECT TOP " & HandicapDivisor & " * " & _
                        " FROM " & DetailTableName & " " & _
                        " WHERE (MasterKey = " & MasterKey & ") " & _
                        " ORDER BY HDCP"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        dblTotalHDCP = dblTotalHDCP + CDbl(rt!HDCP)
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    dblGrossPts = 0
                    t = "SELECT TOP " & TeamPlayer2Cnt & " * " & _
                        " FROM " & DetailTableName & " " & _
                        " WHERE (MasterKey = " & MasterKey & ") " & _
                        " ORDER BY GrossPts DESC"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        dblGrossPts = dblGrossPts + CDbl(rt!GrossPts)
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    dblCntBackPlayer = 0
                    t = "SELECT TOP 1 * " & _
                        " FROM " & DetailTableName & " " & _
                        " WHERE (MasterKey = " & MasterKey & ") " & _
                        " ORDER BY GrossPts"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        dblCntBackPlayer = CDbl(rt!GrossPts)
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ConnOmega.Execute "UPDATE " & TableName & " " & _
                                      " SET TotalHDCP = " & dblTotalHDCP & ", " & _
                                      " GrossPts = " & dblGrossPts & ", " & _
                                      " CntBackPlayer = " & dblCntBackPlayer & " " & _
                                      " WHERE (PK = " & MasterKey & ")"
                    
                    UpdateProgress picProgressBar, cnt / rs.RecordCount
                    
                    rs.MoveNext
                Wend
                rs.Close
                
                i = 0
                s = "SELECT * " & _
                    " FROM " & TableName & " " & _
                    " ORDER BY GrossPts DESC, CntBackPlayer DESC"
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
                
                ColTop = 0: RowTop = 0
                ColCount = 0: RowCount = 0
                cnt = 0
                Set xlsApp = CreateObject("Excel.Application")
                With xlsApp
                    .Visible = False
                    
                    .Workbooks.Add
                    .DisplayAlerts = False
                    .Workbooks(1).Sheets(1).Activate
                    .Workbooks(1).Sheets(1).Name = cmbReportTypeStableford.List(cmbReportTypeStableford.ListIndex) & " (" & Replace(cmbDivisionStableford.List(cmbDivisionStableford.ListIndex), "SORT BY ", "") & ")" '"Report"
                    .Workbooks(1).Sheets(2).Delete
                    .Workbooks(1).Sheets(2).Delete
                    
                    With xlsApp.ActiveWorkbook.Sheets(1)
                        s = "SELECT * FROM " & TableName & "" & _
                            " ORDER BY Sorting"
                        If rs.State = adStateOpen Then rs.Close
                        rs.Open s, ConnOmega
                        '== Header
                        RowCount = RowCount + 1
                        ColCount = ColCount + 1
                        strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                        strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 3) > 26, 64 + 1, 64) + rs.Fields.Count - 3)) & CStr(RowCount)
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
                        strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 3) > 26, 64 + 1, 64) + rs.Fields.Count - 3)) & CStr(RowCount)
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
                        strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 3) > 26, 64 + 1, 64) + rs.Fields.Count - 3)) & CStr(RowCount)
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
                        strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 3) > 26, 64 + 1, 64) + rs.Fields.Count - 3)) & CStr(RowCount)
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
                        strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 3) > 26, 64 + 1, 64) + rs.Fields.Count - 3)) & CStr(RowCount)
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
                        strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 3) > 26, 64 + 1, 64) + rs.Fields.Count - 3)) & CStr(RowCount)
                        .Range(strRange, strRange1).Select
                        xlsApp.Selection.Merge
                        strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                        .Range(strRange).Value = TournamentRange
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        .Range(strRange).HorizontalAlignment = 3
                        .Range(strRange).VerticalAlignment = 2
                                        
                                        
                        If Trim(cmbDivisionStableford.List(cmbDivisionStableford.ListIndex)) <> "ALL" Then
                            ColCount = 0
                            RowCount = RowCount + 1
                            ColCount = ColCount + 1
                            strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                            strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 3) > 26, 64 + 1, 64) + rs.Fields.Count - 3)) & CStr(RowCount)
                            .Range(strRange, strRange1).Select
                            xlsApp.Selection.Merge
                            
                            strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                            .Range(strRange).Value = "CLASS " & Trim(cmbDivisionStableford.List(cmbDivisionStableford.ListIndex))
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 10
                            .Range(strRange).Font.Bold = True
                            .Range(strRange).HorizontalAlignment = 3
                            .Range(strRange).VerticalAlignment = 2
                        End If
                       
                        ColCount = 0
                        RowCount = RowCount + 1
                        ColCount = ColCount + 1
                        strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                        strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 3) > 26, 64 + 1, 64) + rs.Fields.Count - 3)) & CStr(RowCount)
                        .Range(strRange, strRange1).Select
                        xlsApp.Selection.Merge
                        
                        strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                        '.Range(strRange).Value = "DAY " & Trim(cmbDay.List(cmbDay.ListIndex))
                        .Range(strRange).Value = "Location : " & cmbDay.List(cmbDay.ListIndex)
                        'cmbDay.ItemData(cmbDay.ListIndex)
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 10
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).HorizontalAlignment = 3
                        .Range(strRange).VerticalAlignment = 2
                            
                        ColCount = 0
                        RowCount = RowCount + 1
                        ColCount = ColCount + 1
                        strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                        strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 3) > 26, 64 + 1, 64) + rs.Fields.Count - 3)) & CStr(RowCount)
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
                        .Range(strRange).Value = "Team Name"
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
                        .Range(strRange).Value = "Gross Pts"
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
                        .Range(strRange).Value = "Gross Pts"
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        .Range(strRange).HorizontalAlignment = 3
                        .Range(strRange).Interior.ColorIndex = 15
                        .Range(strRange).Interior.Pattern = 1 'xlSolid
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
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
                                " ORDER BY HDCP"
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
                            .Range(strRange).Value = rs!TeamName
                            .Range(strRange).Font.Name = "Courier New"
                            .Range(strRange).Font.Size = 9
                            .Range(strRange).Font.Bold = False
                            .Columns(ColCount).ColumnWidth = 30
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
                                " ORDER BY HDCP"
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
                                
                                ColCountDet = ColCountDet + 1
                                strRange = (Chr$(IIf(CDbl(ColCountDet) > 26, 64 + 1, 64) + ColCountDet)) & CStr(RowCountDet)
                                .Range(strRange).Value = rt!GrossPts
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
                            .Range(strRange).Value = rs!GrossPts
                            .Range(strRange).Font.Name = "Courier New"
                            .Range(strRange).Font.Size = 13
                            .Range(strRange).Font.Bold = True
                            .Range(strRange).HorizontalAlignment = 3
                            .Range(strRange).VerticalAlignment = 2
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                            
                            RowCount = RowTo
                            
                            UpdateProgress_Caption "Generating Excel Output", picProgressBar, cnt / rs.RecordCount
                            
                            rs.MoveNext
                        Wend
                        rs.Close
                        
                        .PageSetup.PaperSize = 1 'Letter
                        .PageSetup.Orientation = 2 '2 'LandScape
                        .PageSetup.TopMargin = 3
                        .PageSetup.LeftMargin = 3
                        .PageSetup.RightMargin = 3
                        .PageSetup.BottomMargin = 3
                        .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
                    End With
                    
                    If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
                    .ActiveWorkbook.SaveAs Filename:=WorkbookName
                    .Visible = True
                    Set xlsApp = Nothing
                End With
                picProgress.Visible = False
                picMain.Enabled = True
                picToolbar.Enabled = True
                Screen.MousePointer = vbDefault
                
            Case 1  'Individual
            
                Screen.MousePointer = vbHourglass
                TableName = "tmp_" & gbl_UserName & "_Ind_Report"
                Columns = ""
                Columns = Columns & "|Sorting:int:NOT NULL:DEFAULT(0)"
                Columns = Columns & "|PlayerName:varchar:(50):NOT NULL:DEFAULT('')"
                Columns = Columns & "|HDCP:float:NOT NULL:DEFAULT(0)"
                Columns = Columns & "|Class:varchar:(5):NOT NULL:DEFAULT('')"
                Columns = Columns & "|Front9Pts:float:NOT NULL:DEFAULT(0)"
                Columns = Columns & "|Back9Pts:float:NOT NULL:DEFAULT(0)"
                Columns = Columns & "|GrossPts:float:NOT NULL:DEFAULT(0)"
                
                Clustered = ""
                Clustered = Clustered & "|Sorting"
                
                ColumnsDet = ""
                
                DetailTableName = ""
                CreateTable gbl_Database, TableName, Columns, CStr(Clustered), 0, CStr(DetailTableName), CStr(ColumnsDet)
                
                sMasterFields = ""
                Arr = Split(Columns, "|", -1, 1)
                For i = 1 To UBound(Arr)
                    Arr1 = Split(Arr(i), ":", -1, 1)
                    sMasterFields = sMasterFields & Arr1(0) & ", "
                Next i
                sMasterFields = Mid(Trim(CStr(sMasterFields)), 1, Len(Trim(CStr(sMasterFields))) - 1)
                cnt = 0
                
                If Trim(cmbDivisionStableford.List(cmbDivisionStableford.ListIndex)) = "ALL" Then
                    s = "SELECT PK, LastName, FirstName, MiddleName, HandiCap, " & _
                        " (SELECT tbl_Scoring_TournamentInfo_Class.Class " & _
                        " From tbl_Scoring_TournamentInfo_Class " & _
                        " WHERE (tbl_Scoring_TournamentInfo_Class.TournamentKey = tbl_Scoring_PlayerName.TournamentKey) AND " & _
                        " (tbl_Scoring_TournamentInfo_Class.HFrom <= tbl_Scoring_PlayerName.HandiCap) AND " & _
                        " (tbl_Scoring_TournamentInfo_Class.HTo >= tbl_Scoring_PlayerName.HandiCap)) AS Class, ISNULL " & _
                        " ((SELECT sum(tbl_Scoring_ScoreCard.Front9Gross) AS Front9Gross " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.TournamentKey = tbl_Scoring_PlayerName.TournamentKey) AND " & _
                        " (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK)), 0) AS Front9GrsPts, ISNULL " & _
                        " ((SELECT sum(tbl_Scoring_ScoreCard.Back9Gross) AS Back9Gross " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.TournamentKey = tbl_Scoring_PlayerName.TournamentKey) AND " & _
                        " (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK)), 0) AS Back9GrsPts, ISNULL " & _
                        " ((SELECT sum(tbl_Scoring_ScoreCard.GrossPoints) AS GrossPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.TournamentKey = tbl_Scoring_PlayerName.TournamentKey) AND " & _
                        " (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK)), 0) AS GrossPoints " & _
                        " From tbl_Scoring_PlayerName " & _
                        " WHERE (TournamentKey = " & TournamentKey & ") "
                Else
                    s = "SELECT PK, LastName, FirstName, MiddleName, HandiCap, " & _
                        " (SELECT tbl_Scoring_TournamentInfo_Class.Class " & _
                        " From tbl_Scoring_TournamentInfo_Class " & _
                        " WHERE (tbl_Scoring_TournamentInfo_Class.TournamentKey = tbl_Scoring_PlayerName.TournamentKey) AND " & _
                        " (tbl_Scoring_TournamentInfo_Class.HFrom <= tbl_Scoring_PlayerName.HandiCap) AND " & _
                        " (tbl_Scoring_TournamentInfo_Class.HTo >= tbl_Scoring_PlayerName.HandiCap)) AS Class, ISNULL " & _
                        " ((SELECT sum(tbl_Scoring_ScoreCard.Front9Gross) AS Front9Gross " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.TournamentKey = tbl_Scoring_PlayerName.TournamentKey) AND " & _
                        " (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK)), 0) AS Front9GrsPts, ISNULL " & _
                        " ((SELECT sum(tbl_Scoring_ScoreCard.Back9Gross) AS Back9Gross " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.TournamentKey = tbl_Scoring_PlayerName.TournamentKey) AND " & _
                        " (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK)), 0) AS Back9GrsPts, ISNULL " & _
                        " ((SELECT sum(tbl_Scoring_ScoreCard.GrossPoints) AS GrossPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.TournamentKey = tbl_Scoring_PlayerName.TournamentKey) AND " & _
                        " (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK)), 0) AS GrossPoints " & _
                        " From tbl_Scoring_PlayerName " & _
                        " WHERE (TournamentKey = " & TournamentKey & ") " & _
                        " AND ((SELECT tbl_Scoring_TournamentInfo_Class.Class " & _
                        " From tbl_Scoring_TournamentInfo_Class " & _
                        " WHERE (tbl_Scoring_TournamentInfo_Class.TournamentKey = tbl_Scoring_PlayerName.TournamentKey) AND " & _
                        " (tbl_Scoring_TournamentInfo_Class.HFrom <= tbl_Scoring_PlayerName.HandiCap) AND " & _
                        " (tbl_Scoring_TournamentInfo_Class.HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivisionStableford.List(cmbDivisionStableford.ListIndex) & "')"
                End If
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                While Not rs.EOF
                    DoEvents
                    cnt = cnt + 1
                    
                    strPlayerName = Trim(rs!LastName) & ",  " & Trim(rs!FirstName) & IIf(Trim(rs!MiddleName) = "", "", "  " & rs!MiddleName)
                    
                    sGrossPts = ""
                    t = "SELECT GrossPoints " & _
                        " From dbo.tbl_Scoring_ScoreCard " & _
                        " Where (TournamentKey = " & TournamentKey & ") " & _
                        " And (PlayerKey = " & rs!PK & ") " & _
                        " And (LocationKey = 1)"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    If rt.RecordCount > 0 Then
                        sGrossPts = rt!GrossPoints
                    Else
                        sGrossPts = "0"
                    End If
                    rt.Close
                    t = "SELECT GrossPoints " & _
                        " From dbo.tbl_Scoring_ScoreCard " & _
                        " Where (TournamentKey = " & TournamentKey & ") " & _
                        " And (PlayerKey = " & rs!PK & ") " & _
                        " And (LocationKey = 2)"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    If rt.RecordCount > 0 Then
                        sGrossPts = sGrossPts & "|" & rt!GrossPoints
                    Else
                        sGrossPts = sGrossPts & "|0"
                    End If
                    rt.Close
                    
                    'sGrossPts = Mid(sGrossPts, 2, Len(sGrossPts))
                    Arr2 = Split(sGrossPts, "|", -1, 1)
'                    For i = 0 To UBound(Arr2)
                        ConnOmega.Execute "INSERT INTO " & TableName & " " & _
                                          " (" & sMasterFields & ") " & _
                                          " VALUES (0, '" & FORMATSQL(CStr(strPlayerName)) & "', " & _
                                          " " & rs!Handicap & ", " & _
                                          " '" & rs!Class & "', " & _
                                          " " & CDbl(Arr2(0)) & ", " & _
                                          " " & CDbl(Arr2(1)) & ", " & _
                                          " " & CDbl(rs!GrossPoints) & ")"
'                    Next i
                    
                    UpdateProgress picProgressBar, cnt / rs.RecordCount
                    
                    rs.MoveNext
                Wend
                rs.Close
                
                i = 0
                s = "SELECT * " & _
                    " FROM " & TableName & " " & _
                    " ORDER BY GrossPts DESC, Back9Pts DESC, Front9Pts DESC"
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
                
                ColTop = 0: RowTop = 0
                ColCount = 0: RowCount = 0
                cnt = 0
                Set xlsApp = CreateObject("Excel.Application")
                With xlsApp
                    .Visible = False
                    
                    .Workbooks.Add
                    .DisplayAlerts = False
                    .Workbooks(1).Sheets(1).Activate
                    .Workbooks(1).Sheets(1).Name = cmbReportTypeStableford.List(cmbReportTypeStableford.ListIndex) & " (" & cmbGroupStableford.List(cmbGroupStableford.ListIndex) & ") (" & Replace(cmbDivisionStableford.List(cmbDivisionStableford.ListIndex), "SORT BY ", "") & ")"    '"Report"
                    .Workbooks(1).Sheets(2).Delete
                    .Workbooks(1).Sheets(2).Delete
                    With xlsApp.ActiveWorkbook.Sheets(1)
                        s = "SELECT * FROM " & TableName & "" & _
                            " ORDER BY Sorting"
                        If rs.State = adStateOpen Then rs.Close
                        rs.Open s, ConnOmega
                        '== Header
                        RowCount = RowCount + 1
                        ColCount = ColCount + 1
                        strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                        strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 2) > 26, 64 + 1, 64) + rs.Fields.Count - 2)) & CStr(RowCount)
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
                        strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 2) > 26, 64 + 1, 64) + rs.Fields.Count - 2)) & CStr(RowCount)
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
                        strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 2) > 26, 64 + 1, 64) + rs.Fields.Count - 2)) & CStr(RowCount)
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
                        strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 2) > 26, 64 + 1, 64) + rs.Fields.Count - 2)) & CStr(RowCount)
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
                        strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 2) > 26, 64 + 1, 64) + rs.Fields.Count - 2)) & CStr(RowCount)
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
                        strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 2) > 26, 64 + 1, 64) + rs.Fields.Count - 2)) & CStr(RowCount)
                        .Range(strRange, strRange1).Select
                        xlsApp.Selection.Merge
                        strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                        .Range(strRange).Value = TournamentRange
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        .Range(strRange).HorizontalAlignment = 3
                        .Range(strRange).VerticalAlignment = 2
                                        
                        If Trim(cmbDivisionStableford.List(cmbDivisionStableford.ListIndex)) <> "ALL" Then
                            ColCount = 0
                            RowCount = RowCount + 1
                            ColCount = ColCount + 1
                            strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                            strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 3) > 26, 64 + 1, 64) + rs.Fields.Count - 2)) & CStr(RowCount)
                            .Range(strRange, strRange1).Select
                            xlsApp.Selection.Merge
                            
                            strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                            .Range(strRange).Value = "CLASS " & Trim(cmbDivisionStableford.List(cmbDivisionStableford.ListIndex))
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 10
                            .Range(strRange).Font.Bold = True
                            .Range(strRange).HorizontalAlignment = 3
                            .Range(strRange).VerticalAlignment = 2
                        End If
                                        
                        ColCount = 0
                        RowCount = RowCount + 1
                        ColCount = ColCount + 1
                        strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                        strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 2) > 26, 64 + 1, 64) + rs.Fields.Count - 2)) & CStr(RowCount)
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
                        .Range(strRange).Value = "Class"
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
                        '.Range(strRange).Value = "Front 9"
                        .Range(strRange).Value = "Loc 1"
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
                        '.Range(strRange).Value = "Back 9"
                        .Range(strRange).Value = "Loc 2"
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
                        .Range(strRange).Value = "Gross Pts"
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        .Range(strRange).HorizontalAlignment = 3
                        .Range(strRange).Interior.ColorIndex = 15
                        .Range(strRange).Interior.Pattern = 1 'xlSolid
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        HeaderRow = RowCount
                           
                        While Not rs.EOF
                            cnt = cnt + 1
                            ColCount = 0
                            RowCount = RowCount + 1
                            ColCount = ColCount + 1
                            strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                            .Range(strRange).Value = rs!PlayerName
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).HorizontalAlignment = 1
                            .Columns(ColCount).ColumnWidth = 28
                            .Range(strRange).VerticalAlignment = 2
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                            
                            ColCount = ColCount + 1
                            strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                            .Range(strRange).Value = rs!HDCP
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).HorizontalAlignment = 4
                            .Range(strRange).VerticalAlignment = 2
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                            
                            ColCount = ColCount + 1
                            strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                            .Range(strRange).Value = rs!Class
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).HorizontalAlignment = 4
                            .Range(strRange).VerticalAlignment = 2
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                            
                            ColCount = ColCount + 1
                            strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                            .Range(strRange).Value = rs!Front9Pts
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).HorizontalAlignment = 3
                            .Range(strRange).VerticalAlignment = 2
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                            
                            ColCount = ColCount + 1
                            strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                            .Range(strRange).Value = rs!Back9Pts
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).HorizontalAlignment = 3
                            .Range(strRange).VerticalAlignment = 2
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                            
                            ColCount = ColCount + 1
                            strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                            .Range(strRange).Value = rs!GrossPts
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).HorizontalAlignment = 3
                            .Range(strRange).VerticalAlignment = 2
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                            
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
                        .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
                    End With
                    
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
        
End Select

Exit Sub
ErrorHandler:

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
Screen.MousePointer = vbHourglass
s = "SELECT PK, CtrlNo, TournamentKey, LocationKey, PlayerKey, " & _
    " DDate, Front9Score, Back9Score, Front9Gross, Back9Gross, " & _
    " Front9Net, Back9Net, LastModified " & _
    " FROM dbo.tbl_Scoring_ScoreCard_ModMolave"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    t = "SELECT tbl_Scoring_ScoreCard.* " & _
        " FROM tbl_Scoring_ScoreCard " & _
        " WHERE (TournamentKey = " & rs!TournamentKey & ") " & _
        " AND (PlayerKey = " & rs!PlayerKey & ") " & _
        " AND (DDate = '" & FormatDateTime(rs!dDate, vbShortDate) & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount = 0 Then
        strCtrlNo = rs!CtrlNo
        Do
            u = "SELECT tbl_Scoring_ScoreCard.* " & _
                " FROM tbl_Scoring_ScoreCard " & _
                " WHERE (TournamentKey = " & rs!TournamentKey & ") " & _
                " AND (CtrlNo = '" & strCtrlNo & "')"
            If ru.State = adStateOpen Then ru.Close
            ru.Open u, ConnOmega
            If ru.RecordCount = 0 Then
                ru.Close
                Exit Do
            End If
            ru.Close
            strCtrlNo = Format(CDbl(strCtrlNo) + 1, "0000000#")
        Loop
        ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard " & _
                          " (CtrlNo, TournamentKey, LocationKey, PlayerKey, " & _
                          " DDate, Front9Score, Back9Score, Front9Gross, Back9Gross, " & _
                          " Front9Net, Back9Net, LastModified) " & _
                          " VALUES ('" & strCtrlNo & "', " & rs!TournamentKey & ", " & _
                          " " & rs!LocationKey & ", " & rs!PlayerKey & ", " & _
                          " '" & FormatDateTime(rs!dDate, vbShortDate) & "' , " & _
                          " " & rs!Front9Score & ", " & rs!Back9Score & ", " & _
                          " " & rs!Front9Gross & ", " & rs!Back9Gross & ", " & _
                          " " & rs!Front9Net & ", " & rs!Back9Net & ", " & _
                          " '" & rs!LastModified & "')"
    Else
        ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard " & _
                          " SET LocationKey = " & rs!LocationKey & ", " & _
                          " Front9Score = " & rs!Front9Score & ", " & _
                          " Back9Score = " & rs!Back9Score & ", " & _
                          " Front9Gross = " & rs!Front9Gross & ", " & _
                          " Back9Gross = " & rs!Back9Gross & ", " & _
                          " Front9Net = " & rs!Front9Net & ", " & _
                          " Back9Net = " & rs!Back9Net & ", " & _
                          " LastModified = '" & rs!LastModified & "' " & _
                          " WHERE (PK = " & rt!PK & ")"
    End If
    rt.Close
    
    t = "SELECT tbl_Scoring_ScoreCard.* " & _
        " FROM tbl_Scoring_ScoreCard " & _
        " WHERE (TournamentKey = " & rs!TournamentKey & ") " & _
        " AND (PlayerKey = " & rs!PlayerKey & ") " & _
        " AND (DDate = '" & FormatDateTime(rs!dDate, vbShortDate) & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_Detail WHERE (ScoreCardKey = " & rt!PK & ")"
        u = "SELECT ScoreCardKey, Hole, Par, Handicap, Score, Gross, Net " & _
            " FROM tbl_Scoring_ScoreCard_ModMolave_Detail " & _
            " WHERE (ScoreCardKey = " & rs!PK & ")"
        If ru.State = adStateOpen Then ru.Close
        ru.Open u, ConnOmega
        While Not ru.EOF
            ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_Detail " & _
                              " (ScoreCardKey, Hole, Par, Handicap, Score, Gross, Net) " & _
                              " VALUES (" & rt!PK & ", " & ru!Hole & ", " & ru!Par & ", " & _
                              " " & ru!Handicap & ", " & ru!Score & ", " & ru!Gross & ", " & ru!Net & ")"
            ru.MoveNext
        Wend
        ru.Close
    End If
    rt.Close
    
    rs.MoveNext
Wend
rs.Close
Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Screen.MousePointer = vbHourglass
s = "SELECT PK, CtrlNo, TournamentKey, LocationKey, PlayerKey, " & _
    " DDate, Front9Score, Back9Score, Front9Gross, Back9Gross, " & _
    " Front9Net, Back9Net, LastModified " & _
    " FROM dbo.tbl_Scoring_ScoreCard_ModStableFord " & _
    " ORDER BY TournamentKey, CtrlNo "
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    t = "SELECT tbl_Scoring_ScoreCard.* " & _
        " FROM tbl_Scoring_ScoreCard " & _
        " WHERE (TournamentKey = " & rs!TournamentKey & ") " & _
        " AND (PlayerKey = " & rs!PlayerKey & ") " & _
        " AND (DDate = '" & FormatDateTime(rs!dDate, vbShortDate) & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount = 0 Then
        strCtrlNo = rs!CtrlNo
        Do
            u = "SELECT tbl_Scoring_ScoreCard.* " & _
                " FROM tbl_Scoring_ScoreCard " & _
                " WHERE (TournamentKey = " & rs!TournamentKey & ") " & _
                " AND (CtrlNo = '" & strCtrlNo & "')"
            If ru.State = adStateOpen Then ru.Close
            ru.Open u, ConnOmega
            If ru.RecordCount = 0 Then
                ru.Close
                Exit Do
            End If
            ru.Close
            strCtrlNo = Format(CDbl(strCtrlNo) + 1, "0000000#")
        Loop
        
        ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard " & _
                          " (CtrlNo, TournamentKey, LocationKey, PlayerKey, " & _
                          " DDate, Front9Score, Back9Score, Front9Gross, Back9Gross, " & _
                          " Front9Net, Back9Net, LastModified) " & _
                          " VALUES ('" & strCtrlNo & "', " & rs!TournamentKey & ", " & _
                          " " & rs!LocationKey & ", " & rs!PlayerKey & ", " & _
                          " '" & FormatDateTime(rs!dDate, vbShortDate) & "' , " & _
                          " " & rs!Front9Score & ", " & rs!Back9Score & ", " & _
                          " " & rs!Front9Gross & ", " & rs!Back9Gross & ", " & _
                          " " & rs!Front9Net & ", " & rs!Back9Net & ", " & _
                          " '" & rs!LastModified & "')"
    Else
        ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard " & _
                          " SET LocationKey = " & rs!LocationKey & ", " & _
                          " Front9Score = " & rs!Front9Score & ", " & _
                          " Back9Score = " & rs!Back9Score & ", " & _
                          " Front9Gross = " & rs!Front9Gross & ", " & _
                          " Back9Gross = " & rs!Back9Gross & ", " & _
                          " Front9Net = " & rs!Front9Net & ", " & _
                          " Back9Net = " & rs!Back9Net & ", " & _
                          " LastModified = '" & rs!LastModified & "' " & _
                          " WHERE (PK = " & rt!PK & ")"
    End If
    rt.Close
    
    t = "SELECT tbl_Scoring_ScoreCard.* " & _
        " FROM tbl_Scoring_ScoreCard " & _
        " WHERE (TournamentKey = " & rs!TournamentKey & ") " & _
        " AND (PlayerKey = " & rs!PlayerKey & ") " & _
        " AND (DDate = '" & FormatDateTime(rs!dDate, vbShortDate) & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_Detail WHERE (ScoreCardKey = " & rt!PK & ")"
        u = "SELECT ScoreCardKey, Hole, Par, Handicap, Score, Gross, Net " & _
            " FROM tbl_Scoring_ScoreCard_ModStableFord_Detail " & _
            " WHERE (ScoreCardKey = " & rs!PK & ")"
        If ru.State = adStateOpen Then ru.Close
        ru.Open u, ConnOmega
        While Not ru.EOF
            ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_Detail " & _
                              " (ScoreCardKey, Hole, Par, Handicap, Score, Gross, Net) " & _
                              " VALUES (" & rt!PK & ", " & ru!Hole & ", " & ru!Par & ", " & _
                              " " & ru!Handicap & ", " & ru!Score & ", " & ru!Gross & ", " & ru!Net & ")"
            ru.MoveNext
        Wend
        ru.Close
    End If
    rt.Close
    
    rs.MoveNext
Wend
rs.Close
Screen.MousePointer = vbDefault
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
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "ScoreCardControlAll", "ScoreCardControlAll", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "ScoreCardControlAll", "ScoreCardControlAll", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "ScoreCardControlAll", "ScoreCardControlAll", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "ScoreCardControlAll", "ScoreCardControlAll", ""), "is_END"
    Case Else: Exit Sub
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Height = 7695
Me.Width = 13980
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2

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

With cmbLocationSearch
    .Clear
    s = "SELECT dbo.tbl_Scoring_TournamentInfo_Location.LocationKey, " & _
        " dbo.tbl_Scoring_Location.ScoringLocation " & _
        " FROM dbo.tbl_Scoring_TournamentInfo_Location LEFT OUTER JOIN " & _
        " dbo.tbl_Scoring_Location ON dbo.tbl_Scoring_TournamentInfo_Location.LocationKey = dbo.tbl_Scoring_Location.PK " & _
        " Where (dbo.tbl_Scoring_TournamentInfo_Location.MasterKey = " & TournamentKey & ") " & _
        " ORDER BY dbo.tbl_Scoring_TournamentInfo_Location.LocationKey"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        .AddItem rs!ScoringLocation
        .ItemData(.NewIndex) = rs!LocationKey
        rs.MoveNext
    Wend
    rs.Close
End With

With cmbReportType
    .Clear
    .AddItem "INDIVIDUAL"
    .AddItem "TEAM"
    '.AddItem "RESULT"
End With

With cmbGroup
    .Clear
    .AddItem "NET POINTS"
    .AddItem "GROSS POINTS"
    .AddItem "GROSS SCORE"
End With

With cmbGender
    .Clear
    .AddItem "-- ALL --"
    .AddItem "MALE"
    .AddItem "FEMALE"
End With

With cmbReportTypeStableford
    .Clear
    .AddItem "RESULT"
    .AddItem "SUMMARY"
End With

With cmbGroupStableford
    .Clear
    .AddItem "TEAM"
    .AddItem "INDIVIDUAL"
End With

cmbDay.Clear
For i = 1 To DaysPlayerToPlay
    cmbDay.AddItem i
Next i

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

LOAD_CARD dDateEnd, FGrid

CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH

BROWSER GetSetting(App.EXEName, "ScoreCardControlAll", "ScoreCardControlAll", ""), "is_LOAD"
If Trim(txtPlayer.Text) = "" Then BROWSER GetSetting(App.EXEName, "ScoreCardControlAll", "ScoreCardControlAll", ""), "is_HOME"

tmp = SetWindowLong(txtSearchAdd.hwnd, GWL_STYLE, GetWindowLong(txtSearchAdd.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picPrintStableford.Visible = True Then Cancel = -1
If picProgress.Visible = True Then Cancel = -1
If picSearchAdd.Visible = True Then Cancel = -1
If picSearch.Visible = True Then Cancel = -1
If picPrint.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstResult_Click()
If cmbLocationSearch.ListIndex = -1 Then Exit Sub
If lstResult.ListIndex = -1 Then cmbDate.Clear: Exit Sub
cmbDate.Clear
s = "SELECT PK, DDate " & _
    " From tbl_Scoring_ScoreCard " & _
    " Where (TournamentKey = " & TournamentKey & ") " & _
    " And (LocationKey = " & cmbLocationSearch.ItemData(cmbLocationSearch.ListIndex) & ") " & _
    " And (PlayerKey = " & lstResult.ItemData(lstResult.ListIndex) & ") " & _
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

Private Sub lstResult_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbDate.SetFocus
End Sub

Private Sub lstResultAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtDateAdd.SetFocus
End Sub

Private Sub TimerAddLocation_Timer()
TimerAddLocation.Enabled = False

Arr = Split(Trim(txtPath.Text), "\", -1, 1)
Arr1 = Split(CStr(Arr(UBound(Arr))), ".", -1, 1)

Screen.MousePointer = vbHourglass
On Error GoTo PG:
If Arr1(UBound(Arr1)) = "txt" Then    'Text File
    sFileNameMaster = Trim(txtPath.Text)
    'sFileNameDetail = Replace(Trim(txtPath.Text), CStr(Arr(UBound(Arr))), Arr1(0) & "_det.txt")
    
    Open CStr(sFileNameMaster) For Input As #1
        Do Until EOF(1)
            Line Input #1, StrFile
            sFileArr = Split(StrFile, "[", -1, 1)
            Select Case sFileArr(0)
                Case "ScoreCard"
                    sFileArrDet = Split(sFileArr(1), "|", -1, 1)
                    u = "SELECT tbl_Scoring_ScoreCard.* " & _
                        " FROM tbl_Scoring_ScoreCard " & _
                        " WHERE (TournamentKey = " & TournamentKey & ") " & _
                        " AND (PlayerKey = " & sFileArrDet(3) & ") " & _
                        " AND (DDate = '" & FormatDateTime(sFileArrDet(4), vbShortDate) & "')"
                    If ru.State = adStateOpen Then ru.Close
                    ru.Open u, ConnOmega
                    If ru.RecordCount = 0 Then
                        strCtrlNo = "00000001"
                        v = "SELECT TOP 1 CtrlNo " & _
                            " FROM tbl_Scoring_ScoreCard " & _
                            " WHERE (TournamentKey = " & TournamentKey & ") " & _
                            " ORDER BY CtrlNo DESC"
                        If rv.State = adStateOpen Then rv.Close
                        rv.Open v, ConnOmega
                        If rv.RecordCount > 0 Then
                            strCtrlNo = Format(CDbl(rv!CtrlNo) + 1, "0000000#")
                        End If
                        rv.Close
                        
                        Do
                            v = "SELECT tbl_Scoring_ScoreCard.* " & _
                                " FROM tbl_Scoring_ScoreCard " & _
                                " WHERE (TournamentKey = " & TournamentKey & ") " & _
                                " AND (CtrlNo = '" & strCtrlNo & "')"
                            If rv.State = adStateOpen Then rv.Close
                            rv.Open v, ConnOmega
                            If rv.RecordCount = 0 Then
                                rv.Close
                                Exit Do
                            End If
                            rv.Close
                            strCtrlNo = Format(CDbl(strCtrlNo) + 1, "0000000#")
                        Loop
                        
                        ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard " & _
                                          " (TournamentKey, PlayerKey, DDate, LastModified, CtrlNo, LocationKey) " & _
                                          " VALUES (" & TournamentKey & ", " & sFileArrDet(3) & ", " & _
                                          " '" & FormatDateTime(sFileArrDet(4), vbShortDate) & "', " & _
                                          " '" & CStr(Now) & " - " & gbl_CompleteName & "', '" & strCtrlNo & "', " & sFileArrDet(2) & ")"
                        
                        
                        iScoreKey = 0
                        v = "SELECT PK " & _
                            " FROM tbl_Scoring_ScoreCard " & _
                            " WHERE (TournamentKey = " & TournamentKey & ") " & _
                            " AND (PlayerKey = " & sFileArrDet(3) & ") " & _
                            " AND (DDate = '" & FormatDateTime(sFileArrDet(4), vbShortDate) & "')"
                        If rv.State = adStateOpen Then rv.Close
                        rv.Open v, ConnOmega
                        If rv.RecordCount > 0 Then
                            iScoreKey = rv!PK
                        End If
                        rv.Close
                        
                        ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_Detail WHERE (ScoreCardKey = " & iScoreKey & ")"
                        
                    Else
                        
                        iScoreKey = ru!PK
                        strCtrlNo = ru!CtrlNo
                        ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_Detail WHERE (ScoreCardKey = " & iScoreKey & ")"
                    
                    End If
                    ru.Close
                Case "ScoreCardDetail"
                    sFileArrDet = Split(sFileArr(1), "|", -1, 1)
                    If CDbl(iScoreKey) > 0 Then
                        ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_Detail " & _
                                          " (ScoreCardKey, Hole, Par, Handicap, Score, Gross, Net) " & _
                                          " VALUES (" & iScoreKey & ", " & sFileArrDet(1) & ", " & sFileArrDet(2) & ", " & _
                                          " " & sFileArrDet(3) & ", " & sFileArrDet(4) & ", " & sFileArrDet(5) & ", " & _
                                          " " & sFileArrDet(6) & ")"
                    End If
            End Select
        Loop
    Close #1
    
    BROWSER strCtrlNo, "is_LOAD"
    If Trim(txtPlayer.Text) = "" Then BROWSER GetSetting(App.EXEName, "ScoreCardControlAll", "ScoreCardControlAll", ""), "is_HOME"
    
    Screen.MousePointer = vbDefault
Else
    Set cn = New ADODB.Connection
    'cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cn.Provider = "Microsoft.ACE.OLEDB.12.0;"
    cn.ConnectionString = _
        "Data Source= " & Trim(txtPath.Text) & ";" & _
        "Extended Properties=Excel 8.0;"
    cn.CursorLocation = adUseClient
    If cn.State = adStateOpen Then cn.Close
    cn.Open
    
    Set rs = New ADODB.Recordset
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT * FROM [ScoreCard$] ", cn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        If IsNull(rs!PK) = False Then
            u = "SELECT tbl_Scoring_ScoreCard.* " & _
                " FROM tbl_Scoring_ScoreCard " & _
                " WHERE (TournamentKey = " & TournamentKey & ") " & _
                " AND (PlayerKey = " & rs!PlayerKey & ") " & _
                " AND (DDate = '" & FormatDateTime(rs!dDate, vbShortDate) & "')"
            If ru.State = adStateOpen Then ru.Close
            ru.Open u, ConnOmega
            If ru.RecordCount = 0 Then
                strCtrlNo = "00000001"
                v = "SELECT TOP 1 CtrlNo " & _
                    " FROM tbl_Scoring_ScoreCard " & _
                    " WHERE (TournamentKey = " & TournamentKey & ") " & _
                    " ORDER BY CtrlNo DESC"
                If rv.State = adStateOpen Then rv.Close
                rv.Open v, ConnOmega
                If rv.RecordCount > 0 Then
                    strCtrlNo = Format(CDbl(rv!CtrlNo) + 1, "0000000#")
                End If
                rv.Close
                
                Do
                    v = "SELECT tbl_Scoring_ScoreCard.* " & _
                        " FROM tbl_Scoring_ScoreCard " & _
                        " WHERE (TournamentKey = " & TournamentKey & ") " & _
                        " AND (CtrlNo = '" & strCtrlNo & "')"
                    If rv.State = adStateOpen Then rv.Close
                    rv.Open v, ConnOmega
                    If rv.RecordCount = 0 Then
                        rv.Close
                        Exit Do
                    End If
                    rv.Close
                    strCtrlNo = Format(CDbl(strCtrlNo) + 1, "0000000#")
                Loop
                
                ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard " & _
                                  " (TournamentKey, PlayerKey, DDate, LastModified, CtrlNo, LocationKey) " & _
                                  " VALUES (" & TournamentKey & ", " & rs!PlayerKey & ", " & _
                                  " '" & FormatDateTime(rs!dDate, vbShortDate) & "', " & _
                                  " '" & CStr(Now) & " - " & gbl_CompleteName & "', '" & strCtrlNo & "', " & rs!LocationKey & ")"
                
                
                iScoreKey = 0
                v = "SELECT PK " & _
                    " FROM tbl_Scoring_ScoreCard " & _
                    " WHERE (TournamentKey = " & TournamentKey & ") " & _
                    " AND (PlayerKey = " & rs!PlayerKey & ") " & _
                    " AND (DDate = '" & FormatDateTime(rs!dDate, vbShortDate) & "')"
                If rv.State = adStateOpen Then rv.Close
                rv.Open v, ConnOmega
                If rv.RecordCount > 0 Then
                    iScoreKey = rv!PK
                End If
                rv.Close
            Else
                iScoreKey = ru!PK
                strCtrlNo = ru!CtrlNo
            End If
            ru.Close
            If CDbl(iScoreKey) > 0 Then
                ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_Detail WHERE (ScoreCardKey = " & iScoreKey & ")"
                Set rt = New ADODB.Recordset
                If rt.State = adStateOpen Then rt.Close
                rt.Open "SELECT * FROM [ScoreCardDetails$] WHERE (ScoreCardKey = " & rs!PK & ")", cn, adOpenDynamic, adLockOptimistic
                While Not rt.EOF
                    If IsNull(rt!ScoreCardKey) = False Then
                        ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_Detail " & _
                                          " (ScoreCardKey, Hole, Par, Handicap, Score, Gross, Net) " & _
                                          " VALUES (" & iScoreKey & ", " & rt!Hole & ", " & rt!Par & ", " & _
                                          " " & rt!Handicap & ", " & rt!Score & ", " & rt!Gross & ", " & rt!Net & ")"
                    End If
                    rt.MoveNext
                Wend
                rt.Close
            End If
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    BROWSER strCtrlNo, "is_LOAD"
    If Trim(txtPlayer.Text) = "" Then BROWSER GetSetting(App.EXEName, "ScoreCardControlAll", "ScoreCardControlAll", ""), "is_HOME"
    
    Screen.MousePointer = vbDefault
    If cn.State = adStateOpen Then cn.Close
End If

Exit Sub
PG:
Screen.MousePointer = vbDefault
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub TimerReportModifiedMolave_Timer()
TimerReportModifiedMolave.Enabled = False
Select Case cmbReportType.ListIndex
    Case 0  'INDIVIDUAL
        Select Case cmbGroup.ListIndex
            Case 0  'NetPoints
                
                picPrint.Enabled = False
                picProgress.ZOrder 0
                picProgressBar.BackColor = &HFFFFFF
                picProgress.Visible = True
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
                    .Name = "Top " & CStr(Trim(txtTop.Text))
                End With
                With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = TournamentName
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 10
                    .Range(strRange).Font.Bold = True
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Range : " & TournamentRange
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    If cmbDivisionStableford.ListIndex = 0 Then
                        .Range(strRange).Value = "Individual (Net Points) [" & IIf(cmbGender.ListIndex = 0, "MALE", "FEMALE") & "]"
                    Else
                        .Range(strRange).Value = "Individual [Class " & cmbDivisionStableford.List(cmbDivisionStableford.ListIndex) & "] (Net Points) [" & IIf(cmbGender.ListIndex = 0, "MALE", "FEMALE") & "]"
                    End If
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = ""
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "#"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCnt).ColumnWidth = 3
                    .Range(strRange).HorizontalAlignment = 4
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Name"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Handicap"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    
                    Arr = Split(TournamentRange, " - ", -1, 1)
                    iDay = 0
                    For i = 0 To DateDiff("d", Arr(0), Arr(1), vbMonday)
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = "Day " & i + 1
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).HorizontalAlignment = 4
                        iDay = iDay + 1
                    Next i
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Total"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    j = 0
                    If cmbDivision.ListIndex = 0 Then
                        If IsDate(txtDatePrint.Text) = True Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Net) AS Holes_B9, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") " & _
                                " AND (tbl_Scoring_ScoreCard.DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "') " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.NetPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Net) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Net) AS Holes_B9, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.NetPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Net) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        End If
                    Else
                        If IsDate(txtDatePrint.Text) = True Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Net) AS Holes_B9, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " AND (tbl_Scoring_ScoreCard.DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "') " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.NetPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Net) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Net) AS Holes_B9, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.NetPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Net) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        End If
                    End If
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open s, ConnOmega
                    While Not rs.EOF
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
                        .Range(strRange).Value = rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        .Columns(ColCnt).ColumnWidth = 25
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = rs!Handicap
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        sTotal = "="
                        For i = 1 To iDay
                            dDate = DateAdd("d", CDbl(i - 1), Arr(0))
                            ColCnt = ColCnt + 1
                            strRange = EXCEL_RANGE(ColCnt, RowCnt)
                            sTotal = sTotal & strRange & "+"
                            If IsDate(txtDatePrint.Text) = True Then
                                If DateValue(dDate) = DateValue(FormatDateTime(txtDatePrint.Text, vbShortDate)) Then
                                    t = "SELECT NetPoints " & _
                                        " From tbl_Scoring_ScoreCard " & _
                                        " WHERE (PlayerKey = " & rs!PlayerKey & ") " & _
                                        " AND (DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "')"
                                    If rt.State = adStateOpen Then rt.Close
                                    rt.Open t, ConnOmega
                                    If rt.RecordCount > 0 Then
                                        .Range(strRange).Value = rt!NetPoints
                                        .Range(strRange).Font.Name = "Tahoma"
                                        .Range(strRange).Font.Size = 8
                                        .Range(strRange).Font.Bold = False
                                    Else
                                        .Range(strRange).Value = ""
                                        .Range(strRange).Font.Name = "Tahoma"
                                        .Range(strRange).Font.Size = 8
                                        .Range(strRange).Font.Bold = False
                                    End If
                                    rt.Close
                                Else
                                    .Range(strRange).Value = ""
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                End If
                            Else
                                t = "SELECT NetPoints " & _
                                    " From tbl_Scoring_ScoreCard " & _
                                    " WHERE (PlayerKey = " & rs!PlayerKey & ") " & _
                                    " AND (DDate = '" & FormatDateTime(dDate, vbShortDate) & "')"
                                If rt.State = adStateOpen Then rt.Close
                                rt.Open t, ConnOmega
                                If rt.RecordCount > 0 Then
                                    .Range(strRange).Value = rt!NetPoints
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                Else
                                    .Range(strRange).Value = ""
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                End If
                                rt.Close
                            End If
                            
                        Next i
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = Mid(sTotal, 1, Len(sTotal) - 1)
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        UpdateProgress picProgressBar, j / rs.RecordCount
                        
                        rs.MoveNext
                    Wend
                    rs.Close
                    
                    .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
                    
                End With
                
SAVING1:
                On Error GoTo err_saving1:
                If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
                xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName
                
                xlsApp.Visible = True
                
                picProgress.Visible = False
                picPrint.Enabled = True
                
            Case 1  'GrossPoints
            
                picPrint.Enabled = False
                picProgress.ZOrder 0
                picProgressBar.BackColor = &HFFFFFF
                picProgress.Visible = True
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
                    .Name = "Top " & CStr(Trim(txtTop.Text))
                End With
                With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = TournamentName
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 10
                    .Range(strRange).Font.Bold = True
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Range : " & TournamentRange
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    If cmbDivision.ListIndex = 0 Then
                        .Range(strRange).Value = "Individual (Gross Points) [" & IIf(cmbGender.ListIndex = 0, "MALE", "FEMALE") & "]"
                    Else
                        .Range(strRange).Value = "Individual [Class " & cmbDivision.List(cmbDivision.ListIndex) & "] (Gross Points) [" & IIf(cmbGender.ListIndex = 0, "MALE", "FEMALE") & "]"
                    End If
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = ""
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "#"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCnt).ColumnWidth = 3
                    .Range(strRange).HorizontalAlignment = 4
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Name"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Handicap"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    
                    Arr = Split(TournamentRange, " - ", -1, 1)
                    iDay = 0
                    For i = 0 To DateDiff("d", Arr(0), Arr(1), vbMonday)
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = "Day " & i + 1
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).HorizontalAlignment = 4
                        iDay = iDay + 1
                    Next i
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Total"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    
                    j = 0
                    If cmbDivision.ListIndex = 0 Then
                        If IsDate(txtDatePrint.Text) = True Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Gross) AS Holes_B9, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") " & _
                                " AND (tbl_Scoring_ScoreCard.DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "') " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.GrossPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Gross) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Gross) AS Holes_B9, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.GrossPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Gross) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        End If
                    Else
                        If IsDate(txtDatePrint.Text) = True Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Gross) AS Holes_B9, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ")  AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " AND (tbl_Scoring_ScoreCard.DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "') " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.GrossPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Gross) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Gross) AS Holes_B9, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ")  AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.GrossPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Gross) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        End If
                    End If
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open s, ConnOmega
                    While Not rs.EOF
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
                        .Range(strRange).Value = rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        .Columns(ColCnt).ColumnWidth = 25
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = rs!Handicap
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        sTotal = "="
                        For i = 1 To iDay
                            dDate = DateAdd("d", CDbl(i - 1), Arr(0))
                            ColCnt = ColCnt + 1
                            strRange = EXCEL_RANGE(ColCnt, RowCnt)
                            sTotal = sTotal & strRange & "+"
                            If IsDate(txtDatePrint.Text) = True Then
                                If DateValue(dDate) = DateValue(FormatDateTime(txtDatePrint.Text, vbShortDate)) Then
                                    t = "SELECT GrossPoints as NetPoints " & _
                                    " From tbl_Scoring_ScoreCard " & _
                                    " WHERE (PlayerKey = " & rs!PlayerKey & ") " & _
                                    " AND (DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "')"
                                    If rt.State = adStateOpen Then rt.Close
                                    rt.Open t, ConnOmega
                                    If rt.RecordCount > 0 Then
                                        .Range(strRange).Value = rt!NetPoints
                                        .Range(strRange).Font.Name = "Tahoma"
                                        .Range(strRange).Font.Size = 8
                                        .Range(strRange).Font.Bold = False
                                    Else
                                        .Range(strRange).Value = ""
                                        .Range(strRange).Font.Name = "Tahoma"
                                        .Range(strRange).Font.Size = 8
                                        .Range(strRange).Font.Bold = False
                                    End If
                                    rt.Close
                                Else
                                    .Range(strRange).Value = ""
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                End If
                            Else
                                t = "SELECT GrossPoints as NetPoints " & _
                                    " From tbl_Scoring_ScoreCard " & _
                                    " WHERE (PlayerKey = " & rs!PlayerKey & ") " & _
                                    " AND (DDate = '" & FormatDateTime(dDate, vbShortDate) & "')"
                                If rt.State = adStateOpen Then rt.Close
                                rt.Open t, ConnOmega
                                If rt.RecordCount > 0 Then
                                    .Range(strRange).Value = rt!NetPoints
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                Else
                                    .Range(strRange).Value = ""
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                End If
                                rt.Close
                            End If
                        Next i
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = Mid(sTotal, 1, Len(sTotal) - 1)
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        UpdateProgress picProgressBar, j / rs.RecordCount
                        
                        rs.MoveNext
                    Wend
                    rs.Close

                    .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
                    
                End With
SAVING2:
                On Error GoTo err_saving2:
                If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
                xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName
                
                xlsApp.Visible = True
                
                picProgress.Visible = False
                picPrint.Enabled = True
                
            Case 2      'Gross Score
                
                picPrint.Enabled = False
                picProgress.ZOrder 0
                picProgressBar.BackColor = &HFFFFFF
                picProgress.Visible = True
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
                    .Name = "Top " & CStr(Trim(txtTop.Text))
                End With
                With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = TournamentName
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 10
                    .Range(strRange).Font.Bold = True
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Range : " & TournamentRange
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    If cmbDivision.ListIndex = 0 Then
                        .Range(strRange).Value = "Individual (Gross Score) [" & IIf(cmbGender.ListIndex = 0, "MALE", "FEMALE") & "]"
                    Else
                        .Range(strRange).Value = "Individual [Class " & cmbDivision.List(cmbDivision.ListIndex) & "] (Gross Score) [" & IIf(cmbGender.ListIndex = 0, "MALE", "FEMALE") & "]"
                    End If
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = ""
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "#"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCnt).ColumnWidth = 3
                    .Range(strRange).HorizontalAlignment = 4
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Name"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Handicap"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    
                    Arr = Split(TournamentRange, " - ", -1, 1)
                    iDay = 0
                    For i = 0 To DateDiff("d", Arr(0), Arr(1), vbMonday)
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = "Day " & i + 1
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).HorizontalAlignment = 4
                        iDay = iDay + 1
                    Next i
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Total"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    
                    j = 0
                    If cmbDivision.ListIndex = 0 Then
                        If IsDate(txtDatePrint.Text) = True Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.Score) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Score) AS Holes_B9, (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") " & _
                                " AND (tbl_Scoring_ScoreCard.DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "') " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.Score), SUM(tbl_Scoring_ScoreCard.Back9Score) , (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)), (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) "
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.Score) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Score) AS Holes_B9, (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.Score), SUM(tbl_Scoring_ScoreCard.Back9Score) , (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)), (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) "
                        End If
                    Else
                        If IsDate(txtDatePrint.Text) = True Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.Score) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Score) AS Holes_B9, (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ")  AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " AND (tbl_Scoring_ScoreCard.DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "') " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.Score) , SUM(tbl_Scoring_ScoreCard.Back9Score) , (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) , (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) "
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.Score) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Score) AS Holes_B9, (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ")  AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.Score) , SUM(tbl_Scoring_ScoreCard.Back9Score) , (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) , (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) "
                        End If
                    End If
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open s, ConnOmega
                    While Not rs.EOF
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
                        .Range(strRange).Value = rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        .Columns(ColCnt).ColumnWidth = 25
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = rs!Handicap
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        sTotal = "="
                        For i = 1 To iDay
                            dDate = DateAdd("d", CDbl(i - 1), Arr(0))
                            ColCnt = ColCnt + 1
                            strRange = EXCEL_RANGE(ColCnt, RowCnt)
                            sTotal = sTotal & strRange & "+"
                            If IsDate(txtDatePrint.Text) = True Then
                                If DateValue(dDate) = DateValue(FormatDateTime(txtDatePrint.Text, vbShortDate)) Then
                                    t = "SELECT Score as NetPoints " & _
                                        " From tbl_Scoring_ScoreCard " & _
                                        " WHERE (PlayerKey = " & rs!PlayerKey & ") " & _
                                        " AND (DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "')"
                                    If rt.State = adStateOpen Then rt.Close
                                    rt.Open t, ConnOmega
                                    If rt.RecordCount > 0 Then
                                        .Range(strRange).Value = rt!NetPoints
                                        .Range(strRange).Font.Name = "Tahoma"
                                        .Range(strRange).Font.Size = 8
                                        .Range(strRange).Font.Bold = False
                                    Else
                                        .Range(strRange).Value = ""
                                        .Range(strRange).Font.Name = "Tahoma"
                                        .Range(strRange).Font.Size = 8
                                        .Range(strRange).Font.Bold = False
                                    End If
                                    rt.Close
                                Else
                                    .Range(strRange).Value = ""
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                End If
                            Else
                                t = "SELECT Score as NetPoints " & _
                                    " From tbl_Scoring_ScoreCard " & _
                                    " WHERE (PlayerKey = " & rs!PlayerKey & ") " & _
                                    " AND (DDate = '" & FormatDateTime(dDate, vbShortDate) & "')"
                                If rt.State = adStateOpen Then rt.Close
                                rt.Open t, ConnOmega
                                If rt.RecordCount > 0 Then
                                    .Range(strRange).Value = rt!NetPoints
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                Else
                                    .Range(strRange).Value = ""
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                End If
                                rt.Close
                            End If
                        Next i
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = Mid(sTotal, 1, Len(sTotal) - 1)
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        UpdateProgress picProgressBar, j / rs.RecordCount
                        
                        rs.MoveNext
                    Wend
                    rs.Close
                    
                    .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
                    
                End With
SAVING6:
                On Error GoTo err_saving6:
                If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
                xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName
                
                xlsApp.Visible = True
                
                picProgress.Visible = False
                picPrint.Enabled = True
                
        End Select
        
    Case 1  'TEAM
    
        Select Case cmbGroup.ListIndex
            Case 0  'NetPoints
                
                picPrint.Enabled = False
                picProgress.ZOrder 0
                picProgressBar.BackColor = &HFFFFFF
                picProgress.Visible = True
                DoEvents
                
                ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_Team_Rep " & _
                                  " WHERE (LogInName = '" & gbl_UserName & "')"
                
                If cmbDivision.ListIndex = 0 Then
                    s = "SELECT PK, TeamName " & _
                        " From tbl_Scoring_Team " & _
                        " WHERE (TournamentKey = " & TournamentKey & ")"
                Else
                    s = "SELECT PK, TeamName " & _
                        " From tbl_Scoring_Team " & _
                        " WHERE ((SELECT Class " & _
                        " FROM tbl_Scoring_TournamentInfo_Class AS tbl_Scoring_TournamentInfo_Class_1 " & _
                        " WHERE (HFrom <= tbl_Scoring_Team.TeamHDCP) AND (HTo >= tbl_Scoring_Team.TeamHDCP) " & _
                        " AND (TournamentKey = " & TournamentKey & ")) = '" & cmbDivision.List(cmbDivision.ListIndex) & "') " & _
                        " AND (TournamentKey = " & TournamentKey & ")"
                End If
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                While Not rs.EOF
                    ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_Team_Rep " & _
                                      " (LogInName, TeamKey, TeamName) " & _
                                      " VALUES ('" & gbl_UserName & "', " & rs!PK & ", '" & FORMATSQL(rs!TeamName) & "')"
                    
                    dTotalTeam = 0
                    t = "SELECT  TOP " & TeamPlayer2Cnt & " tbl_Scoring_Team_Detail.PlayerKey, " & _
                        " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                        " tbl_Scoring_PlayerName.MiddleName, " & _
                        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS NetPoints " & _
                        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                        " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                        " Order By ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) DESC"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        dTotalTeam = dTotalTeam + CDbl(rt!NetPoints)
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_Team_Rep " & _
                                      " SET Score = " & CDbl(dTotalTeam) & " " & _
                                      " WHERE (LogInName = '" & gbl_UserName & "') " & _
                                      " AND (TeamKey = " & rs!PK & ")"
                    
                    dTotalTeam = 0: iTeamCounter = 0
                    t = "SELECT tbl_Scoring_Team_Detail.PlayerKey, " & _
                        " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                        " tbl_Scoring_PlayerName.MiddleName, " & _
                        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS NetPoints " & _
                        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                        " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                        " Order By ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) DESC"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        iTeamCounter = iTeamCounter + 1
                        If CDbl(iTeamCounter) > CDbl(TeamPlayer2Cnt) Then
                            dTotalTeam = dTotalTeam + CDbl(rt!NetPoints)
                        End If
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_Team_Rep " & _
                                      " SET CountBck = " & CDbl(dTotalTeam) & " " & _
                                      " WHERE (LogInName = '" & gbl_UserName & "') " & _
                                      " AND (TeamKey = " & rs!PK & ")"
                    
                    rs.MoveNext
                Wend
                rs.Close
                
                
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
                    .Name = "Top " & CStr(Trim(txtTop.Text))
                End With
                With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = TournamentName
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 10
                    .Range(strRange).Font.Bold = True

                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Range : " & TournamentRange
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    If cmbDivision.ListIndex > 0 Then
                        RowCnt = RowCnt + 1
                        ColCnt = 0
                        ColCnt = ColCnt + 1
                        HeaderRow = HeaderRow + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = "CLASS " & cmbDivision.List(cmbDivision.ListIndex)
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                    End If
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Team (Net Points)"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False

                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = ""
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True

                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "#"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCnt).ColumnWidth = 3
                    .Range(strRange).HorizontalAlignment = 4
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                        
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Name"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Columns(ColCnt).ColumnWidth = 25
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    Arr = Split(TournamentRange, " - ", -1, 1)
                    iDay = 0
                    For i = 0 To DateDiff("d", Arr(0), Arr(1), vbMonday)
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = "Day " & i + 1
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).HorizontalAlignment = 4
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        iDay = iDay + 1
                    Next i

                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Total"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Team Total"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCnt).ColumnWidth = 10
                    .Range(strRange).HorizontalAlignment = 4
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1

                    j = 0
                    
                    s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " TeamKey as PK, TeamName, Score  " & _
                        " FROM tbl_Scoring_ScoreCard_Team_Rep " & _
                        " WHERE (LogInName = '" & gbl_UserName & "') " & _
                        " ORDER BY Score DESC, CountBck DESC"
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open s, ConnOmega
                    While Not rs.EOF
                        j = j + 1
                        RowCnt = RowCnt + 1
                        ColCnt = 0
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = j
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        RowCntTmp = RowCnt - 1
                        
                        ColCntTmp = ColCnt
                        RowCntTmp = RowCntTmp + 1
                        ColCntTmp = ColCntTmp + 1
                        strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                        .Range(strRange).Value = rs!TeamName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        For i = 1 To iDay
                            ColCntTmp = ColCntTmp + 1
                            strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                            .Range(strRange).Value = ""
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                        Next i
                        
                        t = "SELECT  tbl_Scoring_Team_Detail.PlayerKey, " & _
                            " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                            " tbl_Scoring_PlayerName.MiddleName, " & _
                            " (SELECT SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints " & _
                            " From tbl_Scoring_ScoreCard " & _
                            " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)) AS NetPoints " & _
                            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                            " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                            " Order By (SELECT SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints " & _
                            " From tbl_Scoring_ScoreCard " & _
                            " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)) DESC"
                        If rt.State = adStateOpen Then rt.Close
                        rt.Open t, ConnOmega
                        While Not rt.EOF
                            
                            
                            ColCntTmp = ColCnt
                            RowCntTmp = RowCntTmp + 1
                            ColCntTmp = ColCntTmp + 1
                            strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                            .Range(strRange).Value = rt!LastName & ",  " & rt!FirstName & "  " & rt!MiddleName
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                            sTotal = "="
                            For i = 1 To iDay
                                dDate = DateAdd("d", CDbl(i - 1), Arr(0))
                                ColCntTmp = ColCntTmp + 1
                                strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                                sTotal = sTotal & strRange & "+"
                                u = "SELECT NetPoints " & _
                                    " From tbl_Scoring_ScoreCard " & _
                                    " WHERE (PlayerKey = " & rt!PlayerKey & ") " & _
                                    " AND (DDate = '" & FormatDateTime(dDate, vbShortDate) & "')"
                                If ru.State = adStateOpen Then ru.Close
                                ru.Open u, ConnOmega
                                If ru.RecordCount > 0 Then
                                    .Range(strRange).Value = ru!NetPoints
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                Else
                                    .Range(strRange).Value = ""
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                End If
                                .Range(strRange).Select
                                xlsApp.Selection.Borders.LineStyle = 1
                                ru.Close
                            Next i
                            
                            ColCntTmp = ColCntTmp + 1
                            strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                            .Range(strRange).Value = Mid(sTotal, 1, Len(sTotal) - 1)
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                            rt.MoveNext
                        Wend
                        rt.Close
                        
                        ColCnt = ColCntTmp
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        strRangeFrom = EXCEL_RANGE(ColCnt, RowCnt)
                        strRangeTo = EXCEL_RANGE(ColCnt, RowCntTmp)
                        .Range(strRange).Value = rs!Score 'rs!NetPoints
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRangeFrom, strRangeTo).Select
                        xlsApp.Selection.Merge
                        .Range(strRange).VerticalAlignment = 2
                        .Range(strRange).HorizontalAlignment = 3
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        strRange = EXCEL_RANGE(1, RowCnt)
                        strRangeFrom = EXCEL_RANGE(1, RowCnt)
                        strRangeTo = EXCEL_RANGE(1, RowCntTmp)
                        .Range(strRangeFrom, strRangeTo).Select
                        xlsApp.Selection.Merge
                        .Range(strRange).VerticalAlignment = 2
                        .Range(strRange).HorizontalAlignment = 3
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        RowCnt = RowCntTmp
                        
                        UpdateProgress picProgressBar, j / rs.RecordCount
                        
                        rs.MoveNext
                    Wend
                    rs.Close
                    
                    .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
                    
                End With

SAVING3:
                On Error GoTo err_saving3:
                If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
                xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

                xlsApp.Visible = True
                
                picProgress.Visible = False
                picPrint.Enabled = True
                
            Case 1  'Gross Points
                
                picPrint.Enabled = False
                picProgress.ZOrder 0
                picProgressBar.BackColor = &HFFFFFF
                picProgress.Visible = True
                DoEvents
                
                ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_Team_Rep " & _
                                  " WHERE (LogInName = '" & gbl_UserName & "')"
                
                If cmbDivision.ListIndex = 0 Then
                    s = "SELECT PK, TeamName " & _
                        " From tbl_Scoring_Team " & _
                        " WHERE (TournamentKey = " & TournamentKey & ")"
                Else
                    s = "SELECT PK, TeamName " & _
                        " From tbl_Scoring_Team " & _
                        " WHERE ((SELECT Class " & _
                        " FROM tbl_Scoring_TournamentInfo_Class AS tbl_Scoring_TournamentInfo_Class_1 " & _
                        " WHERE (HFrom <= tbl_Scoring_Team.TeamHDCP) AND (HTo >= tbl_Scoring_Team.TeamHDCP) " & _
                        " AND (TournamentKey = " & TournamentKey & ")) = '" & cmbDivision.List(cmbDivision.ListIndex) & "') " & _
                        " AND (TournamentKey = " & TournamentKey & ")"
                End If
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                While Not rs.EOF
                    ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_Team_Rep " & _
                                      " (LogInName, TeamKey, TeamName) " & _
                                      " VALUES ('" & gbl_UserName & "', " & rs!PK & ", '" & FORMATSQL(rs!TeamName) & "')"
                    
                    dTotalTeam = 0
                    t = "SELECT  TOP " & TeamPlayer2Cnt & " tbl_Scoring_Team_Detail.PlayerKey, " & _
                        " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                        " tbl_Scoring_PlayerName.MiddleName, " & _
                        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS NetPoints " & _
                        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                        " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                        " Order By ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) DESC"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        dTotalTeam = dTotalTeam + CDbl(rt!NetPoints)
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_Team_Rep " & _
                                      " SET Score = " & CDbl(dTotalTeam) & " " & _
                                      " WHERE (LogInName = '" & gbl_UserName & "') " & _
                                      " AND (TeamKey = " & rs!PK & ")"
                    
                    dTotalTeam = 0: iTeamCounter = 0
                    t = "SELECT tbl_Scoring_Team_Detail.PlayerKey, " & _
                        " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                        " tbl_Scoring_PlayerName.MiddleName, " & _
                        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS NetPoints " & _
                        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                        " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                        " Order By ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) DESC"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        iTeamCounter = iTeamCounter + 1
                        If CDbl(iTeamCounter) > CDbl(TeamPlayer2Cnt) Then
                            dTotalTeam = dTotalTeam + CDbl(rt!NetPoints)
                        End If
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_Team_Rep " & _
                                      " SET CountBck = " & CDbl(dTotalTeam) & " " & _
                                      " WHERE (LogInName = '" & gbl_UserName & "') " & _
                                      " AND (TeamKey = " & rs!PK & ")"
                    
                    rs.MoveNext
                Wend
                rs.Close
                
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
                    .Name = "Top " & CStr(Trim(txtTop.Text))
                End With
                With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = TournamentName
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 10
                    .Range(strRange).Font.Bold = True

                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Range : " & TournamentRange
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    If cmbDivision.ListIndex > 0 Then
                        RowCnt = RowCnt + 1
                        ColCnt = 0
                        ColCnt = ColCnt + 1
                        HeaderRow = HeaderRow + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = "CLASS " & cmbDivision.List(cmbDivision.ListIndex)
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                    End If
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Team (Gross Points)"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False

                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = ""
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True

                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "#"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCnt).ColumnWidth = 3
                    .Range(strRange).HorizontalAlignment = 4
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                        
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Name"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Columns(ColCnt).ColumnWidth = 25
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    Arr = Split(TournamentRange, " - ", -1, 1)
                    iDay = 0
                    For i = 0 To DateDiff("d", Arr(0), Arr(1), vbMonday)
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = "Day " & i + 1
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).HorizontalAlignment = 4
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        iDay = iDay + 1
                    Next i

                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Total"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Team Total"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCnt).ColumnWidth = 10
                    .Range(strRange).HorizontalAlignment = 4
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1

                    j = 0
                    s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " TeamKey as PK, TeamName, Score  " & _
                        " FROM tbl_Scoring_ScoreCard_Team_Rep " & _
                        " WHERE (LogInName = '" & gbl_UserName & "') " & _
                        " ORDER BY Score DESC, CountBck DESC"
                        
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open s, ConnOmega
                    While Not rs.EOF
                        j = j + 1
                        RowCnt = RowCnt + 1
                        ColCnt = 0
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = j
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        RowCntTmp = RowCnt - 1
                        
                        ColCntTmp = ColCnt
                        RowCntTmp = RowCntTmp + 1
                        ColCntTmp = ColCntTmp + 1
                        strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                        .Range(strRange).Value = rs!TeamName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        For i = 1 To iDay
                            ColCntTmp = ColCntTmp + 1
                            strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                            .Range(strRange).Value = ""
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                        Next i
                        
                        t = "SELECT  tbl_Scoring_Team_Detail.PlayerKey, " & _
                            " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                            " tbl_Scoring_PlayerName.MiddleName, " & _
                            " (SELECT SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints " & _
                            " From tbl_Scoring_ScoreCard " & _
                            " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)) AS NetPoints " & _
                            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                            " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                            " Order By (SELECT SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints " & _
                            " From tbl_Scoring_ScoreCard " & _
                            " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)) DESC"
                        If rt.State = adStateOpen Then rt.Close
                        rt.Open t, ConnOmega
                        While Not rt.EOF
                            
                            ColCntTmp = ColCnt
                            RowCntTmp = RowCntTmp + 1
                            ColCntTmp = ColCntTmp + 1
                            strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                            .Range(strRange).Value = rt!LastName & ",  " & rt!FirstName & "  " & rt!MiddleName
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                            sTotal = "="
                            For i = 1 To iDay
                                dDate = DateAdd("d", CDbl(i - 1), Arr(0))
                                ColCntTmp = ColCntTmp + 1
                                strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                                sTotal = sTotal & strRange & "+"
                                u = "SELECT GrossPoints as NetPoints " & _
                                    " From tbl_Scoring_ScoreCard " & _
                                    " WHERE (PlayerKey = " & rt!PlayerKey & ") " & _
                                    " AND (DDate = '" & FormatDateTime(dDate, vbShortDate) & "')"
                                If ru.State = adStateOpen Then ru.Close
                                ru.Open u, ConnOmega
                                If ru.RecordCount > 0 Then
                                    .Range(strRange).Value = ru!NetPoints
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                Else
                                    .Range(strRange).Value = ""
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                End If
                                .Range(strRange).Select
                                xlsApp.Selection.Borders.LineStyle = 1
                                ru.Close
                            Next i
                            
                            ColCntTmp = ColCntTmp + 1
                            strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                            .Range(strRange).Value = Mid(sTotal, 1, Len(sTotal) - 1)
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                            rt.MoveNext
                        Wend
                        rt.Close
                        
                        ColCnt = ColCntTmp
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        strRangeFrom = EXCEL_RANGE(ColCnt, RowCnt)
                        strRangeTo = EXCEL_RANGE(ColCnt, RowCntTmp)
                        .Range(strRange).Value = rs!Score 'rs!NetPoints
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRangeFrom, strRangeTo).Select
                        xlsApp.Selection.Merge
                        .Range(strRange).VerticalAlignment = 2
                        .Range(strRange).HorizontalAlignment = 3
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        strRange = EXCEL_RANGE(1, RowCnt)
                        strRangeFrom = EXCEL_RANGE(1, RowCnt)
                        strRangeTo = EXCEL_RANGE(1, RowCntTmp)
                        .Range(strRangeFrom, strRangeTo).Select
                        xlsApp.Selection.Merge
                        .Range(strRange).VerticalAlignment = 2
                        .Range(strRange).HorizontalAlignment = 3
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        RowCnt = RowCntTmp
                        
                        UpdateProgress picProgressBar, j / rs.RecordCount
                        
                        rs.MoveNext
                    Wend
                    rs.Close
                    
                    .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
                    
                End With

SAVING4:
                On Error GoTo err_saving4:
                If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
                xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

                xlsApp.Visible = True
                
                picProgress.Visible = False
                picPrint.Enabled = True
                
            Case 2      'Gross Score
                
                picPrint.Enabled = False
                picProgress.ZOrder 0
                picProgressBar.BackColor = &HFFFFFF
                picProgress.Visible = True
                DoEvents
                
                ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_Team_Rep " & _
                                  " WHERE (LogInName = '" & gbl_UserName & "')"
                
                If cmbDivision.ListIndex = 0 Then
                    s = "SELECT PK, TeamName " & _
                        " From tbl_Scoring_Team " & _
                        " WHERE (TournamentKey = " & TournamentKey & ")"
                Else
                    s = "SELECT PK, TeamName " & _
                        " From tbl_Scoring_Team " & _
                        " WHERE ((SELECT Class " & _
                        " FROM tbl_Scoring_TournamentInfo_Class AS tbl_Scoring_TournamentInfo_Class_1 " & _
                        " WHERE (HFrom <= tbl_Scoring_Team.TeamHDCP) AND (HTo >= tbl_Scoring_Team.TeamHDCP) " & _
                        " AND (TournamentKey = " & TournamentKey & ")) = '" & cmbDivision.List(cmbDivision.ListIndex) & "') " & _
                        " AND (TournamentKey = " & TournamentKey & ")"
                End If
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                While Not rs.EOF
                    ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_Team_Rep " & _
                                      " (LogInName, TeamKey, TeamName) " & _
                                      " VALUES ('" & gbl_UserName & "', " & rs!PK & ", '" & FORMATSQL(rs!TeamName) & "')"
                    
                    dTotalTeam = 0
                    t = "SELECT  TOP " & TeamPlayer2Cnt & " tbl_Scoring_Team_Detail.PlayerKey, " & _
                        " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                        " tbl_Scoring_PlayerName.MiddleName, " & _
                        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.Score) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS NetPoints " & _
                        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                        " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                        " Order By ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.Score) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0)"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        dTotalTeam = dTotalTeam + CDbl(rt!NetPoints)
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_Team_Rep " & _
                                      " SET Score = " & CDbl(dTotalTeam) & " " & _
                                      " WHERE (LogInName = '" & gbl_UserName & "') " & _
                                      " AND (TeamKey = " & rs!PK & ")"
                    
                    dTotalTeam = 0: iTeamCounter = 0
                    t = "SELECT tbl_Scoring_Team_Detail.PlayerKey, " & _
                        " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                        " tbl_Scoring_PlayerName.MiddleName, " & _
                        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.Score) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS NetPoints " & _
                        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                        " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                        " Order By ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.Score) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0)"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        iTeamCounter = iTeamCounter + 1
                        If CDbl(iTeamCounter) > CDbl(TeamPlayer2Cnt) Then
                            dTotalTeam = dTotalTeam + CDbl(rt!NetPoints)
                        End If
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_Team_Rep " & _
                                      " SET CountBck = " & CDbl(dTotalTeam) & " " & _
                                      " WHERE (LogInName = '" & gbl_UserName & "') " & _
                                      " AND (TeamKey = " & rs!PK & ")"
                    
                    rs.MoveNext
                Wend
                rs.Close
                
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
                    .Name = "Top " & CStr(Trim(txtTop.Text))
                End With
                With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = TournamentName
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 10
                    .Range(strRange).Font.Bold = True

                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Range : " & TournamentRange
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    If cmbDivision.ListIndex > 0 Then
                        RowCnt = RowCnt + 1
                        ColCnt = 0
                        ColCnt = ColCnt + 1
                        HeaderRow = HeaderRow + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = "CLASS " & cmbDivision.List(cmbDivision.ListIndex)
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                    End If
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Team (Gross Points)"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False

                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = ""
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True

                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "#"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCnt).ColumnWidth = 3
                    .Range(strRange).HorizontalAlignment = 4
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                        
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Name"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Columns(ColCnt).ColumnWidth = 25
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    Arr = Split(TournamentRange, " - ", -1, 1)
                    iDay = 0
                    For i = 0 To DateDiff("d", Arr(0), Arr(1), vbMonday)
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = "Day " & i + 1
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).HorizontalAlignment = 4
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        iDay = iDay + 1
                    Next i

                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Total"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Team Total"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCnt).ColumnWidth = 10
                    .Range(strRange).HorizontalAlignment = 4
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1

                    j = 0
                    s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " TeamKey as PK, TeamName, Score  " & _
                        " FROM tbl_Scoring_ScoreCard_Team_Rep " & _
                        " WHERE (LogInName = '" & gbl_UserName & "') " & _
                        " ORDER BY Score , CountBck "
                        
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open s, ConnOmega
                    While Not rs.EOF
                        j = j + 1
                        RowCnt = RowCnt + 1
                        ColCnt = 0
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = j
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        RowCntTmp = RowCnt - 1
                        
                        ColCntTmp = ColCnt
                        RowCntTmp = RowCntTmp + 1
                        ColCntTmp = ColCntTmp + 1
                        strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                        .Range(strRange).Value = rs!TeamName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        For i = 1 To iDay
                            ColCntTmp = ColCntTmp + 1
                            strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                            .Range(strRange).Value = ""
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                        Next i
                        
                        t = "SELECT  tbl_Scoring_Team_Detail.PlayerKey, " & _
                            " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                            " tbl_Scoring_PlayerName.MiddleName, " & _
                            " (SELECT SUM(tbl_Scoring_ScoreCard.Score) AS NetPoints " & _
                            " From tbl_Scoring_ScoreCard " & _
                            " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)) AS NetPoints " & _
                            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                            " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                            " Order By (SELECT SUM(tbl_Scoring_ScoreCard.Score) AS NetPoints " & _
                            " From tbl_Scoring_ScoreCard " & _
                            " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)) "
                        If rt.State = adStateOpen Then rt.Close
                        rt.Open t, ConnOmega
                        While Not rt.EOF
                            
                            ColCntTmp = ColCnt
                            RowCntTmp = RowCntTmp + 1
                            ColCntTmp = ColCntTmp + 1
                            strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                            .Range(strRange).Value = rt!LastName & ",  " & rt!FirstName & "  " & rt!MiddleName
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                            sTotal = "="
                            For i = 1 To iDay
                                dDate = DateAdd("d", CDbl(i - 1), Arr(0))
                                ColCntTmp = ColCntTmp + 1
                                strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                                sTotal = sTotal & strRange & "+"
                                u = "SELECT GrossPoints as NetPoints " & _
                                    " From tbl_Scoring_ScoreCard " & _
                                    " WHERE (PlayerKey = " & rt!PlayerKey & ") " & _
                                    " AND (DDate = '" & FormatDateTime(dDate, vbShortDate) & "')"
                                If ru.State = adStateOpen Then ru.Close
                                ru.Open u, ConnOmega
                                If ru.RecordCount > 0 Then
                                    .Range(strRange).Value = ru!NetPoints
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                Else
                                    .Range(strRange).Value = ""
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                End If
                                .Range(strRange).Select
                                xlsApp.Selection.Borders.LineStyle = 1
                                ru.Close
                            Next i
                            
                            ColCntTmp = ColCntTmp + 1
                            strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                            .Range(strRange).Value = Mid(sTotal, 1, Len(sTotal) - 1)
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                            rt.MoveNext
                        Wend
                        rt.Close
                        
                        ColCnt = ColCntTmp
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        strRangeFrom = EXCEL_RANGE(ColCnt, RowCnt)
                        strRangeTo = EXCEL_RANGE(ColCnt, RowCntTmp)
                        .Range(strRange).Value = rs!Score 'rs!NetPoints
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRangeFrom, strRangeTo).Select
                        xlsApp.Selection.Merge
                        .Range(strRange).VerticalAlignment = 2
                        .Range(strRange).HorizontalAlignment = 3
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        strRange = EXCEL_RANGE(1, RowCnt)
                        strRangeFrom = EXCEL_RANGE(1, RowCnt)
                        strRangeTo = EXCEL_RANGE(1, RowCntTmp)
                        .Range(strRangeFrom, strRangeTo).Select
                        xlsApp.Selection.Merge
                        .Range(strRange).VerticalAlignment = 2
                        .Range(strRange).HorizontalAlignment = 3
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        RowCnt = RowCntTmp
                        
                        UpdateProgress picProgressBar, j / rs.RecordCount
                        
                        rs.MoveNext
                    Wend
                    rs.Close
                    
                    .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
                    
                End With

SAVING5:
                On Error GoTo err_saving5:
                If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
                xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

                xlsApp.Visible = True
                
                picProgress.Visible = False
                picPrint.Enabled = True
                
        End Select
        
    Case 2  'Result
        
        Exit Sub
        
End Select
Exit Sub
err_saving1:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING1:

Exit Sub
err_saving2:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING2:

Exit Sub
err_saving3:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING3:

Exit Sub
err_saving4:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING4:

Exit Sub
err_saving5:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING5:

Exit Sub
err_saving6:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING6:
End Sub

Private Sub TimerReportModifiedStableford_Timer()
TimerReportModifiedStableford.Enabled = False
ArrDate = Split(TournamentRange, " - ", -1, 1)
iDateDiff = DateDiff("d", ArrDate(0), ArrDate(1))
Select Case cmbReportType.ListIndex
    Case 0  'INDIVIDUAL
        Select Case cmbGroup.ListIndex
            Case 0  'NetPoints
                
                picPrint.Visible = False
                picProgress.Visible = True
                picProgressBar.BackColor = &HFFFFFF
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
                    .Name = "Top " & CStr(Trim(txtTop.Text))
                End With
                With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = TournamentName
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 10
                    .Range(strRange).Font.Bold = True
                    
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Range : " & TournamentRange
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    If cmbDivision.ListIndex = 0 Then
                        .Range(strRange).Value = "Individual (Net Points) [" & IIf(cmbGender.ListIndex = 1, "MALE", IIf(cmbGender.ListIndex = 2, "FEMALE", "")) & "]"
                    Else
                        .Range(strRange).Value = "Individual [Class " & cmbDivision.List(cmbDivision.ListIndex) & "] (Net Points) [" & IIf(cmbGender.ListIndex = 1, "MALE", IIf(cmbGender.ListIndex = 2, "FEMALE", "")) & "]"
                    End If
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = ""
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "#"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCnt).ColumnWidth = 3
                    .Range(strRange).HorizontalAlignment = 4
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Name"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Handicap"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    
                    Arr = Split(TournamentRange, " - ", -1, 1)
                    iDay = 0
                    For i = 0 To DateDiff("d", Arr(0), Arr(1), vbMonday)
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = "Day " & i + 1
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).HorizontalAlignment = 4
                        iDay = iDay + 1
                    Next i
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Total"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    j = 0
                    If cmbDivision.ListIndex = 0 Then
                        'All Class
                        If cmbGender.ListIndex = 0 Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Net) AS Holes_B9, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ")  " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.NetPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Net) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                            
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Net) AS Holes_B9, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex & ") " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.NetPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Net) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        End If
                    Else
                        'Class A to
                        If cmbGender.ListIndex = 0 Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Net) AS Holes_B9, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.NetPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Net) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Net) AS Holes_B9, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex & ") AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.NetPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Net) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        End If
                    End If
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open s, ConnOmega
                    While Not rs.EOF
                        DoEvents
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
                        .Range(strRange).Value = rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        .Columns(ColCnt).ColumnWidth = 25
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = rs!Handicap
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        sTotal = "="
                        For i = 1 To iDay
                            dDate = DateAdd("d", CDbl(i - 1), Arr(0))
                            ColCnt = ColCnt + 1
                            strRange = EXCEL_RANGE(ColCnt, RowCnt)
                            sTotal = sTotal & strRange & "+"
                            t = "SELECT NetPoints " & _
                                " From tbl_Scoring_ScoreCard " & _
                                " WHERE (PlayerKey = " & rs!PlayerKey & ") " & _
                                " AND (DDate = '" & FormatDateTime(dDate, vbShortDate) & "')"
                            If rt.State = adStateOpen Then rt.Close
                            rt.Open t, ConnOmega
                            If rt.RecordCount > 0 Then
                                .Range(strRange).Value = rt!NetPoints
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                            Else
                                .Range(strRange).Value = ""
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                            End If
                            rt.Close
                        Next i
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = Mid(sTotal, 1, Len(sTotal) - 1)
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        UpdateProgress picProgressBar, iProgressValue / rs.RecordCount
                        
                        rs.MoveNext
                    Wend
                    rs.Close
                    '(tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex & ")
                    If cmbDivision.ListIndex = 0 Then
                        If cmbGender.ListIndex = 0 Then
                            s = "SELECT LastName, FirstName, MiddleName, HandiCap, " & _
                                " ISNULL((SELECT SUM(NetPoints) AS NetPoints " & _
                                " From dbo.tbl_Scoring_ScoreCard " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) AS NPoints " & _
                                " From dbo.tbl_Scoring_PlayerName " & _
                                " WHERE (ISNULL((SELECT  SUM(NetPoints) AS NetPoints " & _
                                " FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) = 0) " & _
                                " AND (TournamentKey = " & TournamentKey & ") " & _
                                " ORDER BY HandiCap, LastName, FirstName, MiddleName"
                        Else
                            s = "SELECT LastName, FirstName, MiddleName, HandiCap, " & _
                                " ISNULL((SELECT SUM(NetPoints) AS NetPoints " & _
                                " From dbo.tbl_Scoring_ScoreCard " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) AS NPoints " & _
                                " From dbo.tbl_Scoring_PlayerName " & _
                                " WHERE (ISNULL((SELECT  SUM(NetPoints) AS NetPoints " & _
                                " FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) = 0) " & _
                                " AND (TournamentKey = " & TournamentKey & ") " & _
                                " AND (Gender = " & cmbGender.ListIndex & ") " & _
                                " ORDER BY HandiCap, LastName, FirstName, MiddleName"
                        End If
                    Else
                        If cmbGender.ListIndex = 0 Then
                            s = "SELECT LastName, FirstName, MiddleName, HandiCap, " & _
                                " ISNULL((SELECT SUM(NetPoints) AS NetPoints " & _
                                " From dbo.tbl_Scoring_ScoreCard " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) AS NPoints " & _
                                " From dbo.tbl_Scoring_PlayerName " & _
                                " WHERE ((SELECT Class FROM dbo.tbl_Scoring_TournamentInfo_Class AS Class " & _
                                " WHERE (TournamentKey = dbo.tbl_Scoring_PlayerName.TournamentKey) AND (HFrom <= dbo.tbl_Scoring_PlayerName.HandiCap) " & _
                                " AND (HTo >= dbo.tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "') " & _
                                " AND (TournamentKey = " & TournamentKey & ") " & _
                                " AND (ISNULL((SELECT SUM(NetPoints) AS NetPoints " & _
                                " FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) = 0) " & _
                                " ORDER BY HandiCap, LastName, FirstName, MiddleName"
                        Else
                            s = "SELECT LastName, FirstName, MiddleName, HandiCap, " & _
                                " ISNULL((SELECT SUM(NetPoints) AS NetPoints " & _
                                " From dbo.tbl_Scoring_ScoreCard " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) AS NPoints " & _
                                " From dbo.tbl_Scoring_PlayerName " & _
                                " WHERE ((SELECT Class FROM dbo.tbl_Scoring_TournamentInfo_Class AS Class " & _
                                " WHERE (TournamentKey = dbo.tbl_Scoring_PlayerName.TournamentKey) AND (HFrom <= dbo.tbl_Scoring_PlayerName.HandiCap) " & _
                                " AND (HTo >= dbo.tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "') " & _
                                " AND (TournamentKey = " & TournamentKey & ") " & _
                                " AND (Gender = " & cmbGender.ListIndex & ") " & _
                                " AND (ISNULL((SELECT SUM(NetPoints) AS NetPoints " & _
                                " FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) = 0) " & _
                                " ORDER BY HandiCap, LastName, FirstName, MiddleName"
                        End If
                    End If
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open s, ConnOmega
                    While Not rs.EOF
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
                        .Range(strRange).Value = rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        .Columns(ColCnt).ColumnWidth = 25
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = rs!Handicap
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        rs.MoveNext
                    Wend
                    rs.Close
                    
                    .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
                    
                End With
                
SAVING1:
                On Error GoTo err_saving1:
                If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
                xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName
                
                xlsApp.Visible = True
                
            Case 1  'GrossPoints
                
                picPrint.Visible = False
                picProgress.Visible = True
                picProgressBar.BackColor = &HFFFFFF
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
                    .Name = "Top " & CStr(Trim(txtTop.Text))
                End With
                With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = TournamentName
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 10
                    .Range(strRange).Font.Bold = True
                    
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Range : " & TournamentRange
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    If cmbDivision.ListIndex = 0 Then
                        .Range(strRange).Value = "Individual (Gross Points) [" & IIf(cmbGender.ListIndex = 1, "MALE", IIf(cmbGender.ListIndex = 2, "FEMALE", "")) & "]"
                    Else
                        .Range(strRange).Value = "Individual [Class " & cmbDivision.List(cmbDivision.ListIndex) & "] (Gross Points) [" & IIf(cmbGender.ListIndex = 1, "MALE", IIf(cmbGender.ListIndex = 2, "FEMALE", "")) & "]"
                    End If
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = ""
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "#"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCnt).ColumnWidth = 3
                    .Range(strRange).HorizontalAlignment = 4
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Name"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Handicap"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    
                    Arr = Split(TournamentRange, " - ", -1, 1)
                    iDay = 0
                    For i = 0 To DateDiff("d", Arr(0), Arr(1), vbMonday)
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = "Day " & i + 1
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).HorizontalAlignment = 4
                        iDay = iDay + 1
                    Next i
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Total"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    
                    j = 0
                    If cmbDivision.ListIndex = 0 Then
                        If cmbGender.ListIndex = 0 Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Gross) AS Holes_B9, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.GrossPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Gross) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Gross) AS Holes_B9, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex & ") " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.GrossPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Gross) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        End If
                    Else
                        If cmbGender.ListIndex = 0 Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Gross) AS Holes_B9, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.GrossPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Gross) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Gross) AS Holes_B9, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex & ")  AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.GrossPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Gross) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        End If
                    End If
                    If ra.State = adStateOpen Then ra.Close
                    ra.Open s, ConnOmega
                    While Not ra.EOF
                        DoEvents
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
                        .Range(strRange).Value = ra!LastName & ",  " & ra!FirstName & "  " & ra!MiddleName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        .Columns(ColCnt).ColumnWidth = 25
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = ra!Handicap
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        sTotal = "="
                        For i = 1 To iDay
                            dDate = DateAdd("d", CDbl(i - 1), Arr(0))
                            ColCnt = ColCnt + 1
                            strRange = EXCEL_RANGE(ColCnt, RowCnt)
                            sTotal = sTotal & strRange & "+"
                            t = "SELECT GrossPoints as NetPoints " & _
                                " From tbl_Scoring_ScoreCard " & _
                                " WHERE (PlayerKey = " & ra!PlayerKey & ") " & _
                                " AND (DDate = '" & FormatDateTime(dDate, vbShortDate) & "')"
                            If rt.State = adStateOpen Then rt.Close
                            rt.Open t, ConnOmega
                            If rt.RecordCount > 0 Then
                                .Range(strRange).Value = rt!NetPoints
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                            Else
                                .Range(strRange).Value = ""
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                            End If
                            rt.Close
                        Next i
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = Mid(sTotal, 1, Len(sTotal) - 1)
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        UpdateProgress picProgressBar, iProgressValue / ra.RecordCount
                        
                        ra.MoveNext
                    Wend
                    ra.Close
                    
                    's = "SELECT LastName, FirstName, MiddleName, HandiCap, " & _
                        " ISNULL((SELECT SUM(GrossPoints) AS NetPoints " & _
                        " From dbo.tbl_Scoring_ScoreCard " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) AS NPoints " & _
                        " From dbo.tbl_Scoring_PlayerName " & _
                        " WHERE (ISNULL((SELECT  SUM(GrossPoints) AS NetPoints " & _
                        " FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) = 0) " & _
                        " AND (TournamentKey = " & TournamentKey & ") " & _
                        " ORDER BY HandiCap, LastName, FirstName, MiddleName"
                    
                    If cmbDivision.ListIndex = 0 Then
                        If cmbGender.ListIndex = 0 Then
                            s = "SELECT LastName, FirstName, MiddleName, HandiCap, " & _
                                " ISNULL((SELECT SUM(GrossPoints) AS NetPoints " & _
                                " From dbo.tbl_Scoring_ScoreCard " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) AS NPoints " & _
                                " From dbo.tbl_Scoring_PlayerName " & _
                                " WHERE (ISNULL((SELECT  SUM(GrossPoints) AS NetPoints " & _
                                " FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) = 0) " & _
                                " AND (TournamentKey = " & TournamentKey & ") " & _
                                " ORDER BY HandiCap, LastName, FirstName, MiddleName"
                        Else
                            s = "SELECT LastName, FirstName, MiddleName, HandiCap, " & _
                                " ISNULL((SELECT SUM(GrossPoints) AS NetPoints " & _
                                " From dbo.tbl_Scoring_ScoreCard " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) AS NPoints " & _
                                " From dbo.tbl_Scoring_PlayerName " & _
                                " WHERE (ISNULL((SELECT  SUM(GrossPoints) AS NetPoints " & _
                                " FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) = 0) " & _
                                " AND (TournamentKey = " & TournamentKey & ") " & _
                                " AND (Gender = " & cmbGender.ListIndex & ") " & _
                                " ORDER BY HandiCap, LastName, FirstName, MiddleName"
                        End If
                    Else
                        If cmbGender.ListIndex = 0 Then
                            s = "SELECT LastName, FirstName, MiddleName, HandiCap, " & _
                                " ISNULL((SELECT SUM(GrossPoints) AS NetPoints " & _
                                " From dbo.tbl_Scoring_ScoreCard " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) AS NPoints " & _
                                " From dbo.tbl_Scoring_PlayerName " & _
                                " WHERE ((SELECT Class FROM dbo.tbl_Scoring_TournamentInfo_Class AS Class " & _
                                " WHERE (TournamentKey = dbo.tbl_Scoring_PlayerName.TournamentKey) AND (HFrom <= dbo.tbl_Scoring_PlayerName.HandiCap) " & _
                                " AND (HTo >= dbo.tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "') " & _
                                " AND (TournamentKey = " & TournamentKey & ") " & _
                                " AND (ISNULL((SELECT SUM(GrossPoints) AS NetPoints " & _
                                " FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) = 0) " & _
                                " ORDER BY HandiCap, LastName, FirstName, MiddleName"
                        Else
                            s = "SELECT LastName, FirstName, MiddleName, HandiCap, " & _
                                " ISNULL((SELECT SUM(GrossPoints) AS NetPoints " & _
                                " From dbo.tbl_Scoring_ScoreCard " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) AS NPoints " & _
                                " From dbo.tbl_Scoring_PlayerName " & _
                                " WHERE ((SELECT Class FROM dbo.tbl_Scoring_TournamentInfo_Class AS Class " & _
                                " WHERE (TournamentKey = dbo.tbl_Scoring_PlayerName.TournamentKey) AND (HFrom <= dbo.tbl_Scoring_PlayerName.HandiCap) " & _
                                " AND (HTo >= dbo.tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "') " & _
                                " AND (TournamentKey = " & TournamentKey & ") " & _
                                " AND (Gender = " & cmbGender.ListIndex & ") " & _
                                " AND (ISNULL((SELECT SUM(GrossPoints) AS NetPoints " & _
                                " FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) = 0) " & _
                                " ORDER BY HandiCap, LastName, FirstName, MiddleName"
                        End If
                    End If
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open s, ConnOmega
                    While Not rs.EOF
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
                        .Range(strRange).Value = rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        .Columns(ColCnt).ColumnWidth = 25
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = rs!Handicap
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        rs.MoveNext
                    Wend
                    rs.Close
                    
                    .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
                    
                End With
SAVING2:
                On Error GoTo err_saving2:
                If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
                xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName
                
                xlsApp.Visible = True
                
        End Select
        
    Case 1  'TEAM
        
        'Create_Table
        TableName = "tmp_" & gbl_UserName & "_Scoring_ModStableFord"
        DetailTableName = TableName & "_Detail"
        CREATE_MODIFIED_STABLE_FORD TableName
        
        Select Case cmbGroup.ListIndex
        
            Case 0  'NetPoints
                
                j = 0
                picPrint.Visible = False
                picProgress.Visible = True
                picProgressBar.BackColor = &HFFFFFF
                DoEvents
                                
                s = "SELECT PK, TeamName, TeamHDCP, TeamIndex " & _
                    " From tbl_Scoring_Team " & _
                    " WHERE (TournamentKey = " & TournamentKey & ")"
                If rs.State = adStateOpen Then ra.Close
                ra.Open s, ConnOmega
                While Not ra.EOF
                    DoEvents
                    j = j + 1
                    
                    iPlayCnt = 0: dTeamTotal = 0: dB9Total = 0: dF9Total = 0: iDLine = 0: dLastMan = 0
                    
                    t = "SELECT ISNULL ((SELECT SUM(NetPoints) AS Points From dbo.tbl_Scoring_ScoreCard " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) AS Points, " & _
                        " ISNULL((SELECT SUM(Back9Net) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) as B9, " & _
                        " ISNULL((SELECT SUM(Front9Net) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) as F9 " & _
                        " FROM dbo.tbl_Scoring_Team LEFT OUTER JOIN " & _
                        " dbo.tbl_Scoring_Team_Detail ON dbo.tbl_Scoring_Team.PK = dbo.tbl_Scoring_Team_Detail.TeamKey " & _
                        " Where (dbo.tbl_Scoring_Team.PK = " & ra!PK & ") " & _
                        " ORDER BY ISNULL((SELECT SUM(NetPoints) AS Points From dbo.tbl_Scoring_ScoreCard " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC, " & _
                        " ISNULL((SELECT SUM(Back9Net) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC, " & _
                        " ISNULL((SELECT SUM(Front9Net) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        iPlayCnt = iPlayCnt + 1
                        If CDbl(iPlayCnt) <= CDbl(TeamPlayer2Cnt) Then
                            dTeamTotal = dTeamTotal + CDbl(rt!Points)
                            dB9Total = dB9Total + CDbl(rt!B9)
                            dF9Total = dF9Total + CDbl(rt!F9)
                        End If
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ConnOmega.Execute "INSERT INTO " & TableName & " " & _
                                      " (TeamName, AveHandicap, TeamKey, TeamIndex) " & _
                                      " VALUES ('" & FORMATSQL(ra!TeamName) & "', " & _
                                      " " & ra!TeamHDCP & ", " & ra!PK & ", " & _
                                      " " & ra!TeamIndex & ")"
                    t = "SELECT PK " & _
                        " FROM " & TableName & " " & _
                        " WHERE (TeamKey = " & ra!PK & ")"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    If rt.RecordCount > 0 Then
                        
                        ConnOmega.Execute "UPDATE " & TableName & " " & _
                                          " SET TeamTotal = " & CDbl(dTeamTotal) & ", " & _
                                          " Back9 = " & CDbl(dB9Total) & ", " & _
                                          " Front9 = " & CDbl(dF9Total) & " " & _
                                          " WHERE (PK = " & rt!PK & ")"
                        
                        u = "SELECT dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName, dbo.tbl_Scoring_Team_Detail.PlayerKey, " & _
                            " ISNULL((SELECT SUM(NetPoints) AS Points From dbo.tbl_Scoring_ScoreCard WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) AS Points " & _
                            " FROM dbo.tbl_Scoring_PlayerName RIGHT OUTER JOIN " & _
                            " dbo.tbl_Scoring_Team_Detail ON dbo.tbl_Scoring_PlayerName.PK = dbo.tbl_Scoring_Team_Detail.PlayerKey RIGHT OUTER JOIN " & _
                            " dbo.tbl_Scoring_Team ON dbo.tbl_Scoring_Team_Detail.TeamKey = dbo.tbl_Scoring_Team.PK " & _
                            " Where (dbo.tbl_Scoring_Team.PK = " & ra!PK & ") " & _
                            " ORDER BY ISNULL((SELECT SUM(NetPoints) AS Points From dbo.tbl_Scoring_ScoreCard " & _
                            " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC, " & _
                            " ISNULL((SELECT SUM(Back9Net) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                            " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC, " & _
                            " ISNULL((SELECT SUM(Front9Net) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                            " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC"
                        If ru.State = adStateOpen Then ru.Close
                        ru.Open u, ConnOmega
                        While Not ru.EOF
                            iDLine = iDLine + 1
                            ConnOmega.Execute "INSERT INTO " & DetailTableName & " " & _
                                              " (MasterKey, Line, PlayerName) " & _
                                              " VALUES (" & rt!PK & ", " & iDLine & ", '" & FORMATSQL(ru!LastName & ",  " & ru!FirstName & "  " & ru!MiddleName) & "')"
                            
                            Arr = Split(TournamentRange, " - ", -1, 1)
                            iDay = DateDiff("d", Arr(0), Arr(1), vbMonday) + 1
                            
                            For i = 1 To iDay
                                dDate = DateAdd("d", CDbl(i - 1), Arr(0))
                                sFieldNum = "Day" & CStr(i)
                                v = "SELECT NetPoints as NetPoints " & _
                                    " From tbl_Scoring_ScoreCard " & _
                                    " WHERE (PlayerKey = " & ru!PlayerKey & ") " & _
                                    " AND (DDate = '" & FormatDateTime(dDate, vbShortDate) & "')"
                                If rv.State = adStateOpen Then rv.Close
                                rv.Open v, ConnOmega
                                If rv.RecordCount > 0 Then
                                    ConnOmega.Execute "UPDATE " & DetailTableName & " " & _
                                                      " SET " & sFieldNum & " = " & rv!NetPoints & " " & _
                                                      " WHERE (MasterKey = " & rt!PK & ") " & _
                                                      " AND (Line = " & iDLine & ")"
                                End If
                                rv.Close
                            Next i
                            
                            If CDbl(iDLine) > CDbl(TeamPlayer2Cnt) Then
                                dLastMan = dLastMan + CDbl(ru!Points)
                            End If
                            ru.MoveNext
                        Wend
                        ru.Close
                        
                        ConnOmega.Execute "UPDATE " & TableName & " " & _
                                          " SET LastPlayer = " & CDbl(dLastMan) & " " & _
                                          " WHERE (PK = " & rt!PK & ")"
                        
                    End If
                    rt.Close
                    
                    UpdateProgress picProgressBar, j / ra.RecordCount
                    ra.MoveNext
                Wend
                ra.Close
            
            Case 1 'Gross Points
                
                j = 0
                picPrint.Visible = False
                picProgress.Visible = True
                picProgressBar.BackColor = &HFFFFFF
                DoEvents
                              
                s = "SELECT PK, TeamName, TeamHDCP, TeamIndex " & _
                    " From tbl_Scoring_Team " & _
                    " WHERE (TournamentKey = " & TournamentKey & ")"
                If ra.State = adStateOpen Then ra.Close
                ra.Open s, ConnOmega
                While Not ra.EOF
                    DoEvents
                    j = j + 1
                    
                    iPlayCnt = 0: dTeamTotal = 0: dB9Total = 0: dF9Total = 0: iDLine = 0: dLastMan = 0
                    
                    t = "SELECT ISNULL ((SELECT SUM(GrossPoints) AS Points From dbo.tbl_Scoring_ScoreCard " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) AS Points, " & _
                        " ISNULL((SELECT SUM(Back9Gross) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) as B9, " & _
                        " ISNULL((SELECT SUM(Front9Gross) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) as F9 " & _
                        " FROM dbo.tbl_Scoring_Team LEFT OUTER JOIN " & _
                        " dbo.tbl_Scoring_Team_Detail ON dbo.tbl_Scoring_Team.PK = dbo.tbl_Scoring_Team_Detail.TeamKey " & _
                        " Where (dbo.tbl_Scoring_Team.PK = " & ra!PK & ") " & _
                        " ORDER BY ISNULL((SELECT SUM(GrossPoints) AS Points From dbo.tbl_Scoring_ScoreCard " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC, " & _
                        " ISNULL((SELECT SUM(Back9Gross) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC, " & _
                        " ISNULL((SELECT SUM(Front9Gross) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        iPlayCnt = iPlayCnt + 1
                        If CDbl(iPlayCnt) <= CDbl(TeamPlayer2Cnt) Then
                            dTeamTotal = dTeamTotal + CDbl(rt!Points)
                            dB9Total = dB9Total + CDbl(rt!B9)
                            dF9Total = dF9Total + CDbl(rt!F9)
                        End If
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ConnOmega.Execute "INSERT INTO " & TableName & " " & _
                                      " (TeamName, AveHandicap, TeamKey, TeamIndex) " & _
                                      " VALUES ('" & FORMATSQL(ra!TeamName) & "', " & _
                                    " " & ra!TeamHDCP & ", " & ra!PK & ", " & _
                                      " " & ra!TeamIndex & ")"
                    t = "SELECT PK " & _
                        " FROM " & TableName & " " & _
                        " WHERE (TeamKey = " & ra!PK & ")"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    If rt.RecordCount > 0 Then
                        
                        ConnOmega.Execute "UPDATE " & TableName & " " & _
                                          " SET TeamTotal = " & CDbl(dTeamTotal) & ", " & _
                                          " Back9 = " & CDbl(dB9Total) & ", " & _
                                          " Front9 = " & CDbl(dF9Total) & " " & _
                                          " WHERE (PK = " & rt!PK & ")"
                        
                        u = "SELECT dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName, dbo.tbl_Scoring_Team_Detail.PlayerKey, " & _
                            " ISNULL((SELECT SUM(GrossPoints) AS Points From dbo.tbl_Scoring_ScoreCard WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) AS Points " & _
                            " FROM dbo.tbl_Scoring_PlayerName RIGHT OUTER JOIN " & _
                            " dbo.tbl_Scoring_Team_Detail ON dbo.tbl_Scoring_PlayerName.PK = dbo.tbl_Scoring_Team_Detail.PlayerKey RIGHT OUTER JOIN " & _
                            " dbo.tbl_Scoring_Team ON dbo.tbl_Scoring_Team_Detail.TeamKey = dbo.tbl_Scoring_Team.PK " & _
                            " Where (dbo.tbl_Scoring_Team.PK = " & ra!PK & ") " & _
                            " ORDER BY ISNULL((SELECT SUM(GrossPoints) AS Points From dbo.tbl_Scoring_ScoreCard " & _
                            " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC, " & _
                            " ISNULL((SELECT SUM(Back9Gross) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                            " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC, " & _
                            " ISNULL((SELECT SUM(Front9Gross) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                            " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC"
                        If ru.State = adStateOpen Then ru.Close
                        ru.Open u, ConnOmega
                        While Not ru.EOF
                            iDLine = iDLine + 1
                            ConnOmega.Execute "INSERT INTO " & DetailTableName & " " & _
                                              " (MasterKey, Line, PlayerName) " & _
                                              " VALUES (" & rt!PK & ", " & iDLine & ", '" & FORMATSQL(ru!LastName & ",  " & ru!FirstName & "  " & ru!MiddleName) & "')"
                            
                            Arr = Split(TournamentRange, " - ", -1, 1)
                            iDay = DateDiff("d", Arr(0), Arr(1), vbMonday) + 1
                            
                            For i = 1 To iDay
                                dDate = DateAdd("d", CDbl(i - 1), Arr(0))
                                sFieldNum = "Day" & CStr(i)
                                v = "SELECT GrossPoints as NetPoints " & _
                                    " From tbl_Scoring_ScoreCard " & _
                                    " WHERE (PlayerKey = " & ru!PlayerKey & ") " & _
                                    " AND (DDate = '" & FormatDateTime(dDate, vbShortDate) & "')"
                                If rv.State = adStateOpen Then rv.Close
                                rv.Open v, ConnOmega
                                If rv.RecordCount > 0 Then
                                    ConnOmega.Execute "UPDATE " & DetailTableName & " " & _
                                                      " SET " & sFieldNum & " = " & rv!NetPoints & " " & _
                                                      " WHERE (MasterKey = " & rt!PK & ") " & _
                                                      " AND (Line = " & iDLine & ")"
                                End If
                                rv.Close
                            Next i
                            
                            If CDbl(iDLine) > CDbl(TeamPlayer2Cnt) Then
                                dLastMan = dLastMan + CDbl(ru!Points)
                            End If
                            ru.MoveNext
                        Wend
                        ru.Close
                        
                        ConnOmega.Execute "UPDATE " & TableName & " " & _
                                          " SET LastPlayer = " & CDbl(dLastMan) & " " & _
                                          " WHERE (PK = " & rt!PK & ")"
                        
                    End If
                    rt.Close
                    
                    UpdateProgress picProgressBar, j / ra.RecordCount
                    ra.MoveNext
                Wend
                ra.Close
                
            End Select
            
            picProgressBar.BackColor = &HFFFFFF
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
                .Name = "Top " & CStr(Trim(txtTop.Text))
            End With
            With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                RowCnt = RowCnt + 1
                HeaderRow = HeaderRow + 1
                ColCnt = 0
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = TournamentName
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 10
                .Range(strRange).Font.Bold = True

                RowCnt = RowCnt + 1
                HeaderRow = HeaderRow + 1
                ColCnt = 0
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = "Range : " & TournamentRange
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False

                RowCnt = RowCnt + 1
                HeaderRow = HeaderRow + 1
                ColCnt = 0
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                If cmbGroup.ListIndex = 0 Then
                    .Range(strRange).Value = "Team (Net Points)"
                ElseIf cmbGroup.ListIndex = 1 Then
                    .Range(strRange).Value = "Team (Gross Points)"
                End If
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                
                If cmbDivision.ListIndex > 0 Then
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Division : " & cmbDivision.List(cmbDivision.ListIndex)
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                End If
                
                RowCnt = RowCnt + 1
                HeaderRow = HeaderRow + 1
                ColCnt = 0
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = ""
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = True

                RowCnt = RowCnt + 1
                HeaderRow = HeaderRow + 1
                ColCnt = 0
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = "TeamName"
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = True
                .Columns(ColCnt).ColumnWidth = 15
                .Range(strRange).HorizontalAlignment = 3
                .Range(strRange).Select
                xlsApp.Selection.Borders.LineStyle = 1
                
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                If TeamAverage = 1 Then
                    .Range(strRange).Value = "Ave. Handicap"
                ElseIf TeamAverage = 2 Then
                    .Range(strRange).Value = "Ave. Index"
                End If
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Columns(ColCnt).ColumnWidth = 9
                .Range(strRange).Font.Bold = True
                .Range(strRange).HorizontalAlignment = 3
                .Range(strRange).Select
                xlsApp.Selection.Borders.LineStyle = 1
                    
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = "Name"
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Columns(ColCnt).ColumnWidth = 25
                .Range(strRange).Font.Bold = True
                .Range(strRange).Select
                xlsApp.Selection.Borders.LineStyle = 1
                
                Arr = Split(TournamentRange, " - ", -1, 1)
                iDay = 0
                For i = 0 To DateDiff("d", Arr(0), Arr(1), vbMonday)
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Day " & i + 1
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    iDay = iDay + 1
                Next i

                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = "Total"
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = True
                .Range(strRange).HorizontalAlignment = 4
                .Range(strRange).Select
                xlsApp.Selection.Borders.LineStyle = 1
                
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = "Team Total"
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = True
                .Columns(ColCnt).ColumnWidth = 10
                .Range(strRange).HorizontalAlignment = 4
                .Range(strRange).Select
                xlsApp.Selection.Borders.LineStyle = 1

                j = 0
                If cmbDivision.ListIndex = 0 Or _
                cmbDivision.ListIndex = -1 Then
                
                    s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " PK, TeamName, AveHandicap, TeamIndex, TeamTotal " & _
                        " From " & TableName & " " & _
                        " ORDER BY TeamTotal DESC, LastPlayer DESC, Back9 DESC, Front9 DESC"
                Else
                    If TeamAverage = 1 Then
                        s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " PK, TeamName, AveHandicap, TeamIndex, TeamTotal, " & _
                            " ISNULL((SELECT Class From dbo.tbl_Scoring_TournamentInfo_Class " & _
                            " WHERE (TournamentKey = " & TournamentKey & ") " & _
                            " AND (HFrom <= dbo." & TableName & ".AveHandicap) " & _
                            " AND (HTo >= dbo." & TableName & ".AveHandicap)), '') AS Class " & _
                            " From " & TableName & " " & _
                            " WHERE (ISNULL((SELECT Class FROM dbo.tbl_Scoring_TournamentInfo_Class AS tbl_Scoring_TournamentInfo_Class_1 " & _
                            " WHERE (TournamentKey = " & TournamentKey & ") " & _
                            " AND (HFrom <= dbo." & TableName & ".AveHandicap) " & _
                            " AND (HTo >= dbo." & TableName & ".AveHandicap)), '') = '" & cmbDivision.List(cmbDivision.ListIndex) & "') " & _
                            " ORDER BY TeamTotal DESC, LastPlayer DESC, Back9 DESC, Front9 DESC"
                            
                    ElseIf TeamAverage = 2 Then
                        
                        s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " PK, TeamName, AveHandicap, TeamIndex, TeamTotal, " & _
                            " ISNULL((SELECT Class From dbo.tbl_Scoring_TournamentInfo_Index " & _
                            " WHERE (TournamentKey = " & TournamentKey & ") " & _
                            " AND (HFrom <= dbo." & TableName & ".TeamIndex) " & _
                            " AND (HTo >= dbo." & TableName & ".TeamIndex)), '') AS Class " & _
                            " From " & TableName & " " & _
                            " WHERE (ISNULL((SELECT Class FROM dbo.tbl_Scoring_TournamentInfo_Index AS tbl_Scoring_TournamentInfo_Index_1 " & _
                            " WHERE (TournamentKey = " & TournamentKey & ") " & _
                            " AND (HFrom <= dbo." & TableName & ".TeamIndex) " & _
                            " AND (HTo >= dbo." & TableName & ".TeamIndex)), '') = '" & cmbDivision.List(cmbDivision.ListIndex) & "') " & _
                            " ORDER BY TeamTotal DESC, LastPlayer DESC, Back9 DESC, Front9 DESC"
                    End If
                End If
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                While Not rs.EOF
                    DoEvents
                    iProgressValue = iProgressValue + 1
                    j = j + 1
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = rs!TeamName
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = rs!TeamIndex
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCntTmp = RowCnt - 1
                    
                    t = "SELECT * " & _
                        " FROM " & DetailTableName & " " & _
                        " WHERE (MasterKey = " & rs!PK & ") " & _
                        " ORDER BY Line"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    l = 0
                    dTeamPoints = 0
                    While Not rt.EOF
                        l = l + 1
                        ColCntTmp = ColCnt
                        RowCntTmp = RowCntTmp + 1
                        ColCntTmp = ColCntTmp + 1
                        strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                        .Range(strRange).Value = rt!PlayerName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        iDay = rt.Fields.Count - 2
                        For i = 3 To iDay
                            ColCntTmp = ColCntTmp + 1
                            strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                            .Range(strRange).Value = rt.Fields(i).Value
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                        Next i
                        
                        ColCntTmp = ColCntTmp + 1
                        strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                        .Range(strRange).Value = rt!Total
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ColCnt = ColCntTmp
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    strRangeFrom = EXCEL_RANGE(ColCnt, RowCnt)
                    strRangeTo = EXCEL_RANGE(ColCnt, RowCntTmp)
                    .Range(strRange).Value = rs!TeamTotal
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRangeFrom, strRangeTo).Select
                    xlsApp.Selection.Merge
                    .Range(strRange).VerticalAlignment = 2
                    .Range(strRange).HorizontalAlignment = 3
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    strRange = EXCEL_RANGE(1, RowCnt)
                    strRangeFrom = EXCEL_RANGE(1, RowCnt)
                    strRangeTo = EXCEL_RANGE(1, RowCntTmp)
                    .Range(strRangeFrom, strRangeTo).Select
                    xlsApp.Selection.Merge
                    .Range(strRange).VerticalAlignment = 2
                    .Range(strRange).HorizontalAlignment = 3
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    strRange = EXCEL_RANGE(2, RowCnt)
                    strRangeFrom = EXCEL_RANGE(2, RowCnt)
                    strRangeTo = EXCEL_RANGE(2, RowCntTmp)
                    .Range(strRangeFrom, strRangeTo).Select
                    xlsApp.Selection.Merge
                    .Range(strRange).VerticalAlignment = 2
                    .Range(strRange).HorizontalAlignment = 3
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    RowCnt = RowCntTmp
                    
                    UpdateProgress_Caption "Exporting to Excel", picProgressBar, iProgressValue / rs.RecordCount
                    
                    rs.MoveNext
                Wend
                rs.Close
                
                .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
                
            End With
        
SAVING3:
            On Error GoTo err_saving3:
            If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
            xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

            xlsApp.Visible = True
                    
    Case 2  'Result
        
        Exit Sub
    
    Case 3  'Scores
        
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
            RowCnt = RowCnt + 1
            HeaderRow = HeaderRow + 1
            ColCnt = 0
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = TournamentName
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 10
            .Range(strRange).Font.Bold = True

            RowCnt = RowCnt + 1
            HeaderRow = HeaderRow + 1
            ColCnt = 0
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = "Range : " & TournamentRange
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = False
            
            s = "SELECT dbo.tbl_Scoring_System.ScoringSystem " & _
                " FROM dbo.tbl_Scoring_TournamentInfo LEFT OUTER JOIN " & _
                " dbo.tbl_Scoring_System ON dbo.tbl_Scoring_TournamentInfo.Scoring = dbo.tbl_Scoring_System.PK " & _
                " WHERE (dbo.tbl_Scoring_TournamentInfo.PK = " & TournamentKey & ")"
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
            ColCnt = 0
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = ""
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = False
            
            RowCnt = RowCnt + 1
            HeaderRow = HeaderRow + 1
            ColCnt = 0
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = "#"
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = "Name"
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = "Date"
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = ""
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = False
            
            For i = 1 To 9
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = "'" & Format(i, "0#")
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                .Range(strRange).HorizontalAlignment = 3
            Next i
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = "F-9"
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = False
            .Range(strRange).HorizontalAlignment = 3
            
            For i = 1 To 9
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = "'" & Format(CDbl(i) + 9, "0#")
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                .Range(strRange).HorizontalAlignment = 3
            Next i
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = "B-9"
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = False
            .Range(strRange).HorizontalAlignment = 3
            
            strRange = EXCEL_RANGE(1, HeaderRow + 1)
            .Range(strRange).Select
            xlsApp.ActiveWindow.FreezePanes = True
            
            picProgressBar.Visible = True
            s = "SELECT dbo.tbl_Scoring_ScoreCard.PlayerKey, dbo.tbl_Scoring_PlayerName.LastName, " & _
                " dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName " & _
                " FROM dbo.tbl_Scoring_ScoreCard LEFT OUTER JOIN " & _
                " dbo.tbl_Scoring_PlayerName ON dbo.tbl_Scoring_ScoreCard.PlayerKey = dbo.tbl_Scoring_PlayerName.PK " & _
                " Where (dbo.tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") " & _
                " GROUP BY dbo.tbl_Scoring_ScoreCard.PlayerKey, dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName " & _
                " ORDER BY dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName"
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
                
                t = "SELECT PK, PlayerKey, DDate " & _
                    " From dbo.tbl_Scoring_ScoreCard " & _
                    " Where (PlayerKey = " & rs!PlayerKey & ") " & _
                    " ORDER BY DDate"
                If rt.State = adStateOpen Then rt.Close
                rt.Open t, ConnOmega
                If rt.RecordCount = 1 Then
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = Format(rt!dDate, "mm/dd/yyyy")
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    For i = 1 To 3
                        sMinRange = "": sMaxRange = ""
                        Select Case i
                            Case 1
                                ColCnt = ColCnt + 1
                                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                .Range(strRange).Value = "Score"
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                                u = "SELECT Score " & _
                                    " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                    " Where (ScoreCardKey = " & rt!PK & ") " & _
                                    " ORDER BY Hole"
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
                            Case 2
                                RowCnt = RowCnt + 1
                                ColCnt = 3
                                ColCnt = ColCnt + 1
                                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                .Range(strRange).Value = "Gross"
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                                u = "SELECT Gross " & _
                                    " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                    " Where (ScoreCardKey = " & rt!PK & ") " & _
                                    " ORDER BY Hole"
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
                            Case 3
                                RowCnt = RowCnt + 1
                                ColCnt = 3
                                ColCnt = ColCnt + 1
                                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                .Range(strRange).Value = "Net"
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                                u = "SELECT Net " & _
                                    " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                    " Where (ScoreCardKey = " & rt!PK & ") " & _
                                    " ORDER BY Hole"
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
                        End Select
                    Next i
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
                        For i = 1 To 3
                            sMinRange = "": sMaxRange = ""
                            Select Case i
                                Case 1
                                    ColCnt = ColCnt + 1
                                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                    .Range(strRange).Value = "Score"
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                    u = "SELECT Score " & _
                                        " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                        " Where (ScoreCardKey = " & rt!PK & ") " & _
                                        " ORDER BY Hole"
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
                                Case 2
                                    RowCnt = RowCnt + 1
                                    ColCnt = 3
                                    ColCnt = ColCnt + 1
                                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                    .Range(strRange).Value = "Gross"
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                    u = "SELECT Gross " & _
                                        " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                        " Where (ScoreCardKey = " & rt!PK & ") " & _
                                        " ORDER BY Hole"
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
                                Case 3
                                    RowCnt = RowCnt + 1
                                    ColCnt = 3
                                    ColCnt = ColCnt + 1
                                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                    .Range(strRange).Value = "Net"
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                    u = "SELECT Net " & _
                                        " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                        " Where (ScoreCardKey = " & rt!PK & ") " & _
                                        " ORDER BY Hole"
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
                            End Select
                        Next i
                        rt.MoveNext
                    Wend
                End If
                rt.Close
                
                
                UpdateProgress_Caption "Exporting to Excel", picProgressBar, iProgressValue / rs.RecordCount
                rs.MoveNext
            Wend
            rs.Close
            
            .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
            
        End With

SAVING4:
            On Error GoTo err_saving4:
            If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
            xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

            xlsApp.Visible = True
End Select

picProgress.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True

'cmdCancelPrint_Click

Exit Sub
err_saving1:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING1:

Exit Sub
err_saving2:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING2:

Exit Sub
err_saving3:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING3:

Exit Sub

err_saving4:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING4:

Exit Sub
'err_saving4:
'MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
'GoTo SAVING4:
End Sub

Private Sub TimerReportMolave_Timer()
TimerReportMolave.Enabled = False
Select Case cmbReportType.ListIndex
    Case 0  'INDIVIDUAL
        Select Case cmbGroup.ListIndex
            Case 0  'NetPoints
                
                picPrint.Enabled = False
                picProgress.ZOrder 0
                picProgressBar.BackColor = &HFFFFFF
                picProgress.Visible = True
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
                    .Name = "Top " & CStr(Trim(txtTop.Text))
                End With
                With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = TournamentName
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 10
                    .Range(strRange).Font.Bold = True
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Range : " & TournamentRange
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    If cmbDivisionStableford.ListIndex = 0 Then
                        .Range(strRange).Value = "Individual (Net Points) [" & IIf(cmbGender.ListIndex = 0, "MALE", "FEMALE") & "]"
                    Else
                        .Range(strRange).Value = "Individual [Class " & cmbDivisionStableford.List(cmbDivisionStableford.ListIndex) & "] (Net Points) [" & IIf(cmbGender.ListIndex = 0, "MALE", "FEMALE") & "]"
                    End If
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = ""
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "#"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCnt).ColumnWidth = 3
                    .Range(strRange).HorizontalAlignment = 4
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Name"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Handicap"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    
                    Arr = Split(TournamentRange, " - ", -1, 1)
                    iDay = 0
                    For i = 0 To DateDiff("d", Arr(0), Arr(1), vbMonday)
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = "Day " & i + 1
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).HorizontalAlignment = 4
                        iDay = iDay + 1
                    Next i
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Total"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    j = 0
                    If cmbDivision.ListIndex = 0 Then
                        If IsDate(txtDatePrint.Text) = True Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Net) AS Holes_B9, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") " & _
                                " AND (tbl_Scoring_ScoreCard.DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "') " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.NetPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Net) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Net) AS Holes_B9, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.NetPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Net) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        End If
                    Else
                        If IsDate(txtDatePrint.Text) = True Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Net) AS Holes_B9, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " AND (tbl_Scoring_ScoreCard.DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "') " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.NetPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Net) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Net) AS Holes_B9, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.NetPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Net) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        End If
                    End If
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open s, ConnOmega
                    While Not rs.EOF
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
                        .Range(strRange).Value = rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        .Columns(ColCnt).ColumnWidth = 25
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = rs!Handicap
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        sTotal = "="
                        For i = 1 To iDay
                            dDate = DateAdd("d", CDbl(i - 1), Arr(0))
                            ColCnt = ColCnt + 1
                            strRange = EXCEL_RANGE(ColCnt, RowCnt)
                            sTotal = sTotal & strRange & "+"
                            If IsDate(txtDatePrint.Text) = True Then
                                If DateValue(dDate) = DateValue(FormatDateTime(txtDatePrint.Text, vbShortDate)) Then
                                    t = "SELECT NetPoints " & _
                                        " From tbl_Scoring_ScoreCard " & _
                                        " WHERE (PlayerKey = " & rs!PlayerKey & ") " & _
                                        " AND (DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "')"
                                    If rt.State = adStateOpen Then rt.Close
                                    rt.Open t, ConnOmega
                                    If rt.RecordCount > 0 Then
                                        .Range(strRange).Value = rt!NetPoints
                                        .Range(strRange).Font.Name = "Tahoma"
                                        .Range(strRange).Font.Size = 8
                                        .Range(strRange).Font.Bold = False
                                    Else
                                        .Range(strRange).Value = ""
                                        .Range(strRange).Font.Name = "Tahoma"
                                        .Range(strRange).Font.Size = 8
                                        .Range(strRange).Font.Bold = False
                                    End If
                                    rt.Close
                                Else
                                    .Range(strRange).Value = ""
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                End If
                            Else
                                t = "SELECT NetPoints " & _
                                    " From tbl_Scoring_ScoreCard " & _
                                    " WHERE (PlayerKey = " & rs!PlayerKey & ") " & _
                                    " AND (DDate = '" & FormatDateTime(dDate, vbShortDate) & "')"
                                If rt.State = adStateOpen Then rt.Close
                                rt.Open t, ConnOmega
                                If rt.RecordCount > 0 Then
                                    .Range(strRange).Value = rt!NetPoints
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                Else
                                    .Range(strRange).Value = ""
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                End If
                                rt.Close
                            End If
                            
                        Next i
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = Mid(sTotal, 1, Len(sTotal) - 1)
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        UpdateProgress picProgressBar, j / rs.RecordCount
                        
                        rs.MoveNext
                    Wend
                    rs.Close
                    
                    .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
                    
                End With
                
SAVING1:
                On Error GoTo err_saving1:
                If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
                xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName
                
                xlsApp.Visible = True
                
                picProgress.Visible = False
                picPrint.Enabled = True
                
            Case 1  'GrossPoints
            
                picPrint.Enabled = False
                picProgress.ZOrder 0
                picProgressBar.BackColor = &HFFFFFF
                picProgress.Visible = True
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
                    .Name = "Top " & CStr(Trim(txtTop.Text))
                End With
                With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = TournamentName
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 10
                    .Range(strRange).Font.Bold = True
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Range : " & TournamentRange
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    If cmbDivision.ListIndex = 0 Then
                        .Range(strRange).Value = "Individual (Gross Points) [" & IIf(cmbGender.ListIndex = 0, "MALE", "FEMALE") & "]"
                    Else
                        .Range(strRange).Value = "Individual [Class " & cmbDivision.List(cmbDivision.ListIndex) & "] (Gross Points) [" & IIf(cmbGender.ListIndex = 0, "MALE", "FEMALE") & "]"
                    End If
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = ""
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "#"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCnt).ColumnWidth = 3
                    .Range(strRange).HorizontalAlignment = 4
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Name"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Handicap"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    
                    Arr = Split(TournamentRange, " - ", -1, 1)
                    iDay = 0
                    For i = 0 To DateDiff("d", Arr(0), Arr(1), vbMonday)
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = "Day " & i + 1
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).HorizontalAlignment = 4
                        iDay = iDay + 1
                    Next i
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Total"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    
                    j = 0
                    If cmbDivision.ListIndex = 0 Then
                        If IsDate(txtDatePrint.Text) = True Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Gross) AS Holes_B9, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") " & _
                                " AND (tbl_Scoring_ScoreCard.DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "') " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.GrossPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Gross) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Gross) AS Holes_B9, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.GrossPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Gross) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        End If
                    Else
                        If IsDate(txtDatePrint.Text) = True Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Gross) AS Holes_B9, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ")  AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " AND (tbl_Scoring_ScoreCard.DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "') " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.GrossPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Gross) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Gross) AS Holes_B9, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ")  AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.GrossPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Gross) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        End If
                    End If
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open s, ConnOmega
                    While Not rs.EOF
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
                        .Range(strRange).Value = rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        .Columns(ColCnt).ColumnWidth = 25
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = rs!Handicap
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        sTotal = "="
                        For i = 1 To iDay
                            dDate = DateAdd("d", CDbl(i - 1), Arr(0))
                            ColCnt = ColCnt + 1
                            strRange = EXCEL_RANGE(ColCnt, RowCnt)
                            sTotal = sTotal & strRange & "+"
                            If IsDate(txtDatePrint.Text) = True Then
                                If DateValue(dDate) = DateValue(FormatDateTime(txtDatePrint.Text, vbShortDate)) Then
                                    t = "SELECT GrossPoints as NetPoints " & _
                                    " From tbl_Scoring_ScoreCard " & _
                                    " WHERE (PlayerKey = " & rs!PlayerKey & ") " & _
                                    " AND (DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "')"
                                    If rt.State = adStateOpen Then rt.Close
                                    rt.Open t, ConnOmega
                                    If rt.RecordCount > 0 Then
                                        .Range(strRange).Value = rt!NetPoints
                                        .Range(strRange).Font.Name = "Tahoma"
                                        .Range(strRange).Font.Size = 8
                                        .Range(strRange).Font.Bold = False
                                    Else
                                        .Range(strRange).Value = ""
                                        .Range(strRange).Font.Name = "Tahoma"
                                        .Range(strRange).Font.Size = 8
                                        .Range(strRange).Font.Bold = False
                                    End If
                                    rt.Close
                                Else
                                    .Range(strRange).Value = ""
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                End If
                            Else
                                t = "SELECT GrossPoints as NetPoints " & _
                                    " From tbl_Scoring_ScoreCard " & _
                                    " WHERE (PlayerKey = " & rs!PlayerKey & ") " & _
                                    " AND (DDate = '" & FormatDateTime(dDate, vbShortDate) & "')"
                                If rt.State = adStateOpen Then rt.Close
                                rt.Open t, ConnOmega
                                If rt.RecordCount > 0 Then
                                    .Range(strRange).Value = rt!NetPoints
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                Else
                                    .Range(strRange).Value = ""
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                End If
                                rt.Close
                            End If
                        Next i
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = Mid(sTotal, 1, Len(sTotal) - 1)
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        UpdateProgress picProgressBar, j / rs.RecordCount
                        
                        rs.MoveNext
                    Wend
                    rs.Close

                    .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
                    
                End With
SAVING2:
                On Error GoTo err_saving2:
                If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
                xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName
                
                xlsApp.Visible = True
                
                picProgress.Visible = False
                picPrint.Enabled = True
                
            Case 2      'Gross Score
                
                picPrint.Enabled = False
                picProgress.ZOrder 0
                picProgressBar.BackColor = &HFFFFFF
                picProgress.Visible = True
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
                    .Name = "Top " & CStr(Trim(txtTop.Text))
                End With
                With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = TournamentName
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 10
                    .Range(strRange).Font.Bold = True
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Range : " & TournamentRange
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    If cmbDivision.ListIndex = 0 Then
                        .Range(strRange).Value = "Individual (Gross Score) [" & IIf(cmbGender.ListIndex = 0, "MALE", "FEMALE") & "]"
                    Else
                        .Range(strRange).Value = "Individual [Class " & cmbDivision.List(cmbDivision.ListIndex) & "] (Gross Score) [" & IIf(cmbGender.ListIndex = 0, "MALE", "FEMALE") & "]"
                    End If
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = ""
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "#"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCnt).ColumnWidth = 3
                    .Range(strRange).HorizontalAlignment = 4
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Name"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Handicap"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    
                    Arr = Split(TournamentRange, " - ", -1, 1)
                    iDay = 0
                    For i = 0 To DateDiff("d", Arr(0), Arr(1), vbMonday)
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = "Day " & i + 1
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).HorizontalAlignment = 4
                        iDay = iDay + 1
                    Next i
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Total"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    
                    j = 0
                    If cmbDivision.ListIndex = 0 Then
                        If IsDate(txtDatePrint.Text) = True Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.Score) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Score) AS Holes_B9, (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") " & _
                                " AND (tbl_Scoring_ScoreCard.DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "') " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.Score), SUM(tbl_Scoring_ScoreCard.Back9Score) , (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)), (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) "
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.Score) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Score) AS Holes_B9, (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.Score), SUM(tbl_Scoring_ScoreCard.Back9Score) , (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)), (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) "
                        End If
                    Else
                        If IsDate(txtDatePrint.Text) = True Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.Score) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Score) AS Holes_B9, (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ")  AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " AND (tbl_Scoring_ScoreCard.DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "') " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.Score) , SUM(tbl_Scoring_ScoreCard.Back9Score) , (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) , (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) "
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.Score) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Score) AS Holes_B9, (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ")  AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.Score) , SUM(tbl_Scoring_ScoreCard.Back9Score) , (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) , (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) "
                        End If
                    End If
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open s, ConnOmega
                    While Not rs.EOF
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
                        .Range(strRange).Value = rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        .Columns(ColCnt).ColumnWidth = 25
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = rs!Handicap
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        sTotal = "="
                        For i = 1 To iDay
                            dDate = DateAdd("d", CDbl(i - 1), Arr(0))
                            ColCnt = ColCnt + 1
                            strRange = EXCEL_RANGE(ColCnt, RowCnt)
                            sTotal = sTotal & strRange & "+"
                            If IsDate(txtDatePrint.Text) = True Then
                                If DateValue(dDate) = DateValue(FormatDateTime(txtDatePrint.Text, vbShortDate)) Then
                                    t = "SELECT Score as NetPoints " & _
                                        " From tbl_Scoring_ScoreCard " & _
                                        " WHERE (PlayerKey = " & rs!PlayerKey & ") " & _
                                        " AND (DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "')"
                                    If rt.State = adStateOpen Then rt.Close
                                    rt.Open t, ConnOmega
                                    If rt.RecordCount > 0 Then
                                        .Range(strRange).Value = rt!NetPoints
                                        .Range(strRange).Font.Name = "Tahoma"
                                        .Range(strRange).Font.Size = 8
                                        .Range(strRange).Font.Bold = False
                                    Else
                                        .Range(strRange).Value = ""
                                        .Range(strRange).Font.Name = "Tahoma"
                                        .Range(strRange).Font.Size = 8
                                        .Range(strRange).Font.Bold = False
                                    End If
                                    rt.Close
                                Else
                                    .Range(strRange).Value = ""
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                End If
                            Else
                                t = "SELECT Score as NetPoints " & _
                                    " From tbl_Scoring_ScoreCard " & _
                                    " WHERE (PlayerKey = " & rs!PlayerKey & ") " & _
                                    " AND (DDate = '" & FormatDateTime(dDate, vbShortDate) & "')"
                                If rt.State = adStateOpen Then rt.Close
                                rt.Open t, ConnOmega
                                If rt.RecordCount > 0 Then
                                    .Range(strRange).Value = rt!NetPoints
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                Else
                                    .Range(strRange).Value = ""
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                End If
                                rt.Close
                            End If
                        Next i
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = Mid(sTotal, 1, Len(sTotal) - 1)
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        UpdateProgress picProgressBar, j / rs.RecordCount
                        
                        rs.MoveNext
                    Wend
                    rs.Close
                    
                    .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
                    
                End With
SAVING6:
                On Error GoTo err_saving6:
                If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
                xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName
                
                xlsApp.Visible = True
                
                picProgress.Visible = False
                picPrint.Enabled = True
                
        End Select
        
    Case 1  'TEAM
    
        Select Case cmbGroup.ListIndex
            Case 0  'NetPoints
                
                picPrint.Enabled = False
                picProgress.ZOrder 0
                picProgressBar.BackColor = &HFFFFFF
                picProgress.Visible = True
                DoEvents
                
                ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_Team_Rep " & _
                                  " WHERE (LogInName = '" & gbl_UserName & "')"
                
                If cmbDivision.ListIndex = 0 Then
                    s = "SELECT PK, TeamName " & _
                        " From tbl_Scoring_Team " & _
                        " WHERE (TournamentKey = " & TournamentKey & ")"
                Else
                    s = "SELECT PK, TeamName " & _
                        " From tbl_Scoring_Team " & _
                        " WHERE ((SELECT Class " & _
                        " FROM tbl_Scoring_TournamentInfo_Class AS tbl_Scoring_TournamentInfo_Class_1 " & _
                        " WHERE (HFrom <= tbl_Scoring_Team.TeamHDCP) AND (HTo >= tbl_Scoring_Team.TeamHDCP) " & _
                        " AND (TournamentKey = " & TournamentKey & ")) = '" & cmbDivision.List(cmbDivision.ListIndex) & "') " & _
                        " AND (TournamentKey = " & TournamentKey & ")"
                End If
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                While Not rs.EOF
                    ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_Team_Rep " & _
                                      " (LogInName, TeamKey, TeamName) " & _
                                      " VALUES ('" & gbl_UserName & "', " & rs!PK & ", '" & FORMATSQL(rs!TeamName) & "')"
                    
                    dTotalTeam = 0
                    t = "SELECT  TOP " & TeamPlayer2Cnt & " tbl_Scoring_Team_Detail.PlayerKey, " & _
                        " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                        " tbl_Scoring_PlayerName.MiddleName, " & _
                        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS NetPoints " & _
                        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                        " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                        " Order By ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) DESC"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        dTotalTeam = dTotalTeam + CDbl(rt!NetPoints)
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_Team_Rep " & _
                                      " SET Score = " & CDbl(dTotalTeam) & " " & _
                                      " WHERE (LogInName = '" & gbl_UserName & "') " & _
                                      " AND (TeamKey = " & rs!PK & ")"
                    
                    dTotalTeam = 0: iTeamCounter = 0
                    t = "SELECT tbl_Scoring_Team_Detail.PlayerKey, " & _
                        " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                        " tbl_Scoring_PlayerName.MiddleName, " & _
                        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS NetPoints " & _
                        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                        " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                        " Order By ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) DESC"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        iTeamCounter = iTeamCounter + 1
                        If CDbl(iTeamCounter) > CDbl(TeamPlayer2Cnt) Then
                            dTotalTeam = dTotalTeam + CDbl(rt!NetPoints)
                        End If
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_Team_Rep " & _
                                      " SET CountBck = " & CDbl(dTotalTeam) & " " & _
                                      " WHERE (LogInName = '" & gbl_UserName & "') " & _
                                      " AND (TeamKey = " & rs!PK & ")"
                    
                    rs.MoveNext
                Wend
                rs.Close
                
                
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
                    .Name = "Top " & CStr(Trim(txtTop.Text))
                End With
                With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = TournamentName
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 10
                    .Range(strRange).Font.Bold = True

                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Range : " & TournamentRange
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    If cmbDivision.ListIndex > 0 Then
                        RowCnt = RowCnt + 1
                        ColCnt = 0
                        ColCnt = ColCnt + 1
                        HeaderRow = HeaderRow + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = "CLASS " & cmbDivision.List(cmbDivision.ListIndex)
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                    End If
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Team (Net Points)"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False

                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = ""
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True

                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "#"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCnt).ColumnWidth = 3
                    .Range(strRange).HorizontalAlignment = 4
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                        
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Name"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Columns(ColCnt).ColumnWidth = 25
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    Arr = Split(TournamentRange, " - ", -1, 1)
                    iDay = 0
                    For i = 0 To DateDiff("d", Arr(0), Arr(1), vbMonday)
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = "Day " & i + 1
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).HorizontalAlignment = 4
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        iDay = iDay + 1
                    Next i

                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Total"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Team Total"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCnt).ColumnWidth = 10
                    .Range(strRange).HorizontalAlignment = 4
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1

                    j = 0
                    
                    s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " TeamKey as PK, TeamName, Score  " & _
                        " FROM tbl_Scoring_ScoreCard_Team_Rep " & _
                        " WHERE (LogInName = '" & gbl_UserName & "') " & _
                        " ORDER BY Score DESC, CountBck DESC"
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open s, ConnOmega
                    While Not rs.EOF
                        j = j + 1
                        RowCnt = RowCnt + 1
                        ColCnt = 0
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = j
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        RowCntTmp = RowCnt - 1
                        
                        ColCntTmp = ColCnt
                        RowCntTmp = RowCntTmp + 1
                        ColCntTmp = ColCntTmp + 1
                        strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                        .Range(strRange).Value = rs!TeamName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        For i = 1 To iDay
                            ColCntTmp = ColCntTmp + 1
                            strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                            .Range(strRange).Value = ""
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                        Next i
                        
                        t = "SELECT  tbl_Scoring_Team_Detail.PlayerKey, " & _
                            " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                            " tbl_Scoring_PlayerName.MiddleName, " & _
                            " (SELECT SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints " & _
                            " From tbl_Scoring_ScoreCard " & _
                            " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)) AS NetPoints " & _
                            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                            " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                            " Order By (SELECT SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints " & _
                            " From tbl_Scoring_ScoreCard " & _
                            " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)) DESC"
                        If rt.State = adStateOpen Then rt.Close
                        rt.Open t, ConnOmega
                        While Not rt.EOF
                            
                            
                            ColCntTmp = ColCnt
                            RowCntTmp = RowCntTmp + 1
                            ColCntTmp = ColCntTmp + 1
                            strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                            .Range(strRange).Value = rt!LastName & ",  " & rt!FirstName & "  " & rt!MiddleName
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                            sTotal = "="
                            For i = 1 To iDay
                                dDate = DateAdd("d", CDbl(i - 1), Arr(0))
                                ColCntTmp = ColCntTmp + 1
                                strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                                sTotal = sTotal & strRange & "+"
                                u = "SELECT NetPoints " & _
                                    " From tbl_Scoring_ScoreCard " & _
                                    " WHERE (PlayerKey = " & rt!PlayerKey & ") " & _
                                    " AND (DDate = '" & FormatDateTime(dDate, vbShortDate) & "')"
                                If ru.State = adStateOpen Then ru.Close
                                ru.Open u, ConnOmega
                                If ru.RecordCount > 0 Then
                                    .Range(strRange).Value = ru!NetPoints
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                Else
                                    .Range(strRange).Value = ""
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                End If
                                .Range(strRange).Select
                                xlsApp.Selection.Borders.LineStyle = 1
                                ru.Close
                            Next i
                            
                            ColCntTmp = ColCntTmp + 1
                            strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                            .Range(strRange).Value = Mid(sTotal, 1, Len(sTotal) - 1)
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                            rt.MoveNext
                        Wend
                        rt.Close
                        
                        ColCnt = ColCntTmp
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        strRangeFrom = EXCEL_RANGE(ColCnt, RowCnt)
                        strRangeTo = EXCEL_RANGE(ColCnt, RowCntTmp)
                        .Range(strRange).Value = rs!Score 'rs!NetPoints
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRangeFrom, strRangeTo).Select
                        xlsApp.Selection.Merge
                        .Range(strRange).VerticalAlignment = 2
                        .Range(strRange).HorizontalAlignment = 3
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        strRange = EXCEL_RANGE(1, RowCnt)
                        strRangeFrom = EXCEL_RANGE(1, RowCnt)
                        strRangeTo = EXCEL_RANGE(1, RowCntTmp)
                        .Range(strRangeFrom, strRangeTo).Select
                        xlsApp.Selection.Merge
                        .Range(strRange).VerticalAlignment = 2
                        .Range(strRange).HorizontalAlignment = 3
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        RowCnt = RowCntTmp
                        
                        UpdateProgress picProgressBar, j / rs.RecordCount
                        
                        rs.MoveNext
                    Wend
                    rs.Close
                    
                    .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
                    
                End With

SAVING3:
                On Error GoTo err_saving3:
                If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
                xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

                xlsApp.Visible = True
                
                picProgress.Visible = False
                picPrint.Enabled = True
                
            Case 1  'Gross Points
                
                picPrint.Enabled = False
                picProgress.ZOrder 0
                picProgressBar.BackColor = &HFFFFFF
                picProgress.Visible = True
                DoEvents
                
                ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_Team_Rep " & _
                                  " WHERE (LogInName = '" & gbl_UserName & "')"
                
                If cmbDivision.ListIndex = 0 Then
                    s = "SELECT PK, TeamName " & _
                        " From tbl_Scoring_Team " & _
                        " WHERE (TournamentKey = " & TournamentKey & ")"
                Else
                    s = "SELECT PK, TeamName " & _
                        " From tbl_Scoring_Team " & _
                        " WHERE ((SELECT Class " & _
                        " FROM tbl_Scoring_TournamentInfo_Class AS tbl_Scoring_TournamentInfo_Class_1 " & _
                        " WHERE (HFrom <= tbl_Scoring_Team.TeamHDCP) AND (HTo >= tbl_Scoring_Team.TeamHDCP) " & _
                        " AND (TournamentKey = " & TournamentKey & ")) = '" & cmbDivision.List(cmbDivision.ListIndex) & "') " & _
                        " AND (TournamentKey = " & TournamentKey & ")"
                End If
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                While Not rs.EOF
                    ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_Team_Rep " & _
                                      " (LogInName, TeamKey, TeamName) " & _
                                      " VALUES ('" & gbl_UserName & "', " & rs!PK & ", '" & FORMATSQL(rs!TeamName) & "')"
                    
                    dTotalTeam = 0
                    t = "SELECT  TOP " & TeamPlayer2Cnt & " tbl_Scoring_Team_Detail.PlayerKey, " & _
                        " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                        " tbl_Scoring_PlayerName.MiddleName, " & _
                        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS NetPoints " & _
                        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                        " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                        " Order By ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) DESC"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        dTotalTeam = dTotalTeam + CDbl(rt!NetPoints)
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_Team_Rep " & _
                                      " SET Score = " & CDbl(dTotalTeam) & " " & _
                                      " WHERE (LogInName = '" & gbl_UserName & "') " & _
                                      " AND (TeamKey = " & rs!PK & ")"
                    
                    dTotalTeam = 0: iTeamCounter = 0
                    t = "SELECT tbl_Scoring_Team_Detail.PlayerKey, " & _
                        " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                        " tbl_Scoring_PlayerName.MiddleName, " & _
                        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS NetPoints " & _
                        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                        " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                        " Order By ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) DESC"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        iTeamCounter = iTeamCounter + 1
                        If CDbl(iTeamCounter) > CDbl(TeamPlayer2Cnt) Then
                            dTotalTeam = dTotalTeam + CDbl(rt!NetPoints)
                        End If
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_Team_Rep " & _
                                      " SET CountBck = " & CDbl(dTotalTeam) & " " & _
                                      " WHERE (LogInName = '" & gbl_UserName & "') " & _
                                      " AND (TeamKey = " & rs!PK & ")"
                    
                    rs.MoveNext
                Wend
                rs.Close
                
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
                    .Name = "Top " & CStr(Trim(txtTop.Text))
                End With
                With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = TournamentName
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 10
                    .Range(strRange).Font.Bold = True

                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Range : " & TournamentRange
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    If cmbDivision.ListIndex > 0 Then
                        RowCnt = RowCnt + 1
                        ColCnt = 0
                        ColCnt = ColCnt + 1
                        HeaderRow = HeaderRow + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = "CLASS " & cmbDivision.List(cmbDivision.ListIndex)
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                    End If
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Team (Gross Points)"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False

                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = ""
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True

                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "#"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCnt).ColumnWidth = 3
                    .Range(strRange).HorizontalAlignment = 4
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                        
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Name"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Columns(ColCnt).ColumnWidth = 25
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    Arr = Split(TournamentRange, " - ", -1, 1)
                    iDay = 0
                    For i = 0 To DateDiff("d", Arr(0), Arr(1), vbMonday)
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = "Day " & i + 1
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).HorizontalAlignment = 4
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        iDay = iDay + 1
                    Next i

                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Total"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Team Total"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCnt).ColumnWidth = 10
                    .Range(strRange).HorizontalAlignment = 4
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1

                    j = 0
                    s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " TeamKey as PK, TeamName, Score  " & _
                        " FROM tbl_Scoring_ScoreCard_Team_Rep " & _
                        " WHERE (LogInName = '" & gbl_UserName & "') " & _
                        " ORDER BY Score DESC, CountBck DESC"
                        
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open s, ConnOmega
                    While Not rs.EOF
                        j = j + 1
                        RowCnt = RowCnt + 1
                        ColCnt = 0
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = j
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        RowCntTmp = RowCnt - 1
                        
                        ColCntTmp = ColCnt
                        RowCntTmp = RowCntTmp + 1
                        ColCntTmp = ColCntTmp + 1
                        strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                        .Range(strRange).Value = rs!TeamName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        For i = 1 To iDay
                            ColCntTmp = ColCntTmp + 1
                            strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                            .Range(strRange).Value = ""
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                        Next i
                        
                        t = "SELECT  tbl_Scoring_Team_Detail.PlayerKey, " & _
                            " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                            " tbl_Scoring_PlayerName.MiddleName, " & _
                            " (SELECT SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints " & _
                            " From tbl_Scoring_ScoreCard " & _
                            " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)) AS NetPoints " & _
                            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                            " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                            " Order By (SELECT SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints " & _
                            " From tbl_Scoring_ScoreCard " & _
                            " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)) DESC"
                        If rt.State = adStateOpen Then rt.Close
                        rt.Open t, ConnOmega
                        While Not rt.EOF
                            
                            ColCntTmp = ColCnt
                            RowCntTmp = RowCntTmp + 1
                            ColCntTmp = ColCntTmp + 1
                            strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                            .Range(strRange).Value = rt!LastName & ",  " & rt!FirstName & "  " & rt!MiddleName
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                            sTotal = "="
                            For i = 1 To iDay
                                dDate = DateAdd("d", CDbl(i - 1), Arr(0))
                                ColCntTmp = ColCntTmp + 1
                                strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                                sTotal = sTotal & strRange & "+"
                                u = "SELECT GrossPoints as NetPoints " & _
                                    " From tbl_Scoring_ScoreCard " & _
                                    " WHERE (PlayerKey = " & rt!PlayerKey & ") " & _
                                    " AND (DDate = '" & FormatDateTime(dDate, vbShortDate) & "')"
                                If ru.State = adStateOpen Then ru.Close
                                ru.Open u, ConnOmega
                                If ru.RecordCount > 0 Then
                                    .Range(strRange).Value = ru!NetPoints
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                Else
                                    .Range(strRange).Value = ""
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                End If
                                .Range(strRange).Select
                                xlsApp.Selection.Borders.LineStyle = 1
                                ru.Close
                            Next i
                            
                            ColCntTmp = ColCntTmp + 1
                            strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                            .Range(strRange).Value = Mid(sTotal, 1, Len(sTotal) - 1)
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                            rt.MoveNext
                        Wend
                        rt.Close
                        
                        ColCnt = ColCntTmp
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        strRangeFrom = EXCEL_RANGE(ColCnt, RowCnt)
                        strRangeTo = EXCEL_RANGE(ColCnt, RowCntTmp)
                        .Range(strRange).Value = rs!Score 'rs!NetPoints
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRangeFrom, strRangeTo).Select
                        xlsApp.Selection.Merge
                        .Range(strRange).VerticalAlignment = 2
                        .Range(strRange).HorizontalAlignment = 3
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        strRange = EXCEL_RANGE(1, RowCnt)
                        strRangeFrom = EXCEL_RANGE(1, RowCnt)
                        strRangeTo = EXCEL_RANGE(1, RowCntTmp)
                        .Range(strRangeFrom, strRangeTo).Select
                        xlsApp.Selection.Merge
                        .Range(strRange).VerticalAlignment = 2
                        .Range(strRange).HorizontalAlignment = 3
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        RowCnt = RowCntTmp
                        
                        UpdateProgress picProgressBar, j / rs.RecordCount
                        
                        rs.MoveNext
                    Wend
                    rs.Close
                    
                    .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
                    
                End With

SAVING4:
                On Error GoTo err_saving4:
                If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
                xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

                xlsApp.Visible = True
                
                picProgress.Visible = False
                picPrint.Enabled = True
                
            Case 2      'Gross Score
                
                picPrint.Enabled = False
                picProgress.ZOrder 0
                picProgressBar.BackColor = &HFFFFFF
                picProgress.Visible = True
                DoEvents
                
                ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_Team_Rep " & _
                                  " WHERE (LogInName = '" & gbl_UserName & "')"
                
                If cmbDivision.ListIndex = 0 Then
                    s = "SELECT PK, TeamName " & _
                        " From tbl_Scoring_Team " & _
                        " WHERE (TournamentKey = " & TournamentKey & ")"
                Else
                    s = "SELECT PK, TeamName " & _
                        " From tbl_Scoring_Team " & _
                        " WHERE ((SELECT Class " & _
                        " FROM tbl_Scoring_TournamentInfo_Class AS tbl_Scoring_TournamentInfo_Class_1 " & _
                        " WHERE (HFrom <= tbl_Scoring_Team.TeamHDCP) AND (HTo >= tbl_Scoring_Team.TeamHDCP) " & _
                        " AND (TournamentKey = " & TournamentKey & ")) = '" & cmbDivision.List(cmbDivision.ListIndex) & "') " & _
                        " AND (TournamentKey = " & TournamentKey & ")"
                End If
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                While Not rs.EOF
                    ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_Team_Rep " & _
                                      " (LogInName, TeamKey, TeamName) " & _
                                      " VALUES ('" & gbl_UserName & "', " & rs!PK & ", '" & FORMATSQL(rs!TeamName) & "')"
                    
                    dTotalTeam = 0
                    t = "SELECT  TOP " & TeamPlayer2Cnt & " tbl_Scoring_Team_Detail.PlayerKey, " & _
                        " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                        " tbl_Scoring_PlayerName.MiddleName, " & _
                        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.Score) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS NetPoints " & _
                        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                        " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                        " Order By ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.Score) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0)"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        dTotalTeam = dTotalTeam + CDbl(rt!NetPoints)
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_Team_Rep " & _
                                      " SET Score = " & CDbl(dTotalTeam) & " " & _
                                      " WHERE (LogInName = '" & gbl_UserName & "') " & _
                                      " AND (TeamKey = " & rs!PK & ")"
                    
                    dTotalTeam = 0: iTeamCounter = 0
                    t = "SELECT tbl_Scoring_Team_Detail.PlayerKey, " & _
                        " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                        " tbl_Scoring_PlayerName.MiddleName, " & _
                        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.Score) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS NetPoints " & _
                        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                        " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                        " Order By ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.Score) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard " & _
                        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0)"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        iTeamCounter = iTeamCounter + 1
                        If CDbl(iTeamCounter) > CDbl(TeamPlayer2Cnt) Then
                            dTotalTeam = dTotalTeam + CDbl(rt!NetPoints)
                        End If
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_Team_Rep " & _
                                      " SET CountBck = " & CDbl(dTotalTeam) & " " & _
                                      " WHERE (LogInName = '" & gbl_UserName & "') " & _
                                      " AND (TeamKey = " & rs!PK & ")"
                    
                    rs.MoveNext
                Wend
                rs.Close
                
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
                    .Name = "Top " & CStr(Trim(txtTop.Text))
                End With
                With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = TournamentName
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 10
                    .Range(strRange).Font.Bold = True

                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Range : " & TournamentRange
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    If cmbDivision.ListIndex > 0 Then
                        RowCnt = RowCnt + 1
                        ColCnt = 0
                        ColCnt = ColCnt + 1
                        HeaderRow = HeaderRow + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = "CLASS " & cmbDivision.List(cmbDivision.ListIndex)
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                    End If
                    
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Team (Gross Points)"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False

                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = ""
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True

                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    HeaderRow = HeaderRow + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "#"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCnt).ColumnWidth = 3
                    .Range(strRange).HorizontalAlignment = 4
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                        
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Name"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Columns(ColCnt).ColumnWidth = 25
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    Arr = Split(TournamentRange, " - ", -1, 1)
                    iDay = 0
                    For i = 0 To DateDiff("d", Arr(0), Arr(1), vbMonday)
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = "Day " & i + 1
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).HorizontalAlignment = 4
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        iDay = iDay + 1
                    Next i

                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Total"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Team Total"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCnt).ColumnWidth = 10
                    .Range(strRange).HorizontalAlignment = 4
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1

                    j = 0
                    s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " TeamKey as PK, TeamName, Score  " & _
                        " FROM tbl_Scoring_ScoreCard_Team_Rep " & _
                        " WHERE (LogInName = '" & gbl_UserName & "') " & _
                        " ORDER BY Score , CountBck "
                        
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open s, ConnOmega
                    While Not rs.EOF
                        j = j + 1
                        RowCnt = RowCnt + 1
                        ColCnt = 0
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = j
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        RowCntTmp = RowCnt - 1
                        
                        ColCntTmp = ColCnt
                        RowCntTmp = RowCntTmp + 1
                        ColCntTmp = ColCntTmp + 1
                        strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                        .Range(strRange).Value = rs!TeamName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        For i = 1 To iDay
                            ColCntTmp = ColCntTmp + 1
                            strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                            .Range(strRange).Value = ""
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                        Next i
                        
                        t = "SELECT  tbl_Scoring_Team_Detail.PlayerKey, " & _
                            " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                            " tbl_Scoring_PlayerName.MiddleName, " & _
                            " (SELECT SUM(tbl_Scoring_ScoreCard.Score) AS NetPoints " & _
                            " From tbl_Scoring_ScoreCard " & _
                            " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)) AS NetPoints " & _
                            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                            " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                            " Order By (SELECT SUM(tbl_Scoring_ScoreCard.Score) AS NetPoints " & _
                            " From tbl_Scoring_ScoreCard " & _
                            " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)) "
                        If rt.State = adStateOpen Then rt.Close
                        rt.Open t, ConnOmega
                        While Not rt.EOF
                            
                            ColCntTmp = ColCnt
                            RowCntTmp = RowCntTmp + 1
                            ColCntTmp = ColCntTmp + 1
                            strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                            .Range(strRange).Value = rt!LastName & ",  " & rt!FirstName & "  " & rt!MiddleName
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                            sTotal = "="
                            For i = 1 To iDay
                                dDate = DateAdd("d", CDbl(i - 1), Arr(0))
                                ColCntTmp = ColCntTmp + 1
                                strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                                sTotal = sTotal & strRange & "+"
                                u = "SELECT GrossPoints as NetPoints " & _
                                    " From tbl_Scoring_ScoreCard " & _
                                    " WHERE (PlayerKey = " & rt!PlayerKey & ") " & _
                                    " AND (DDate = '" & FormatDateTime(dDate, vbShortDate) & "')"
                                If ru.State = adStateOpen Then ru.Close
                                ru.Open u, ConnOmega
                                If ru.RecordCount > 0 Then
                                    .Range(strRange).Value = ru!NetPoints
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                Else
                                    .Range(strRange).Value = ""
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                End If
                                .Range(strRange).Select
                                xlsApp.Selection.Borders.LineStyle = 1
                                ru.Close
                            Next i
                            
                            ColCntTmp = ColCntTmp + 1
                            strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                            .Range(strRange).Value = Mid(sTotal, 1, Len(sTotal) - 1)
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                            rt.MoveNext
                        Wend
                        rt.Close
                        
                        ColCnt = ColCntTmp
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        strRangeFrom = EXCEL_RANGE(ColCnt, RowCnt)
                        strRangeTo = EXCEL_RANGE(ColCnt, RowCntTmp)
                        .Range(strRange).Value = rs!Score 'rs!NetPoints
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRangeFrom, strRangeTo).Select
                        xlsApp.Selection.Merge
                        .Range(strRange).VerticalAlignment = 2
                        .Range(strRange).HorizontalAlignment = 3
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        strRange = EXCEL_RANGE(1, RowCnt)
                        strRangeFrom = EXCEL_RANGE(1, RowCnt)
                        strRangeTo = EXCEL_RANGE(1, RowCntTmp)
                        .Range(strRangeFrom, strRangeTo).Select
                        xlsApp.Selection.Merge
                        .Range(strRange).VerticalAlignment = 2
                        .Range(strRange).HorizontalAlignment = 3
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        
                        RowCnt = RowCntTmp
                        
                        UpdateProgress picProgressBar, j / rs.RecordCount
                        
                        rs.MoveNext
                    Wend
                    rs.Close
                    
                    .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
                    
                End With

SAVING5:
                On Error GoTo err_saving5:
                If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
                xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

                xlsApp.Visible = True
                
                picProgress.Visible = False
                picPrint.Enabled = True
                
        End Select
        
    Case 2  'Result
        
        Exit Sub
        
End Select
Exit Sub
err_saving1:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING1:

Exit Sub
err_saving2:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING2:

Exit Sub
err_saving3:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING3:

Exit Sub
err_saving4:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING4:

Exit Sub
err_saving5:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING5:

Exit Sub
err_saving6:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING6:
End Sub

Private Sub TimerReportStableford_Timer()
TimerReportStableford.Enabled = False
ArrDate = Split(TournamentRange, " - ", -1, 1)
iDateDiff = DateDiff("d", ArrDate(0), ArrDate(1))
Select Case cmbReportType.ListIndex
    Case 0  'INDIVIDUAL
        Select Case cmbGroup.ListIndex
            Case 0  'NetPoints
                
                picPrint.Visible = False
                picProgress.Visible = True
                picProgressBar.BackColor = &HFFFFFF
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
                    .Name = "Top " & CStr(Trim(txtTop.Text))
                End With
                With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = TournamentName
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 10
                    .Range(strRange).Font.Bold = True
                    
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Range : " & TournamentRange
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    If cmbDivision.ListIndex = 0 Then
                        .Range(strRange).Value = "Individual (Net Points) [" & IIf(cmbGender.ListIndex = 1, "MALE", IIf(cmbGender.ListIndex = 2, "FEMALE", "")) & "]"
                    Else
                        .Range(strRange).Value = "Individual [Class " & cmbDivision.List(cmbDivision.ListIndex) & "] (Net Points) [" & IIf(cmbGender.ListIndex = 1, "MALE", IIf(cmbGender.ListIndex = 2, "FEMALE", "")) & "]"
                    End If
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = ""
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "#"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCnt).ColumnWidth = 3
                    .Range(strRange).HorizontalAlignment = 4
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Name"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Handicap"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    
                    Arr = Split(TournamentRange, " - ", -1, 1)
                    iDay = 0
                    For i = 0 To DateDiff("d", Arr(0), Arr(1), vbMonday)
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = "Day " & i + 1
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).HorizontalAlignment = 4
                        iDay = iDay + 1
                    Next i
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Total"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    j = 0
                    If cmbDivision.ListIndex = 0 Then
                        'All Class
                        If cmbGender.ListIndex = 0 Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Net) AS Holes_B9, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ")  " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.NetPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Net) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                            
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Net) AS Holes_B9, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex & ") " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.NetPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Net) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        End If
                    Else
                        'Class A to
                        If cmbGender.ListIndex = 0 Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Net) AS Holes_B9, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.NetPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Net) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.NetPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Net) AS Holes_B9, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex & ") AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.NetPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Net) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        End If
                    End If
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open s, ConnOmega
                    While Not rs.EOF
                        DoEvents
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
                        .Range(strRange).Value = rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        .Columns(ColCnt).ColumnWidth = 25
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = rs!Handicap
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        sTotal = "="
                        For i = 1 To iDay
                            dDate = DateAdd("d", CDbl(i - 1), Arr(0))
                            ColCnt = ColCnt + 1
                            strRange = EXCEL_RANGE(ColCnt, RowCnt)
                            sTotal = sTotal & strRange & "+"
                            t = "SELECT NetPoints " & _
                                " From tbl_Scoring_ScoreCard " & _
                                " WHERE (PlayerKey = " & rs!PlayerKey & ") " & _
                                " AND (DDate = '" & FormatDateTime(dDate, vbShortDate) & "')"
                            If rt.State = adStateOpen Then rt.Close
                            rt.Open t, ConnOmega
                            If rt.RecordCount > 0 Then
                                .Range(strRange).Value = rt!NetPoints
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                            Else
                                .Range(strRange).Value = ""
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                            End If
                            rt.Close
                        Next i
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = Mid(sTotal, 1, Len(sTotal) - 1)
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        UpdateProgress picProgressBar, iProgressValue / rs.RecordCount
                        
                        rs.MoveNext
                    Wend
                    rs.Close
                    '(tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex & ")
                    If cmbDivision.ListIndex = 0 Then
                        If cmbGender.ListIndex = 0 Then
                            s = "SELECT LastName, FirstName, MiddleName, HandiCap, " & _
                                " ISNULL((SELECT SUM(NetPoints) AS NetPoints " & _
                                " From dbo.tbl_Scoring_ScoreCard " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) AS NPoints " & _
                                " From dbo.tbl_Scoring_PlayerName " & _
                                " WHERE (ISNULL((SELECT  SUM(NetPoints) AS NetPoints " & _
                                " FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) = 0) " & _
                                " AND (TournamentKey = " & TournamentKey & ") " & _
                                " ORDER BY HandiCap, LastName, FirstName, MiddleName"
                        Else
                            s = "SELECT LastName, FirstName, MiddleName, HandiCap, " & _
                                " ISNULL((SELECT SUM(NetPoints) AS NetPoints " & _
                                " From dbo.tbl_Scoring_ScoreCard " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) AS NPoints " & _
                                " From dbo.tbl_Scoring_PlayerName " & _
                                " WHERE (ISNULL((SELECT  SUM(NetPoints) AS NetPoints " & _
                                " FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) = 0) " & _
                                " AND (TournamentKey = " & TournamentKey & ") " & _
                                " AND (Gender = " & cmbGender.ListIndex & ") " & _
                                " ORDER BY HandiCap, LastName, FirstName, MiddleName"
                        End If
                    Else
                        If cmbGender.ListIndex = 0 Then
                            s = "SELECT LastName, FirstName, MiddleName, HandiCap, " & _
                                " ISNULL((SELECT SUM(NetPoints) AS NetPoints " & _
                                " From dbo.tbl_Scoring_ScoreCard " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) AS NPoints " & _
                                " From dbo.tbl_Scoring_PlayerName " & _
                                " WHERE ((SELECT Class FROM dbo.tbl_Scoring_TournamentInfo_Class AS Class " & _
                                " WHERE (TournamentKey = dbo.tbl_Scoring_PlayerName.TournamentKey) AND (HFrom <= dbo.tbl_Scoring_PlayerName.HandiCap) " & _
                                " AND (HTo >= dbo.tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "') " & _
                                " AND (TournamentKey = " & TournamentKey & ") " & _
                                " AND (ISNULL((SELECT SUM(NetPoints) AS NetPoints " & _
                                " FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) = 0) " & _
                                " ORDER BY HandiCap, LastName, FirstName, MiddleName"
                        Else
                            s = "SELECT LastName, FirstName, MiddleName, HandiCap, " & _
                                " ISNULL((SELECT SUM(NetPoints) AS NetPoints " & _
                                " From dbo.tbl_Scoring_ScoreCard " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) AS NPoints " & _
                                " From dbo.tbl_Scoring_PlayerName " & _
                                " WHERE ((SELECT Class FROM dbo.tbl_Scoring_TournamentInfo_Class AS Class " & _
                                " WHERE (TournamentKey = dbo.tbl_Scoring_PlayerName.TournamentKey) AND (HFrom <= dbo.tbl_Scoring_PlayerName.HandiCap) " & _
                                " AND (HTo >= dbo.tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "') " & _
                                " AND (TournamentKey = " & TournamentKey & ") " & _
                                " AND (Gender = " & cmbGender.ListIndex & ") " & _
                                " AND (ISNULL((SELECT SUM(NetPoints) AS NetPoints " & _
                                " FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) = 0) " & _
                                " ORDER BY HandiCap, LastName, FirstName, MiddleName"
                        End If
                    End If
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open s, ConnOmega
                    While Not rs.EOF
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
                        .Range(strRange).Value = rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        .Columns(ColCnt).ColumnWidth = 25
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = rs!Handicap
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        rs.MoveNext
                    Wend
                    rs.Close
                    
                    .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
                    
                End With
                
SAVING1:
                On Error GoTo err_saving1:
                If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
                xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName
                
                xlsApp.Visible = True
                
            Case 1  'GrossPoints
                
                picPrint.Visible = False
                picProgress.Visible = True
                picProgressBar.BackColor = &HFFFFFF
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
                    .Name = "Top " & CStr(Trim(txtTop.Text))
                End With
                With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = TournamentName
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 10
                    .Range(strRange).Font.Bold = True
                    
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Range : " & TournamentRange
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    If cmbDivision.ListIndex = 0 Then
                        .Range(strRange).Value = "Individual (Gross Points) [" & IIf(cmbGender.ListIndex = 1, "MALE", IIf(cmbGender.ListIndex = 2, "FEMALE", "")) & "]"
                    Else
                        .Range(strRange).Value = "Individual [Class " & cmbDivision.List(cmbDivision.ListIndex) & "] (Gross Points) [" & IIf(cmbGender.ListIndex = 1, "MALE", IIf(cmbGender.ListIndex = 2, "FEMALE", "")) & "]"
                    End If
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = ""
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "#"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Columns(ColCnt).ColumnWidth = 3
                    .Range(strRange).HorizontalAlignment = 4
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Name"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Handicap"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    
                    Arr = Split(TournamentRange, " - ", -1, 1)
                    iDay = 0
                    For i = 0 To DateDiff("d", Arr(0), Arr(1), vbMonday)
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = "Day " & i + 1
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = True
                        .Range(strRange).HorizontalAlignment = 4
                        iDay = iDay + 1
                    Next i
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Total"
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    
                    j = 0
                    If cmbDivision.ListIndex = 0 Then
                        If cmbGender.ListIndex = 0 Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Gross) AS Holes_B9, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.GrossPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Gross) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Gross) AS Holes_B9, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex & ") " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.GrossPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Gross) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        End If
                    Else
                        If cmbGender.ListIndex = 0 Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Gross) AS Holes_B9, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.GrossPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Gross) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard.TournamentKey, tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard.GrossPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard.Back9Gross) AS Holes_B9, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex & ")  AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard.PlayerKey, tbl_Scoring_ScoreCard.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard.GrossPoints) DESC, SUM(tbl_Scoring_ScoreCard.Back9Gross) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        End If
                    End If
                    If ra.State = adStateOpen Then ra.Close
                    ra.Open s, ConnOmega
                    While Not ra.EOF
                        DoEvents
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
                        .Range(strRange).Value = ra!LastName & ",  " & ra!FirstName & "  " & ra!MiddleName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        .Columns(ColCnt).ColumnWidth = 25
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = ra!Handicap
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        sTotal = "="
                        For i = 1 To iDay
                            dDate = DateAdd("d", CDbl(i - 1), Arr(0))
                            ColCnt = ColCnt + 1
                            strRange = EXCEL_RANGE(ColCnt, RowCnt)
                            sTotal = sTotal & strRange & "+"
                            t = "SELECT GrossPoints as NetPoints " & _
                                " From tbl_Scoring_ScoreCard " & _
                                " WHERE (PlayerKey = " & ra!PlayerKey & ") " & _
                                " AND (DDate = '" & FormatDateTime(dDate, vbShortDate) & "')"
                            If rt.State = adStateOpen Then rt.Close
                            rt.Open t, ConnOmega
                            If rt.RecordCount > 0 Then
                                .Range(strRange).Value = rt!NetPoints
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                            Else
                                .Range(strRange).Value = ""
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                            End If
                            rt.Close
                        Next i
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = Mid(sTotal, 1, Len(sTotal) - 1)
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        
                        UpdateProgress picProgressBar, iProgressValue / ra.RecordCount
                        
                        ra.MoveNext
                    Wend
                    ra.Close
                    
                    's = "SELECT LastName, FirstName, MiddleName, HandiCap, " & _
                        " ISNULL((SELECT SUM(GrossPoints) AS NetPoints " & _
                        " From dbo.tbl_Scoring_ScoreCard " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) AS NPoints " & _
                        " From dbo.tbl_Scoring_PlayerName " & _
                        " WHERE (ISNULL((SELECT  SUM(GrossPoints) AS NetPoints " & _
                        " FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) = 0) " & _
                        " AND (TournamentKey = " & TournamentKey & ") " & _
                        " ORDER BY HandiCap, LastName, FirstName, MiddleName"
                    
                    If cmbDivision.ListIndex = 0 Then
                        If cmbGender.ListIndex = 0 Then
                            s = "SELECT LastName, FirstName, MiddleName, HandiCap, " & _
                                " ISNULL((SELECT SUM(GrossPoints) AS NetPoints " & _
                                " From dbo.tbl_Scoring_ScoreCard " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) AS NPoints " & _
                                " From dbo.tbl_Scoring_PlayerName " & _
                                " WHERE (ISNULL((SELECT  SUM(GrossPoints) AS NetPoints " & _
                                " FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) = 0) " & _
                                " AND (TournamentKey = " & TournamentKey & ") " & _
                                " ORDER BY HandiCap, LastName, FirstName, MiddleName"
                        Else
                            s = "SELECT LastName, FirstName, MiddleName, HandiCap, " & _
                                " ISNULL((SELECT SUM(GrossPoints) AS NetPoints " & _
                                " From dbo.tbl_Scoring_ScoreCard " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) AS NPoints " & _
                                " From dbo.tbl_Scoring_PlayerName " & _
                                " WHERE (ISNULL((SELECT  SUM(GrossPoints) AS NetPoints " & _
                                " FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) = 0) " & _
                                " AND (TournamentKey = " & TournamentKey & ") " & _
                                " AND (Gender = " & cmbGender.ListIndex & ") " & _
                                " ORDER BY HandiCap, LastName, FirstName, MiddleName"
                        End If
                    Else
                        If cmbGender.ListIndex = 0 Then
                            s = "SELECT LastName, FirstName, MiddleName, HandiCap, " & _
                                " ISNULL((SELECT SUM(GrossPoints) AS NetPoints " & _
                                " From dbo.tbl_Scoring_ScoreCard " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) AS NPoints " & _
                                " From dbo.tbl_Scoring_PlayerName " & _
                                " WHERE ((SELECT Class FROM dbo.tbl_Scoring_TournamentInfo_Class AS Class " & _
                                " WHERE (TournamentKey = dbo.tbl_Scoring_PlayerName.TournamentKey) AND (HFrom <= dbo.tbl_Scoring_PlayerName.HandiCap) " & _
                                " AND (HTo >= dbo.tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "') " & _
                                " AND (TournamentKey = " & TournamentKey & ") " & _
                                " AND (ISNULL((SELECT SUM(GrossPoints) AS NetPoints " & _
                                " FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) = 0) " & _
                                " ORDER BY HandiCap, LastName, FirstName, MiddleName"
                        Else
                            s = "SELECT LastName, FirstName, MiddleName, HandiCap, " & _
                                " ISNULL((SELECT SUM(GrossPoints) AS NetPoints " & _
                                " From dbo.tbl_Scoring_ScoreCard " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) AS NPoints " & _
                                " From dbo.tbl_Scoring_PlayerName " & _
                                " WHERE ((SELECT Class FROM dbo.tbl_Scoring_TournamentInfo_Class AS Class " & _
                                " WHERE (TournamentKey = dbo.tbl_Scoring_PlayerName.TournamentKey) AND (HFrom <= dbo.tbl_Scoring_PlayerName.HandiCap) " & _
                                " AND (HTo >= dbo.tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "') " & _
                                " AND (TournamentKey = " & TournamentKey & ") " & _
                                " AND (Gender = " & cmbGender.ListIndex & ") " & _
                                " AND (ISNULL((SELECT SUM(GrossPoints) AS NetPoints " & _
                                " FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                                " WHERE (PlayerKey = dbo.tbl_Scoring_PlayerName.PK)), 0) = 0) " & _
                                " ORDER BY HandiCap, LastName, FirstName, MiddleName"
                        End If
                    End If
                    If rs.State = adStateOpen Then rs.Close
                    rs.Open s, ConnOmega
                    While Not rs.EOF
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
                        .Range(strRange).Value = rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        .Columns(ColCnt).ColumnWidth = 25
                        
                        ColCnt = ColCnt + 1
                        strRange = EXCEL_RANGE(ColCnt, RowCnt)
                        .Range(strRange).Value = rs!Handicap
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        rs.MoveNext
                    Wend
                    rs.Close
                    
                    .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
                    
                End With
SAVING2:
                On Error GoTo err_saving2:
                If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
                xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName
                
                xlsApp.Visible = True
                
        End Select
        
    Case 1  'TEAM
        
        'Create_Table
        TableName = "tmp_" & gbl_UserName & "_Scoring_ModStableFord"
        DetailTableName = TableName & "_Detail"
        CREATE_MODIFIED_STABLE_FORD TableName
        
        Select Case cmbGroup.ListIndex
        
            Case 0  'NetPoints
                
                j = 0
                picPrint.Visible = False
                picProgress.Visible = True
                picProgressBar.BackColor = &HFFFFFF
                DoEvents
                                
                s = "SELECT PK, TeamName, TeamHDCP, TeamIndex " & _
                    " From tbl_Scoring_Team " & _
                    " WHERE (TournamentKey = " & TournamentKey & ")"
                If rs.State = adStateOpen Then ra.Close
                ra.Open s, ConnOmega
                While Not ra.EOF
                    DoEvents
                    j = j + 1
                    
                    iPlayCnt = 0: dTeamTotal = 0: dB9Total = 0: dF9Total = 0: iDLine = 0: dLastMan = 0
                    
                    t = "SELECT ISNULL ((SELECT SUM(NetPoints) AS Points From dbo.tbl_Scoring_ScoreCard " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) AS Points, " & _
                        " ISNULL((SELECT SUM(Back9Net) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) as B9, " & _
                        " ISNULL((SELECT SUM(Front9Net) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) as F9 " & _
                        " FROM dbo.tbl_Scoring_Team LEFT OUTER JOIN " & _
                        " dbo.tbl_Scoring_Team_Detail ON dbo.tbl_Scoring_Team.PK = dbo.tbl_Scoring_Team_Detail.TeamKey " & _
                        " Where (dbo.tbl_Scoring_Team.PK = " & ra!PK & ") " & _
                        " ORDER BY ISNULL((SELECT SUM(NetPoints) AS Points From dbo.tbl_Scoring_ScoreCard " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC, " & _
                        " ISNULL((SELECT SUM(Back9Net) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC, " & _
                        " ISNULL((SELECT SUM(Front9Net) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        iPlayCnt = iPlayCnt + 1
                        If CDbl(iPlayCnt) <= CDbl(TeamPlayer2Cnt) Then
                            dTeamTotal = dTeamTotal + CDbl(rt!Points)
                            dB9Total = dB9Total + CDbl(rt!B9)
                            dF9Total = dF9Total + CDbl(rt!F9)
                        End If
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ConnOmega.Execute "INSERT INTO " & TableName & " " & _
                                      " (TeamName, AveHandicap, TeamKey, TeamIndex) " & _
                                      " VALUES ('" & FORMATSQL(ra!TeamName) & "', " & _
                                      " " & ra!TeamHDCP & ", " & ra!PK & ", " & _
                                      " " & ra!TeamIndex & ")"
                    t = "SELECT PK " & _
                        " FROM " & TableName & " " & _
                        " WHERE (TeamKey = " & ra!PK & ")"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    If rt.RecordCount > 0 Then
                        
                        ConnOmega.Execute "UPDATE " & TableName & " " & _
                                          " SET TeamTotal = " & CDbl(dTeamTotal) & ", " & _
                                          " Back9 = " & CDbl(dB9Total) & ", " & _
                                          " Front9 = " & CDbl(dF9Total) & " " & _
                                          " WHERE (PK = " & rt!PK & ")"
                        
                        u = "SELECT dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName, dbo.tbl_Scoring_Team_Detail.PlayerKey, " & _
                            " ISNULL((SELECT SUM(NetPoints) AS Points From dbo.tbl_Scoring_ScoreCard WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) AS Points " & _
                            " FROM dbo.tbl_Scoring_PlayerName RIGHT OUTER JOIN " & _
                            " dbo.tbl_Scoring_Team_Detail ON dbo.tbl_Scoring_PlayerName.PK = dbo.tbl_Scoring_Team_Detail.PlayerKey RIGHT OUTER JOIN " & _
                            " dbo.tbl_Scoring_Team ON dbo.tbl_Scoring_Team_Detail.TeamKey = dbo.tbl_Scoring_Team.PK " & _
                            " Where (dbo.tbl_Scoring_Team.PK = " & ra!PK & ") " & _
                            " ORDER BY ISNULL((SELECT SUM(NetPoints) AS Points From dbo.tbl_Scoring_ScoreCard " & _
                            " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC, " & _
                            " ISNULL((SELECT SUM(Back9Net) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                            " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC, " & _
                            " ISNULL((SELECT SUM(Front9Net) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                            " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC"
                        If ru.State = adStateOpen Then ru.Close
                        ru.Open u, ConnOmega
                        While Not ru.EOF
                            iDLine = iDLine + 1
                            ConnOmega.Execute "INSERT INTO " & DetailTableName & " " & _
                                              " (MasterKey, Line, PlayerName) " & _
                                              " VALUES (" & rt!PK & ", " & iDLine & ", '" & FORMATSQL(ru!LastName & ",  " & ru!FirstName & "  " & ru!MiddleName) & "')"
                            
                            Arr = Split(TournamentRange, " - ", -1, 1)
                            iDay = DateDiff("d", Arr(0), Arr(1), vbMonday) + 1
                            
                            For i = 1 To iDay
                                dDate = DateAdd("d", CDbl(i - 1), Arr(0))
                                sFieldNum = "Day" & CStr(i)
                                v = "SELECT NetPoints as NetPoints " & _
                                    " From tbl_Scoring_ScoreCard " & _
                                    " WHERE (PlayerKey = " & ru!PlayerKey & ") " & _
                                    " AND (DDate = '" & FormatDateTime(dDate, vbShortDate) & "')"
                                If rv.State = adStateOpen Then rv.Close
                                rv.Open v, ConnOmega
                                If rv.RecordCount > 0 Then
                                    ConnOmega.Execute "UPDATE " & DetailTableName & " " & _
                                                      " SET " & sFieldNum & " = " & rv!NetPoints & " " & _
                                                      " WHERE (MasterKey = " & rt!PK & ") " & _
                                                      " AND (Line = " & iDLine & ")"
                                End If
                                rv.Close
                            Next i
                            
                            If CDbl(iDLine) > CDbl(TeamPlayer2Cnt) Then
                                dLastMan = dLastMan + CDbl(ru!Points)
                            End If
                            ru.MoveNext
                        Wend
                        ru.Close
                        
                        ConnOmega.Execute "UPDATE " & TableName & " " & _
                                          " SET LastPlayer = " & CDbl(dLastMan) & " " & _
                                          " WHERE (PK = " & rt!PK & ")"
                        
                    End If
                    rt.Close
                    
                    UpdateProgress picProgressBar, j / ra.RecordCount
                    ra.MoveNext
                Wend
                ra.Close
            
            Case 1 'Gross Points
                
                j = 0
                picPrint.Visible = False
                picProgress.Visible = True
                picProgressBar.BackColor = &HFFFFFF
                DoEvents
                              
                s = "SELECT PK, TeamName, TeamHDCP, TeamIndex " & _
                    " From tbl_Scoring_Team " & _
                    " WHERE (TournamentKey = " & TournamentKey & ")"
                If ra.State = adStateOpen Then ra.Close
                ra.Open s, ConnOmega
                While Not ra.EOF
                    DoEvents
                    j = j + 1
                    
                    iPlayCnt = 0: dTeamTotal = 0: dB9Total = 0: dF9Total = 0: iDLine = 0: dLastMan = 0
                    
                    t = "SELECT ISNULL ((SELECT SUM(GrossPoints) AS Points From dbo.tbl_Scoring_ScoreCard " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) AS Points, " & _
                        " ISNULL((SELECT SUM(Back9Gross) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) as B9, " & _
                        " ISNULL((SELECT SUM(Front9Gross) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) as F9 " & _
                        " FROM dbo.tbl_Scoring_Team LEFT OUTER JOIN " & _
                        " dbo.tbl_Scoring_Team_Detail ON dbo.tbl_Scoring_Team.PK = dbo.tbl_Scoring_Team_Detail.TeamKey " & _
                        " Where (dbo.tbl_Scoring_Team.PK = " & ra!PK & ") " & _
                        " ORDER BY ISNULL((SELECT SUM(GrossPoints) AS Points From dbo.tbl_Scoring_ScoreCard " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC, " & _
                        " ISNULL((SELECT SUM(Back9Gross) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC, " & _
                        " ISNULL((SELECT SUM(Front9Gross) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                        " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        iPlayCnt = iPlayCnt + 1
                        If CDbl(iPlayCnt) <= CDbl(TeamPlayer2Cnt) Then
                            dTeamTotal = dTeamTotal + CDbl(rt!Points)
                            dB9Total = dB9Total + CDbl(rt!B9)
                            dF9Total = dF9Total + CDbl(rt!F9)
                        End If
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ConnOmega.Execute "INSERT INTO " & TableName & " " & _
                                      " (TeamName, AveHandicap, TeamKey, TeamIndex) " & _
                                      " VALUES ('" & FORMATSQL(ra!TeamName) & "', " & _
                                    " " & ra!TeamHDCP & ", " & ra!PK & ", " & _
                                      " " & ra!TeamIndex & ")"
                    t = "SELECT PK " & _
                        " FROM " & TableName & " " & _
                        " WHERE (TeamKey = " & ra!PK & ")"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    If rt.RecordCount > 0 Then
                        
                        ConnOmega.Execute "UPDATE " & TableName & " " & _
                                          " SET TeamTotal = " & CDbl(dTeamTotal) & ", " & _
                                          " Back9 = " & CDbl(dB9Total) & ", " & _
                                          " Front9 = " & CDbl(dF9Total) & " " & _
                                          " WHERE (PK = " & rt!PK & ")"
                        
                        u = "SELECT dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName, dbo.tbl_Scoring_Team_Detail.PlayerKey, " & _
                            " ISNULL((SELECT SUM(GrossPoints) AS Points From dbo.tbl_Scoring_ScoreCard WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) AS Points " & _
                            " FROM dbo.tbl_Scoring_PlayerName RIGHT OUTER JOIN " & _
                            " dbo.tbl_Scoring_Team_Detail ON dbo.tbl_Scoring_PlayerName.PK = dbo.tbl_Scoring_Team_Detail.PlayerKey RIGHT OUTER JOIN " & _
                            " dbo.tbl_Scoring_Team ON dbo.tbl_Scoring_Team_Detail.TeamKey = dbo.tbl_Scoring_Team.PK " & _
                            " Where (dbo.tbl_Scoring_Team.PK = " & ra!PK & ") " & _
                            " ORDER BY ISNULL((SELECT SUM(GrossPoints) AS Points From dbo.tbl_Scoring_ScoreCard " & _
                            " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC, " & _
                            " ISNULL((SELECT SUM(Back9Gross) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                            " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC, " & _
                            " ISNULL((SELECT SUM(Front9Gross) AS Points FROM dbo.tbl_Scoring_ScoreCard AS tbl_Scoring_ScoreCard_1 " & _
                            " WHERE (PlayerKey = dbo.tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC"
                        If ru.State = adStateOpen Then ru.Close
                        ru.Open u, ConnOmega
                        While Not ru.EOF
                            iDLine = iDLine + 1
                            ConnOmega.Execute "INSERT INTO " & DetailTableName & " " & _
                                              " (MasterKey, Line, PlayerName) " & _
                                              " VALUES (" & rt!PK & ", " & iDLine & ", '" & FORMATSQL(ru!LastName & ",  " & ru!FirstName & "  " & ru!MiddleName) & "')"
                            
                            Arr = Split(TournamentRange, " - ", -1, 1)
                            iDay = DateDiff("d", Arr(0), Arr(1), vbMonday) + 1
                            
                            For i = 1 To iDay
                                dDate = DateAdd("d", CDbl(i - 1), Arr(0))
                                sFieldNum = "Day" & CStr(i)
                                v = "SELECT GrossPoints as NetPoints " & _
                                    " From tbl_Scoring_ScoreCard " & _
                                    " WHERE (PlayerKey = " & ru!PlayerKey & ") " & _
                                    " AND (DDate = '" & FormatDateTime(dDate, vbShortDate) & "')"
                                If rv.State = adStateOpen Then rv.Close
                                rv.Open v, ConnOmega
                                If rv.RecordCount > 0 Then
                                    ConnOmega.Execute "UPDATE " & DetailTableName & " " & _
                                                      " SET " & sFieldNum & " = " & rv!NetPoints & " " & _
                                                      " WHERE (MasterKey = " & rt!PK & ") " & _
                                                      " AND (Line = " & iDLine & ")"
                                End If
                                rv.Close
                            Next i
                            
                            If CDbl(iDLine) > CDbl(TeamPlayer2Cnt) Then
                                dLastMan = dLastMan + CDbl(ru!Points)
                            End If
                            ru.MoveNext
                        Wend
                        ru.Close
                        
                        ConnOmega.Execute "UPDATE " & TableName & " " & _
                                          " SET LastPlayer = " & CDbl(dLastMan) & " " & _
                                          " WHERE (PK = " & rt!PK & ")"
                        
                    End If
                    rt.Close
                    
                    UpdateProgress picProgressBar, j / ra.RecordCount
                    ra.MoveNext
                Wend
                ra.Close
                
            End Select
            
            picProgressBar.BackColor = &HFFFFFF
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
                .Name = "Top " & CStr(Trim(txtTop.Text))
            End With
            With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
                RowCnt = RowCnt + 1
                HeaderRow = HeaderRow + 1
                ColCnt = 0
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = TournamentName
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 10
                .Range(strRange).Font.Bold = True

                RowCnt = RowCnt + 1
                HeaderRow = HeaderRow + 1
                ColCnt = 0
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = "Range : " & TournamentRange
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False

                RowCnt = RowCnt + 1
                HeaderRow = HeaderRow + 1
                ColCnt = 0
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                If cmbGroup.ListIndex = 0 Then
                    .Range(strRange).Value = "Team (Net Points)"
                ElseIf cmbGroup.ListIndex = 1 Then
                    .Range(strRange).Value = "Team (Gross Points)"
                End If
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                
                If cmbDivision.ListIndex > 0 Then
                    RowCnt = RowCnt + 1
                    HeaderRow = HeaderRow + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Division : " & cmbDivision.List(cmbDivision.ListIndex)
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                End If
                
                RowCnt = RowCnt + 1
                HeaderRow = HeaderRow + 1
                ColCnt = 0
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = ""
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = True

                RowCnt = RowCnt + 1
                HeaderRow = HeaderRow + 1
                ColCnt = 0
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = "TeamName"
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = True
                .Columns(ColCnt).ColumnWidth = 15
                .Range(strRange).HorizontalAlignment = 3
                .Range(strRange).Select
                xlsApp.Selection.Borders.LineStyle = 1
                
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                If TeamAverage = 1 Then
                    .Range(strRange).Value = "Ave. Handicap"
                ElseIf TeamAverage = 2 Then
                    .Range(strRange).Value = "Ave. Index"
                End If
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Columns(ColCnt).ColumnWidth = 9
                .Range(strRange).Font.Bold = True
                .Range(strRange).HorizontalAlignment = 3
                .Range(strRange).Select
                xlsApp.Selection.Borders.LineStyle = 1
                    
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = "Name"
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Columns(ColCnt).ColumnWidth = 25
                .Range(strRange).Font.Bold = True
                .Range(strRange).Select
                xlsApp.Selection.Borders.LineStyle = 1
                
                Arr = Split(TournamentRange, " - ", -1, 1)
                iDay = 0
                For i = 0 To DateDiff("d", Arr(0), Arr(1), vbMonday)
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = "Day " & i + 1
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRange).HorizontalAlignment = 4
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    iDay = iDay + 1
                Next i

                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = "Total"
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = True
                .Range(strRange).HorizontalAlignment = 4
                .Range(strRange).Select
                xlsApp.Selection.Borders.LineStyle = 1
                
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = "Team Total"
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = True
                .Columns(ColCnt).ColumnWidth = 10
                .Range(strRange).HorizontalAlignment = 4
                .Range(strRange).Select
                xlsApp.Selection.Borders.LineStyle = 1

                j = 0
                If cmbDivision.ListIndex = 0 Or _
                cmbDivision.ListIndex = -1 Then
                
                    s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " PK, TeamName, AveHandicap, TeamIndex, TeamTotal " & _
                        " From " & TableName & " " & _
                        " ORDER BY TeamTotal DESC, LastPlayer DESC, Back9 DESC, Front9 DESC"
                Else
                    If TeamAverage = 1 Then
                        s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " PK, TeamName, AveHandicap, TeamIndex, TeamTotal, " & _
                            " ISNULL((SELECT Class From dbo.tbl_Scoring_TournamentInfo_Class " & _
                            " WHERE (TournamentKey = " & TournamentKey & ") " & _
                            " AND (HFrom <= dbo." & TableName & ".AveHandicap) " & _
                            " AND (HTo >= dbo." & TableName & ".AveHandicap)), '') AS Class " & _
                            " From " & TableName & " " & _
                            " WHERE (ISNULL((SELECT Class FROM dbo.tbl_Scoring_TournamentInfo_Class AS tbl_Scoring_TournamentInfo_Class_1 " & _
                            " WHERE (TournamentKey = " & TournamentKey & ") " & _
                            " AND (HFrom <= dbo." & TableName & ".AveHandicap) " & _
                            " AND (HTo >= dbo." & TableName & ".AveHandicap)), '') = '" & cmbDivision.List(cmbDivision.ListIndex) & "') " & _
                            " ORDER BY TeamTotal DESC, LastPlayer DESC, Back9 DESC, Front9 DESC"
                            
                    ElseIf TeamAverage = 2 Then
                        
                        s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " PK, TeamName, AveHandicap, TeamIndex, TeamTotal, " & _
                            " ISNULL((SELECT Class From dbo.tbl_Scoring_TournamentInfo_Index " & _
                            " WHERE (TournamentKey = " & TournamentKey & ") " & _
                            " AND (HFrom <= dbo." & TableName & ".TeamIndex) " & _
                            " AND (HTo >= dbo." & TableName & ".TeamIndex)), '') AS Class " & _
                            " From " & TableName & " " & _
                            " WHERE (ISNULL((SELECT Class FROM dbo.tbl_Scoring_TournamentInfo_Index AS tbl_Scoring_TournamentInfo_Index_1 " & _
                            " WHERE (TournamentKey = " & TournamentKey & ") " & _
                            " AND (HFrom <= dbo." & TableName & ".TeamIndex) " & _
                            " AND (HTo >= dbo." & TableName & ".TeamIndex)), '') = '" & cmbDivision.List(cmbDivision.ListIndex) & "') " & _
                            " ORDER BY TeamTotal DESC, LastPlayer DESC, Back9 DESC, Front9 DESC"
                    End If
                End If
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                While Not rs.EOF
                    DoEvents
                    iProgressValue = iProgressValue + 1
                    j = j + 1
                    RowCnt = RowCnt + 1
                    ColCnt = 0
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = rs!TeamName
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = rs!TeamIndex
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    
                    RowCntTmp = RowCnt - 1
                    
                    t = "SELECT * " & _
                        " FROM " & DetailTableName & " " & _
                        " WHERE (MasterKey = " & rs!PK & ") " & _
                        " ORDER BY Line"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    l = 0
                    dTeamPoints = 0
                    While Not rt.EOF
                        l = l + 1
                        ColCntTmp = ColCnt
                        RowCntTmp = RowCntTmp + 1
                        ColCntTmp = ColCntTmp + 1
                        strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                        .Range(strRange).Value = rt!PlayerName
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        iDay = rt.Fields.Count - 2
                        For i = 3 To iDay
                            ColCntTmp = ColCntTmp + 1
                            strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                            .Range(strRange).Value = rt.Fields(i).Value
                            .Range(strRange).Font.Name = "Tahoma"
                            .Range(strRange).Font.Size = 8
                            .Range(strRange).Font.Bold = False
                            
                            .Range(strRange).Select
                            xlsApp.Selection.Borders.LineStyle = 1
                        Next i
                        
                        ColCntTmp = ColCntTmp + 1
                        strRange = EXCEL_RANGE(ColCntTmp, RowCntTmp)
                        .Range(strRange).Value = rt!Total
                        .Range(strRange).Font.Name = "Tahoma"
                        .Range(strRange).Font.Size = 8
                        .Range(strRange).Font.Bold = False
                        .Range(strRange).Select
                        xlsApp.Selection.Borders.LineStyle = 1
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ColCnt = ColCntTmp
                    
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    strRangeFrom = EXCEL_RANGE(ColCnt, RowCnt)
                    strRangeTo = EXCEL_RANGE(ColCnt, RowCntTmp)
                    .Range(strRange).Value = rs!TeamTotal
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = True
                    .Range(strRangeFrom, strRangeTo).Select
                    xlsApp.Selection.Merge
                    .Range(strRange).VerticalAlignment = 2
                    .Range(strRange).HorizontalAlignment = 3
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    strRange = EXCEL_RANGE(1, RowCnt)
                    strRangeFrom = EXCEL_RANGE(1, RowCnt)
                    strRangeTo = EXCEL_RANGE(1, RowCntTmp)
                    .Range(strRangeFrom, strRangeTo).Select
                    xlsApp.Selection.Merge
                    .Range(strRange).VerticalAlignment = 2
                    .Range(strRange).HorizontalAlignment = 3
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    strRange = EXCEL_RANGE(2, RowCnt)
                    strRangeFrom = EXCEL_RANGE(2, RowCnt)
                    strRangeTo = EXCEL_RANGE(2, RowCntTmp)
                    .Range(strRangeFrom, strRangeTo).Select
                    xlsApp.Selection.Merge
                    .Range(strRange).VerticalAlignment = 2
                    .Range(strRange).HorizontalAlignment = 3
                    .Range(strRange).Select
                    xlsApp.Selection.Borders.LineStyle = 1
                    
                    RowCnt = RowCntTmp
                    
                    UpdateProgress_Caption "Exporting to Excel", picProgressBar, iProgressValue / rs.RecordCount
                    
                    rs.MoveNext
                Wend
                rs.Close
                
                .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
                
            End With
        
SAVING3:
            On Error GoTo err_saving3:
            If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
            xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

            xlsApp.Visible = True
                    
    Case 2  'Result
        
        Exit Sub
    
    Case 3  'Scores
        
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
            RowCnt = RowCnt + 1
            HeaderRow = HeaderRow + 1
            ColCnt = 0
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = TournamentName
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 10
            .Range(strRange).Font.Bold = True

            RowCnt = RowCnt + 1
            HeaderRow = HeaderRow + 1
            ColCnt = 0
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = "Range : " & TournamentRange
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = False
            
            s = "SELECT dbo.tbl_Scoring_System.ScoringSystem " & _
                " FROM dbo.tbl_Scoring_TournamentInfo LEFT OUTER JOIN " & _
                " dbo.tbl_Scoring_System ON dbo.tbl_Scoring_TournamentInfo.Scoring = dbo.tbl_Scoring_System.PK " & _
                " WHERE (dbo.tbl_Scoring_TournamentInfo.PK = " & TournamentKey & ")"
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
            ColCnt = 0
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = ""
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = False
            
            RowCnt = RowCnt + 1
            HeaderRow = HeaderRow + 1
            ColCnt = 0
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = "#"
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = "Name"
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = "Date"
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = ""
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = False
            
            For i = 1 To 9
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = "'" & Format(i, "0#")
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                .Range(strRange).HorizontalAlignment = 3
            Next i
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = "F-9"
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = False
            .Range(strRange).HorizontalAlignment = 3
            
            For i = 1 To 9
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = "'" & Format(CDbl(i) + 9, "0#")
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
                .Range(strRange).HorizontalAlignment = 3
            Next i
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = "B-9"
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = False
            .Range(strRange).HorizontalAlignment = 3
            
            strRange = EXCEL_RANGE(1, HeaderRow + 1)
            .Range(strRange).Select
            xlsApp.ActiveWindow.FreezePanes = True
            
            picProgressBar.Visible = True
            s = "SELECT dbo.tbl_Scoring_ScoreCard.PlayerKey, dbo.tbl_Scoring_PlayerName.LastName, " & _
                " dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName " & _
                " FROM dbo.tbl_Scoring_ScoreCard LEFT OUTER JOIN " & _
                " dbo.tbl_Scoring_PlayerName ON dbo.tbl_Scoring_ScoreCard.PlayerKey = dbo.tbl_Scoring_PlayerName.PK " & _
                " Where (dbo.tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") " & _
                " GROUP BY dbo.tbl_Scoring_ScoreCard.PlayerKey, dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName, dbo.tbl_Scoring_PlayerName.MiddleName " & _
                " ORDER BY dbo.tbl_Scoring_PlayerName.LastName, dbo.tbl_Scoring_PlayerName.FirstName"
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
                
                t = "SELECT PK, PlayerKey, DDate " & _
                    " From dbo.tbl_Scoring_ScoreCard " & _
                    " Where (PlayerKey = " & rs!PlayerKey & ") " & _
                    " ORDER BY DDate"
                If rt.State = adStateOpen Then rt.Close
                rt.Open t, ConnOmega
                If rt.RecordCount = 1 Then
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = Format(rt!dDate, "mm/dd/yyyy")
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                    For i = 1 To 3
                        sMinRange = "": sMaxRange = ""
                        Select Case i
                            Case 1
                                ColCnt = ColCnt + 1
                                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                .Range(strRange).Value = "Score"
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                                u = "SELECT Score " & _
                                    " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                    " Where (ScoreCardKey = " & rt!PK & ") " & _
                                    " ORDER BY Hole"
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
                            Case 2
                                RowCnt = RowCnt + 1
                                ColCnt = 3
                                ColCnt = ColCnt + 1
                                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                .Range(strRange).Value = "Gross"
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                                u = "SELECT Gross " & _
                                    " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                    " Where (ScoreCardKey = " & rt!PK & ") " & _
                                    " ORDER BY Hole"
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
                            Case 3
                                RowCnt = RowCnt + 1
                                ColCnt = 3
                                ColCnt = ColCnt + 1
                                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                .Range(strRange).Value = "Net"
                                .Range(strRange).Font.Name = "Tahoma"
                                .Range(strRange).Font.Size = 8
                                .Range(strRange).Font.Bold = False
                                u = "SELECT Net " & _
                                    " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                    " Where (ScoreCardKey = " & rt!PK & ") " & _
                                    " ORDER BY Hole"
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
                        End Select
                    Next i
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
                        For i = 1 To 3
                            sMinRange = "": sMaxRange = ""
                            Select Case i
                                Case 1
                                    ColCnt = ColCnt + 1
                                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                    .Range(strRange).Value = "Score"
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                    u = "SELECT Score " & _
                                        " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                        " Where (ScoreCardKey = " & rt!PK & ") " & _
                                        " ORDER BY Hole"
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
                                Case 2
                                    RowCnt = RowCnt + 1
                                    ColCnt = 3
                                    ColCnt = ColCnt + 1
                                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                    .Range(strRange).Value = "Gross"
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                    u = "SELECT Gross " & _
                                        " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                        " Where (ScoreCardKey = " & rt!PK & ") " & _
                                        " ORDER BY Hole"
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
                                Case 3
                                    RowCnt = RowCnt + 1
                                    ColCnt = 3
                                    ColCnt = ColCnt + 1
                                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                                    .Range(strRange).Value = "Net"
                                    .Range(strRange).Font.Name = "Tahoma"
                                    .Range(strRange).Font.Size = 8
                                    .Range(strRange).Font.Bold = False
                                    u = "SELECT Net " & _
                                        " From dbo.tbl_Scoring_ScoreCard_Detail " & _
                                        " Where (ScoreCardKey = " & rt!PK & ") " & _
                                        " ORDER BY Hole"
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
                            End Select
                        Next i
                        rt.MoveNext
                    Wend
                End If
                rt.Close
                
                
                UpdateProgress_Caption "Exporting to Excel", picProgressBar, iProgressValue / rs.RecordCount
                rs.MoveNext
            Wend
            rs.Close
            
            .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
            
        End With

SAVING4:
            On Error GoTo err_saving4:
            If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
            xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

            xlsApp.Visible = True
End Select

picProgress.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True

'cmdCancelPrint_Click

Exit Sub
err_saving1:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING1:

Exit Sub
err_saving2:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING2:

Exit Sub
err_saving3:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING3:

Exit Sub

err_saving4:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING4:

Exit Sub
'err_saving4:
'MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
'GoTo SAVING4:
End Sub

Private Sub TimerTeamAllStableFord_Timer()
TimerTeamAllStableFord.Enabled = False

With MainForm.CommonDialog1
    .CancelError = True
    On Error GoTo ErrorHandler
    .DialogTitle = "Save"
    .Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
    .ShowSave
    Filename = Trim(.Filename)
End With

On Error GoTo PG:
WorkbookName = Filename
picProgressBar.BackColor = &HFFFFFF
picProgress.ZOrder 0
picPrintStableford.Visible = False
picProgress.Visible = True

Screen.MousePointer = vbHourglass
TableName = "tmp_" & gbl_UserName & "_Report"
Columns = ""
Columns = Columns & "|Sorting:int:NOT NULL:DEFAULT(0)"
Columns = Columns & "|TeamKey:int:NOT NULL"
Columns = Columns & "|TeamID:varchar:(50):NOT NULL:DEFAULT('')"
Columns = Columns & "|TeamName:varchar:(50):NOT NULL:DEFAULT('')"
Columns = Columns & "|TeamHDCPIndex:float:NOT NULL:DEFAULT(0)"
Columns = Columns & "|TeamClass:varchar:(5):NOT NULL:DEFAULT('')"
Columns = Columns & "|CntBackPlayer1:float:NOT NULL:DEFAULT(0)"
Columns = Columns & "|CntBackPlayer2:float:NOT NULL:DEFAULT(0)"
Columns = Columns & "|CntBackPlayerTot:float:NOT NULL:DEFAULT(0)"
Columns = Columns & "|GrossPts1:float:NOT NULL:DEFAULT(0)"
Columns = Columns & "|GrossPts2:float:NOT NULL:DEFAULT(0)"
Columns = Columns & "|GrossPtsTot:float:NOT NULL:DEFAULT(0)"

Clustered = ""
Clustered = Clustered & "|Sorting"

ColumnsDet = ""
ColumnsDet = ColumnsDet & "|PlayerName:varchar:(100):NOT NULL:DEFAULT('')"
ColumnsDet = ColumnsDet & "|HDCPIndex:float:NOT NULL:DEFAULT(0)"
ColumnsDet = ColumnsDet & "|GrossPts1:float:NOT NULL:DEFAULT(0)"
ColumnsDet = ColumnsDet & "|GrossPts2:float:NOT NULL:DEFAULT(0)"
ColumnsDet = ColumnsDet & "|GrossPtsTot:float:NOT NULL:DEFAULT(0)"

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

'TeamAverage
If Trim(cmbDivisionStableford.List(cmbDivisionStableford.ListIndex)) = "ALL" Then
    If TeamAverage = 2 Then
        s = "SELECT PK, (CASE LTRIM(RTRIM(TeamName)) WHEN '' THEN TeamID ELSE LTRIM(RTRIM(TeamName)) END) as TeamName, TeamIndex as TeamHDCP, " & _
            " TeamID, (SELECT tbl_Scoring_TournamentInfo_Index.Class " & _
            " From tbl_Scoring_TournamentInfo_Index " & _
            " WHERE (tbl_Scoring_TournamentInfo_Index.TournamentKey = " & TournamentKey & ") " & _
            " AND (tbl_Scoring_TournamentInfo_Index.HFrom <= tbl_Scoring_Team.TeamIndex) " & _
            " AND (tbl_Scoring_TournamentInfo_Index.HTo >= tbl_Scoring_Team.TeamIndex)) AS Class " & _
            " From tbl_Scoring_Team " & _
            " WHERE (TournamentKey = " & TournamentKey & ") "
    Else
        s = "SELECT PK, (CASE LTRIM(RTRIM(TeamName)) WHEN '' THEN TeamID ELSE LTRIM(RTRIM(TeamName)) END) as TeamName, TeamHDCP, " & _
            " TeamID, (SELECT tbl_Scoring_TournamentInfo_Class.Class " & _
            " From tbl_Scoring_TournamentInfo_Class " & _
            " WHERE (tbl_Scoring_TournamentInfo_Class.TournamentKey = " & TournamentKey & ") " & _
            " AND (tbl_Scoring_TournamentInfo_Class.HFrom <= tbl_Scoring_Team.TeamHDCP) " & _
            " AND (tbl_Scoring_TournamentInfo_Class.HTo >= tbl_Scoring_Team.TeamHDCP)) AS Class " & _
            " From tbl_Scoring_Team " & _
            " WHERE (TournamentKey = " & TournamentKey & ") "
    End If
Else
    If TeamAverage = 2 Then
        s = "SELECT PK, (CASE LTRIM(RTRIM(TeamName)) WHEN '' THEN TeamID ELSE LTRIM(RTRIM(TeamName)) END) as TeamName, TeamIndex as TeamHDCP, " & _
            " TeamID, (SELECT tbl_Scoring_TournamentInfo_Index.Class " & _
            " From tbl_Scoring_TournamentInfo_Index " & _
            " WHERE (tbl_Scoring_TournamentInfo_Index.TournamentKey = " & TournamentKey & ") " & _
            " AND (tbl_Scoring_TournamentInfo_Index.HFrom <= tbl_Scoring_Team.TeamIndex) " & _
            " AND (tbl_Scoring_TournamentInfo_Index.HTo >= tbl_Scoring_Team.TeamIndex)) AS Class " & _
            " From tbl_Scoring_Team " & _
            " WHERE (TournamentKey = " & TournamentKey & ") " & _
            " AND ((SELECT tbl_Scoring_TournamentInfo_Index.Class " & _
            " From tbl_Scoring_TournamentInfo_Index " & _
            " WHERE (tbl_Scoring_TournamentInfo_Index.TournamentKey = " & TournamentKey & ") " & _
            " AND (tbl_Scoring_TournamentInfo_Index.HFrom <= tbl_Scoring_Team.TeamIndex) " & _
            " AND (tbl_Scoring_TournamentInfo_Index.HTo >= tbl_Scoring_Team.TeamIndex)) = '" & cmbDivisionStableford.List(cmbDivisionStableford.ListIndex) & "')"
    Else
        s = "SELECT PK, (CASE LTRIM(RTRIM(TeamName)) WHEN '' THEN TeamID ELSE LTRIM(RTRIM(TeamName)) END) as TeamName, TeamHDCP, " & _
            " TeamID, (SELECT tbl_Scoring_TournamentInfo_Class.Class " & _
            " From tbl_Scoring_TournamentInfo_Class " & _
            " WHERE (tbl_Scoring_TournamentInfo_Class.TournamentKey = " & TournamentKey & ") " & _
            " AND (tbl_Scoring_TournamentInfo_Class.HFrom <= tbl_Scoring_Team.TeamHDCP) " & _
            " AND (tbl_Scoring_TournamentInfo_Class.HTo >= tbl_Scoring_Team.TeamHDCP)) AS Class " & _
            " From tbl_Scoring_Team " & _
            " WHERE (TournamentKey = " & TournamentKey & ") " & _
            " AND ((SELECT tbl_Scoring_TournamentInfo_Class.Class " & _
            " From tbl_Scoring_TournamentInfo_Class " & _
            " WHERE (tbl_Scoring_TournamentInfo_Class.TournamentKey = " & TournamentKey & ") " & _
            " AND (tbl_Scoring_TournamentInfo_Class.HFrom <= tbl_Scoring_Team.TeamHDCP) " & _
            " AND (tbl_Scoring_TournamentInfo_Class.HTo >= tbl_Scoring_Team.TeamHDCP)) = '" & cmbDivisionStableford.List(cmbDivisionStableford.ListIndex) & "')"
    End If
End If
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    DoEvents
    cnt = cnt + 1
    ConnOmega.Execute "INSERT INTO " & TableName & " " & _
                      " (" & sMasterFields & ") " & _
                      " VALUES (0, " & rs!PK & ", " & _
                      " '" & rs!TeamID & "', " & _
                      " '" & FORMATSQL(rs!TeamName) & "', " & _
                      " " & rs!TeamHDCP & ", " & _
                      " '" & rs!Class & "', " & _
                      " 0, 0, 0, 0, 0, 0)"

    MasterKey = 0
    t = "SELECT PK " & _
        " FROM " & TableName & " " & _
        " WHERE (TeamID = '" & rs!TeamID & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        MasterKey = rt!PK
    End If
    rt.Close
    
    j = 0
    If TeamAverage = 2 Then
        t = "SELECT tbl_Scoring_Team_Detail.PlayerKey, tbl_Scoring_PlayerName.LastName, " & _
            " tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, " & _
            " tbl_Scoring_PlayerName.iIndex as HandiCap, tbl_Scoring_Team_Detail.TeamKey " & _
            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " WHERE (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ")"
    Else
        t = "SELECT tbl_Scoring_Team_Detail.PlayerKey, tbl_Scoring_PlayerName.LastName, " & _
            " tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, " & _
            " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_Team_Detail.TeamKey " & _
            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " WHERE (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ")"
    End If
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        j = j + 1
        
        ConnOmega.Execute "INSERT INTO " & DetailTableName & " " & _
                          " (" & sDetailFields & ") " & _
                          " VALUES (" & MasterKey & ", " & j & ", " & _
                          " '" & FORMATSQL(rt!LastName & ",  " & rt!FirstName & IIf(Trim(rt!MiddleName) = "", "", "  " & rt!MiddleName)) & "', " & _
                          " " & CDbl(rt!Handicap) & ", " & _
                          " 0, 0, 0)"
        
        x = 0
        For i = 2 To cmbDay.ListCount
            x = x + 1
            u = "SELECT TOP 1 GrossPoints " & _
                " FROM tbl_Scoring_ScoreCard " & _
                " WHERE (PlayerKey = " & rt!PlayerKey & ") " & _
                " AND (LocationKey = " & cmbDay.ItemData(i - 1) & ")"
            If ru.State = adStateOpen Then ru.Close
            ru.Open u, ConnOmega
            If ru.RecordCount > 0 Then
                ConnOmega.Execute "UPDATE " & DetailTableName & " " & _
                                  " SET GrossPts" & x & " = " & ru!GrossPoints & " " & _
                                  " WHERE (MasterKey = " & MasterKey & ") " & _
                                  " AND (Line = " & j & ")"
            End If
            ru.Close
        Next i
        rt.MoveNext
    Wend
    rt.Close
    
    dblGrossPts1 = 0
    t = "SELECT TOP " & TeamPlayer2Cnt & " * " & _
        " FROM " & DetailTableName & " " & _
        " WHERE (MasterKey = " & MasterKey & ") " & _
        " ORDER BY GrossPts1 DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        dblGrossPts1 = dblGrossPts1 + CDbl(rt!GrossPts1)
        rt.MoveNext
    Wend
    rt.Close
    
    dblCntBackPlayer1 = 0
    t = "SELECT TOP 1 * " & _
        " FROM " & DetailTableName & " " & _
        " WHERE (MasterKey = " & MasterKey & ") " & _
        " ORDER BY GrossPts1"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        dblCntBackPlayer1 = CDbl(rt!GrossPts1)
        rt.MoveNext
    Wend
    rt.Close
    
    dblGrossPts2 = 0
    t = "SELECT TOP " & TeamPlayer2Cnt & " * " & _
        " FROM " & DetailTableName & " " & _
        " WHERE (MasterKey = " & MasterKey & ") " & _
        " ORDER BY GrossPts2 DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        dblGrossPts2 = dblGrossPts2 + CDbl(rt!GrossPts2)
        rt.MoveNext
    Wend
    rt.Close
    
    dblCntBackPlayer2 = 0
    t = "SELECT TOP 1 * " & _
        " FROM " & DetailTableName & " " & _
        " WHERE (MasterKey = " & MasterKey & ") " & _
        " ORDER BY GrossPts2"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        dblCntBackPlayer2 = CDbl(rt!GrossPts2)
        rt.MoveNext
    Wend
    rt.Close
    
    dblGrossPtsTot = CDbl(dblGrossPts1) + CDbl(dblGrossPts2)
    dblCntBackPlayerTot = CDbl(dblCntBackPlayer1) + CDbl(dblCntBackPlayer2)
    
    ConnOmega.Execute "UPDATE " & TableName & " " & _
                      " SET GrossPts1 = " & dblGrossPts1 & ", " & _
                      " CntBackPlayer1 = " & dblCntBackPlayer1 & ", " & _
                      " GrossPts2 = " & dblGrossPts2 & ", " & _
                      " CntBackPlayer2 = " & dblCntBackPlayer2 & ", " & _
                      " CntBackPlayerTot = " & CDbl(dblCntBackPlayerTot) & ", " & _
                      " GrossPtsTot = " & CDbl(dblGrossPtsTot) & " " & _
                      " WHERE (PK = " & MasterKey & ")"

    
    UpdateProgress picProgressBar, cnt / rs.RecordCount
        
    rs.MoveNext
Wend
rs.Close

i = 0
s = "SELECT * " & _
    " FROM " & TableName & " " & _
    " ORDER BY GrossPtsTot DESC, CntBackPlayerTot DESC"
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

'Excel File
ColTop = 0: RowTop = 0
ColCount = 0: RowCount = 0
cnt = 0
Set xlsApp = CreateObject("Excel.Application")
With xlsApp
    .Visible = False
    
    .Workbooks.Add
    .DisplayAlerts = False
    .Workbooks(1).Sheets(1).Activate
    .Workbooks(1).Sheets(1).Name = cmbReportTypeStableford.List(cmbReportTypeStableford.ListIndex) & " (" & Replace(cmbDivisionStableford.List(cmbDivisionStableford.ListIndex), "SORT BY ", "") & ")" '"Report"
    .Workbooks(1).Sheets(2).Delete
    .Workbooks(1).Sheets(2).Delete
    
    With xlsApp.ActiveWorkbook.Sheets(1)
        s = "SELECT * FROM " & TableName & "" & _
            " ORDER BY Sorting"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        '== Header
        RowCount = RowCount + 1
        ColCount = ColCount + 1
        strRange = EXCEL_RANGE(ColCount, RowCount)
        strRange1 = EXCEL_RANGE(rs.Fields.Count - 3, RowCount)
        .Range(strRange, strRange1).Select
        xlsApp.Selection.Merge
        
        strRange = EXCEL_RANGE(ColCount, RowCount)
        .Range(strRange).Value = gbl_CompanyName
        .Range(strRange).Font.Name = "Script MT Bold" '"Tahoma"
        .Range(strRange).Font.Size = 10
        .Range(strRange).Font.Bold = True
        .Range(strRange).HorizontalAlignment = 3
        .Range(strRange).VerticalAlignment = 2
        
        ColCount = 0
        RowCount = RowCount + 1
        ColCount = ColCount + 1
        strRange = EXCEL_RANGE(ColCount, RowCount)
        strRange1 = EXCEL_RANGE(rs.Fields.Count - 3, RowCount)
        .Range(strRange, strRange1).Select
        xlsApp.Selection.Merge
        
        strRange = EXCEL_RANGE(ColCount, RowCount)
        .Range(strRange).Value = gbl_CompanyAddress1
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 8
        .Range(strRange).Font.Bold = False
        .Range(strRange).HorizontalAlignment = 3
        .Range(strRange).VerticalAlignment = 2
        
        ColCount = 0
        RowCount = RowCount + 1
        ColCount = ColCount + 1
        strRange = EXCEL_RANGE(ColCount, RowCount)
        strRange1 = EXCEL_RANGE(rs.Fields.Count - 3, RowCount)
        .Range(strRange, strRange1).Select
        xlsApp.Selection.Merge
        
        strRange = EXCEL_RANGE(ColCount, RowCount)
        .Range(strRange).Value = gbl_CompanyAddress2
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 8
        .Range(strRange).Font.Bold = False
        .Range(strRange).HorizontalAlignment = 3
        .Range(strRange).VerticalAlignment = 2
        
        ColCount = 0
        RowCount = RowCount + 1
        ColCount = ColCount + 1
        strRange = EXCEL_RANGE(ColCount, RowCount)
        strRange1 = EXCEL_RANGE(rs.Fields.Count - 3, RowCount)
        .Range(strRange, strRange1).Select
        xlsApp.Selection.Merge
        
        strRange = EXCEL_RANGE(ColCount, RowCount)
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
        strRange = EXCEL_RANGE(ColCount, RowCount)
        strRange1 = EXCEL_RANGE(rs.Fields.Count - 3, RowCount)
        .Range(strRange, strRange1).Select
        xlsApp.Selection.Merge
        strRange = EXCEL_RANGE(ColCount, RowCount)
        .Range(strRange).Value = TournamentName
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 9
        .Range(strRange).Font.Bold = True
        .Range(strRange).HorizontalAlignment = 3
        .Range(strRange).VerticalAlignment = 2
        
        ColCount = 0
        RowCount = RowCount + 1
        ColCount = ColCount + 1
        strRange = EXCEL_RANGE(ColCount, RowCount)
        strRange1 = EXCEL_RANGE(rs.Fields.Count - 3, RowCount)
        .Range(strRange, strRange1).Select
        xlsApp.Selection.Merge
        strRange = EXCEL_RANGE(ColCount, RowCount)
        .Range(strRange).Value = TournamentRange
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 8
        .Range(strRange).Font.Bold = False
        .Range(strRange).HorizontalAlignment = 3
        .Range(strRange).VerticalAlignment = 2
                        
                        
        If Trim(cmbDivisionStableford.List(cmbDivisionStableford.ListIndex)) <> "ALL" Then
            ColCount = 0
            RowCount = RowCount + 1
            ColCount = ColCount + 1
            strRange = EXCEL_RANGE(ColCount, RowCount)
            strRange1 = EXCEL_RANGE(rs.Fields.Count - 3, RowCount)
            .Range(strRange, strRange1).Select
            xlsApp.Selection.Merge
            
            strRange = EXCEL_RANGE(ColCount, RowCount)
            .Range(strRange).Value = "CLASS " & Trim(cmbDivisionStableford.List(cmbDivisionStableford.ListIndex))
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 10
            .Range(strRange).Font.Bold = True
            .Range(strRange).HorizontalAlignment = 3
            .Range(strRange).VerticalAlignment = 2
        End If
       
        ColCount = 0
        RowCount = RowCount + 1
        ColCount = ColCount + 1
        strRange = EXCEL_RANGE(ColCount, RowCount)
        strRange1 = EXCEL_RANGE(rs.Fields.Count - 3, RowCount)
        .Range(strRange, strRange1).Select
        xlsApp.Selection.Merge
        
        strRange = EXCEL_RANGE(ColCount, RowCount)
        .Range(strRange).Value = "LEGEND"
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 10
        .Range(strRange).Font.Bold = True
        .Range(strRange).HorizontalAlignment = 1 '3
        .Range(strRange).VerticalAlignment = 2
        
        x = 0
        For i = 2 To cmbDay.ListCount
            x = x + 1
            ColCount = 0
            RowCount = RowCount + 1
            ColCount = ColCount + 1
            strRange = EXCEL_RANGE(ColCount, RowCount)
            strRange1 = EXCEL_RANGE(rs.Fields.Count - 3, RowCount)
            .Range(strRange, strRange1).Select
            xlsApp.Selection.Merge
            
            strRange = EXCEL_RANGE(ColCount, RowCount)
            .Range(strRange).Value = "Loc " & x & " : " & cmbDay.List(i - 1)
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = False
            .Range(strRange).HorizontalAlignment = 1 '3
            .Range(strRange).VerticalAlignment = 2
        Next i
        
        ColCount = 0
        RowCount = RowCount + 1
        ColCount = ColCount + 1
        strRange = EXCEL_RANGE(ColCount, RowCount)
        strRange1 = EXCEL_RANGE(rs.Fields.Count - 3, RowCount)
        .Range(strRange, strRange1).Select
        xlsApp.Selection.Merge
        
        strRange = EXCEL_RANGE(ColCount, RowCount)
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
        strRange = EXCEL_RANGE(ColCount, RowCount)
        .Range(strRange).Value = "Team Name"
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 8
        .Range(strRange).Font.Bold = False
        .Range(strRange).HorizontalAlignment = 3
        .Range(strRange).Interior.ColorIndex = 15
        .Range(strRange).Interior.Pattern = 1 'xlSolid
        .Range(strRange).Select
        xlsApp.Selection.Borders.LineStyle = 1
            
        ColCount = ColCount + 1
        strRange = EXCEL_RANGE(ColCount, RowCount)
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
        strRange = EXCEL_RANGE(ColCount, RowCount)
        If TeamAverage = 2 Then
            .Range(strRange).Value = "Index"
        Else
            .Range(strRange).Value = "H D C P"
        End If
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 8
        .Range(strRange).Font.Bold = False
        .Range(strRange).HorizontalAlignment = 4
        .Range(strRange).Interior.ColorIndex = 15
        .Range(strRange).Interior.Pattern = 1 'xlSolid
        .Range(strRange).Select
        xlsApp.Selection.Borders.LineStyle = 1
        
        ColCount = ColCount + 1
        strRange = EXCEL_RANGE(ColCount, RowCount)
        If TeamAverage = 2 Then
            .Range(strRange).Value = "Team Index"
        Else
            .Range(strRange).Value = "Team HDCP"
        End If
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 8
        .Range(strRange).Font.Bold = False
        .Range(strRange).HorizontalAlignment = 3
        .Range(strRange).Interior.ColorIndex = 15
        .Range(strRange).Interior.Pattern = 1 'xlSolid
        .Range(strRange).Select
        xlsApp.Selection.Borders.LineStyle = 1
                
        ColCount = ColCount + 1
        strRange = EXCEL_RANGE(ColCount, RowCount)
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
        strRange = EXCEL_RANGE(ColCount, RowCount)
        .Range(strRange).Value = "Loc 1"
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 8
        .Range(strRange).Font.Bold = False
        .Range(strRange).HorizontalAlignment = 4
        .Range(strRange).Interior.ColorIndex = 15
        .Range(strRange).Interior.Pattern = 1 'xlSolid
        .Range(strRange).Select
        xlsApp.Selection.Borders.LineStyle = 1
        
        ColCount = ColCount + 1
        strRange = EXCEL_RANGE(ColCount, RowCount)
        .Range(strRange).Value = "Loc 1 Tot"
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 8
        .Range(strRange).Font.Bold = False
        .Range(strRange).HorizontalAlignment = 3
        .Range(strRange).Interior.ColorIndex = 15
        .Range(strRange).Interior.Pattern = 1 'xlSolid
        .Range(strRange).Select
        xlsApp.Selection.Borders.LineStyle = 1
        
        ColCount = ColCount + 1
        strRange = EXCEL_RANGE(ColCount, RowCount)
        .Range(strRange).Value = "Loc 2"
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 8
        .Range(strRange).Font.Bold = False
        .Range(strRange).HorizontalAlignment = 4
        .Range(strRange).Interior.ColorIndex = 15
        .Range(strRange).Interior.Pattern = 1 'xlSolid
        .Range(strRange).Select
        xlsApp.Selection.Borders.LineStyle = 1
        
        ColCount = ColCount + 1
        strRange = EXCEL_RANGE(ColCount, RowCount)
        .Range(strRange).Value = "Loc 2 Tot"
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 8
        .Range(strRange).Font.Bold = False
        .Range(strRange).HorizontalAlignment = 3
        .Range(strRange).Interior.ColorIndex = 15
        .Range(strRange).Interior.Pattern = 1 'xlSolid
        .Range(strRange).Select
        xlsApp.Selection.Borders.LineStyle = 1
        
        ColCount = ColCount + 1
        strRange = EXCEL_RANGE(ColCount, RowCount)
        .Range(strRange).Value = "Total Pts"
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 8
        .Range(strRange).Font.Bold = False
        .Range(strRange).HorizontalAlignment = 3
        .Range(strRange).Interior.ColorIndex = 15
        .Range(strRange).Interior.Pattern = 1 'xlSolid
        .Range(strRange).Select
        xlsApp.Selection.Borders.LineStyle = 1
        
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
                " ORDER BY HDCPIndex"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            While Not rt.EOF
                RowTo = RowTo + 1
                rt.MoveNext
            Wend
            rt.Close
            RowTo = RowTo - 1
            
            strRange = EXCEL_RANGE(ColCount, RowFrom)
            strRange1 = EXCEL_RANGE(ColCount, RowTo)
            .Range(strRange, strRange1).Select
            xlsApp.Selection.Merge
            strRange = EXCEL_RANGE(ColCount, RowFrom)
            .Range(strRange).Value = rs!TeamName
            .Range(strRange).Font.Name = "Courier New"
            .Range(strRange).Font.Size = 9
            .Range(strRange).Font.Bold = False
            .Columns(ColCount).ColumnWidth = 30
            .Range(strRange).HorizontalAlignment = 3
            .Range(strRange).VerticalAlignment = 2
            .Range(strRange).Select
            xlsApp.Selection.Borders.LineStyle = 1
                        
            ColCount = ColCount + 1
            
            '== Player Info
            ColCountDet = ColCount
            RowCountDet = RowFrom
            t = "SELECT * FROM " & DetailTableName & " " & _
                " WHERE (MasterKey = " & rs!PK & ") " & _
                " ORDER BY HDCPIndex"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            RowFrom = RowCount
            RowTo = RowFrom
            While Not rt.EOF
                strRange = EXCEL_RANGE(ColCountDet, RowCountDet)
                .Range(strRange).Value = rt!PlayerName
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 10
                .Range(strRange).Font.Color = vbBlue
                .Columns(ColCountDet).ColumnWidth = 28
                .Range(strRange).Select
                xlsApp.Selection.Borders.LineStyle = 1
                
                ColCountDet = ColCountDet + 1
                strRange = EXCEL_RANGE(ColCountDet, RowCountDet)
                .Range(strRange).Value = rt!HDCPIndex
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
            
            ColCount = ColCount + (rt.Fields.Count) - 5 '2
            rt.Close
            
            RowTo = RowTo - 1
            
            strRange = EXCEL_RANGE(ColCount, RowFrom)
            strRange1 = EXCEL_RANGE(ColCount, RowTo)
            .Range(strRange, strRange1).Select
            xlsApp.Selection.Merge
            strRange = EXCEL_RANGE(ColCount, RowFrom)
            .Range(strRange).Value = rs!TeamHDCPIndex
            .Range(strRange).Font.Name = "Courier New"
            .Range(strRange).Font.Size = 13
            .Range(strRange).Font.Bold = True
            .Range(strRange).HorizontalAlignment = 3
            .Range(strRange).VerticalAlignment = 2
            .Range(strRange).Select
            xlsApp.Selection.Borders.LineStyle = 1
                
            
            ColCount = ColCount + 1
            strRange = EXCEL_RANGE(ColCount, RowFrom)
            strRange1 = EXCEL_RANGE(ColCount, RowTo)
            .Range(strRange, strRange1).Select
            xlsApp.Selection.Merge
            strRange = EXCEL_RANGE(ColCount, RowFrom)
            .Range(strRange).Value = rs!TeamClass
            .Range(strRange).Font.Name = "Courier New"
            .Range(strRange).Font.Size = 13
            .Range(strRange).Font.Bold = True
            .Range(strRange).HorizontalAlignment = 3
            .Range(strRange).VerticalAlignment = 2
            .Range(strRange).Select
            xlsApp.Selection.Borders.LineStyle = 1
                        
            
            ' Player Pts1
            ColCountDet = ColCount
            RowCountDet = RowFrom
            t = "SELECT * FROM " & DetailTableName & " " & _
                " WHERE (MasterKey = " & rs!PK & ") " & _
                " ORDER BY HDCPIndex"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            RowFrom = RowCount
            RowTo = RowFrom
            While Not rt.EOF

                ColCountDet = ColCountDet + 1
                strRange = EXCEL_RANGE(ColCountDet, RowCountDet)
                .Range(strRange).Value = rt!GrossPts1
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
            ColCount = ColCount + (rt.Fields.Count) - 5 '2
            rt.Close
            RowTo = RowTo - 1
            'RowCount = RowTo
            
            strRange = EXCEL_RANGE(ColCount, RowFrom)
            strRange1 = EXCEL_RANGE(ColCount, RowTo)
            .Range(strRange, strRange1).Select
            xlsApp.Selection.Merge
            strRange = EXCEL_RANGE(ColCount, RowFrom)
            .Range(strRange).Value = rs!GrossPts1
            .Range(strRange).Font.Name = "Courier New"
            .Range(strRange).Font.Size = 13
            .Range(strRange).Font.Bold = True
            .Range(strRange).HorizontalAlignment = 3
            .Range(strRange).VerticalAlignment = 2
            .Range(strRange).Select
            xlsApp.Selection.Borders.LineStyle = 1
            
            
            
            
            ' Player Pts2
            ColCountDet = ColCount
            RowCountDet = RowFrom
            t = "SELECT * FROM " & DetailTableName & " " & _
                " WHERE (MasterKey = " & rs!PK & ") " & _
                " ORDER BY HDCPIndex"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            RowFrom = RowCount
            RowTo = RowFrom
            While Not rt.EOF

                ColCountDet = ColCountDet + 1
                strRange = EXCEL_RANGE(ColCountDet, RowCountDet)
                .Range(strRange).Value = rt!GrossPts2
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
            
            ColCount = ColCount + (rt.Fields.Count) - 5 '2
            rt.Close
            RowTo = RowTo - 1
            'RowCount = RowTo
            
            strRange = EXCEL_RANGE(ColCount, RowFrom)
            strRange1 = EXCEL_RANGE(ColCount, RowTo)
            .Range(strRange, strRange1).Select
            xlsApp.Selection.Merge
            strRange = EXCEL_RANGE(ColCount, RowFrom)
            .Range(strRange).Value = rs!GrossPts2
            .Range(strRange).Font.Name = "Courier New"
            .Range(strRange).Font.Size = 13
            .Range(strRange).Font.Bold = True
            .Range(strRange).HorizontalAlignment = 3
            .Range(strRange).VerticalAlignment = 2
            .Range(strRange).Select
            xlsApp.Selection.Borders.LineStyle = 1
            
            ColCount = ColCount + 1
            strRange = EXCEL_RANGE(ColCount, RowFrom)
            strRange1 = EXCEL_RANGE(ColCount, RowTo)
            .Range(strRange, strRange1).Select
            xlsApp.Selection.Merge
            strRange = EXCEL_RANGE(ColCount, RowFrom)
            .Range(strRange).Value = rs!GrossPtsTot
            .Range(strRange).Font.Name = "Courier New"
            .Range(strRange).Font.Size = 13
            .Range(strRange).Font.Bold = True
            .Range(strRange).HorizontalAlignment = 3
            .Range(strRange).VerticalAlignment = 2
            .Range(strRange).Select
            xlsApp.Selection.Borders.LineStyle = 1
            
            RowCount = RowTo
            
            UpdateProgress_Caption "Generating Excel Output", picProgressBar, cnt / rs.RecordCount
            
            rs.MoveNext
        Wend
        rs.Close
        
        '.Range("A12").Select
'        .Range(sFreezePane).Select
'        .ActiveWindow.FreezePanes = True
        
        .PageSetup.PaperSize = 1 'Letter
        .PageSetup.Orientation = 2 '2 'LandScape
        .PageSetup.TopMargin = 3
        .PageSetup.LeftMargin = 3
        .PageSetup.RightMargin = 3
        .PageSetup.BottomMargin = 3
        .PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
    End With
    
    If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
    .ActiveWorkbook.SaveAs Filename:=WorkbookName
    .Visible = True
    Set xlsApp = Nothing
End With


picProgress.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True
Screen.MousePointer = vbDefault

Exit Sub
ErrorHandler:
Screen.MousePointer = vbDefault
Exit Sub

Exit Sub
PG:
Screen.MousePointer = vbDefault
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "ScoreCardControlAll", "ScoreCardControlAll", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "ScoreCardControlAll", "ScoreCardControlAll", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "ScoreCardControlAll", "ScoreCardControlAll", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "ScoreCardControlAll", "ScoreCardControlAll", ""), "is_END"
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

Private Sub txtGrossPts_Change(Index As Integer)
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
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
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
txtGrossPtsTot.Text = RETURNTEXTVALUE(txtGrossPtsF) + _
                      RETURNTEXTVALUE(txtGrossPtsB)
txtSGrossF.Text = RETURNTEXTVALUE(txtGrossPtsF)
End Sub

Private Sub txtGrossPtsF_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
txtGrossPtsTot.Text = RETURNTEXTVALUE(txtGrossPtsF) + _
                      RETURNTEXTVALUE(txtGrossPtsB)
txtSGrossF.Text = RETURNTEXTVALUE(txtGrossPtsF)
End Sub

Private Sub txtGrossScore_Change(Index As Integer)
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub

If RETURNTEXTVALUE(txtGrossScore(Index)) <= 0 Then txtGrossPts(Index).Text = "0": txtNetPts(Index).Text = "0": Exit Sub

With FGrid
    Select Case ScoringType
        Case 3
        Case 4
        Case Else   'Stableford & Modified Stableford & Modified Molave
            Select Case Index
                Case 0
                    dblPar = .TextMatrix(1, 2)
                    txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                    dblHandicap = .TextMatrix(2, 2)
                    txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                Case 1
                    dblPar = .TextMatrix(1, 3)
                    txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                    dblHandicap = .TextMatrix(2, 3)
                    txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                Case 2
                    dblPar = .TextMatrix(1, 4)
                    txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                    dblHandicap = .TextMatrix(2, 4)
                    txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                Case 3
                    dblPar = .TextMatrix(1, 5)
                    txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                    dblHandicap = .TextMatrix(2, 5)
                    txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                Case 4
                    dblPar = .TextMatrix(1, 6)
                    txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                    dblHandicap = .TextMatrix(2, 6)
                    txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                Case 5
                    dblPar = .TextMatrix(1, 7)
                    txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                    dblHandicap = .TextMatrix(2, 7)
                    txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                Case 6
                    dblPar = .TextMatrix(1, 8)
                    txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                    dblHandicap = .TextMatrix(2, 8)
                    txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                Case 7
                    dblPar = .TextMatrix(1, 9)
                    txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                    dblHandicap = .TextMatrix(2, 9)
                    txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                Case 8
                    dblPar = .TextMatrix(1, 10)
                    txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                    dblHandicap = .TextMatrix(2, 10)
                    txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                Case 9
                    dblPar = .TextMatrix(1, 12)
                    txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                    dblHandicap = .TextMatrix(2, 12)
                    txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                Case 10
                    dblPar = .TextMatrix(1, 13)
                    txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                    dblHandicap = .TextMatrix(2, 13)
                    txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                Case 11
                    dblPar = .TextMatrix(1, 14)
                    txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                    dblHandicap = .TextMatrix(2, 14)
                    txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                Case 12
                    dblPar = .TextMatrix(1, 15)
                    txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                    dblHandicap = .TextMatrix(2, 15)
                    txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                Case 13
                    dblPar = .TextMatrix(1, 16)
                    txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) '+ 1
                    dblHandicap = .TextMatrix(2, 16)
                    txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                Case 14
                    dblPar = .TextMatrix(1, 17)
                    txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                    dblHandicap = .TextMatrix(2, 17)
                    txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                Case 15
                    dblPar = .TextMatrix(1, 18)
                    txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                    dblHandicap = .TextMatrix(2, 18)
                    txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                Case 16
                    dblPar = .TextMatrix(1, 19)
                    txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                    dblHandicap = .TextMatrix(2, 19)
                    txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                Case 17
                    dblPar = .TextMatrix(1, 20)
                    txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
                    dblHandicap = .TextMatrix(2, 20)
                    txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index)))) ' + 1
            End Select
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
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
txtGrossScoreTot.Text = RETURNTEXTVALUE(txtGrossScoreF) + _
                        RETURNTEXTVALUE(txtGrossScoreB)
End Sub

Private Sub txtGrossScoreF_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
txtGrossScoreTot.Text = RETURNTEXTVALUE(txtGrossScoreF) + _
                        RETURNTEXTVALUE(txtGrossScoreB)
End Sub

Private Sub txtNetPts_Change(Index As Integer)
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
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
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
txtNetPtsTot.Text = RETURNTEXTVALUE(txtNetPtsF) + _
                    RETURNTEXTVALUE(txtNetPtsB)
txtSNetF.Text = RETURNTEXTVALUE(txtNetPtsF)
End Sub

Private Sub txtNetPtsF_Change()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
txtNetPtsTot.Text = RETURNTEXTVALUE(txtNetPtsF) + _
                    RETURNTEXTVALUE(txtNetPtsB)
txtSNetF.Text = RETURNTEXTVALUE(txtNetPtsF)
End Sub

Private Sub txtSearch_Change()
If cmbLocationSearch.ListIndex = -1 Then Exit Sub
If Trim(txtSearch.Text) = "" Then lstResult.Clear: cmbDate.Clear: Exit Sub
lstResult.Clear: cmbDate.Clear
    s = "SELECT tbl_Scoring_PlayerName.PK, " & _
        " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName " & _
        " FROM tbl_Scoring_ScoreCard LEFT OUTER JOIN " & _
        " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
        " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") " & _
        " AND (tbl_Scoring_ScoreCard.LocationKey = " & cmbLocationSearch.ItemData(cmbLocationSearch.ListIndex) & ") " & _
        " AND (tbl_Scoring_PlayerName.LastName LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
        " GROUP BY tbl_Scoring_PlayerName.PK, tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName " & _
        " ORDER BY tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResult.AddItem rs!PlayerName
    lstResult.ItemData(lstResult.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResult.ListCount Then lstResult.ListIndex = 0
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResult.SetFocus
End Sub

Private Sub txtSearchAdd_Change()
If Trim(txtSearchAdd.Text) = "" Then lstResultAdd.Clear: Exit Sub
lstResultAdd.Clear
s = "SELECT PK, LastName + ',  ' + FirstName + '  ' + MiddleName AS PlayerName " & _
    " From tbl_Scoring_PlayerName " & _
    " WHERE (LastName LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%') " & _
    " AND (TournamentKey = " & TournamentKey & ") " & _
    " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultAdd.AddItem rs!PlayerName
    lstResultAdd.ItemData(lstResultAdd.NewIndex) = rs!PK
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

Private Sub txtTop_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKPrint_Click
End Sub
