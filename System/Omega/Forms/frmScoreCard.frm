VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmScoreCard 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   330
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
   Icon            =   "frmScoreCard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picPrint 
      Height          =   2775
      Left            =   4080
      TabIndex        =   129
      Top             =   2160
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4895
      BackColor       =   15396057
      Begin VB.Timer TimerTeamAll 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   120
         Top             =   2280
      End
      Begin VB.ComboBox cmbDay 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   142
         Top             =   1680
         Width           =   4095
      End
      Begin VB.ComboBox cmbReportType 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   137
         Top             =   480
         Width           =   4095
      End
      Begin VB.ComboBox cmbDivision 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   136
         Top             =   1320
         Width           =   4095
      End
      Begin VB.ComboBox cmbGroup 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   134
         Top             =   915
         Width           =   4095
      End
      Begin VB.CommandButton cmdCancelPrint 
         Height          =   480
         Left            =   2235
         Picture         =   "frmScoreCard.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   133
         Top             =   2115
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKPrint 
         Height          =   480
         Left            =   555
         Picture         =   "frmScoreCard.frx":1026
         Style           =   1  'Graphical
         TabIndex        =   132
         Top             =   2115
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar3 
         Height          =   345
         Left            =   40
         TabIndex        =   135
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
         Icon            =   "frmScoreCard.frx":1698
      End
   End
   Begin VB.Timer TimerAddLocation 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6000
      Top             =   6720
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   840
      TabIndex        =   145
      Top             =   6840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin RPVGCC.b8Container picSearch 
      Height          =   4695
      Left            =   4080
      TabIndex        =   119
      Top             =   840
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   8281
      BackColor       =   15396057
      Begin VB.ComboBox cmbDate 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   126
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   123
         Top             =   480
         Width           =   4095
      End
      Begin VB.ListBox lstResult 
         Height          =   2595
         Left            =   120
         TabIndex        =   122
         Top             =   840
         Width           =   4095
      End
      Begin VB.CommandButton cmdOKSearch 
         Height          =   480
         Left            =   480
         Picture         =   "frmScoreCard.frx":1C32
         Style           =   1  'Graphical
         TabIndex        =   121
         Top             =   3960
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelSearch 
         Height          =   480
         Left            =   2280
         Picture         =   "frmScoreCard.frx":22A4
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   3960
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   124
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
         Icon            =   "frmScoreCard.frx":2A00
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   255
         Left            =   1200
         TabIndex        =   125
         Top             =   3480
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2880
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2160
      Top             =   6120
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
            Picture         =   "frmScoreCard.frx":2F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCard.frx":309C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCard.frx":3220
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCard.frx":353A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCard.frx":38F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCard.frx":3D45
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCard.frx":4197
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCard.frx":454F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCard.frx":4661
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCard.frx":4BA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCard.frx":4CFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCard.frx":523F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   7425
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
   Begin RPVGCC.b8Container picProgress 
      Height          =   975
      Left            =   3480
      TabIndex        =   130
      Top             =   2880
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
         TabIndex        =   131
         Top             =   120
         Width           =   5295
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6440
      Left            =   0
      Picture         =   "frmScoreCard.frx":5463
      ScaleHeight     =   6435
      ScaleWidth      =   12135
      TabIndex        =   1
      Top             =   0
      Width           =   12135
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   10800
         TabIndex        =   141
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox picToolbar 
         BorderStyle     =   0  'None
         Height          =   770
         Left            =   0
         ScaleHeight     =   765
         ScaleWidth      =   15000
         TabIndex        =   127
         Top             =   0
         Width           =   15000
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   570
            Left            =   0
            TabIndex        =   128
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
         Left            =   10440
         TabIndex        =   118
         Top             =   6240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin RPVGCC.b8Container b8Container3 
         Height          =   1455
         Left            =   5880
         TabIndex        =   2
         Top             =   840
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2566
         BackColor       =   49152
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1215
            Left            =   120
            ScaleHeight     =   1215
            ScaleWidth      =   5895
            TabIndex        =   3
            Top             =   120
            Width           =   5895
            Begin VB.TextBox txtLocation 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1080
               TabIndex        =   143
               Top             =   840
               Width           =   4695
            End
            Begin VB.TextBox txtTournament 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1095
               TabIndex        =   5
               Top             =   120
               Width           =   4695
            End
            Begin VB.TextBox txtTourDate 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1095
               TabIndex        =   4
               Text            =   "06/01/2010 - 06/04/2010"
               Top             =   480
               Width           =   4695
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "Location"
               Height          =   255
               Left            =   120
               TabIndex        =   144
               Top             =   840
               Width           =   975
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Tournament"
               Height          =   255
               Left            =   120
               TabIndex        =   7
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Date Range"
               Height          =   255
               Left            =   120
               TabIndex        =   6
               Top             =   480
               Width           =   975
            End
         End
      End
      Begin RPVGCC.b8Container b8Container5 
         Height          =   1215
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2143
         BackColor       =   49152
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   975
            Left            =   120
            ScaleHeight     =   975
            ScaleWidth      =   5415
            TabIndex        =   9
            Top             =   120
            Width           =   5415
            Begin VB.TextBox txtDate 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   720
               TabIndex        =   12
               Top             =   480
               Width           =   1575
            End
            Begin VB.TextBox txtDay 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   4560
               TabIndex        =   11
               Top             =   480
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox txtPlayer 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   720
               TabIndex        =   10
               Top             =   120
               Width           =   4575
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   480
               Width           =   495
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Day"
               Height          =   255
               Left            =   3720
               TabIndex        =   14
               Top             =   480
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Player"
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   120
               Width           =   975
            End
         End
      End
      Begin RPVGCC.b8Container b8Container2 
         Height          =   1335
         Left            =   8520
         TabIndex        =   16
         Top             =   2400
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
            TabIndex        =   17
            Top             =   120
            Width           =   3255
            Begin VB.TextBox txtSGrossF 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1080
               TabIndex        =   23
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox txtSNetF 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1080
               TabIndex        =   22
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtSGrossB 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1680
               TabIndex        =   21
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox txtSNetB 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1680
               TabIndex        =   20
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtSGrossTot 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   2400
               TabIndex        =   19
               Top             =   360
               Width           =   735
            End
            Begin VB.TextBox txtSNetTot 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   2400
               TabIndex        =   18
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Gross Points"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Net Points"
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   720
               Width           =   975
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
               TabIndex        =   27
               Top             =   0
               Width           =   975
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "F - 9"
               Height          =   255
               Left            =   1080
               TabIndex        =   26
               Top             =   120
               Width           =   495
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "B - 9"
               Height          =   255
               Left            =   1680
               TabIndex        =   25
               Top             =   120
               Width           =   495
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Total"
               Height          =   255
               Left            =   2400
               TabIndex        =   24
               Top             =   120
               Width           =   735
            End
         End
      End
      Begin RPVGCC.b8Container b8Container1 
         Height          =   2490
         Left            =   120
         TabIndex        =   30
         Top             =   3840
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
            TabIndex        =   31
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
               TabIndex        =   34
               Top             =   1650
               Width           =   12330
               Begin VB.PictureBox Picture4 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   10640
                  ScaleHeight     =   225
                  ScaleWidth      =   1860
                  TabIndex        =   98
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
                     TabIndex        =   100
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
                     TabIndex        =   99
                     Text            =   "0"
                     Top             =   -10
                     Width           =   570
                  End
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
                  TabIndex        =   96
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
                     TabIndex        =   97
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
                  TabIndex        =   95
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
                  TabIndex        =   94
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
                  TabIndex        =   93
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
                  TabIndex        =   92
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
                  TabIndex        =   91
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
                  TabIndex        =   90
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
                  TabIndex        =   89
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
                  TabIndex        =   88
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
                  TabIndex        =   87
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
                  TabIndex        =   86
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
                  TabIndex        =   85
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
                  TabIndex        =   84
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
                  TabIndex        =   83
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
                  TabIndex        =   82
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
                  TabIndex        =   81
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
                  TabIndex        =   80
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
                  TabIndex        =   79
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
                  TabIndex        =   78
                  Text            =   "0"
                  Top             =   -10
                  Width           =   460
               End
               Begin VB.PictureBox Picture2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   495
                  Left            =   1980
                  ScaleHeight     =   465
                  ScaleWidth      =   10305
                  TabIndex        =   35
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
                     TabIndex        =   77
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
                     TabIndex        =   76
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
                     TabIndex        =   75
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
                     TabIndex        =   74
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
                     TabIndex        =   73
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
                     TabIndex        =   72
                     Text            =   "0"
                     Top             =   -10
                     Width           =   570
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   17
                     Left            =   8190
                     TabIndex        =   71
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   17
                     Left            =   8190
                     TabIndex        =   70
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   16
                     Left            =   7740
                     TabIndex        =   69
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   16
                     Left            =   7740
                     TabIndex        =   68
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   15
                     Left            =   7290
                     TabIndex        =   67
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   15
                     Left            =   7290
                     TabIndex        =   66
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   14
                     Left            =   6840
                     TabIndex        =   65
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   14
                     Left            =   6840
                     TabIndex        =   64
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   13
                     Left            =   6390
                     TabIndex        =   63
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   13
                     Left            =   6390
                     TabIndex        =   62
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   12
                     Left            =   5940
                     TabIndex        =   61
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   12
                     Left            =   5940
                     TabIndex        =   60
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   11
                     Left            =   5490
                     TabIndex        =   59
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   11
                     Left            =   5490
                     TabIndex        =   58
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   10
                     Left            =   5040
                     TabIndex        =   57
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   10
                     Left            =   5040
                     TabIndex        =   56
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   9
                     Left            =   4590
                     TabIndex        =   55
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   9
                     Left            =   4590
                     TabIndex        =   54
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   8
                     Left            =   3580
                     TabIndex        =   53
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   8
                     Left            =   3580
                     TabIndex        =   52
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   7
                     Left            =   3130
                     TabIndex        =   51
                     Text            =   "0"
                     Top             =   225
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   7
                     Left            =   3130
                     TabIndex        =   50
                     Text            =   "0"
                     Top             =   -15
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   6
                     Left            =   2680
                     TabIndex        =   49
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   6
                     Left            =   2680
                     TabIndex        =   48
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   5
                     Left            =   2240
                     TabIndex        =   47
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   5
                     Left            =   2240
                     TabIndex        =   46
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   4
                     Left            =   1780
                     TabIndex        =   45
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   4
                     Left            =   1780
                     TabIndex        =   44
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   3
                     Left            =   1340
                     TabIndex        =   43
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   3
                     Left            =   1340
                     TabIndex        =   42
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   2
                     Left            =   890
                     TabIndex        =   41
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   2
                     Left            =   890
                     TabIndex        =   40
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   1
                     Left            =   440
                     TabIndex        =   39
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   1
                     Left            =   440
                     TabIndex        =   38
                     Text            =   "0"
                     Top             =   -10
                     Width           =   460
                  End
                  Begin VB.TextBox txtNetPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   0
                     Left            =   -10
                     TabIndex        =   37
                     Text            =   "0"
                     Top             =   230
                     Width           =   460
                  End
                  Begin VB.TextBox txtGrossPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   0
                     Left            =   -10
                     TabIndex        =   36
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
                  TabIndex        =   103
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
                  TabIndex        =   102
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
                  TabIndex        =   101
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
               TabIndex        =   32
               Top             =   -10
               Width           =   12330
               Begin MSFlexGridLib.MSFlexGrid FGrid 
                  Height          =   2025
                  Left            =   -105
                  TabIndex        =   33
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
      Begin RPVGCC.b8Container b8Container4 
         Height          =   1335
         Left            =   5880
         TabIndex        =   104
         Top             =   2400
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   2355
         BackColor       =   49152
         Begin VB.PictureBox Picture7 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1095
            Left            =   120
            ScaleHeight     =   1095
            ScaleWidth      =   2295
            TabIndex        =   105
            Top             =   120
            Width           =   2295
            Begin VB.TextBox txtClass 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1080
               TabIndex        =   107
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtHandicap 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1080
               TabIndex        =   106
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Class"
               Height          =   255
               Left            =   120
               TabIndex        =   109
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Handicap"
               Height          =   255
               Left            =   120
               TabIndex        =   108
               Top             =   240
               Width           =   975
            End
         End
      End
      Begin RPVGCC.b8Container b8Container6 
         Height          =   1575
         Left            =   120
         TabIndex        =   138
         Top             =   2160
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2778
         BackColor       =   49152
         Begin VB.PictureBox Picture8 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   120
            ScaleHeight     =   1335
            ScaleWidth      =   5415
            TabIndex        =   139
            Top             =   120
            Width           =   5415
            Begin MSComctlLib.ListView lstTeamMates 
               Height          =   1335
               Left            =   0
               TabIndex        =   140
               Top             =   0
               Width           =   5415
               _ExtentX        =   9551
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
               NumItems        =   3
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "PlayerName"
                  Object.Width           =   7056
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   2
                  Text            =   "Gross Pts"
                  Object.Width           =   1764
               EndProperty
            End
         End
      End
   End
   Begin RPVGCC.b8Container picSearchAdd 
      Height          =   4695
      Left            =   4080
      TabIndex        =   110
      Top             =   840
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   8281
      BackColor       =   15396057
      Begin VB.TextBox txtDateAdd 
         Height          =   315
         Left            =   1800
         TabIndex        =   116
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelAdd 
         Height          =   480
         Left            =   2280
         Picture         =   "frmScoreCard.frx":3CC96
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   3960
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKAdd 
         Height          =   480
         Left            =   480
         Picture         =   "frmScoreCard.frx":3D3F2
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   3960
         Width           =   1560
      End
      Begin VB.ListBox lstResultAdd 
         Height          =   2595
         Left            =   120
         TabIndex        =   113
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtSearchAdd 
         Height          =   315
         Left            =   120
         TabIndex        =   112
         Top             =   480
         Width           =   4095
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   45
         TabIndex        =   111
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
         Icon            =   "frmScoreCard.frx":3DA64
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   255
         Left            =   1320
         TabIndex        =   117
         Top             =   3480
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmScoreCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TourNoOfPlays    As Double
Dim PlayerKey           As Double
Dim s                   As String
Dim rs                  As New ADODB.Recordset
Dim t                   As String
Dim rt                  As New ADODB.Recordset
Dim u                   As String
Dim ru                  As New ADODB.Recordset
Dim v                   As String
Dim rv                  As New ADODB.Recordset

Dim TRANSACTIONTYPE     As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim iScoreKey, dDateEnd

Private Function BROWSER(strCtrl, isAction As String)
Dim i, TeamTmp, x
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
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    LocationKey = rs!LocationKey
    txtCtrl.Text = rs!CtrlNo
    txtDate.Text = Format(rs!dDate, "mm/dd/yyyy")
    txtPlayer.Text = rs!PlayerName
    txtHandicap.Text = rs!Handicap
    txtClass.Text = rs!Class
    
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
    t = "SELECT TeamKey " & _
        " From tbl_Scoring_Team_Detail " & _
        " WHERE (PlayerKey = " & rs!PlayerKey & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        TeamTmp = rt!TeamKey
    End If
    rt.Close
    lstTeamMates.ListItems.Clear
    If CDbl(TeamTmp) > 0 Then
        t = "SELECT tbl_Scoring_Team_Detail.TeamKey, tbl_Scoring_Team_Detail.Line, tbl_Scoring_Team_Detail.PlayerKey, " & _
            " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, " & _
            " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.Class, " & _
            " IsNull((SELECT SUM(GrossPoints) AS GrossPoints " & _
            " From tbl_Scoring_ScoreCard " & _
            " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)), 0) AS GrossPts " & _
            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " Where (tbl_Scoring_Team_Detail.TeamKey = " & TeamTmp & ") And (tbl_Scoring_Team_Detail.PlayerKey <> " & rs!PlayerKey & ") " & _
            " ORDER BY ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.GrossPoints) AS GrossPoints " & _
            " From tbl_Scoring_ScoreCard " & _
            " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        While Not rt.EOF
            Set x = lstTeamMates.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = Trim(rt!LastName) & ",  " & Trim(rt!FirstName) & IIf(Trim(rt!MiddleName) = "", "", "  " & rt!MiddleName)
            x.SubItems(2) = rt!GrossPts
            rt.MoveNext
        Wend
        rt.Close
    End If
    
    StatusBar1.Panels(1).Text = rs!PK
    StatusBar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", "Last Modified : " & rs!LastModified)
    
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
    
    SaveSetting App.EXEName, "ScoreCardControl", "ScoreCardCtrl", rs!CtrlNo
    
End If
rs.Close
End Function


Private Function PRESS_INSERT()
Dim i
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
    PopupMenu MainFormPopupF.mnuScoringLocationAdd, , 200, 500
End If
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
ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard WHERE (PK = " & StatusBar1.Panels(1).Text & ")"
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
        " From tbl_Scoring_ScoreCard " & _
        " WHERE (TournamentKey = " & TournamentKey & ") " & _
        " AND (PlayerKey = " & PlayerKey & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    TourNoOfPlaysTmp = rs!NoofRec
    rs.Close
    
    If CDbl(TourNoOfPlaysTmp) + 1 > CDbl(DaysPlayerToPlay) Then MsgBox "Number of Plays Exceeded!                  ", vbCritical, "Error...": Exit Function
    
    strCtrlNo = "00000001"
    s = "SELECT TOP 1 CtrlNo " & _
        " FROM tbl_Scoring_ScoreCard " & _
        " WHERE (TournamentKey = " & TournamentKey & ") " & _
        " ORDER BY CtrlNo DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        strCtrlNo = Format(CDbl(rs!CtrlNo) + 1, "0000000#")
    End If
    rs.Close
    
    Do
        s = "SELECT tbl_Scoring_ScoreCard.* " & _
            " FROM tbl_Scoring_ScoreCard " & _
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
            
            ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_Detail " & _
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
    ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard " & _
                      " SET LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & SCardKey & ")"
    
    ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_Detail WHERE (ScoreCardKey = " & SCardKey & ")"
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
        
        ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_Detail " & _
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
picToolbar.Enabled = False
picMain.Enabled = False
picPrint.ZOrder 0
picPrint.Visible = True
cmbReportType.SetFocus
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
Dim i
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
txtLocation.Text = ""
lstTeamMates.ListItems.Clear
StatusBar1.Panels(1).Text = ""
StatusBar1.Panels(2).Text = ""
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
            .Buttons(19).Enabled = False
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
            .Buttons(19).Enabled = False
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


Private Sub cmbDivision_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKPrint_Click
End Sub

Private Sub cmbGroup_Click()
If cmbGroup.ListIndex = -1 Then Exit Sub
If cmbGroup.ListIndex = 0 Then
    cmbDivision.Clear
    cmbDivision.AddItem "ALL"
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
        cmbDivision.AddItem ru!Class
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
    cmbDivision.Clear
    cmbDivision.AddItem "ALL"
    u = "SELECT Class" & _
        " From tbl_Scoring_TournamentInfo_Class " & _
        " Where (TournamentKey = " & TournamentKey & ") " & _
        " ORDER BY Class"
    If ru.State = adStateOpen Then ru.Close
    ru.Open u, ConnOmega
    While Not ru.EOF
        cmbDivision.AddItem ru!Class
        ru.MoveNext
    Wend
    ru.Close
    
    cmbDay.Clear
    cmbDay.AddItem "1"
    cmbDay.AddItem "2"
End If
End Sub

Private Sub cmbGroup_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbDivision.SetFocus
End Sub

Private Sub cmbReportType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbGroup.SetFocus
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
Dim Array1, TourNoOfPlaysTmp, x, TeamTmp
If lstResultAdd.ListIndex = -1 Then Exit Sub
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


LOAD_CARD_LOCATION LocationKey, FormatDateTime(txtDateAdd.Text, vbShortDate), FGrid

CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
PlayerKey = lstResultAdd.ItemData(lstResultAdd.ListIndex)
txtPlayer.Text = lstResultAdd.List(lstResultAdd.ListIndex)
txtDate.Text = Format(FormatDateTime(txtDateAdd.Text, vbShortDate), "mm/dd/yyyy")


txtLocation.Text = ""
t = "SELECT tbl_Scoring_Location.* " & _
    " FROM tbl_Scoring_Location " & _
    " WHERE (PK = " & LocationKey & ")"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then
    txtLocation.Text = rt!ScoringLocation
End If
rt.Close


TeamTmp = 0
s = "SELECT TeamKey " & _
    " From tbl_Scoring_Team_Detail " & _
    " WHERE (PlayerKey = " & PlayerKey & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    TeamTmp = rs!TeamKey
End If
rs.Close

If CDbl(TeamTmp) > 0 Then
    's = "SELECT tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
        " tbl_Scoring_PlayerName.MiddleName, " & _
        " ISNULL((SELECT tbl_Scoring_ScoreCard.GrossPoints " & _
        " From tbl_Scoring_ScoreCard " & _
        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS GrossPts " & _
        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
        " Where (tbl_Scoring_Team_Detail.TeamKey = " & TeamTmp & ") " & _
        " And (tbl_Scoring_Team_Detail.PlayerKey <> " & PlayerKey & ") " & _
        " Order By ISNULL((SELECT tbl_Scoring_ScoreCard.GrossPoints " & _
        " From tbl_Scoring_ScoreCard " & _
        " WHERE (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) DESC"
    s = "SELECT tbl_Scoring_Team_Detail.TeamKey, tbl_Scoring_Team_Detail.Line, tbl_Scoring_Team_Detail.PlayerKey, " & _
        " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, " & _
        " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.Class, " & _
        " IsNull((SELECT SUM(GrossPoints) AS GrossPoints " & _
        " From tbl_Scoring_ScoreCard " & _
        " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)), 0) AS GrossPts " & _
        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
        " Where (tbl_Scoring_Team_Detail.TeamKey = " & TeamTmp & ") And (tbl_Scoring_Team_Detail.PlayerKey <> " & PlayerKey & ") " & _
        " ORDER BY ISNULL((SELECT SUM(tbl_Scoring_ScoreCard.GrossPoints) AS GrossPoints " & _
        " From tbl_Scoring_ScoreCard " & _
        " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)), 0) DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        Set x = lstTeamMates.ListItems.Add()
        x.Text = ""
        x.SubItems(1) = Trim(rs!LastName) & ",  " & Trim(rs!FirstName) & IIf(Trim(rs!MiddleName) = "", "", "  " & rs!MiddleName)
        x.SubItems(2) = rs!GrossPts
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
    txtClass.Text = rs!Class
End If
rs.Close
cmdCancelAdd_Click
txtGrossScore(0).SetFocus

End Sub

Private Sub cmdOKPrint_Click()
If cmbReportType.ListIndex = -1 Then Exit Sub
If cmbReportType.ListIndex = 1 Then
    If cmbGroup.ListIndex = -1 Then Exit Sub
    If cmbDivision.ListIndex = -1 Then Exit Sub
End If
If cmbDay.ListIndex = -1 Then Exit Sub
If cmbReportType.ListIndex = 1 Then
    If cmbGroup.ListIndex = 0 Then
        If cmbDay.ListIndex = 0 Then TimerTeamAll.Enabled = True: Exit Sub
    End If
End If

Dim i, j, strClass, TableName, DetailTableName, _
Columns, ColumnsDet, Clustered, sMasterFields, _
sDetailFields, Arr, Arr1, Arr2, MasterKey, dblGrossPts, _
Filename, cnt, HeaderRow, dblCntBackPlayer, _
dblEagle, dblTotalHDCP, strPlayerName, sGrossPts

Dim WorkbookName    As String
Dim ColTop, RowTop, ColCount, RowCount, strRange, strRange1, _
ColCountDet, RowCountDet, RowFrom, RowTo

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

Select Case cmbReportType.ListIndex
    Case 0  'Result
    
    Case 1  'Summary
        
        Select Case cmbGroup.ListIndex
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
                If Trim(cmbDivision.List(cmbDivision.ListIndex)) = "ALL" Then
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
                            " AND (tbl_Scoring_TournamentInfo_Class.HTo >= tbl_Scoring_Team.TeamHDCP)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')"
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
                            " AND (tbl_Scoring_TournamentInfo_Index.HTo >= tbl_Scoring_Team.TeamIndex)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')"
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
                    .Workbooks(1).Sheets(1).Name = cmbReportType.List(cmbReportType.ListIndex) & " (" & Replace(cmbDivision.List(cmbDivision.ListIndex), "SORT BY ", "") & ")" '"Report"
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
                                        
                                        
                        If Trim(cmbDivision.List(cmbDivision.ListIndex)) <> "ALL" Then
                            ColCount = 0
                            RowCount = RowCount + 1
                            ColCount = ColCount + 1
                            strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                            strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 3) > 26, 64 + 1, 64) + rs.Fields.Count - 3)) & CStr(RowCount)
                            .Range(strRange, strRange1).Select
                            xlsApp.Selection.Merge
                            
                            strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                            .Range(strRange).Value = "CLASS " & Trim(cmbDivision.List(cmbDivision.ListIndex))
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
                
                If Trim(cmbDivision.List(cmbDivision.ListIndex)) = "ALL" Then
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
                        " (tbl_Scoring_TournamentInfo_Class.HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')"
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
                    .Workbooks(1).Sheets(1).Name = cmbReportType.List(cmbReportType.ListIndex) & " (" & cmbGroup.List(cmbGroup.ListIndex) & ") (" & Replace(cmbDivision.List(cmbDivision.ListIndex), "SORT BY ", "") & ")"    '"Report"
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
                                        
                        If Trim(cmbDivision.List(cmbDivision.ListIndex)) <> "ALL" Then
                            ColCount = 0
                            RowCount = RowCount + 1
                            ColCount = ColCount + 1
                            strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                            strRange1 = (Chr$(IIf(CDbl(rs.Fields.Count - 3) > 26, 64 + 1, 64) + rs.Fields.Count - 2)) & CStr(RowCount)
                            .Range(strRange, strRange1).Select
                            xlsApp.Selection.Merge
                            
                            strRange = (Chr$(IIf(CDbl(ColCount) > 26, 64 + 1, 64) + ColCount)) & CStr(RowCount)
                            .Range(strRange).Value = "CLASS " & Trim(cmbDivision.List(cmbDivision.ListIndex))
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
ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard WHERE (TournamentKey = 7) AND (DDate = '" & FormatDateTime("9/25/2011", vbShortDate) & "')"
Screen.MousePointer = vbDefault
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
Me.Height = 7170 '6825
Me.Width = 12195
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

With cmbReportType
    .Clear
    .AddItem "RESULT"
    .AddItem "SUMMARY"
'    .AddItem "SCORES"
End With

With cmbGroup
    .Clear
    .AddItem "TEAM"
    .AddItem "INDIVIDUAL"
End With

'cmbDivision.Clear
'cmbDivision.AddItem "ALL"
'If CDbl(TeamAverage) = 2 Then
'    s = "SELECT Class" & _
'        " From tbl_Scoring_TournamentInfo_Index " & _
'        " Where (TournamentKey = " & TournamentKey & ") " & _
'        " ORDER BY Class"
'Else
'    s = "SELECT Class" & _
'        " From tbl_Scoring_TournamentInfo_Class " & _
'        " Where (TournamentKey = " & TournamentKey & ") " & _
'        " ORDER BY Class"
'End If
'If rs.State = adStateOpen Then rs.Close
'rs.Open s, ConnOmega
'While Not rs.EOF
'    cmbDivision.AddItem rs!Class
'    rs.MoveNext
'Wend
'rs.Close

Dim i

cmbDay.Clear
For i = 1 To DaysPlayerToPlay
    cmbDay.AddItem i
Next i

'Me.Caption = "Score Card"
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
'LocationKey
LOAD_CARD_LOCATION LocationKey, dDateEnd, FGrid
'LOAD_CARD
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH

BROWSER GetSetting(App.EXEName, "ScoreCardControl", "ScoreCardCtrl", ""), "is_LOAD"
If Trim(txtPlayer.Text) = "" Then BROWSER GetSetting(App.EXEName, "ScoreCardControl", "ScoreCardCtrl", ""), "is_HOME"

Dim tmp As Long
tmp = SetWindowLong(txtSearchAdd.hwnd, GWL_STYLE, GetWindowLong(txtSearchAdd.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picPrint.Visible = True Then Cancel = -1
If picSearchAdd.Visible = True Then Cancel = -1
If picSearch.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstResult_Click()
If lstResult.ListIndex = -1 Then cmbDate.Clear: Exit Sub
cmbDate.Clear
s = "SELECT PK, DDate " & _
    " From tbl_Scoring_ScoreCard " & _
    " Where (TournamentKey = " & TournamentKey & ") " & _
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
Dim strCtrlNo, Arr, Arr1, sFileNameMaster, sFileNameDetail, StrFile, sFileArrDet, sFileArr
Dim cn As ADODB.Connection
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
    If Trim(txtPlayer.Text) = "" Then BROWSER GetSetting(App.EXEName, "ScoreCardControl", "ScoreCardCtrl", ""), "is_HOME"
    
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
    If Trim(txtPlayer.Text) = "" Then BROWSER GetSetting(App.EXEName, "ScoreCardControl", "ScoreCardCtrl", ""), "is_HOME"
    
    Screen.MousePointer = vbDefault
    If cn.State = adStateOpen Then cn.Close
End If

Exit Sub
PG:
Screen.MousePointer = vbDefault
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub TimerTeamAll_Timer()
TimerTeamAll.Enabled = False
Dim i, j, x, strClass, TableName, DetailTableName, _
Columns, ColumnsDet, Clustered, sMasterFields, _
sDetailFields, Arr, Arr1, Arr2, MasterKey, dblGrossPts, _
Filename, cnt, HeaderRow, dblCntBackPlayer1, dblCntBackPlayer2, _
dblEagle, dblTotalHDCP, strPlayerName, dblGrossPts1, _
dblGrossPts2, dblGrossPtsTot, dblCntBackPlayerTot, sFreezePane

Dim WorkbookName    As String
Dim ColTop, RowTop, ColCount, RowCount, strRange, strRange1, _
ColCountDet, RowCountDet, RowFrom, RowTo

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
If Trim(cmbDivision.List(cmbDivision.ListIndex)) = "ALL" Then
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
            " AND (tbl_Scoring_TournamentInfo_Index.HTo >= tbl_Scoring_Team.TeamIndex)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')"
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
            " AND (tbl_Scoring_TournamentInfo_Class.HTo >= tbl_Scoring_Team.TeamHDCP)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')"
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
    .Workbooks(1).Sheets(1).Name = cmbReportType.List(cmbReportType.ListIndex) & " (" & Replace(cmbDivision.List(cmbDivision.ListIndex), "SORT BY ", "") & ")" '"Report"
    If .Workbooks(1).Sheets.Count = 3 Then
        .Workbooks(1).Sheets(2).Delete
        .Workbooks(1).Sheets(2).Delete
    End If
    
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
                        
                        
        If Trim(cmbDivision.List(cmbDivision.ListIndex)) <> "ALL" Then
            ColCount = 0
            RowCount = RowCount + 1
            ColCount = ColCount + 1
            strRange = EXCEL_RANGE(ColCount, RowCount)
            strRange1 = EXCEL_RANGE(rs.Fields.Count - 3, RowCount)
            .Range(strRange, strRange1).Select
            xlsApp.Selection.Merge
            
            strRange = EXCEL_RANGE(ColCount, RowCount)
            .Range(strRange).Value = "CLASS " & Trim(cmbDivision.List(cmbDivision.ListIndex))
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
txtSGrossB.Text = RETURNTEXTVALUE(txtGrossPtsB)
End Sub

Private Sub txtGrossPtsF_Change()
txtGrossPtsTot.Text = RETURNTEXTVALUE(txtGrossPtsF) + _
                      RETURNTEXTVALUE(txtGrossPtsB)
txtSGrossF.Text = RETURNTEXTVALUE(txtGrossPtsF)
End Sub

Private Sub txtGrossScore_Change(Index As Integer)

If TRANSACTIONTYPE = is_REFRESH Then Exit Sub

If RETURNTEXTVALUE(txtGrossScore(Index)) <= 0 Then txtGrossPts(Index).Text = "0": txtNetPts(Index).Text = "0": Exit Sub

Dim dblPar, dblHandicap
With FGrid
    Select Case Index
        Case 0
            dblPar = .TextMatrix(1, 2)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 2)
            txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 1
            dblPar = .TextMatrix(1, 3)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 3)
            txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 2
            dblPar = .TextMatrix(1, 4)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 4)
            txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 3
            dblPar = .TextMatrix(1, 5)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 5)
            txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 4
            dblPar = .TextMatrix(1, 6)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 6)
            txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 5
            dblPar = .TextMatrix(1, 7)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 7)
            txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 6
            dblPar = .TextMatrix(1, 8)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 8)
            txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 7
            dblPar = .TextMatrix(1, 9)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 9)
            txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 8
            dblPar = .TextMatrix(1, 10)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 10)
            txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 9
            dblPar = .TextMatrix(1, 12)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 12)
            txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 10
            dblPar = .TextMatrix(1, 13)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 13)
            txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 11
            dblPar = .TextMatrix(1, 14)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 14)
            txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 12
            dblPar = .TextMatrix(1, 15)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 15)
            txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 13
            dblPar = .TextMatrix(1, 16)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 16)
            txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 14
            dblPar = .TextMatrix(1, 17)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 17)
            txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 15
            dblPar = .TextMatrix(1, 18)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 18)
            txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 16
            dblPar = .TextMatrix(1, 19)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 19)
            txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
        Case 17
            dblPar = .TextMatrix(1, 20)
            txtGrossPts(Index).Text = IIf(Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))) <= 0, 0, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
            dblHandicap = .TextMatrix(2, 20)
            txtNetPts(Index).Text = Get_Net_Points(RETURNTEXTVALUE(txtHandicap), dblHandicap, Get_Gross_Points(dblPar, RETURNTEXTVALUE(txtGrossScore(Index))))
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

End Sub

Private Sub txtGrossScoreF_Change()
txtGrossScoreTot.Text = RETURNTEXTVALUE(txtGrossScoreF) + _
                        RETURNTEXTVALUE(txtGrossScoreB)
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
txtSNetB.Text = RETURNTEXTVALUE(txtNetPtsB)
End Sub

Private Sub txtNetPtsF_Change()
txtNetPtsTot.Text = RETURNTEXTVALUE(txtNetPtsF) + _
                    RETURNTEXTVALUE(txtNetPtsB)
txtSNetF.Text = RETURNTEXTVALUE(txtNetPtsF)
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstResult.Clear: cmbDate.Clear: Exit Sub
lstResult.Clear: cmbDate.Clear
's = "SELECT tbl_Scoring_PlayerName.PK, " & _
    " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName " & _
    " FROM tbl_Scoring_ScoreCard LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") " & _
    " AND (tbl_Scoring_PlayerName.LastName LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
    " ORDER BY tbl_Scoring_ScoreCard.TournamentKey"
s = "SELECT tbl_Scoring_PlayerName.PK, " & _
    " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName " & _
    " FROM tbl_Scoring_ScoreCard LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_ScoreCard.TournamentKey = " & TournamentKey & ") " & _
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

Private Sub txtSearch_GotFocus()
HTEXT txtSearch
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResult.SetFocus
End Sub

Private Sub txtSearchAdd_Change()
If Trim(txtSearchAdd.Text) = "" Then lstResultAdd.Clear: Exit Sub
Dim s As String
Dim rs As New ADODB.Recordset
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

Private Sub txtSGrossB_Change()
txtSGrossTot.Text = RETURNTEXTVALUE(txtSGrossF) + _
                    RETURNTEXTVALUE(txtSGrossB)
End Sub

Private Sub txtSGrossF_Change()
txtSGrossTot.Text = RETURNTEXTVALUE(txtSGrossF) + _
                    RETURNTEXTVALUE(txtSGrossB)
End Sub

Private Sub txtSNetB_Change()
txtSNetTot.Text = RETURNTEXTVALUE(txtSNetF) + _
                  RETURNTEXTVALUE(txtSNetB)
End Sub

Private Sub txtSNetF_Change()
txtSNetTot.Text = RETURNTEXTVALUE(txtSNetF) + _
                  RETURNTEXTVALUE(txtSNetB)
End Sub
