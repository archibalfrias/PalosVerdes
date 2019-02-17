VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmScoreCardModifiedMolave 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12075
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmScoreCardModifiedMolave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   12075
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picPrint 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   3960
      ScaleHeight     =   3015
      ScaleWidth      =   4335
      TabIndex        =   136
      Top             =   1920
      Visible         =   0   'False
      Width           =   4335
      Begin RPVGCC.b8Container picPrint1 
         Height          =   3015
         Left            =   0
         TabIndex        =   137
         Top             =   0
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   5318
         BackColor       =   15396057
         Begin VB.TextBox txtDatePrint 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1440
            TabIndex        =   149
            Top             =   1920
            Width           =   1815
         End
         Begin VB.ComboBox cmbGender 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   144
            Top             =   840
            Width           =   4095
         End
         Begin VB.Timer TimerReport 
            Enabled         =   0   'False
            Interval        =   200
            Left            =   120
            Top             =   2400
         End
         Begin VB.TextBox txtTop 
            Height          =   315
            Left            =   3240
            TabIndex        =   143
            Top             =   1560
            Width           =   975
         End
         Begin VB.CommandButton cmdOKPrint 
            Height          =   480
            Left            =   555
            Picture         =   "frmScoreCardModifiedMolave.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   142
            Top             =   2355
            Width           =   1560
         End
         Begin VB.CommandButton cmdCancelPrint 
            Height          =   480
            Left            =   2235
            Picture         =   "frmScoreCardModifiedMolave.frx":0F3C
            Style           =   1  'Graphical
            TabIndex        =   141
            Top             =   2355
            Width           =   1560
         End
         Begin VB.ComboBox cmbGroup 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   140
            Top             =   1200
            Width           =   4095
         End
         Begin VB.ComboBox cmbDivision 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   139
            Top             =   1560
            Width           =   2535
         End
         Begin VB.ComboBox cmbReportType 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   138
            Top             =   480
            Width           =   4095
         End
         Begin RPVGCC.b8TitleBar b8TitleBar3 
            Height          =   345
            Left            =   40
            TabIndex        =   145
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
            Icon            =   "frmScoreCardModifiedMolave.frx":1698
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   255
            Left            =   960
            TabIndex        =   150
            Top             =   1920
            Width           =   495
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Top"
            Height          =   255
            Left            =   2760
            TabIndex        =   146
            Top             =   1560
            Width           =   495
         End
      End
   End
   Begin RPVGCC.b8Container picProgress 
      Height          =   975
      Left            =   3240
      TabIndex        =   116
      Top             =   2640
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
         TabIndex        =   117
         Top             =   120
         Width           =   5415
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11280
      Top             =   0
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
            Picture         =   "frmScoreCardModifiedMolave.frx":1C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardModifiedMolave.frx":1D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardModifiedMolave.frx":1EB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardModifiedMolave.frx":21D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardModifiedMolave.frx":258B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardModifiedMolave.frx":29DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardModifiedMolave.frx":2E2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardModifiedMolave.frx":31E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardModifiedMolave.frx":32F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardModifiedMolave.frx":383B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardModifiedMolave.frx":3995
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardModifiedMolave.frx":3ED7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   7560
      Width           =   12075
      _ExtentX        =   21299
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
   Begin RPVGCC.b8Container b8Container7 
      Height          =   975
      Left            =   0
      TabIndex        =   134
      Top             =   0
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1720
      BackColor       =   13023396
      Begin VB.PictureBox Picture9 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5235
         TabIndex        =   135
         Top             =   120
         Width           =   5295
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   0
      Picture         =   "frmScoreCardModifiedMolave.frx":40FB
      ScaleHeight     =   6375
      ScaleWidth      =   12135
      TabIndex        =   1
      Top             =   0
      Width           =   12135
      Begin VB.PictureBox picToolbar 
         BorderStyle     =   0  'None
         Height          =   770
         Left            =   0
         ScaleHeight     =   765
         ScaleWidth      =   15000
         TabIndex        =   3
         Top             =   0
         Width           =   15000
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   570
            Left            =   0
            TabIndex        =   4
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
            Begin VB.PictureBox Picture10 
               BorderStyle     =   0  'None
               Height          =   375
               Left            =   7320
               ScaleHeight     =   375
               ScaleWidth      =   4695
               TabIndex        =   147
               Top             =   120
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
                  TabIndex        =   148
                  Text            =   "Text1"
                  Top             =   0
                  Width           =   4095
               End
            End
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
         Left            =   10920
         TabIndex        =   2
         Top             =   3480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin RPVGCC.b8Container b8Container3 
         Height          =   1095
         Left            =   5880
         TabIndex        =   5
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
            TabIndex        =   6
            Top             =   120
            Width           =   5895
            Begin VB.TextBox txtTournament 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1095
               TabIndex        =   8
               Top             =   120
               Width           =   4695
            End
            Begin VB.TextBox txtTourDate 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1095
               TabIndex        =   7
               Text            =   "06/01/2010 - 06/04/2010"
               Top             =   480
               Width           =   4695
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Tournament"
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   120
               Width           =   1335
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
         End
      End
      Begin RPVGCC.b8Container b8Container5 
         Height          =   1095
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1931
         BackColor       =   49152
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   855
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   5415
            TabIndex        =   12
            Top             =   120
            Width           =   5415
            Begin VB.TextBox txtDate 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   720
               TabIndex        =   15
               Top             =   480
               Width           =   1575
            End
            Begin VB.TextBox txtDay 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   4560
               TabIndex        =   14
               Top             =   480
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox txtPlayer 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   720
               TabIndex        =   13
               Top             =   120
               Width           =   4575
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   480
               Width           =   495
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Day"
               Height          =   255
               Left            =   3720
               TabIndex        =   17
               Top             =   480
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Player"
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   120
               Width           =   975
            End
         End
      End
      Begin RPVGCC.b8Container b8Container2 
         Height          =   1335
         Left            =   2280
         TabIndex        =   19
         Top             =   2040
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
            TabIndex        =   20
            Top             =   120
            Width           =   3255
            Begin VB.TextBox txtSGrossF 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1080
               TabIndex        =   26
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox txtSNetF 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1080
               TabIndex        =   25
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtSGrossB 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1680
               TabIndex        =   24
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox txtSNetB 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1680
               TabIndex        =   23
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtSGrossTot 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   2400
               TabIndex        =   22
               Top             =   360
               Width           =   735
            End
            Begin VB.TextBox txtSNetTot 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   2400
               TabIndex        =   21
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Gross Points"
               Height          =   255
               Left            =   120
               TabIndex        =   32
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Net Points"
               Height          =   255
               Left            =   120
               TabIndex        =   31
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
               TabIndex        =   30
               Top             =   0
               Width           =   975
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "F - 9"
               Height          =   255
               Left            =   1080
               TabIndex        =   29
               Top             =   120
               Width           =   495
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "B - 9"
               Height          =   255
               Left            =   1680
               TabIndex        =   28
               Top             =   120
               Width           =   495
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Total"
               Height          =   255
               Left            =   2400
               TabIndex        =   27
               Top             =   120
               Width           =   735
            End
         End
      End
      Begin RPVGCC.b8Container b8Container1 
         Height          =   2490
         Left            =   120
         TabIndex        =   33
         Top             =   3720
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
            TabIndex        =   34
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
               TabIndex        =   37
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
                  TabIndex        =   101
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
                     TabIndex        =   103
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
                     TabIndex        =   102
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
                  TabIndex        =   99
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
                     TabIndex        =   100
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
                  TabIndex        =   98
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
                  TabIndex        =   97
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
                  TabIndex        =   96
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
                  Index           =   13
                  Left            =   8380
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
                  Index           =   14
                  Left            =   8830
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
                  Index           =   15
                  Left            =   9280
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
                  Index           =   16
                  Left            =   9730
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
                  Index           =   17
                  Left            =   10180
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
                  Index           =   8
                  Left            =   5580
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
                  Index           =   7
                  Left            =   5130
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
                  Index           =   6
                  Left            =   4680
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
                  Index           =   5
                  Left            =   4230
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
                  Index           =   4
                  Left            =   3780
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
                  Index           =   3
                  Left            =   3330
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
                  Index           =   2
                  Left            =   2880
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
                  Index           =   1
                  Left            =   2430
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
                  Index           =   0
                  Left            =   1980
                  TabIndex        =   81
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
                  TabIndex        =   38
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
                     TabIndex        =   80
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
                     TabIndex        =   79
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
                     TabIndex        =   78
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
                     TabIndex        =   77
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
                     TabIndex        =   76
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
                     TabIndex        =   75
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
                     Index           =   17
                     Left            =   8190
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
                     Index           =   16
                     Left            =   7740
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
                     Index           =   16
                     Left            =   7740
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
                     Index           =   15
                     Left            =   7290
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
                     Index           =   15
                     Left            =   7290
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
                     Index           =   14
                     Left            =   6840
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
                     Index           =   13
                     Left            =   6390
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
                     Index           =   13
                     Left            =   6390
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
                     Index           =   12
                     Left            =   5940
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
                     Index           =   12
                     Left            =   5940
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
                     Index           =   11
                     Left            =   5490
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
                     Index           =   11
                     Left            =   5490
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
                     Index           =   10
                     Left            =   5040
                     TabIndex        =   60
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
                     TabIndex        =   59
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
                     TabIndex        =   58
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
                     TabIndex        =   57
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
                     TabIndex        =   56
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
                     TabIndex        =   55
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
                     TabIndex        =   54
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
                     TabIndex        =   53
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
                     TabIndex        =   52
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
                     TabIndex        =   51
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
                     TabIndex        =   50
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
                     TabIndex        =   49
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
                     Index           =   4
                     Left            =   1780
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
                     Index           =   3
                     Left            =   1340
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
                     Index           =   3
                     Left            =   1340
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
                     Index           =   2
                     Left            =   890
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
                     Index           =   2
                     Left            =   890
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
                     Index           =   1
                     Left            =   440
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
                     Index           =   1
                     Left            =   440
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
                     Index           =   0
                     Left            =   -10
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
                     Index           =   0
                     Left            =   -10
                     TabIndex        =   39
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
                  TabIndex        =   106
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
                  TabIndex        =   105
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
                  TabIndex        =   104
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
               TabIndex        =   35
               Top             =   -10
               Width           =   12330
               Begin MSFlexGridLib.MSFlexGrid FGrid 
                  Height          =   2025
                  Left            =   -105
                  TabIndex        =   36
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
         Left            =   120
         TabIndex        =   107
         Top             =   2040
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   2355
         BackColor       =   49152
         Begin VB.PictureBox Picture7 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1095
            Left            =   120
            ScaleHeight     =   1095
            ScaleWidth      =   1815
            TabIndex        =   108
            Top             =   120
            Width           =   1815
            Begin VB.TextBox txtClass 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   960
               TabIndex        =   110
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox txtHandicap 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   960
               TabIndex        =   109
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Class"
               Height          =   255
               Left            =   120
               TabIndex        =   112
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Handicap"
               Height          =   255
               Left            =   120
               TabIndex        =   111
               Top             =   240
               Width           =   975
            End
         End
      End
      Begin RPVGCC.b8Container b8Container6 
         Height          =   1575
         Left            =   5880
         TabIndex        =   113
         Top             =   2040
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2778
         BackColor       =   49152
         Begin VB.PictureBox Picture8 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   120
            ScaleHeight     =   1335
            ScaleWidth      =   5895
            TabIndex        =   114
            Top             =   120
            Width           =   5895
            Begin MSComctlLib.ListView lstTeamMates 
               Height          =   1335
               Left            =   0
               TabIndex        =   115
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
      Left            =   3960
      TabIndex        =   118
      Top             =   720
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   8281
      BackColor       =   15396057
      Begin VB.TextBox txtDateAdd 
         Height          =   315
         Left            =   1800
         TabIndex        =   123
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelAdd 
         Height          =   480
         Left            =   2280
         Picture         =   "frmScoreCardModifiedMolave.frx":3B92E
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   3960
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKAdd 
         Height          =   480
         Left            =   480
         Picture         =   "frmScoreCardModifiedMolave.frx":3C08A
         Style           =   1  'Graphical
         TabIndex        =   121
         Top             =   3960
         Width           =   1560
      End
      Begin VB.ListBox lstResultAdd 
         Height          =   2595
         Left            =   120
         TabIndex        =   120
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtSearchAdd 
         Height          =   315
         Left            =   120
         TabIndex        =   119
         Top             =   480
         Width           =   4095
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
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
         Icon            =   "frmScoreCardModifiedMolave.frx":3C6FC
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   255
         Left            =   1320
         TabIndex        =   125
         Top             =   3480
         Width           =   495
      End
   End
   Begin RPVGCC.b8Container picSearch 
      Height          =   4695
      Left            =   3960
      TabIndex        =   126
      Top             =   720
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   8281
      BackColor       =   15396057
      Begin VB.ComboBox cmbDate 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   131
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   130
         Top             =   480
         Width           =   4095
      End
      Begin VB.ListBox lstResult 
         Height          =   2595
         Left            =   120
         TabIndex        =   129
         Top             =   840
         Width           =   4095
      End
      Begin VB.CommandButton cmdOKSearch 
         Height          =   480
         Left            =   480
         Picture         =   "frmScoreCardModifiedMolave.frx":3CC96
         Style           =   1  'Graphical
         TabIndex        =   128
         Top             =   3960
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelSearch 
         Height          =   480
         Left            =   2280
         Picture         =   "frmScoreCardModifiedMolave.frx":3D308
         Style           =   1  'Graphical
         TabIndex        =   127
         Top             =   3960
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   132
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
         Icon            =   "frmScoreCardModifiedMolave.frx":3DA64
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   255
         Left            =   1200
         TabIndex        =   133
         Top             =   3480
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmScoreCardModifiedMolave"
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
Dim sFileName           As String

Dim Arr, dDateEnd

Dim TRANSACTIONTYPE     As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Private Function BROWSER(strCtrl, isAction As String)
Dim i, TeamTmp, x
Select Case isAction
    Case "is_LOAD"
        If strCtrl <> "" Then
            s = "SELECT TOP 1 tbl_Scoring_ScoreCard_ModMolave.PK, tbl_Scoring_ScoreCard_ModMolave.CtrlNo, " & _
                " tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
                " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.Class, tbl_Scoring_ScoreCard_ModMolave.DDate, tbl_Scoring_ScoreCard_ModMolave.Score, " & _
                " tbl_Scoring_ScoreCard_ModMolave.Front9Gross, tbl_Scoring_ScoreCard_ModMolave.Back9Gross, tbl_Scoring_ScoreCard_ModMolave.GrossPoints, " & _
                " tbl_Scoring_ScoreCard_ModMolave.Front9Net, tbl_Scoring_ScoreCard_ModMolave.Back9Net, tbl_Scoring_ScoreCard_ModMolave.NetPoints, " & _
                " tbl_Scoring_ScoreCard_ModMolave.LastModified, tbl_Scoring_ScoreCard_ModMolave.Front9Score, " & _
                " tbl_Scoring_ScoreCard_ModMolave.Back9Score " & _
                " FROM tbl_Scoring_ScoreCard_ModMolave LEFT OUTER JOIN " & _
                " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                " WHERE (tbl_Scoring_ScoreCard_ModMolave.TournamentKey = " & TournamentKey & ") " & _
                " AND (tbl_Scoring_ScoreCard_ModMolave.CtrlNo = '" & strCtrl & "') " & _
                " ORDER BY tbl_Scoring_ScoreCard_ModMolave.CtrlNo"
        Else
            s = "SELECT TOP 1 tbl_Scoring_ScoreCard_ModMolave.PK, tbl_Scoring_ScoreCard_ModMolave.CtrlNo, " & _
                " tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
                " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.Class, tbl_Scoring_ScoreCard_ModMolave.DDate, tbl_Scoring_ScoreCard_ModMolave.Score, " & _
                " tbl_Scoring_ScoreCard_ModMolave.Front9Gross, tbl_Scoring_ScoreCard_ModMolave.Back9Gross, tbl_Scoring_ScoreCard_ModMolave.GrossPoints, " & _
                " tbl_Scoring_ScoreCard_ModMolave.Front9Net, tbl_Scoring_ScoreCard_ModMolave.Back9Net, tbl_Scoring_ScoreCard_ModMolave.NetPoints, " & _
                " tbl_Scoring_ScoreCard_ModMolave.LastModified, tbl_Scoring_ScoreCard_ModMolave.Front9Score, " & _
                " tbl_Scoring_ScoreCard_ModMolave.Back9Score " & _
                " FROM tbl_Scoring_ScoreCard_ModMolave LEFT OUTER JOIN " & _
                " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                " WHERE (tbl_Scoring_ScoreCard_ModMolave.TournamentKey = " & TournamentKey & ") " & _
                " ORDER BY tbl_Scoring_ScoreCard_ModMolave.CtrlNo"
        End If
    Case "is_FIND"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Scoring_ScoreCard_ModMolave.PK, tbl_Scoring_ScoreCard_ModMolave.CtrlNo, " & _
            " tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
            " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.Class, tbl_Scoring_ScoreCard_ModMolave.DDate, tbl_Scoring_ScoreCard_ModMolave.Score, " & _
            " tbl_Scoring_ScoreCard_ModMolave.Front9Gross, tbl_Scoring_ScoreCard_ModMolave.Back9Gross, tbl_Scoring_ScoreCard_ModMolave.GrossPoints, " & _
            " tbl_Scoring_ScoreCard_ModMolave.Front9Net, tbl_Scoring_ScoreCard_ModMolave.Back9Net, tbl_Scoring_ScoreCard_ModMolave.NetPoints, " & _
            " tbl_Scoring_ScoreCard_ModMolave.LastModified, tbl_Scoring_ScoreCard_ModMolave.Front9Score, " & _
            " tbl_Scoring_ScoreCard_ModMolave.Back9Score " & _
            " FROM tbl_Scoring_ScoreCard_ModMolave LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " WHERE (tbl_Scoring_ScoreCard_ModMolave.PK = " & strCtrl & ") " & _
            " ORDER BY tbl_Scoring_ScoreCard_ModMolave.CtrlNo DESC"
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Scoring_ScoreCard_ModMolave.PK, tbl_Scoring_ScoreCard_ModMolave.CtrlNo, " & _
            " tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
            " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.Class, tbl_Scoring_ScoreCard_ModMolave.DDate, tbl_Scoring_ScoreCard_ModMolave.Score, " & _
            " tbl_Scoring_ScoreCard_ModMolave.Front9Gross, tbl_Scoring_ScoreCard_ModMolave.Back9Gross, tbl_Scoring_ScoreCard_ModMolave.GrossPoints, " & _
            " tbl_Scoring_ScoreCard_ModMolave.Front9Net, tbl_Scoring_ScoreCard_ModMolave.Back9Net, tbl_Scoring_ScoreCard_ModMolave.NetPoints, " & _
            " tbl_Scoring_ScoreCard_ModMolave.LastModified, tbl_Scoring_ScoreCard_ModMolave.Front9Score, " & _
            " tbl_Scoring_ScoreCard_ModMolave.Back9Score " & _
            " FROM tbl_Scoring_ScoreCard_ModMolave LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " WHERE (tbl_Scoring_ScoreCard_ModMolave.TournamentKey = " & TournamentKey & ") " & _
            " ORDER BY tbl_Scoring_ScoreCard_ModMolave.CtrlNo"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Scoring_ScoreCard_ModMolave.PK, tbl_Scoring_ScoreCard_ModMolave.CtrlNo, " & _
            " tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
            " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.Class, tbl_Scoring_ScoreCard_ModMolave.DDate, tbl_Scoring_ScoreCard_ModMolave.Score, " & _
            " tbl_Scoring_ScoreCard_ModMolave.Front9Gross, tbl_Scoring_ScoreCard_ModMolave.Back9Gross, tbl_Scoring_ScoreCard_ModMolave.GrossPoints, " & _
            " tbl_Scoring_ScoreCard_ModMolave.Front9Net, tbl_Scoring_ScoreCard_ModMolave.Back9Net, tbl_Scoring_ScoreCard_ModMolave.NetPoints, " & _
            " tbl_Scoring_ScoreCard_ModMolave.LastModified, tbl_Scoring_ScoreCard_ModMolave.Front9Score, " & _
            " tbl_Scoring_ScoreCard_ModMolave.Back9Score " & _
            " FROM tbl_Scoring_ScoreCard_ModMolave LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " WHERE (tbl_Scoring_ScoreCard_ModMolave.TournamentKey = " & TournamentKey & ") " & _
            " AND (tbl_Scoring_ScoreCard_ModMolave.CtrlNo < '" & strCtrl & "') " & _
            " ORDER BY tbl_Scoring_ScoreCard_ModMolave.CtrlNo DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Scoring_ScoreCard_ModMolave.PK, tbl_Scoring_ScoreCard_ModMolave.CtrlNo, " & _
            " tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
            " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.Class, tbl_Scoring_ScoreCard_ModMolave.DDate, tbl_Scoring_ScoreCard_ModMolave.Score, " & _
            " tbl_Scoring_ScoreCard_ModMolave.Front9Gross, tbl_Scoring_ScoreCard_ModMolave.Back9Gross, tbl_Scoring_ScoreCard_ModMolave.GrossPoints, " & _
            " tbl_Scoring_ScoreCard_ModMolave.Front9Net, tbl_Scoring_ScoreCard_ModMolave.Back9Net, tbl_Scoring_ScoreCard_ModMolave.NetPoints, " & _
            " tbl_Scoring_ScoreCard_ModMolave.LastModified, tbl_Scoring_ScoreCard_ModMolave.Front9Score, " & _
            " tbl_Scoring_ScoreCard_ModMolave.Back9Score " & _
            " FROM tbl_Scoring_ScoreCard_ModMolave LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " WHERE (tbl_Scoring_ScoreCard_ModMolave.TournamentKey = " & TournamentKey & ") " & _
            " AND (tbl_Scoring_ScoreCard_ModMolave.CtrlNo > '" & strCtrl & "') " & _
            " ORDER BY tbl_Scoring_ScoreCard_ModMolave.CtrlNo "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Scoring_ScoreCard_ModMolave.PK, tbl_Scoring_ScoreCard_ModMolave.CtrlNo, " & _
            " tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
            " tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_PlayerName.Class, tbl_Scoring_ScoreCard_ModMolave.DDate, tbl_Scoring_ScoreCard_ModMolave.Score, " & _
            " tbl_Scoring_ScoreCard_ModMolave.Front9Gross, tbl_Scoring_ScoreCard_ModMolave.Back9Gross, tbl_Scoring_ScoreCard_ModMolave.GrossPoints, " & _
            " tbl_Scoring_ScoreCard_ModMolave.Front9Net, tbl_Scoring_ScoreCard_ModMolave.Back9Net, tbl_Scoring_ScoreCard_ModMolave.NetPoints, " & _
            " tbl_Scoring_ScoreCard_ModMolave.LastModified, tbl_Scoring_ScoreCard_ModMolave.Front9Score, " & _
            " tbl_Scoring_ScoreCard_ModMolave.Back9Score " & _
            " FROM tbl_Scoring_ScoreCard_ModMolave LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " WHERE (tbl_Scoring_ScoreCard_ModMolave.TournamentKey = " & TournamentKey & ") " & _
            " ORDER BY tbl_Scoring_ScoreCard_ModMolave.CtrlNo DESC"
    Case Else: Exit Function
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtCtrl.Text = rs!CtrlNo
    txtDate.Text = Format(rs!dDate, "mm/dd/yyyy")
    txtPlayer.Text = rs!PlayerName
    txtHandicap.Text = rs!HandiCap
    
    If Trim(rs!Class) = "" Then
        u = "SELECT Class " & _
            " From tbl_Scoring_TournamentInfo_Class " & _
            " WHERE (TournamentKey = " & TournamentKey & ") " & _
            " AND (HFrom <= " & rs!HandiCap & ") " & _
            " AND (HTo >= " & rs!HandiCap & ")"
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
    't = "SELECT TeamKey " & _
        " From tbl_Scoring_Team_Detail " & _
        " WHERE (PlayerKey = " & rs!PlayerKey & ")"
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
    If CDbl(TeamTmp) > 0 Then
        t = "SELECT tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
            " tbl_Scoring_PlayerName.MiddleName, " & _
            " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.Score) " & _
            " From tbl_Scoring_ScoreCard_ModMolave " & _
            " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS Score, " & _
            " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.GrossPoints) " & _
            " From tbl_Scoring_ScoreCard_ModMolave " & _
            " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS GrossPts, " & _
            " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.NetPoints) " & _
            " From tbl_Scoring_ScoreCard_ModMolave " & _
            " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS NetPoints " & _
            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " Where (tbl_Scoring_Team_Detail.TeamKey = " & TeamTmp & ") " & _
            " And (tbl_Scoring_Team_Detail.PlayerKey <> " & rs!PlayerKey & ") " & _
            " Order By ISNULL((SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.GrossPoints) " & _
            " From tbl_Scoring_ScoreCard_ModMolave " & _
            " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) DESC"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        While Not rt.EOF
            Set x = lstTeamMates.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = Trim(rt!LastName) & ",  " & Trim(rt!FirstName) & IIf(Trim(rt!MiddleName) = "", "", "  " & rt!MiddleName)
            x.SubItems(2) = rt!Score
            x.SubItems(3) = rt!GrossPts
            x.SubItems(4) = rt!NetPoints
            rt.MoveNext
        Wend
        rt.Close
    End If
    
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", "Last Modified : " & rs!LastModified)
    
    i = -1
    t = "SELECT Par, Handicap, Score, Gross, Net " & _
        " From tbl_Scoring_ScoreCard_ModMolave_Detail " & _
        " Where (ScoreCardKey = " & rs!PK & ") " & _
        " ORDER BY Hole "
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
txtDateAdd.Text = Format(FormatDateTime(Date, vbShortDate), "mm/dd/yyyy")
picSearchAdd.Visible = True
txtSearchAdd.SetFocus
End Function

Private Function PRESS_F2()
Dim i
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
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
ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_ModMolave WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
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
        " From tbl_Scoring_ScoreCard_ModMolave " & _
        " WHERE (TournamentKey = " & TournamentKey & ") " & _
        " AND (PlayerKey = " & PlayerKey & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    TourNoOfPlaysTmp = rs!NoofRec
    rs.Close
    
    If CDbl(TourNoOfPlaysTmp) + 1 > CDbl(DaysPlayerToPlay) Then MsgBox "Number of Plays Exceeded!                  ", vbCritical, "Error...": Exit Function
    
    strCtrlNo = "00000001"
    s = "SELECT TOP 1 CtrlNo " & _
        " FROM tbl_Scoring_ScoreCard_ModMolave " & _
        " WHERE (TournamentKey = " & TournamentKey & ") " & _
        " ORDER BY CtrlNo DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        strCtrlNo = Format(CDbl(rs!CtrlNo) + 1, "0000000#")
    End If
    rs.Close
    
    Do
        s = "SELECT tbl_Scoring_ScoreCard_ModMolave.* " & _
            " FROM tbl_Scoring_ScoreCard_ModMolave " & _
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
    
    ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_ModMolave " & _
                      " (TournamentKey, PlayerKey, DDate, LastModified, CtrlNo) " & _
                      " VALUES (" & TournamentKey & ", " & PlayerKey & ", " & _
                      " '" & FormatDateTime(txtDate.Text, vbShortDate) & "', " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "', '" & strCtrlNo & "')"
    
    SCardKey = 0
    s = "SELECT PK " & _
        " FROM tbl_Scoring_ScoreCard_ModMolave " & _
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
            
            ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_ModMolave_Detail " & _
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
    SCardKey = Statusbar1.Panels(1).Text
    ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_ModMolave " & _
                      " SET LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & SCardKey & ")"
    
    ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_ModMolave_Detail WHERE (ScoreCardKey = " & SCardKey & ")"
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
        
        ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_ModMolave_Detail " & _
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
cmbReportType.ListIndex = -1
cmbGroup.ListIndex = -1
cmbDivision.ListIndex = -1
txtTop.Text = "10"
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


Private Sub CLEARTEXT()
Dim i
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
'txtGrossPtsF.Text = "0"
'txtNetPtsF.Text = "0"
txtGrossScoreB.Text = "0"
'txtGrossPtsB.Text = "0"
'txtNetPtsB.Text = "0"
'txtGrossScoreTot.Text = "0"
'txtGrossPtsTot.Text = "0"
'txtNetPtsTot.Text = "0"

lstTeamMates.ListItems.Clear
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
End Sub

Private Sub LOCKTEXT(bln As Boolean)
Dim i
For i = 0 To 17
    txtGrossScore(i).Locked = bln
Next i
End Sub

Public Sub TOOLBARFUNC(intSel As Integer)
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
End Sub

'Private Sub LOAD_CARD()
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
'End Sub

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
If KeyCode = vbKeyReturn Then txtTop.SetFocus
End Sub


Private Sub cmbGender_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbGroup.SetFocus
End Sub

Private Sub cmbGroup_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbDivision.SetFocus
End Sub


Private Sub cmbReportType_Click()
If cmbReportType.ListIndex = -1 Then Exit Sub
'If cmbReportType.ListIndex = 0 Then
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
'Else
'    With cmbDivision
'        .Clear
'        .AddItem "-- ALL --"
'        s = "SELECT Class " & _
'            " From tbl_Scoring_TournamentInfo_Class " & _
'            " Where (TournamentKey = " & TournamentKey & ") " & _
'            " ORDER BY Class"
'        If rs.State = adStateOpen Then rs.Close
'        rs.Open s, ConnOmega
'        While Not rs.EOF
'            .AddItem rs!Class
'            rs.MoveNext
'        Wend
'        rs.Close
'    End With
'End If
End Sub

Private Sub cmbReportType_GotFocus()
'If KeyCode = vbKeyReturn Then cmbGroup.SetFocus
End Sub

Private Sub cmbReportType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbGender.SetFocus
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

s = "SELECT COUNT(*) AS NoofRec " & _
    " From tbl_Scoring_ScoreCard_ModMolave " & _
    " WHERE (TournamentKey = " & TournamentKey & ") " & _
    " AND (PlayerKey = " & lstResultAdd.ItemData(lstResultAdd.ListIndex) & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
TourNoOfPlaysTmp = rs!NoofRec
rs.Close

If CDbl(TourNoOfPlaysTmp) + 1 > CDbl(DaysPlayerToPlay) Then MsgBox "Number of Plays Exceeded!                  ", vbCritical, "Error...": Exit Sub

CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
PlayerKey = lstResultAdd.ItemData(lstResultAdd.ListIndex)
txtPlayer.Text = lstResultAdd.List(lstResultAdd.ListIndex)
txtDate.Text = Format(FormatDateTime(txtDateAdd.Text, vbShortDate), "mm/dd/yyyy")

TeamTmp = 0
txtTeamName.Text = ""
's = "SELECT TeamKey " & _
'    " From tbl_Scoring_Team_Detail " & _
'    " WHERE (PlayerKey = " & PlayerKey & ")"
'If rs.State = adStateOpen Then rs.Close
'rs.Open s, ConnOmega
'If rs.RecordCount > 0 Then
'    TeamTmp = rs!TeamKey
'End If
'rs.Close
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
        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.Score) " & _
        " From tbl_Scoring_ScoreCard_ModMolave " & _
        " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS Score, " & _
        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.GrossPoints) " & _
        " From tbl_Scoring_ScoreCard_ModMolave " & _
        " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS GrossPts, " & _
        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.NetPoints) " & _
        " From tbl_Scoring_ScoreCard_ModMolave " & _
        " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS NetPoints " & _
        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
        " Where (tbl_Scoring_Team_Detail.TeamKey = " & TeamTmp & ") " & _
        " And (tbl_Scoring_Team_Detail.PlayerKey <> " & PlayerKey & ") " & _
        " Order By ISNULL((SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.GrossPoints) " & _
        " From tbl_Scoring_ScoreCard_ModMolave " & _
        " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) DESC"
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
    txtHandicap.Text = rs!HandiCap
    If Trim(IIf(IsNull(rs!Class), "", rs!Class)) = "" Then
        u = "SELECT Class " & _
            " From tbl_Scoring_TournamentInfo_Class " & _
            " WHERE (TournamentKey = " & TournamentKey & ") " & _
            " AND (HFrom <= " & rs!HandiCap & ") " & _
            " AND (HTo >= " & rs!HandiCap & ")"
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
    'MsgBox Arr(0) & " - " & Arr(1)
    If DateValue(FormatDateTime(txtDatePrint.Text, vbShortDate)) < DateValue(Arr(0)) Then
        MsgBox "Date out of Range!                          ", vbCritical, "Error...": Exit Sub
    ElseIf DateValue(FormatDateTime(txtDatePrint.Text, vbShortDate)) > DateValue(Arr(1)) Then
        MsgBox "Date out of Range!                          ", vbCritical, "Error...": Exit Sub
    End If
End If

'Exit Sub

With MainForm.CommonDialog1
    .CancelError = True
    On Error GoTo ErrorHandler
    .DialogTitle = "Save"
    '.Filter = "Excel(*.xls)|*.xls"
    .Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
    .ShowSave
    sFileName = Trim(.Filename)
End With
Screen.MousePointer = vbHourglass
TimerReport.Enabled = True
Screen.MousePointer = vbDefault
Exit Sub
ErrorHandler:
Exit Sub
End Sub

Private Sub cmdOKSearch_Click()
If cmbDate.ListIndex = -1 Then Exit Sub
BROWSER cmbDate.ItemData(cmbDate.ListIndex), "is_FIND"
cmdCancelSearch_Click
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

'MsgBox "pass"

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
    .AddItem "INDIVIDUAL"
    .AddItem "TEAM"
    .AddItem "RESULT"
End With

With cmbGroup
    .Clear
    .AddItem "NET POINTS"
    .AddItem "GROSS POINTS"
    .AddItem "GROSS SCORE"
End With

With cmbGender
    .Clear
    .AddItem "MALE"
    .AddItem "FEMALE"
End With


'Me.Caption = "Score Card (Modified Molave)"
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

LOAD_CARD dDateEnd, FGrid
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
If picSearchAdd.Visible = True Then Cancel = -1
If picSearch.Visible = True Then Cancel = -1
If picPrint.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstResult_Click()
If lstResult.ListIndex = -1 Then cmbDate.Clear: Exit Sub
cmbDate.Clear
s = "SELECT PK, DDate " & _
    " From tbl_Scoring_ScoreCard_ModMolave " & _
    " Where (TournamentKey = " & TournamentKey & ") " & _
    " And (PlayerKey = " & lstResult.ItemData(lstResult.ListIndex) & ") " & _
    " ORDER BY DDate DESC"
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

Private Sub TimerReport_Timer()
TimerReport.Enabled = False
Dim RowCnt, ColCnt, strRange, i, j, Arr, iDay, _
dDate, sTotal, RowCntTmp, ColCntTmp, strRangeFrom, _
strRangeTo, dTotalTeam, iTeamCounter, HeaderRow
Dim WorkbookName As String
Dim iWorkSheet As Integer
Select Case cmbReportType.ListIndex
    Case 0  'INDIVIDUAL
        Select Case cmbGroup.ListIndex
            Case 0  'NetPoints
                
                picPrint.Enabled = False
                picProgress.ZOrder 0
                picProgressBar.BackColor = &HFFFFFF
                picProgress.Visible = True
                
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
                        .Range(strRange).Value = "Individual (Net Points) [" & IIf(cmbGender.ListIndex = 0, "MALE", "FEMALE") & "]"
                    Else
                        .Range(strRange).Value = "Individual [Class " & cmbDivision.List(cmbDivision.ListIndex) & "] (Net Points) [" & IIf(cmbGender.ListIndex = 0, "MALE", "FEMALE") & "]"
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
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard_ModMolave.TournamentKey, tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard_ModMolave.NetPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Net) AS Holes_B9, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard_ModMolave AS tbl_Scoring_ScoreCard_ModMolave LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard_ModMolave.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") " & _
                                " AND (tbl_Scoring_ScoreCard_ModMolave.DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "') " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_ScoreCard_ModMolave.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard_ModMolave.NetPoints) DESC, SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Net) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard_ModMolave.TournamentKey, tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard_ModMolave.NetPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Net) AS Holes_B9, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard_ModMolave AS tbl_Scoring_ScoreCard_ModMolave LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard_ModMolave.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_ScoreCard_ModMolave.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard_ModMolave.NetPoints) DESC, SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Net) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        End If
                    Else
                        If IsDate(txtDatePrint.Text) = True Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard_ModMolave.TournamentKey, tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard_ModMolave.NetPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Net) AS Holes_B9, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard_ModMolave AS tbl_Scoring_ScoreCard_ModMolave LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard_ModMolave.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " AND (tbl_Scoring_ScoreCard_ModMolave.DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "') " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_ScoreCard_ModMolave.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard_ModMolave.NetPoints) DESC, SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Net) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard_ModMolave.TournamentKey, tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard_ModMolave.NetPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Net) AS Holes_B9, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard_ModMolave AS tbl_Scoring_ScoreCard_ModMolave LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard_ModMolave.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_ScoreCard_ModMolave.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard_ModMolave.NetPoints) DESC, SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Net) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Net) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
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
                        .Range(strRange).Value = rs!HandiCap
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
                                        " From tbl_Scoring_ScoreCard_ModMolave " & _
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
                                    " From tbl_Scoring_ScoreCard_ModMolave " & _
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
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard_ModMolave.TournamentKey, tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard_ModMolave.GrossPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Gross) AS Holes_B9, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard_ModMolave AS tbl_Scoring_ScoreCard_ModMolave LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard_ModMolave.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") " & _
                                " AND (tbl_Scoring_ScoreCard_ModMolave.DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "') " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_ScoreCard_ModMolave.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard_ModMolave.GrossPoints) DESC, SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Gross) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard_ModMolave.TournamentKey, tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard_ModMolave.GrossPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Gross) AS Holes_B9, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard_ModMolave AS tbl_Scoring_ScoreCard_ModMolave LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard_ModMolave.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_ScoreCard_ModMolave.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard_ModMolave.GrossPoints) DESC, SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Gross) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        End If
                    Else
                        If IsDate(txtDatePrint.Text) = True Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard_ModMolave.TournamentKey, tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard_ModMolave.GrossPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Gross) AS Holes_B9, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard_ModMolave AS tbl_Scoring_ScoreCard_ModMolave LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard_ModMolave.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ")  AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " AND (tbl_Scoring_ScoreCard_ModMolave.DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "') " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_ScoreCard_ModMolave.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard_ModMolave.GrossPoints) DESC, SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Gross) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard_ModMolave.TournamentKey, tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard_ModMolave.GrossPoints) AS NetPoints, SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Gross) AS Holes_B9, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard_ModMolave AS tbl_Scoring_ScoreCard_ModMolave LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard_ModMolave.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ")  AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_ScoreCard_ModMolave.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard_ModMolave.GrossPoints) DESC, SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Gross) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) DESC, (SELECT SUM(T_Detail_1.Gross) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) DESC"
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
                        .Range(strRange).Value = rs!HandiCap
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
                                    " From tbl_Scoring_ScoreCard_ModMolave " & _
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
                                    " From tbl_Scoring_ScoreCard_ModMolave " & _
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
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard_ModMolave.TournamentKey, tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard_ModMolave.Score) AS NetPoints, SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Score) AS Holes_B9, (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard_ModMolave AS tbl_Scoring_ScoreCard_ModMolave LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard_ModMolave.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") " & _
                                " AND (tbl_Scoring_ScoreCard_ModMolave.DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "') " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_ScoreCard_ModMolave.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard_ModMolave.Score), SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Score) , (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)), (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) "
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard_ModMolave.TournamentKey, tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard_ModMolave.Score) AS NetPoints, SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Score) AS Holes_B9, (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard_ModMolave AS tbl_Scoring_ScoreCard_ModMolave LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard_ModMolave.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ") " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_ScoreCard_ModMolave.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard_ModMolave.Score), SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Score) , (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)), (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) "
                        End If
                    Else
                        If IsDate(txtDatePrint.Text) = True Then
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard_ModMolave.TournamentKey, tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard_ModMolave.Score) AS NetPoints, SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Score) AS Holes_B9, (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard_ModMolave AS tbl_Scoring_ScoreCard_ModMolave LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard_ModMolave.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ")  AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " AND (tbl_Scoring_ScoreCard_ModMolave.DDate = '" & FormatDateTime(txtDatePrint.Text, vbShortDate) & "') " & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_ScoreCard_ModMolave.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard_ModMolave.Score) , SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Score) , (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) , (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) "
                        Else
                            s = "SELECT TOP " & RETURNTEXTVALUE(txtTop) & " tbl_Scoring_ScoreCard_ModMolave.TournamentKey, tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap, " & _
                                " SUM(tbl_Scoring_ScoreCard_ModMolave.Score) AS NetPoints, SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Score) AS Holes_B9, (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) AS Holes_B6, " & _
                                " (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK " & _
                                " WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND " & _
                                " (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) AS Holes_B3 FROM tbl_Scoring_ScoreCard_ModMolave AS tbl_Scoring_ScoreCard_ModMolave LEFT OUTER JOIN tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                                " WHERE (tbl_Scoring_ScoreCard_ModMolave.TournamentKey = " & TournamentKey & ") AND (tbl_Scoring_PlayerName.Gender = " & cmbGender.ListIndex + 1 & ")  AND ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS Class WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & cmbDivision.List(cmbDivision.ListIndex) & "')" & _
                                " GROUP BY tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, tbl_Scoring_PlayerName.MiddleName, tbl_Scoring_PlayerName.HandiCap , tbl_Scoring_ScoreCard_ModMolave.PlayerKey, tbl_Scoring_ScoreCard_ModMolave.TournamentKey " & _
                                " ORDER BY SUM(tbl_Scoring_ScoreCard_ModMolave.Score) , SUM(tbl_Scoring_ScoreCard_ModMolave.Back9Score) , (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date " & _
                                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 13) AND (T_Detail_1.Hole <= 18)) , (SELECT SUM(T_Detail_1.Score) AS Holes_B6 FROM tbl_Scoring_ScoreCard_ModMolave_Detail AS T_Detail_1 LEFT OUTER JOIN " & _
                                " tbl_Scoring_ScoreCard_ModMolave AS T_Master_1 ON T_Detail_1.ScoreCardKey = T_Master_1.PK WHERE (T_Master_1.TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (T_Master_1.PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) AND (T_Master_1.DDate = " & _
                                " (SELECT TOP 1 DDate FROM tbl_Scoring_ScoreCard_ModMolave AS Top_Date WHERE (TournamentKey = tbl_Scoring_ScoreCard_ModMolave.TournamentKey) AND (PlayerKey = tbl_Scoring_ScoreCard_ModMolave.PlayerKey) ORDER BY DDate DESC)) AND (T_Detail_1.Hole >= 15) AND (T_Detail_1.Hole <= 18)) "
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
                        .Range(strRange).Value = rs!HandiCap
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
                                        " From tbl_Scoring_ScoreCard_ModMolave " & _
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
                                    " From tbl_Scoring_ScoreCard_ModMolave " & _
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
                
                ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_ModMolave_Team_Rep " & _
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
                    ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_ModMolave_Team_Rep " & _
                                      " (LogInName, TeamKey, TeamName) " & _
                                      " VALUES ('" & gbl_UserName & "', " & rs!PK & ", '" & FORMATSQL(rs!TeamName) & "')"
                    
                    dTotalTeam = 0
                    t = "SELECT  TOP " & TeamPlayer2Cnt & " tbl_Scoring_Team_Detail.PlayerKey, " & _
                        " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                        " tbl_Scoring_PlayerName.MiddleName, " & _
                        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.NetPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard_ModMolave " & _
                        " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS NetPoints " & _
                        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                        " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                        " Order By ISNULL((SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.NetPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard_ModMolave " & _
                        " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) DESC"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        dTotalTeam = dTotalTeam + CDbl(rt!NetPoints)
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_ModMolave_Team_Rep " & _
                                      " SET Score = " & CDbl(dTotalTeam) & " " & _
                                      " WHERE (LogInName = '" & gbl_UserName & "') " & _
                                      " AND (TeamKey = " & rs!PK & ")"
                    
                    dTotalTeam = 0: iTeamCounter = 0
                    t = "SELECT tbl_Scoring_Team_Detail.PlayerKey, " & _
                        " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                        " tbl_Scoring_PlayerName.MiddleName, " & _
                        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.NetPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard_ModMolave " & _
                        " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS NetPoints " & _
                        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                        " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                        " Order By ISNULL((SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.NetPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard_ModMolave " & _
                        " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) DESC"
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
                    
                    ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_ModMolave_Team_Rep " & _
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
                        " FROM tbl_Scoring_ScoreCard_ModMolave_Team_Rep " & _
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
                            " (SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.NetPoints) AS NetPoints " & _
                            " From tbl_Scoring_ScoreCard_ModMolave " & _
                            " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)) AS NetPoints " & _
                            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                            " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                            " Order By (SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.NetPoints) AS NetPoints " & _
                            " From tbl_Scoring_ScoreCard_ModMolave " & _
                            " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)) DESC"
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
                                    " From tbl_Scoring_ScoreCard_ModMolave " & _
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
                
                ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_ModMolave_Team_Rep " & _
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
                    ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_ModMolave_Team_Rep " & _
                                      " (LogInName, TeamKey, TeamName) " & _
                                      " VALUES ('" & gbl_UserName & "', " & rs!PK & ", '" & FORMATSQL(rs!TeamName) & "')"
                    
                    dTotalTeam = 0
                    t = "SELECT  TOP " & TeamPlayer2Cnt & " tbl_Scoring_Team_Detail.PlayerKey, " & _
                        " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                        " tbl_Scoring_PlayerName.MiddleName, " & _
                        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.GrossPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard_ModMolave " & _
                        " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS NetPoints " & _
                        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                        " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                        " Order By ISNULL((SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.GrossPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard_ModMolave " & _
                        " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) DESC"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        dTotalTeam = dTotalTeam + CDbl(rt!NetPoints)
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_ModMolave_Team_Rep " & _
                                      " SET Score = " & CDbl(dTotalTeam) & " " & _
                                      " WHERE (LogInName = '" & gbl_UserName & "') " & _
                                      " AND (TeamKey = " & rs!PK & ")"
                    
                    dTotalTeam = 0: iTeamCounter = 0
                    t = "SELECT tbl_Scoring_Team_Detail.PlayerKey, " & _
                        " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                        " tbl_Scoring_PlayerName.MiddleName, " & _
                        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.GrossPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard_ModMolave " & _
                        " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS NetPoints " & _
                        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                        " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                        " Order By ISNULL((SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.GrossPoints) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard_ModMolave " & _
                        " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) DESC"
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
                    
                    ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_ModMolave_Team_Rep " & _
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
                        " FROM tbl_Scoring_ScoreCard_ModMolave_Team_Rep " & _
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
                            " (SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.GrossPoints) AS NetPoints " & _
                            " From tbl_Scoring_ScoreCard_ModMolave " & _
                            " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)) AS NetPoints " & _
                            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                            " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                            " Order By (SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.GrossPoints) AS NetPoints " & _
                            " From tbl_Scoring_ScoreCard_ModMolave " & _
                            " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)) DESC"
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
                                    " From tbl_Scoring_ScoreCard_ModMolave " & _
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
                
                ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_ModMolave_Team_Rep " & _
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
                    ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_ModMolave_Team_Rep " & _
                                      " (LogInName, TeamKey, TeamName) " & _
                                      " VALUES ('" & gbl_UserName & "', " & rs!PK & ", '" & FORMATSQL(rs!TeamName) & "')"
                    
                    dTotalTeam = 0
                    t = "SELECT  TOP " & TeamPlayer2Cnt & " tbl_Scoring_Team_Detail.PlayerKey, " & _
                        " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                        " tbl_Scoring_PlayerName.MiddleName, " & _
                        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.Score) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard_ModMolave " & _
                        " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS NetPoints " & _
                        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                        " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                        " Order By ISNULL((SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.Score) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard_ModMolave " & _
                        " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0)"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    While Not rt.EOF
                        dTotalTeam = dTotalTeam + CDbl(rt!NetPoints)
                        rt.MoveNext
                    Wend
                    rt.Close
                    
                    ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_ModMolave_Team_Rep " & _
                                      " SET Score = " & CDbl(dTotalTeam) & " " & _
                                      " WHERE (LogInName = '" & gbl_UserName & "') " & _
                                      " AND (TeamKey = " & rs!PK & ")"
                    
                    dTotalTeam = 0: iTeamCounter = 0
                    t = "SELECT tbl_Scoring_Team_Detail.PlayerKey, " & _
                        " tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
                        " tbl_Scoring_PlayerName.MiddleName, " & _
                        " ISNULL((SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.Score) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard_ModMolave " & _
                        " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) AS NetPoints " & _
                        " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                        " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                        " Order By ISNULL((SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.Score) AS NetPoints " & _
                        " From tbl_Scoring_ScoreCard_ModMolave " & _
                        " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0)"
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
                    
                    ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_ModMolave_Team_Rep " & _
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
                        " FROM tbl_Scoring_ScoreCard_ModMolave_Team_Rep " & _
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
                            " (SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.Score) AS NetPoints " & _
                            " From tbl_Scoring_ScoreCard_ModMolave " & _
                            " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)) AS NetPoints " & _
                            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
                            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                            " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!PK & ") " & _
                            " Order By (SELECT SUM(tbl_Scoring_ScoreCard_ModMolave.Score) AS NetPoints " & _
                            " From tbl_Scoring_ScoreCard_ModMolave " & _
                            " WHERE (tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)) "
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
                                    " From tbl_Scoring_ScoreCard_ModMolave " & _
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
        
'        WorkbookName = sFileName
'        iWorkSheet = 1: RowCnt = 0
'        Set xlsApp = CreateObject("Excel.Application")
'        xlsApp.Visible = False
'        xlsApp.Workbooks.Add
'        xlsApp.DisplayAlerts = False
'        xlsApp.Workbooks(1).Sheets(2).Delete
'        xlsApp.Workbooks(1).Sheets(2).Delete
'        With xlsApp.Workbooks(1).Sheets(iWorkSheet)
'            .Activate
'            .Name = "Top " & CStr(Trim(txtTop.Text))
'        End With
'        With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
'
'        End With
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


Private Sub txtDatePrint_LostFocus()
If IsDate(txtDatePrint.Text) = True Then
    txtDatePrint.Text = Format(FormatDateTime(txtDatePrint.Text, vbShortDate), "mm/dd/yyyy")
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
txtSGrossF.Text = RETURNTEXTVALUE(txtGrossPtsF)
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
txtSNetF.Text = RETURNTEXTVALUE(txtNetPtsF)
End Sub

Private Sub txtNetPtsF_Change()
txtNetPtsTot.Text = RETURNTEXTVALUE(txtNetPtsF) + _
                    RETURNTEXTVALUE(txtNetPtsB)
txtSNetF.Text = RETURNTEXTVALUE(txtNetPtsF)
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstResult.Clear: cmbDate.Clear: Exit Sub
lstResult.Clear: cmbDate.Clear
s = "SELECT tbl_Scoring_PlayerName.PK, " & _
    " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName " & _
    " FROM tbl_Scoring_ScoreCard_ModMolave LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_ModMolave.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_ScoreCard_ModMolave.TournamentKey = " & TournamentKey & ") " & _
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

Private Sub txtTop_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKPrint_Click
End Sub

Private Sub txtTop_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub


