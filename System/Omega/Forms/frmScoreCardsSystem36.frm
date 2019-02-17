VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmScoreCardsSystem36 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmScoreCardsSystem36.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   13335
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5535
      ScaleWidth      =   13215
      TabIndex        =   20
      Top             =   0
      Width           =   13215
      Begin VB.Timer TimerPrintResult 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2040
         Top             =   1920
      End
      Begin RPVGCC.b8Container b8Container7 
         Height          =   1575
         Left            =   10320
         TabIndex        =   230
         Top             =   2040
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2778
         BackColor       =   49152
         Begin VB.PictureBox Picture10 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   120
            ScaleHeight     =   1335
            ScaleWidth      =   2655
            TabIndex        =   231
            Top             =   120
            Width           =   2655
            Begin MSComctlLib.ListView lstScores 
               Height          =   1335
               Left            =   0
               TabIndex        =   232
               Top             =   0
               Width           =   2655
               _ExtentX        =   4683
               _ExtentY        =   2355
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               HideColumnHeaders=   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   3
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "PlayerName"
                  Object.Width           =   3087
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   2
                  Text            =   "Gross Pts"
                  Object.Width           =   1411
               EndProperty
            End
         End
      End
      Begin VB.TextBox txtCtrl 
         Height          =   285
         Left            =   0
         TabIndex        =   23
         Top             =   1920
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   0
         ScaleHeight     =   735
         ScaleWidth      =   15000
         TabIndex        =   21
         Top             =   0
         Width           =   15000
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   570
            Left            =   0
            TabIndex        =   22
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
         Begin VB.Line Line8 
            BorderColor     =   &H00808080&
            X1              =   0
            X2              =   15000
            Y1              =   650
            Y2              =   650
         End
         Begin VB.Line Line7 
            BorderColor     =   &H00808080&
            X1              =   0
            X2              =   15000
            Y1              =   90
            Y2              =   90
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00C0C0C0&
            X1              =   0
            X2              =   15000
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            X1              =   0
            X2              =   15000
            Y1              =   660
            Y2              =   660
         End
      End
      Begin RPVGCC.b8Container b8Container3 
         Height          =   1095
         Left            =   5880
         TabIndex        =   24
         Top             =   840
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   1931
         BackColor       =   49152
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   855
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   7095
            TabIndex        =   25
            Top             =   120
            Width           =   7095
            Begin VB.TextBox txtTourDate 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1095
               TabIndex        =   27
               Text            =   "06/01/2010 - 06/04/2010"
               Top             =   480
               Width           =   5895
            End
            Begin VB.TextBox txtTournament 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1095
               TabIndex        =   26
               Top             =   120
               Width           =   5895
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Date Range"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Tournament"
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   120
               Width           =   1335
            End
         End
      End
      Begin RPVGCC.b8Container b8Container5 
         Height          =   1095
         Left            =   120
         TabIndex        =   30
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
            TabIndex        =   31
            Top             =   120
            Width           =   5415
            Begin VB.TextBox txtPlayer 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   720
               TabIndex        =   34
               Top             =   120
               Width           =   4575
            End
            Begin VB.TextBox txtDay 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   4560
               TabIndex        =   33
               Top             =   480
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox txtDate 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   720
               TabIndex        =   32
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Player"
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Day"
               Height          =   255
               Left            =   3720
               TabIndex        =   36
               Top             =   480
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               Height          =   255
               Left            =   120
               TabIndex        =   35
               Top             =   480
               Width           =   495
            End
         End
      End
      Begin RPVGCC.b8Container b8Container2 
         Height          =   1335
         Left            =   2280
         TabIndex        =   38
         Top             =   2160
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
            TabIndex        =   39
            Top             =   120
            Width           =   3255
            Begin VB.TextBox txtSNetTot 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   2400
               TabIndex        =   45
               Top             =   720
               Width           =   735
            End
            Begin VB.TextBox txtSGrossTot 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   2400
               TabIndex        =   44
               Top             =   360
               Width           =   735
            End
            Begin VB.TextBox txtSNetB 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1680
               TabIndex        =   43
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtSGrossB 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1680
               TabIndex        =   42
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox txtSNetF 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1080
               TabIndex        =   41
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtSGrossF 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   1080
               TabIndex        =   40
               Top             =   360
               Width           =   495
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Total"
               Height          =   255
               Left            =   2400
               TabIndex        =   51
               Top             =   120
               Width           =   735
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "B - 9"
               Height          =   255
               Left            =   1680
               TabIndex        =   50
               Top             =   120
               Width           =   495
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "F - 9"
               Height          =   255
               Left            =   1080
               TabIndex        =   49
               Top             =   120
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
               TabIndex        =   48
               Top             =   0
               Width           =   975
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Net Points"
               Height          =   255
               Left            =   120
               TabIndex        =   47
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Gross Points"
               Height          =   255
               Left            =   120
               TabIndex        =   46
               Top             =   360
               Width           =   975
            End
         End
      End
      Begin RPVGCC.b8Container b8Container1 
         Height          =   1770
         Left            =   120
         TabIndex        =   52
         Top             =   3720
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   3122
         BackColor       =   13023396
         ShadowColor1    =   49152
         ShadowColor2    =   8454016
         Begin VB.PictureBox picScoreMain 
            Appearance      =   0  'Flat
            BackColor       =   &H00C6B8A4&
            ForeColor       =   &H80000008&
            Height          =   1680
            Left            =   50
            ScaleHeight     =   1650
            ScaleWidth      =   12990
            TabIndex        =   53
            Top             =   50
            Width           =   13020
            Begin VB.PictureBox picScoreDis 
               Appearance      =   0  'Flat
               BackColor       =   &H00C6B8A4&
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   1440
               Left            =   -10
               ScaleHeight     =   1410
               ScaleWidth      =   12990
               TabIndex        =   106
               Top             =   -10
               Width           =   13020
               Begin MSFlexGridLib.MSFlexGrid FGrid 
                  Height          =   1545
                  Left            =   -105
                  TabIndex        =   107
                  Top             =   -30
                  Width           =   13635
                  _ExtentX        =   24051
                  _ExtentY        =   2725
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
               Height          =   2415
               Left            =   -10
               ScaleHeight     =   2385
               ScaleWidth      =   12990
               TabIndex        =   54
               Top             =   1410
               Width           =   13020
               Begin VB.PictureBox Picture7 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   2415
                  Left            =   1980
                  ScaleHeight     =   2385
                  ScaleWidth      =   11070
                  TabIndex        =   60
                  Top             =   230
                  Width           =   11100
                  Begin VB.TextBox txtNetFB 
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
                     Left            =   10200
                     TabIndex        =   229
                     Text            =   "0"
                     Top             =   1200
                     Width           =   620
                  End
                  Begin VB.TextBox txtNetF 
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
                     Left            =   8760
                     TabIndex        =   228
                     Text            =   "0"
                     Top             =   1200
                     Width           =   620
                  End
                  Begin VB.TextBox txt3Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   17
                     Left            =   8145
                     TabIndex        =   227
                     Text            =   "0"
                     Top             =   1800
                     Width           =   460
                  End
                  Begin VB.TextBox txt3Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   16
                     Left            =   7695
                     TabIndex        =   226
                     Text            =   "0"
                     Top             =   1800
                     Width           =   460
                  End
                  Begin VB.TextBox txt3Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   15
                     Left            =   7245
                     TabIndex        =   225
                     Text            =   "0"
                     Top             =   1800
                     Width           =   460
                  End
                  Begin VB.TextBox txt3Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   14
                     Left            =   6810
                     TabIndex        =   224
                     Text            =   "0"
                     Top             =   1800
                     Width           =   460
                  End
                  Begin VB.TextBox txt3Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   13
                     Left            =   6345
                     TabIndex        =   223
                     Text            =   "0"
                     Top             =   1800
                     Width           =   460
                  End
                  Begin VB.TextBox txt3Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   12
                     Left            =   5880
                     TabIndex        =   222
                     Text            =   "0"
                     Top             =   1800
                     Width           =   460
                  End
                  Begin VB.TextBox txt3Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   11
                     Left            =   5460
                     TabIndex        =   221
                     Text            =   "0"
                     Top             =   1800
                     Width           =   460
                  End
                  Begin VB.TextBox txt3Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   10
                     Left            =   5010
                     TabIndex        =   220
                     Text            =   "0"
                     Top             =   1800
                     Width           =   460
                  End
                  Begin VB.TextBox txt3Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   9
                     Left            =   4560
                     TabIndex        =   219
                     Text            =   "0"
                     Top             =   1800
                     Width           =   460
                  End
                  Begin VB.TextBox txt2Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   17
                     Left            =   8145
                     TabIndex        =   218
                     Text            =   "0"
                     Top             =   1560
                     Width           =   460
                  End
                  Begin VB.TextBox txt2Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   16
                     Left            =   7695
                     TabIndex        =   217
                     Text            =   "0"
                     Top             =   1560
                     Width           =   460
                  End
                  Begin VB.TextBox txt2Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   15
                     Left            =   7245
                     TabIndex        =   216
                     Text            =   "0"
                     Top             =   1560
                     Width           =   460
                  End
                  Begin VB.TextBox txt2Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   14
                     Left            =   6810
                     TabIndex        =   215
                     Text            =   "0"
                     Top             =   1560
                     Width           =   460
                  End
                  Begin VB.TextBox txt2Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   13
                     Left            =   6345
                     TabIndex        =   214
                     Text            =   "0"
                     Top             =   1560
                     Width           =   460
                  End
                  Begin VB.TextBox txt2Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   12
                     Left            =   5910
                     TabIndex        =   213
                     Text            =   "0"
                     Top             =   1560
                     Width           =   460
                  End
                  Begin VB.TextBox txt2Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   11
                     Left            =   5460
                     TabIndex        =   212
                     Text            =   "0"
                     Top             =   1560
                     Width           =   460
                  End
                  Begin VB.TextBox txt2Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   10
                     Left            =   5010
                     TabIndex        =   211
                     Text            =   "0"
                     Top             =   1560
                     Width           =   460
                  End
                  Begin VB.TextBox txt2Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   9
                     Left            =   4560
                     TabIndex        =   210
                     Text            =   "0"
                     Top             =   1560
                     Width           =   460
                  End
                  Begin VB.TextBox txtBoogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   17
                     Left            =   8145
                     TabIndex        =   209
                     Text            =   "0"
                     Top             =   1320
                     Width           =   460
                  End
                  Begin VB.TextBox txtBoogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   16
                     Left            =   7695
                     TabIndex        =   208
                     Text            =   "0"
                     Top             =   1320
                     Width           =   460
                  End
                  Begin VB.TextBox txtBoogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   15
                     Left            =   7245
                     TabIndex        =   207
                     Text            =   "0"
                     Top             =   1320
                     Width           =   460
                  End
                  Begin VB.TextBox txtBoogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   14
                     Left            =   6810
                     TabIndex        =   206
                     Text            =   "0"
                     Top             =   1320
                     Width           =   460
                  End
                  Begin VB.TextBox txtBoogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   13
                     Left            =   6345
                     TabIndex        =   205
                     Text            =   "0"
                     Top             =   1320
                     Width           =   460
                  End
                  Begin VB.TextBox txtBoogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   12
                     Left            =   5910
                     TabIndex        =   204
                     Text            =   "0"
                     Top             =   1320
                     Width           =   460
                  End
                  Begin VB.TextBox txtBoogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   11
                     Left            =   5460
                     TabIndex        =   203
                     Text            =   "0"
                     Top             =   1320
                     Width           =   460
                  End
                  Begin VB.TextBox txtBoogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   10
                     Left            =   5010
                     TabIndex        =   202
                     Text            =   "0"
                     Top             =   1320
                     Width           =   460
                  End
                  Begin VB.TextBox txtBoogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   9
                     Left            =   4560
                     TabIndex        =   201
                     Text            =   "0"
                     Top             =   1320
                     Width           =   460
                  End
                  Begin VB.TextBox txtPars 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   17
                     Left            =   8145
                     TabIndex        =   200
                     Text            =   "0"
                     Top             =   1080
                     Width           =   460
                  End
                  Begin VB.TextBox txtPars 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   16
                     Left            =   7695
                     TabIndex        =   199
                     Text            =   "0"
                     Top             =   1080
                     Width           =   460
                  End
                  Begin VB.TextBox txtPars 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   15
                     Left            =   7245
                     TabIndex        =   198
                     Text            =   "0"
                     Top             =   1080
                     Width           =   460
                  End
                  Begin VB.TextBox txtPars 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   14
                     Left            =   6810
                     TabIndex        =   197
                     Text            =   "0"
                     Top             =   1080
                     Width           =   460
                  End
                  Begin VB.TextBox txtPars 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   13
                     Left            =   6345
                     TabIndex        =   196
                     Text            =   "0"
                     Top             =   1080
                     Width           =   460
                  End
                  Begin VB.TextBox txtPars 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   12
                     Left            =   5910
                     TabIndex        =   195
                     Text            =   "0"
                     Top             =   1080
                     Width           =   460
                  End
                  Begin VB.TextBox txtPars 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   11
                     Left            =   5460
                     TabIndex        =   194
                     Text            =   "0"
                     Top             =   1080
                     Width           =   460
                  End
                  Begin VB.TextBox txtPars 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   10
                     Left            =   5010
                     TabIndex        =   193
                     Text            =   "0"
                     Top             =   1080
                     Width           =   460
                  End
                  Begin VB.TextBox txtPars 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   9
                     Left            =   4560
                     TabIndex        =   192
                     Text            =   "0"
                     Top             =   1080
                     Width           =   460
                  End
                  Begin VB.TextBox txtBirdies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   17
                     Left            =   8145
                     TabIndex        =   191
                     Text            =   "0"
                     Top             =   840
                     Width           =   460
                  End
                  Begin VB.TextBox txtBirdies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   16
                     Left            =   7695
                     TabIndex        =   190
                     Text            =   "0"
                     Top             =   840
                     Width           =   460
                  End
                  Begin VB.TextBox txtBirdies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   15
                     Left            =   7245
                     TabIndex        =   189
                     Text            =   "0"
                     Top             =   840
                     Width           =   460
                  End
                  Begin VB.TextBox txtBirdies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   14
                     Left            =   6810
                     TabIndex        =   188
                     Text            =   "0"
                     Top             =   840
                     Width           =   460
                  End
                  Begin VB.TextBox txtBirdies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   13
                     Left            =   6345
                     TabIndex        =   187
                     Text            =   "0"
                     Top             =   840
                     Width           =   460
                  End
                  Begin VB.TextBox txtBirdies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   12
                     Left            =   5910
                     TabIndex        =   186
                     Text            =   "0"
                     Top             =   840
                     Width           =   460
                  End
                  Begin VB.TextBox txtBirdies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   11
                     Left            =   5460
                     TabIndex        =   185
                     Text            =   "0"
                     Top             =   840
                     Width           =   460
                  End
                  Begin VB.TextBox txtBirdies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   10
                     Left            =   5010
                     TabIndex        =   184
                     Text            =   "0"
                     Top             =   840
                     Width           =   460
                  End
                  Begin VB.TextBox txtBirdies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   9
                     Left            =   4560
                     TabIndex        =   183
                     Text            =   "0"
                     Top             =   840
                     Width           =   460
                  End
                  Begin VB.TextBox txtEagle 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   17
                     Left            =   8145
                     TabIndex        =   182
                     Text            =   "0"
                     Top             =   600
                     Width           =   460
                  End
                  Begin VB.TextBox txtEagle 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   16
                     Left            =   7695
                     TabIndex        =   181
                     Text            =   "0"
                     Top             =   600
                     Width           =   460
                  End
                  Begin VB.TextBox txtEagle 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   15
                     Left            =   7245
                     TabIndex        =   180
                     Text            =   "0"
                     Top             =   600
                     Width           =   460
                  End
                  Begin VB.TextBox txtEagle 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   14
                     Left            =   6810
                     TabIndex        =   179
                     Text            =   "0"
                     Top             =   600
                     Width           =   460
                  End
                  Begin VB.TextBox txtEagle 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   13
                     Left            =   6345
                     TabIndex        =   178
                     Text            =   "0"
                     Top             =   600
                     Width           =   460
                  End
                  Begin VB.TextBox txtEagle 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   12
                     Left            =   5910
                     TabIndex        =   177
                     Text            =   "0"
                     Top             =   600
                     Width           =   460
                  End
                  Begin VB.TextBox txtEagle 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   11
                     Left            =   5460
                     TabIndex        =   176
                     Text            =   "0"
                     Top             =   600
                     Width           =   460
                  End
                  Begin VB.TextBox txtEagle 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   10
                     Left            =   5010
                     TabIndex        =   175
                     Text            =   "0"
                     Top             =   600
                     Width           =   460
                  End
                  Begin VB.TextBox txtEagle 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   9
                     Left            =   4560
                     TabIndex        =   174
                     Text            =   "0"
                     Top             =   600
                     Width           =   460
                  End
                  Begin VB.TextBox txt3Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   8
                     Left            =   3590
                     TabIndex        =   173
                     Text            =   "0"
                     Top             =   1800
                     Width           =   460
                  End
                  Begin VB.TextBox txt3Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   7
                     Left            =   3140
                     TabIndex        =   172
                     Text            =   "0"
                     Top             =   1800
                     Width           =   460
                  End
                  Begin VB.TextBox txt3Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   6
                     Left            =   2690
                     TabIndex        =   171
                     Text            =   "0"
                     Top             =   1800
                     Width           =   460
                  End
                  Begin VB.TextBox txt3Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   5
                     Left            =   2250
                     TabIndex        =   170
                     Text            =   "0"
                     Top             =   1800
                     Width           =   460
                  End
                  Begin VB.TextBox txt3Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   4
                     Left            =   1790
                     TabIndex        =   169
                     Text            =   "0"
                     Top             =   1800
                     Width           =   460
                  End
                  Begin VB.TextBox txt3Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   3
                     Left            =   1350
                     TabIndex        =   168
                     Text            =   "0"
                     Top             =   1800
                     Width           =   460
                  End
                  Begin VB.TextBox txt3Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   2
                     Left            =   900
                     TabIndex        =   167
                     Text            =   "0"
                     Top             =   1800
                     Width           =   460
                  End
                  Begin VB.TextBox txt3Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   1
                     Left            =   450
                     TabIndex        =   166
                     Text            =   "0"
                     Top             =   1800
                     Width           =   460
                  End
                  Begin VB.TextBox txt3Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   0
                     Left            =   0
                     TabIndex        =   165
                     Text            =   "0"
                     Top             =   1800
                     Width           =   460
                  End
                  Begin VB.TextBox txt2Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   8
                     Left            =   3590
                     TabIndex        =   164
                     Text            =   "0"
                     Top             =   1560
                     Width           =   460
                  End
                  Begin VB.TextBox txt2Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   7
                     Left            =   3140
                     TabIndex        =   163
                     Text            =   "0"
                     Top             =   1560
                     Width           =   460
                  End
                  Begin VB.TextBox txt2Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   6
                     Left            =   2690
                     TabIndex        =   162
                     Text            =   "0"
                     Top             =   1560
                     Width           =   460
                  End
                  Begin VB.TextBox txt2Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   5
                     Left            =   2250
                     TabIndex        =   161
                     Text            =   "0"
                     Top             =   1560
                     Width           =   460
                  End
                  Begin VB.TextBox txt2Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   4
                     Left            =   1790
                     TabIndex        =   160
                     Text            =   "0"
                     Top             =   1560
                     Width           =   460
                  End
                  Begin VB.TextBox txt2Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   3
                     Left            =   1350
                     TabIndex        =   159
                     Text            =   "0"
                     Top             =   1560
                     Width           =   460
                  End
                  Begin VB.TextBox txt2Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   2
                     Left            =   900
                     TabIndex        =   158
                     Text            =   "0"
                     Top             =   1560
                     Width           =   460
                  End
                  Begin VB.TextBox txt2Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   1
                     Left            =   450
                     TabIndex        =   157
                     Text            =   "0"
                     Top             =   1560
                     Width           =   460
                  End
                  Begin VB.TextBox txt2Boogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   345
                     Index           =   0
                     Left            =   0
                     TabIndex        =   156
                     Text            =   "0"
                     Top             =   1560
                     Width           =   460
                  End
                  Begin VB.TextBox txtBoogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   8
                     Left            =   3590
                     TabIndex        =   155
                     Text            =   "0"
                     Top             =   1320
                     Width           =   460
                  End
                  Begin VB.TextBox txtBoogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   7
                     Left            =   3140
                     TabIndex        =   154
                     Text            =   "0"
                     Top             =   1320
                     Width           =   460
                  End
                  Begin VB.TextBox txtBoogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   6
                     Left            =   2690
                     TabIndex        =   153
                     Text            =   "0"
                     Top             =   1320
                     Width           =   460
                  End
                  Begin VB.TextBox txtBoogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   5
                     Left            =   2250
                     TabIndex        =   152
                     Text            =   "0"
                     Top             =   1320
                     Width           =   460
                  End
                  Begin VB.TextBox txtBoogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   4
                     Left            =   1790
                     TabIndex        =   151
                     Text            =   "0"
                     Top             =   1320
                     Width           =   460
                  End
                  Begin VB.TextBox txtBoogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   3
                     Left            =   1350
                     TabIndex        =   150
                     Text            =   "0"
                     Top             =   1320
                     Width           =   460
                  End
                  Begin VB.TextBox txtBoogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   2
                     Left            =   900
                     TabIndex        =   149
                     Text            =   "0"
                     Top             =   1320
                     Width           =   460
                  End
                  Begin VB.TextBox txtBoogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   1
                     Left            =   450
                     TabIndex        =   148
                     Text            =   "0"
                     Top             =   1320
                     Width           =   460
                  End
                  Begin VB.TextBox txtBoogies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   0
                     Left            =   0
                     TabIndex        =   147
                     Text            =   "0"
                     Top             =   1320
                     Width           =   460
                  End
                  Begin VB.TextBox txtPars 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   8
                     Left            =   3590
                     TabIndex        =   146
                     Text            =   "0"
                     Top             =   1080
                     Width           =   460
                  End
                  Begin VB.TextBox txtPars 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   7
                     Left            =   3140
                     TabIndex        =   145
                     Text            =   "0"
                     Top             =   1080
                     Width           =   460
                  End
                  Begin VB.TextBox txtPars 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   6
                     Left            =   2690
                     TabIndex        =   144
                     Text            =   "0"
                     Top             =   1080
                     Width           =   460
                  End
                  Begin VB.TextBox txtPars 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   5
                     Left            =   2250
                     TabIndex        =   143
                     Text            =   "0"
                     Top             =   1080
                     Width           =   460
                  End
                  Begin VB.TextBox txtPars 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   4
                     Left            =   1790
                     TabIndex        =   142
                     Text            =   "0"
                     Top             =   1080
                     Width           =   460
                  End
                  Begin VB.TextBox txtPars 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   3
                     Left            =   1350
                     TabIndex        =   141
                     Text            =   "0"
                     Top             =   1080
                     Width           =   460
                  End
                  Begin VB.TextBox txtPars 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   2
                     Left            =   900
                     TabIndex        =   140
                     Text            =   "0"
                     Top             =   1080
                     Width           =   460
                  End
                  Begin VB.TextBox txtPars 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   1
                     Left            =   450
                     TabIndex        =   139
                     Text            =   "0"
                     Top             =   1080
                     Width           =   460
                  End
                  Begin VB.TextBox txtPars 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   0
                     Left            =   0
                     TabIndex        =   138
                     Text            =   "0"
                     Top             =   1080
                     Width           =   460
                  End
                  Begin VB.TextBox txtBirdies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   8
                     Left            =   3590
                     TabIndex        =   137
                     Text            =   "0"
                     Top             =   840
                     Width           =   460
                  End
                  Begin VB.TextBox txtBirdies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   7
                     Left            =   3140
                     TabIndex        =   136
                     Text            =   "0"
                     Top             =   840
                     Width           =   460
                  End
                  Begin VB.TextBox txtBirdies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   6
                     Left            =   2690
                     TabIndex        =   135
                     Text            =   "0"
                     Top             =   840
                     Width           =   460
                  End
                  Begin VB.TextBox txtBirdies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   5
                     Left            =   2250
                     TabIndex        =   134
                     Text            =   "0"
                     Top             =   840
                     Width           =   460
                  End
                  Begin VB.TextBox txtBirdies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   4
                     Left            =   1790
                     TabIndex        =   133
                     Text            =   "0"
                     Top             =   840
                     Width           =   460
                  End
                  Begin VB.TextBox txtBirdies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   3
                     Left            =   1350
                     TabIndex        =   132
                     Text            =   "0"
                     Top             =   840
                     Width           =   460
                  End
                  Begin VB.TextBox txtBirdies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   2
                     Left            =   900
                     TabIndex        =   131
                     Text            =   "0"
                     Top             =   840
                     Width           =   460
                  End
                  Begin VB.TextBox txtBirdies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   1
                     Left            =   450
                     TabIndex        =   130
                     Text            =   "0"
                     Top             =   840
                     Width           =   460
                  End
                  Begin VB.TextBox txtBirdies 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   0
                     Left            =   0
                     TabIndex        =   129
                     Text            =   "0"
                     Top             =   840
                     Width           =   460
                  End
                  Begin VB.TextBox txtEagle 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   8
                     Left            =   3590
                     TabIndex        =   128
                     Text            =   "0"
                     Top             =   600
                     Width           =   460
                  End
                  Begin VB.TextBox txtEagle 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   7
                     Left            =   3140
                     TabIndex        =   127
                     Text            =   "0"
                     Top             =   600
                     Width           =   460
                  End
                  Begin VB.TextBox txtEagle 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   6
                     Left            =   2690
                     TabIndex        =   126
                     Text            =   "0"
                     Top             =   600
                     Width           =   460
                  End
                  Begin VB.TextBox txtEagle 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   5
                     Left            =   2250
                     TabIndex        =   125
                     Text            =   "0"
                     Top             =   600
                     Width           =   460
                  End
                  Begin VB.TextBox txtEagle 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   4
                     Left            =   1790
                     TabIndex        =   124
                     Text            =   "0"
                     Top             =   600
                     Width           =   460
                  End
                  Begin VB.TextBox txtEagle 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   3
                     Left            =   1350
                     TabIndex        =   123
                     Text            =   "0"
                     Top             =   600
                     Width           =   460
                  End
                  Begin VB.TextBox txtEagle 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   2
                     Left            =   900
                     TabIndex        =   122
                     Text            =   "0"
                     Top             =   600
                     Width           =   460
                  End
                  Begin VB.TextBox txtEagle 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   1
                     Left            =   450
                     TabIndex        =   121
                     Text            =   "0"
                     Top             =   600
                     Width           =   460
                  End
                  Begin VB.TextBox txtEagle 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   0
                     Left            =   0
                     TabIndex        =   120
                     Text            =   "0"
                     Top             =   600
                     Width           =   460
                  End
                  Begin VB.TextBox txtPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   0
                     Left            =   -10
                     TabIndex        =   102
                     Text            =   "0"
                     Top             =   105
                     Width           =   460
                  End
                  Begin VB.TextBox txtiHDCP 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   0
                     Left            =   -10
                     TabIndex        =   101
                     Text            =   "0"
                     Top             =   345
                     Width           =   460
                  End
                  Begin VB.TextBox txtPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   1
                     Left            =   440
                     TabIndex        =   100
                     Text            =   "0"
                     Top             =   105
                     Width           =   460
                  End
                  Begin VB.TextBox txtiHDCP 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   1
                     Left            =   440
                     TabIndex        =   99
                     Text            =   "0"
                     Top             =   345
                     Width           =   460
                  End
                  Begin VB.TextBox txtPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   2
                     Left            =   890
                     TabIndex        =   98
                     Text            =   "0"
                     Top             =   105
                     Width           =   460
                  End
                  Begin VB.TextBox txtiHDCP 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   2
                     Left            =   890
                     TabIndex        =   97
                     Text            =   "0"
                     Top             =   345
                     Width           =   460
                  End
                  Begin VB.TextBox txtPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   3
                     Left            =   1340
                     TabIndex        =   96
                     Text            =   "0"
                     Top             =   105
                     Width           =   460
                  End
                  Begin VB.TextBox txtiHDCP 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   3
                     Left            =   1340
                     TabIndex        =   95
                     Text            =   "0"
                     Top             =   345
                     Width           =   460
                  End
                  Begin VB.TextBox txtPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   4
                     Left            =   1780
                     TabIndex        =   94
                     Text            =   "0"
                     Top             =   105
                     Width           =   460
                  End
                  Begin VB.TextBox txtiHDCP 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   4
                     Left            =   1780
                     TabIndex        =   93
                     Text            =   "0"
                     Top             =   345
                     Width           =   460
                  End
                  Begin VB.TextBox txtPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   5
                     Left            =   2240
                     TabIndex        =   92
                     Text            =   "0"
                     Top             =   105
                     Width           =   460
                  End
                  Begin VB.TextBox txtiHDCP 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   5
                     Left            =   2240
                     TabIndex        =   91
                     Text            =   "0"
                     Top             =   345
                     Width           =   460
                  End
                  Begin VB.TextBox txtPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   6
                     Left            =   2680
                     TabIndex        =   90
                     Text            =   "0"
                     Top             =   105
                     Width           =   460
                  End
                  Begin VB.TextBox txtiHDCP 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   6
                     Left            =   2680
                     TabIndex        =   89
                     Text            =   "0"
                     Top             =   345
                     Width           =   460
                  End
                  Begin VB.TextBox txtPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   7
                     Left            =   3130
                     TabIndex        =   88
                     Text            =   "0"
                     Top             =   105
                     Width           =   460
                  End
                  Begin VB.TextBox txtiHDCP 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   7
                     Left            =   3130
                     TabIndex        =   87
                     Text            =   "0"
                     Top             =   345
                     Width           =   460
                  End
                  Begin VB.TextBox txtPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   8
                     Left            =   3580
                     TabIndex        =   86
                     Text            =   "0"
                     Top             =   105
                     Width           =   460
                  End
                  Begin VB.TextBox txtiHDCP 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   8
                     Left            =   3580
                     TabIndex        =   85
                     Text            =   "0"
                     Top             =   345
                     Width           =   460
                  End
                  Begin VB.TextBox txtPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   9
                     Left            =   4590
                     TabIndex        =   84
                     Text            =   "0"
                     Top             =   105
                     Width           =   460
                  End
                  Begin VB.TextBox txtiHDCP 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   9
                     Left            =   4590
                     TabIndex        =   83
                     Text            =   "0"
                     Top             =   345
                     Width           =   460
                  End
                  Begin VB.TextBox txtPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   10
                     Left            =   5040
                     TabIndex        =   82
                     Text            =   "0"
                     Top             =   105
                     Width           =   460
                  End
                  Begin VB.TextBox txtiHDCP 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   10
                     Left            =   5040
                     TabIndex        =   81
                     Text            =   "0"
                     Top             =   345
                     Width           =   460
                  End
                  Begin VB.TextBox txtPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   11
                     Left            =   5490
                     TabIndex        =   80
                     Text            =   "0"
                     Top             =   105
                     Width           =   460
                  End
                  Begin VB.TextBox txtiHDCP 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   11
                     Left            =   5490
                     TabIndex        =   79
                     Text            =   "0"
                     Top             =   345
                     Width           =   460
                  End
                  Begin VB.TextBox txtPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   12
                     Left            =   5940
                     TabIndex        =   78
                     Text            =   "0"
                     Top             =   105
                     Width           =   460
                  End
                  Begin VB.TextBox txtiHDCP 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   12
                     Left            =   5940
                     TabIndex        =   77
                     Text            =   "0"
                     Top             =   345
                     Width           =   460
                  End
                  Begin VB.TextBox txtPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   13
                     Left            =   6390
                     TabIndex        =   76
                     Text            =   "0"
                     Top             =   105
                     Width           =   460
                  End
                  Begin VB.TextBox txtiHDCP 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   13
                     Left            =   6390
                     TabIndex        =   75
                     Text            =   "0"
                     Top             =   345
                     Width           =   460
                  End
                  Begin VB.TextBox txtPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   14
                     Left            =   6840
                     TabIndex        =   74
                     Text            =   "0"
                     Top             =   105
                     Width           =   460
                  End
                  Begin VB.TextBox txtiHDCP 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   14
                     Left            =   6840
                     TabIndex        =   73
                     Text            =   "0"
                     Top             =   345
                     Width           =   460
                  End
                  Begin VB.TextBox txtPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   15
                     Left            =   7290
                     TabIndex        =   72
                     Text            =   "0"
                     Top             =   105
                     Width           =   460
                  End
                  Begin VB.TextBox txtiHDCP 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   15
                     Left            =   7290
                     TabIndex        =   71
                     Text            =   "0"
                     Top             =   345
                     Width           =   460
                  End
                  Begin VB.TextBox txtPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   16
                     Left            =   7740
                     TabIndex        =   70
                     Text            =   "0"
                     Top             =   105
                     Width           =   460
                  End
                  Begin VB.TextBox txtiHDCP 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   16
                     Left            =   7740
                     TabIndex        =   69
                     Text            =   "0"
                     Top             =   345
                     Width           =   460
                  End
                  Begin VB.TextBox txtPts 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   17
                     Left            =   8190
                     TabIndex        =   68
                     Text            =   "0"
                     Top             =   105
                     Width           =   460
                  End
                  Begin VB.TextBox txtiHDCP 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Index           =   17
                     Left            =   8190
                     TabIndex        =   67
                     Text            =   "0"
                     Top             =   345
                     Width           =   460
                  End
                  Begin VB.TextBox txtPtsF 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Left            =   4040
                     TabIndex        =   66
                     Text            =   "0"
                     Top             =   105
                     Width           =   570
                  End
                  Begin VB.TextBox txtiHDCPF 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C6B8A4&
                     Height          =   255
                     Left            =   4040
                     TabIndex        =   65
                     Text            =   "0"
                     Top             =   345
                     Width           =   570
                  End
                  Begin VB.TextBox txtPtsB 
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
                     TabIndex        =   64
                     Text            =   "0"
                     Top             =   105
                     Width           =   570
                  End
                  Begin VB.TextBox txtiHDCPB 
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
                     TabIndex        =   63
                     Text            =   "0"
                     Top             =   345
                     Width           =   570
                  End
                  Begin VB.TextBox txtPtsTot 
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
                     TabIndex        =   62
                     Text            =   "0"
                     Top             =   105
                     Width           =   620
                  End
                  Begin VB.TextBox txtiHDCPTot 
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
                     TabIndex        =   61
                     Text            =   "0"
                     Top             =   345
                     Width           =   620
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
                  TabIndex        =   0
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
                  TabIndex        =   1
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
                  TabIndex        =   2
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
                  TabIndex        =   3
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
                  TabIndex        =   4
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
                  TabIndex        =   5
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
                  TabIndex        =   6
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
                  TabIndex        =   7
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
                  TabIndex        =   8
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
                  TabIndex        =   17
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
                  TabIndex        =   16
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
                  TabIndex        =   15
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
                  TabIndex        =   14
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
                  TabIndex        =   13
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
                  TabIndex        =   12
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
                  TabIndex        =   11
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
                  TabIndex        =   10
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
                  TabIndex        =   9
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
                  TabIndex        =   58
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
                     TabIndex        =   59
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
                  ScaleWidth      =   2775
                  TabIndex        =   55
                  Top             =   -10
                  Width           =   2800
                  Begin VB.TextBox txtNet 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H0000FFFF&
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
                     Left            =   1740
                     TabIndex        =   119
                     Text            =   "0"
                     Top             =   -10
                     Width           =   620
                  End
                  Begin VB.TextBox txtScoreHDCP 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00E0E0E0&
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   255
                     Left            =   1140
                     TabIndex        =   118
                     Text            =   "0"
                     Top             =   -10
                     Width           =   620
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
                     TabIndex        =   57
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
                     ForeColor       =   &H000000FF&
                     Height          =   255
                     Left            =   540
                     TabIndex        =   56
                     Text            =   "0"
                     Top             =   -10
                     Width           =   620
                  End
               End
               Begin VB.Label Label1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " SCORE"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   -15
                  TabIndex        =   105
                  Top             =   -15
                  Width           =   2010
               End
               Begin VB.Label Label2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   -15
                  TabIndex        =   104
                  Top             =   345
                  Width           =   2010
               End
               Begin VB.Label Label3 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   -15
                  TabIndex        =   103
                  Top             =   585
                  Width           =   2010
               End
            End
         End
      End
      Begin RPVGCC.b8Container b8Container4 
         Height          =   1335
         Left            =   120
         TabIndex        =   108
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   2355
         BackColor       =   49152
         Begin VB.PictureBox Picture8 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1095
            Left            =   120
            ScaleHeight     =   1095
            ScaleWidth      =   1815
            TabIndex        =   109
            Top             =   120
            Width           =   1815
            Begin VB.TextBox txtHandicap 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   960
               TabIndex        =   111
               Top             =   240
               Width           =   735
            End
            Begin VB.TextBox txtClass 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   960
               TabIndex        =   110
               Top             =   600
               Width           =   735
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Handicap"
               Height          =   255
               Left            =   120
               TabIndex        =   113
               Top             =   240
               Width           =   975
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
         End
      End
      Begin RPVGCC.b8Container b8Container6 
         Height          =   1575
         Left            =   5880
         TabIndex        =   114
         Top             =   2040
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   2778
         BackColor       =   49152
         Begin VB.PictureBox Picture9 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   120
            ScaleHeight     =   1335
            ScaleWidth      =   4095
            TabIndex        =   115
            Top             =   120
            Width           =   4095
            Begin MSComctlLib.ListView lstTeamMates 
               Height          =   1335
               Left            =   0
               TabIndex        =   116
               Top             =   0
               Width           =   4095
               _ExtentX        =   7223
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
                  Object.Width           =   5292
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   2
                  Text            =   "Net Pts"
                  Object.Width           =   1764
               EndProperty
            End
         End
      End
   End
   Begin VB.Timer TimerPrintSummary 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9360
      Top             =   6000
   End
   Begin RPVGCC.b8Container picPrint 
      Height          =   2055
      Left            =   4560
      TabIndex        =   249
      Top             =   1800
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3625
      BackColor       =   15396057
      Begin VB.CommandButton cmdOKPrint 
         Height          =   480
         Left            =   720
         Picture         =   "frmScoreCardsSystem36.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   252
         Top             =   1275
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelPrint 
         Height          =   480
         Left            =   2355
         Picture         =   "frmScoreCardsSystem36.frx":0F3C
         Style           =   1  'Graphical
         TabIndex        =   251
         Top             =   1275
         Width           =   1560
      End
      Begin VB.ComboBox cmbGrossNet 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   250
         Top             =   720
         Width           =   4095
      End
      Begin RPVGCC.b8TitleBar b8TitleBar3 
         Height          =   345
         Left            =   45
         TabIndex        =   253
         Top             =   45
         Width           =   4485
         _ExtentX        =   7911
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
         Icon            =   "frmScoreCardsSystem36.frx":1698
      End
   End
   Begin RPVGCC.b8Container picProgress 
      Height          =   975
      Left            =   4080
      TabIndex        =   254
      Top             =   2400
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
         TabIndex        =   255
         Top             =   120
         Width           =   5295
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   770
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   15000
      TabIndex        =   18
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   570
         Left            =   0
         TabIndex        =   19
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
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   15000
         Y1              =   690
         Y2              =   690
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
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   15000
         Y1              =   660
         Y2              =   660
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   14280
      Top             =   1800
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
            Picture         =   "frmScoreCardsSystem36.frx":1C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardsSystem36.frx":1D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardsSystem36.frx":1EB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardsSystem36.frx":21D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardsSystem36.frx":258B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardsSystem36.frx":29DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardsSystem36.frx":2E2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardsSystem36.frx":31E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardsSystem36.frx":32F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardsSystem36.frx":383B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardsSystem36.frx":3995
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScoreCardsSystem36.frx":3ED7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   117
      Top             =   7860
      Width           =   13335
      _ExtentX        =   23521
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
   Begin RPVGCC.b8Container picSearch 
      Height          =   4695
      Left            =   4440
      TabIndex        =   241
      Top             =   480
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   8281
      BackColor       =   15396057
      Begin VB.CommandButton cmdCancelSearch 
         Height          =   480
         Left            =   2280
         Picture         =   "frmScoreCardsSystem36.frx":40FB
         Style           =   1  'Graphical
         TabIndex        =   246
         Top             =   3960
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKSearch 
         Height          =   480
         Left            =   480
         Picture         =   "frmScoreCardsSystem36.frx":4857
         Style           =   1  'Graphical
         TabIndex        =   245
         Top             =   3960
         Width           =   1560
      End
      Begin VB.ListBox lstResult 
         Height          =   2595
         Left            =   120
         TabIndex        =   244
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   243
         Top             =   480
         Width           =   4095
      End
      Begin VB.ComboBox cmbDate 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   242
         Top             =   3480
         Width           =   1695
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   247
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
         Icon            =   "frmScoreCardsSystem36.frx":4EC9
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   255
         Left            =   1200
         TabIndex        =   248
         Top             =   3480
         Width           =   495
      End
   End
   Begin RPVGCC.b8Container picSearchAdd 
      Height          =   4695
      Left            =   4440
      TabIndex        =   233
      Top             =   480
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   8281
      BackColor       =   15396057
      Begin VB.TextBox txtSearchAdd 
         Height          =   315
         Left            =   120
         TabIndex        =   238
         Top             =   480
         Width           =   4095
      End
      Begin VB.ListBox lstResultAdd 
         Height          =   2595
         Left            =   120
         TabIndex        =   237
         Top             =   840
         Width           =   4095
      End
      Begin VB.CommandButton cmdOKAdd 
         Height          =   480
         Left            =   480
         Picture         =   "frmScoreCardsSystem36.frx":5463
         Style           =   1  'Graphical
         TabIndex        =   236
         Top             =   3960
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelAdd 
         Height          =   480
         Left            =   2280
         Picture         =   "frmScoreCardsSystem36.frx":5AD5
         Style           =   1  'Graphical
         TabIndex        =   235
         Top             =   3960
         Width           =   1560
      End
      Begin VB.TextBox txtDateAdd 
         Height          =   315
         Left            =   1800
         TabIndex        =   234
         Top             =   3480
         Width           =   1215
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   45
         TabIndex        =   239
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
         Icon            =   "frmScoreCardsSystem36.frx":6231
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   255
         Left            =   1320
         TabIndex        =   240
         Top             =   3480
         Width           =   495
      End
   End
   Begin MSComctlLib.ListView lstScoreResult 
      Height          =   1095
      Left            =   120
      TabIndex        =   256
      Top             =   6240
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   1931
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PlayerName"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Pts"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Class"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "1 : Gross 2 : Net"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Eagle"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Birdie"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Par"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmScoreCardsSystem36"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ReportClass      As String

Public TourNoOfPlays    As Double
Dim PlayerKey           As Double
Dim s                   As String
Dim rs                  As New ADODB.Recordset
Dim t                   As String
Dim rt                  As New ADODB.Recordset
Dim u                   As String
Dim ru                  As New ADODB.Recordset


Dim i, x, j, TeamTmp, dDateEnd
Dim TRANSACTIONTYPE     As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Private Sub BROWSER(sCtrl, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If sCtrl <> "" Then
            s = "SELECT TOP 1 tbl_Scoring_ScoreCard_System36.PK, " & _
                " tbl_Scoring_ScoreCard_System36.CtrlNo, " & _
                " tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
                " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
                " tbl_Scoring_ScoreCard_System36.DDate, " & _
                " tbl_Scoring_ScoreCard_System36.LastModified " & _
                " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
                " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
                " AND (tbl_Scoring_ScoreCard_System36.CtrlNo = '" & sCtrl & "') " & _
                " ORDER BY tbl_Scoring_ScoreCard_System36.CtrlNo"
        Else
            s = "SELECT TOP 1 tbl_Scoring_ScoreCard_System36.PK, " & _
                " tbl_Scoring_ScoreCard_System36.CtrlNo, " & _
                " tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
                " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
                " tbl_Scoring_ScoreCard_System36.DDate, " & _
                " tbl_Scoring_ScoreCard_System36.LastModified " & _
                " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
                " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
                " ORDER BY tbl_Scoring_ScoreCard_System36.CtrlNo"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Scoring_ScoreCard_System36.PK, " & _
            " tbl_Scoring_ScoreCard_System36.CtrlNo, " & _
            " tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
            " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
            " tbl_Scoring_ScoreCard_System36.DDate, " & _
            " tbl_Scoring_ScoreCard_System36.LastModified " & _
            " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
            " ORDER BY tbl_Scoring_ScoreCard_System36.CtrlNo"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Scoring_ScoreCard_System36.PK, " & _
            " tbl_Scoring_ScoreCard_System36.CtrlNo, " & _
            " tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
            " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
            " tbl_Scoring_ScoreCard_System36.DDate, " & _
            " tbl_Scoring_ScoreCard_System36.LastModified " & _
            " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
            " AND (tbl_Scoring_ScoreCard_System36.CtrlNo < '" & sCtrl & "') " & _
            " ORDER BY tbl_Scoring_ScoreCard_System36.CtrlNo DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Scoring_ScoreCard_System36.PK, " & _
            " tbl_Scoring_ScoreCard_System36.CtrlNo, " & _
            " tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
            " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
            " tbl_Scoring_ScoreCard_System36.DDate, " & _
            " tbl_Scoring_ScoreCard_System36.LastModified " & _
            " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
            " AND (tbl_Scoring_ScoreCard_System36.CtrlNo > '" & sCtrl & "') " & _
            " ORDER BY tbl_Scoring_ScoreCard_System36.CtrlNo "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Scoring_ScoreCard_System36.PK, " & _
            " tbl_Scoring_ScoreCard_System36.CtrlNo, " & _
            " tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
            " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
            " tbl_Scoring_ScoreCard_System36.DDate, " & _
            " tbl_Scoring_ScoreCard_System36.LastModified " & _
            " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
            " ORDER BY tbl_Scoring_ScoreCard_System36.CtrlNo DESC"
    Case "is_FIND"
        s = "SELECT TOP 1 tbl_Scoring_ScoreCard_System36.PK, " & _
            " tbl_Scoring_ScoreCard_System36.CtrlNo, " & _
            " tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
            " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
            " tbl_Scoring_ScoreCard_System36.DDate, " & _
            " tbl_Scoring_ScoreCard_System36.LastModified " & _
            " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
            " AND (tbl_Scoring_ScoreCard_System36.PK = " & sCtrl & ") " & _
            " ORDER BY tbl_Scoring_ScoreCard_System36.CtrlNo "
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then

    PlayerKey = rs!PlayerKey
    txtCtrl.Text = rs!CtrlNo
    txtPlayer.Text = rs!PlayerName
    txtDate.Text = Format(rs!dDate, "mm/dd/yyyy")
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    
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
        't = "SELECT tbl_Scoring_PlayerName.LastName, tbl_Scoring_PlayerName.FirstName, " & _
            " tbl_Scoring_PlayerName.MiddleName, " & _
            " ISNULL((SELECT tbl_Scoring_ScoreCard_System36.GrossPoints " & _
            " From tbl_Scoring_ScoreCard_System36 " & _
            " WHERE (tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_ScoreCard_System36_Detail.PlayerKey)),0) AS GrossPts " & _
            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " Where (tbl_Scoring_Team_Detail.TeamKey = " & TeamTmp & ") " & _
            " And (tbl_Scoring_Team_Detail.PlayerKey <> " & rs!PlayerKey & ") " & _
            " Order By ISNULL((SELECT tbl_Scoring_ScoreCard_System36.GrossPoints " & _
            " From tbl_Scoring_ScoreCard_System36 " & _
            " WHERE (tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)),0) DESC"
        
        t = "SELECT tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
            " (SELECT SUM(tbl_Scoring_ScoreCard_System36.NetPoints) AS NetPoints " & _
            " From tbl_Scoring_ScoreCard_System36 " & _
            " WHERE (tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)) AS NetPts " & _
            " FROM tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
            " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
            " Where (tbl_Scoring_Team_Detail.TeamKey = " & TeamTmp & ") " & _
            " And (tbl_Scoring_Team_Detail.PlayerKey <> " & rs!PlayerKey & ") " & _
            " Order By (SELECT SUM(tbl_Scoring_ScoreCard_System36.NetPoints) AS NetPoints " & _
            " From tbl_Scoring_ScoreCard_System36 " & _
            " WHERE (tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)) DESC"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        While Not rt.EOF
            Set x = lstTeamMates.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rt!PlayerName 'Trim(rt!LastName) & ",  " & Trim(rt!FirstName) & IIf(Trim(rt!MiddleName) = "", "", "  " & rt!MiddleName)
            x.SubItems(2) = IIf(IsNull(rt!NetPts), 0, rt!NetPts)
            rt.MoveNext
        Wend
        rt.Close
    End If
    
    i = -1
    t = "SELECT ScoreCardKey, Hole, Par, Score " & _
        " From tbl_Scoring_ScoreCard_System36_Detail " & _
        " Where (ScoreCardKey = " & rs!PK & ") " & _
        " ORDER BY Hole"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        DoEvents
        i = i + 1
        txtGrossScore(i).Text = rt!Score
        rt.MoveNext
    Wend
    rt.Close
    
    SaveSetting App.EXEName, "ScoreCardControl36", "ScoreCardCtrl36", rs!CtrlNo
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSearchAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If picProgress.Visible = True Then Exit Sub
If AccessRights("Scoring Score Card", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
picSearchAdd.ZOrder 0
txtSearchAdd.Text = ""
txtDateAdd.Text = Format(Date, "mm/dd/yyyy")
picSearchAdd.Visible = True
txtSearchAdd.SetFocus
End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If picSearchAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If picProgress.Visible = True Then Exit Sub
If AccessRights("Scoring Score Card", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_EDITTING
txtGrossScore(0).SetFocus
End Sub

Private Sub PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If picSearchAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If picProgress.Visible = True Then Exit Sub
If AccessRights("Scoring Score Card", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If CHECK_TOURNAMENT_STATUS(TournamentKey) <> 0 Then MsgBox "Tournament was already locked!               ", vbCritical, "Error...": Exit Sub
If MsgBox("ARE YOU SURE IN DELETING THIS RECORD?                    ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_System36 WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "ScoreCardControl36", "ScoreCardCtrl36", ""), "is_PAGEDOWN"
If Trim(txtPlayer.Text) = "" Then BROWSER GetSetting(App.EXEName, "ScoreCardControl36", "ScoreCardCtrl36", ""), "is_HOME"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
Dim TourNoOfPlaysTmp, SCardKey, dblPar, dblHandicap, _
dblScore, dblGross, dblNet, strCtrlNo

If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
If picSearchAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If picProgress.Visible = True Then Exit Sub

On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    s = "SELECT COUNT(*) AS NoofRec " & _
        " From tbl_Scoring_ScoreCard_System36 " & _
        " WHERE (TournamentKey = " & TournamentKey & ") " & _
        " AND (PlayerKey = " & PlayerKey & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    TourNoOfPlaysTmp = rs!NoofRec
    rs.Close
    
    If CDbl(TourNoOfPlaysTmp) + 1 > CDbl(DaysPlayerToPlay) Then MsgBox "Number of Plays Exceeded!                  ", vbCritical, "Error...": Exit Sub
    
    strCtrlNo = "00000001"
    s = "SELECT TOP 1 CtrlNo " & _
        " FROM tbl_Scoring_ScoreCard_System36 " & _
        " WHERE (TournamentKey = " & TournamentKey & ") " & _
        " ORDER BY CtrlNo DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        strCtrlNo = Format(CDbl(rs!CtrlNo) + 1, "0000000#")
    End If
    rs.Close
    
    Do
        s = "SELECT tbl_Scoring_ScoreCard_System36.* " & _
            " FROM tbl_Scoring_ScoreCard_System36 " & _
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
    
    ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_System36 " & _
                      " (CtrlNo, TournamentKey, PlayerKey, DDate, Front9Gross, Back9Gross, " & _
                      " Front9Net, Back9Net, HDCP, Eagle, Birdie, Par, Boogie, Boogie_2, Boogie_3, LastModified) " & _
                      " VALUES ('" & strCtrlNo & "', " & TournamentKey & ", " & PlayerKey & ", " & _
                      " '" & FormatDateTime(txtDate.Text, vbShortDate) & "', " & RETURNTEXTVALUE(txtSGrossF) & ", " & _
                      " " & RETURNTEXTVALUE(txtSGrossB) & ", " & RETURNTEXTVALUE(txtSNetF) & ", " & _
                      " " & RETURNTEXTVALUE(txtSNetB) & ", " & RETURNTEXTVALUE(txtScoreHDCP) & ", " & _
                      " " & lstScores.ListItems.Item(1).SubItems(2) & ", " & lstScores.ListItems.Item(2).SubItems(2) & ", " & _
                      " " & lstScores.ListItems.Item(3).SubItems(2) & ", " & lstScores.ListItems.Item(4).SubItems(2) & ", " & _
                      " " & lstScores.ListItems.Item(5).SubItems(2) & ", " & lstScores.ListItems.Item(6).SubItems(2) & ", " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
    
    ConnOmega.Execute "UPDATE tbl_Scoring_PlayerName " & _
                      " SET HandiCap = " & RETURNTEXTVALUE(txtScoreHDCP) & " " & _
                      " WHERE (PK = " & PlayerKey & ")"
    SCardKey = 0
    s = "SELECT PK " & _
        " FROM tbl_Scoring_ScoreCard_System36 " & _
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
        ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_System36_Detail WHERE (ScoreCardKey = " & SCardKey & ")"
        j = 0
        For i = 1 To 18
            With FGrid
                j = j + 1
                If j >= 1 And j <= 9 Then
                    dblPar = .TextMatrix(5, i + 1)
                Else
                    dblPar = .TextMatrix(5, i + 2)
                End If
            End With
            
            dblScore = RETURNTEXTVALUE(txtGrossScore(i - 1))
            
            ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_System36_Detail " & _
                              " (ScoreCardKey, Hole, Par, Score) " & _
                              " VALUES (" & SCardKey & ", " & i & ", " & CDbl(dblPar) & ", " & _
                              " " & CDbl(dblScore) & ")"
        Next i
    End If
    
End If

If TRANSACTIONTYPE = is_EDITTING Then

    strCtrlNo = txtCtrl.Text
    SCardKey = Statusbar1.Panels(1).Text
    
    ConnOmega.Execute "UPDATE tbl_Scoring_ScoreCard_System36 " & _
                      " SET Front9Gross = " & RETURNTEXTVALUE(txtSGrossF) & ", " & _
                      " Back9Gross = " & RETURNTEXTVALUE(txtSGrossB) & ", " & _
                      " Front9Net = " & RETURNTEXTVALUE(txtSNetF) & ", " & _
                      " Back9Net = " & RETURNTEXTVALUE(txtSNetB) & ", " & _
                      " HDCP = " & RETURNTEXTVALUE(txtScoreHDCP) & ", " & _
                      " Eagle = " & lstScores.ListItems.Item(1).SubItems(2) & ", " & _
                      " Birdie = " & lstScores.ListItems.Item(2).SubItems(2) & ", " & _
                      " Par =  " & lstScores.ListItems.Item(3).SubItems(2) & ", " & _
                      " Boogie = " & lstScores.ListItems.Item(4).SubItems(2) & ", " & _
                      " Boogie_2 = " & lstScores.ListItems.Item(5).SubItems(2) & ", " & _
                      " Boogie_3 = " & lstScores.ListItems.Item(6).SubItems(2) & ", " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & SCardKey & ")"
    
    ConnOmega.Execute "UPDATE tbl_Scoring_PlayerName " & _
                      " SET HandiCap = " & RETURNTEXTVALUE(txtScoreHDCP) & " " & _
                      " WHERE (PK = " & PlayerKey & ")"
                      
    If CDbl(SCardKey) <> 0 Then
        ConnOmega.Execute "DELETE FROM tbl_Scoring_ScoreCard_System36_Detail WHERE (ScoreCardKey = " & SCardKey & ")"
        j = 0
        For i = 1 To 18
            With FGrid
                j = j + 1
                If j >= 1 And j <= 9 Then
                    dblPar = .TextMatrix(5, i + 1)
                Else
                    dblPar = .TextMatrix(5, i + 2)
                End If
            End With
            
            dblScore = RETURNTEXTVALUE(txtGrossScore(i - 1))
            
            ConnOmega.Execute "INSERT INTO tbl_Scoring_ScoreCard_System36_Detail " & _
                              " (ScoreCardKey, Hole, Par, Score) " & _
                              " VALUES (" & SCardKey & ", " & i & ", " & CDbl(dblPar) & ", " & _
                              " " & CDbl(dblScore) & ")"
        Next i
    End If
End If

If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER strCtrlNo, "is_LOAD"
End If

Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSearchAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If picProgress.Visible = True Then Exit Sub
picSearch.ZOrder 0
txtSearch.Text = ""
picSearch.Visible = True
txtSearch.SetFocus
End Sub

Private Sub PRESS_F9()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSearchAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If picProgress.Visible = True Then Exit Sub
MainFormPopupF.mnuPrintSystem36Class(0).Caption = "ALL"
For i = 1 To MainFormPopupF.mnuPrintSystem36Class.UBound
    Unload MainFormPopupF.mnuPrintSystem36Class(i)
Next i

i = 0
i = i + 1
Load MainFormPopupF.mnuPrintSystem36Class(i)
MainFormPopupF.mnuPrintSystem36Class(i).Caption = "-"

s = "SELECT Class " & _
    " From tbl_Scoring_TournamentInfo_Class " & _
    " Where (TournamentKey = " & TournamentKey & ") " & _
    " ORDER BY Class"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    i = i + 1
    Load MainFormPopupF.mnuPrintSystem36Class(i)
    MainFormPopupF.mnuPrintSystem36Class(i).Caption = "CLASS " & rs!Class
    rs.MoveNext
Wend
rs.Close
PopupMenu MainFormPopupF.mnuPrintSystem36, , 5800, 500
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSearchAdd.Visible = True Then cmdCancelAdd_Click: Exit Sub
    If picSearch.Visible = True Then cmdCancelSearch_Click: Exit Sub
    If picPrint.Visible = True Then cmdCancelPrint_Click: Exit Sub
    If picProgress.Visible = True Then Exit Sub
    Unload Me
Else
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "ScoreCardControl36", "ScoreCardCtrl36", ""), "is_LOAD"
    If Trim(txtPlayer.Text) = "" Then BROWSER GetSetting(App.EXEName, "ScoreCardControl36", "ScoreCardCtrl36", ""), "is_HOME"
End If
End Sub

Private Sub CLEARTEXT()
PlayerKey = 0
txtCtrl.Text = ""
txtPlayer.Text = ""
txtDate.Text = ""
txtHandicap.Text = ""
txtClass.Text = ""
txtGrossScoreF.Text = ""
txtGrossScoreB.Text = ""
lstTeamMates.ListItems.Clear
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
For i = 1 To 6
    lstScores.ListItems.Item(i).SubItems(2) = "0"
Next i
For i = 0 To 17
    txtGrossScore(i).Text = ""
    txtEagle(i).Text = "0"
    txtBirdies(i).Text = "0"
    txtPars(i).Text = "0"
    txtBoogies(i).Text = "0"
    txt2Boogies(i).Text = "0"
    txt3Boogies(i).Text = "0"
Next i
End Sub

Private Sub LOCKTEXT(bln As Boolean)
For i = 0 To 17
    txtGrossScore(i).Locked = bln
Next i
End Sub

Private Sub TOOLBARFUNC(intSel As Integer)
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
'        HEADER1$ = HEADER1$ & "|" & "IN" & "|" & "GROSS" & "|" & "HDCP" & "|" & "NET"
'        .FormatString = HEADER1$
'        For i = 1 To .Cols - 1
'            If i = 1 Then
'                .ColWidth(i) = 2000
'                .ColAlignment(i) = 1
'            ElseIf i = 11 Or _
'            i = 21 Then
'                .ColWidth(i) = 550
'                .ColAlignment(i) = flexAlignRightCenter
'            ElseIf i = 22 Or _
'            i = 23 Or _
'            i = 24 Then
'                .ColWidth(i) = 600
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
''Handicap
''i = 1
''GRow = GRow + 1
''a = 0: Tot1 = 0: Tot2 = 0: Tot = 0
''s = "SELECT Hole, Par, HandicapIndex, " & _
''    " Gold, Blue, White, Red " & _
''    " From tbl_Scoring_Yardage_Par_HandicapIndex " & _
''    " ORDER BY Hole"
''If rs.State = adStateOpen Then rs.Close
''rs.Open s, ConnOmega
''If rs.RecordCount > 0 Then
''    With FGrid
''        rs.MoveFirst
''        .Rows = .Rows + 1
''        .TextMatrix(GRow, i) = "HANDICAP"
''        While Not rs.EOF
''            i = i + 1
''            If a = 9 Then
''                .TextMatrix(GRow, i) = ""
''                i = i + 1
''                .TextMatrix(GRow, i) = rs!HandicapIndex
''                Tot1 = Tot
''                a = 0
''                Tot = 0
''            Else
''                .TextMatrix(GRow, i) = rs!HandicapIndex
''            End If
''            Tot = Tot + CDbl(rs!Par)
''            a = a + 1
''            rs.MoveNext
''        Wend
''        Tot2 = CDbl(Tot) + CDbl(Tot1)
''        i = i + 1
''        .TextMatrix(GRow, i) = ""
''        i = i + 1
''        .TextMatrix(GRow, i) = ""
''        'i = i + 1
''        '.TextMatrix(GRow, i) = ""
''    End With
''End If
''rs.Close
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
''For i = 1 To 6
''    Set x = lstScores.ListItems.Add()
''    x.Text = ""
''    Select Case i
''        Case 1
''            x.SubItems(1) = "EAGLES"
''        Case 2
''            x.SubItems(1) = "BIRDIES"
''        Case 3
''            x.SubItems(1) = "PARS"
''        Case 4
''            x.SubItems(1) = "BOGEYS"
''        Case 5
''            x.SubItems(1) = "DOUBLE BOGEYS"
''        Case 6
''            x.SubItems(1) = "TRIPLE BOGEYS"
''    End Select
''    x.SubItems(2) = "0"
''Next i
'
'End Sub


Private Sub b8TitleBar3_CLoseClick()
cmdCancelPrint_Click
End Sub

Private Sub cmbDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKSearch_Click
End Sub

Private Sub cmdCancelAdd_Click()
picToolbar.Enabled = True
picMain.Enabled = True
picSearchAdd.Visible = False
End Sub

Private Sub cmdCancelPrint_Click()
picPrint.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True
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
    " From tbl_Scoring_ScoreCard_System36 " & _
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

cmdCancelAdd_Click
txtGrossScore(0).SetFocus

End Sub

Private Sub cmdOKPrint_Click()
If cmbGrossNet.ListIndex = -1 Then Exit Sub
Dim TotalRec, RowCnt, ColCnt, iWorkSheet
Dim WorkbookName, Filename, cnt, ResetCnt, strRange
Dim dblClassFrom, dblClassTo, iWinner, dblEagle, dblBirdie, dblPar
With MainForm.CommonDialog1
    .CancelError = True
    On Error GoTo ErrorHandler
    .DialogTitle = "Save"
    .Filter = "Excel(*.xls)|*.xls"
    .ShowSave
    Filename = Trim(.Filename)
End With

On Error GoTo PG:

picPrint.Visible = False
picProgressBar.BackColor = &HFFFFFF
picProgress.ZOrder 0
picProgress.Visible = True

WorkbookName = CStr(Filename)
'lstScoreResult.ListItems.Clear
RowCnt = 0
iWorkSheet = 1
Set xlsApp = CreateObject("Excel.Application")
xlsApp.Visible = False
xlsApp.Workbooks.Add
xlsApp.DisplayAlerts = False
xlsApp.Workbooks(1).Sheets(2).Delete
xlsApp.Workbooks(1).Sheets(2).Delete
xlsApp.Workbooks(1).Sheets(iWorkSheet).Activate
xlsApp.Workbooks(1).Sheets(iWorkSheet).Name = ReportClass
RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
    .Range(strRange).Value = TournamentName
    .Range(strRange).Font.Name = "Arial"
    .Range(strRange).Font.Size = 14
    .Range(strRange).Font.Bold = True
    
    RowCnt = RowCnt + 1
    ColCnt = 0
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Date : " & TournamentRange
    .Range(strRange).Font.Name = "Arial"
    .Range(strRange).Font.Size = 10
    .Range(strRange).Font.Bold = True
    
    RowCnt = RowCnt + 1
    ColCnt = 0
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = ""
    .Range(strRange).Font.Name = "Arial"
    .Range(strRange).Font.Size = 10
    .Range(strRange).Font.Bold = True

    RowCnt = RowCnt + 1
    ColCnt = 0
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
'    If cmbGrossNet.ListIndex = 0 Then
'        .Range(strRange).Value = "Individual (Net Points) [" & IIf(cmbGender.ListIndex = 1, "MALE", IIf(cmbGender.ListIndex = 2, "FEMALE", "")) & "]"
'    Else
'        .Range(strRange).Value = "Individual [Class " & cmbDivision.List(cmbDivision.ListIndex) & "] (Net Points) [" & IIf(cmbGender.ListIndex = 0, "MALE", "FEMALE") & "]"
'    End If
    .Range(strRange).Value = cmbGrossNet.Text
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Range(strRange).Font.Bold = False
    
    RowCnt = RowCnt + 1
    ColCnt = 0
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = ""
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Range(strRange).Font.Bold = False
    
    RowCnt = RowCnt + 1
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
    .Columns(ColCnt).ColumnWidth = 20
    .Range(strRange).Font.Bold = True
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Gross Pts"
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Range(strRange).Font.Bold = True
    .Range(strRange).HorizontalAlignment = 4
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Handicap"
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Range(strRange).Font.Bold = True
    .Range(strRange).HorizontalAlignment = 4
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Net Pts"
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Range(strRange).Font.Bold = True
    .Range(strRange).HorizontalAlignment = 4
    
'    ColCnt = ColCnt + 1
'    strRange = EXCEL_RANGE(ColCnt, RowCnt)
'    .Range(strRange).Value = ""
'    .Range(strRange).Font.Name = "Tahoma"
'    .Range(strRange).Font.Size = 8
'    .Range(strRange).Font.Bold = True
'    .Range(strRange).ColumnWidth = 1
'    .Range(strRange).HorizontalAlignment = 4
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Eagle"
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Range(strRange).Font.Bold = True
    .Range(strRange).HorizontalAlignment = 4
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Birdie"
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Range(strRange).Font.Bold = True
    .Range(strRange).HorizontalAlignment = 4
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Par"
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Range(strRange).Font.Bold = True
    .Range(strRange).HorizontalAlignment = 4
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Boogie"
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Range(strRange).Font.Bold = True
    .Range(strRange).HorizontalAlignment = 4
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Double Boogie"
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Range(strRange).ColumnWidth = 10
    .Range(strRange).Font.Bold = True
    .Range(strRange).HorizontalAlignment = 4
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Triple Boogie"
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Range(strRange).ColumnWidth = 10
    .Range(strRange).Font.Bold = True
    .Range(strRange).HorizontalAlignment = 4
    i = 0
    Select Case ReportClass
        Case "ALL"
            s = "SELECT tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
                " tbl_Scoring_ScoreCard_System36.NetPoints, tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_ScoreCard_System36.GrossPoints, " & _
                " tbl_Scoring_ScoreCard_System36.Eagle, tbl_Scoring_ScoreCard_System36.Birdie, tbl_Scoring_ScoreCard_System36.Par, " & _
                " tbl_Scoring_ScoreCard_System36.Boogie, tbl_Scoring_ScoreCard_System36.Boogie_2, tbl_Scoring_ScoreCard_System36.Boogie_3 " & _
                " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
                " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                " Where (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
                " ORDER BY tbl_Scoring_ScoreCard_System36.NetPoints, tbl_Scoring_ScoreCard_System36.Eagle DESC, tbl_Scoring_ScoreCard_System36.Birdie DESC, " & _
                " tbl_Scoring_ScoreCard_System36.Par DESC, tbl_Scoring_ScoreCard_System36.Boogie DESC, tbl_Scoring_ScoreCard_System36.Boogie_2 DESC, " & _
                " tbl_Scoring_ScoreCard_System36.Boogie_3 DESC"
        Case Else
            s = "SELECT tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
                " tbl_Scoring_ScoreCard_System36.NetPoints, tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_ScoreCard_System36.GrossPoints, " & _
                " tbl_Scoring_ScoreCard_System36.Eagle, tbl_Scoring_ScoreCard_System36.Birdie, tbl_Scoring_ScoreCard_System36.Par, " & _
                " tbl_Scoring_ScoreCard_System36.Boogie , tbl_Scoring_ScoreCard_System36.Boogie_2, tbl_Scoring_ScoreCard_System36.Boogie_3 " & _
                " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
                " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                " WHERE ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS tbl_Scoring_TournamentInfo_Class_1 " & _
                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_System36.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND " & _
                " (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & Replace(CStr(ReportClass), "CLASS ", "") & "') " & _
                " AND (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
                " ORDER BY tbl_Scoring_ScoreCard_System36.NetPoints, tbl_Scoring_ScoreCard_System36.Eagle DESC, tbl_Scoring_ScoreCard_System36.Birdie DESC, " & _
                " tbl_Scoring_ScoreCard_System36.Par DESC, tbl_Scoring_ScoreCard_System36.Boogie DESC, tbl_Scoring_ScoreCard_System36.Boogie_2 DESC, " & _
                " tbl_Scoring_ScoreCard_System36.Boogie_3 DESC"
    End Select
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        i = i + 1
        RowCnt = RowCnt + 1
        ColCnt = 0
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = i
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 8
        .Range(strRange).Font.Bold = False
        .Columns(ColCnt).ColumnWidth = 3
        .Range(strRange).HorizontalAlignment = 4
        
        For j = 0 To rs.Fields.Count - 1
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = rs.Fields(j).Value
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = False
            '.Range(strRange).HorizontalAlignment = 4
        Next j
        
        UpdateProgress picProgressBar, i / rs.RecordCount
        rs.MoveNext
    Wend
    rs.Close
End With

picProgress.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True

SAVING:
On Error GoTo err_saving:
If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

xlsApp.Visible = True
        
Exit Sub
PG:
picProgress.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True

MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub

Exit Sub
ErrorHandler:
Exit Sub

Exit Sub
err_saving:
picProgress.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True

MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING:

End Sub

Private Sub cmdOKSearch_Click()
If lstResult.ListIndex = -1 Then Exit Sub
If cmbDate.ListIndex = -1 Then Exit Sub
BROWSER cmbDate.ItemData(cmbDate.ListIndex), "is_FIND"
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
    Case vbKeyF9:       PRESS_F9
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "ScoreCardControl36", "ScoreCardCtrl36", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "ScoreCardControl36", "ScoreCardCtrl36", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "ScoreCardControl36", "ScoreCardCtrl36", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "ScoreCardControl36", "ScoreCardCtrl36", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Height = 6405 '6825
Me.Width = 13425
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
'Me.Caption = "Score Card (System 36)"

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

With cmbGrossNet
    .Clear
    .AddItem "Gross"
    .AddItem "Net"
'    .AddItem "SCORES"
End With

For i = 1 To 6
    Set x = lstScores.ListItems.Add()
    x.Text = ""
    Select Case i
        Case 1
            x.SubItems(1) = "EAGLES"
        Case 2
            x.SubItems(1) = "BIRDIES"
        Case 3
            x.SubItems(1) = "PARS"
        Case 4
            x.SubItems(1) = "BOGEYS"
        Case 5
            x.SubItems(1) = "DOUBLE BOGEYS"
        Case 6
            x.SubItems(1) = "TRIPLE BOGEYS"
    End Select
    x.SubItems(2) = "0"
Next i

LOAD_CARD dDateEnd, FGrid
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "ScoreCardControl36", "ScoreCardCtrl36", ""), "is_LOAD"
If Trim(txtPlayer.Text) = "" Then BROWSER GetSetting(App.EXEName, "ScoreCardControl36", "ScoreCardCtrl36", ""), "is_HOME"
Dim tmp As Long
tmp = SetWindowLong(txtSearchAdd.hwnd, GWL_STYLE, GetWindowLong(txtSearchAdd.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)

End Sub

Private Sub Form_Unload(Cancel As Integer)
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
If picSearchAdd.Visible = True Then Cancel = -1
If picSearch.Visible = True Then Cancel = -1
If picProgress.Visible = True Then Cancel = -1
If picPrint.Visible = True Then Cancel = -1
End Sub

Private Sub lstResult_Click()
If lstResult.ListIndex = -1 Then cmbDate.Clear: Exit Sub
cmbDate.Clear
s = "SELECT PK, DDate " & _
    " From tbl_Scoring_ScoreCard_System36 " & _
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

Private Sub TimerPrintResult_Timer()
Dim TotalRec, RowCnt, ColCnt, iWorkSheet
Dim WorkbookName, Filename, cnt, ResetCnt, strRange
Dim dblClassFrom, dblClassTo, iWinner, dblEagle, dblBirdie, dblPar
TimerPrintResult.Enabled = False
With MainForm.CommonDialog1
    .CancelError = True
    On Error GoTo ErrorHandler
    .DialogTitle = "Save"
    .Filter = "Excel(*.xls)|*.xls"
    .ShowSave
    Filename = Trim(.Filename)
End With
TotalRec = 25: i = 0
On Error GoTo PG:
WorkbookName = Filename
picProgressBar.BackColor = &HFFFFFF
picProgress.ZOrder 0
picMain.Enabled = False
picToolbar.Enabled = False
picProgress.Visible = True

WorkbookName = CStr(Filename)
lstScoreResult.ListItems.Clear
RowCnt = 0
iWorkSheet = 1
Set xlsApp = CreateObject("Excel.Application")
xlsApp.Visible = False
xlsApp.Workbooks.Add
xlsApp.DisplayAlerts = False
xlsApp.Workbooks(1).Sheets(2).Delete
xlsApp.Workbooks(1).Sheets(2).Delete
xlsApp.Workbooks(1).Sheets(iWorkSheet).Activate
xlsApp.Workbooks(1).Sheets(iWorkSheet).Name = "R E S U L T"
RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = TournamentName
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 14
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Date : " & TournamentRange
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "GOLF TOURNAMENT RESULT"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 12
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "NAME OF WINNER"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 12
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 12
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "SCORE"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 12
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

's = ""
'If rs.State = adStateOpen Then rs.Close
'rs.Open s, ConnOmega
'While Not rs.EOF
'
'    rs.MoveNext
'Wend
'rs.Close

ResetCnt = 0
s = "SELECT TOP 4 tbl_Scoring_Team_Detail.TeamKey, " & _
    " ISNULL(SUM(tbl_Scoring_ScoreCard_System36.NetPoints), 0) AS NetPoints " & _
    " FROM tbl_Scoring_Team LEFT OUTER JOIN " & _
    " tbl_Scoring_Team_Detail ON tbl_Scoring_Team.PK = tbl_Scoring_Team_Detail.TeamKey LEFT OUTER JOIN " & _
    " tbl_Scoring_ScoreCard_System36 ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_ScoreCard_System36.PlayerKey " & _
    " Where (tbl_Scoring_Team.TournamentKey = " & TournamentKey & ") " & _
    " GROUP BY tbl_Scoring_Team_Detail.TeamKey " & _
    " Having (IsNull(Sum(tbl_Scoring_ScoreCard_System36.NetPoints), 0) <> 0) " & _
    " ORDER BY ISNULL(SUM(tbl_Scoring_ScoreCard_System36.NetPoints), 0), SUM(tbl_Scoring_ScoreCard_System36.Eagle) DESC, " & _
    " SUM(tbl_Scoring_ScoreCard_System36.Birdie) DESC, SUM(tbl_Scoring_ScoreCard_System36.Par) DESC, " & _
    " SUM(tbl_Scoring_ScoreCard_System36.Boogie) DESC, SUM(tbl_Scoring_ScoreCard_System36.Boogie_2) DESC, " & _
    " SUM(tbl_Scoring_ScoreCard_System36.Boogie_3) DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    i = i + 1
    ResetCnt = ResetCnt + 1
    Select Case ResetCnt
        Case 1
            RowCnt = RowCnt + 1
            ColCnt = 0
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "TEAM NET CHAMPION"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
        Case 2
            RowCnt = RowCnt + 1
            ColCnt = 0
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "TEAM NET 1st RUNNER-UP"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
        Case 3
            RowCnt = RowCnt + 1
            ColCnt = 0
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "TEAM NET 2nd RUNNER-UP"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
        Case 4
            RowCnt = RowCnt + 1
            ColCnt = 0
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "TEAM NET 3rd RUNNER-UP"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
    End Select
    
    cnt = 0
    t = "SELECT tbl_Scoring_Team_Detail.TeamKey, " & _
        " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
        " (SELECT SUM(tbl_Scoring_ScoreCard_System36.NetPoints) AS NetPoints " & _
        " From tbl_Scoring_ScoreCard_System36 " & _
        " WHERE (tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_Team_Detail.PlayerKey)) AS NetPts " & _
        " FROM  tbl_Scoring_Team_Detail LEFT OUTER JOIN " & _
        " tbl_Scoring_PlayerName ON tbl_Scoring_Team_Detail.PlayerKey = tbl_Scoring_PlayerName.PK " & _
        " Where (tbl_Scoring_Team_Detail.TeamKey = " & rs!TeamKey & ") " & _
        " ORDER BY tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        cnt = cnt + 1
        ColCnt = 1
        If CDbl(cnt) = 1 Then
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rt!PlayerName
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rt!NetPts
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            strRange = EXCEL_RANGE(ColCnt + 1, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!NetPoints
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
            
        ElseIf CDbl(cnt) = 2 Then
            RowCnt = RowCnt + 1
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rt!PlayerName
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rt!NetPts
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
        End If
        
        rt.MoveNext
    Wend
    rt.Close
    
    UpdateProgress picProgressBar, i / TotalRec
    rs.MoveNext
Wend
rs.Close

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "OVER-ALL GROSS CHAMPION"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

s = "SELECT tbl_Scoring_ScoreCard_System36.TournamentKey, " & _
    " tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
    " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
    " tbl_Scoring_ScoreCard_System36.GrossPoints " & _
    " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " Where (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
    " ORDER BY tbl_Scoring_ScoreCard_System36.GrossPoints, tbl_Scoring_ScoreCard_System36.Eagle DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Birdie DESC, tbl_Scoring_ScoreCard_System36.Par DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie DESC, tbl_Scoring_ScoreCard_System36.Boogie_2 DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie_3 DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    i = i + 1
    rs.MoveFirst
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!PlayerName
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!GrossPoints
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
    Set x = lstScoreResult.ListItems.Add()
    x.Text = ""
    x.SubItems(1) = rs!PlayerKey
    x.SubItems(2) = rs!GrossPoints
    x.SubItems(4) = 1
    UpdateProgress picProgressBar, i / TotalRec
    
End If
rs.Close

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "OVER-ALL NET CHAMPION"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

s = "SELECT tbl_Scoring_ScoreCard_System36.TournamentKey, " & _
    " tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
    " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
    " tbl_Scoring_ScoreCard_System36.NetPoints " & _
    " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " Where (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
    " ORDER BY tbl_Scoring_ScoreCard_System36.NetPoints, tbl_Scoring_ScoreCard_System36.Eagle DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Birdie DESC, tbl_Scoring_ScoreCard_System36.Par DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie DESC, tbl_Scoring_ScoreCard_System36.Boogie_2 DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie_3 DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
        With lstScoreResult.ListItems
            iWinner = 0
            For j = 1 To .Count
                If CDbl(.Item(j).SubItems(1)) = CDbl(rs!PlayerKey) Then
                    iWinner = 1
                End If
            Next j
        End With
        If CDbl(iWinner) = 0 Then
            i = i + 1
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!PlayerName
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!NetPoints
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
            Set x = lstScoreResult.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rs!PlayerKey
            x.SubItems(2) = rs!NetPoints
            x.SubItems(4) = 2
            UpdateProgress picProgressBar, i / TotalRec
            GoTo a:
        End If
        rs.MoveNext
    Wend
End If
rs.Close

a:

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "CLASS A"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 12
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

s = "SELECT HFrom, HTo " & _
    " From tbl_Scoring_TournamentInfo_Class " & _
    " WHERE (TournamentKey = " & TournamentKey & ") " & _
    " AND (Class = 'A')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    dblClassFrom = rs!HFrom
    dblClassTo = rs!HTo
End If
rs.Close

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "GROSS CHAMPION"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 11
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
s = "SELECT tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
    " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
    " tbl_Scoring_ScoreCard_System36.GrossPoints " & _
    " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap >= " & dblClassFrom & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap <= " & dblClassTo & ") " & _
    " ORDER BY tbl_Scoring_ScoreCard_System36.GrossPoints, tbl_Scoring_ScoreCard_System36.Eagle DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Birdie DESC, tbl_Scoring_ScoreCard_System36.Par DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie DESC, tbl_Scoring_ScoreCard_System36.Boogie_2 DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie_3 DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
        With lstScoreResult.ListItems
            iWinner = 0
            For j = 1 To .Count
                If CDbl(.Item(j).SubItems(1)) = CDbl(rs!PlayerKey) Then
                    iWinner = 1
                End If
            Next j
        End With
        If CDbl(iWinner) = 0 Then
            i = i + 1
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!PlayerName
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!GrossPoints
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
            
'            With lstScoreResult.ListItems
'                For j = 1 To .Count
'                    If .Item(j).SubItems(3) = "A" And _
'                    .Item(j).SubItems(4) = 1 And _
'                    CDbl(.Item(j).SubItems(2)) = CDbl(rs!GrossPoints) Then
'                        dblEagle , dblBirdie, dblPar
'                    End If
'                Next j
'            End With
            
            Set x = lstScoreResult.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rs!PlayerKey
            x.SubItems(2) = rs!GrossPoints
            x.SubItems(3) = "A"
            x.SubItems(4) = 1
            UpdateProgress picProgressBar, i / TotalRec
            GoTo AA:
        End If
        rs.MoveNext
    Wend
End If
rs.Close

AA:
RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "GROSS RUNNER-UP"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 11
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
s = "SELECT tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
    " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
    " tbl_Scoring_ScoreCard_System36.GrossPoints " & _
    " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap >= " & dblClassFrom & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap <= " & dblClassTo & ") " & _
    " ORDER BY tbl_Scoring_ScoreCard_System36.GrossPoints, tbl_Scoring_ScoreCard_System36.Eagle DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Birdie DESC, tbl_Scoring_ScoreCard_System36.Par DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie DESC, tbl_Scoring_ScoreCard_System36.Boogie_2 DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie_3 DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
        With lstScoreResult.ListItems
            iWinner = 0
            For j = 1 To .Count
                If CDbl(.Item(j).SubItems(1)) = CDbl(rs!PlayerKey) Then
                    iWinner = 1
                End If
            Next j
        End With
        If CDbl(iWinner) = 0 Then
            i = i + 1
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!PlayerName
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!GrossPoints
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
            Set x = lstScoreResult.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rs!PlayerKey
            x.SubItems(2) = rs!GrossPoints
            x.SubItems(3) = "A"
            x.SubItems(4) = 2
            UpdateProgress picProgressBar, i / TotalRec
            GoTo AB:
        End If
        rs.MoveNext
    Wend
End If
rs.Close

AB:

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "NET CHAMPION"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 11
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
s = "SELECT tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
    " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
    " tbl_Scoring_ScoreCard_System36.NetPoints " & _
    " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap >= " & dblClassFrom & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap <= " & dblClassTo & ") " & _
    " ORDER BY tbl_Scoring_ScoreCard_System36.NetPoints, tbl_Scoring_ScoreCard_System36.Eagle DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Birdie DESC, tbl_Scoring_ScoreCard_System36.Par DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie DESC, tbl_Scoring_ScoreCard_System36.Boogie_2 DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie_3 DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
        With lstScoreResult.ListItems
            iWinner = 0
            For j = 1 To .Count
                If CDbl(.Item(j).SubItems(1)) = CDbl(rs!PlayerKey) Then
                    iWinner = 1
                End If
            Next j
        End With
        If CDbl(iWinner) = 0 Then
            i = i + 1
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!PlayerName
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!NetPoints
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
            Set x = lstScoreResult.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rs!PlayerKey
            x.SubItems(2) = rs!NetPoints
            x.SubItems(3) = "A"
            x.SubItems(4) = 2
            UpdateProgress picProgressBar, i / TotalRec
            GoTo AC:
        End If
        rs.MoveNext
    Wend
End If
rs.Close

AC:
RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "NET RUNNER-UP"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 11
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
s = "SELECT tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
    " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
    " tbl_Scoring_ScoreCard_System36.NetPoints " & _
    " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap >= " & dblClassFrom & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap <= " & dblClassTo & ") " & _
    " ORDER BY tbl_Scoring_ScoreCard_System36.NetPoints, tbl_Scoring_ScoreCard_System36.Eagle DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Birdie DESC, tbl_Scoring_ScoreCard_System36.Par DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie DESC, tbl_Scoring_ScoreCard_System36.Boogie_2 DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie_3 DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
        With lstScoreResult.ListItems
            iWinner = 0
            For j = 1 To .Count
                If CDbl(.Item(j).SubItems(1)) = CDbl(rs!PlayerKey) Then
                    iWinner = 1
                End If
            Next j
        End With
        If CDbl(iWinner) = 0 Then
            i = i + 1
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!PlayerName
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!NetPoints
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
            Set x = lstScoreResult.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rs!PlayerKey
            x.SubItems(2) = rs!NetPoints
            x.SubItems(3) = "A"
            x.SubItems(4) = 2
            UpdateProgress picProgressBar, i / TotalRec
            GoTo b:
        End If
        rs.MoveNext
    Wend
End If
rs.Close

'============ CLASS B
b:

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "CLASS B"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 12
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

s = "SELECT HFrom, HTo " & _
    " From tbl_Scoring_TournamentInfo_Class " & _
    " WHERE (TournamentKey = " & TournamentKey & ") " & _
    " AND (Class = 'B')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    dblClassFrom = rs!HFrom
    dblClassTo = rs!HTo
End If
rs.Close

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "GROSS CHAMPION"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 11
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
s = "SELECT tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
    " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
    " tbl_Scoring_ScoreCard_System36.GrossPoints " & _
    " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap >= " & dblClassFrom & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap <= " & dblClassTo & ") " & _
    " ORDER BY tbl_Scoring_ScoreCard_System36.GrossPoints, tbl_Scoring_ScoreCard_System36.Eagle DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Birdie DESC, tbl_Scoring_ScoreCard_System36.Par DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie DESC, tbl_Scoring_ScoreCard_System36.Boogie_2 DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie_3 DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
        With lstScoreResult.ListItems
            iWinner = 0
            For j = 1 To .Count
                If CDbl(.Item(j).SubItems(1)) = CDbl(rs!PlayerKey) Then
                    iWinner = 1
                End If
            Next j
        End With
        If CDbl(iWinner) = 0 Then
            i = i + 1
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!PlayerName
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!GrossPoints
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
            Set x = lstScoreResult.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rs!PlayerKey
            x.SubItems(2) = rs!GrossPoints
            x.SubItems(3) = "B"
            x.SubItems(4) = 1
            UpdateProgress picProgressBar, i / TotalRec
            GoTo BA:
        End If
        rs.MoveNext
    Wend
End If
rs.Close

BA:
RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "GROSS RUNNER-UP"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 11
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
s = "SELECT tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
    " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
    " tbl_Scoring_ScoreCard_System36.GrossPoints " & _
    " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap >= " & dblClassFrom & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap <= " & dblClassTo & ") " & _
    " ORDER BY tbl_Scoring_ScoreCard_System36.GrossPoints, tbl_Scoring_ScoreCard_System36.Eagle DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Birdie DESC, tbl_Scoring_ScoreCard_System36.Par DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie DESC, tbl_Scoring_ScoreCard_System36.Boogie_2 DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie_3 DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
        With lstScoreResult.ListItems
            iWinner = 0
            For j = 1 To .Count
                If CDbl(.Item(j).SubItems(1)) = CDbl(rs!PlayerKey) Then
                    iWinner = 1
                End If
            Next j
        End With
        If CDbl(iWinner) = 0 Then
            i = i + 1
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!PlayerName
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!GrossPoints
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
            Set x = lstScoreResult.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rs!PlayerKey
            x.SubItems(2) = rs!GrossPoints
            x.SubItems(3) = "B"
            x.SubItems(4) = 1
            UpdateProgress picProgressBar, i / TotalRec
            GoTo BB:
        End If
        rs.MoveNext
    Wend
End If
rs.Close

BB:

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "NET CHAMPION"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 11
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
s = "SELECT tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
    " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
    " tbl_Scoring_ScoreCard_System36.NetPoints " & _
    " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap >= " & dblClassFrom & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap <= " & dblClassTo & ") " & _
    " ORDER BY tbl_Scoring_ScoreCard_System36.NetPoints, tbl_Scoring_ScoreCard_System36.Eagle DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Birdie DESC, tbl_Scoring_ScoreCard_System36.Par DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie DESC, tbl_Scoring_ScoreCard_System36.Boogie_2 DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie_3 DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
        With lstScoreResult.ListItems
            iWinner = 0
            For j = 1 To .Count
                If CDbl(.Item(j).SubItems(1)) = CDbl(rs!PlayerKey) Then
                    iWinner = 1
                End If
            Next j
        End With
        If CDbl(iWinner) = 0 Then
            i = i + 1
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!PlayerName
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!NetPoints
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
            Set x = lstScoreResult.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rs!PlayerKey
            x.SubItems(2) = rs!NetPoints
            x.SubItems(3) = "B"
            x.SubItems(4) = 2
            UpdateProgress picProgressBar, i / TotalRec
            GoTo BC:
        End If
        rs.MoveNext
    Wend
End If
rs.Close

BC:
RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "NET RUNNER-UP"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 11
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
s = "SELECT tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
    " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
    " tbl_Scoring_ScoreCard_System36.NetPoints " & _
    " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap >= " & dblClassFrom & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap <= " & dblClassTo & ") " & _
    " ORDER BY tbl_Scoring_ScoreCard_System36.NetPoints, tbl_Scoring_ScoreCard_System36.Eagle DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Birdie DESC, tbl_Scoring_ScoreCard_System36.Par DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie DESC, tbl_Scoring_ScoreCard_System36.Boogie_2 DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie_3 DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
        With lstScoreResult.ListItems
            iWinner = 0
            For j = 1 To .Count
                If CDbl(.Item(j).SubItems(1)) = CDbl(rs!PlayerKey) Then
                    iWinner = 1
                End If
            Next j
        End With
        If CDbl(iWinner) = 0 Then
            i = i + 1
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!PlayerName
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!NetPoints
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
            Set x = lstScoreResult.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rs!PlayerKey
            x.SubItems(2) = rs!NetPoints
            x.SubItems(3) = "B"
            x.SubItems(4) = 2
            UpdateProgress picProgressBar, i / TotalRec
            GoTo C:
        End If
        rs.MoveNext
    Wend
End If
rs.Close

'====== CLASS C
C:

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "CLASS C"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 12
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

s = "SELECT HFrom, HTo " & _
    " From tbl_Scoring_TournamentInfo_Class " & _
    " WHERE (TournamentKey = " & TournamentKey & ") " & _
    " AND (Class = 'C')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    dblClassFrom = rs!HFrom
    dblClassTo = rs!HTo
End If
rs.Close

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "GROSS CHAMPION"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 11
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
s = "SELECT tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
    " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
    " tbl_Scoring_ScoreCard_System36.GrossPoints " & _
    " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap >= " & dblClassFrom & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap <= " & dblClassTo & ") " & _
    " ORDER BY tbl_Scoring_ScoreCard_System36.GrossPoints, tbl_Scoring_ScoreCard_System36.Eagle DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Birdie DESC, tbl_Scoring_ScoreCard_System36.Par DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie DESC, tbl_Scoring_ScoreCard_System36.Boogie_2 DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie_3 DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
        With lstScoreResult.ListItems
            iWinner = 0
            For j = 1 To .Count
                If CDbl(.Item(j).SubItems(1)) = CDbl(rs!PlayerKey) Then
                    iWinner = 1
                End If
            Next j
        End With
        If CDbl(iWinner) = 0 Then
            i = i + 1
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!PlayerName
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!GrossPoints
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
            Set x = lstScoreResult.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rs!PlayerKey
            x.SubItems(2) = rs!GrossPoints
            x.SubItems(3) = "C"
            x.SubItems(4) = 1
            UpdateProgress picProgressBar, i / TotalRec
            GoTo CA:
        End If
        rs.MoveNext
    Wend
End If
rs.Close

CA:
RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "GROSS RUNNER-UP"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 11
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
s = "SELECT tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
    " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
    " tbl_Scoring_ScoreCard_System36.GrossPoints " & _
    " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap >= " & dblClassFrom & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap <= " & dblClassTo & ") " & _
    " ORDER BY tbl_Scoring_ScoreCard_System36.GrossPoints, tbl_Scoring_ScoreCard_System36.Eagle DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Birdie DESC, tbl_Scoring_ScoreCard_System36.Par DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie DESC, tbl_Scoring_ScoreCard_System36.Boogie_2 DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie_3 DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
        With lstScoreResult.ListItems
            iWinner = 0
            For j = 1 To .Count
                If CDbl(.Item(j).SubItems(1)) = CDbl(rs!PlayerKey) Then
                    iWinner = 1
                End If
            Next j
        End With
        If CDbl(iWinner) = 0 Then
            i = i + 1
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!PlayerName
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!GrossPoints
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
            Set x = lstScoreResult.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rs!PlayerKey
            x.SubItems(2) = rs!GrossPoints
            x.SubItems(3) = "C"
            x.SubItems(4) = 1
            UpdateProgress picProgressBar, i / TotalRec
            GoTo CB:
        End If
        rs.MoveNext
    Wend
End If
rs.Close

CB:

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "NET CHAMPION"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 11
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
s = "SELECT tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
    " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
    " tbl_Scoring_ScoreCard_System36.NetPoints " & _
    " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap >= " & dblClassFrom & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap <= " & dblClassTo & ") " & _
    " ORDER BY tbl_Scoring_ScoreCard_System36.NetPoints, tbl_Scoring_ScoreCard_System36.Eagle DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Birdie DESC, tbl_Scoring_ScoreCard_System36.Par DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie DESC, tbl_Scoring_ScoreCard_System36.Boogie_2 DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie_3 DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
        With lstScoreResult.ListItems
            iWinner = 0
            For j = 1 To .Count
                If CDbl(.Item(j).SubItems(1)) = CDbl(rs!PlayerKey) Then
                    iWinner = 1
                End If
            Next j
        End With
        If CDbl(iWinner) = 0 Then
            i = i + 1
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!PlayerName
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!NetPoints
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
            Set x = lstScoreResult.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rs!PlayerKey
            x.SubItems(2) = rs!NetPoints
            x.SubItems(3) = "C"
            x.SubItems(4) = 2
            UpdateProgress picProgressBar, i / TotalRec
            GoTo CC:
        End If
        rs.MoveNext
    Wend
End If
rs.Close

CC:
RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "NET RUNNER-UP"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 11
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
s = "SELECT tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
    " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
    " tbl_Scoring_ScoreCard_System36.NetPoints " & _
    " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap >= " & dblClassFrom & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap <= " & dblClassTo & ") " & _
    " ORDER BY tbl_Scoring_ScoreCard_System36.NetPoints, tbl_Scoring_ScoreCard_System36.Eagle DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Birdie DESC, tbl_Scoring_ScoreCard_System36.Par DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie DESC, tbl_Scoring_ScoreCard_System36.Boogie_2 DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie_3 DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
        With lstScoreResult.ListItems
            iWinner = 0
            For j = 1 To .Count
                If CDbl(.Item(j).SubItems(1)) = CDbl(rs!PlayerKey) Then
                    iWinner = 1
                End If
            Next j
        End With
        If CDbl(iWinner) = 0 Then
            i = i + 1
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!PlayerName
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!NetPoints
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
            Set x = lstScoreResult.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rs!PlayerKey
            x.SubItems(2) = rs!NetPoints
            x.SubItems(3) = "C"
            x.SubItems(4) = 2
            UpdateProgress picProgressBar, i / TotalRec
            GoTo D:
        End If
        rs.MoveNext
    Wend
End If
rs.Close

'==== CLASS D
D:

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "CLASS D"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 12
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

s = "SELECT HFrom, HTo " & _
    " From tbl_Scoring_TournamentInfo_Class " & _
    " WHERE (TournamentKey = " & TournamentKey & ") " & _
    " AND (Class = 'D')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    dblClassFrom = rs!HFrom
    dblClassTo = rs!HTo
End If
rs.Close

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "GROSS CHAMPION"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 11
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
s = "SELECT tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
    " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
    " tbl_Scoring_ScoreCard_System36.GrossPoints " & _
    " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap >= " & dblClassFrom & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap <= " & dblClassTo & ") " & _
    " ORDER BY tbl_Scoring_ScoreCard_System36.GrossPoints, tbl_Scoring_ScoreCard_System36.Eagle DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Birdie DESC, tbl_Scoring_ScoreCard_System36.Par DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie DESC, tbl_Scoring_ScoreCard_System36.Boogie_2 DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie_3 DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
        With lstScoreResult.ListItems
            iWinner = 0
            For j = 1 To .Count
                If CDbl(.Item(j).SubItems(1)) = CDbl(rs!PlayerKey) Then
                    iWinner = 1
                End If
            Next j
        End With
        If CDbl(iWinner) = 0 Then
            i = i + 1
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!PlayerName
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!GrossPoints
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
            Set x = lstScoreResult.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rs!PlayerKey
            x.SubItems(2) = rs!GrossPoints
            x.SubItems(3) = "D"
            x.SubItems(4) = 1
            UpdateProgress picProgressBar, i / TotalRec
            GoTo DA:
        End If
        rs.MoveNext
    Wend
End If
rs.Close

DA:
RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "GROSS RUNNER-UP"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 11
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
s = "SELECT tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
    " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
    " tbl_Scoring_ScoreCard_System36.GrossPoints " & _
    " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap >= " & dblClassFrom & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap <= " & dblClassTo & ") " & _
    " ORDER BY tbl_Scoring_ScoreCard_System36.GrossPoints, tbl_Scoring_ScoreCard_System36.Eagle DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Birdie DESC, tbl_Scoring_ScoreCard_System36.Par DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie DESC, tbl_Scoring_ScoreCard_System36.Boogie_2 DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie_3 DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
        With lstScoreResult.ListItems
            iWinner = 0
            For j = 1 To .Count
                If CDbl(.Item(j).SubItems(1)) = CDbl(rs!PlayerKey) Then
                    iWinner = 1
                End If
            Next j
        End With
        If CDbl(iWinner) = 0 Then
            i = i + 1
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!PlayerName
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!GrossPoints
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
            Set x = lstScoreResult.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rs!PlayerKey
            x.SubItems(2) = rs!GrossPoints
            x.SubItems(3) = "D"
            x.SubItems(4) = 1
            UpdateProgress picProgressBar, i / TotalRec
            GoTo DB:
        End If
        rs.MoveNext
    Wend
End If
rs.Close

DB:

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "NET CHAMPION"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 11
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
s = "SELECT tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
    " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
    " tbl_Scoring_ScoreCard_System36.NetPoints " & _
    " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap >= " & dblClassFrom & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap <= " & dblClassTo & ") " & _
    " ORDER BY tbl_Scoring_ScoreCard_System36.NetPoints, tbl_Scoring_ScoreCard_System36.Eagle DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Birdie DESC, tbl_Scoring_ScoreCard_System36.Par DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie DESC, tbl_Scoring_ScoreCard_System36.Boogie_2 DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie_3 DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
        With lstScoreResult.ListItems
            iWinner = 0
            For j = 1 To .Count
                If CDbl(.Item(j).SubItems(1)) = CDbl(rs!PlayerKey) Then
                    iWinner = 1
                End If
            Next j
        End With
        If CDbl(iWinner) = 0 Then
            i = i + 1
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!PlayerName
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!NetPoints
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
            Set x = lstScoreResult.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rs!PlayerKey
            x.SubItems(2) = rs!NetPoints
            x.SubItems(3) = "D"
            x.SubItems(4) = 2
            UpdateProgress picProgressBar, i / TotalRec
            GoTo DC:
        End If
        rs.MoveNext
    Wend
End If
rs.Close

DC:
RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "NET RUNNER-UP"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 11
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
s = "SELECT tbl_Scoring_ScoreCard_System36.PlayerKey, " & _
    " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
    " tbl_Scoring_ScoreCard_System36.NetPoints " & _
    " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap >= " & dblClassFrom & ") " & _
    " AND (tbl_Scoring_PlayerName.HandiCap <= " & dblClassTo & ") " & _
    " ORDER BY tbl_Scoring_ScoreCard_System36.NetPoints, tbl_Scoring_ScoreCard_System36.Eagle DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Birdie DESC, tbl_Scoring_ScoreCard_System36.Par DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie DESC, tbl_Scoring_ScoreCard_System36.Boogie_2 DESC, " & _
    " tbl_Scoring_ScoreCard_System36.Boogie_3 DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
        With lstScoreResult.ListItems
            iWinner = 0
            For j = 1 To .Count
                If CDbl(.Item(j).SubItems(1)) = CDbl(rs!PlayerKey) Then
                    iWinner = 1
                End If
            Next j
        End With
        If CDbl(iWinner) = 0 Then
            i = i + 1
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!PlayerName
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!NetPoints
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
            Set x = lstScoreResult.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rs!PlayerKey
            x.SubItems(2) = rs!NetPoints
            x.SubItems(3) = "D"
            x.SubItems(4) = 2
            UpdateProgress picProgressBar, i / TotalRec
            GoTo E:
        End If
        rs.MoveNext
    Wend
End If
rs.Close

E:

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "SPECIAL AWARD"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 12
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "MOST NUMBER OF PAR"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 11
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
s = "SELECT tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName  AS PlayerName, " & _
    " tbl_Scoring_ScoreCard_System36.Par " & _
    " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " Where (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
    " ORDER BY tbl_Scoring_ScoreCard_System36.Par DESC, tbl_Scoring_ScoreCard_System36.Boogie DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    rs.MoveFirst
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!PlayerName
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!Par
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
End If
rs.Close
i = i + 1
UpdateProgress picProgressBar, i / TotalRec

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "MOST NUMBER OF BERDIE"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 11
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
s = "SELECT tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName  AS PlayerName, " & _
    " tbl_Scoring_ScoreCard_System36.Birdie " & _
    " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " Where (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
    " ORDER BY tbl_Scoring_ScoreCard_System36.Birdie DESC, tbl_Scoring_ScoreCard_System36.Par DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    rs.MoveFirst
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!PlayerName
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!Birdie
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
End If
rs.Close
i = i + 1
UpdateProgress picProgressBar, i / TotalRec

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "MOST NUMBER OF EAGLE"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 11
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
s = "SELECT tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName  AS PlayerName, " & _
    " tbl_Scoring_ScoreCard_System36.Eagle " & _
    " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " Where (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
    " ORDER BY tbl_Scoring_ScoreCard_System36.Eagle DESC, tbl_Scoring_ScoreCard_System36.Birdie DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    rs.MoveFirst
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!PlayerName
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs!Eagle
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
End If
rs.Close
i = i + 1
UpdateProgress picProgressBar, i / TotalRec

SAVING:
On Error GoTo err_saving:
If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

xlsApp.Visible = True
        
picMain.Enabled = True
picToolbar.Enabled = True
picProgress.Visible = False
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub

Exit Sub
ErrorHandler:
Exit Sub

Exit Sub
err_saving:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING:
End Sub

Private Sub TimerPrintSummary_Timer()
TimerPrintSummary.Enabled = False
Dim TotalRec, RowCnt, ColCnt, iWorkSheet
Dim WorkbookName, Filename, cnt, ResetCnt, strRange
Dim dblClassFrom, dblClassTo, iWinner, dblEagle, dblBirdie, dblPar
With MainForm.CommonDialog1
    .CancelError = True
    On Error GoTo ErrorHandler
    .DialogTitle = "Save"
    .Filter = "Excel(*.xls)|*.xls"
    .ShowSave
    Filename = Trim(.Filename)
End With

On Error GoTo PG:

'picPrint.Visible = False
picProgressBar.BackColor = &HFFFFFF
picProgress.ZOrder 0
picProgress.Visible = True

WorkbookName = CStr(Filename)
'lstScoreResult.ListItems.Clear
RowCnt = 0
iWorkSheet = 1
Set xlsApp = CreateObject("Excel.Application")
xlsApp.Visible = False
xlsApp.Workbooks.Add
xlsApp.DisplayAlerts = False
xlsApp.Workbooks(1).Sheets(2).Delete
xlsApp.Workbooks(1).Sheets(2).Delete
xlsApp.Workbooks(1).Sheets(iWorkSheet).Activate
xlsApp.Workbooks(1).Sheets(iWorkSheet).Name = ReportClass
RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
    .Range(strRange).Value = TournamentName
    .Range(strRange).Font.Name = "Arial"
    .Range(strRange).Font.Size = 14
    .Range(strRange).Font.Bold = True
    
    RowCnt = RowCnt + 1
    ColCnt = 0
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Date : " & TournamentRange
    .Range(strRange).Font.Name = "Arial"
    .Range(strRange).Font.Size = 10
    .Range(strRange).Font.Bold = True
    
    If ReportClass <> "ALL" Then
        RowCnt = RowCnt + 1
        ColCnt = 0
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = ReportClass
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 10
        .Range(strRange).Font.Bold = True
    End If
    
    RowCnt = RowCnt + 1
    ColCnt = 0
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = ""
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Range(strRange).Font.Bold = False
    
    RowCnt = RowCnt + 1
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
    .Range(strRange).Value = ""
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Range(strRange).Font.Bold = True
    .Columns(ColCnt).ColumnWidth = 1
    .Range(strRange).HorizontalAlignment = 4
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Name"
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Columns(ColCnt).ColumnWidth = 20
    .Range(strRange).Font.Bold = True
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Gross Pts"
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Range(strRange).Font.Bold = True
    .Range(strRange).HorizontalAlignment = 4
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Handicap"
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Range(strRange).Font.Bold = True
    .Range(strRange).HorizontalAlignment = 4
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Net Pts"
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Range(strRange).Font.Bold = True
    .Range(strRange).HorizontalAlignment = 4
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Eagle"
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Range(strRange).Font.Bold = True
    .Range(strRange).HorizontalAlignment = 4
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Birdie"
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Range(strRange).Font.Bold = True
    .Range(strRange).HorizontalAlignment = 4
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Par"
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Range(strRange).Font.Bold = True
    .Range(strRange).HorizontalAlignment = 4
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Boogie"
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Range(strRange).Font.Bold = True
    .Range(strRange).HorizontalAlignment = 4
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Double Boogie"
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Range(strRange).ColumnWidth = 13
    .Range(strRange).Font.Bold = True
    .Range(strRange).HorizontalAlignment = 4
    
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    .Range(strRange).Value = "Triple Boogie"
    .Range(strRange).Font.Name = "Tahoma"
    .Range(strRange).Font.Size = 8
    .Range(strRange).ColumnWidth = 11
    .Range(strRange).Font.Bold = True
    .Range(strRange).HorizontalAlignment = 4
    i = 0
    Select Case ReportClass
        Case "ALL"
            s = "SELECT tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
                " tbl_Scoring_ScoreCard_System36.GrossPoints, tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_ScoreCard_System36.NetPoints, " & _
                " tbl_Scoring_ScoreCard_System36.Eagle, tbl_Scoring_ScoreCard_System36.Birdie, tbl_Scoring_ScoreCard_System36.Par, " & _
                " tbl_Scoring_ScoreCard_System36.Boogie, tbl_Scoring_ScoreCard_System36.Boogie_2, tbl_Scoring_ScoreCard_System36.Boogie_3 " & _
                " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
                " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                " Where (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
                " ORDER BY tbl_Scoring_ScoreCard_System36.NetPoints, tbl_Scoring_ScoreCard_System36.Eagle DESC, tbl_Scoring_ScoreCard_System36.Birdie DESC, " & _
                " tbl_Scoring_ScoreCard_System36.Par DESC, tbl_Scoring_ScoreCard_System36.Boogie DESC, tbl_Scoring_ScoreCard_System36.Boogie_2 DESC, " & _
                " tbl_Scoring_ScoreCard_System36.Boogie_3 DESC"
        Case Else
            s = "SELECT tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName, " & _
                " tbl_Scoring_ScoreCard_System36.GrossPoints, tbl_Scoring_PlayerName.HandiCap, tbl_Scoring_ScoreCard_System36.NetPoints, " & _
                " tbl_Scoring_ScoreCard_System36.Eagle, tbl_Scoring_ScoreCard_System36.Birdie, tbl_Scoring_ScoreCard_System36.Par, " & _
                " tbl_Scoring_ScoreCard_System36.Boogie , tbl_Scoring_ScoreCard_System36.Boogie_2, tbl_Scoring_ScoreCard_System36.Boogie_3 " & _
                " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
                " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
                " WHERE ((SELECT Class FROM tbl_Scoring_TournamentInfo_Class AS tbl_Scoring_TournamentInfo_Class_1 " & _
                " WHERE (TournamentKey = tbl_Scoring_ScoreCard_System36.TournamentKey) AND (HFrom <= tbl_Scoring_PlayerName.HandiCap) AND " & _
                " (HTo >= tbl_Scoring_PlayerName.HandiCap)) = '" & Replace(CStr(ReportClass), "CLASS ", "") & "') " & _
                " AND (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
                " ORDER BY tbl_Scoring_ScoreCard_System36.NetPoints, tbl_Scoring_ScoreCard_System36.Eagle DESC, tbl_Scoring_ScoreCard_System36.Birdie DESC, " & _
                " tbl_Scoring_ScoreCard_System36.Par DESC, tbl_Scoring_ScoreCard_System36.Boogie DESC, tbl_Scoring_ScoreCard_System36.Boogie_2 DESC, " & _
                " tbl_Scoring_ScoreCard_System36.Boogie_3 DESC"
    End Select
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        i = i + 1
        RowCnt = RowCnt + 1
        ColCnt = 0
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = i
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 8
        .Range(strRange).Font.Bold = False
        .Columns(ColCnt).ColumnWidth = 3
        .Range(strRange).HorizontalAlignment = 4
        
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = "."
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 8
        .Range(strRange).Font.Bold = False
        '.Columns(ColCnt).ColumnWidth = 3
        .Range(strRange).HorizontalAlignment = 4
        
        For j = 0 To rs.Fields.Count - 1
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = rs.Fields(j).Value
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = False
            '.Range(strRange).HorizontalAlignment = 4
        Next j
        
        UpdateProgress picProgressBar, i / rs.RecordCount
        rs.MoveNext
    Wend
    rs.Close
End With

picProgress.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True

SAVING:
On Error GoTo err_saving:
If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

xlsApp.Visible = True
        
Exit Sub
PG:
picProgress.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True

MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub

Exit Sub
ErrorHandler:
Exit Sub

Exit Sub
err_saving:
picProgress.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True

MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING:

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "ScoreCardControl36", "ScoreCardCtrl36", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "ScoreCardControl36", "ScoreCardCtrl36", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "ScoreCardControl36", "ScoreCardCtrl36", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "ScoreCardControl36", "ScoreCardCtrl36", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":   PRESS_F9
    Case "Close":   PRESS_ESCAPE
    Case Else: Exit Sub
End Select
End Sub

Private Sub txt2Boogies_Change(Index As Integer)
lstScores.ListItems.Item(5).SubItems(2) = RETURNTEXTVALUE(txt2Boogies(0)) + _
                                          RETURNTEXTVALUE(txt2Boogies(1)) + _
                                          RETURNTEXTVALUE(txt2Boogies(2)) + _
                                          RETURNTEXTVALUE(txt2Boogies(3)) + _
                                          RETURNTEXTVALUE(txt2Boogies(4)) + _
                                          RETURNTEXTVALUE(txt2Boogies(5)) + _
                                          RETURNTEXTVALUE(txt2Boogies(6)) + _
                                          RETURNTEXTVALUE(txt2Boogies(7)) + _
                                          RETURNTEXTVALUE(txt2Boogies(8)) + _
                                          RETURNTEXTVALUE(txt2Boogies(9)) + _
                                          RETURNTEXTVALUE(txt2Boogies(10)) + _
                                          RETURNTEXTVALUE(txt2Boogies(11)) + _
                                          RETURNTEXTVALUE(txt2Boogies(12)) + _
                                          RETURNTEXTVALUE(txt2Boogies(13)) + _
                                          RETURNTEXTVALUE(txt2Boogies(14)) + _
                                          RETURNTEXTVALUE(txt2Boogies(15)) + _
                                          RETURNTEXTVALUE(txt2Boogies(16)) + _
                                          RETURNTEXTVALUE(txt2Boogies(17))
End Sub

Private Sub txt3Boogies_Change(Index As Integer)
lstScores.ListItems.Item(6).SubItems(2) = RETURNTEXTVALUE(txt3Boogies(0)) + _
                                          RETURNTEXTVALUE(txt3Boogies(1)) + _
                                          RETURNTEXTVALUE(txt3Boogies(2)) + _
                                          RETURNTEXTVALUE(txt3Boogies(3)) + _
                                          RETURNTEXTVALUE(txt3Boogies(4)) + _
                                          RETURNTEXTVALUE(txt3Boogies(5)) + _
                                          RETURNTEXTVALUE(txt3Boogies(6)) + _
                                          RETURNTEXTVALUE(txt3Boogies(7)) + _
                                          RETURNTEXTVALUE(txt3Boogies(8)) + _
                                          RETURNTEXTVALUE(txt3Boogies(9)) + _
                                          RETURNTEXTVALUE(txt3Boogies(10)) + _
                                          RETURNTEXTVALUE(txt3Boogies(11)) + _
                                          RETURNTEXTVALUE(txt3Boogies(12)) + _
                                          RETURNTEXTVALUE(txt3Boogies(13)) + _
                                          RETURNTEXTVALUE(txt3Boogies(14)) + _
                                          RETURNTEXTVALUE(txt3Boogies(15)) + _
                                          RETURNTEXTVALUE(txt3Boogies(16)) + _
                                          RETURNTEXTVALUE(txt3Boogies(17))
End Sub

Private Sub txtBirdies_Change(Index As Integer)
lstScores.ListItems.Item(2).SubItems(2) = RETURNTEXTVALUE(txtBirdies(0)) + _
                                          RETURNTEXTVALUE(txtBirdies(1)) + _
                                          RETURNTEXTVALUE(txtBirdies(2)) + _
                                          RETURNTEXTVALUE(txtBirdies(3)) + _
                                          RETURNTEXTVALUE(txtBirdies(4)) + _
                                          RETURNTEXTVALUE(txtBirdies(5)) + _
                                          RETURNTEXTVALUE(txtBirdies(6)) + _
                                          RETURNTEXTVALUE(txtBirdies(7)) + _
                                          RETURNTEXTVALUE(txtBirdies(8)) + _
                                          RETURNTEXTVALUE(txtBirdies(9)) + _
                                          RETURNTEXTVALUE(txtBirdies(10)) + _
                                          RETURNTEXTVALUE(txtBirdies(11)) + _
                                          RETURNTEXTVALUE(txtBirdies(12)) + _
                                          RETURNTEXTVALUE(txtBirdies(13)) + _
                                          RETURNTEXTVALUE(txtBirdies(14)) + _
                                          RETURNTEXTVALUE(txtBirdies(15)) + _
                                          RETURNTEXTVALUE(txtBirdies(16)) + _
                                          RETURNTEXTVALUE(txtBirdies(17))
End Sub

Private Sub txtBoogies_Change(Index As Integer)
lstScores.ListItems.Item(4).SubItems(2) = RETURNTEXTVALUE(txtBoogies(0)) + _
                                          RETURNTEXTVALUE(txtBoogies(1)) + _
                                          RETURNTEXTVALUE(txtBoogies(2)) + _
                                          RETURNTEXTVALUE(txtBoogies(3)) + _
                                          RETURNTEXTVALUE(txtBoogies(4)) + _
                                          RETURNTEXTVALUE(txtBoogies(5)) + _
                                          RETURNTEXTVALUE(txtBoogies(6)) + _
                                          RETURNTEXTVALUE(txtBoogies(7)) + _
                                          RETURNTEXTVALUE(txtBoogies(8)) + _
                                          RETURNTEXTVALUE(txtBoogies(9)) + _
                                          RETURNTEXTVALUE(txtBoogies(10)) + _
                                          RETURNTEXTVALUE(txtBoogies(11)) + _
                                          RETURNTEXTVALUE(txtBoogies(12)) + _
                                          RETURNTEXTVALUE(txtBoogies(13)) + _
                                          RETURNTEXTVALUE(txtBoogies(14)) + _
                                          RETURNTEXTVALUE(txtBoogies(15)) + _
                                          RETURNTEXTVALUE(txtBoogies(16)) + _
                                          RETURNTEXTVALUE(txtBoogies(17))
End Sub

Private Sub txtDateAdd_GotFocus()
HTEXT txtDateAdd
End Sub

Private Sub txtDateAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAdd_Click
End Sub

Private Sub txtEagle_Change(Index As Integer)
lstScores.ListItems.Item(1).SubItems(2) = RETURNTEXTVALUE(txtEagle(0)) + _
                                          RETURNTEXTVALUE(txtEagle(1)) + _
                                          RETURNTEXTVALUE(txtEagle(2)) + _
                                          RETURNTEXTVALUE(txtEagle(3)) + _
                                          RETURNTEXTVALUE(txtEagle(4)) + _
                                          RETURNTEXTVALUE(txtEagle(5)) + _
                                          RETURNTEXTVALUE(txtEagle(6)) + _
                                          RETURNTEXTVALUE(txtEagle(7)) + _
                                          RETURNTEXTVALUE(txtEagle(8)) + _
                                          RETURNTEXTVALUE(txtEagle(9)) + _
                                          RETURNTEXTVALUE(txtEagle(10)) + _
                                          RETURNTEXTVALUE(txtEagle(11)) + _
                                          RETURNTEXTVALUE(txtEagle(12)) + _
                                          RETURNTEXTVALUE(txtEagle(13)) + _
                                          RETURNTEXTVALUE(txtEagle(14)) + _
                                          RETURNTEXTVALUE(txtEagle(15)) + _
                                          RETURNTEXTVALUE(txtEagle(16)) + _
                                          RETURNTEXTVALUE(txtEagle(17))
End Sub

Private Sub txtGrossScore_Change(Index As Integer)
Dim dblPar

If Index >= 0 And Index <= 8 Then
    dblPar = FGrid.TextMatrix(5, Index + 2)
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
    dblPar = FGrid.TextMatrix(5, Index + 3)
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

txtPts(Index).Text = RETURNTEXTVALUE(txtGrossScore(Index)) - CDbl(dblPar)

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

txtSGrossB.Text = RETURNTEXTVALUE(txtGrossScoreB)

txtGrossScoreTot.Text = RETURNTEXTVALUE(txtGrossScoreF) + _
                        RETURNTEXTVALUE(txtGrossScoreB)
End Sub

Private Sub txtGrossScoreF_Change()

txtSGrossF.Text = RETURNTEXTVALUE(txtGrossScoreF)

txtGrossScoreTot.Text = RETURNTEXTVALUE(txtGrossScoreF) + _
                        RETURNTEXTVALUE(txtGrossScoreB)
End Sub

Private Sub txtGrossScoreTot_Change()
txtScoreHDCP.Text = RETURNTEXTVALUE(txtGrossScoreTot) - RETURNTEXTVALUE(txtNet)
End Sub

Private Sub txtHandicap_Change()
Dim h As String
Dim rh As New ADODB.Recordset
txtClass.Text = ""
h = "SELECT Class " & _
    " From tbl_Scoring_TournamentInfo_Class " & _
    " WHERE (TournamentKey = " & TournamentKey & ") " & _
    " AND (HFrom <= " & RETURNTEXTVALUE(txtHandicap) & ") " & _
    " AND (HTo >= " & RETURNTEXTVALUE(txtHandicap) & ")"
If rh.State = adStateOpen Then rh.Close
rh.Open h, ConnOmega
If rh.RecordCount > 0 Then
    txtClass.Text = rh!Class
End If
rh.Close
End Sub

Private Sub txtiHDCP_Change(Index As Integer)
If Index >= 0 And Index <= 8 Then
    txtiHDCPF.Text = RETURNTEXTVALUE(txtiHDCP(0)) + _
                    RETURNTEXTVALUE(txtiHDCP(1)) + _
                    RETURNTEXTVALUE(txtiHDCP(2)) + _
                    RETURNTEXTVALUE(txtiHDCP(3)) + _
                    RETURNTEXTVALUE(txtiHDCP(4)) + _
                    RETURNTEXTVALUE(txtiHDCP(5)) + _
                    RETURNTEXTVALUE(txtiHDCP(6)) + _
                    RETURNTEXTVALUE(txtiHDCP(7)) + _
                    RETURNTEXTVALUE(txtiHDCP(8))
ElseIf Index >= 9 And Index <= 17 Then
    txtiHDCPB.Text = RETURNTEXTVALUE(txtiHDCP(9)) + _
                    RETURNTEXTVALUE(txtiHDCP(10)) + _
                    RETURNTEXTVALUE(txtiHDCP(11)) + _
                    RETURNTEXTVALUE(txtiHDCP(12)) + _
                    RETURNTEXTVALUE(txtiHDCP(13)) + _
                    RETURNTEXTVALUE(txtiHDCP(14)) + _
                    RETURNTEXTVALUE(txtiHDCP(15)) + _
                    RETURNTEXTVALUE(txtiHDCP(16)) + _
                    RETURNTEXTVALUE(txtiHDCP(17))
End If
End Sub

Private Sub txtiHDCPB_Change()

txtNetFB.Text = FGrid.TextMatrix(5, 21) + RETURNTEXTVALUE(txtiHDCPB)

txtiHDCPTot.Text = RETURNTEXTVALUE(txtiHDCPF) + _
                   RETURNTEXTVALUE(txtiHDCPB)
End Sub

Private Sub txtiHDCPF_Change()

txtNetF.Text = FGrid.TextMatrix(5, 11) + RETURNTEXTVALUE(txtiHDCPF)

txtiHDCPTot.Text = RETURNTEXTVALUE(txtiHDCPF) + _
                   RETURNTEXTVALUE(txtiHDCPB)
End Sub

Private Sub txtiHDCPTot_Change()
txtNet.Text = FGrid.TextMatrix(5, 22) + RETURNTEXTVALUE(txtiHDCPTot)
End Sub

Private Sub txtNet_Change()
txtScoreHDCP.Text = RETURNTEXTVALUE(txtGrossScoreTot) - RETURNTEXTVALUE(txtNet)
End Sub

Private Sub txtNetF_Change()
txtSNetF.Text = RETURNTEXTVALUE(txtNetF)
End Sub

Private Sub txtNetFB_Change()
txtSNetB.Text = RETURNTEXTVALUE(txtNetFB)
End Sub

Private Sub txtPars_Change(Index As Integer)
lstScores.ListItems.Item(3).SubItems(2) = RETURNTEXTVALUE(txtPars(0)) + _
                                          RETURNTEXTVALUE(txtPars(1)) + _
                                          RETURNTEXTVALUE(txtPars(2)) + _
                                          RETURNTEXTVALUE(txtPars(3)) + _
                                          RETURNTEXTVALUE(txtPars(4)) + _
                                          RETURNTEXTVALUE(txtPars(5)) + _
                                          RETURNTEXTVALUE(txtPars(6)) + _
                                          RETURNTEXTVALUE(txtPars(7)) + _
                                          RETURNTEXTVALUE(txtPars(8)) + _
                                          RETURNTEXTVALUE(txtPars(9)) + _
                                          RETURNTEXTVALUE(txtPars(10)) + _
                                          RETURNTEXTVALUE(txtPars(11)) + _
                                          RETURNTEXTVALUE(txtPars(12)) + _
                                          RETURNTEXTVALUE(txtPars(13)) + _
                                          RETURNTEXTVALUE(txtPars(14)) + _
                                          RETURNTEXTVALUE(txtPars(15)) + _
                                          RETURNTEXTVALUE(txtPars(16)) + _
                                          RETURNTEXTVALUE(txtPars(17))
End Sub

Private Sub txtPts_Change(Index As Integer)

txtiHDCP(Index).Text = IIf(RETURNTEXTVALUE(txtPts(Index)) < 0, RETURNTEXTVALUE(txtPts(Index)), IIf(RETURNTEXTVALUE(txtPts(Index)) > 2, RETURNTEXTVALUE(txtPts(Index)) - 2, 0))

Select Case RETURNTEXTVALUE(txtPts(Index))
    Case -5
        txtEagle(Index).Text = "0"
        txtBirdies(Index).Text = "0"
        txtPars(Index).Text = "0"
        txtBoogies(Index).Text = "0"
        txt2Boogies(Index).Text = "0"
        txt3Boogies(Index).Text = "0"
    Case -4
        txtEagle(Index).Text = "0"
        txtBirdies(Index).Text = "0"
        txtPars(Index).Text = "0"
        txtBoogies(Index).Text = "0"
        txt2Boogies(Index).Text = "0"
        txt3Boogies(Index).Text = "0"
    Case -3
        txtEagle(Index).Text = "0"
        txtBirdies(Index).Text = "0"
        txtPars(Index).Text = "0"
        txtBoogies(Index).Text = "0"
        txt2Boogies(Index).Text = "0"
        txt3Boogies(Index).Text = "0"
    Case -2
        txtEagle(Index).Text = "1"
        txtBirdies(Index).Text = "0"
        txtPars(Index).Text = "0"
        txtBoogies(Index).Text = "0"
        txt2Boogies(Index).Text = "0"
        txt3Boogies(Index).Text = "0"
    Case -1
        txtEagle(Index).Text = "0"
        txtBirdies(Index).Text = "1"
        txtPars(Index).Text = "0"
        txtBoogies(Index).Text = "0"
        txt2Boogies(Index).Text = "0"
        txt3Boogies(Index).Text = "0"
    Case 0
        txtEagle(Index).Text = "0"
        txtBirdies(Index).Text = "0"
        txtPars(Index).Text = "1"
        txtBoogies(Index).Text = "0"
        txt2Boogies(Index).Text = "0"
        txt3Boogies(Index).Text = "0"
    Case 1
        txtEagle(Index).Text = "0"
        txtBirdies(Index).Text = "0"
        txtPars(Index).Text = "0"
        txtBoogies(Index).Text = "1"
        txt2Boogies(Index).Text = "0"
        txt3Boogies(Index).Text = "0"
    Case 2
        txtEagle(Index).Text = "0"
        txtBirdies(Index).Text = "0"
        txtPars(Index).Text = "0"
        txtBoogies(Index).Text = "0"
        txt2Boogies(Index).Text = "1"
        txt3Boogies(Index).Text = "0"
    Case Else
        txtEagle(Index).Text = "0"
        txtBirdies(Index).Text = "0"
        txtPars(Index).Text = "0"
        txtBoogies(Index).Text = "0"
        txt2Boogies(Index).Text = "0"
        txt3Boogies(Index).Text = "1"
End Select
End Sub

Private Sub txtScoreHDCP_Change()
txtHandicap.Text = RETURNTEXTVALUE(txtScoreHDCP)
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstResult.Clear: cmbDate.Clear: Exit Sub
lstResult.Clear: cmbDate.Clear
s = "SELECT tbl_Scoring_PlayerName.PK, " & _
    " tbl_Scoring_PlayerName.LastName + ',  ' + tbl_Scoring_PlayerName.FirstName + '  ' + tbl_Scoring_PlayerName.MiddleName AS PlayerName " & _
    " FROM tbl_Scoring_ScoreCard_System36 LEFT OUTER JOIN " & _
    " tbl_Scoring_PlayerName ON tbl_Scoring_ScoreCard_System36.PlayerKey = tbl_Scoring_PlayerName.PK " & _
    " WHERE (tbl_Scoring_ScoreCard_System36.TournamentKey = " & TournamentKey & ") " & _
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
txtSGrossTot.Text = RETURNTEXTVALUE(txtSGrossF) + RETURNTEXTVALUE(txtSGrossB)
End Sub

Private Sub txtSGrossF_Change()
txtSGrossTot.Text = RETURNTEXTVALUE(txtSGrossF) + RETURNTEXTVALUE(txtSGrossB)
End Sub

Private Sub txtSNetB_Change()
txtSNetTot.Text = RETURNTEXTVALUE(txtSNetF) + RETURNTEXTVALUE(txtSNetB)
End Sub

Private Sub txtSNetF_Change()
txtSNetTot.Text = RETURNTEXTVALUE(txtSNetF) + RETURNTEXTVALUE(txtSNetB)
End Sub
