VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMembershipAction 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMembershipAction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   14175
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9960
      Top             =   600
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
            Picture         =   "frmMembershipAction.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipAction.frx":15A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipAction.frx":227E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipAction.frx":2F58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipAction.frx":3C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipAction.frx":490C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipAction.frx":55E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipAction.frx":62C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipAction.frx":6F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipAction.frx":7874
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipAction.frx":854E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipAction.frx":9228
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipAction.frx":9F02
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipAction.frx":ABDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipAction.frx":B8B6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   39
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   40
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
         MouseIcon       =   "frmMembershipAction.frx":C590
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
      Height          =   300
      Left            =   0
      TabIndex        =   8
      Top             =   7095
      Width           =   14175
      _ExtentX        =   25003
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
   Begin RPVGCC.b8Container picAdd 
      Height          =   4875
      Left            =   4800
      TabIndex        =   20
      Top             =   1320
      Visible         =   0   'False
      Width           =   4455
      _extentx        =   7858
      _extenty        =   8599
      backcolor       =   15396057
      Begin VB.TextBox txtSearchAssignor 
         Height          =   315
         Left            =   120
         TabIndex        =   46
         Top             =   2400
         Width           =   4215
      End
      Begin VB.ListBox lstResultAssignorAdd 
         Height          =   1425
         Left            =   120
         TabIndex        =   45
         Top             =   2760
         Width           =   4215
      End
      Begin VB.ListBox lstResultAdd 
         Height          =   1230
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtSearchAdd 
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   4215
      End
      Begin VB.CommandButton cmdCancelAdd 
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
         Picture         =   "frmMembershipAction.frx":C8AA
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   4275
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKAdd 
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
         Picture         =   "frmMembershipAction.frx":D006
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4275
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   25
         Top             =   45
         Width           =   4365
         _extentx        =   7699
         _extenty        =   609
         caption         =   "Search"
         font            =   "frmMembershipAction.frx":D678
         fontbold        =   -1  'True
         fontname        =   "Tahoma"
         fontsize        =   8.25
         autofunction    =   0   'False
         icon            =   "frmMembershipAction.frx":D6A0
         shadowvisible   =   0   'False
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Assignor"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   2160
         Width           =   3495
      End
   End
   Begin RPVGCC.b8Container picSearch 
      Height          =   4875
      Left            =   4800
      TabIndex        =   31
      Top             =   1320
      Visible         =   0   'False
      Width           =   4455
      _extentx        =   7858
      _extenty        =   8599
      backcolor       =   15396057
      Begin VB.ListBox lstSearchIDNumbers 
         Height          =   1035
         Left            =   120
         TabIndex        =   37
         Top             =   3000
         Width           =   4215
      End
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
         Left            =   600
         Picture         =   "frmMembershipAction.frx":DC3C
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   4155
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
         Left            =   2280
         Picture         =   "frmMembershipAction.frx":E2AE
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   4155
         Width           =   1560
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   4215
      End
      Begin VB.ListBox lstSearch 
         Height          =   1815
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   4215
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   45
         TabIndex        =   36
         Top             =   45
         Width           =   4365
         _extentx        =   7699
         _extenty        =   609
         caption         =   "Search"
         font            =   "frmMembershipAction.frx":EA0A
         fontbold        =   -1  'True
         fontname        =   "Tahoma"
         fontsize        =   8.25
         autofunction    =   0   'False
         icon            =   "frmMembershipAction.frx":EA32
         shadowvisible   =   0   'False
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Number/s"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   2760
         Width           =   1335
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   240
      ScaleHeight     =   5655
      ScaleWidth      =   13695
      TabIndex        =   9
      Top             =   1200
      Width           =   13695
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5535
         Left            =   7920
         ScaleHeight     =   5505
         ScaleWidth      =   5625
         TabIndex        =   41
         Top             =   0
         Width           =   5655
         Begin VB.Image imgPictureMovement 
            Height          =   5505
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   5625
         End
      End
      Begin VB.TextBox txtEffectDateTo 
         Height          =   315
         Left            =   6600
         MaxLength       =   100
         TabIndex        =   29
         Text            =   "07/17/2011"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtRemarks 
         Height          =   315
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   27
         Top             =   1080
         Width           =   6375
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C6B8A4&
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   0
         ScaleHeight     =   1905
         ScaleWidth      =   7665
         TabIndex        =   17
         Top             =   3720
         Width           =   7695
         Begin VB.TextBox txtBoughtFromIDSeries 
            Height          =   315
            Left            =   3240
            MaxLength       =   100
            TabIndex        =   44
            Top             =   840
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1575
            Left            =   5880
            ScaleHeight     =   1545
            ScaleWidth      =   1545
            TabIndex        =   43
            Top             =   120
            Width           =   1575
            Begin VB.Image imgPictureBought 
               Height          =   1545
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1545
            End
         End
         Begin VB.TextBox txtBoughtFromID 
            Height          =   315
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   7
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox txtBoughtFrom 
            Height          =   315
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   6
            Top             =   480
            Width           =   4455
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00404040&
            Height          =   1575
            Left            =   5955
            Top             =   195
            Width           =   1575
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Bought From ID"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   870
            Width           =   1215
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Bought From"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   510
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C6B8A4&
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   0
         ScaleHeight     =   1905
         ScaleWidth      =   7665
         TabIndex        =   14
         Top             =   1560
         Width           =   7695
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1575
            Left            =   5880
            ScaleHeight     =   1545
            ScaleWidth      =   1545
            TabIndex        =   42
            Top             =   120
            Width           =   1575
            Begin VB.Image imgPictureAssignor 
               Height          =   1545
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1545
            End
         End
         Begin VB.TextBox txtAssignorIDSeries 
            Height          =   315
            Left            =   3240
            MaxLength       =   100
            TabIndex        =   30
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtAssignor 
            Height          =   315
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   4
            Top             =   600
            Width           =   4455
         End
         Begin VB.TextBox txtAssignorID 
            Height          =   315
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   5
            Top             =   960
            Width           =   1815
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00404040&
            Height          =   1575
            Left            =   5955
            Top             =   195
            Width           =   1575
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Assignor"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   630
            Width           =   1095
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Assignor ID"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   990
            Width           =   1095
         End
      End
      Begin VB.TextBox txtID 
         Height          =   315
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtEffectDateFrom 
         Height          =   315
         Left            =   5160
         MaxLength       =   100
         TabIndex        =   3
         Text            =   "07/17/2011"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtCtrl 
         Height          =   315
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   0
         Top             =   0
         Width           =   1815
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   1
         Top             =   360
         Width           =   6375
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         Height          =   5535
         Left            =   7995
         Top             =   75
         Width           =   5655
      End
      Begin VB.Line Line1 
         X1              =   6360
         X2              =   6500
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label lblActionType 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SHARE HOLDER"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3720
         TabIndex        =   26
         Top             =   80
         Width           =   3975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Assign ID"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Effectivity Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl #"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   30
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   390
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmMembershipAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim tmp As Long

Public iActionType, iMemberKey, iIDNumberKey, iAssignorKey, iBoughtKey, iSearchAdd

Dim sCtrl, iPK, TmpIDNumber, FIDNumber, sFIDNumber, _
sFullName, Arr, TmpID01, TmpID02, TmpID03, i, j, _
iIDCounter, iIDKey, dMemberIDSeries, iRefKey


Private Sub BROWSER(sCtrl, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If sCtrl <> "" Then
            s = "SELECT TOP 1 dbo.tbl_Member_Action.* " & _
                " FROM dbo.tbl_Member_Action " & _
                " WHERE (Ctrl = '" & sCtrl & "') " & _
                " ORDER BY Ctrl "
        Else
            s = "SELECT TOP 1 dbo.tbl_Member_Action.* " & _
                " FROM dbo.tbl_Member_Action " & _
                " ORDER BY Ctrl "
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picAdd.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        s = "SELECT TOP 1 dbo.tbl_Member_Action.* " & _
            " FROM dbo.tbl_Member_Action " & _
            " ORDER BY Ctrl "
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picAdd.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        s = "SELECT TOP 1 dbo.tbl_Member_Action.* " & _
            " FROM dbo.tbl_Member_Action " & _
            " WHERE (Ctrl < '" & sCtrl & "') " & _
            " ORDER BY Ctrl DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picAdd.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        s = "SELECT TOP 1 dbo.tbl_Member_Action.* " & _
            " FROM dbo.tbl_Member_Action " & _
            " WHERE (Ctrl > '" & sCtrl & "') " & _
            " ORDER BY Ctrl"
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picAdd.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        s = "SELECT TOP 1 dbo.tbl_Member_Action.* " & _
            " FROM dbo.tbl_Member_Action " & _
            " ORDER BY Ctrl DESC"
    Case "is_FIND"
        s = "SELECT TOP 1 dbo.tbl_Member_Action.* " & _
            " FROM dbo.tbl_Member_Action " & _
            " WHERE (PK = " & sCtrl & ") "
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then

    'iMemberKey = rs!MemberKey
    iIDNumberKey = rs!IDNumberKey
    iAssignorKey = rs!AssignorKey
    iBoughtKey = rs!SoldToKey
    iActionType = rs!ActionType
    lblActionType.Caption = IIf(rs!ActionType = 1, "ASSIGNEE", IIf(rs!ActionType = 2, "ASSIGNOR", IIf(rs!ActionType = 3, "BOUGHT SHARE", "")))
    lblActionType.ForeColor = IIf(rs!ActionType = 1, &HFF&, IIf(rs!ActionType = 2, &HFF00&, IIf(rs!ActionType = 3, &H800000, &H0&)))
    txtCtrl.Text = rs!Ctrl
    txtEffectDateFrom.Text = Format(rs!EffectDateFrom, "mm/dd/yyyy")
    txtEffectDateTo.Text = IIf(IsNull(rs!EffectDateTo), "", rs!EffectDateTo)
    txtRemarks.Text = rs!Remarks
    txtName.Text = ""
    txtID.Text = ""
    imgPictureMovement.Picture = LoadPicture("")
    t = "SELECT dbo.tbl_Member_Information.LastName, " & _
        " dbo.tbl_Member_Information.FirstName, " & _
        " dbo.tbl_Member_Information.MiddleName, " & _
        " dbo.tbl_Member_IDNumber.IDNumber " & _
        " FROM dbo.tbl_Member_IDNumber LEFT OUTER JOIN " & _
        " dbo.tbl_Member_Information ON dbo.tbl_Member_IDNumber.MemberKey = dbo.tbl_Member_Information.PK " & _
        " WHERE (dbo.tbl_Member_IDNumber.PK = " & iIDNumberKey & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtName.Text = rt!LastName & ",  " & rt!FirstName & "  " & rt!MiddleName
        txtID.Text = rt!IDNumber
        imgPictureMovement.Picture = LoadPicture(SHOW_IMAGES(iIDNumberKey, 0, "Member ID Number"))
    End If
    rt.Close
    
    txtAssignor.Text = ""
    txtAssignorID.Text = ""
    imgPictureAssignor.Picture = LoadPicture("")
    t = "SELECT dbo.tbl_Member_Information.LastName, " & _
        " dbo.tbl_Member_Information.FirstName, " & _
        " dbo.tbl_Member_Information.MiddleName, " & _
        " dbo.tbl_Member_IDNumber.IDNumber " & _
        " FROM dbo.tbl_Member_IDNumber LEFT OUTER JOIN " & _
        " dbo.tbl_Member_Information ON dbo.tbl_Member_IDNumber.MemberKey = dbo.tbl_Member_Information.PK " & _
        " WHERE (dbo.tbl_Member_IDNumber.PK = " & iAssignorKey & ") " & _
        " AND (dbo.tbl_Member_IDNumber.ViewNot = 0)"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtAssignor.Text = rt!LastName & ",  " & rt!FirstName & "  " & rt!MiddleName
        txtAssignorID.Text = rt!IDNumber
        imgPictureAssignor.Picture = LoadPicture(SHOW_IMAGES(iAssignorKey, 0, "Member ID Number"))
    End If
    rt.Close
    
    txtBoughtFrom.Text = ""
    txtBoughtFromID.Text = ""
    imgPictureBought.Picture = LoadPicture("")
    t = "SELECT dbo.tbl_Member_Information.LastName, " & _
        " dbo.tbl_Member_Information.FirstName, " & _
        " dbo.tbl_Member_Information.MiddleName, " & _
        " dbo.tbl_Member_IDNumber.IDNumber " & _
        " FROM dbo.tbl_Member_IDNumber LEFT OUTER JOIN " & _
        " dbo.tbl_Member_Information ON dbo.tbl_Member_IDNumber.MemberKey = dbo.tbl_Member_Information.PK " & _
        " WHERE (dbo.tbl_Member_IDNumber.PK = " & iBoughtKey & ") " & _
        " AND (dbo.tbl_Member_IDNumber.ViewNot = 0)"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtBoughtFrom.Text = rt!LastName & ",  " & rt!FirstName & "  " & rt!MiddleName
        txtBoughtFromID.Text = rt!IDNumber
        imgPictureBought.Picture = LoadPicture(SHOW_IMAGES(iBoughtKey, 0, "Member ID Number"))
    End If
    rt.Close
    
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    
    SaveSetting App.EXEName, "MemberAction", "MemAction", rs!Ctrl
    
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If AccessRights("Membership Action", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
PopupMenu MainFormPopupF.mnuMemberActionAdd, , Toolbar1.Buttons(1).Left, Toolbar1.Buttons(1).Height
End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If AccessRights("Membership Action", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
End Sub

Private Sub PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If AccessRights("Membership Action", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
's = "SELECT tbl_Member_Action.* " & _
'    " FROM tbl_Member_Action " & _
'    " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
'If rs.State = adStateOpen Then rs.Close
'rs.Open s, ConnOmega
'If rs.RecordCount > 0 Then
'    If rs!ActionType = 2 Then
'        t = ""
'        If rt.State = adStateOpen Then rt.Close
'        rt.Open t, ConnOmega
'
'    End If
'End If
'rs.Close
'On Error GoTo PG:
'ConnOmega.Execute "DELETE FROM"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub

If iActionType = 0 Then MsgBox "Please Select Action Taken!                 ", vbCritical, "Error...": Exit Sub
If iMemberKey = 0 Then MsgBox "Please Select Member!                ", vbCritical, "Error...": Exit Sub
If iActionType = 1 Then
    If iAssignorKey = 0 Then MsgBox "Please Select Assignor!                ", vbCritical, "Error...": Exit Sub
End If
If IsDate(txtEffectDateFrom.Text) = False Then MsgBox "Please Supply a Valid Date!                  ", vbCritical, "Error...": txtEffectDateFrom.SetFocus: Exit Sub
If IsDate(txtEffectDateTo.Text) = True Then
    If DateValue(FormatDateTime(txtEffectDateFrom.Text, vbShortDate)) > DateValue(FormatDateTime(txtEffectDateTo.Text, vbShortDate)) Then
        MsgBox "Invalid Range!                              ", vbCritical, "Error...", "Error...": txtEffectDateTo.SetFocus: Exit Sub
    End If
End If

On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    sCtrl = ""
    s = "SELECT TOP 1 Ctrl " & _
        " FROM tbl_Member_Action " & _
        " WHERE (Year(DateEncoded) = " & Year(Now) & ") " & _
        " ORDER BY Ctrl DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        sCtrl = CDbl(rs!Ctrl) + 1
    Else
        sCtrl = Format(Now, "yyyy") & "0000"
    End If
    rs.Close
    
    Do
        s = "SELECT Ctrl " & _
            " FROM tbl_Member_Action " & _
            " WHERE (Ctrl = '" & sCtrl & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If
        rs.Close
        sCtrl = CDbl(sCtrl) + 1
    Loop
    
    iIDNumberKey = 0
    Select Case iActionType
        Case 1      'Assignee
            'Add to Member's ID
            iIDCounter = 0
            s = "SELECT TOP 1 IDNumber, IDCounter " & _
                " From tbl_Member_IDNumber " & _
                " WHERE (IDNumber = '" & Trim(txtID.Text) & "') " & _
                " ORDER BY IDCounter DESC"
            If rs.State = adStateOpen Then rs.Close
            rs.Open s, ConnOmega
            If rs.RecordCount > 0 Then
                iIDCounter = rs!IDCounter + 1
            Else
                iIDCounter = 1
            End If
            rs.Close
            
            Arr = Split(Trim(txtID.Text), "-", -1, 1)
            TmpID01 = Arr(0): TmpID02 = Arr(1): TmpID03 = Arr(2)
            t = "SELECT LastName, FirstName, MiddleName, " & _
                " CivilStatus, SpouseLName, SpouseGName, " & _
                " SpouseMName, MemberPicture, SpousePicture " & _
                " From tbl_Member_Information " & _
                " WHERE (PK = " & iMemberKey & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                FIDNumber = TmpID01 & "-" & TmpID02 & "-" & TmpID03
                sFIDNumber = FIDNumber
                sFullName = rt!LastName & ",  " & rt!FirstName & "  " & rt!MiddleName
                ConnOmega.Execute "INSERT INTO tbl_Member_IDNumber " & _
                                  " (MemberKey, MemberName, IDNumber, MemberType, LastModified, IDCounter, MemberChildLine) " & _
                                  " VALUES (" & iMemberKey & ", '" & FORMATSQL(CStr(sFullName)) & "', " & _
                                  " '" & FIDNumber & "', 1, '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                                  " " & iIDCounter & ", 0)"
                u = "SELECT PK " & _
                    " FROM tbl_Member_IDNumber " & _
                    " WHERE (MemberKey = " & iMemberKey & ") " & _
                    " AND (MemberName = '" & FORMATSQL(CStr(sFullName)) & "') " & _
                    " AND (IDNumber = '" & FIDNumber & "') " & _
                    " AND (MemberType = 1) " & _
                    " AND (IDCounter = " & iIDCounter & ")"
                If ru.State = adStateOpen Then ru.Close
                ru.Open u, ConnOmega
                If ru.RecordCount > 0 Then
                    iIDNumberKey = ru!PK
                End If
                ru.Close
                
                If IsNull(rt!MemberPicture) = False Then
                    u = "SELECT PK " & _
                        " FROM tbl_Member_IDNumber " & _
                        " WHERE (IDNumber = '" & FIDNumber & "')"
                    If ru.State = adStateOpen Then ru.Close
                    ru.Open u, ConnOmega
                    If ru.RecordCount > 0 Then
                        SAVE_IMAGES ru!PK, 0, SHOW_IMAGES(iMemberKey, 0, "Member"), "Member ID Number"
                    End If
                    ru.Close
                End If
                
                
                If rt!CivilStatus = 2 Then
                    If Trim(rt!SpouseLName) <> "" And Trim(rt!SpouseGName) <> "" Then
                        FIDNumber = TmpID01 & "-" & TmpID02 & "-" & Format(CDbl(TmpID03) + 1, "0#")
                        sFullName = rt!SpouseLName & ",  " & rt!SpouseGName & "  " & rt!SpouseMName
                        ConnOmega.Execute "INSERT INTO tbl_Member_IDNumber " & _
                                          " (MemberKey, MemberName, IDNumber, MemberType, LastModified, IDCounter, MemberChildLine) " & _
                                          " VALUES (" & iMemberKey & ", '" & FORMATSQL(CStr(sFullName)) & "', " & _
                                          " '" & FIDNumber & "', 2, '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                                          " " & iIDCounter & ", 0)"
                                        
                        If IsNull(rt!SpousePicture) = False Then
                            u = "SELECT PK " & _
                                " FROM tbl_Member_IDNumber " & _
                                " WHERE (IDNumber = '" & FIDNumber & "')"
                            If ru.State = adStateOpen Then ru.Close
                            ru.Open u, ConnOmega
                            If ru.RecordCount > 0 Then
                                SAVE_IMAGES ru!PK, 0, SHOW_IMAGES(iMemberKey, 0, "Member Spouse"), "Member ID Number"
                            End If
                            ru.Close
                        End If
                    End If
                End If
                
                j = 0
                i = 2
                v = "SELECT ChildLName, ChildGName, ChildMName, " & _
                    " ChildStatus, ChildBirthDate, ChildPicture " & _
                    " From tbl_Member_Dependent " & _
                    " Where (MemberKey = " & iMemberKey & ") " & _
                    " ORDER BY ChildBirthDate"
                If rv.State = adStateOpen Then rv.Close
                rv.Open v, ConnOmega
                While Not rv.EOF
                    j = j + 1
                    If rv!ChildStatus = 1 Then
                        If Get_Age(FormatDateTime(rv!ChildBirthDate, vbShortDate), FormatDateTime(Date, vbShortDate)) <= 25 Then
                            i = i + 1
                            FIDNumber = TmpID01 & "-" & TmpID02 & "-" & Format(i, "0#")
                            sFullName = rv!ChildLName & ",  " & rv!ChildGName & "  " & rv!ChildMName
                            ConnOmega.Execute "INSERT INTO tbl_Member_IDNumber " & _
                                              " (MemberKey, MemberName, IDNumber, MemberType, MemberCStatus, MemberBDay, LastModified, IDCounter, MemberChildLine) " & _
                                              " VALUES (" & iMemberKey & ", '" & FORMATSQL(CStr(sFullName)) & "', " & _
                                              " '" & FIDNumber & "', 3, " & rv!ChildStatus & ", '" & FormatDateTime(rv!ChildBirthDate, vbShortDate) & "', " & _
                                              " '" & CStr(Now) & " - " & gbl_CompleteName & "', " & iIDCounter & ", " & j & ")"
                            
                            If IsNull(rv!ChildPicture) = False Then
                                u = "SELECT PK " & _
                                    " FROM tbl_Member_IDNumber " & _
                                    " WHERE (IDNumber = '" & FIDNumber & "')"
                                If ru.State = adStateOpen Then ru.Close
                                ru.Open u, ConnOmega
                                If ru.RecordCount > 0 Then
                                    SAVE_IMAGES ru!PK, j, SHOW_IMAGES(iMemberKey, j, "Member Child"), "Member ID Number (Child)"
                                End If
                                ru.Close
                            End If
                            
                        End If
                    End If
                    rv.MoveNext
                Wend
                rv.Close
            
            End If
            rt.Close
            
        Case 2      'Share Holder
        
            iIDCounter = 0
            u = "SELECT PK " & _
                " FROM tbl_Member_IDNumber " & _
                " WHERE (IDNumber = '" & Trim(txtID.Text) & "')"
            If ru.State = adStateOpen Then ru.Close
            ru.Open u, ConnOmega
            If ru.RecordCount > 0 Then
                iIDNumberKey = ru!PK
            End If
            ru.Close
        
        Case 3      'Bought Share
            
            iIDCounter = 0
            s = "SELECT TOP 1 IDNumber, IDCounter " & _
                " From tbl_Member_IDNumber " & _
                " WHERE (IDNumber = '" & Trim(txtID.Text) & "') " & _
                " ORDER BY IDCounter DESC"
            If rs.State = adStateOpen Then rs.Close
            rs.Open s, ConnOmega
            If rs.RecordCount > 0 Then
                iIDCounter = rs!IDCounter + 1
            Else
                iIDCounter = 1
            End If
            rs.Close
            
            Arr = Split(Trim(txtID.Text), "-", -1, 1)
            TmpID01 = Arr(0): TmpID02 = Arr(1): TmpID03 = Arr(2)
            t = "SELECT LastName, FirstName, MiddleName, " & _
                " CivilStatus, SpouseLName, SpouseGName, " & _
                " SpouseMName, MemberPicture, SpousePicture " & _
                " From tbl_Member_Information " & _
                " WHERE (PK = " & iMemberKey & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                FIDNumber = TmpID01 & "-" & TmpID02 & "-" & TmpID03
                sFIDNumber = FIDNumber
                sFullName = rt!LastName & ",  " & rt!FirstName & "  " & rt!MiddleName
                ConnOmega.Execute "INSERT INTO tbl_Member_IDNumber " & _
                                  " (MemberKey, MemberName, IDNumber, MemberType, LastModified, IDCounter, MemberChildLine) " & _
                                  " VALUES (" & iMemberKey & ", '" & FORMATSQL(CStr(sFullName)) & "', " & _
                                  " '" & FIDNumber & "', 1, '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                                  " " & iIDCounter & ", 0)"
                u = "SELECT PK " & _
                    " FROM tbl_Member_IDNumber " & _
                    " WHERE (MemberKey = " & iMemberKey & ") " & _
                    " AND (MemberName = '" & FORMATSQL(CStr(sFullName)) & "') " & _
                    " AND (IDNumber = '" & FIDNumber & "') " & _
                    " AND (MemberType = 1) " & _
                    " AND (IDCounter = " & iIDCounter & ")"
                If ru.State = adStateOpen Then ru.Close
                ru.Open u, ConnOmega
                If ru.RecordCount > 0 Then
                    iIDNumberKey = ru!PK
                End If
                ru.Close
                
                If IsNull(rt!MemberPicture) = False Then
                    u = "SELECT PK " & _
                        " FROM tbl_Member_IDNumber " & _
                        " WHERE (IDNumber = '" & FIDNumber & "')"
                    If ru.State = adStateOpen Then ru.Close
                    ru.Open u, ConnOmega
                    If ru.RecordCount > 0 Then
                        SAVE_IMAGES ru!PK, 0, SHOW_IMAGES(iMemberKey, 0, "Member"), "Member ID Number"
                    End If
                    ru.Close
                End If
                
                
                If rt!CivilStatus = 2 Then
                    If Trim(rt!SpouseLName) <> "" And Trim(rt!SpouseGName) <> "" Then
                        FIDNumber = TmpID01 & "-" & TmpID02 & "-" & Format(CDbl(TmpID03) + 1, "0#")
                        sFullName = rt!SpouseLName & ",  " & rt!SpouseGName & "  " & rt!SpouseMName
                        ConnOmega.Execute "INSERT INTO tbl_Member_IDNumber " & _
                                          " (MemberKey, MemberName, IDNumber, MemberType, LastModified, IDCounter, MemberChildLine) " & _
                                          " VALUES (" & iMemberKey & ", '" & FORMATSQL(CStr(sFullName)) & "', " & _
                                          " '" & FIDNumber & "', 2, '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                                          " " & iIDCounter & ", 0)"
                                        
                        If IsNull(rt!SpousePicture) = False Then
                            u = "SELECT PK " & _
                                " FROM tbl_Member_IDNumber " & _
                                " WHERE (IDNumber = '" & FIDNumber & "')"
                            If ru.State = adStateOpen Then ru.Close
                            ru.Open u, ConnOmega
                            If ru.RecordCount > 0 Then
                                SAVE_IMAGES ru!PK, 0, SHOW_IMAGES(iMemberKey, 0, "Member Spouse"), "Member ID Number"
                            End If
                            ru.Close
                        End If
                    End If
                End If
                
                j = 0
                i = 2
                v = "SELECT ChildLName, ChildGName, ChildMName, " & _
                    " ChildStatus, ChildBirthDate, ChildPicture " & _
                    " From tbl_Member_Dependent " & _
                    " Where (MemberKey = " & iMemberKey & ") " & _
                    " ORDER BY ChildBirthDate"
                If rv.State = adStateOpen Then rv.Close
                rv.Open v, ConnOmega
                While Not rv.EOF
                    j = j + 1
                    If rv!ChildStatus = 1 Then
                        If Get_Age(FormatDateTime(rv!ChildBirthDate, vbShortDate), FormatDateTime(Date, vbShortDate)) <= 25 Then
                            i = i + 1
                            FIDNumber = TmpID01 & "-" & TmpID02 & "-" & Format(i, "0#")
                            sFullName = rv!ChildLName & ",  " & rv!ChildGName & "  " & rv!ChildMName
                            ConnOmega.Execute "INSERT INTO tbl_Member_IDNumber " & _
                                              " (MemberKey, MemberName, IDNumber, MemberType, MemberCStatus, MemberBDay, LastModified, IDCounter, MemberChildLine) " & _
                                              " VALUES (" & iMemberKey & ", '" & FORMATSQL(CStr(sFullName)) & "', " & _
                                              " '" & FIDNumber & "', 3, " & rv!ChildStatus & ", '" & FormatDateTime(rv!ChildBirthDate, vbShortDate) & "', " & _
                                              " '" & CStr(Now) & " - " & gbl_CompleteName & "', " & iIDCounter & ", " & j & ")"
                            
                            If IsNull(rv!ChildPicture) = False Then
                                u = "SELECT PK " & _
                                    " FROM tbl_Member_IDNumber " & _
                                    " WHERE (IDNumber = '" & FIDNumber & "')"
                                If ru.State = adStateOpen Then ru.Close
                                ru.Open u, ConnOmega
                                If ru.RecordCount > 0 Then
                                    SAVE_IMAGES ru!PK, j, SHOW_IMAGES(iMemberKey, j, "Member Child"), "Member ID Number (Child)"
                                End If
                                ru.Close
                            End If
                            
                        End If
                    End If
                    rv.MoveNext
                Wend
                rv.Close
            
            End If
            rt.Close
                        
    End Select
    
    ConnOmega.Execute "INSERT INTO tbl_Member_Action " & _
                      " (Ctrl, DateEncoded, ActionType, EffectDateFrom, IDNumberKey, " & _
                      " MemberID, MemberIDSeries, Remarks, LastModified) " & _
                      " VALUES ('" & sCtrl & "', '" & FormatDateTime(Now, vbShortDate) & "', " & _
                      " " & iActionType & ", '" & FormatDateTime(txtEffectDateFrom.Text, vbShortDate) & "', " & _
                      " " & iIDNumberKey & ", '" & Trim(txtID.Text) & "', " & iIDCounter & ", " & _
                      " '" & FORMATSQL(Trim(txtRemarks.Text)) & "', " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
    
    
    
    iPK = 0
    s = "SELECT PK " & _
        " FROM tbl_Member_Action " & _
        " WHERE (Ctrl = '" & sCtrl & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        iPK = rs!PK
    End If
    rs.Close
    
    If iPK > 0 Then
        If IsDate(txtEffectDateTo.Text) = True Then
            ConnOmega.Execute "UPDATE tbl_Member_Action " & _
                              " SET EffectDateTo = '" & FormatDateTime(txtEffectDateTo.Text, vbShortDate) & "' " & _
                              " WHERE (PK = " & iPK & ")"
        End If
        Select Case iActionType
            Case 1      'Assignee
                ConnOmega.Execute "UPDATE tbl_Member_Action " & _
                                  " SET AssignorKey = " & iAssignorKey & ", " & _
                                  " AssignorID = '" & Trim(txtAssignorID.Text) & "', " & _
                                  " AssignorIDSeries = " & RETURNTEXTVALUE(txtAssignorIDSeries) & " " & _
                                  " WHERE (PK = " & iPK & ")"
                
                s = "SELECT dbo.tbl_Member_IDNumber.* " & _
                    " From dbo.tbl_Member_IDNumber " & _
                    " WHERE (PK = " & iAssignorKey & ")"
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                If rs.RecordCount > 0 Then
                    ConnOmega.Execute "UPDATE tbl_Member_IDNumber " & _
                                      " SET PlayingRights = 0 " & _
                                      " WHERE (Memberkey = " & rs!MemberKey & ")"
                End If
                rs.Close
                
                s = "SELECT dbo.tbl_Member_IDNumber.* " & _
                    " From dbo.tbl_Member_IDNumber " & _
                    " WHERE (PK = " & iIDNumberKey & ")"
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                If rs.RecordCount > 0 Then
                    ConnOmega.Execute "UPDATE tbl_Member_IDNumber " & _
                                      " SET PlayingRights = 1 " & _
                                      " WHERE (Memberkey = " & rs!MemberKey & ")"
                End If
                rs.Close
            Case 2      'Share Holder
                
                s = "SELECT dbo.tbl_Member_IDNumber.* " & _
                    " From dbo.tbl_Member_IDNumber " & _
                    " WHERE (PK = " & iIDNumberKey & ")"
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                If rs.RecordCount > 0 Then
                    ConnOmega.Execute "UPDATE tbl_Member_IDNumber " & _
                                      " SET PlayingRights = 1 " & _
                                      " WHERE (Memberkey = " & rs!MemberKey & ")"
                End If
                rs.Close
                
            Case 3      'Bougth Share
                          
                ConnOmega.Execute "UPDATE tbl_Member_Action " & _
                                  " SET SoldToKey = " & iBoughtKey & ", " & _
                                  " SoldToID = '" & Trim(txtBoughtFromID.Text) & "', " & _
                                  " SoldToIDSeries = " & RETURNTEXTVALUE(txtBoughtFromIDSeries) & " " & _
                                  " WHERE (PK = " & iPK & ")"
                
                s = "SELECT dbo.tbl_Member_IDNumber.* " & _
                    " From dbo.tbl_Member_IDNumber " & _
                    " WHERE (PK = " & iBoughtKey & ")"
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                If rs.RecordCount > 0 Then
                    ConnOmega.Execute "UPDATE tbl_Member_IDNumber " & _
                                      " SET PlayingRights = 0 " & _
                                      " WHERE (Memberkey = " & rs!MemberKey & ")"
                End If
                rs.Close
                
                s = "SELECT dbo.tbl_Member_IDNumber.* " & _
                    " From dbo.tbl_Member_IDNumber " & _
                    " WHERE (PK = " & iIDNumberKey & ")"
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                If rs.RecordCount > 0 Then
                    ConnOmega.Execute "UPDATE tbl_Share_IDNumber " & _
                                      " SET MemberKey = " & rs!MemberKey & " " & _
                                      " WHERE (IDNumber = '" & Trim(txtID.Text) & "')"
                    ConnOmega.Execute "UPDATE tbl_Member_IDNumber " & _
                                      " SET PlayingRights = 1 " & _
                                      " WHERE (Memberkey = " & rs!MemberKey & ")"
                End If
                rs.Close
        End Select
    End If
End If
If TRANSACTIONTYPE = is_EDITTING Then

End If
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER sCtrl, "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
picSearch.ZOrder 0
txtSearch.Text = ""
picMain.Enabled = False
picToolbar.Enabled = False
picSearch.Visible = True
txtSearch.SetFocus
End Sub

Private Sub PRESS_F9()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picAdd.Visible = True Then cmdCancelAdd_Click: Exit Sub
    If picSearch.Visible = True Then cmdCancelSearch_Click: Exit Sub
    Unload Me
Else
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "MemberAction", "MemAction", ""), "is_LOAD"
    If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "MemberAction", "MemAction", ""), "is_HOME"
End If
End Sub


Private Sub CLEARTEXT()
iMemberKey = 0
iAssignorKey = 0
'iActionType = 0
iIDNumberKey = 0
iBoughtKey = 0
lblActionType.Caption = ""
txtCtrl.Text = ""
txtEffectDateFrom.Text = ""
txtEffectDateTo.Text = ""
txtName.Text = ""
txtAssignor.Text = ""
txtAssignorID.Text = ""
txtID.Text = ""
txtBoughtFrom.Text = ""
txtBoughtFromID.Text = ""
txtRemarks.Text = ""
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
imgPictureMovement.Picture = LoadPicture("")
imgPictureAssignor.Picture = LoadPicture("")
imgPictureBought.Picture = LoadPicture("")
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtEffectDateFrom.Locked = bln
txtEffectDateTo.Locked = bln
txtRemarks.Locked = bln
txtCtrl.Locked = True
txtName.Locked = True
txtAssignor.Locked = True
txtAssignorID.Locked = True
txtID.Locked = True
txtBoughtFrom.Locked = True
txtBoughtFromID.Locked = True
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
cmdCancelSearch_Click
End Sub

Private Sub b8TitleBar2_CLoseClick()
cmdCancelAdd_Click
End Sub

Private Sub cmdCancelAdd_Click()
picAdd.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True
End Sub

Private Sub cmdCancelSearch_Click()
picSearch.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True
End Sub

Private Sub cmdOKAdd_Click()
If lstResultAdd.ListIndex = -1 Then Exit Sub
Select Case iActionType
    Case 1
        If lstResultAssignorAdd.ListIndex = -1 Then MsgBox "Please Supply Assignor!                         ", vbCritical, "Error...": txtSearchAssignor.SetFocus: Exit Sub
        CLEARTEXT
        iMemberKey = lstResultAdd.ItemData(lstResultAdd.ListIndex)
        txtName.Text = lstResultAdd.List(lstResultAdd.ListIndex)
        imgPictureMovement.Picture = LoadPicture(SHOW_IMAGES(iMemberKey, 0, "Member"))
        iAssignorKey = lstResultAssignorAdd.ItemData(lstResultAssignorAdd.ListIndex)
        txtAssignor.Text = lstResultAssignorAdd.List(lstResultAssignorAdd.ListIndex)
        imgPictureAssignor.Picture = LoadPicture("")
        s = "SELECT IDNumber, IDCounter " & _
            " FROM tbl_Member_IDNumber " & _
            " WHERE (PK = " & iAssignorKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            txtAssignorID.Text = rs!IDNumber
            txtAssignorIDSeries.Text = rs!IDCounter
            txtID.Text = "A" & Mid(rs!IDNumber, 2, Len(rs!IDNumber))
            imgPictureAssignor.Picture = LoadPicture(SHOW_IMAGES(iAssignorKey, 0, "Member ID Number"))
        End If
        rs.Close
        LOCKTEXT False
        TOOLBARFUNC 2
        TRANSACTIONTYPE = is_ADDING
        picMain.Enabled = True
        picToolbar.Enabled = True
        picAdd.Visible = False
        txtEffectDateFrom.SetFocus
    Case 2
        CLEARTEXT
        iMemberKey = lstResultAdd.ItemData(lstResultAdd.ListIndex)
        txtName.Text = lstResultAdd.List(lstResultAdd.ListIndex)
        imgPictureMovement.Picture = LoadPicture("")
        txtID.Text = ""
        s = "SELECT IDNumber " & _
            " FROM tbl_Member_IDNumber " & _
            " WHERE (MemberKey = " & iMemberKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            txtID.Text = rs!IDNumber
            imgPictureMovement.Picture = LoadPicture(SHOW_IMAGES(iMemberKey, 0, "Member"))
        End If
        rs.Close
        LOCKTEXT False
        TOOLBARFUNC 2
        TRANSACTIONTYPE = is_ADDING
        picMain.Enabled = True
        picToolbar.Enabled = True
        picAdd.Visible = False
        txtEffectDateFrom.SetFocus
    Case 3
        CLEARTEXT
        iMemberKey = lstResultAdd.ItemData(lstResultAdd.ListIndex)
        txtName.Text = lstResultAdd.List(lstResultAdd.ListIndex)
        imgPictureMovement.Picture = LoadPicture(SHOW_IMAGES(iMemberKey, 0, "Member"))
        iBoughtKey = lstResultAssignorAdd.ItemData(lstResultAssignorAdd.ListIndex)
        txtBoughtFrom.Text = lstResultAssignorAdd.List(lstResultAssignorAdd.ListIndex)
        imgPictureBought.Picture = LoadPicture("")
        s = "SELECT IDNumber, IDCounter " & _
            " FROM tbl_Member_IDNumber " & _
            " WHERE (PK = " & iBoughtKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            txtBoughtFromID.Text = rs!IDNumber
            txtBoughtFromIDSeries.Text = rs!IDCounter
            txtID.Text = rs!IDNumber '"A" & Mid(rs!IDNumber, 2, Len(rs!IDNumber))
            imgPictureBought.Picture = LoadPicture(SHOW_IMAGES(iBoughtKey, 0, "Member ID Number"))
        End If
        rs.Close
        LOCKTEXT False
        TOOLBARFUNC 2
        TRANSACTIONTYPE = is_ADDING
        picMain.Enabled = True
        picToolbar.Enabled = True
        picAdd.Visible = False
        txtEffectDateFrom.SetFocus
End Select
End Sub

Private Sub cmdOKSearch_Click()
If lstSearchIDNumbers.ListIndex = -1 Then Exit Sub
BROWSER lstSearchIDNumbers.ItemData(lstSearchIDNumbers.ListIndex), "is_FIND"
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
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "MemberAction", "MemAction", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "MemberAction", "MemAction", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "MemberAction", "MemAction", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "MemberAction", "MemAction", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
'Me.Caption = "Member's Movement"
Me.Top = (MainForm.Height - Me.Height) / 4
Me.Left = (MainForm.Width - Me.Width) / 5
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "MemberAction", "MemAction", ""), "is_LOAD"
If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "MemberAction", "MemAction", ""), "is_HOME"

tmp = SetWindowLong(txtSearchAdd.hwnd, GWL_STYLE, GetWindowLong(txtSearchAdd.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtRemarks.hwnd, GWL_STYLE, GetWindowLong(txtRemarks.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearchAssignor.hwnd, GWL_STYLE, GetWindowLong(txtSearchAssignor.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picAdd.Visible = True Then Cancel = -1
If picSearch.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstResultAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If iActionType = 1 Then
        txtSearchAssignor.SetFocus
    Else
        cmdOKAdd_Click
    End If
'    If lstResultAdd.ListIndex = -1 Then Exit Sub
'    txtSearchAssignor.SetFocus ' cmdOKAdd_Click
End If
End Sub

Private Sub lstResultAssignorAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAdd_Click
End Sub

Private Sub lstSearch_Click()
If lstSearch.ListIndex = -1 Then lstSearchIDNumbers.Clear: Exit Sub
lstSearchIDNumbers.Clear
t = "SELECT PK, Ctrl, MemberID, ActionType, EffectDateFrom " & _
    " From dbo.tbl_Member_Action " & _
    " Where (IDNumberKey = " & lstSearch.ItemData(lstSearch.ListIndex) & ") " & _
    " ORDER BY Ctrl"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
While Not rt.EOF
    lstSearchIDNumbers.AddItem rt!MemberID & "   " & Format(rt!EffectDateFrom, "mm/dd/yyyy") & "   " & IIf(rt!ActionType = 1, "ASSIGNEE", IIf(rt!ActionType = 2, "SHARE HOLDER", IIf(rt!ActionType = 3, "BOUGTH SHARE", "")))
    lstSearchIDNumbers.ItemData(lstSearchIDNumbers.NewIndex) = rt!PK
    rt.MoveNext
Wend
rt.Close
If lstSearchIDNumbers.ListCount Then lstSearchIDNumbers.ListIndex = 0
End Sub

Private Sub lstSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstSearchIDNumbers.SetFocus
End Sub

Private Sub lstSearchIDNumbers_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKSearch_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "MemberAction", "MemAction", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "MemberAction", "MemAction", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "MemberAction", "MemAction", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "MemberAction", "MemAction", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstSearch.Clear: lstSearchIDNumbers.Clear: Exit Sub
lstSearch.Clear: lstSearchIDNumbers.Clear
's = "SELECT dbo.tbl_Member_Action.MemberKey, dbo.tbl_Member_Information.LastName, " & _
    " dbo.tbl_Member_Information.FirstName, dbo.tbl_Member_Information.MiddleName " & _
    " FROM dbo.tbl_Member_Action LEFT OUTER JOIN " & _
    " dbo.tbl_Member_Information ON dbo.tbl_Member_Action.MemberKey = dbo.tbl_Member_Information.PK " & _
    " GROUP BY dbo.tbl_Member_Action.MemberKey, dbo.tbl_Member_Information.LastName, dbo.tbl_Member_Information.FirstName, " & _
    " dbo.tbl_Member_Information.MiddleName " & _
    " HAVING (dbo.tbl_Member_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
    " ORDER BY dbo.tbl_Member_Information.LastName, dbo.tbl_Member_Information.FirstName, dbo.tbl_Member_Information.MiddleName"
s = "SELECT dbo.tbl_Member_Action.IDNumberKey, dbo.tbl_Member_Information.LastName, " & _
    " dbo.tbl_Member_Information.FirstName, dbo.tbl_Member_Information.MiddleName " & _
    " FROM dbo.tbl_Member_Action LEFT OUTER JOIN " & _
    " dbo.tbl_Member_IDNumber ON dbo.tbl_Member_Action.IDNumberKey = dbo.tbl_Member_IDNumber.PK LEFT OUTER JOIN " & _
    " dbo.tbl_Member_Information ON dbo.tbl_Member_IDNumber.MemberKey = dbo.tbl_Member_Information.PK " & _
    " GROUP BY dbo.tbl_Member_Action.IDNumberKey, dbo.tbl_Member_Information.LastName, dbo.tbl_Member_Information.FirstName, " & _
    " dbo.tbl_Member_Information.MiddleName " & _
    " HAVING (dbo.tbl_Member_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
    " ORDER BY dbo.tbl_Member_Information.LastName, dbo.tbl_Member_Information.FirstName, dbo.tbl_Member_Information.MiddleName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstSearch.AddItem rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
    lstSearch.ItemData(lstSearch.NewIndex) = rs!IDNumberKey
    rs.MoveNext
Wend
rs.Close
If lstSearch.ListCount Then lstSearch.ListIndex = 0
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstSearch.SetFocus
End Sub

Private Sub txtSearchAdd_Change()
If Trim(txtSearchAdd.Text) = "" Then lstResultAdd.Clear: txtSearchAssignor.Text = "": lstResultAssignorAdd.Clear:  Exit Sub
lstResultAdd.Clear: txtSearchAssignor.Text = "": lstResultAssignorAdd.Clear
Select Case iActionType
    Case 1      'Assignee
        s = "SELECT PK, LastName + ',  ' + FirstName + '  ' + MiddleName AS MemberName " & _
            " From dbo.tbl_Member_Information " & _
            " WHERE (NOT (PK IN (SELECT dbo.tbl_Member_IDNumber.MemberKey " & _
            " FROM dbo.tbl_Member_Action LEFT OUTER JOIN " & _
            " dbo.tbl_Member_IDNumber ON dbo.tbl_Member_Action.IDNumberKey = dbo.tbl_Member_IDNumber.PK " & _
            " WHERE (dbo.tbl_Member_Action.Status = 1) AND (dbo.tbl_Member_IDNumber.ViewNot = 0)))) " & _
            " AND (LastName LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%') " & _
            " ORDER BY MemberName"
    Case 2      'Share Holder
        s = "SELECT PK, LastName + ',  ' + FirstName + '  ' + MiddleName AS MemberName " & _
            " From dbo.tbl_Member_Information " & _
            " WHERE (NOT (PK IN (SELECT dbo.tbl_Member_IDNumber.MemberKey " & _
            " FROM dbo.tbl_Member_Action LEFT OUTER JOIN " & _
            " dbo.tbl_Member_IDNumber ON dbo.tbl_Member_Action.IDNumberKey = dbo.tbl_Member_IDNumber.PK " & _
            " WHERE (dbo.tbl_Member_Action.Status = 1) AND (dbo.tbl_Member_IDNumber.ViewNot = 0)))) " & _
            " AND (PK IN (SELECT MemberKey FROM dbo.tbl_Member_IDNumber AS tbl_Member_IDNumber_1 " & _
            " WHERE (MemberAssignor = 1))) " & _
            " AND (LastName LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%') " & _
            " ORDER BY MemberName"
    Case 3      'Bought Share
        s = "SELECT PK, LastName + ',  ' + FirstName + '  ' + MiddleName AS MemberName " & _
            " From dbo.tbl_Member_Information " & _
            " WHERE (NOT (PK IN (SELECT dbo.tbl_Member_IDNumber.MemberKey " & _
            " FROM dbo.tbl_Member_Action LEFT OUTER JOIN " & _
            " dbo.tbl_Member_IDNumber ON dbo.tbl_Member_Action.IDNumberKey = dbo.tbl_Member_IDNumber.PK " & _
            " WHERE (dbo.tbl_Member_Action.Status = 1) AND (dbo.tbl_Member_IDNumber.ViewNot = 0)))) " & _
            " AND (LastName LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%') " & _
            " ORDER BY MemberName"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultAdd.AddItem rs!MemberName
    lstResultAdd.ItemData(lstResultAdd.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResultAdd.ListCount Then lstResultAdd.ListIndex = 0
End Sub

Private Sub txtSearchAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultAdd.SetFocus
End Sub
 
Private Sub txtSearchAssignor_Change()
If Trim(txtSearchAssignor.Text) = "" Then lstResultAssignorAdd.Clear:  Exit Sub
lstResultAssignorAdd.Clear
's = "SELECT PK, LastName + ',  ' + FirstName + '  ' + MiddleName AS MemberName " & _
    " From dbo.tbl_Member_Information " & _
    " WHERE (PK IN (SELECT dbo.tbl_Member_IDNumber.MemberKey " & _
    " FROM dbo.tbl_Member_Action LEFT OUTER JOIN " & _
    " dbo.tbl_Member_IDNumber ON dbo.tbl_Member_Action.IDNumberKey = dbo.tbl_Member_IDNumber.PK " & _
    " WHERE (dbo.tbl_Member_Action.Status = 1) AND (dbo.tbl_Member_Action.ActionType = 2) AND (dbo.tbl_Member_IDNumber.ViewNot = 0) " & _
    " OR (dbo.tbl_Member_Action.Status = 1) AND (dbo.tbl_Member_Action.ActionType = 3) AND (dbo.tbl_Member_IDNumber.ViewNot = 0))) " & _
    " AND (LastName LIKE '" & FORMATSQL(Trim(txtSearchAssignor.Text)) & "%') " & _
    " ORDER BY MemberName"
s = "SELECT dbo.tbl_Member_IDNumber.PK, " & _
    " dbo.tbl_Member_Information.LastName + ',  ' + dbo.tbl_Member_Information.FirstName + '  ' + dbo.tbl_Member_Information.MiddleName AS MemberName " & _
    " FROM dbo.tbl_Member_IDNumber LEFT OUTER JOIN " & _
    " dbo.tbl_Member_Information ON dbo.tbl_Member_IDNumber.MemberKey = dbo.tbl_Member_Information.PK " & _
    " WHERE (dbo.tbl_Member_IDNumber.ViewNot = 0) AND (dbo.tbl_Member_Information.PK IN " & _
    " (SELECT tbl_Member_IDNumber_1.MemberKey " & _
    " FROM dbo.tbl_Member_Action LEFT OUTER JOIN " & _
    " dbo.tbl_Member_IDNumber AS tbl_Member_IDNumber_1 ON dbo.tbl_Member_Action.IDNumberKey = tbl_Member_IDNumber_1.PK " & _
    " WHERE (dbo.tbl_Member_Action.Status = 1) AND (dbo.tbl_Member_Action.ActionType = 2) AND (tbl_Member_IDNumber_1.ViewNot = 0) OR " & _
    " (dbo.tbl_Member_Action.Status = 1) AND (dbo.tbl_Member_Action.ActionType = 3) AND (tbl_Member_IDNumber_1.ViewNot = 0))) AND " & _
    " (dbo.tbl_Member_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearchAssignor.Text)) & "%') " & _
    " AND (dbo.tbl_Member_IDNumber.MemberType = 1) " & _
    " AND (dbo.tbl_Member_IDNumber.PlayingRights <> 0)" & _
    " ORDER BY MemberName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultAssignorAdd.AddItem rs!MemberName
    lstResultAssignorAdd.ItemData(lstResultAssignorAdd.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResultAssignorAdd.ListCount Then lstResultAssignorAdd.ListIndex = 0
End Sub

Private Sub txtSearchAssignor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultAssignorAdd.SetFocus
End Sub
