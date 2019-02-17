VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOperationBagDrop 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11385
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBagTag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15600
      TabIndex        =   60
      Top             =   0
      Width           =   15600
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   61
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
               Caption         =   " Post   "
               Key             =   "Post"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               Key             =   "Refresh"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         MousePointer    =   99
         MouseIcon       =   "frmBagTag.frx":0CCA
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   9900
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   62
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmBagTag.frx":0FE4
               Top             =   120
               Visible         =   0   'False
               Width           =   1395
            End
         End
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
   Begin RPVGCC.b8Container picAddSearchMember 
      Height          =   3375
      Left            =   3360
      TabIndex        =   51
      Top             =   1080
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5953
      BackColor       =   15396057
      Begin VB.ListBox lstResulMember 
         Height          =   1815
         Left            =   120
         TabIndex        =   55
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtSearchMember 
         Height          =   315
         Left            =   120
         TabIndex        =   54
         Top             =   480
         Width           =   4215
      End
      Begin VB.CommandButton cmdCancelMember 
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
         Picture         =   "frmBagTag.frx":16F7
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   2745
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKMember 
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
         Picture         =   "frmBagTag.frx":1E53
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   2745
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar4 
         Height          =   345
         Left            =   45
         TabIndex        =   56
         Top             =   45
         Width           =   4365
         _ExtentX        =   7699
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
         Icon            =   "frmBagTag.frx":24C5
         ShadowVisible   =   0   'False
      End
      Begin RPVGCC.b8TitleBar b8TitleBar5 
         Height          =   345
         Left            =   40
         TabIndex        =   57
         Top             =   40
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   609
         Caption         =   "Search Member"
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
         Icon            =   "frmBagTag.frx":2A5F
         ShadowVisible   =   0   'False
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   4785
      Width           =   11385
      _ExtentX        =   20082
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11280
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
            Picture         =   "frmBagTag.frx":2FF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBagTag.frx":3CD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBagTag.frx":49AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBagTag.frx":5687
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBagTag.frx":6361
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBagTag.frx":703B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBagTag.frx":7D15
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBagTag.frx":89EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBagTag.frx":96C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBagTag.frx":9FA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBagTag.frx":AC7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBagTag.frx":B957
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBagTag.frx":C631
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBagTag.frx":D30B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBagTag.frx":DFE5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   600
      ScaleHeight     =   3375
      ScaleWidth      =   10335
      TabIndex        =   1
      Top             =   1200
      Width           =   10335
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   6960
         ScaleHeight     =   3225
         ScaleWidth      =   3225
         TabIndex        =   63
         Top             =   0
         Width           =   3255
         Begin VB.Image imgPicture 
            Height          =   3225
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   3225
         End
      End
      Begin VB.TextBox txtMemberName 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   58
         Top             =   1080
         Width           =   5775
      End
      Begin VB.TextBox txtCaddyNo 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   36
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   2640
         MaxLength       =   100
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtBagTagNo 
         Height          =   315
         Left            =   5160
         MaxLength       =   100
         TabIndex        =   14
         Top             =   0
         Width           =   1695
      End
      Begin VB.PictureBox picDetail 
         BackColor       =   &H00C6B8A4&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   600
         ScaleHeight     =   855
         ScaleWidth      =   5655
         TabIndex        =   12
         Top             =   2520
         Width           =   5655
         Begin VB.TextBox txtNoOthers 
            Height          =   315
            Left            =   2160
            MaxLength       =   100
            TabIndex        =   26
            Top             =   360
            Width           =   3375
         End
         Begin VB.TextBox txtNoUmbrella 
            Height          =   315
            Left            =   720
            MaxLength       =   100
            TabIndex        =   24
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtNoPutter 
            Height          =   315
            Left            =   4800
            MaxLength       =   100
            TabIndex        =   22
            Top             =   0
            Width           =   735
         End
         Begin VB.TextBox txtNoIron 
            Height          =   315
            Left            =   3480
            MaxLength       =   100
            TabIndex        =   20
            Top             =   0
            Width           =   735
         End
         Begin VB.TextBox txtNoWood 
            Height          =   315
            Left            =   2160
            MaxLength       =   100
            TabIndex        =   18
            Top             =   0
            Width           =   735
         End
         Begin VB.TextBox txtNoDriver 
            Height          =   315
            Left            =   720
            MaxLength       =   100
            TabIndex        =   16
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Others"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1560
            TabIndex        =   27
            Top             =   390
            Width           =   495
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Umbrella"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   25
            Top             =   390
            Width           =   615
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Putter"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4320
            TabIndex        =   23
            Top             =   30
            Width           =   495
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Iron"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3000
            TabIndex        =   21
            Top             =   30
            Width           =   495
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Wood"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1560
            TabIndex        =   19
            Top             =   30
            Width           =   495
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Driver"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   30
            Width           =   495
         End
      End
      Begin VB.ComboBox cmbPlayerType 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   720
         Width           =   5775
      End
      Begin VB.TextBox txtCaddyName 
         Height          =   315
         Left            =   2040
         MaxLength       =   100
         TabIndex        =   9
         Top             =   1800
         Width           =   4815
      End
      Begin VB.TextBox txtPorterName 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   7
         Top             =   1440
         Width           =   5775
      End
      Begin VB.TextBox txtPlayerName 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   3
         Top             =   360
         Width           =   5775
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   2
         Top             =   0
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         Height          =   3255
         Left            =   7035
         Top             =   75
         Width           =   3255
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Guest Of"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   59
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Your bag contains the following:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   6495
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Bag Tag #"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4320
         TabIndex        =   15
         Top             =   30
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Caddy"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   1830
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Porter"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   1470
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Player Type"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Player Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   390
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   30
         Width           =   1095
      End
   End
   Begin RPVGCC.b8Container picAdd 
      Height          =   4575
      Left            =   480
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8070
      BackColor       =   15396057
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   4440
         ScaleHeight     =   3225
         ScaleWidth      =   3225
         TabIndex        =   64
         Top             =   480
         Width           =   3255
         Begin VB.Image imgPictureSearchAdd 
            Height          =   3225
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   3225
         End
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
         Left            =   2280
         Picture         =   "frmBagTag.frx":ECBF
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3945
         Width           =   1560
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
         Left            =   3960
         Picture         =   "frmBagTag.frx":F331
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3945
         Width           =   1560
      End
      Begin VB.TextBox txtAdd 
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   4215
      End
      Begin VB.ListBox lstAdd 
         Height          =   2985
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   4215
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   34
         Top             =   45
         Width           =   7845
         _ExtentX        =   13838
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
         Icon            =   "frmBagTag.frx":FA8D
         ShadowVisible   =   0   'False
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   40
         TabIndex        =   35
         Top             =   40
         Width           =   4365
         _ExtentX        =   7699
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
         Icon            =   "frmBagTag.frx":10027
         ShadowVisible   =   0   'False
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         Height          =   3255
         Left            =   4515
         Top             =   555
         Width           =   3255
      End
   End
   Begin VB.PictureBox picAddPlayer 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   1680
      ScaleHeight     =   2535
      ScaleWidth      =   7335
      TabIndex        =   37
      Top             =   1440
      Visible         =   0   'False
      Width           =   7335
      Begin RPVGCC.b8Container picAddPlayer1 
         Height          =   2415
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   4260
         BackColor       =   15396057
         Begin VB.TextBox txtPlayerNameAdd3 
            Height          =   315
            Left            =   5160
            MaxLength       =   100
            TabIndex        =   50
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txtPlayerNameAdd2 
            Height          =   315
            Left            =   3000
            MaxLength       =   100
            TabIndex        =   49
            Top             =   480
            Width           =   2055
         End
         Begin VB.ComboBox cmbPlayerTypeAdd 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   840
            Width           =   5775
         End
         Begin VB.TextBox txtPlayerNameAdd1 
            Height          =   315
            Left            =   1200
            MaxLength       =   100
            TabIndex        =   43
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtMemberAdd 
            Height          =   315
            Left            =   1200
            MaxLength       =   100
            TabIndex        =   42
            Top             =   1200
            Width           =   5490
         End
         Begin VB.CommandButton cmdOKAddAdd 
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
            Picture         =   "frmBagTag.frx":105C1
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   1680
            Width           =   1560
         End
         Begin VB.CommandButton cmdCancelAddAdd 
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
            Left            =   3960
            Picture         =   "frmBagTag.frx":10C33
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   1680
            Width           =   1560
         End
         Begin VB.CommandButton cmdAddSearchMember 
            Caption         =   ".."
            Height          =   315
            Left            =   6690
            TabIndex        =   39
            Top             =   1200
            Width           =   255
         End
         Begin RPVGCC.b8TitleBar b8TitleBar3 
            Height          =   345
            Left            =   45
            TabIndex        =   45
            Top             =   45
            Width           =   7005
            _ExtentX        =   12356
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
            Icon            =   "frmBagTag.frx":1138F
            ShadowVisible   =   0   'False
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Player Type"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Player Name"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   510
            Width           =   1095
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Member"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   1230
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmOperationBagDrop"
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

Dim iMemberKey, MemberType, iPlayerType, Arr, iCaddyKey, iPK, sBagTag

Private Sub BROWSER(sBagTagNo, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If sBagTagNo <> "" Then
            s = "SELECT TOP 1 tbl_Operation_Passport.* " & _
                " FROM tbl_Operation_Passport " & _
                " WHERE (BagTagNo = '" & sBagTagNo & "') " & _
                " ORDER BY BagTagNo"
        Else
            s = "SELECT TOP 1 tbl_Operation_Passport.* " & _
                " FROM tbl_Operation_Passport " & _
                " ORDER BY BagTagNo"
        End If
    Case "is_HOME"
        If picAddSearchMember.Visible = True Then Exit Sub
        If picAdd.Visible = True Then Exit Sub
        If picAddPlayer.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Operation_Passport.* " & _
            " FROM tbl_Operation_Passport " & _
            " ORDER BY BagTagNo"
    Case "is_PAGEUP"
        If picAddSearchMember.Visible = True Then Exit Sub
        If picAdd.Visible = True Then Exit Sub
        If picAddPlayer.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Operation_Passport.* " & _
            " FROM tbl_Operation_Passport " & _
            " WHERE (BagTagNo < '" & sBagTagNo & "') " & _
            " ORDER BY BagTagNo DESC"
    Case "is_PAGEDOWN"
        If picAddSearchMember.Visible = True Then Exit Sub
        If picAdd.Visible = True Then Exit Sub
        If picAddPlayer.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Operation_Passport.* " & _
            " FROM tbl_Operation_Passport " & _
            " WHERE (BagTagNo > '" & sBagTagNo & "') " & _
            " ORDER BY BagTagNo"
    Case "is_END"
        If picAddSearchMember.Visible = True Then Exit Sub
        If picAdd.Visible = True Then Exit Sub
        If picAddPlayer.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Operation_Passport.* " & _
            " FROM tbl_Operation_Passport " & _
            " ORDER BY BagTagNo DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iCaddyKey = rs!CaddyKey
    iMemberKey = IIf(IsNull(rs!MemberKey), 0, rs!MemberKey)
    iPlayerType = rs!PlayerTypeKey
    txtDate.Text = Format(rs!DateAll, "mm/dd/yyyy")
    txtBagTagNo.Text = rs!BagTagNo
    txtPlayerName.Text = rs!PlayerName
    txtMemberName.Text = IIf(IsNull(rs!GuestOf), "", rs!GuestOf)
    imgPicture.Picture = LoadPicture(SHOW_IMAGES(IIf(IsNull(rs!MemberKey), 0, rs!MemberKey), 0, "Member ID Number"))
    
    txtCaddyNo.Text = ""
    txtCaddyName.Text = ""
    t = "SELECT tbl_Caddy_Information.* " & _
        " FROM tbl_Caddy_Information " & _
        " WHERE (PK = " & rs!CaddyKey & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtCaddyNo.Text = rt!CaddyNo
        txtCaddyName.Text = rt!CaddyLName & ",  " & rt!CaddyFName & "  " & rt!CaddyMName
    End If
    rt.Close
    
    cmbPlayerType.ListIndex = rs!PlayerTypeKey - 1
    imgPosted.Visible = IIf(rs!PostedBagDrop = 1, True, False)
    Statusbar1.Panels(1).Text = rs!PK
    
    txtNoDriver.Text = ""
    txtNoWood.Text = ""
    txtNoIron.Text = ""
    txtNoPutter.Text = ""
    txtNoUmbrella.Text = ""
    txtNoOthers.Text = ""
    Statusbar1.Panels(2).Text = ""
    t = "SELECT tbl_Operation_Passport_BagDrop.* " & _
        " FROM tbl_Operation_Passport_BagDrop " & _
        " WHERE (PassportKey = " & rs!PK & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtPorterName.Text = rt!Porter
        txtNoDriver.Text = rt!Driver
        txtNoWood.Text = rt!Wood
        txtNoIron.Text = rt!Iron
        txtNoPutter.Text = rt!Putter
        txtNoUmbrella.Text = rt!Umbrella
        txtNoOthers.Text = rt!Others
        Statusbar1.Panels(2).Text = IIf(IsNull(rt!LastModified), "", rt!LastModified)
    End If
    rt.Close
    
    SaveSetting App.EXEName, "BagDropNum", "BagDropNum", rs!BagTagNo
    
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If picAddSearchMember.Visible = True Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picAddPlayer.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If AccessRights("Bag Drop", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
picMain.Enabled = False
picToolbar.Enabled = False
picAdd.ZOrder 0
txtAdd.Text = ""
picAdd.Visible = True
txtAdd.SetFocus
End Sub

Private Sub PRESS_F2()
If picAddSearchMember.Visible = True Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picAddPlayer.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If imgPosted.Visible = True Then MsgBox "Already Posted!                ", vbCritical, "Error...": Exit Sub
If AccessRights("Bag Drop", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_EDITTING
End Sub

Private Sub PRESS_DELETE()
If picAddSearchMember.Visible = True Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picAddPlayer.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If imgPosted.Visible = True Then MsgBox "Already Posted!                ", vbCritical, "Error...": Exit Sub
If AccessRights("Bag Drop", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
If MsgBox("ARE YOU SURE IN DELETING THIS RECORD?                    ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_Operation_Passport_BagDrop WHERE (PassportKey = " & Statusbar1.Panels(1).Text & ")"
ConnOmega.Execute "DELETE FROM tbl_Operation_Passport WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "BagDropNum", "BagDropNum", ""), "is_PAGEDOWN"
If Trim(txtBagTagNo.Text) = "" Then BROWSER GetSetting(App.EXEName, "BagDropNum", "BagDropNum", ""), "is_HOME"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If picAddSearchMember.Visible = True Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picAddPlayer.Visible = True Then Exit Sub
If IsDate(txtDate.Text) = False Then MsgBox "Please Supply a Valid Date!                  ", vbCritical, "Error...": txtDate.SetFocus: Exit Sub
If Trim(txtBagTagNo.Text) = "" Then MsgBox "Please Supply Bag Tag No.!                    ", vbCritical, "Error...": txtBagTagNo.SetFocus: Exit Sub
If Trim(txtCaddyNo.Text) = "" Then MsgBox "Please Supply Caddy!                          ", vbCritical, "Error...": txtCaddyNo.SetFocus: Exit Sub

txtCaddyName.Text = "": iCaddyKey = 0
t = "SELECT PK, CaddyNo, CaddyLName, CaddyFName, CaddyMName " & _
    " From dbo.tbl_Caddy_Information " & _
    " WHERE (CaddyNo = '" & Trim(txtCaddyNo.Text) & "')"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then
    iCaddyKey = rt!PK
    txtCaddyName.Text = rt!CaddyLName & ", " & rt!CaddyFName & " " & rt!CaddyMName
End If
rt.Close
If CDbl(iCaddyKey) = 0 Then MsgBox "Please Supply Caddy!                    ", vbCritical, "Error...": txtCaddyNo.SetFocus: Exit Sub
sBagTag = Trim(txtBagTagNo.Text)
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    ConnOmega.Execute "INSERT INTO tbl_Operation_Passport " & _
                      " (DateAll, BagTagNo, PlayerTypeKey, CaddyKey, PlayerName) " & _
                      " VALUES ('" & FormatDateTime(txtDate.Text, vbShortDate) & "', " & _
                      " '" & Trim(txtBagTagNo.Text) & "', " & iPlayerType & ", " & _
                      " " & iCaddyKey & ", '" & FORMATSQL(Trim(txtPlayerName.Text)) & "')"
    iPK = 0
    s = "SELECT tbl_Operation_Passport.* " & _
        " FROM tbl_Operation_Passport " & _
        " WHERE (BagTagNo = '" & Trim(txtBagTagNo.Text) & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        iPK = rs!PK
    End If
    rs.Close
    If CDbl(iMemberKey) <> 0 Then
        ConnOmega.Execute "UPDATE tbl_Operation_Passport " & _
                          " SET MemberKey = " & iMemberKey & ", " & _
                          " GuestOf = '" & FORMATSQL(Trim(txtMemberName.Text)) & "' " & _
                          " WHERE (PK = " & iPK & ")"
    End If
End If
If TRANSACTIONTYPE = is_EDITTING Then
    iPK = Statusbar1.Panels(1).Text
    ConnOmega.Execute "UPDATE tbl_Operation_Passport " & _
                      " SET DateAll = '" & FormatDateTime(txtDate.Text, vbShortDate) & "', " & _
                      " BagTagNo = '" & Trim(txtBagTagNo.Text) & "', " & _
                      " CaddyKey = " & iCaddyKey & ", " & _
                      " PlayerName = '" & FORMATSQL(Trim(txtPlayerName.Text)) & "' " & _
                      " WHERE (PK = " & iPK & ")"
End If
ConnOmega.Execute "DELETE FROM tbl_Operation_Passport_BagDrop WHERE (PassportKey = " & iPK & ")"
ConnOmega.Execute "INSERT INTO tbl_Operation_Passport_BagDrop " & _
                  " (PassportKey, Porter, Driver, Wood, Iron, Putter, Umbrella, Others, LastModified) " & _
                  " VALUES (" & iPK & ", '" & FORMATSQL(Trim(txtPorterName.Text)) & "', " & _
                  " " & RETURNTEXTVALUE(txtNoDriver) & ", " & RETURNTEXTVALUE(txtNoWood) & ", " & _
                  " " & RETURNTEXTVALUE(txtNoIron) & ", " & RETURNTEXTVALUE(txtNoPutter) & ", " & _
                  " " & RETURNTEXTVALUE(txtNoUmbrella) & ", '" & FORMATSQL(Trim(txtNoOthers.Text)) & "', " & _
                  " '" & CStr(Now) & " - " & gbl_CompleteName & "')"

CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER sBagTag, "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If picAddSearchMember.Visible = True Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picAddPlayer.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
End Sub

Private Sub PRESS_F8()
If picAddSearchMember.Visible = True Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picAddPlayer.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("Bag Drop", "Post") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
If MsgBox("ARE YOU SURE IN POSTING THIS TRANSACTION?                    ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
On Error GoTo PG:
ConnOmega.Execute "UPDATE tbl_Operation_Passport " & _
                  " SET PostedBagDrop = 1 " & _
                  " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
                  
ConnOmega.Execute "UPDATE tbl_Operation_Passport_BagDrop " & _
                  " SET LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                  " WHERE (PassportKey = " & Statusbar1.Panels(1).Text & ")"
    
BROWSER GetSetting(App.EXEName, "BagDropNum", "BagDropNum", ""), "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error...'"
Exit Sub
End Sub

Private Sub PRESS_F9()
If picAddSearchMember.Visible = True Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picAddPlayer.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picAddSearchMember.Visible = True Then cmdCancelMember_Click: Exit Sub
    If picAdd.Visible = True Then cmdCancelAdd_Click: Exit Sub
    If picAddPlayer.Visible = True Then cmdCancelAddAdd_Click: Exit Sub
    Unload Me
Else
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "BagDropNum", "BagDropNum", ""), "is_LOAD"
    If Trim(txtBagTagNo.Text) = "" Then BROWSER GetSetting(App.EXEName, "BagDropNum", "BagDropNum", ""), "is_HOME"
End If
End Sub

Private Sub CLEARTEXT()
iCaddyKey = 0
iMemberKey = 0
iPlayerType = 0
txtDate.Text = ""
txtBagTagNo.Text = ""
txtPlayerName.Text = ""
txtMemberName.Text = ""
txtPorterName.Text = ""
txtCaddyNo.Text = ""
txtCaddyName.Text = ""
txtNoDriver.Text = ""
txtNoWood.Text = ""
txtNoIron.Text = ""
txtNoPutter.Text = ""
txtNoUmbrella.Text = ""
txtNoOthers.Text = ""
cmbPlayerType.ListIndex = -1
imgPosted.Visible = False
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
imgPicture.Picture = LoadPicture("")
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtMemberAdd.Locked = True
txtDate.Locked = bln
txtBagTagNo.Locked = bln
txtPlayerName.Locked = bln
txtPorterName.Locked = bln 'True
txtCaddyNo.Locked = bln
txtCaddyName.Locked = True
txtMemberName.Locked = True
txtNoDriver.Locked = bln
txtNoWood.Locked = bln
txtNoIron.Locked = bln
txtNoPutter.Locked = bln
txtNoUmbrella.Locked = bln
txtNoOthers.Locked = bln
cmbPlayerType.Locked = bln
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
    End Select
End With
End Sub

Private Sub b8TitleBar1_CLoseClick()
cmdCancelAdd_Click
End Sub

Private Sub b8TitleBar2_CLoseClick()
cmdCancelAdd_Click
End Sub

Private Sub b8TitleBar3_CLoseClick()
cmdCancelAddAdd_Click
End Sub

Private Sub b8TitleBar5_CLoseClick()
cmdCancelMember_Click
End Sub

Private Sub cmbPlayerTypeAdd_Click()
If cmbPlayerTypeAdd.ListIndex = -1 Then Exit Sub
iPlayerType = cmbPlayerTypeAdd.ItemData(cmbPlayerTypeAdd.ListIndex)
End Sub

Private Sub cmdAddSearchMember_Click()
txtMemberAdd.SetFocus
If cmbPlayerTypeAdd.ListIndex = -1 Then Exit Sub
If cmbPlayerTypeAdd.ItemData(cmbPlayerTypeAdd.ListIndex) = 2 Or _
cmbPlayerTypeAdd.ItemData(cmbPlayerTypeAdd.ListIndex) = 3 Then
    picAddSearchMember.ZOrder 0
    txtSearchMember.Text = ""
    picAddPlayer.Enabled = False
    picAddSearchMember.Visible = True
    txtSearchMember.SetFocus
End If
End Sub

Private Sub cmdCancelAdd_Click()
picAdd.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True
End Sub

Private Sub cmdCancelAddAdd_Click()
picAddPlayer.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True
End Sub

Private Sub cmdCancelMember_Click()
picAddPlayer.Enabled = True
picAddSearchMember.Visible = False
txtMemberAdd.SetFocus
End Sub

Private Sub cmdOKAdd_Click()
If lstAdd.ListIndex = -1 Then GoTo Add_as_Guest:
CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
txtPlayerName.Locked = True
cmbPlayerType.Locked = True
Arr = Split(lstAdd.List(lstAdd.ListIndex), " - ", -1, 1)
iMemberKey = lstAdd.ItemData(lstAdd.ListIndex)
imgPicture.Picture = LoadPicture(SHOW_IMAGES(iMemberKey, 0, "Member ID Number"))
txtPlayerName.Text = CStr(Arr(1))
txtDate.Text = Format(FormatDateTime(Date, vbShortDate), "mm/dd/yyyy")
cmbPlayerType.ListIndex = 0
'txtPorterName.Text = gbl_CompleteName
iPlayerType = 1
cmdCancelAdd_Click
txtDate.SetFocus
Exit Sub
Add_as_Guest:
picAdd.Visible = False
picAddPlayer.Width = 7095
picAddPlayer.Height = 2415
picAddPlayer.ZOrder 0
txtPlayerNameAdd1.Text = ""
txtPlayerNameAdd2.Text = ""
txtPlayerNameAdd3.Text = ""
cmbPlayerTypeAdd.ListIndex = -1
txtMemberAdd.Text = ""
picAddPlayer.Visible = True
txtPlayerNameAdd1.SetFocus
End Sub

Private Sub cmdOKAddAdd_Click()
If Trim(txtPlayerNameAdd1.Text) = "" Then MsgBox "Please Supply LastName!                 ", vbCritical, "Error...": txtPlayerNameAdd1.SetFocus: Exit Sub
If Trim(txtPlayerNameAdd2.Text) = "" Then MsgBox "Please Supply FirstName!                    ", vbCritical, "Error...": txtPlayerNameAdd2.SetFocus: Exit Sub
iPlayerType = cmbPlayerTypeAdd.ItemData(cmbPlayerTypeAdd.ListIndex)
If CDbl(iPlayerType) = 2 Or _
CDbl(iPlayerType) = 3 Then
    If CDbl(iMemberKey) = 0 Then MsgBox "Please Supply Member!                  ", vbCritical, "Error...": txtMemberAdd.SetFocus: Exit Sub
End If
CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
txtDate.Text = Format(FormatDateTime(Date, vbShortDate), "mm/dd/yyyy")
txtPlayerName.Text = Trim(txtPlayerNameAdd1.Text) & ", " & txtPlayerNameAdd2.Text & " " & txtPlayerNameAdd3.Text
txtMemberName.Text = Trim(txtMemberAdd.Text)
cmbPlayerType.ListIndex = cmbPlayerTypeAdd.ListIndex + 1
'txtPorterName.Text = gbl_CompleteName
cmdCancelAddAdd_Click
txtDate.SetFocus
End Sub

Private Sub cmdOKMember_Click()
If lstResulMember.ListIndex = -1 Then Exit Sub
Arr = Split(lstResulMember.List(lstResulMember.ListIndex), " - ", -1, 1)
txtMemberAdd.Text = CStr(Arr(1))
iMemberKey = lstResulMember.ItemData(lstResulMember.ListIndex)
cmdCancelMember_Click
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
    Case vbKeyF8::      PRESS_F8
    Case vbKeyF9:       PRESS_F9
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "BagDropNum", "BagDropNum", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "BagDropNum", "BagDropNum", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "BagDropNum", "BagDropNum", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "BagDropNum", "BagDropNum", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.Height - Me.Height) / 3
Me.Left = (MainForm.Width - Me.Width) / 5
'cmbPlayerType
POPULATE_COMBO "PK", "PlayerType", "tbl_Operation_PlayerType", "PK", cmbPlayerType
POPULATE_COMBO_EXEMPTION "PK", "PlayerType", "tbl_Operation_PlayerType", "PK", "Visible", "" & 1 & "", cmbPlayerTypeAdd
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "BagDropNum", "BagDropNum", ""), "is_LOAD"
If Trim(txtBagTagNo.Text) = "" Then BROWSER GetSetting(App.EXEName, "BagDropNum", "BagDropNum", ""), "is_HOME"
tmp = SetWindowLong(txtAdd.hwnd, GWL_STYLE, GetWindowLong(txtAdd.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPlayerName.hwnd, GWL_STYLE, GetWindowLong(txtPlayerName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPorterName.hwnd, GWL_STYLE, GetWindowLong(txtPorterName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtCaddyNo.hwnd, GWL_STYLE, GetWindowLong(txtCaddyNo.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtCaddyName.hwnd, GWL_STYLE, GetWindowLong(txtCaddyName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPlayerNameAdd1.hwnd, GWL_STYLE, GetWindowLong(txtPlayerNameAdd1.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPlayerNameAdd2.hwnd, GWL_STYLE, GetWindowLong(txtPlayerNameAdd2.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPlayerNameAdd3.hwnd, GWL_STYLE, GetWindowLong(txtPlayerNameAdd3.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearchMember.hwnd, GWL_STYLE, GetWindowLong(txtSearchMember.hwnd, GWL_STYLE) Or ES_UPPERCASE)

End Sub


Private Sub Form_Unload(Cancel As Integer)
If picAddSearchMember.Visible = True Then Cancel = -1
If picAdd.Visible = True Then Cancel = -1
If picAddPlayer.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstAdd_Click()
If lstAdd.ListIndex = -1 Then imgPictureSearchAdd.Picture = LoadPicture(""): Exit Sub
imgPictureSearchAdd.Picture = LoadPicture(SHOW_IMAGES(lstAdd.ItemData(lstAdd.ListIndex), 0, "Member ID Number"))
End Sub

Private Sub lstAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAdd_Click
End Sub


Private Sub lstResulMember_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKMember_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "BagDropNum", "BagDropNum", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "BagDropNum", "BagDropNum", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "BagDropNum", "BagDropNum", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "BagDropNum", "BagDropNum", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Post":    PRESS_F8
    Case "Print":   PRESS_F9
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtAdd_Change()
If Trim(txtAdd.Text) = "" Then lstAdd.Clear: imgPictureSearchAdd.Picture = LoadPicture(""): Exit Sub
lstAdd.Clear: imgPictureSearchAdd.Picture = LoadPicture("")
s = "SELECT dbo.tbl_Member_IDNumber.* " & _
    " From dbo.tbl_Member_IDNumber " & _
    " WHERE (MemberName LIKE '" & FORMATSQL(Trim(txtAdd.Text)) & "%') " & _
    " AND (PlayingRights = 1) " & _
    " ORDER BY MemberName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstAdd.AddItem rs!IDNumber & " - " & rs!MemberName
    lstAdd.ItemData(lstAdd.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstAdd.ListCount Then lstAdd.ListIndex = 0
End Sub

Private Sub txtAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstAdd.SetFocus
End Sub

Private Sub txtCaddyNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtNoDriver.SetFocus
End If
End Sub

Private Sub txtCaddyNo_LostFocus()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    txtCaddyName.Text = "": iCaddyKey = 0
    t = "SELECT PK, CaddyNo, CaddyLName, CaddyFName, CaddyMName " & _
        " From dbo.tbl_Caddy_Information " & _
        " WHERE (CaddyNo = '" & Trim(txtCaddyNo.Text) & "') "
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        iCaddyKey = rt!PK
        txtCaddyName.Text = rt!CaddyLName & ", " & rt!CaddyFName & " " & rt!CaddyMName
    End If
    rt.Close
End If
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtBagTagNo.SetFocus
End Sub

Private Sub txtSearchMember_Change()
If Trim(txtSearchMember.Text) = "" Then lstResulMember.Clear: Exit Sub
lstResulMember.Clear
s = "SELECT dbo.tbl_Member_IDNumber.* " & _
    " From dbo.tbl_Member_IDNumber " & _
    " WHERE (MemberName LIKE '" & FORMATSQL(Trim(txtSearchMember.Text)) & "%') " & _
    " AND (PlayingRights = 1) " & _
    " ORDER BY MemberName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResulMember.AddItem rs!IDNumber & " - " & rs!MemberName
    lstResulMember.ItemData(lstResulMember.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResulMember.ListCount Then lstResulMember.ListIndex = 0
End Sub

Private Sub txtSearchMember_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResulMember.SetFocus
End Sub
