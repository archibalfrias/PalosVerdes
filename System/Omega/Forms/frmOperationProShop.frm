VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOperationProShop 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11355
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOperationProShop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picAddProShop 
      Height          =   3375
      Left            =   2880
      TabIndex        =   25
      Top             =   1200
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5953
      BackColor       =   15396057
      Begin VB.ListBox lstResultProShop 
         Height          =   1815
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtSearchProShop 
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   4215
      End
      Begin VB.CommandButton cmdCancelProShop 
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
         Picture         =   "frmOperationProShop.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2745
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKProShop 
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
         Picture         =   "frmOperationProShop.frx":1026
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2745
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar5 
         Height          =   345
         Left            =   40
         TabIndex        =   30
         Top             =   40
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   609
         Caption         =   "Search Passport"
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
         Icon            =   "frmOperationProShop.frx":1698
         ShadowVisible   =   0   'False
      End
   End
   Begin RPVGCC.b8Container picSearchItem 
      Height          =   3375
      Left            =   1800
      TabIndex        =   38
      Top             =   1680
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5953
      BackColor       =   15396057
      Begin VB.CommandButton cmdOKSearchItem 
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
         Picture         =   "frmOperationProShop.frx":1C32
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2745
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelSearchItem 
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
         Picture         =   "frmOperationProShop.frx":22A4
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   2745
         Width           =   1560
      End
      Begin VB.TextBox txtSearchItem 
         Height          =   315
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   4215
      End
      Begin VB.ListBox lstResultSearchItem 
         Height          =   1815
         Left            =   120
         TabIndex        =   39
         Top             =   840
         Width           =   4215
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   40
         TabIndex        =   43
         Top             =   40
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   609
         Caption         =   "Search Item"
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
         Icon            =   "frmOperationProShop.frx":2A00
         ShadowVisible   =   0   'False
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15600
      TabIndex        =   47
      Top             =   0
      Width           =   15600
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   48
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
               ImageKey        =   "IMG1"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Edit"
               Key             =   "Edit"
               ImageKey        =   "IMG2"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Delete"
               Key             =   "Delete"
               ImageKey        =   "IMG3"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "First"
               Key             =   "First"
               ImageKey        =   "IMG4"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Back"
               Key             =   "Back"
               ImageKey        =   "IMG5"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Next"
               Key             =   "Next"
               ImageKey        =   "IMG6"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Last"
               Key             =   "Last"
               ImageKey        =   "IMG7"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Find"
               Key             =   "Find"
               ImageKey        =   "IMG8"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Print"
               Key             =   "Print"
               ImageKey        =   "IMG9"
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   " Post   "
               Key             =   "Post"
               ImageKey        =   "IMG10"
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               Key             =   "Refresh"
               ImageKey        =   "IMG12"
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
               ImageKey        =   "IMG13"
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         MousePointer    =   99
         MouseIcon       =   "frmOperationProShop.frx":2F9A
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   9900
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   49
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmOperationProShop.frx":32B4
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
   Begin VB.PictureBox picSLine 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1320
      ScaleHeight     =   855
      ScaleWidth      =   9015
      TabIndex        =   13
      Top             =   4680
      Visible         =   0   'False
      Width           =   9015
      Begin RPVGCC.b8Container picSline1 
         Height          =   855
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   1508
         BackColor       =   8438015
         Begin VB.TextBox txtTotalSRP1 
            Height          =   315
            Left            =   3960
            MaxLength       =   100
            TabIndex        =   37
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtSRP1 
            Height          =   315
            Left            =   3720
            MaxLength       =   100
            TabIndex        =   36
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtQty1 
            Height          =   315
            Left            =   3480
            MaxLength       =   100
            TabIndex        =   35
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtDescription1 
            Height          =   315
            Left            =   3240
            MaxLength       =   100
            TabIndex        =   34
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtCode1 
            Height          =   315
            Left            =   3000
            MaxLength       =   100
            TabIndex        =   33
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtItemKey1 
            Height          =   315
            Left            =   2760
            MaxLength       =   100
            TabIndex        =   32
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtItemKey 
            Height          =   315
            Left            =   720
            MaxLength       =   100
            TabIndex        =   31
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtTotalSRP 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7320
            MaxLength       =   100
            TabIndex        =   23
            Text            =   "10,000.00"
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtSRP 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6120
            MaxLength       =   100
            TabIndex        =   21
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtQty 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5160
            MaxLength       =   100
            TabIndex        =   19
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtDescription 
            Height          =   315
            Left            =   1200
            MaxLength       =   100
            TabIndex        =   17
            Top             =   360
            Width           =   3855
         End
         Begin VB.TextBox txtCode 
            Height          =   315
            Left            =   120
            MaxLength       =   100
            TabIndex        =   15
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total SRP"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   7320
            TabIndex        =   24
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "SRP"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6120
            TabIndex        =   22
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Qty"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5160
            TabIndex        =   20
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1200
            TabIndex        =   18
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Code"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   1095
         End
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   5445
      Width           =   11355
      _ExtentX        =   20029
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
      Left            =   10560
      Top             =   1200
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
            Picture         =   "frmOperationProShop.frx":39C7
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShop.frx":46A1
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShop.frx":537B
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShop.frx":6055
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShop.frx":6D2F
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShop.frx":7A09
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShop.frx":86E3
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShop.frx":93BD
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShop.frx":A097
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShop.frx":A971
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShop.frx":B64B
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShop.frx":C325
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShop.frx":CFFF
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShop.frx":DCD9
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShop.frx":E9B3
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   1320
      ScaleHeight     =   3975
      ScaleWidth      =   8295
      TabIndex        =   1
      Top             =   1200
      Width           =   8295
      Begin VB.TextBox txtPassportKey 
         Height          =   315
         Left            =   840
         MaxLength       =   100
         TabIndex        =   46
         Top             =   360
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.PictureBox picSetFocus 
         BackColor       =   &H00C6B8A4&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3120
         ScaleHeight     =   375
         ScaleWidth      =   735
         TabIndex        =   45
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdSelectPassport 
         Caption         =   "..."
         Height          =   315
         Left            =   2550
         MouseIcon       =   "frmOperationProShop.frx":F68D
         MousePointer    =   99  'Custom
         TabIndex        =   44
         Top             =   360
         Width           =   300
      End
      Begin MSComctlLib.ListView lstDetails 
         Height          =   2535
         Left            =   0
         TabIndex        =   8
         Top             =   840
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   4471
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
         NumItems        =   8
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
            Text            =   "ItemKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Code"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Description"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Qty"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "SRP"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Total SRP"
            Object.Width           =   2293
         EndProperty
      End
      Begin VB.TextBox txtPassport 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtCtrlNo 
         Height          =   315
         Left            =   6600
         MaxLength       =   100
         TabIndex        =   4
         Top             =   0
         Width           =   1695
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   2
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   6600
         TabIndex        =   12
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label lblQty 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   6600
         TabIndex        =   11
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Total SRP"
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
         Left            =   5760
         TabIndex        =   10
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Qty"
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
         Left            =   5760
         TabIndex        =   9
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Passport No"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   390
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl #"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5760
         TabIndex        =   5
         Top             =   30
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   30
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmOperationProShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim TRANS_DET As Long
Const is_DET_REFRESH = 0
Const is_DET_ADDING = 1
Const is_DET_EDITTING = 2

Dim iRow As Long
Dim iFocus As Long
Dim iItemCodeFocus As Long

Dim tmp As Long

Dim x, Arr, iPK, sCtrlNo, i, l

Private Function ProShopItemSRP(iKey) As Double
ProShopItemSRP = 0
t = "SELECT tbl_Operation_ProShop_Items.* " & _
    " FROM tbl_Operation_ProShop_Items " & _
    " WHERE (PK = " & iKey & ")"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then
    ProShopItemSRP = rt!SRP
End If
rt.Close
End Function

Private Sub BROWSER(sCtrl, isAction As String)
'MsgBox sCtrl
Select Case isAction
    Case "is_LOAD"
        If sCtrl <> "" Then
            s = "SELECT TOP 1 tbl_Operation_ProShop.* " & _
                " FROM tbl_Operation_ProShop " & _
                " WHERE (CtrlNo = '" & sCtrl & "') " & _
                " ORDER BY CtrlNo"
        Else
            s = "SELECT TOP 1 tbl_Operation_ProShop.* " & _
                " FROM tbl_Operation_ProShop " & _
                " ORDER BY CtrlNo"
        End If
    Case "is_HOME"
        If picAddProShop.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Operation_ProShop.* " & _
            " FROM tbl_Operation_ProShop " & _
            " WHERE (CtrlNo = '" & sCtrl & "') " & _
            " ORDER BY CtrlNo"
    Case "is_PAGEUP"
        If picAddProShop.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Operation_ProShop.* " & _
            " FROM tbl_Operation_ProShop " & _
            " WHERE (CtrlNo < '" & sCtrl & "') " & _
            " ORDER BY CtrlNo DESC"
    Case "is_PAGEDOWN"
        If picAddProShop.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Operation_ProShop.* " & _
            " FROM tbl_Operation_ProShop " & _
            " WHERE (CtrlNo > '" & sCtrl & "') " & _
            " ORDER BY CtrlNo"
    Case "is_END"
        If picAddProShop.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Operation_ProShop.* " & _
            " FROM tbl_Operation_ProShop " & _
            " ORDER BY CtrlNo DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
'MsgBox rs.RecordCount
If rs.RecordCount > 0 Then
    txtCtrlNo.Text = rs!CtrlNo
    txtDate.Text = Format(rs!TransDate, "mm/dd/yyyy")
    txtPassport.Text = rs!PassportNo
    txtPassportKey.Text = rs!PassportKey
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    CLEAR_DETAILS
    l = 0
    t = "SELECT dbo.tbl_Operation_ProShop_Details.ItemKey, dbo.tbl_Operation_ProShop_Items.Code, " & _
        " dbo.tbl_Operation_ProShop_Items.ItemDescription, dbo.tbl_Operation_ProShop_Details.Qty, " & _
        " dbo.tbl_Operation_ProShop_Details.SRP, dbo.tbl_Operation_ProShop_Details.TotalSRP " & _
        " FROM dbo.tbl_Operation_ProShop_Details LEFT OUTER JOIN " & _
        " dbo.tbl_Operation_ProShop_Items ON dbo.tbl_Operation_ProShop_Details.ItemKey = dbo.tbl_Operation_ProShop_Items.PK " & _
        " Where (dbo.tbl_Operation_ProShop_Details.MasterKey = " & rs!PK & ") " & _
        " ORDER BY dbo.tbl_Operation_ProShop_Details.Line"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstDetails.ListItems.Clear
        While Not rt.EOF
            l = l + 1
            Set x = lstDetails.ListItems.Add()
            x.SubItems(1) = Format(l, "0#")
            x.SubItems(2) = rt!ItemKey
            x.SubItems(3) = rt!Code
            x.SubItems(4) = rt!ItemDescription
            x.SubItems(5) = Format(rt!Qty, "#0.00")
            x.SubItems(6) = Format(rt!SRP, "#,##0.00")
            x.SubItems(7) = Format(rt!TotalSRP, "#,##0.00")
            rt.MoveNext
        Wend
    End If
    rt.Close
    
    SaveSetting App.EXEName, "PassportTrans", "PassportTrans", rs!CtrlNo
    
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSLine.Visible = True Then Exit Sub
    If picAddProShop.Visible = True Then Exit Sub
    If picSearchItem.Visible = True Then Exit Sub
    If AccessRights("Pro Shop", "Add") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    picAddProShop.ZOrder 0
    txtSearchProShop.Text = ""
    picAddProShop.Visible = True
    txtSearchProShop.SetFocus
Else
    If iFocus = 0 Then Exit Sub
    If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
    With lstDetails.ListItems
        If CDbl(.Item(iRow).SubItems(2)) <> 0 Then
            Set x = .Add()
            x.Text = ""
            x.SubItems(1) = Format(.Count, "0#")
            x.SubItems(2) = "0"
            x.SubItems(3) = " "
            x.SubItems(4) = " "
            x.SubItems(5) = " "
            x.SubItems(6) = " "
            x.SubItems(7) = " "
        Else
            .Item(1).SubItems(1) = Format(.Count, "0#")
            .Item(1).SubItems(2) = "0"
            .Item(1).SubItems(3) = " "
            .Item(1).SubItems(4) = " "
            .Item(1).SubItems(5) = " "
            .Item(1).SubItems(6) = " "
            .Item(1).SubItems(7) = " "
        End If
        iRow = .Count
        lstDetails.ListItems(iRow).EnsureVisible
        lstDetails.ListItems(iRow).Selected = True
        txtCode.Text = ""
        txtDescription.Text = ""
        txtQty.Text = ""
        txtSRP.Text = ""
        txtTotalSRP.Text = ""
        picMain.Enabled = False
        picToolbar.Enabled = False
        picSLine.ZOrder 0
        picSLine.Visible = True
        TRANS_DET = is_DET_ADDING
        txtCode.SetFocus
    End With
End If
End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSLine.Visible = True Then Exit Sub
    If picAddProShop.Visible = True Then Exit Sub
    If picSearchItem.Visible = True Then Exit Sub
    If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
    If AccessRights("Pro Shop", "Edit") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_EDITTING
    txtDate.SetFocus
Else
    If iFocus = 0 Then Exit Sub
    If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
    With lstDetails.ListItems
        txtItemKey.Text = .Item(iRow).SubItems(2)
        txtCode.Text = .Item(iRow).SubItems(3)
        txtDescription.Text = .Item(iRow).SubItems(4)
        txtQty.Text = .Item(iRow).SubItems(5)
        txtSRP.Text = .Item(iRow).SubItems(6)
        txtTotalSRP.Text = .Item(iRow).SubItems(7)
        
        txtItemKey1.Text = .Item(iRow).SubItems(2)
        txtCode1.Text = .Item(iRow).SubItems(3)
        txtDescription1.Text = .Item(iRow).SubItems(4)
        txtQty1.Text = .Item(iRow).SubItems(5)
        txtSRP1.Text = .Item(iRow).SubItems(6)
        txtTotalSRP1.Text = .Item(iRow).SubItems(7)
        
        picMain.Enabled = False
        picToolbar.Enabled = False
        picSLine.ZOrder 0
        picSLine.Visible = True
        TRANS_DET = is_DET_EDITTING
        txtCode.SetFocus
    End With
End If
End Sub

Private Sub PRESS_DELETE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSLine.Visible = True Then Exit Sub
    If picAddProShop.Visible = True Then Exit Sub
    If picSearchItem.Visible = True Then Exit Sub
    If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
    If AccessRights("Pro Shop", "Delete") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If MsgBox("ARE YOU SURE IN DELETING THIS TRANSACTION?                       ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    On Error GoTo PG:
    ConnOmega.Execute "DELETE FROM tbl_Operation_ProShop WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    CLEARTEXT
    BROWSER GetSetting(App.EXEName, "PassportTrans", "PassportTrans", ""), "is_PAGEDOWN"
If Trim(txtCtrlNo.Text) = "" Then BROWSER GetSetting(App.EXEName, "PassportTrans", "PassportTrans", ""), "is_HOME"
Else
    If iFocus = 0 Then Exit Sub
    If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
    With lstDetails.ListItems
        If .Count > 1 Then
            .Remove iRow
            If CDbl(iRow) > CDbl(.Count) Then
                iRow = .Count
            End If
        Else
            .Item(1).SubItems(1) = " "
            .Item(1).SubItems(2) = "0"
            .Item(1).SubItems(3) = " "
            .Item(1).SubItems(4) = " "
            .Item(1).SubItems(5) = " "
            .Item(1).SubItems(6) = " "
            .Item(1).SubItems(7) = " "
            iRow = 1
        End If
        lstDetails.ListItems(iRow).EnsureVisible
        lstDetails.ListItems(iRow).Selected = True
    End With
End If
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If IsDate(txtDate.Text) = False Then MsgBox "Please supply a valid Date!                  ", vbCritical, "Error...": txtDate.SetFocus: Exit Sub
'If Trim(txtCtrlNo.Text) = "" Then MsgBox "Please supply Control Number!                       ", vbCritical, "Error...": txtCtrlNo.SetFocus: Exit Sub
If Trim(txtPassport.Text) = "" Then MsgBox "Please select Passport Number!                        ", vbCritical, "Error...": txtPassport.SetFocus: Exit Sub
' Validate Passport
s = "SELECT tbl_Operation_Passport.* " & _
    " FROM tbl_Operation_Passport " & _
    " WHERE (PassportNo = '" & Trim(txtPassport.Text) & "') " '& _
    " AND (DateAll = '" & FormatDateTime(txtDate.Text, vbShortDate) & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount = 0 Then
    MsgBox "Invalid Passport!               ", vbCritical, "Error..."
    rs.Close
    Exit Sub
Else
    If DateValue(FormatDateTime(rs!DateAll, vbShortDate)) <> DateValue(FormatDateTime(txtDate.Text, vbShortDate)) Then
        MsgBox "Date not match with the passport!                   ", vbCritical, "Error..."
        rs.Close
        Exit Sub
    End If
End If
rs.Close

l = 0
For i = 1 To lstDetails.ListItems.Count
    If CDbl(lstDetails.ListItems.Item(i).SubItems(2)) > 0 Then
        l = l + 1
    End If
Next i

If CDbl(l) = 0 Then MsgBox "No Details!                     ", vbCritical, "Error...": Exit Sub

On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    sCtrlNo = ""
    s = "SELECT TOP 1 tbl_Operation_ProShop.* " & _
        " FROM tbl_Operation_ProShop " & _
        " WHERE (Year(TransDate) = " & Format(FormatDateTime(txtDate.Text, vbShortDate), "yyyy") & ") " & _
        " ORDER BY CtrlNo DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        sCtrlNo = Format(CDbl(rs!CtrlNo) + 1, "0000000#")
    Else
        sCtrlNo = Format(FormatDateTime(txtDate.Text, vbShortDate), "yyyy") & "0000"
    End If
    rs.Close
    Do
        s = "SELECT tbl_Operation_ProShop.* " & _
            " FROM tbl_Operation_ProShop " & _
            " WHERE (CtrlNo = '" & sCtrlNo & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If
        rs.Close
        sCtrlNo = Format(CDbl(sCtrlNo) + 1, "0000000#")
    Loop
    
    ConnOmega.Execute "INSERT INTO tbl_Operation_ProShop " & _
                      " (CtrlNo, TransDate, PassportNo, PassportKey, LastModified) " & _
                      " VALUES ('" & sCtrlNo & "', '" & FormatDateTime(txtDate.Text, vbShortDate) & "', " & _
                      " '" & Trim(txtPassport.Text) & "', " & txtPassportKey.Text & ", '" & CStr(Now) & " - " & gbl_CompleteName & "')"
    iPK = 0
    s = "SELECT PK " & _
        " FROM tbl_Operation_ProShop " & _
        " WHERE (CtrlNo = '" & sCtrlNo & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        iPK = rs!PK
    End If
    rs.Close
End If
If TRANSACTIONTYPE = is_EDITTING Then
    sCtrlNo = Trim(txtCtrlNo.Text)
    iPK = Statusbar1.Panels(1).Text
    
    ConnOmega.Execute "UPDATE tbl_Operation_ProShop " & _
                      " SET TransDate = '" & FormatDateTime(txtDate.Text, vbShortDate) & "', " & _
                      " PassportNo = '" & Trim(txtPassport.Text) & "', " & _
                      " PassportKey = " & txtPassportKey.Text & ", " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & iPK & ")"
    
    
    
End If
If CDbl(iPK) > 0 Then
    l = 0
    ConnOmega.Execute "DELETE FROM tbl_Operation_ProShop_Details WHERE (MasterKey = " & iPK & ")"
    With lstDetails.ListItems
        For i = 1 To .Count
            If CDbl(.Item(i).SubItems(2)) > 0 Then
                l = l + 1
                ConnOmega.Execute "INSERT INTO tbl_Operation_ProShop_Details " & _
                                  " (MasterKey, Line, ItemKey, Qty, SRP) " & _
                                  " VALUES (" & iPK & ", " & l & ", " & .Item(i).SubItems(2) & ", " & _
                                  " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(5)) = False, 0, .Item(i).SubItems(5))) & ", " & _
                                  " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(6)) = False, 0, .Item(i).SubItems(6))) & ")"
            End If
        Next i
    End With
End If
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER sCtrlNo, "is_LOAD"
txtDate.SetFocus
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If picAddProShop.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
End Sub

Private Sub PRESS_F8()
If picAddProShop.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
If AccessRights("Pro Shop", "Post") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If

End Sub

Private Sub PRESS_F9()
If picAddProShop.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picAddProShop.Visible = True Then cmdCancelProShop_Click
    Unload Me
Else
    If picSearchItem.Visible = True Then
        cmdCancelSearchItem_Click
        Exit Sub
    End If
    If TRANS_DET = is_ADDING Then
        With lstDetails.ListItems
            If .Count > 1 Then
                .Remove iRow
                If CDbl(iRow) > CDbl(.Count) Then
                    iRow = .Count
                End If
            Else
                .Item(1).SubItems(1) = " "
                .Item(1).SubItems(2) = "0"
                .Item(1).SubItems(3) = " "
                .Item(1).SubItems(4) = " "
                .Item(1).SubItems(5) = " "
                .Item(1).SubItems(6) = " "
                .Item(1).SubItems(7) = " "
                iRow = 1
            End If
        End With
        lstDetails.ListItems(iRow).EnsureVisible
        lstDetails.ListItems(iRow).Selected = True
        picSLine.Visible = False
        picMain.Enabled = True
        picToolbar.Enabled = True
        lstDetails.SetFocus
        Exit Sub
    End If
    If TRANS_DET = is_DET_EDITTING Then
        With lstDetails.ListItems
            .Item(iRow).SubItems(2) = txtItemKey1.Text
            .Item(iRow).SubItems(3) = txtCode1.Text
            .Item(iRow).SubItems(4) = txtDescription1.Text
            .Item(iRow).SubItems(5) = txtQty1.Text
            .Item(iRow).SubItems(6) = txtSRP1.Text
            .Item(iRow).SubItems(7) = txtTotalSRP1.Text
            picSLine.Visible = False
            picMain.Enabled = True
            picToolbar.Enabled = True
            lstDetails.SetFocus
        End With
        Exit Sub
    End If
    If iFocus = 1 Then txtDate.SetFocus
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "PassportTrans", "PassportTrans", ""), "is_LOAD"
    If Trim(txtCtrlNo.Text) = "" Then BROWSER GetSetting(App.EXEName, "PassportTrans", "PassportTrans", ""), "is_HOME"
End If
End Sub

Private Sub CLEARTEXT()
txtDate.Text = ""
txtPassport.Text = ""
txtCtrlNo.Text = ""
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
CLEAR_DETAILS
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtDate.Locked = bln
txtPassport.Locked = bln
txtCtrlNo.Locked = True
txtDescription.Locked = True
txtTotalSRP.Locked = True
End Sub


Private Sub CLEAR_DETAILS()
With lstDetails.ListItems
    .Clear
    Set x = .Add()
    x.Text = ""
    x.SubItems(1) = " "
    x.SubItems(2) = "0"
    x.SubItems(3) = " "
    x.SubItems(4) = " "
    x.SubItems(5) = " "
    x.SubItems(6) = " "
    x.SubItems(7) = " "
End With
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
cmdCancelSearchItem_Click
End Sub

Private Sub b8TitleBar5_CLoseClick()
cmdCancelProShop_Click
End Sub

Private Sub cmdCancelProShop_Click()
picAddProShop.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True
txtDate.SetFocus
End Sub

Private Sub cmdCancelSearchItem_Click()
picSearchItem.Visible = False
picSLine.Enabled = True
txtCode.SetFocus
End Sub

Private Sub cmdOKProShop_Click()
If lstResultProShop.ListIndex = -1 Then Exit Sub
CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
Arr = Split(lstResultProShop.List(lstResultProShop.ListIndex), " - ", -1, 1)
txtPassportKey.Text = lstResultProShop.ItemData(lstResultProShop.ListIndex)
t = "SELECT tbl_Operation_Passport.* " & _
    " FROM tbl_Operation_Passport " & _
    " WHERE (PK = " & RETURNTEXTVALUE(txtPassportKey) & ")"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then
    txtDate.Text = Format(rt!DateAll, "mm/dd/yyyy")
Else
    txtDate.Text = Format(Date, "mm/dd/yyyy")
End If
rt.Close
txtPassport.Text = Arr(0)
cmdCancelProShop_Click
End Sub

Private Sub cmdOKSearchItem_Click()
If lstResultSearchItem.ListIndex = -1 Then Exit Sub
Arr = Split(lstResultSearchItem.List(lstResultSearchItem.ListIndex), " - ", -1, 1)
txtItemKey.Text = lstResultSearchItem.ItemData(lstResultSearchItem.ListIndex)
txtCode.Text = Arr(0)
txtDescription.Text = Arr(1)
txtSRP.Text = Format(ProShopItemSRP(lstResultSearchItem.ItemData(lstResultSearchItem.ListIndex)), "#,##0.00")
cmdCancelSearchItem_Click
End Sub

Private Sub cmdSelectPassport_Click()
picSetFocus.SetFocus
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert:   PRESS_INSERT
    Case vbKeyF2:       PRESS_F2
    Case vbKeyDelete:   PRESS_DELETE
    Case vbKeyF5:       PRESS_F5
    Case vbKeyF6:       PRESS_F6
    Case vbKeyF8:       PRESS_F8
    Case vbKeyF9:       PRESS_F9
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "PassportTrans", "PassportTrans", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "PassportTrans", "PassportTrans", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "PassportTrans", "PassportTrans", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "PassportTrans", "PassportTrans", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.Height - Me.Height) / 3
Me.Left = (MainForm.Width - Me.Width) / 5
picSLine.Width = 8535
iRow = 0
iFocus = 0
iItemCodeFocus = 0
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "PassportTrans", "PassportTrans", ""), "is_LOAD"
If Trim(txtCtrlNo.Text) = "" Then BROWSER GetSetting(App.EXEName, "PassportTrans", "PassportTrans", ""), "is_HOME"
tmp = SetWindowLong(txtSearchItem.hwnd, GWL_STYLE, GetWindowLong(txtSearchItem.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picAddProShop.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstDetails_GotFocus()
iFocus = 1
TRANS_DET = is_DET_REFRESH
iRow = lstDetails.SelectedItem.Index
With lstDetails.ListItems
    If .Count = 1 Then
        If CDbl(.Item(iRow).SubItems(2)) > 0 Then
            TOOLBARFUNC 5
        Else
            TOOLBARFUNC 4
        End If
    ElseIf .Count > 1 Then
        TOOLBARFUNC 5
    End If
End With
End Sub

Private Sub lstDetails_ItemClick(ByVal Item As MSComctlLib.ListItem)
iRow = lstDetails.SelectedItem.Index
End Sub

Private Sub lstDetails_LostFocus()
iFocus = 0
End Sub



Private Sub lstResultProShop_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKProShop_Click
End Sub

Private Sub lstResultSearchItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKSearchItem_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Refresh"
            'ToDo: Add 'Refresh' button code.
            'MsgBox "Add 'Refresh' button code."
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "PassportTrans", "PassportTrans", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "PassportTrans", "PassportTrans", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "PassportTrans", "PassportTrans", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "PassportTrans", "PassportTrans", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":   PRESS_F9
    Case "Post":    PRESS_F8
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtCode_Change()
If TRANS_DET = is_DET_ADDING Or _
TRANS_DET = is_DET_EDITTING Then
    With lstDetails.ListItems
        .Item(iRow).SubItems(3) = txtCode.Text
    End With
End If
End Sub

Private Sub txtCode_GotFocus()
iItemCodeFocus = 1
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF6 Then
    picSearchItem.ZOrder 0
    picSLine.Enabled = False
    txtSearchItem.Text = ""
    picSearchItem.Visible = True
    txtSearchItem.SetFocus
End If
If KeyCode = vbKeyReturn Then txtQty.SetFocus
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtCode_LostFocus()
iItemCodeFocus = 0
End Sub

Private Sub txtDescription_Change()
If TRANS_DET = is_DET_ADDING Or _
TRANS_DET = is_DET_EDITTING Then
    With lstDetails.ListItems
        .Item(iRow).SubItems(4) = txtDescription.Text
    End With
End If
End Sub

Private Sub txtItemKey_Change()
If TRANS_DET = is_DET_ADDING Or _
TRANS_DET = is_DET_EDITTING Then
    With lstDetails.ListItems
        .Item(iRow).SubItems(2) = txtItemKey.Text
    End With
End If
End Sub

Private Sub txtQty_Change()
txtTotalSRP.Text = Format(RETURNTEXTVALUE(txtQty) * RETURNTEXTVALUE(txtSRP), "#,##0.00")
If TRANS_DET = is_DET_ADDING Or _
TRANS_DET = is_DET_EDITTING Then
    With lstDetails.ListItems
        .Item(iRow).SubItems(5) = Format(RETURNTEXTVALUE(txtQty), "#0.00")
    End With
End If
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtSRP.SetFocus
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtSearchItem_Change()
If Trim(txtSearchItem.Text) = "" Then lstResultSearchItem.Clear
lstResultSearchItem.Clear
s = "SELECT PK, Code, ItemDescription " & _
    " From dbo.tbl_Operation_ProShop_Items " & _
    " WHERE (ItemDescription LIKE '" & FORMATSQL(Trim(txtSearchItem.Text)) & "%') " & _
    " ORDER BY ItemDescription, Code"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultSearchItem.AddItem rs!Code & " - " & rs!ItemDescription
    lstResultSearchItem.ItemData(lstResultSearchItem.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResultSearchItem.ListCount Then lstResultSearchItem.ListIndex = 0
End Sub

Private Sub txtSearchItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultSearchItem.SetFocus
End Sub

Private Sub txtSearchProShop_Change()
If Trim(txtSearchProShop.Text) = "" Then lstResultProShop.Clear: Exit Sub
lstResultProShop.Clear
s = "SELECT PK, PassportNo, PlayerName " & _
    " FROM tbl_Operation_Passport " & _
    " WHERE (PassportNo LIKE '" & FORMATSQL(Trim(txtSearchProShop.Text)) & "%') " & _
    " AND (RegistrationAdded = 1) " & _
    " ORDER BY PassportNo"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultProShop.AddItem rs!PassportNo & " - " & rs!PlayerName
    lstResultProShop.ItemData(lstResultProShop.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResultProShop.ListCount Then lstResultProShop.ListIndex = 0
End Sub

Private Sub txtSearchProShop_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultProShop.SetFocus
End Sub

Private Sub txtSearchProShop_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtSRP_Change()
txtTotalSRP.Text = Format(RETURNTEXTVALUE(txtQty) * RETURNTEXTVALUE(txtSRP), "#,##0.00")
If TRANS_DET = is_DET_ADDING Or _
TRANS_DET = is_DET_EDITTING Then
    With lstDetails.ListItems
        .Item(iRow).SubItems(6) = Format(RETURNTEXTVALUE(txtSRP), "#,##0.00")
    End With
End If
End Sub

Private Sub txtSRP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If TRANS_DET = is_DET_ADDING Then
        TRANS_DET = is_DET_REFRESH
        With lstDetails.ListItems
            If CDbl(.Item(iRow).SubItems(2)) <> 0 Then
                Set x = .Add()
                x.Text = ""
                x.SubItems(1) = Format(.Count, "0#")
                x.SubItems(2) = "0"
                x.SubItems(3) = " "
                x.SubItems(4) = " "
                x.SubItems(5) = " "
                x.SubItems(6) = " "
                x.SubItems(7) = " "
                iRow = .Count
            End If
            txtCode.Text = ""
            txtDescription.Text = ""
            txtQty.Text = ""
            txtSRP.Text = ""
            txtTotalSRP.Text = ""
            TRANS_DET = is_DET_ADDING
            txtCode.SetFocus
        End With
    End If
    If TRANS_DET = is_DET_EDITTING Then
        picSLine.Visible = False
        TRANS_DET = is_DET_REFRESH
        picMain.Enabled = True
        lstDetails.SetFocus
    End If
End If
End Sub

Private Sub txtSRP_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtTotalSRP_Change()
If TRANS_DET = is_DET_ADDING Or _
TRANS_DET = is_DET_EDITTING Then
    With lstDetails.ListItems
        .Item(iRow).SubItems(7) = Format(RETURNTEXTVALUE(txtTotalSRP), "#,##0.00")
    End With
End If
End Sub
