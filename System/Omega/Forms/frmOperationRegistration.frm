VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOperationRegistration 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOperationRegistration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   11445
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picAddBagDrop 
      Height          =   4575
      Left            =   1680
      TabIndex        =   29
      Top             =   240
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
         TabIndex        =   47
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
      Begin VB.CommandButton cmdOKBagTag 
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
         Picture         =   "frmOperationRegistration.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3945
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelBagTag 
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
         Picture         =   "frmOperationRegistration.frx":0F3C
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3945
         Width           =   1560
      End
      Begin VB.TextBox txtSearchBagTag 
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   4215
      End
      Begin VB.ListBox lstResulBagTag 
         Height          =   2985
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   4215
      End
      Begin RPVGCC.b8TitleBar b8TitleBar4 
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
         Icon            =   "frmOperationRegistration.frx":1698
         ShadowVisible   =   0   'False
      End
      Begin RPVGCC.b8TitleBar b8TitleBar5 
         Height          =   345
         Left            =   40
         TabIndex        =   35
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
         Icon            =   "frmOperationRegistration.frx":1C32
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
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15600
      TabIndex        =   41
      Top             =   0
      Width           =   15600
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   42
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
         MouseIcon       =   "frmOperationRegistration.frx":21CC
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   9900
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   43
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmOperationRegistration.frx":24E6
               Top             =   120
               Visible         =   0   'False
               Width           =   1395
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
   Begin RPVGCC.b8Container picSLine 
      Height          =   855
      Left            =   3240
      TabIndex        =   22
      Top             =   2040
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox txtQty1 
         Height          =   315
         Left            =   2760
         TabIndex        =   40
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtUOM1 
         Height          =   315
         Left            =   2520
         TabIndex        =   39
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtEquipment1 
         Height          =   315
         Left            =   2280
         TabIndex        =   38
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtEQKey1 
         Height          =   315
         Left            =   2040
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtEQKey 
         Height          =   315
         Left            =   1560
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.ComboBox cmbEquipment 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtUOM 
         Height          =   315
         Left            =   2040
         TabIndex        =   26
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3240
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "UOM"
         Height          =   255
         Left            =   2040
         TabIndex        =   27
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   255
         Left            =   3000
         TabIndex        =   25
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "Equipment"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   1815
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   4770
      Width           =   11445
      _ExtentX        =   20188
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
            Picture         =   "frmOperationRegistration.frx":2BF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationRegistration.frx":38D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationRegistration.frx":45AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationRegistration.frx":5287
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationRegistration.frx":5F61
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationRegistration.frx":6C3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationRegistration.frx":7915
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationRegistration.frx":85EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationRegistration.frx":92C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationRegistration.frx":9BA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationRegistration.frx":A87D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationRegistration.frx":B557
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationRegistration.frx":C231
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationRegistration.frx":CF0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationRegistration.frx":DBE5
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
      Begin VB.TextBox txtBagTagNo 
         Height          =   315
         Left            =   3480
         MaxLength       =   100
         TabIndex        =   45
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   6960
         ScaleHeight     =   3225
         ScaleWidth      =   3225
         TabIndex        =   44
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
      Begin VB.TextBox txtCheckInBy 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   20
         Top             =   3000
         Width           =   5775
      End
      Begin MSComctlLib.ListView lstEquipmentRental 
         Height          =   1215
         Left            =   2640
         TabIndex        =   19
         Top             =   1680
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2143
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
         NumItems        =   6
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
            Text            =   "EquipmentKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Equipment"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "UOM"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Qty"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.TextBox txtLockerKeyNo 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   16
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtFlightNo 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   14
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtTeeTime 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   12
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtPassportNo 
         Height          =   315
         Left            =   5640
         MaxLength       =   100
         TabIndex        =   10
         Top             =   0
         Width           =   1215
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   5
         Top             =   0
         Width           =   1455
      End
      Begin VB.TextBox txtPlayerName 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   4
         Top             =   360
         Width           =   5775
      End
      Begin VB.ComboBox cmbPlayerType 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   5775
      End
      Begin VB.TextBox txtMemberName 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   2
         Top             =   1080
         Width           =   5775
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "BagTag #"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2640
         TabIndex        =   46
         Top             =   30
         Width           =   855
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
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Check In By"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   21
         Top             =   3030
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Equipment Rental"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   1440
         Width           =   4095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Locker Key #"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Flight No"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tee Time"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Passport #"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4800
         TabIndex        =   11
         Top             =   30
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   30
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Player Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   390
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Player Type"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Guest Of"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   1110
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmOperationRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iAddType As Long

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim TRANS_DETAIL  As Long
Const is_DET_REFRESH = 0
Const is_DET_ADDING = 1
Const is_DET_EDITTING = 2

Dim tmp As Long

Dim iFocus As Long
Dim iRow As Long

Dim iPK As Long
Dim iMemberKey As Long
Dim iPlayerType As Long

Dim i, l, x, sCtrl

Private Sub BROWSER(sPassNo, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If sPassNo <> "" Then
            s = "SELECT TOP 1 tbl_Operation_Passport.* " & _
                " FROM tbl_Operation_Passport " & _
                " WHERE (PassportNo = '" & sPassNo & "') " & _
                " AND (RegistrationAdded = 1) " & _
                " ORDER BY PassportNo"
        Else
            s = "SELECT TOP 1 tbl_Operation_Passport.* " & _
                " FROM tbl_Operation_Passport " & _
                " WHERE (RegistrationAdded = 1) " & _
                " ORDER BY PassportNo"
        End If
    Case "is_HOME"
        If picAddBagDrop.Visible = True Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Operation_Passport.* " & _
            " FROM tbl_Operation_Passport " & _
            " WHERE (RegistrationAdded = 1) " & _
            " ORDER BY PassportNo"
    Case "is_PAGEUP"
        If picAddBagDrop.Visible = True Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Operation_Passport.* " & _
            " FROM tbl_Operation_Passport " & _
            " WHERE (PassportNo < '" & sPassNo & "') " & _
            " AND (RegistrationAdded = 1) " & _
            " ORDER BY PassportNo DESC"
    Case "is_PAGEDOWN"
        If picAddBagDrop.Visible = True Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Operation_Passport.* " & _
            " FROM tbl_Operation_Passport " & _
            " WHERE (PassportNo > '" & sPassNo & "') " & _
            " AND (RegistrationAdded = 1) " & _
            " ORDER BY PassportNo "
    Case "is_END"
        If picAddBagDrop.Visible = True Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Operation_Passport.* " & _
            " FROM tbl_Operation_Passport " & _
            " WHERE (RegistrationAdded = 1) " & _
            " ORDER BY PassportNo DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iMemberKey = IIf(IsNull(rs!MemberKey), 0, rs!MemberKey)
    iPlayerType = rs!PlayerTypeKey
    txtDate.Text = Format(rs!DateAll, "mm/dd/yyyy")
    txtPassportNo.Text = IIf(IsNull(rs!PassportNo), "", rs!PassportNo)
    txtBagTagNo.Text = rs!BagTagNo
    txtPlayerName.Text = IIf(IsNull(rs!PlayerName), "", rs!PlayerName)
    txtMemberName.Text = ""
    If iPlayerType <> 1 Then
        If IsNull(rs!MemberKey) = False Then
            t = "SELECT MemberName " & _
                " From dbo.tbl_Member_IDNumber " & _
                " WHERE (PK = " & IIf(IsNull(rs!MemberKey), 0, rs!MemberKey) & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                txtMemberName.Text = rt!MemberName
                imgPicture.Picture = LoadPicture(SHOW_IMAGES(IIf(IsNull(rs!MemberKey), 0, rs!MemberKey), 0, "Member ID Number"))
            End If
            rt.Close
        End If
    Else
        imgPicture.Picture = LoadPicture("")
    End If
    
    txtTeeTime.Text = ""
    If IsNull(rs!TeeTime) = False Then
        txtTeeTime.Text = Format(rs!TeeTime, "hh:mm AM/PM")
    End If
    txtFlightNo.Text = IIf(IsNull(rs!FlightNo), "", rs!FlightNo)
    txtLockerKeyNo.Text = rs!LockerKeyNo
    txtCheckInBy.Text = IIf(IIf(IsNull(rs!RegCheckInBy), "", rs!RegCheckInBy) <> "", rs!RegCheckInBy & " on " & Format(rs!RegCheckIn, "mm/dd/yyyy hh:mm:ss AM/PM"), "")
    cmbPlayerType.ListIndex = rs!PlayerTypeKey - 1
    imgPosted.Visible = IIf(rs!PostedRegistration = 1, True, False)
    
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModifiedRegistratrion), "", rs!LastModifiedRegistratrion)
    
    
    CLEARDETAIL
    t = "SELECT dbo.tbl_Operation_Passport_EquipmentRental.Line, " & _
        " dbo.tbl_Operation_Passport_EquipmentRental.EquipmentKey, " & _
        " dbo.tbl_Operation_Equipment_For_Rent.EquipmentName, " & _
        " dbo.tbl_Operation_Equipment_For_Rent.UnitOfMeasure, " & _
        " dbo.tbl_Operation_Passport_EquipmentRental.Qty " & _
        " FROM dbo.tbl_Operation_Passport_EquipmentRental LEFT OUTER JOIN " & _
        " dbo.tbl_Operation_Equipment_For_Rent ON dbo.tbl_Operation_Passport_EquipmentRental.EquipmentKey = dbo.tbl_Operation_Equipment_For_Rent.PK " & _
        " Where (dbo.tbl_Operation_Passport_EquipmentRental.MasterKey = 1) " & _
        " ORDER BY dbo.tbl_Operation_Passport_EquipmentRental.Line"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstEquipmentRental.ListItems.Clear
        i = 0
        While Not rt.EOF
            i = i + 1
            Set x = lstEquipmentRental.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = Format(i, "0#") '#
            x.SubItems(2) = rt!EquipmentKey 'EquipmentKey
            x.SubItems(3) = rt!EquipmentName 'Equipment
            x.SubItems(4) = rt!UnitOfMeasure 'UOM
            x.SubItems(5) = rt!Qty 'Qty
            rt.MoveNext
        Wend
    End If
    rt.Close
    
    
    SaveSetting App.EXEName, "OperationRegPassNo", "OperationRegPassNo", IIf(IsNull(rs!PassportNo), "", rs!PassportNo)
    
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If TRANSACTIONTYPE = is_REFRESH Then
    If picAddBagDrop.Visible = True Then Exit Sub
    PopupMenu MainFormPopupF.mnuRegistrationAdd, , Toolbar1.Buttons(1).Left, Toolbar1.Buttons(1).Top + Toolbar1.Buttons(1).Height
'    If AccessRights("Registration", "Add") = False Then
'        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
'               "ACCESS DENIED!                                      ", vbCritical, "Alert"
'        Exit Sub
'    End If
'    picAddBagDrop.ZOrder 0
'    txtSearchBagTag.Text = ""
'    picAddBagDrop.Visible = True
'    txtSearchBagTag.SetFocus
Else
    If iFocus = 0 Then Exit Sub
    If picSLine.Visible = True Then Exit Sub
    With lstEquipmentRental.ListItems
        If CDbl(.Item(iRow).SubItems(2)) = 0 Then
            .Item(iRow).SubItems(1) = Format(iRow, "0#")
            .Item(iRow).SubItems(2) = "0"
            .Item(iRow).SubItems(3) = " "
            .Item(iRow).SubItems(4) = " "
            .Item(iRow).SubItems(5) = " "
        Else
            Set x = .Add()
            x.Text = ""
            x.SubItems(1) = Format(.Count, "0#")
            x.SubItems(2) = "0"
            x.SubItems(3) = " "
            x.SubItems(4) = " "
            x.SubItems(5) = " "
            iRow = .Count
        End If
    End With
    lstEquipmentRental.ListItems(iRow).EnsureVisible
    lstEquipmentRental.ListItems(iRow).Selected = True
    cmbEquipment.ListIndex = -1
    txtEQKey.Text = ""
    txtUOM.Text = ""
    txtQty.Text = ""
    TRANS_DETAIL = is_DET_ADDING
    TOOLBARFUNC 2
    picSLine.ZOrder 0
    picMain.Enabled = False
    picToolbar.Enabled = False
    picSLine.Visible = True
    cmbEquipment.SetFocus
End If
End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE = is_REFRESH Then
    If picAddBagDrop.Visible = True Then Exit Sub
    If imgPosted.Visible = True Then MsgBox "Already Posted!                ", vbCritical, "Error...": Exit Sub
    If AccessRights("Registration", "Edit") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_EDITTING
Else
    If iFocus = 0 Then Exit Sub
    If picSLine.Visible = True Then Exit Sub
    With lstEquipmentRental.ListItems
        txtEQKey.Text = .Item(iRow).SubItems(2)
        cmbEquipment.ListIndex = .Item(iRow).SubItems(2) - 1
        txtUOM.Text = .Item(iRow).SubItems(4)
        txtQty.Text = .Item(iRow).SubItems(5)
    
        txtEQKey1.Text = .Item(iRow).SubItems(2)
        txtEquipment1.Text = .Item(iRow).SubItems(3)
        txtUOM1.Text = .Item(iRow).SubItems(4)
        txtQty1.Text = .Item(iRow).SubItems(5)
    End With
    lstEquipmentRental.ListItems(iRow).EnsureVisible
    lstEquipmentRental.ListItems(iRow).Selected = True
    TRANS_DETAIL = is_DET_EDITTING
    TOOLBARFUNC 2
    picSLine.ZOrder 0
    picMain.Enabled = False
    picToolbar.Enabled = False
    picSLine.Visible = True
    cmbEquipment.SetFocus
End If
End Sub

Private Sub PRESS_DELETE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picAddBagDrop.Visible = True Then Exit Sub
    If picSLine.Visible = True Then Exit Sub
    If imgPosted.Visible = True Then MsgBox "Already Posted!                ", vbCritical, "Error...": Exit Sub
    If AccessRights("Registration", "Delete") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If MsgBox("ARE YOU SURE IN DELETING THIS TRANSACTION?                   ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    On Error GoTo PG:
    ConnOmega.Execute "UPDATE tbl_Operation_Passport SET RegistrationAdded = 0 WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    CLEARTEXT
    BROWSER GetSetting(App.EXEName, "OperationRegPassNo", "OperationRegPassNo", ""), "is_PAGEDOWN"
    If Trim(txtPassportNo.Text) = "" Then BROWSER GetSetting(App.EXEName, "OperationRegPassNo", "OperationRegPassNo", ""), "is_HOME"
Else
    If iFocus = 0 Then Exit Sub
    If picSLine.Visible = True Then Exit Sub
    With lstEquipmentRental.ListItems
        If .Count = 1 Then
            .Item(iRow).SubItems(1) = " "
            .Item(iRow).SubItems(2) = "0"
            .Item(iRow).SubItems(3) = " "
            .Item(iRow).SubItems(4) = " "
            .Item(iRow).SubItems(5) = " "
        Else
            .Remove iRow
            If CDbl(iRow) > CDbl(.Count) Then
                iRow = .Count
            End If
        End If
        lstEquipmentRental.ListItems(iRow).EnsureVisible
        lstEquipmentRental.ListItems(iRow).Selected = True
    End With
End If
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If picSLine.Visible = True Then Exit Sub
If picAddBagDrop.Visible = True Then Exit Sub
If Trim(txtPassportNo.Text) = "" Then MsgBox "Please Supply Passport Number!                  ", vbCritical, "Error...": txtPassportNo.SetFocus: Exit Sub
s = "SELECT tbl_Operation_Passport.* " & _
    " FROM tbl_Operation_Passport " & _
    " WHERE (PassportNo = '" & Trim(txtPassportNo.Text) & "') " & _
    " AND (PK <> " & IIf(Trim(Statusbar1.Panels(1).Text) = "", 0, Statusbar1.Panels(1).Text) & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    MsgBox "Found Duplicate Passport Number!                            ", vbCritical, "Error...": rs.Close: Exit Sub
End If
rs.Close
If Trim(txtLockerKeyNo.Text) = "" Then
    If MsgBox("Proceed without Locker Key?                          ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm...") = vbNo Then
        txtLockerKeyNo.SetFocus: Exit Sub
    End If
End If
iPK = Statusbar1.Panels(1).Text
If TRANSACTIONTYPE = is_ADDING Then
    ConnOmega.Execute "UPDATE tbl_Operation_Passport " & _
                      " SET RegistrationAdded = 1, " & _
                      " LastModifiedRegistratrion = '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                      " PassportNo = '" & Trim(txtPassportNo.Text) & "', " & _
                      " PlayerTypeKey = " & iPlayerType & ", " & _
                      " PlayerName = '" & FORMATSQL(Trim(txtPlayerName.Text)) & "', " & _
                      " LockerKeyNo = '" & FORMATSQL(Trim(txtLockerKeyNo.Text)) & "', " & _
                      " FlightNo = '" & FORMATSQL(Trim(txtFlightNo.Text)) & "' " & _
                      " WHERE (PK = " & iPK & ")"
End If
If TRANSACTIONTYPE = is_EDITTING Then
    
    ConnOmega.Execute "UPDATE tbl_Operation_Passport " & _
                      " SET LastModifiedRegistratrion = '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                      " PassportNo = '" & Trim(txtPassportNo.Text) & "', " & _
                      " PlayerTypeKey = " & iPlayerType & ", " & _
                      " PlayerName = '" & FORMATSQL(Trim(txtPlayerName.Text)) & "', " & _
                      " LockerKeyNo = '" & FORMATSQL(Trim(txtLockerKeyNo.Text)) & "', " & _
                      " FlightNo = '" & FORMATSQL(Trim(txtFlightNo.Text)) & "' " & _
                      " WHERE (PK = " & iPK & ")"
End If

If IsDate(txtTeeTime.Text) = True Then
    ConnOmega.Execute "UPDATE tbl_Operation_Passport  " & _
                      " SET TeeTime = '" & txtTeeTime.Text & "' " & _
                      " WHERE (PK = " & iPK & ")"
End If

ConnOmega.Execute "DELETE FROM tbl_Operation_Passport_EquipmentRental WHERE (MasterKey = " & iPK & ")"

With lstEquipmentRental.ListItems
    l = 0
    For i = 1 To .Count
        If CDbl(.Item(i).SubItems(2)) <> 0 And CDbl(IIf(IsNumeric(.Item(i).SubItems(5)) = False, 0, .Item(i).SubItems(5))) > 0 Then
            l = l + 1
            ConnOmega.Execute "INSERT INTO tbl_Operation_Passport_EquipmentRental " & _
                              " (MasterKey, Line, EquipmentKey, Qty) " & _
                              " VALUES (" & iPK & ", " & l & ", " & _
                              " " & .Item(i).SubItems(2) & ", " & _
                              " " & CDbl(.Item(i).SubItems(5)) & ")"
        End If
    Next i
End With

CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER Trim(txtPassportNo.Text), "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()

End Sub


Private Sub PRESS_F8()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If imgPosted.Visible = True Then MsgBox "Already Posted!                ", vbCritical, "Error...": Exit Sub
If AccessRights("Registration", "Post") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If MsgBox("ARE YOU SURE IN POSTING THIS TRANSACTION?                   ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
On Error GoTo PG:
ConnOmega.Execute "UPDATE tbl_Operation_Passport " & _
                  " SET PostedRegistration = 1, " & _
                  " RegCheckIn = '" & Now & "', " & _
                  " RegCheckInBy = '" & gbl_CompleteName & "' " & _
                  " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"

If Trim(txtLockerKeyNo.Text) <> "" Then
    sCtrl = ""
    s = "SELECT TOP 1 CtrlNo " & _
        " FROM tbl_Operation_LockerRoom " & _
        " WHERE (Year(DDate) = " & Format(FormatDateTime(txtDate.Text, vbShortDate), "yyyy") & ") " & _
        " ORDER BY CtrlNo DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        sCtrl = Format(CDbl(rs!CtrlNo) + 1, "0000000#")
    Else
        sCtrl = Format(FormatDateTime(txtDate.Text, vbShortDate), "yyyy") & "0000"
    End If
    rs.Close
    Do
        s = "SELECT tbl_Operation_LockerRoom.* " & _
            " FROM tbl_Operation_LockerRoom " & _
            " WHERE (CtrlNo = '" & sCtrl & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If
        rs.Close
        sCtrl = Format(CDbl(sCtrl) + 1, "0000000#")
    Loop
    ConnOmega.Execute "INSERT INTO tbl_Operation_LockerRoom " & _
                      " (CtrlNo, DDate, PassportKey, LockerKeyNo, LastModified) " & _
                      " VALUES ('" & sCtrl & "', '" & FormatDateTime(txtDate.Text, vbShortDate) & "', " & _
                      " " & Statusbar1.Panels(1).Text & ", '" & Trim(txtLockerKeyNo.Text) & "', " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
End If
BROWSER GetSetting(App.EXEName, "OperationRegPassNo", "OperationRegPassNo", ""), "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub


Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picAddBagDrop.Visible = True Then cmdCancelBagTag_Click: Exit Sub
    Unload Me
Else
    If picSLine.Visible = True Then
         With lstEquipmentRental.ListItems
            If TRANS_DETAIL = is_DET_ADDING Then
                TRANS_DETAIL = is_DET_REFRESH
                With lstEquipmentRental.ListItems
                    If .Count = 1 Then
                        .Item(iRow).SubItems(1) = " "
                        .Item(iRow).SubItems(2) = "0"
                        .Item(iRow).SubItems(3) = " "
                        .Item(iRow).SubItems(4) = " "
                        .Item(iRow).SubItems(5) = " "
                    Else
                        .Remove iRow
                        iRow = .Count
                    End If
                End With
            End If
            If TRANS_DETAIL = is_DET_EDITTING Then
                With lstEquipmentRental.ListItems
                    .Item(iRow).SubItems(2) = txtEQKey1.Text
                    .Item(iRow).SubItems(3) = txtEquipment1.Text
                    .Item(iRow).SubItems(4) = txtUOM1.Text
                    .Item(iRow).SubItems(5) = txtQty1.Text
                End With
            End If
            lstEquipmentRental.ListItems(iRow).EnsureVisible
            lstEquipmentRental.ListItems(iRow).Selected = True
            picToolbar.Enabled = True
            picMain.Enabled = True
            picSLine.Visible = False
            lstEquipmentRental.SetFocus
        End With
        Exit Sub
    End If
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "OperationRegPassNo", "OperationRegPassNo", ""), "is_LOAD"
    If Trim(txtPassportNo.Text) = "" Then BROWSER GetSetting(App.EXEName, "OperationRegPassNo", "OperationRegPassNo", ""), "is_HOME"
End If
End Sub

Private Sub CLEARDETAIL()
With lstEquipmentRental.ListItems
    .Clear
    Set x = .Add()
    x.Text = ""
    x.SubItems(1) = " " '#
    x.SubItems(2) = "0" 'EquipmentKey
    x.SubItems(3) = " " 'Equipment
    x.SubItems(4) = " " 'UOM
    x.SubItems(5) = " " 'Qty
End With
End Sub

Private Sub CLEARTEXT()
iPK = 0
iMemberKey = 0
iPlayerType = 0
txtDate.Text = ""
txtPassportNo.Text = ""
txtBagTagNo.Text = ""
txtPlayerName.Text = ""
txtMemberName.Text = ""
txtTeeTime.Text = ""
txtFlightNo.Text = ""
txtLockerKeyNo.Text = ""
txtCheckInBy.Text = ""
cmbPlayerType.ListIndex = -1
imgPosted.Visible = False
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
imgPicture.Picture = LoadPicture("")
CLEARDETAIL
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtDate.Locked = True
txtPassportNo.Locked = bln
txtPlayerName.Locked = bln
txtMemberName.Locked = True
txtTeeTime.Locked = bln
txtFlightNo.Locked = bln
txtLockerKeyNo.Locked = bln
txtCheckInBy.Locked = True
cmbPlayerType.Locked = bln
txtBagTagNo.Locked = True
txtUOM.Locked = True
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

Private Sub b8TitleBar5_CLoseClick()
cmdCancelBagTag_Click
End Sub


Private Sub cmbEquipment_Click()
If cmbEquipment.ListIndex = -1 Then Exit Sub
If TRANS_DETAIL = is_DET_ADDING Or TRANS_DETAIL = is_DET_EDITTING Then
    txtEQKey.Text = cmbEquipment.ItemData(cmbEquipment.ListIndex)
    With lstEquipmentRental.ListItems
        .Item(iRow).SubItems(3) = cmbEquipment.List(cmbEquipment.ListIndex)
    End With
    txtUOM.Text = ""
    t = "SELECT tbl_Operation_Equipment_For_Rent.* " & _
        " FROM tbl_Operation_Equipment_For_Rent " & _
        " WHERE (PK = " & cmbEquipment.ItemData(cmbEquipment.ListIndex) & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtUOM.Text = rt!UnitOfMeasure
    End If
    rt.Close
End If
End Sub


Private Sub cmbPlayerType_Click()
If cmbPlayerType.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    iPlayerType = cmbPlayerType.ItemData(cmbPlayerType.ListIndex)
End If
End Sub

Private Sub cmdCancelBagTag_Click()
picMain.Enabled = True
picToolbar.Enabled = True
picAddBagDrop.Visible = False
End Sub

Private Sub cmdOKBagTag_Click()
If lstResulBagTag.ListIndex = -1 Then Exit Sub
CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
iPK = lstResulBagTag.ItemData(lstResulBagTag.ListIndex)
txtDate.Text = ""
txtPlayerName.Text = ""
txtMemberName.Text = ""
cmbPlayerType.ListIndex = -1
t = "SELECT PK, DateAll, PlayerName, " & _
    " GuestOf, MemberKey, PlayerTypeKey " & _
    " From dbo.tbl_Operation_Passport " & _
    " WHERE (PK = " & iPK & ")"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then
    txtDate.Text = Format(rt!DateAll, "mm/dd/yyyy")
    txtPlayerName.Text = rt!PlayerName
    iPlayerType = rt!PlayerTypeKey
    cmbPlayerType.ListIndex = rt!PlayerTypeKey - 1
    iMemberKey = IIf(IsNull(rt!MemberKey), 0, rt!MemberKey)
    txtMemberName.Text = IIf(IsNull(rt!GuestOf), "", rt!GuestOf)
    imgPicture.Picture = LoadPicture(SHOW_IMAGES(IIf(IsNull(rt!MemberKey), 0, rt!MemberKey), 0, "Member ID Number"))
End If
rt.Close
txtCheckInBy.Text = gbl_CompleteName
Statusbar1.Panels(1).Text = iPK
cmdCancelBagTag_Click
txtPassportNo.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert:   PRESS_INSERT
    Case vbKeyF2:       PRESS_F2
    Case vbKeyDelete:   PRESS_DELETE
    Case vbKeyF5:       PRESS_F5
    Case vbKeyF6:       PRESS_F6
    Case vbKeyF8:       PRESS_F8
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "OperationRegPassNo", "OperationRegPassNo", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "OperationRegPassNo", "OperationRegPassNo", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "OperationRegPassNo", "OperationRegPassNo", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "OperationRegPassNo", "OperationRegPassNo", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.Height - Me.Height) / 3
Me.Left = (MainForm.Width - Me.Width) / 5
POPULATE_COMBO "PK", "PlayerType", "tbl_Operation_PlayerType", "PK", cmbPlayerType
POPULATE_COMBO "PK", "EquipmentName", "tbl_Operation_Equipment_For_Rent", "PK", cmbEquipment
iFocus = 0
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANS_DETAIL = is_DET_REFRESH
BROWSER GetSetting(App.EXEName, "OperationRegPassNo", "OperationRegPassNo", ""), "is_LOAD"
If Trim(txtPassportNo.Text) = "" Then BROWSER GetSetting(App.EXEName, "OperationRegPassNo", "OperationRegPassNo", ""), "is_HOME"
tmp = SetWindowLong(txtSearchBagTag.hwnd, GWL_STYLE, GetWindowLong(txtSearchBagTag.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPlayerName.hwnd, GWL_STYLE, GetWindowLong(txtPlayerName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtMemberName.hwnd, GWL_STYLE, GetWindowLong(txtMemberName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picAddBagDrop.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstEquipmentRental_Click()
If imgPosted.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
iFocus = 1
iRow = lstEquipmentRental.SelectedItem.Index
TRANS_DETAIL = is_DET_REFRESH
If CDbl(lstEquipmentRental.ListItems.Item(iRow).SubItems(2)) = 0 Then
    TOOLBARFUNC 4
Else
    TOOLBARFUNC 5
End If
End Sub

Private Sub lstEquipmentRental_GotFocus()
If imgPosted.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
iFocus = 1
iRow = lstEquipmentRental.SelectedItem.Index
TRANS_DETAIL = is_DET_REFRESH
If CDbl(lstEquipmentRental.ListItems.Item(iRow).SubItems(2)) = 0 Then
    TOOLBARFUNC 4
Else
    TOOLBARFUNC 5
End If
End Sub

Private Sub lstEquipmentRental_ItemClick(ByVal Item As MSComctlLib.ListItem)
iRow = lstEquipmentRental.SelectedItem.Index
End Sub

Private Sub lstEquipmentRental_LostFocus()
iFocus = 0
End Sub

Private Sub lstResulBagTag_Click()
If lstResulBagTag.ListIndex = -1 Then imgPictureSearchAdd.Picture = LoadPicture(""): Exit Sub
imgPictureSearchAdd.Picture = LoadPicture("")
t = "SELECT tbl_Operation_Passport.* " & _
    " FROM tbl_Operation_Passport " & _
    " WHERE (PK= " & lstResulBagTag.ItemData(lstResulBagTag.ListIndex) & ")"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then
    imgPictureSearchAdd.Picture = LoadPicture(SHOW_IMAGES(IIf(IsNull(rt!MemberKey), 0, rt!MemberKey), 0, "Member ID Number"))
End If
rt.Close
End Sub

Private Sub lstResulBagTag_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKBagTag_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "OperationRegPassNo", "OperationRegPassNo", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "OperationRegPassNo", "OperationRegPassNo", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "OperationRegPassNo", "OperationRegPassNo", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "OperationRegPassNo", "OperationRegPassNo", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Post":    PRESS_F8
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtEQKey_Change()
If TRANS_DETAIL = is_DET_ADDING Or TRANS_DETAIL = is_DET_EDITTING Then
    With lstEquipmentRental.ListItems
        .Item(iRow).SubItems(2) = txtEQKey.Text
    End With
End If
End Sub

Private Sub txtFlightNo_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtQty_Change()
If TRANS_DETAIL = is_DET_ADDING Or TRANS_DETAIL = is_DET_EDITTING Then
    With lstEquipmentRental.ListItems
        .Item(iRow).SubItems(5) = RETURNTEXTVALUE(txtQty)
    End With
End If
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    With lstEquipmentRental.ListItems
        If TRANS_DETAIL = is_DET_ADDING Then
            TRANS_DETAIL = is_DET_REFRESH
            With lstEquipmentRental.ListItems
                Set x = .Add()
                x.Text = ""
                x.SubItems(1) = Format(.Count, "0#")
                x.SubItems(2) = "0"
                x.SubItems(3) = " "
                x.SubItems(4) = " "
                x.SubItems(5) = " "
                iRow = .Count
            End With
            lstEquipmentRental.ListItems(iRow).EnsureVisible
            lstEquipmentRental.ListItems(iRow).Selected = True
            cmbEquipment.ListIndex = -1
            txtEQKey.Text = ""
            txtUOM.Text = ""
            txtQty.Text = ""
            TRANS_DETAIL = is_DET_ADDING
            cmbEquipment.SetFocus
        End If
        If TRANS_DETAIL = is_DET_EDITTING Then
            
        End If
    End With
End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtSearchBagTag_Change()
If Trim(txtSearchBagTag.Text) = "" Then lstResulBagTag.Clear: imgPictureSearchAdd.Picture = LoadPicture(""): Exit Sub
lstResulBagTag.Clear: imgPictureSearchAdd.Picture = LoadPicture("")
If iAddType = 1 Then
    s = "SELECT PK, BagTagNo " & _
        " From dbo.tbl_Operation_Passport " & _
        " WHERE (PostedBagDrop = 1) " & _
        " AND (RegistrationAdded = 0) " & _
        " AND (BagTagNo LIKE '" & FORMATSQL(Trim(txtSearchBagTag.Text)) & "%') " & _
        " ORDER BY BagTagNo"
ElseIf iAddType = 2 Then
    s = "SELECT PK, PlayerName as BagTagNo " & _
        " From dbo.tbl_Operation_Passport " & _
        " WHERE (PostedBagDrop = 1) " & _
        " AND (RegistrationAdded = 0) " & _
        " AND (PlayerName LIKE '" & FORMATSQL(Trim(txtSearchBagTag.Text)) & "%') " & _
        " ORDER BY PlayerName"
End If
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResulBagTag.AddItem rs!BagTagNo
    lstResulBagTag.ItemData(lstResulBagTag.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResulBagTag.ListCount Then lstResulBagTag.ListIndex = 0
End Sub

Private Sub txtSearchBagTag_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResulBagTag.SetFocus
End Sub

Private Sub txtUOM_Change()
If TRANS_DETAIL = is_DET_ADDING Or TRANS_DETAIL = is_DET_EDITTING Then
    With lstEquipmentRental.ListItems
        .Item(iRow).SubItems(4) = txtUOM.Text
    End With
End If
End Sub
