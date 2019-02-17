VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelAbsentLateUndertime 
   Appearance      =   0  'Flat
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11430
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
   ScaleHeight     =   5805
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picSearchLine 
      Height          =   3135
      Left            =   2160
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5530
      BackColor       =   15266266
      Begin VB.CommandButton cmdOKSLine 
         Height          =   480
         Left            =   480
         Picture         =   "frmPersonnelAbsentLateUndertime.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2520
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelSLine 
         Height          =   480
         Left            =   2280
         Picture         =   "frmPersonnelAbsentLateUndertime.frx":0672
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2520
         Width           =   1560
      End
      Begin VB.TextBox txtSearchSLine 
         Height          =   315
         Left            =   120
         MaxLength       =   100
         TabIndex        =   12
         Top             =   120
         Width           =   3975
      End
      Begin VB.ListBox lstResultSLine 
         Height          =   2010
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   3975
      End
   End
   Begin VB.PictureBox picSLine 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   720
      ScaleHeight     =   855
      ScaleWidth      =   9795
      TabIndex        =   15
      Top             =   1560
      Visible         =   0   'False
      Width           =   9795
      Begin RPVGCC.b8Container picSLine1 
         Height          =   855
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   1508
         BackColor       =   8438015
         Begin VB.TextBox txtMinutes1 
            Height          =   315
            Left            =   4080
            MaxLength       =   100
            TabIndex        =   32
            Top             =   0
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.TextBox txtHours1 
            Height          =   315
            Left            =   3840
            MaxLength       =   100
            TabIndex        =   31
            Top             =   0
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.TextBox cmbType1 
            Height          =   315
            Left            =   3600
            MaxLength       =   100
            TabIndex        =   30
            Top             =   0
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.TextBox txtType1 
            Height          =   315
            Left            =   3360
            MaxLength       =   100
            TabIndex        =   29
            Top             =   0
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.TextBox txtName1 
            Height          =   315
            Left            =   3120
            MaxLength       =   100
            TabIndex        =   28
            Top             =   0
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.TextBox txtIDNo1 
            Height          =   315
            Left            =   2880
            MaxLength       =   100
            TabIndex        =   27
            Top             =   0
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.TextBox txtEmp1 
            Height          =   315
            Left            =   2640
            MaxLength       =   100
            TabIndex        =   26
            Top             =   0
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.TextBox txtIDNo 
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   25
            Top             =   360
            Width           =   1260
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   24
            Top             =   360
            Width           =   4260
         End
         Begin VB.TextBox txtMinutes 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8760
            MaxLength       =   100
            TabIndex        =   23
            Top             =   360
            Width           =   900
         End
         Begin VB.TextBox txtHours 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7800
            MaxLength       =   100
            TabIndex        =   22
            Top             =   360
            Width           =   900
         End
         Begin VB.ComboBox cmbType 
            Height          =   315
            Left            =   5760
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   360
            Width           =   1960
         End
         Begin VB.TextBox txtEmp 
            Height          =   315
            Left            =   1080
            MaxLength       =   100
            TabIndex        =   20
            Top             =   0
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.TextBox txtType 
            Height          =   315
            Left            =   6480
            MaxLength       =   100
            TabIndex        =   19
            Top             =   0
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.TextBox txtHours2 
            Height          =   315
            Left            =   8880
            MaxLength       =   100
            TabIndex        =   18
            Top             =   0
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.TextBox txtMinutes2 
            Height          =   315
            Left            =   9120
            MaxLength       =   100
            TabIndex        =   17
            Top             =   0
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ID Number"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1440
            TabIndex        =   36
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Minutes"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   8760
            TabIndex        =   35
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Hour/s"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7800
            TabIndex        =   34
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5760
            TabIndex        =   33
            Top             =   120
            Width           =   1935
         End
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   0
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   1
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
         MouseIcon       =   "frmPersonnelAbsentLateUndertime.frx":0DCE
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   9900
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   2
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmPersonnelAbsentLateUndertime.frx":10E8
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12120
      Top             =   2040
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
            Picture         =   "frmPersonnelAbsentLateUndertime.frx":17FB
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAbsentLateUndertime.frx":24D5
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAbsentLateUndertime.frx":31AF
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAbsentLateUndertime.frx":3E89
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAbsentLateUndertime.frx":4B63
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAbsentLateUndertime.frx":583D
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAbsentLateUndertime.frx":6517
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAbsentLateUndertime.frx":71F1
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAbsentLateUndertime.frx":7ECB
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAbsentLateUndertime.frx":87A5
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAbsentLateUndertime.frx":947F
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAbsentLateUndertime.frx":A159
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAbsentLateUndertime.frx":AE33
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAbsentLateUndertime.frx":BB0D
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAbsentLateUndertime.frx":C7E7
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   5490
      Width           =   11430
      _ExtentX        =   20161
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
   Begin RPVGCC.b8Container picSearch 
      Height          =   4455
      Left            =   3600
      TabIndex        =   38
      Top             =   600
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7858
      BackColor       =   15266266
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         MaxLength       =   100
         TabIndex        =   44
         Top             =   480
         Width           =   3975
      End
      Begin VB.CommandButton cmdCancel 
         Height          =   480
         Left            =   2160
         Picture         =   "frmPersonnelAbsentLateUndertime.frx":D4C1
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   3720
         Width           =   1560
      End
      Begin VB.CommandButton cmdOK 
         Height          =   480
         Left            =   480
         Picture         =   "frmPersonnelAbsentLateUndertime.frx":DC1D
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   3720
         Width           =   1560
      End
      Begin VB.ListBox lstResult 
         Height          =   1230
         Left            =   120
         TabIndex        =   40
         Top             =   840
         Width           =   3975
      End
      Begin VB.ListBox lstCtrl 
         Height          =   1425
         Left            =   120
         TabIndex        =   39
         Top             =   2160
         Width           =   3975
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   45
         TabIndex        =   41
         Top             =   45
         Width           =   4125
         _ExtentX        =   7276
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
         Icon            =   "frmPersonnelAbsentLateUndertime.frx":E28F
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   720
      ScaleHeight     =   3855
      ScaleWidth      =   9855
      TabIndex        =   4
      Top             =   1200
      Width           =   9855
      Begin VB.ComboBox cmbDivision 
         Height          =   315
         ItemData        =   "frmPersonnelAbsentLateUndertime.frx":E829
         Left            =   1080
         List            =   "frmPersonnelAbsentLateUndertime.frx":E82B
         TabIndex        =   46
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox txtCtrl 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   7
         Top             =   0
         Width           =   1260
      End
      Begin VB.TextBox txtDateApplied 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   6
         Top             =   360
         Width           =   1260
      End
      Begin MSComctlLib.ListView lstDetail 
         Height          =   2670
         Left            =   0
         TabIndex        =   5
         Top             =   1200
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   4710
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
         NumItems        =   9
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
            Text            =   "EmpKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ID #"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Name"
            Object.Width           =   7762
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "TypeKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Type"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Hour/s"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Minutes"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   45
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   9
         Top             =   45
         Width           =   735
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Applied"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   8
         Top             =   405
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmPersonnelAbsentLateUndertime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE     As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim TRANS_DETAIL        As Long
Const is_DET_REFRESH = 10
Const is_DET_ADDING = 11
Const is_DET_EDITTING = 12

Dim ROW                 As Long
Dim ListViewFocus       As Long
Dim txtNameFocus        As Long
Dim iDivision           As Long

Dim x, i, Arr, sCtrl, iPK, iLine, dblHours, dlMins


Private Sub BROWSER(Ctrl, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If Ctrl <> "" Then
            s = "SELECT TOP 1 tbl_Personnel_AbsentLateUndertime.* " & _
                " FROM tbl_Personnel_AbsentLateUndertime " & _
                " WHERE (Ctrl = '" & Ctrl & "') " & _
                " ORDER BY Ctrl"
        Else
            s = "SELECT TOP 1 tbl_Personnel_AbsentLateUndertime.* " & _
                " FROM tbl_Personnel_AbsentLateUndertime " & _
                " ORDER BY Ctrl"
        End If
    Case "is_HOME"
        If picSLine.Visible = True Then Exit Sub
        If picSearchLine.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Personnel_AbsentLateUndertime.* " & _
            " FROM tbl_Personnel_AbsentLateUndertime " & _
            " ORDER BY Ctrl"
    Case "is_PAGEUP"
        If picSLine.Visible = True Then Exit Sub
        If picSearchLine.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Personnel_AbsentLateUndertime.* " & _
            " FROM tbl_Personnel_AbsentLateUndertime " & _
            " WHERE (Ctrl < '" & Ctrl & "') " & _
            " ORDER BY Ctrl DESC"
    Case "is_PAGEDOWN"
        If picSLine.Visible = True Then Exit Sub
        If picSearchLine.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Personnel_AbsentLateUndertime.* " & _
            " FROM tbl_Personnel_AbsentLateUndertime " & _
            " WHERE (Ctrl > '" & Ctrl & "') " & _
            " ORDER BY Ctrl"
    Case "is_END"
        If picSLine.Visible = True Then Exit Sub
        If picSearchLine.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Personnel_AbsentLateUndertime.* " & _
            " FROM tbl_Personnel_AbsentLateUndertime " & _
            " ORDER BY Ctrl DESC"
    Case Else: Exit Sub
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtCtrl.Text = rs!Ctrl
    iDivision = rs!DivisionKey
    txtDateApplied.Text = Format(rs!DateApplied, "mm/dd/yyyy")
    
    If DIV_NAME(rs!DivisionKey) <> "" Then
        Arr = Split(DIV_NAME(rs!DivisionKey), ";", -1, 1)
        cmbDivision.Text = CStr(Arr(1))
    Else
        cmbDivision.Text = ""
    End If
    
    StatusBar1.Panels(1).Text = rs!PK
    StatusBar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    
    iLine = 0: CLEARDETAIL
    t = "SELECT tbl_Personnel_AbsentLateUndertime_Details.EmployeeKey, " & _
        " tbl_Personnel_IDNumber.IDNumber, " & _
        " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
        " tbl_Personnel_AbsentLateUndertime_Details.AbsType, " & _
        " tbl_Personnel_AbsentLateUndertime_Details.Hours, " & _
        " tbl_Personnel_AbsentLateUndertime_Details.Minutes " & _
        " FROM tbl_Personnel_AbsentLateUndertime_Details LEFT OUTER JOIN " & _
        " tbl_Personnel_IDNumber ON tbl_Personnel_AbsentLateUndertime_Details.EmployeeKey = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
        " Where (tbl_Personnel_AbsentLateUndertime_Details.MasterKey = " & rs!PK & ") " & _
        " ORDER BY tbl_Personnel_Information.LastName, tbl_Personnel_Information.FirstName, tbl_Personnel_Information.MiddleName, tbl_Personnel_IDNumber.IDNumber"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        With lstDetail.ListItems
            .Clear
            While Not rt.EOF
                iLine = iLine + 1
                Set x = .Add()
                x.Text = ""
                x.SubItems(1) = Format(iLine, "0#")
                x.SubItems(2) = rt!EmployeeKey
                x.SubItems(3) = rt!IDNumber
                x.SubItems(4) = rt!EmployeeName
                x.SubItems(5) = rt!AbsType
                x.SubItems(6) = IIf(rt!AbsType = 1, "Absent", IIf(rt!AbsType = 2, "Late", IIf(rt!AbsType = 3, "Undertime", "")))
                x.SubItems(7) = rt!Hours
                x.SubItems(8) = rt!Minutes
                rt.MoveNext
            Wend
        End With
    End If
    rt.Close
    
    imgPosted.Visible = IIf(rs!Posted = 1, True, False)
    Toolbar1.Buttons(19).Caption = IIf(rs!Posted = 1, "UnPost", " Post ")
    Toolbar1.Buttons(19).Image = IIf(rs!Posted = 1, 11, 10)
    
    SaveSetting App.EXEName, "AbsentLateUndertimeCtrlV2", "AbsLatUnderCtrlV2", rs!Ctrl
    
End If
rs.Close
End Sub

Private Function CHECK_ABSENT_DUPLICATE(iEmpKey, iType, iRow) As String
CHECK_ABSENT_DUPLICATE = "False|0"
With lstDetail.ListItems
    For i = 1 To .Count
        If CDbl(iEmpKey) = CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2))) Then
            If CDbl(iType) = CDbl(IIf(IsNumeric(.Item(i).SubItems(5)) = False, 0, .Item(i).SubItems(5))) Then
                If CDbl(iRow) <> CDbl(i) Then
                    CHECK_ABSENT_DUPLICATE = "True|" & CStr(Format(i, "0#"))
                    Exit Function
                End If
            End If
        End If
    Next i
End With
End Function

Private Sub PRESS_INSERT()
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    If picSLine.Visible = True Then Exit Sub
    If picSearchLine.Visible = True Then Exit Sub
    If AccessRights("Absent/Late/Undertime Employee", "Add") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    'If imgPosted.Visible = True Then MsgBox "Already Posted!                   ", vbCritical, "Error...": Exit Sub
    CLEARTEXT
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_ADDING
    'Me.Caption = "Absent / Late / Undertime Employee - New"
    txtDateApplied.SetFocus
Else
    With lstDetail.ListItems
        If ListViewFocus = 0 Then Exit Sub
        If IsDate(txtDateApplied.Text) = False Then Exit Sub
        If TRANS_DETAIL <> is_DET_REFRESH Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        If CDbl(.Item(ROW).SubItems(2)) <> 0 Then
            Set x = .Add()
            x.Text = ""
            x.SubItems(1) = Format(.Count, "0#")
            x.SubItems(2) = "0"
            x.SubItems(3) = " "
            x.SubItems(4) = " "
            x.SubItems(5) = "0"
            x.SubItems(6) = " "
            x.SubItems(7) = " "
            x.SubItems(8) = " "
            ROW = .Count
        Else
            .Item(1).SubItems(1) = Format(.Count, "0#")
            ROW = .Count
        End If
        lstDetail.ListItems(ROW).EnsureVisible
        lstDetail.ListItems(ROW).Selected = True
        txtEmp.Text = ""
        txtIDNo.Text = ""
        txtName.Text = ""
        txtType.Text = ""
        cmbType.ListIndex = -1
        txtHours.Text = ""
        txtMinutes.Text = ""
'        CLEARDETAIL
        picMain.Enabled = False
        picToolbar.Enabled = False
        picSLine.ZOrder 0
        picSLine.Visible = True
        TRANS_DETAIL = is_DET_ADDING
        txtName.SetFocus
    End With
End If
End Sub

Private Sub PRESS_F2()
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    If StatusBar1.Panels(1).Text = "" Then Exit Sub
    If picSLine.Visible = True Then Exit Sub
    If picSearchLine.Visible = True Then Exit Sub
    If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
    If AccessRights("Absent/Late/Undertime Employee", "Edit") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    'If imgPost.Visible = True Then MsgBox "Already Posted!                   ", vbCritical, "Error...": Exit Sub
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_EDITTING
    If ListViewFocus = 1 Then lstDetail_Click
    'Me.Caption = "Absent / Late / Undertime Employee - Edit"
Else
    With lstDetail.ListItems
        If ListViewFocus = 0 Then Exit Sub
        If IsDate(txtDateApplied.Text) = False Then Exit Sub
        If TRANS_DETAIL <> is_DET_REFRESH Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        If CDbl(.Item(ROW).SubItems(2)) <> 0 Then
            txtEmp.Text = .Item(ROW).SubItems(2)
            txtIDNo.Text = .Item(ROW).SubItems(3)
            txtName.Text = .Item(ROW).SubItems(4)
            txtType.Text = .Item(ROW).SubItems(5)
            cmbType.ListIndex = .Item(ROW).SubItems(5) - 1
            txtHours.Text = .Item(ROW).SubItems(7)
            txtMinutes.Text = .Item(ROW).SubItems(8)

            txtEmp1.Text = .Item(ROW).SubItems(2)
            txtIDNo1.Text = .Item(ROW).SubItems(3)
            txtName1.Text = .Item(ROW).SubItems(4)
            txtType1.Text = .Item(ROW).SubItems(5)
            cmbType1.Text = .Item(ROW).SubItems(6)
            txtHours1.Text = .Item(ROW).SubItems(7)
            txtMinutes1.Text = .Item(ROW).SubItems(8)

        End If
        picMain.Enabled = False
        picToolbar.Enabled = False
        picSLine.ZOrder 0
        picSLine.Visible = True
        TRANS_DETAIL = is_DET_EDITTING
        txtName.SetFocus
    End With
End If
End Sub

Private Sub PRESS_DELETE()
If picSearch.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    If picSLine.Visible = True Then Exit Sub
    If picSearchLine.Visible = True Then Exit Sub
    If StatusBar1.Panels(1).Text = "" Then Exit Sub
    If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
    If AccessRights("Absent/Late/Undertime Employee", "Delete") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    'If imgPost.Visible = True Then MsgBox "Already Posted!                   ", vbCritical, "Error...": Exit Sub
    If MsgBox("ARE YOU SURE TO DELETE THIS TRANSACTION?                             ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    ConnOmega.Execute "DELETE FROM tbl_Personnel_AbsentLateUndertime WHERE (PK = " & StatusBar1.Panels(1).Text & ")"
    CLEARTEXT
    BROWSER GetSetting(App.EXEName, "AbsentLateUndertimeCtrlV2", "AbsLatUnderCtrlV2", ""), "is_PAGEDOWN"
    If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "AbsentLateUndertimeCtrlV2", "AbsLatUnderCtrlV2", ""), "is_HOME"
Else
    With lstDetail.ListItems
        If ListViewFocus = 0 Then Exit Sub
        If TRANS_DETAIL <> is_DET_REFRESH Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        If .Count > 1 Then
            .Remove ROW
            ROW = IIf(CDbl(.Count) < CDbl(ROW), .Count, ROW)
        Else
            .Item(ROW).SubItems(1) = " "
            .Item(ROW).SubItems(2) = "0"
            .Item(ROW).SubItems(3) = " "
            .Item(ROW).SubItems(4) = " "
            .Item(ROW).SubItems(5) = "0"
            .Item(ROW).SubItems(6) = " "
            .Item(ROW).SubItems(7) = " "
            .Item(ROW).SubItems(8) = " "
        End If
        lstDetail.ListItems(ROW).EnsureVisible
        lstDetail.ListItems(ROW).Selected = True
    End With
End If
End Sub

Private Sub PRESS_F5()
If picSLine.Visible = True Then Exit Sub
If picSearchLine.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If IsDate(txtDateApplied.Text) = False Then MsgBox "Please Supply a Valid Date!                   ", vbCritical, "Error...": txtDateApplied.SetFocus: Exit Sub
If iDivision = 0 Then MsgBox "Please select Division!                   ", vbCritical, "Error...": cmbDivision.SetFocus: Exit Sub
With lstDetail.ListItems
    For i = 1 To .Count
        If CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2))) <> 0 Then
            If CDbl(IIf(IsNumeric(.Item(i).SubItems(7)) = False, 0, .Item(i).SubItems(7))) <= 0 Then
                If CDbl(IIf(IsNumeric(.Item(i).SubItems(8)) = False, 0, .Item(i).SubItems(8))) <= 0 Then
                    MsgBox "Found Invalid Entry!                        ", vbCritical, "Error..."
                    lstDetail.ListItems(i).EnsureVisible
                    lstDetail.ListItems(i).Selected = True
                    lstDetail.SetFocus
                    Exit Sub
                End If
            End If
            dblHours = CDbl(IIf(IsNumeric(.Item(i).SubItems(7)) = False, 0, .Item(i).SubItems(7)))
            dlMins = CDbl(IIf(IsNumeric(.Item(i).SubItems(8)) = False, 0, .Item(i).SubItems(8)))
            If CDbl(dblHours) > 0 Then
                Arr = Split(Format(dblHours, "#0.00"), ".", -1, 1)
                If CDbl(Arr(1)) > 0 Then
                    MsgBox "Found Invalid Entry!                        " & vbCrLf & vbCrLf & "Invalid Hours!", vbCritical, "Error..."
                    lstDetail.ListItems(i).EnsureVisible
                    lstDetail.ListItems(i).Selected = True
                    lstDetail.SetFocus
                    Exit Sub
                End If
            End If
            If CDbl(dlMins) > 0 Then
                Arr = Split(Format(dlMins, "#0.00"), ".", -1, 1)
                If CDbl(Arr(1)) > 0 Then
                    MsgBox "Found Invalid Entry!                        " & vbCrLf & vbCrLf & "Invalid Minutes!", vbCritical, "Error..."
                    lstDetail.ListItems(i).EnsureVisible
                    lstDetail.ListItems(i).Selected = True
                    lstDetail.SetFocus
                    Exit Sub
                End If
            End If
        End If
    Next i
    
End With

On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    sCtrl = ""
    s = "SELECT TOP 1 Ctrl " & _
        " FROM tbl_Personnel_AbsentLateUndertime " & _
        " ORDER BY Ctrl DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        sCtrl = Format(CDbl(rs!Ctrl) + 1, "0000000#")
    Else
        sCtrl = Format(FormatDateTime(txtDateApplied.Text, vbShortDate), "yyyy") & "0000"
    End If
    rs.Close
    
    Do
        s = "SELECT tbl_Personnel_AbsentLateUndertime.* " & _
            " FROM tbl_Personnel_AbsentLateUndertime " & _
            " WHERE (Ctrl = '" & sCtrl & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If
        rs.Close
        sCtrl = Format(CDbl(sCtrl) + 1, "0000000#")
    Loop
    
    ConnOmega.Execute "INSERT INTO tbl_Personnel_AbsentLateUndertime " & _
                      " (Ctrl, DateApplied, LastModified, DivisionKey) " & _
                      " VALUES ('" & sCtrl & "', " & _
                      " '" & FormatDateTime(txtDateApplied.Text, vbShortDate) & "', " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "', " & iDivision & ")"
                      
    iPK = 0: iLine = 0
    s = "SELECT PK " & _
        " FROM tbl_Personnel_AbsentLateUndertime " & _
        " WHERE (Ctrl = '" & sCtrl & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        iPK = rs!PK
    End If
    rs.Close
    
End If
If TRANSACTIONTYPE = is_EDITTING Then
    iPK = StatusBar1.Panels(1).Text
    sCtrl = Trim(txtCtrl.Text)
    iLine = 0
    ConnOmega.Execute "UPDATE tbl_Personnel_AbsentLateUndertime " & _
                      " SET DateApplied = '" & FormatDateTime(txtDateApplied.Text, vbShortDate) & "', " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                      " DivisionKey = " & iDivision & " " & _
                      " WHERE (PK = " & iPK & ")"
    
'    If CDbl(iPK) <> 0 Then
'        ConnOmega.Execute "DELETE FROM tbl_Personnel_AbsentLateUndertime_Details WHERE (MasterKey = " & iPK & ")"
'        With lstDetail.ListItems
'            For i = 1 To .Count
'                If CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2))) <> 0 Then
'                    iLine = iLine + 1
'                    ConnOmega.Execute "INSERT INTO tbl_Personnel_AbsentLateUndertime_Details " & _
'                                      " (MasterKey, Line, EmployeeKey, AbsType, Hours, Minutes) " & _
'                                      " VALUES (" & iPK & ", " & iLine & ", " & .Item(i).SubItems(2) & ", " & _
'                                      " " & .Item(i).SubItems(5) & ", " & CDbl(.Item(i).SubItems(7)) & ", " & _
'                                      " " & CDbl(.Item(i).SubItems(8)) & ")"
'                End If
'            Next i
'        End With
'    End If
    
End If

If CDbl(iPK) <> 0 Then
    ConnOmega.Execute "DELETE FROM tbl_Personnel_AbsentLateUndertime_Details WHERE (MasterKey = " & iPK & ")"
    With lstDetail.ListItems
        For i = 1 To .Count
            If CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2))) <> 0 Then
                iLine = iLine + 1
                ConnOmega.Execute "INSERT INTO tbl_Personnel_AbsentLateUndertime_Details " & _
                                  " (MasterKey, EmployeeKey, AbsType, Hours, Minutes) " & _
                                  " VALUES (" & iPK & ", " & .Item(i).SubItems(2) & ", " & _
                                  " " & .Item(i).SubItems(5) & ", " & CDbl(IIf(IsNumeric(.Item(i).SubItems(7)) = False, 0, .Item(i).SubItems(7))) & ", " & _
                                  " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(8)) = False, 0, .Item(i).SubItems(8))) & ")"
            End If
        Next i
    End With
End If

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
If TRANSACTIONTYPE = is_REFRESH Then
    If picSLine.Visible = True Then Exit Sub
    If picSearchLine.Visible = True Then Exit Sub
    If picSearch.Visible = True Then Exit Sub
    picSearch.ZOrder 0
    txtSearch.Text = ""
    picMain.Enabled = False
    picToolbar.Enabled = False
    picSearch.Visible = True
    txtSearch.SetFocus
Else
    If txtNameFocus = 1 Then
        txtSearchSLine.Text = ""
        picSearchLine.ZOrder 0
        picSearchLine.Visible = True
        txtSearchSLine.SetFocus
    End If
End If
End Sub


Private Sub PRESS_F8()
If picSLine.Visible = True Then Exit Sub
If picSearchLine.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If StatusBar1.Panels(1).Text = "" Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
On Error GoTo PG:
If imgPosted.Visible = False Then
    If AccessRights("Absent/Late/Undertime Employee", "Post") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If MsgBox("ARE YOU SURE IN POSTING THIS TRANSACTION?                ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    ConnOmega.Execute "UPDATE tbl_Personnel_AbsentLateUndertime SET Posted = 1 WHERE (PK = " & StatusBar1.Panels(1).Text & ")"
End If
If imgPosted.Visible = True Then
    If AccessRights("Absent/Late/Undertime Employee", "UnPost") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    
    BROWSER GetSetting(App.EXEName, "AbsentLateUndertimeCtrlV2", "AbsLatUnderCtrlV2", ""), "is_LOAD"
    With lstDetail.ListItems
        For i = 1 To .Count
            If CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2))) <> 0 Then
                t = "SELECT dbo.tbl_Personnel_Hours.PayrollKey, dbo.tbl_Personnel_Hours.EmployeeKey, " & _
                    " dbo.tbl_Personnel_ActionNew.DivisionKey, dbo.tbl_Personnel_Compensation_Period.DateFrom, " & _
                    " dbo.tbl_Personnel_Compensation_Period.DateTo, dbo.tbl_Personnel_Compensation_Period.PayrollDate " & _
                    " FROM  dbo.tbl_Personnel_Hours LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_Compensation_Period ON dbo.tbl_Personnel_Hours.PayrollPeriodKey = dbo.tbl_Personnel_Compensation_Period.PK LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_ActionNew ON dbo.tbl_Personnel_Hours.ActionMemoKey = dbo.tbl_Personnel_ActionNew.PK " & _
                    " WHERE (dbo.tbl_Personnel_Hours.EmployeeKey = " & .Item(i).SubItems(2) & ") " & _
                    " AND (dbo.tbl_Personnel_ActionNew.DivisionKey = " & iDivision & ") " & _
                    " AND (dbo.tbl_Personnel_Compensation_Period.DateFrom <= '" & FormatDateTime(txtDateApplied.Text, vbShortDate) & "') " & _
                    " AND (dbo.tbl_Personnel_Compensation_Period.DateTo >= '" & FormatDateTime(txtDateApplied.Text, vbShortDate) & "')"
                If rt.State = adStateOpen Then rt.Close
                rt.Open t, ConnOmega
                If rt.RecordCount > 0 Then
                    MsgBox "YOU CAN'T UNPOST THIS TRANSACTION! " & vbCrLf & vbCrLf & _
                           .Item(i).SubItems(4) & vbCrLf & _
                           "ALREADY RECORDED IN PERSONNEL HOURS" & vbCrLf & _
                           "UNDER PAYROLL PERIOD " & Format(rt!PayrollDate, "mm/dd/yyyy"), vbExclamation, "Can't Unpost"
                    rt.Close
                    Exit Sub
                End If
                rt.Close
            End If
        Next i
    End With
    If MsgBox("ARE YOU SURE IN UNPOSTING THIS TRANSACTION?                ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    ConnOmega.Execute "UPDATE tbl_Personnel_AbsentLateUndertime SET Posted = 0 WHERE (PK = " & StatusBar1.Panels(1).Text & ")"
End If
BROWSER GetSetting(App.EXEName, "AbsentLateUndertimeCtrlV2", "AbsLatUnderCtrlV2", ""), "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSearch.Visible = True Then cmdCancel_Click: Exit Sub
    Unload Me
Else
    If picSearchLine.Visible = True Then cmdCancelSLine_Click: Exit Sub
    If TRANS_DETAIL = is_DET_ADDING Then
        With lstDetail.ListItems
            If .Count > 1 Then
                .Remove .Count
            Else
                .Item(1).SubItems(1) = " "
                .Item(1).SubItems(2) = "0"
                .Item(1).SubItems(3) = " "
                .Item(1).SubItems(4) = " "
                .Item(1).SubItems(5) = "0"
                .Item(1).SubItems(6) = " "
                .Item(1).SubItems(7) = " "
                .Item(1).SubItems(8) = " "
            End If
        End With
        picSLine.Visible = False
        picMain.Enabled = True
        picToolbar.Enabled = True
        lstDetail.SetFocus
        Exit Sub
    End If
    If TRANS_DETAIL = is_DET_EDITTING Then
        With lstDetail.ListItems
            .Item(ROW).SubItems(2) = txtEmp1.Text
            .Item(ROW).SubItems(3) = txtIDNo1.Text
            .Item(ROW).SubItems(4) = txtName1.Text
            .Item(ROW).SubItems(5) = txtType1.Text
            .Item(ROW).SubItems(6) = cmbType1.Text
            .Item(ROW).SubItems(7) = txtHours1.Text
            .Item(ROW).SubItems(8) = txtMinutes1.Text
        End With
        picSLine.Visible = False
        picMain.Enabled = True
        picToolbar.Enabled = True
        lstDetail.SetFocus
        Exit Sub
    End If
    If ListViewFocus = 1 Then
        txtDateApplied.SetFocus
        Exit Sub
    End If
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    'Me.Caption = "Absent / Late / Undertime Employee - Browse"
    BROWSER GetSetting(App.EXEName, "AbsentLateUndertimeCtrlV2", "AbsLatUnderCtrlV2", ""), "is_LOAD"
    If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "AbsentLateUndertimeCtrlV2", "AbsLatUnderCtrlV2", ""), "is_HOME"
End If
End Sub

Private Function CLEARTEXT()
iDivision = 0
txtCtrl.Text = ""
txtDateApplied.Text = ""
cmbDivision.Text = ""
cmbDivision.ListIndex = -1
StatusBar1.Panels(1).Text = ""
StatusBar1.Panels(2).Text = ""
imgPosted.Visible = False
CLEARDETAIL
End Function

Private Function CLEARDETAIL()
With lstDetail.ListItems
    .Clear
    Set x = .Add()
    x.Text = ""
    x.SubItems(1) = " "
    x.SubItems(2) = "0"
    x.SubItems(3) = " "
    x.SubItems(4) = " "
    x.SubItems(5) = "0"
    x.SubItems(6) = " "
    x.SubItems(7) = " "
    x.SubItems(8) = " "
End With
'txtEmp.Text = ""
'txtIDNo.Text = ""
'txtName.Text = ""
'txtType.Text = ""
'cmbType.ListIndex = -1
'txtHours.Text = ""
'txtMinutes.Text = ""
End Function

Private Function LOCKTEXT(bln As Boolean)
txtCtrl.Locked = True
txtDateApplied.Locked = bln
cmbDivision.Locked = bln
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
        Case 6      '=== NOT EMPTY DETAIL NAME ===
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
cmdCancel_Click
End Sub

Private Sub cmbDivision_Click()
If cmbDivision.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    iDivision = cmbDivision.ItemData(cmbDivision.ListIndex)
End If
End Sub

Private Sub cmbType_Click()
If cmbType.ListIndex = -1 Then Exit Sub
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    txtType.Text = cmbType.ListIndex + 1
    If cmbType.ListIndex = 0 Then
        txtHours.Text = "8"
        txtMinutes.Text = "0"
        txtHours.Locked = True
        txtMinutes.Locked = True
    Else
        txtHours.Locked = False
        txtMinutes.Locked = False
    End If
    With lstDetail.ListItems
        .Item(ROW).SubItems(6) = cmbType.List(cmbType.ListIndex)
    End With
End If
End Sub

Private Sub cmbType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtHours.SetFocus
End Sub

Private Sub cmdCancel_Click()
picMain.Enabled = True
picToolbar.Enabled = True
picSearch.Visible = False
End Sub

Private Sub cmdCancelSLine_Click()
picSearchLine.Visible = False
picSLine.Enabled = True
txtName.SetFocus
End Sub

Private Sub cmdOK_Click()
If lstCtrl.ListIndex = -1 Then Exit Sub
Arr = Split(lstCtrl.List(lstCtrl.ListIndex), "  ", -1, 1)
BROWSER CStr(Arr(0)), "is_LOAD"
cmdCancel_Click
End Sub

Private Sub cmdOKSLine_Click()
If lstResultSLine.ListIndex = -1 Then Exit Sub
Arr = Split(lstResultSLine.List(lstResultSLine.ListIndex), " - ", -1, 1)
txtEmp.Text = lstResultSLine.ItemData(lstResultSLine.ListIndex)
txtIDNo.Text = Arr(0)
txtName.Text = Arr(1)
cmdCancelSLine_Click
cmbType.SetFocus
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
    Case vbKeyF8:       PRESS_F8
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "AbsentLateUndertimeCtrlV2", "AbsLatUnderCtrlV2", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "AbsentLateUndertimeCtrlV2", "AbsLatUnderCtrlV2", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "AbsentLateUndertimeCtrlV2", "AbsLatUnderCtrlV2", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "AbsentLateUndertimeCtrlV2", "AbsLatUnderCtrlV2", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption & " [V2]"
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
'Me.Caption = "Absent / Late / Undertime Employee - Browse"
POPULATE_COMBO "PK", "Description", "tbl_Personnel_Division", "Description", cmbDivision
With cmbType
    .Clear
    .AddItem "Absent"
    .AddItem "Late"
    .AddItem "Undertime"
End With
ListViewFocus = 0
txtNameFocus = 0
ROW = 0
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "AbsentLateUndertimeCtrlV2", "AbsLatUnderCtrlV2", ""), "is_LOAD"
If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "AbsentLateUndertimeCtrlV2", "AbsLatUnderCtrlV2", ""), "is_HOME"
Dim tmp As Long
tmp = SetWindowLong(txtSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearchSLine.hwnd, GWL_STYLE, GetWindowLong(txtSearchSLine.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picSearch.Visible = True Then Cancel = -1
If picSearchLine.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstCtrl_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub lstDetail_Click()
ListViewFocus = 1
ROW = lstDetail.SelectedItem.Index
TRANS_DETAIL = is_DET_REFRESH
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    With lstDetail.ListItems
        If CDbl(.Item(ROW).SubItems(2)) <> 0 Then
            TOOLBARFUNC 5
        Else
            TOOLBARFUNC 4
        End If
    End With
End If
End Sub

Private Sub lstDetail_GotFocus()
ListViewFocus = 1
ROW = lstDetail.SelectedItem.Index
TRANS_DETAIL = is_DET_REFRESH
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    With lstDetail.ListItems
        If CDbl(.Item(ROW).SubItems(2)) <> 0 Then
            TOOLBARFUNC 5
        Else
            TOOLBARFUNC 4
        End If
    End With
End If
End Sub

Private Sub lstDetail_ItemClick(ByVal Item As MSComctlLib.ListItem)
ROW = lstDetail.SelectedItem.Index
End Sub

Private Sub lstDetail_LostFocus()
ListViewFocus = 0
End Sub

Private Sub lstResult_Click()
If lstResult.ListIndex = -1 Then Exit Sub
lstCtrl.Clear
t = "SELECT tbl_Personnel_AbsentLateUndertime.Ctrl, " & _
    " tbl_Personnel_AbsentLateUndertime.DateApplied " & _
    " FROM tbl_Personnel_AbsentLateUndertime LEFT OUTER JOIN " & _
    " tbl_Personnel_AbsentLateUndertime_Details ON tbl_Personnel_AbsentLateUndertime.PK = tbl_Personnel_AbsentLateUndertime_Details.MasterKey " & _
    " Where (tbl_Personnel_AbsentLateUndertime_Details.EmployeeKey = " & lstResult.ItemData(lstResult.ListIndex) & ") " & _
    " ORDER BY tbl_Personnel_AbsentLateUndertime.DateApplied"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
While Not rt.EOF
    lstCtrl.AddItem rt!Ctrl & "  " & Format(rt!DateApplied, "mm/dd/yyyy")
    rt.MoveNext
Wend
rt.Close
If lstCtrl.ListCount Then lstCtrl.ListIndex = 0
End Sub

Private Sub lstResult_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstCtrl.SetFocus
End Sub

Private Sub lstResultSLine_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKSLine_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
        Case "Refresh"
            'ToDo: Add 'Refresh' button code.
            MsgBox "Add 'Refresh' button code."
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "AbsentLateUndertimeCtrlV2", "AbsLatUnderCtrlV2", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "AbsentLateUndertimeCtrlV2", "AbsLatUnderCtrlV2", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "AbsentLateUndertimeCtrlV2", "AbsLatUnderCtrlV2", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "AbsentLateUndertimeCtrlV2", "AbsLatUnderCtrlV2", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":
    Case "Post":    PRESS_F8
    Case "Close":   PRESS_ESCAPE
    Case Else: Exit Sub
End Select
End Sub

Private Sub txtCtrl_GotFocus()
HTEXT txtCtrl
End Sub

Private Sub txtCtrl_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtDateApplied.SetFocus
End If
End Sub

Private Sub txtDateApplied_GotFocus()
HTEXT txtDateApplied
End Sub

Private Sub txtDateApplied_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
    If TRANSACTIONTYPE = is_ADDING Or _
    TRANSACTIONTYPE = is_EDITTING Then
        lstDetail.SetFocus
    End If
End If
End Sub

Private Sub txtDateApplied_LostFocus()
If IsDate(txtDateApplied.Text) = True Then
    txtDateApplied.Text = Format(FormatDateTime(txtDateApplied.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub

Private Sub txtEmp_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(ROW).SubItems(2) = RETURNTEXTVALUE(txtEmp)
    End With
End If
End Sub

Private Sub txtHours_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(ROW).SubItems(7) = RETURNTEXTVALUE(txtHours)
    End With
End If
End Sub

Private Sub txtHours_GotFocus()
HTEXT txtHours
End Sub

Private Sub txtHours_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtMinutes.SetFocus
If KeyCode = vbKeyUp Then cmbType.SetFocus
End Sub

Private Sub txtHours_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtIDNo_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(ROW).SubItems(3) = Trim(txtIDNo.Text)
    End With
End If
End Sub

Private Sub txtMinutes_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(ROW).SubItems(8) = RETURNTEXTVALUE(txtMinutes)
    End With
End If
End Sub

Private Sub txtMinutes_GotFocus()
HTEXT txtMinutes
End Sub

Private Sub txtMinutes_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then txtHours.SetFocus
If KeyCode = vbKeyReturn Then
    Arr = Split(CHECK_ABSENT_DUPLICATE(RETURNTEXTVALUE(txtEmp), RETURNTEXTVALUE(txtType), ROW), "|", -1, 1)
    If CStr(Arr(0)) = "True" Then
        MsgBox "Found Duplicate Value in Line #" & CStr(Arr(1)) & "             ", vbCrLf, "Error..."
        Exit Sub
    End If
    picMain.Enabled = True
    picToolbar.Enabled = True
    picSLine.Visible = False
    lstDetail.SetFocus
End If
End Sub

Private Sub txtMinutes_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtName_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(ROW).SubItems(4) = Trim(txtName.Text)
    End With
End If
End Sub

Private Sub txtName_GotFocus()
txtNameFocus = 1
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbType.SetFocus
End Sub

Private Sub txtName_LostFocus()
txtNameFocus = 0
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstResult.Clear: lstCtrl.Clear: Exit Sub
lstResult.Clear: lstCtrl.Clear
s = "SELECT tbl_Personnel_AbsentLateUndertime_Details.EmployeeKey, tbl_Personnel_IDNumber.IDNumber, " & _
    " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName " & _
    " FROM tbl_Personnel_AbsentLateUndertime_Details LEFT OUTER JOIN " & _
    " tbl_Personnel_IDNumber ON tbl_Personnel_AbsentLateUndertime_Details.EmployeeKey = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
    " WHERE (tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
    " GROUP BY tbl_Personnel_AbsentLateUndertime_Details.EmployeeKey, tbl_Personnel_IDNumber.IDNumber, " & _
    " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName " & _
    " ORDER BY tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResult.AddItem rs!IDNumber & " - " & rs!EmployeeName
    lstResult.ItemData(lstResult.NewIndex) = rs!EmployeeKey
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

Private Sub txtSearchSLine_Change()
If Trim(txtSearchSLine.Text) = "" Then lstResultSLine.Clear: Exit Sub
If IsDate(txtDateApplied.Text) = False Then lstResultSLine.Clear: Exit Sub
If cmbDivision.ListIndex = -1 Then lstResultSLine.Clear: Exit Sub
lstResultSLine.Clear
's = "SELECT tbl_Personnel_IDNumber.PK, tbl_Personnel_IDNumber.IDNumber, " & _
    " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName " & _
    " FROM tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
    " WHERE (tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearchSLine.Text)) & "%') " & _
    " AND (ISNULL((SELECT TOP 1 tbl_Personnel_EmploymentStatus.Active " & _
    " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
    " tbl_Personnel_EmploymentStatus ON tbl_Personnel_Action.EmpStatus = tbl_Personnel_EmploymentStatus.PK " & _
    " WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_IDNumber.PK) " & _
    " AND (tbl_Personnel_Action.EffectivityDate <= '" & FormatDateTime(txtDateApplied.Text, vbShortDate) & "') " & _
    " ORDER BY tbl_Personnel_Action.EffectivityDate DESC), 0) = 1) " & _
    " ORDER BY tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName"
s = "SELECT dbo.tbl_Personnel_IDNumber.PK, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName, dbo.tbl_Personnel_Information.FirstName, dbo.tbl_Personnel_Information.MiddleName " & _
    " FROM  dbo.tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
    " WHERE (ISNULL((SELECT TOP (1) DivisionKey FROM  dbo.tbl_Personnel_ActionNew AS tbl_Personnel_ActionNew_1 " & _
    " WHERE (EmpPK = dbo.tbl_Personnel_IDNumber.PK) AND (EffectivityDate <= '" & FormatDateTime(txtDateApplied.Text, vbShortDate) & "') " & _
    " ORDER BY EffectivityDate DESC), 0) = " & cmbDivision.ItemData(cmbDivision.ListIndex) & ") " & _
    " AND (ISNULL((SELECT TOP (1) tbl_Personnel_EmploymentStatus_1.Active " & _
    " FROM  dbo.tbl_Personnel_ActionNew AS tbl_Personnel_ActionNew_2 LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_EmploymentStatus AS tbl_Personnel_EmploymentStatus_1 ON tbl_Personnel_ActionNew_2.EmpStatusKey = tbl_Personnel_EmploymentStatus_1.PK " & _
    " WHERE (tbl_Personnel_ActionNew_2.EmpPK = dbo.tbl_Personnel_IDNumber.PK) AND (tbl_Personnel_ActionNew_2.EffectivityDate <= '" & FormatDateTime(txtDateApplied.Text, vbShortDate) & "') " & _
    " ORDER BY tbl_Personnel_ActionNew_2.EffectivityDate DESC), 0) = 1) " & _
    " AND (dbo.tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearchSLine.Text)) & "%') " & _
    " ORDER BY tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultSLine.AddItem rs!IDNumber & " - " & rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
    lstResultSLine.ItemData(lstResultSLine.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResultSLine.ListCount Then lstResultSLine.ListIndex = 0
End Sub

Private Sub txtSearchSLine_GotFocus()
HTEXT txtSearchSLine
End Sub

Private Sub txtSearchSLine_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultSLine.SetFocus
End Sub

Private Sub txtType_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(ROW).SubItems(5) = RETURNTEXTVALUE(txtType)
    End With
End If
End Sub


