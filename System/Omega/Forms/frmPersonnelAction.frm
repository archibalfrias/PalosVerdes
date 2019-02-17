VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelAction 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPersonnelAction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picSearch 
      Height          =   4935
      Left            =   2640
      TabIndex        =   55
      Top             =   360
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   8705
      BackColor       =   15266266
      Begin VB.ComboBox cmbEffectivityDate 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   3960
         Width           =   2055
      End
      Begin VB.CommandButton cmdOKSearch 
         Height          =   480
         Left            =   120
         Picture         =   "frmPersonnelAction.frx":18B02
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   4320
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelSearch 
         Height          =   480
         Left            =   1920
         Picture         =   "frmPersonnelAction.frx":19174
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   4320
         Width           =   1560
      End
      Begin VB.ListBox lstResultSearch 
         Height          =   2985
         Left            =   120
         TabIndex        =   57
         Top             =   885
         Width           =   3375
      End
      Begin VB.TextBox txtSearchSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   56
         Top             =   480
         Width           =   3375
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   40
         TabIndex        =   60
         Top             =   40
         Width           =   3530
         _ExtentX        =   6218
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
         Icon            =   "frmPersonnelAction.frx":198D0
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Effectivity Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   3960
         Width           =   1335
      End
   End
   Begin RPVGCC.b8Container picAdd 
      Height          =   4935
      Left            =   2640
      TabIndex        =   49
      Top             =   360
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   8705
      BackColor       =   15266266
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   54
         Top             =   480
         Width           =   3375
      End
      Begin VB.ListBox lstResult 
         Height          =   3375
         Left            =   120
         TabIndex        =   53
         Top             =   890
         Width           =   3375
      End
      Begin VB.CommandButton cmdCancel 
         Height          =   480
         Left            =   1920
         Picture         =   "frmPersonnelAction.frx":19E6A
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   4320
         Width           =   1560
      End
      Begin VB.CommandButton cmdOK 
         Height          =   480
         Left            =   120
         Picture         =   "frmPersonnelAction.frx":1A5C6
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   4320
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   40
         TabIndex        =   50
         Top             =   40
         Width           =   3530
         _ExtentX        =   6218
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
         Icon            =   "frmPersonnelAction.frx":1AC38
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   240
      ScaleHeight     =   4335
      ScaleWidth      =   8175
      TabIndex        =   0
      Top             =   960
      Width           =   8175
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6720
         TabIndex        =   63
         Top             =   3300
         Width           =   1455
      End
      Begin VB.ComboBox cmbDivision 
         Height          =   315
         ItemData        =   "frmPersonnelAction.frx":1B1D2
         Left            =   1680
         List            =   "frmPersonnelAction.frx":1B1DC
         TabIndex        =   28
         Top             =   660
         Width           =   3255
      End
      Begin VB.TextBox txtRemarks 
         Height          =   315
         Left            =   1680
         TabIndex        =   27
         Top             =   3960
         Width           =   6495
      End
      Begin VB.TextBox txtControl 
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   26
         Top             =   0
         Width           =   1215
      End
      Begin VB.TextBox txtEffectDate 
         Height          =   315
         Left            =   6720
         TabIndex        =   25
         Top             =   2310
         Width           =   1455
      End
      Begin VB.ComboBox cmbTaxStatus 
         Height          =   315
         Left            =   1680
         TabIndex        =   24
         Top             =   1650
         Width           =   3255
      End
      Begin VB.PictureBox picGovt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F1DA&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   6720
         ScaleHeight     =   1185
         ScaleWidth      =   1425
         TabIndex        =   15
         Top             =   960
         Width           =   1455
         Begin VB.CheckBox chkTIN 
            BackColor       =   &H00FDE9C6&
            Height          =   200
            Left            =   120
            TabIndex        =   19
            Top             =   840
            Width           =   200
         End
         Begin VB.CheckBox chkPagIbig 
            BackColor       =   &H00FDE9C6&
            Height          =   200
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   200
         End
         Begin VB.CheckBox chkPHIC 
            BackColor       =   &H00FDE9C6&
            Height          =   200
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   200
         End
         Begin VB.CheckBox chkSSS 
            BackColor       =   &H00FDE9C6&
            Height          =   200
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   200
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "SSS"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   360
            TabIndex        =   23
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "PHIL HEALTH"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   360
            TabIndex        =   22
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "PAG IBIG"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "TIN"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   360
            TabIndex        =   20
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.ComboBox cmbComp 
         Height          =   315
         ItemData        =   "frmPersonnelAction.frx":1B1F9
         Left            =   1680
         List            =   "frmPersonnelAction.frx":1B1FB
         TabIndex        =   14
         Top             =   2310
         Width           =   3255
      End
      Begin VB.TextBox txtCola 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6720
         TabIndex        =   13
         Top             =   2970
         Width           =   1455
      End
      Begin VB.TextBox txtBasic 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6720
         TabIndex        =   12
         Top             =   2640
         Width           =   1455
      End
      Begin VB.ComboBox cmbPost 
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   1980
         Width           =   3255
      End
      Begin VB.ComboBox cmdStatus 
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox txtID 
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Top             =   330
         Width           =   1215
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   2910
         TabIndex        =   8
         Top             =   330
         Width           =   5260
      End
      Begin VB.TextBox txtTIN 
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   3630
         Width           =   3255
      End
      Begin VB.TextBox txtPagIbig 
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   3300
         Width           =   3255
      End
      Begin VB.TextBox txtPHIC 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   2970
         Width           =   3255
      End
      Begin VB.TextBox txtSSS 
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   2640
         Width           =   3255
      End
      Begin VB.ComboBox cmdDept 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   990
         Width           =   3255
      End
      Begin VB.TextBox txtAllow 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6720
         TabIndex        =   2
         Top             =   3630
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   6240
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox picBasic 
         BackColor       =   &H00000000&
         Height          =   315
         Left            =   6720
         ScaleHeight     =   255
         ScaleWidth      =   1395
         TabIndex        =   65
         Top             =   2640
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox picCola 
         BackColor       =   &H00000000&
         Height          =   315
         Left            =   6720
         ScaleHeight     =   255
         ScaleWidth      =   1395
         TabIndex        =   66
         Top             =   2970
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox picTotal 
         BackColor       =   &H00000000&
         Height          =   315
         Left            =   6720
         ScaleHeight     =   255
         ScaleWidth      =   1395
         TabIndex        =   67
         Top             =   3300
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox picAllowance 
         BackColor       =   &H00000000&
         Height          =   315
         Left            =   6720
         ScaleHeight     =   255
         ScaleWidth      =   1395
         TabIndex        =   68
         Top             =   3630
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5160
         TabIndex        =   64
         Top             =   3375
         Width           =   975
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "DIVISION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   45
         Top             =   705
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "REMARKS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   44
         Top             =   4005
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "CTRL NO."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "EFFECTIVITY DATE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5160
         TabIndex        =   42
         Top             =   2355
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "TAX STATUS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   41
         Top             =   1710
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "RATE COMPENSATION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   40
         Top             =   2355
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "COLA"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5160
         TabIndex        =   39
         Top             =   3030
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "BASIC"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5160
         TabIndex        =   38
         Top             =   2685
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NAME"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   37
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "TIN"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   36
         Top             =   3705
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "PAG IBIG"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   35
         Top             =   3375
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "PHIL HEALTH "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   34
         Top             =   3030
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "SSS NO."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   2685
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "EMP STATUS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   32
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "DEPARTMENT"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   1035
         Width           =   1095
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "POSITION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   30
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "ALLOWANCE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5160
         TabIndex        =   29
         Top             =   3670
         Width           =   975
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   770
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   15000
      TabIndex        =   46
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   570
         Left            =   0
         TabIndex        =   47
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
            NumButtons      =   16
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Add"
               Key             =   "Add"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Edit"
               Key             =   "Edit"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Delete"
               Key             =   "Delete"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Save"
               Key             =   "Save"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Undo"
               Key             =   "Undo"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Find"
               Key             =   "Find"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Print"
               Key             =   "Print"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
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
         Y1              =   690
         Y2              =   690
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7920
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483648
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAction.frx":1B1FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAction.frx":1B2FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAction.frx":1B483
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAction.frx":1B79D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAction.frx":1B8AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAction.frx":1BDF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAction.frx":1BF4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelAction.frx":1C48D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   48
      Top             =   5475
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10936
            MinWidth        =   10936
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPersonnelAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public locEmployeePK As Long
Public locTransNo As Long

Dim locDept As Long
Dim locPost As Long
Dim locEmpStatus As Long
Dim locTaxStatus As Long
Public iSupervisory As Long

Dim tmp As Long

Public TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2
Const is_FINDING = 3

Dim Array1, Array2, Array3, Array4, strEmpNo, dblRatePerHour, _
dblAllowanceRate, dblColaPerHour, strCtrl, strDeptFrom, strPostFrom, _
strStatusFrom, strTaxStatus, dblBasic, dblAllowance, Compensation

Private Function BROWSER(strCtrl, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If Trim(strCtrl) <> "" Then
            s = "sp_Personnel_Action_Browse('" & strCtrl & "',0) "
        Else
            s = "sp_Personnel_Action_Browse('" & strCtrl & "',1) "
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "sp_Personnel_Action_Browse('" & strCtrl & "',1) "
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "sp_Personnel_Action_Browse('" & strCtrl & "',2) "
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "sp_Personnel_Action_Browse('" & strCtrl & "',3) "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "sp_Personnel_Action_Browse('" & strCtrl & "',4) "
    Case Else: Exit Function
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iSupervisory = rs!PositionLevel
    locTransNo = rs!PK
    locEmployeePK = rs!EmpPK
    locDept = rs!Dept
    locPost = rs!Positions
    locEmpStatus = rs!EmpStatus
    locTaxStatus = rs!TaxStatus
    
    txtControl.Text = rs!CntrlNo
    txtID.Text = rs!IDNumber
    txtName.Text = rs!EmployeeName
    cmbDivision.ListIndex = rs!Division - 1
    If DEPT_NAME(rs!Dept) <> "" Then
        Array1 = Split(DEPT_NAME(rs!Dept), ";", -1, 1)
        cmdDept.Text = CStr(Array1(1))
    Else
        cmdDept.Text = ""
    End If
    If EMP_STATUS(rs!EmpStatus) <> "" Then
        Array2 = Split(EMP_STATUS(rs!EmpStatus), ";", -1, 1)
        cmdStatus = CStr(Array2(1))
    Else
        cmdStatus.Text = ""
    End If
    If TAX_STATUS_NAME(rs!TaxStatus) <> "" Then
        Array4 = Split(TAX_STATUS_NAME(rs!TaxStatus), ";", -1, 1)
        cmbTaxStatus.Text = CStr(Array4(1))
    Else
        cmbTaxStatus.Text = ""
    End If
    If POSITION_NAME(rs!Positions) <> "" Then
        Array3 = Split(POSITION_NAME(rs!Positions), ";", -1, 1)
        cmbPost.Text = CStr(Array3(1))
    Else
        cmbPost.Text = ""
    End If
    cmbComp.ListIndex = rs!CompensationRate - 1
    chkSSS.Value = rs!Is_SSS
    chkPHIC.Value = rs!Is_PHIC
    chkPagIbig.Value = rs!Is_PAGIBIG
    chkTIN.Value = rs!Is_TIN
    txtSSS.Text = rs!SSS
    txtPHIC.Text = rs!PHIC
    txtPagIbig.Text = rs!PAGIBIG
    txtTIN.Text = rs!TIN
    txtRemarks.Text = rs!Remarks
    txtEffectDate.Text = Format(rs!EffectivityDate, "mm/dd/yyyy")
    
    HIDE_SALARY_RATE iSupervisory
    
'    If iSupervisory = 2 Then
'        If AccessRights("Personnel Action Memo", "Supervisory") = True Then
'            picBasic.Visible = False
'            picCola.Visible = False
'            picTotal.Visible = False
'            picAllowance.Visible = False
''            txtBasic.BackColor = &H80000005
''            txtCola.BackColor = &H80000005
''            txtTotal.BackColor = &H80000005
''            txtAllow.BackColor = &H80000005
'        Else
'            picBasic.ZOrder 0
'            picCola.ZOrder 0
'            picTotal.ZOrder 0
'            picAllowance.ZOrder 0
'            picBasic.Visible = True
'            picCola.Visible = True
'            picTotal.Visible = True
'            picAllowance.Visible = True
''            txtBasic.BackColor = &H0&
''            txtCola.BackColor = &H0&
''            txtTotal.BackColor = &H0&
''            txtAllow.BackColor = &H0&
'        End If
'    Else
'        picBasic.ZOrder 0
'        picCola.ZOrder 0
'        picTotal.ZOrder 0
'        picAllowance.ZOrder 0
'        picBasic.Visible = True
'        picCola.Visible = True
'        picTotal.Visible = True
'        picAllowance.Visible = True
''        txtBasic.BackColor = &H80000005
''        txtCola.BackColor = &H80000005
''        txtTotal.BackColor = &H80000005
''        txtAllow.BackColor = &H80000005
'    End If
    
    txtBasic.Text = Format(rs!Basic, "#,##0.00")
    txtCola.Text = Format(rs!Cola, "#,##0.00")
    txtAllow.Text = Format(rs!Allowance, "#,##0.00")
    
    StatusBar.Panels(1).Text = rs!PK
    StatusBar.Panels(2).Text = IIf(IsNull(rs!LastModified), "", "LAST MODIFIED BY : " & rs!LastModified)
    StatusBar.Panels(3).Text = IIf(rs!Locked = 0, "UNLOCKED", "LOCKED")
    SaveSetting App.EXEName, "PersonnelActionCtrl", "PerActCtrl", rs!CntrlNo
    
End If
rs.Close
End Function

Public Sub HIDE_SALARY_RATE(iLevel As Long)
    If iLevel = 2 Then
        If AccessRights("Personnel Action Memo", "Supervisory") = True Then
            picBasic.Visible = False
            picCola.Visible = False
            picTotal.Visible = False
            picAllowance.Visible = False
'            txtBasic.BackColor = &H80000005
'            txtCola.BackColor = &H80000005
'            txtTotal.BackColor = &H80000005
'            txtAllow.BackColor = &H80000005
        Else
            picBasic.ZOrder 0
            picCola.ZOrder 0
            picTotal.ZOrder 0
            picAllowance.ZOrder 0
            picBasic.Visible = True
            picCola.Visible = True
            picTotal.Visible = True
            picAllowance.Visible = True
'            txtBasic.BackColor = &H0&
'            txtCola.BackColor = &H0&
'            txtTotal.BackColor = &H0&
'            txtAllow.BackColor = &H0&
        End If
    Else
'        picBasic.ZOrder 0
'        picCola.ZOrder 0
'        picTotal.ZOrder 0
'        picAllowance.ZOrder 0
        picBasic.Visible = False
        picCola.Visible = False
        picTotal.Visible = False
        picAllowance.Visible = False
'        picBasic.Visible = True
'        picCola.Visible = True
'        picTotal.Visible = True
'        picAllowance.Visible = True
'        txtBasic.BackColor = &H80000005
'        txtCola.BackColor = &H80000005
'        txtTotal.BackColor = &H80000005
'        txtAllow.BackColor = &H80000005
    End If
End Sub


Private Function PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If picAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
If AccessRights("Personnel Action Memo", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If

PopupMenu MainFormPopupF.mnuActionMemo, , Toolbar1.Buttons(1).Left, 500

'picToolbar.Enabled = False
'picMain.Enabled = False
'picAdd.ZOrder 0
'txtSearch.Text = ""
'picAdd.Visible = True
'txtSearch.SetFocus

End Function

Private Function PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If StatusBar.Panels(1).Text = "" Then Exit Function
If picAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
If AccessRights("Personnel Action Memo", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If

If StatusBar.Panels(3).Text = "LOCKED" Then MsgBox "TRANSACTION ALREADY LOCKED!                         ", vbCritical, "Error...": Exit Function

If iSupervisory = 2 Then
    If AccessRights("Personnel Action Memo", "Supervisory") = False Then
        MsgBox "CAN'T EDIT SUPERVISORY TRANSCTION.                  " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Function
    End If
End If

s = "SELECT ActionMemo " & _
    " From tbl_Personnel_Compensation " & _
    " WHERE (ActionMemo = " & StatusBar.Panels(1).Text & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    MsgBox "CANNOT BE EDITTED!          " & vbCrLf & _
           "" & vbCrLf & _
           "Action used by the Payroll Transaction.   " & vbCrLf & _
           "Any changes have an effect on Payroll Computation.  ", vbCritical, "Error..."
    Exit Function
Else
    'Me.Caption = "Personnel Action Memo - Edit"
    TRANSACTIONTYPE = is_EDITTING
    LOCKTEXT False
    'TOOLBARBUTTON False
    TOOLBAR_FUNC 2
    cmbDivision.SetFocus
End If
rs.Close
End Function


Private Function PRESS_DELETE()

If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If StatusBar.Panels(1).Text = "" Then Exit Function
If picAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function

If AccessRights("Personnel Action Memo", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If

If StatusBar.Panels(3).Text = "LOCKED" Then MsgBox "TRANSACTION ALREADY LOCKED!                         ", vbCritical, "Error...": Exit Function

If iSupervisory = 2 Then
    If AccessRights("Personnel Action Memo", "Supervisory") = False Then
        MsgBox "CAN'T DELETE SUPERVISORY TRANSCTION.                " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Function
    End If
End If

s = "SELECT ActionMemo " & _
    " From tbl_Personnel_Compensation " & _
    " WHERE (ActionMemo = " & StatusBar.Panels(1).Text & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    MsgBox "CANNOT BE DELETED!          " & vbCrLf & _
           "" & vbCrLf & _
           "Action used by the Payroll Transaction.   " & vbCrLf & _
           "Any changes have an effect on Payroll Computation.  ", vbCritical, "Error..."
    Exit Function
Else
    If MsgBox("ARE YOU SURE TO DELETE THIS RECORD?          ", vbCritical + vbYesNo, "CONFIRMATION") = vbNo Then Exit Function
    
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Action " & _
                      " WHERE (PK = " & StatusBar.Panels(1).Text & ")"
    
    strEmpNo = locEmployeePK
        
    CLEARTEXT
    
End If
rs.Close


Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error...."
Exit Function
End Function


Private Function PRESS_F5()
If picAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
If locEmployeePK = 0 Then MsgBox "Please Select Employee!                   ", vbCritical, "Error...": Exit Function
If cmbDivision.ListIndex = -1 Then MsgBox "Please Select Division!              ", vbCritical, "Error...": cmbDivision.SetFocus: Exit Function
If locDept = 0 Then MsgBox "Please Supply Department!                   ", vbCritical, "Error...": Exit Function
If locPost = 0 Then MsgBox "Please Supply Position!                 ", vbCritical, "Error...": Exit Function
If locEmpStatus = 0 Then MsgBox "Please Supply Employment Status!               ", vbCritical, "Error...": Exit Function
If locTaxStatus = 0 Then MsgBox "Please Supply Tax Status!                  ", vbCritical, "Error...": Exit Function
If cmbComp.ListIndex = -1 Then MsgBox "Please Supply Compensation Rate!                 ", vbCritical, "Error...": Exit Function
If IsDate(txtEffectDate.Text) = False Then MsgBox "Please Supply a Valid Effect Date!                 ", vbCritical, "Error...": txtEffectDate.SetFocus: Exit Function
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    If DateValue(GET_LAST_ACTION_EFFECTIVITY(locEmployeePK)) < DateValue(CDate(Trim(txtEffectDate.Text))) Then
        
        strCtrl = ""
        s = "SELECT TOP 1 CntrlNo" & _
            " From tbl_Personnel_Action " & _
            " ORDER BY CntrlNo DESC"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            strCtrl = Format(CLng(rs!CntrlNo) + 1, "0000000#")
        Else
            strCtrl = Format(1, "0000000#")
        End If
        rs.Close
        
        Do
            s = "SELECT tbl_Personnel_Action.* " & _
                " From tbl_Personnel_Action " & _
                " WHERE (CntrlNo = '" & strCtrl & "')"
            If rs.State = adStateOpen Then rs.Close
            rs.Open s, ConnOmega
            If rs.RecordCount = 0 Then
                rs.Close
                Exit Do
            End If
            rs.Close
            strCtrl = Format(CDbl(strCtrl) + 1, "0000000#")
        Loop
        
        dblRatePerHour = 0: dblAllowanceRate = 0: dblColaPerHour = 0
        If CInt(cmbComp.ListIndex + 1) = 1 Then
            dblRatePerHour = ((CDbl(RETURNTEXTVALUE(txtBasic)) / 2) / 13.08333) / 8
            dblAllowanceRate = ((CDbl(RETURNTEXTVALUE(txtAllow)) / 2) / 13.08333) / 8
            'dblColaPerHour = ((CDbl(RETURNTEXTVALUE(txtCola)) / 2) / 13.08333) / 8
        ElseIf CInt(cmbComp.ListIndex + 1) = 2 Then
            dblRatePerHour = CDbl(RETURNTEXTVALUE(txtBasic)) / 8
            dblAllowanceRate = CDbl(RETURNTEXTVALUE(txtAllow)) / 8
            'dblColaPerHour = CDbl(RETURNTEXTVALUE(txtCola)) / 8
        End If
        
        dblColaPerHour = CDbl(RETURNTEXTVALUE(txtCola)) / 8
        
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Action" & _
                          " (CntrlNo, EmpPK, Division, Dept, EmpStatus, TaxStatus," & _
                          " Positions, CompensationRate, Is_PAGIBIG, Is_PHIC, Is_SSS, " & _
                          " Is_TIN, SSS, PAGIBIG, PHIC, TIN, Remarks, EffectivityDate, " & _
                          " Basic, RatePerHourBasic, Allowance, RatePerHourAllow, LastModified, " & _
                          " Cola, RatePerHourCola)" & _
                          " VALUES ('" & strCtrl & "', " & locEmployeePK & ", " & cmbDivision.ListIndex + 1 & ", " & _
                          " " & locDept & ", " & locEmpStatus & ", " & locTaxStatus & ", " & locPost & ", " & _
                          " " & cmbComp.ListIndex + 1 & ", " & chkPagIbig.Value & ", " & chkPHIC.Value & ", " & chkSSS.Value & ", " & _
                          " " & chkTIN.Value & ", '" & Trim(txtSSS.Text) & "', '" & Trim(txtPagIbig.Text) & "', '" & Trim(txtPHIC.Text) & "', " & _
                          " '" & Trim(txtTIN.Text) & "', '" & Trim(txtRemarks.Text) & "', '" & FormatDateTime(txtEffectDate.Text, vbShortDate) & "', " & _
                          " " & RETURNTEXTVALUE(txtBasic) & ", " & CDbl(dblRatePerHour) & "," & RETURNTEXTVALUE(txtAllow) & ", " & _
                          " " & CDbl(dblAllowanceRate) & ",'" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                          " " & RETURNTEXTVALUE(txtCola) & ", " & CDbl(dblColaPerHour) & ")"
        s = "SELECT PK " & _
            " From tbl_Personnel_Action " & _
            " WHERE (EmpPK = " & locEmployeePK & ") " & _
            " AND (EffectivityDate = '" & FormatDateTime(txtEffectDate.Text, vbShortDate) & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            t = "SELECT ProfileKey " & _
                " FROM tbl_Personnel_IDNumber " & _
                " WHERE (PK = " & locEmployeePK & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                ConnOmega.Execute "UPDATE tbl_Personnel_Information " & _
                                  " SET SSSNumber = '" & Trim(txtSSS.Text) & "', " & _
                                  " PHICNumber = '" & Trim(txtPHIC.Text) & "', " & _
                                  " HDMFNumber = '" & Trim(txtPagIbig.Text) & "', " & _
                                  " TIN = '" & Trim(txtTIN.Text) & "', " & _
                                  " TaxStatus = " & locTaxStatus & " " & _
                                  " WHERE (PK = " & rt!ProfileKey & ")"
            End If
            rt.Close
        End If
        rs.Close
                
        LOCKTEXT True
        'TOOLBARBUTTON True
        TOOLBAR_FUNC 1
        TRANSACTIONTYPE = is_REFRESH
        BROWSER strCtrl, "is_LOAD"
        'Me.Caption = "Personnel Action Memo"
    Else
        MsgBox "EFFECTIVITY DATE MUST BE HIGHER THAN THE LAST ACTION MEMO!          ", vbInformation, "Error..."
    End If
ElseIf TRANSACTIONTYPE = is_EDITTING Then
    
    dblRatePerHour = 0: dblAllowanceRate = 0: dblColaPerHour = 0
    If CInt(cmbComp.ListIndex + 1) = 1 Then
        dblRatePerHour = ((CDbl(RETURNTEXTVALUE(txtBasic)) / 2) / 13.08333) / 8
        dblAllowanceRate = ((CDbl(RETURNTEXTVALUE(txtAllow)) / 2) / 13.08333) / 8
        'dblColaPerHour = ((CDbl(RETURNTEXTVALUE(txtCola)) / 2) / 13.08333) / 8
    ElseIf CInt(cmbComp.ListIndex + 1) = 2 Then
        dblRatePerHour = CDbl(RETURNTEXTVALUE(txtBasic)) / 8
        dblAllowanceRate = CDbl(RETURNTEXTVALUE(txtAllow)) / 8
        'dblColaPerHour = CDbl(RETURNTEXTVALUE(txtCola)) / 8
    End If
    
    dblColaPerHour = CDbl(RETURNTEXTVALUE(txtCola)) / 8
    
    ConnOmega.Execute "UPDATE tbl_Personnel_Action" & _
                      " SET Division = " & cmbDivision.ListIndex + 1 & ", Dept = " & locDept & ", " & _
                      " EmpStatus = " & locEmpStatus & ", TaxStatus = " & locTaxStatus & "," & _
                      " Positions =  " & locPost & ", CompensationRate = " & cmbComp.ListIndex + 1 & ", " & _
                      " Is_PAGIBIG = " & chkPagIbig.Value & ", Is_PHIC = " & chkPHIC.Value & ", " & _
                      " Is_SSS = " & chkSSS.Value & ", Is_TIN = " & chkTIN.Value & ", " & _
                      " SSS = '" & Trim(txtSSS.Text) & "', PAGIBIG = '" & Trim(txtPagIbig.Text) & "', " & _
                      " PHIC = '" & Trim(txtPHIC.Text) & "', TIN = '" & Trim(txtTIN.Text) & "', " & _
                      " Remarks = '" & Trim(txtRemarks.Text) & "', EffectivityDate = '" & FormatDateTime(txtEffectDate.Text, vbShortDate) & "', " & _
                      " Basic = " & RETURNTEXTVALUE(txtBasic) & ", RatePerHourBasic = " & CDbl(dblRatePerHour) & ",  " & _
                      " Allowance = " & RETURNTEXTVALUE(txtAllow) & ", RatePerHourAllow = " & CDbl(dblAllowanceRate) & ", " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                      " Cola = " & RETURNTEXTVALUE(txtCola) & ", RatePerHourCola = " & CDbl(dblColaPerHour) & " " & _
                      " WHERE (PK = " & StatusBar.Panels(1).Text & ")"

                
    t = "SELECT ProfileKey " & _
        " FROM tbl_Personnel_IDNumber " & _
        " WHERE (PK = " & locEmployeePK & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        ConnOmega.Execute "UPDATE tbl_Personnel_Information " & _
                          " SET SSSNumber = '" & Trim(txtSSS.Text) & "', " & _
                          " PHICNumber = '" & Trim(txtPHIC.Text) & "', " & _
                          " HDMFNumber = '" & Trim(txtPagIbig.Text) & "', " & _
                          " TIN = '" & Trim(txtTIN.Text) & "', " & _
                          " TaxStatus = " & locTaxStatus & " " & _
                          " WHERE (PK = " & rt!ProfileKey & ")"
    End If
    rt.Close
            
    LOCKTEXT True
    'TOOLBARBUTTON True
    TOOLBAR_FUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER Trim(txtControl.Text), "is_LOAD"
    'Me.Caption = "Personnel Action Memo"
End If
Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error.."
Exit Function
End Function

Private Function PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If picAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
picToolbar.Enabled = False
picMain.Enabled = False
picSearch.ZOrder 0
txtSearchSearch.Text = ""
picSearch.Visible = True
txtSearchSearch.SetFocus
End Function


Private Function PRESS_F9()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If StatusBar.Panels(1).Text = "" Then Exit Function
If picAdd.Visible = True Then Exit Function
If picSearch.Visible = True Then Exit Function
If AccessRights("Personnel Action Memo", "Print") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
If iSupervisory = 2 Then
    If AccessRights("Personnel Action Memo", "Supervisory") = False Then
        MsgBox "THIS TRANSACTION IS FOR SUPERVISORY ACCOUNT ONLY.   " & vbCrLf & vbCrLf & _
               "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Function
    End If
End If

s = "sp_Personnel_Action_Print(" & locEmployeePK & ", '" & FormatDateTime(txtEffectDate.Text, vbShortDate) & "', 0)"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    
    ConnOmega.Execute "DELETE FROM tbl_PersonnelAction_Tmp WHERE (LogInName = '" & gbl_UserName & "')"
    
    If rs!CompensationRate = 1 Then
        Compensation = "MONTHLY"
    ElseIf rs!CompensationRate = 2 Then
        Compensation = "DAILY"
    End If
    
    ConnOmega.Execute "INSERT INTO  tbl_PersonnelAction_Tmp " & _
                      " (CrtlNo, IDNo, Name, DateHired, EffectDate, SSS, " & _
                      " PHIC, PagIbig, TIN, Remarks, StatusTo, " & _
                      " PostTo, DeptTo, CompTo, ColaTo, " & _
                      " BasicTo, AllowanceTo, LogInName, Company) " & _
                      " VALUES('" & rs!CntrlNo & "', '" & rs!IDNumber & "', '" & rs!EmployeeName & "', " & _
                      " '" & CStr(UCase(Format(rs!DHired, "mmmm dd, yyyy"))) & "', '" & CStr(UCase(Format(rs!EffectivityDate, "mmmm dd, yyyy"))) & "', " & _
                      " '" & rs!SSS & "', " & _
                      " '" & rs!PHIC & "', '" & rs!PAGIBIG & "', '" & rs!TIN & "', " & _
                      " '" & FORMATSQL(rs!Remarks) & "', '" & rs!StatusName & "', " & _
                      " '" & rs!PositionName & "', '" & rs!DepartmentName & "', " & _
                      " '" & Compensation & "', " & CDbl(rs!Cola) & ", " & _
                      " " & CDbl(rs!Basic) & ", " & _
                      " " & CDbl(rs!Allowance) & ", '" & gbl_UserName & "', 1)"

    t = "sp_Personnel_Action_Print(" & locEmployeePK & ", '" & FormatDateTime(rs!EffectivityDate, vbShortDate) & "',1)"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        If rt!CompensationRate = 1 Then
            Compensation = "MONTHLY"
        ElseIf rt!CompensationRate = 2 Then
            Compensation = "DAILY"
        End If
        
        ConnOmega.Execute "UPDATE tbl_PersonnelAction_Tmp " & _
                          " SET StatusFrom = '" & rt!StatusName & "', " & _
                          " PostFrom = '" & rt!PositionName & "', " & _
                          " DeptFrom = '" & rt!DepartmentName & "', " & _
                          " CompFrom = '" & Compensation & "', " & _
                          " ColaFrom = " & CDbl(rs!Cola) & ", " & _
                          " BasicFrom = " & CDbl(rt!Basic) & ", " & _
                          " AllowanceFrom = " & CDbl(rt!Allowance) & "" & _
                          " WHERE (CrtlNo = '" & rs!CntrlNo & "') " & _
                          " AND (LogInName ='" & gbl_UserName & "')"
        
    Else
    
        ConnOmega.Execute "UPDATE tbl_PersonnelAction_Tmp " & _
                          " SET StatusFrom = '', " & _
                          " PostFrom = '', " & _
                          " DeptFrom = '', " & _
                          " CompFrom = '', " & _
                          " ColaFrom = 0, " & _
                          " BasicFrom = 0, " & _
                          " AllowanceFrom = 0 " & _
                          " WHERE (CrtlNo = '" & rs!CntrlNo & "')" & _
                          " AND (LogInName ='" & gbl_UserName & "')"
    End If
    rt.Close
    
    
    
End If
rs.Close

s = "SELECT tbl_PersonnelAction_Tmp.*" & _
    " FROM tbl_PersonnelAction_Tmp"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
rs.Requery
rs.Close

frmCrystalReportViewer.PRINT_ACTION_MEMO
If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show

End Function

Private Function PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picAdd.Visible = True Then cmdCancel_Click: Exit Function
    If picSearch.Visible = True Then cmdCancelSearch_Click: Exit Function
    Unload Me
Else
    CLEARTEXT
    TRANSACTIONTYPE = is_REFRESH
    LOCKTEXT True
    'TOOLBARBUTTON True
    If Trim(StatusBar.Panels(1).Text) <> "" Then
        TOOLBAR_FUNC 1
        BROWSER GetSetting(App.EXEName, "PersonnelActionCtrl", "PerActCtrl", ""), "is_LOAD"
    Else
        TOOLBAR_FUNC 0
    End If
    'Me.Caption = "Personnel Action Memo"
End If
End Function

Public Function CLEARTEXT()
iSupervisory = 0
locEmployeePK = 0
locDept = 0
locPost = 0
locEmpStatus = 0
locTaxStatus = 0
txtControl.Text = ""
txtID.Text = ""
txtName.Text = ""
cmbDivision.Text = ""
cmbDivision.ListIndex = -1
cmdDept.Text = ""
cmdDept.ListIndex = -1
cmdStatus.Text = ""
cmdStatus.ListIndex = -1
cmbTaxStatus.Text = ""
cmbTaxStatus.ListIndex = -1
cmbPost.Text = ""
cmbPost.ListIndex = -1
cmbComp.Text = ""
cmbComp.ListIndex = -1
txtSSS.Text = ""
txtPHIC.Text = ""
txtPagIbig.Text = ""
txtTIN.Text = ""
txtEffectDate.Text = ""
txtBasic.Text = ""
txtAllow.Text = ""
txtTotal.Text = ""
txtRemarks.Text = ""
txtCola.Text = ""
chkSSS.Value = 0
chkPHIC.Value = 0
chkPagIbig.Value = 0
chkTIN.Value = 0
StatusBar.Panels(1).Text = ""
StatusBar.Panels(2).Text = ""
StatusBar.Panels(3).Text = ""
txtBasic.BackColor = &H80000005
txtCola.BackColor = &H80000005
txtTotal.BackColor = &H80000005
txtAllow.BackColor = &H80000005
End Function

Public Function LOCKTEXT(bln As Boolean)
txtControl.Locked = True
txtID.Locked = True
txtName.Locked = True
txtTotal.Locked = True
'If bln Then
    cmbDivision.Locked = bln
    cmdDept.Locked = bln
    cmdStatus.Locked = bln
    cmbTaxStatus.Locked = bln
    cmbPost.Locked = bln
    cmbComp.Locked = bln
    txtSSS.Locked = bln
    txtPHIC.Locked = bln
    txtPagIbig.Locked = bln
    txtTIN.Locked = bln
    txtRemarks.Locked = bln
    txtEffectDate.Locked = bln
    txtBasic.Locked = bln
    txtAllow.Locked = bln
    txtCola.Locked = bln
    txtTotal.Locked = bln
    picGovt.Enabled = IIf(bln = True, False, True)
'Else
'    cmbDivision.Locked = False
'    cmdDept.Locked = False
'    cmdStatus.Locked = False
'    cmbTaxStatus.Locked = False
'    cmbPost.Locked = False
'    cmbComp.Locked = False
'    txtSSS.Locked = False
'    txtPHIC.Locked = False
'    txtPagIbig.Locked = False
'    txtTIN.Locked = False
'    txtRemarks.Locked = False
'    txtEffectDate.Locked = False
'    txtBasic.Locked = False
'    txtAllow.Locked = False
'    txtCola.Locked = False
'    txtTotal.Locked = True
'    picGovt.Enabled = True
'End If
End Function

Private Function TOOLBAR_FUNC(isSelect As Integer)
With Toolbar1
    Set .ImageList = ImageList1
    .Buttons(1).Image = 1
    .Buttons(3).Image = 2
    .Buttons(5).Image = 3
    .Buttons(7).Image = 7
    .Buttons(9).Image = 8
    .Buttons(11).Image = 4
    .Buttons(13).Image = 5
    .Buttons(15).Image = 6
    Select Case isSelect
        Case 0  'Empty Fields
            .Buttons(1).Enabled = True
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = False
            .Buttons(7).Enabled = False
            .Buttons(9).Enabled = False
            .Buttons(11).Enabled = True
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = True
            .Buttons(1).ToolTipText = "NEW (Ins)"
            .Buttons(3).ToolTipText = ""
            .Buttons(5).ToolTipText = ""
            .Buttons(7).ToolTipText = ""
            .Buttons(9).ToolTipText = ""
            .Buttons(11).ToolTipText = "FIND (F6)"
            .Buttons(13).ToolTipText = ""
            .Buttons(15).ToolTipText = "CLOSE (Esc)"
        Case 1
            .Buttons(1).Enabled = True
            .Buttons(3).Enabled = True
            .Buttons(5).Enabled = True
            .Buttons(7).Enabled = False
            .Buttons(9).Enabled = False
            .Buttons(11).Enabled = True
            .Buttons(13).Enabled = True
            .Buttons(15).Enabled = True
            .Buttons(1).ToolTipText = "NEW (Ins)"
            .Buttons(3).ToolTipText = "EDIT (F2)"
            .Buttons(5).ToolTipText = "DELETE (Del)"
            .Buttons(7).ToolTipText = ""
            .Buttons(9).ToolTipText = ""
            .Buttons(11).ToolTipText = "FIND (F6)"
            .Buttons(13).ToolTipText = "PRINT (F9)"
            .Buttons(15).ToolTipText = "CLOSE (Esc)"
        Case 2
            .Buttons(1).Enabled = False
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = False
            .Buttons(7).Enabled = True
            .Buttons(9).Enabled = True
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(1).ToolTipText = ""
            .Buttons(3).ToolTipText = ""
            .Buttons(5).ToolTipText = ""
            .Buttons(7).ToolTipText = "Save (F5)"
            .Buttons(9).ToolTipText = "Undo (Esc)"
            .Buttons(11).ToolTipText = ""
            .Buttons(13).ToolTipText = ""
            .Buttons(15).ToolTipText = ""
    End Select
End With
End Function

'Public Sub TOOLBARBUTTON(blnTag As Boolean)
'With Toolbar1
'    If blnTag Then
'        .Buttons(1).Image = 1
'        .Buttons(3).Image = 2
'        .Buttons(5).Image = 3
'        .Buttons(11).Image = 6
'        .Buttons(13).Image = 7
'        .Buttons(15).Image = 8
'        .Buttons(17).Image = 9
'        .Buttons(19).Image = 10
'        .Buttons(1).Enabled = True
'        .Buttons(3).Enabled = True
'        .Buttons(5).Enabled = True
'        .Buttons(7).Image = 4
'        .Buttons(7).Caption = "First"
'        .Buttons(9).Image = 5
'        .Buttons(9).Caption = "Back"
'        .Buttons(7).Enabled = True
'        .Buttons(9).Enabled = True
'        .Buttons(11).Enabled = True
'        .Buttons(13).Enabled = True
'        .Buttons(15).Enabled = True
'        .Buttons(17).Enabled = True
'        .Buttons(19).Enabled = True
'        .Buttons(1).ToolTipText = "NEW (Ins)"
'        .Buttons(3).ToolTipText = "EDIT (F2)"
'        .Buttons(5).ToolTipText = "DELETE (Del)"
'        .Buttons(7).ToolTipText = "FIRST (Home)"
'        .Buttons(9).ToolTipText = "BACK (PgUp)"
'        .Buttons(11).ToolTipText = "NEXT (PgDown)"
'        .Buttons(13).ToolTipText = "LAST (End)"
'        .Buttons(15).ToolTipText = "FIND (F6)"
'        .Buttons(17).ToolTipText = "PRINT (F9)"
'        .Buttons(19).ToolTipText = "CLOSE (Esc)"
'    Else
'        .Buttons(1).Image = 1
'        .Buttons(3).Image = 2
'        .Buttons(5).Image = 3
'        .Buttons(11).Image = 6
'        .Buttons(13).Image = 7
'        .Buttons(15).Image = 8
'        .Buttons(17).Image = 9
'        .Buttons(19).Image = 10
'        .Buttons(1).Enabled = False
'        .Buttons(3).Enabled = False
'        .Buttons(5).Enabled = False
'        .Buttons(7).Image = 11
'        .Buttons(7).Caption = "Save"
'        .Buttons(9).Image = 12
'        .Buttons(9).Caption = "Undo"
'        .Buttons(7).Enabled = True
'        .Buttons(9).Enabled = True
'        .Buttons(11).Enabled = False
'        .Buttons(13).Enabled = False
'        .Buttons(15).Enabled = False
'        .Buttons(17).Enabled = False
'        .Buttons(19).Enabled = False
'        .Buttons(1).ToolTipText = ""
'        .Buttons(3).ToolTipText = ""
'        .Buttons(5).ToolTipText = ""
'        .Buttons(7).ToolTipText = "SAVE (F5)"
'        .Buttons(9).ToolTipText = "UNDO (Esc)"
'        .Buttons(11).ToolTipText = ""
'        .Buttons(13).ToolTipText = ""
'        .Buttons(15).ToolTipText = ""
'        .Buttons(17).ToolTipText = ""
'        .Buttons(19).ToolTipText = ""
'    End If
'End With
'End Sub

Private Sub b8TitleBar1_CLoseClick()
cmdCancel_Click
End Sub

Private Sub b8TitleBar2_CLoseClick()
cmdCancelSearch_Click
End Sub

Private Sub cmbComp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtSSS.SetFocus
End If
End Sub

Private Sub cmbDivision_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    cmdDept.SetFocus
End If
End Sub

Private Sub cmbEffectivityDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKSearch_Click
End Sub

Private Sub cmbPost_Click()
locPost = cmbPost.ItemData(cmbPost.ListIndex)
End Sub

Private Sub cmbPost_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    cmbComp.SetFocus
End If
End Sub

Private Sub cmbTaxStatus_Click()
locTaxStatus = cmbTaxStatus.ItemData(cmbTaxStatus.ListIndex)
End Sub

Private Sub cmbTaxStatus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    cmbPost.SetFocus
End If
End Sub

Private Sub cmdCancel_Click()
picToolbar.Enabled = True
picMain.Enabled = True
picAdd.Visible = False
End Sub

Private Sub cmdCancelSearch_Click()
picToolbar.Enabled = True
picMain.Enabled = True
picSearch.Visible = False
End Sub

Private Sub cmdDept_Click()
locDept = cmdDept.ItemData(cmdDept.ListIndex)
End Sub

Private Sub cmdDept_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    cmdStatus.SetFocus
End If
End Sub

Private Sub cmdOK_Click()
If lstResult.ListIndex = -1 Then Exit Sub
CLEARTEXT
locEmployeePK = lstResult.ItemData(lstResult.ListIndex)
Array1 = Split(lstResult.List(lstResult.ListIndex), " - ", -1, 1)
txtID.Text = CStr(Array1(0))
txtName.Text = CStr(Array1(1))

s = "SELECT tbl_Personnel_Information.SSSNumber AS SSS, " & _
    " tbl_Personnel_Information.PHICNumber AS PHIC, " & _
    " tbl_Personnel_Information.HDMFNumber AS PagIbig, " & _
    " tbl_Personnel_Information.TIN AS TIN " & _
    " FROM tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
    " WHERE (tbl_Personnel_IDNumber.PK = " & locEmployeePK & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtSSS.Text = rs!SSS
    txtPHIC.Text = rs!PHIC
    txtPagIbig.Text = rs!PAGIBIG
    txtTIN.Text = rs!TIN
Else
    txtSSS.Text = "ON PROCESS"
    txtPHIC.Text = "ON PROCESS"
    txtPagIbig.Text = "ON PROCESS"
    txtTIN.Text = "ON PROCESS"
End If
rs.Close



s = "SELECT TOP 1 Is_SSS, Is_PHIC, Is_PAGIBIG, Is_TIN, " & _
    " Division, Dept, EmpStatus, TaxStatus, Positions, CompensationRate " & _
    " From tbl_Personnel_Action " & _
    " Where (EmpPK = " & locEmployeePK & ") " & _
    " And (EffectivityDate <= CONVERT(DateTime, CONVERT(Char(6), GETDATE(), 12), 102)) " & _
    " ORDER BY EffectivityDate DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    chkSSS.Value = rs!Is_SSS
    chkPHIC.Value = rs!Is_PHIC
    chkPagIbig.Value = rs!Is_PAGIBIG
    chkTIN.Value = rs!Is_TIN
    cmbDivision.ListIndex = rs!Division - 1
    Array1 = Split(DEPT_NAME(rs!Dept), ";", -1, 1)
    cmdDept.Text = CStr(Array1(1))
    Array1 = Split(EMP_STATUS(rs!EmpStatus), ";", -1, 1)
    cmdStatus.Text = CStr(Array1(1))
    Array1 = Split(TAX_STATUS_NAME(rs!TaxStatus), ";", -1, 1)
    cmbTaxStatus.Text = CStr(Array1(1))
    Array1 = Split(POSITION_NAME(rs!Positions), ";", -1, 1)
    cmbPost.Text = CStr(Array1(1))
    cmbComp.ListIndex = rs!CompensationRate - 1
    
    locDept = rs!Dept
    locPost = rs!Positions
    locEmpStatus = rs!EmpStatus
    locTaxStatus = rs!TaxStatus
    
End If
rs.Close
    
TRANSACTIONTYPE = 1
TOOLBAR_FUNC 2
'Me.Caption = "Personnel Action Memo - New"
LOCKTEXT False
cmdCancel_Click
cmbDivision.SetFocus
End Sub

Private Sub cmdOKSearch_Click()
If cmbEffectivityDate.ListIndex = -1 Then Exit Sub
s = "SELECT CntrlNo " & _
    " From tbl_Personnel_Action " & _
    " WHERE (EffectivityDate = '" & FormatDateTime(cmbEffectivityDate.List(cmbEffectivityDate.ListIndex), vbShortDate) & "') " & _
    " AND (EmpPK = " & lstResultSearch.ItemData(lstResultSearch.ListIndex) & ")"
ra.Open s, ConnOmega
If ra.State = adStateOpen Then ra.Close
ra.Open s, ConnOmega
If ra.RecordCount > 0 Then
    TOOLBAR_FUNC 1
    BROWSER ra!CntrlNo, "is_LOAD"
    'Me.Caption = "Personnel Action Memo"
    cmdCancelSearch_Click
End If
'ra.Close
If ra.State = adStateOpen Then ra.Close
End Sub

Private Sub cmdStatus_Click()
locEmpStatus = cmdStatus.ItemData(cmdStatus.ListIndex)
End Sub

Private Sub cmdStatus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    cmbTaxStatus.SetFocus
End If
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
    'Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "PersonnelActionCtrl", "PerActCtrl", ""), "is_HOME"
    'Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "PersonnelActionCtrl", "PerActCtrl", ""), "is_PAGEUP"
    'Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "PersonnelActionCtrl", "PerActCtrl", ""), "is_PAGEDOWN"
    'Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "PersonnelActionCtrl", "PerActCtrl", ""), "is_END"
    Case vbKeyEscape:   PRESS_ESCAPE
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 2
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
POPULATE_COMBO "PK", "DepartmentName", "tbl_Personnel_Department", "DepartmentName", cmdDept
'POPULATE_COMBO "PK", "StatusName", "tbl_Personnel_EmploymentStatus", "StatusName", cmdStatus
'POPULATE_COMBO "PK", "PositionName", "tbl_Personnel_Position", "PositionName", cmbPost
POPULATE_COMBO "PK", "TaxStatus", "tbl_Personnel_TaxStatus", "PK", cmbTaxStatus
cmdStatus.Clear
s = "SELECT PK, StatusName " & _
    " From tbl_Personnel_EmploymentStatus " & _
    " Where (Active = 1) " & _
    " ORDER BY StatusName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmdStatus.AddItem rs!StatusName
    cmdStatus.ItemData(cmdStatus.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close

With cmbComp
    .Clear
    .AddItem "MONTHLY"
    .AddItem "DAILY"
End With
CLEARTEXT
LOCKTEXT True
'TOOLBARBUTTON True
TOOLBAR_FUNC 0
'Me.Caption = "Personnel Action Memo"

HIDE_SALARY_RATE 0

'BROWSER GetSetting(App.EXEName, "PersonnelActionCtrl", "PerActCtrl", ""), "is_LOAD"
'If Trim(txtControl.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelActionCtrl", "PerActCtrl", ""), "is_HOME"
TRANSACTIONTYPE = is_REFRESH

'"Personnel Action Memo"

'If iSupervisory = 2 Then
'    If AccessRights("Personnel Action Memo", "Supervisory") = True Then
'        txtBasic.BackColor = &H80000005
'        txtCola.BackColor = &H80000005
'        txtTotal.BackColor = &H80000005
'        txtAllow.BackColor = &H80000005
'    Else
'        txtBasic.BackColor = &H0&
'        txtCola.BackColor = &H0&
'        txtTotal.BackColor = &H0&
'        txtAllow.BackColor = &H0&
'    End If
'Else
'    txtBasic.BackColor = &H0&
'    txtCola.BackColor = &H0&
'    txtTotal.BackColor = &H0&
'    txtAllow.BackColor = &H0&
'End If


tmp = SetWindowLong(txtSSS.hwnd, GWL_STYLE, GetWindowLong(txtSSS.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPHIC.hwnd, GWL_STYLE, GetWindowLong(txtPHIC.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPagIbig.hwnd, GWL_STYLE, GetWindowLong(txtPagIbig.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtTIN.hwnd, GWL_STYLE, GetWindowLong(txtTIN.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtRemarks.hwnd, GWL_STYLE, GetWindowLong(txtRemarks.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearchSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearchSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)

End Sub

Private Sub Form_Unload(Cancel As Integer)
If TRANSACTIONTYPE <> is_REFRESH Then
    Cancel = -1
End If
End Sub


Private Sub lstResult_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub lstResultSearch_Click()
If lstResultSearch.ListIndex = -1 Then cmbEffectivityDate.Clear: Exit Sub
cmbEffectivityDate.Clear
s = "SELECT EffectivityDate" & _
    " From tbl_Personnel_Action  " & _
    " Where (EmpPK = " & lstResultSearch.ItemData(lstResultSearch.ListIndex) & ") " & _
    " ORDER BY EffectivityDate DESC"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbEffectivityDate.AddItem Format(rs!EffectivityDate, "mm/dd/yyyy")
    rs.MoveNext
Wend
If cmbEffectivityDate.ListCount Then
    cmbEffectivityDate.ListIndex = 0
End If
rs.Close
End Sub

Private Sub lstResultSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbEffectivityDate.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "Save":    PRESS_F5
    Case "Undo":    PRESS_ESCAPE
    Case "Find":    PRESS_F6
    Case "Print":   PRESS_F9
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtAllow_GotFocus()
If iSupervisory = 2 Then
    If AccessRights("Personnel Action Memo", "Supervisory") = True Then
        txtAllow.Alignment = 0
        If IsNumeric(txtAllow.Text) Then
            txtAllow.Text = RETURNTEXTVALUE(txtAllow)
        End If
        HTEXT txtAllow
    End If
Else
    txtAllow.Alignment = 0
    If IsNumeric(txtAllow.Text) Then
        txtAllow.Text = RETURNTEXTVALUE(txtAllow)
    End If
    HTEXT txtAllow
End If
End Sub

Private Sub txtAllow_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtControl.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtCola.SetFocus
End If
End Sub

Private Sub txtAllow_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtAllow_LostFocus()
txtAllow.Alignment = 1
If IsNumeric(txtAllow.Text) Then
    txtAllow.Text = Format(RETURNTEXTVALUE(txtAllow), "#,##0.00")
Else
    txtAllow.Text = "0.00"
End If
End Sub

Private Sub txtBasic_Change()
txtTotal.Text = Format(RETURNTEXTVALUE(txtBasic) + RETURNTEXTVALUE(txtCola), "#,##0.00")
End Sub

Private Sub txtBasic_GotFocus()
If iSupervisory = 2 Then
    If AccessRights("Personnel Action Memo", "Supervisory") = True Then
        txtBasic.Alignment = 0
        If IsNumeric(txtBasic.Text) Then
            txtBasic.Text = CDbl(txtBasic.Text)
        End If
        HTEXT txtBasic
    End If
Else
    txtBasic.Alignment = 0
    If IsNumeric(txtBasic.Text) Then
        txtBasic.Text = CDbl(txtBasic.Text)
    End If
    HTEXT txtBasic
End If
End Sub

Private Sub txtBasic_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtCola.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtEffectDate.SetFocus
End If
End Sub

Private Sub txtBasic_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtBasic_LostFocus()
txtBasic.Alignment = 1
If IsNumeric(txtBasic.Text) Then
    txtBasic.Text = Format(CDbl(txtBasic.Text), "##,##0.00")
Else
    txtBasic.Text = "0.00"
End If
End Sub

Private Sub txtCola_Change()
txtTotal.Text = Format(RETURNTEXTVALUE(txtBasic) + RETURNTEXTVALUE(txtCola), "#,##0.00")
End Sub

Private Sub txtCola_GotFocus()
If iSupervisory = 2 Then
    If AccessRights("Personnel Action Memo", "Supervisory") = True Then
        txtCola.Alignment = 0
        If IsNumeric(txtCola.Text) Then
            txtCola.Text = RETURNTEXTVALUE(txtCola)
        End If
        HTEXT txtCola
    End If
Else
    txtCola.Alignment = 0
    If IsNumeric(txtCola.Text) Then
        txtCola.Text = RETURNTEXTVALUE(txtCola)
    End If
    HTEXT txtCola
End If
End Sub

Private Sub txtCola_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtAllow.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtBasic.SetFocus
End If
End Sub

Private Sub txtCola_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtCola_LostFocus()
txtCola.Alignment = 1
If IsNumeric(txtCola.Text) Then
    txtCola.Text = Format(RETURNTEXTVALUE(txtCola), "#,##0.00")
Else
    txtCola.Text = "0.00"
End If
End Sub

Private Sub txtControl_GotFocus()
HTEXT txtControl
End Sub

Private Sub txtControl_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtID.SetFocus
End If
End Sub

Private Sub txtEffectDate_GotFocus()
HTEXT txtEffectDate
End Sub

Private Sub txtEffectDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtBasic.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtRemarks.SetFocus
End If
End Sub

Private Sub txtEffectDate_LostFocus()
If IsDate(txtEffectDate.Text) Then
    txtEffectDate.Text = Format(FormatDateTime(txtEffectDate.Text, vbShortDate), "mm/dd/yyyy")
Else
    MsgBox "PLEASE SUPPLY A VALID DATE!         ", vbCritical, ""
    txtEffectDate.SetFocus
    HTEXT txtEffectDate
End If
End Sub

Private Sub txtID_GotFocus()
HTEXT txtID
End Sub

Private Sub txtID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtName.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtControl.SetFocus
End If
End Sub

Private Sub txtName_GotFocus()
HTEXT txtName
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    cmbDivision.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtID.SetFocus
End If
End Sub

Private Sub txtPagIbig_GotFocus()
HTEXT txtPagIbig
End Sub

Private Sub txtPagIbig_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtTIN.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtPHIC.SetFocus
End If
End Sub

Private Sub txtPHIC_GotFocus()
HTEXT txtPHIC
End Sub

Private Sub txtPHIC_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtPagIbig.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtSSS.SetFocus
End If
End Sub

Private Sub txtRemarks_GotFocus()
HTEXT txtRemarks
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtEffectDate.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtTIN.SetFocus
End If
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstResult.Clear: Exit Sub
lstResult.Clear
's = "sp_Personnel_Action_Search_Add('" & FORMATSQL(Trim(txtSearch.Text)) & "%')"
s = "SELECT tbl_Personnel_IDNumber.PK, " & _
    " tbl_Personnel_IDNumber.IDNumber, " & _
    " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName " & _
    " FROM tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
    " WHERE (tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
    " AND (ISNULL((SELECT TOP 1 tbl_Personnel_EmploymentStatus.Active " & _
    " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
    " tbl_Personnel_EmploymentStatus ON tbl_Personnel_Action.EmpStatus = tbl_Personnel_EmploymentStatus.PK " & _
    " WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_IDNumber.PK) " & _
    " AND (tbl_Personnel_Action.EffectivityDate <= CONVERT(DATETIME, CONVERT(char(6), getdate(), 12), 102)) ORDER BY tbl_Personnel_Action.EffectivityDate DESC), 0) = 1) " & _
    " OR (tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
    " AND (ISNULL((SELECT TOP 1 tbl_Personnel_EmploymentStatus.Active " & _
    " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
    " tbl_Personnel_EmploymentStatus ON tbl_Personnel_Action.EmpStatus = tbl_Personnel_EmploymentStatus.PK " & _
    " WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_IDNumber.PK) " & _
    " AND (tbl_Personnel_Action.EffectivityDate <= CONVERT(DATETIME, CONVERT(char(6), getdate(), 12), 102)) ORDER BY tbl_Personnel_Action.EffectivityDate DESC), 0) = 0) " & _
    " ORDER BY tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResult.AddItem rs!IDNumber & " - " & rs!EmployeeName
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

Private Sub txtSearchSearch_Change()
If Trim(txtSearchSearch.Text) = "" Then lstResultSearch.Clear: cmbEffectivityDate.Clear: Exit Sub
lstResultSearch.Clear
's = "sp_Personnel_Action_Search_Search('" & FORMATSQL(Trim(txtSearchSearch.Text)) & "%')"
's = "SELECT tbl_Personnel_IDNumber.PK, " & _
    " tbl_Personnel_IDNumber.IDNumber, " & _
    " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName " & _
    " FROM tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
    " WHERE (tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearchSearch.Text)) & "%') " & _
    " AND (ISNULL((SELECT TOP 1 tbl_Personnel_EmploymentStatus.Active " & _
    " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
    " tbl_Personnel_EmploymentStatus ON tbl_Personnel_Action.EmpStatus = tbl_Personnel_EmploymentStatus.PK " & _
    " WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_IDNumber.PK) " & _
    " AND (tbl_Personnel_Action.EffectivityDate <= CONVERT(DATETIME, CONVERT(char(6), getdate(), 12), 102)) ORDER BY tbl_Personnel_Action.EffectivityDate DESC), 0) = 1) " & _
    " ORDER BY tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName"

's = "SELECT tbl_Personnel_IDNumber.PK, tbl_Personnel_IDNumber.IDNumber, " & _
    " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName " & _
    " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
    " tbl_Personnel_IDNumber ON tbl_Personnel_Action.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
    " WHERE (tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearchSearch.Text)) & "%') " & _
    " GROUP BY tbl_Personnel_IDNumber.PK, tbl_Personnel_IDNumber.IDNumber, " & _
    " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName " & _
    " ORDER BY tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName, " & _
    " tbl_Personnel_IDNumber.IDNumber "
s = "SELECT tbl_Personnel_IDNumber.PK, tbl_Personnel_IDNumber.IDNumber, " & _
    " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName " & _
    " FROM tbl_Personnel_Action AS tbl_Personnel_Action_1 LEFT OUTER JOIN " & _
    " tbl_Personnel_IDNumber ON tbl_Personnel_Action_1.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
    " WHERE (tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearchSearch.Text)) & "%') " & _
    " GROUP BY tbl_Personnel_IDNumber.PK, tbl_Personnel_IDNumber.IDNumber, " & _
    " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName " & _
    " HAVING (ISNULL ((SELECT TOP 1 tbl_Personnel_EmploymentStatus_1.Active " & _
    " FROM tbl_Personnel_Action AS tbl_Personnel_Action_2 LEFT OUTER JOIN " & _
    " tbl_Personnel_EmploymentStatus AS tbl_Personnel_EmploymentStatus_1 ON " & _
    " tbl_Personnel_Action_2.EmpStatus = tbl_Personnel_EmploymentStatus_1.PK " & _
    " Where (tbl_Personnel_Action_2.EmpPK = tbl_Personnel_IDNumber.PK) " & _
    " ORDER BY tbl_Personnel_Action_2.EffectivityDate DESC), 0) = 1) " & _
    " ORDER BY tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName, " & _
    " tbl_Personnel_IDNumber.IDNumber "
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultSearch.AddItem rs!IDNumber & " - " & rs!EmployeeName
    lstResultSearch.ItemData(lstResultSearch.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResultSearch.ListCount Then lstResultSearch.ListIndex = 0
End Sub

Private Sub txtSearchSearch_GotFocus()
HTEXT txtSearchSearch
End Sub

Private Sub txtSearchSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultSearch.SetFocus
End Sub

Private Sub txtSSS_GotFocus()
HTEXT txtSSS
End Sub

Private Sub txtSSS_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtPHIC.SetFocus
ElseIf KeyCode = vbKeyUp Then
    cmbComp.SetFocus
End If
End Sub

Private Sub txtTIN_GotFocus()
HTEXT txtTIN
End Sub

Private Sub txtTIN_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtRemarks.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtTIN.SetFocus
End If
End Sub




