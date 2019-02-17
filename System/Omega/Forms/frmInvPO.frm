VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInvPO 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11940
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvPO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15600
      TabIndex        =   77
      Top             =   0
      Width           =   15600
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   78
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
         MouseIcon       =   "frmInvPO.frx":08CA
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   10380
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   79
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmInvPO.frx":0BE4
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
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   6390
      Width           =   11940
      _ExtentX        =   21061
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
   Begin RPVGCC.b8Container picSLine 
      Height          =   855
      Left            =   240
      TabIndex        =   42
      Top             =   5040
      Visible         =   0   'False
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox txtTypeDesc1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   76
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtItemKey1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtType1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtType 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   360
         Width           =   1275
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1440
         TabIndex        =   63
         Top             =   360
         Width           =   675
      End
      Begin VB.TextBox txtItemCode 
         Height          =   315
         Left            =   2160
         TabIndex        =   62
         Top             =   360
         Width           =   1155
      End
      Begin VB.TextBox txtItemDesc 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   360
         Width           =   2955
      End
      Begin VB.TextBox txtUnit 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1680
         TabIndex        =   60
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtCost 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7560
         TabIndex        =   59
         Top             =   360
         Width           =   1035
      End
      Begin VB.TextBox txtTotalNetCost 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   360
         Width           =   1275
      End
      Begin VB.TextBox txtQty1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5280
         TabIndex        =   57
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtUnit1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtItemCode1 
         Height          =   285
         Left            =   5760
         TabIndex        =   55
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtItemDesc1 
         Height          =   285
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtCost1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtTotalNetCost1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtItemKey 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.ComboBox cmbUnit 
         Height          =   315
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   360
         Width           =   1155
      End
      Begin VB.TextBox txtOQty 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1440
         TabIndex        =   49
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtOCost 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1920
         TabIndex        =   48
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtOQty1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtOCost1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtNetCost 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   360
         Width           =   1035
      End
      Begin VB.TextBox txtNetCost1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtTotalCost 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "TYPE"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "QTY"
         Height          =   255
         Left            =   1440
         TabIndex        =   70
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "ITEMCODE"
         Height          =   255
         Left            =   2160
         TabIndex        =   69
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM DESCRIPTION"
         Height          =   255
         Left            =   3360
         TabIndex        =   68
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "UNIT"
         Height          =   255
         Left            =   6360
         TabIndex        =   67
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "COST"
         Height          =   255
         Left            =   7560
         TabIndex        =   66
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL NETCOST"
         Height          =   255
         Left            =   9720
         TabIndex        =   65
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NETCOST"
         Height          =   255
         Left            =   8640
         TabIndex        =   64
         Top             =   120
         Width           =   1035
      End
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
            Picture         =   "frmInvPO.frx":12F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPO.frx":1FD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPO.frx":2CAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPO.frx":3985
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPO.frx":465F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPO.frx":5339
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPO.frx":6013
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPO.frx":6CED
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPO.frx":79C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPO.frx":82A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPO.frx":8F7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPO.frx":9C55
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPO.frx":A92F
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPO.frx":B609
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPO.frx":C2E3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBody 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   120
      ScaleHeight     =   5055
      ScaleWidth      =   11655
      TabIndex        =   1
      Top             =   1200
      Width           =   11655
      Begin VB.ComboBox cmbDeptName 
         Height          =   315
         Left            =   4920
         TabIndex        =   41
         Text            =   "Combo1"
         Top             =   990
         Width           =   6735
      End
      Begin VB.ComboBox cmbSuppName 
         Height          =   315
         Left            =   5640
         TabIndex        =   40
         Text            =   "Combo1"
         Top             =   0
         Width           =   6015
      End
      Begin VB.TextBox txtApproved 
         Height          =   315
         Left            =   5040
         TabIndex        =   16
         Top             =   4440
         Width           =   2475
      End
      Begin VB.TextBox txtChecked 
         Height          =   315
         Left            =   2520
         TabIndex        =   15
         Top             =   4440
         Width           =   2475
      End
      Begin VB.TextBox txtRequested 
         Height          =   315
         Left            =   0
         TabIndex        =   14
         Top             =   4440
         Width           =   2475
      End
      Begin VB.TextBox txtFaxNo 
         Height          =   315
         Left            =   8280
         TabIndex        =   13
         Top             =   660
         Width           =   3375
      End
      Begin VB.TextBox txtTelNo 
         Height          =   315
         Left            =   4920
         TabIndex        =   12
         Top             =   660
         Width           =   2115
      End
      Begin VB.TextBox txtAddress 
         Height          =   315
         Left            =   4920
         TabIndex        =   11
         Top             =   330
         Width           =   6735
      End
      Begin VB.TextBox txtTerms 
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Top             =   990
         Width           =   2475
      End
      Begin VB.TextBox txtRefNo 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   660
         Width           =   1155
      End
      Begin VB.TextBox txtPODate 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   330
         Width           =   1155
      End
      Begin VB.TextBox txtPONumber 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   0
         Width           =   1155
      End
      Begin VB.TextBox txtSuppCode 
         Height          =   315
         Left            =   4920
         TabIndex        =   6
         Top             =   0
         Width           =   700
      End
      Begin VB.TextBox txtSuppKey 
         Height          =   315
         Left            =   3240
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtRemarks 
         Height          =   315
         Left            =   840
         TabIndex        =   4
         Top             =   3840
         Width           =   6680
      End
      Begin VB.TextBox txtDisc1 
         Height          =   315
         Left            =   4920
         TabIndex        =   3
         Top             =   1320
         Width           =   2130
      End
      Begin VB.TextBox txtVat 
         Height          =   315
         Left            =   8280
         TabIndex        =   2
         Top             =   1320
         Width           =   3375
      End
      Begin MSComctlLib.ListView lstDetail 
         Height          =   2025
         Left            =   0
         TabIndex        =   17
         Top             =   1680
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   3572
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
         NumItems        =   14
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
            Text            =   "TypeKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Type"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Unit"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ItemCode"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Item Description"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Cost"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "NetCost"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "Total NetCost"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "O_Qty"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "O_Cost"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "TotalCost"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblTotalCost 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10080
         TabIndex        =   39
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL COST"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9000
         TabIndex        =   38
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "General Manager"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5040
         TabIndex        =   37
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "APPROVED BY"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5040
         TabIndex        =   36
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Chief Accountant"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   35
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "CHECKED BY"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   34
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Purchasing Officer"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "REQUESTED BY"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   32
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "FAX NO"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7320
         TabIndex        =   31
         Top             =   675
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "TEL NO"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4080
         TabIndex        =   30
         Top             =   675
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4080
         TabIndex        =   29
         Top             =   345
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Terms of Payment"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   1000
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "REF/PR #"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   27
         Top             =   680
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PO DATE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   26
         Top             =   350
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PO #"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4080
         TabIndex        =   24
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "DEPT"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4080
         TabIndex        =   23
         Top             =   1005
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "REMARKS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "DISCOUNT"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4080
         TabIndex        =   21
         Top             =   1335
         Width           =   975
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "VAT"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7320
         TabIndex        =   20
         Top             =   1335
         Width           =   975
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL NETCOST"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9000
         TabIndex        =   19
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label lblTotalNetCost 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10080
         TabIndex        =   18
         Top             =   4080
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmInvPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2
Const is_FINDING = 3

Dim TRANS_DETAIL As Long
Const is_DET_REFRESH = 0
Const is_DET_ADDING = 1
Const is_DET_EDITTING = 2

Dim is_DET_FOCUS    As Long
Dim ROW             As Long
Dim tmp             As Long

Dim x, sPONumber, iSupplier, iDept, i, _
k, strDisc, dblTotalCost, a, _
dblTotalNetCost, dblTotalQty


Public Sub BROWSER(sPONum, isWant As String)
Select Case isWant
    Case "is_LOAD"
        If sPONum <> "" Then
            s = "SELECT TOP 1 tbl_Inv_PO.PK, tbl_Inv_PO.PONumber, tbl_Inv_PO.PODate, tbl_Inv_PO.PRNumber, " & _
                " tbl_Inv_PO.SuppKey, tbl_Inv_PO.DeptKey, tbl_Inv_PO.Terms, tbl_Inv_PO.Remarks, tbl_Inv_PO.TotalCost, " & _
                " tbl_Inv_PO.TotalDiscount, tbl_Inv_PO.TotalNetCost, tbl_Inv_PO.RequestBy, tbl_Inv_PO.CheckedBy, " & _
                " tbl_Inv_PO.ApprovedBy, tbl_Inv_PO.Posted, tbl_Inv_PO.Printed, tbl_Inv_PO.LastModified, " & _
                " tbl_Inv_PO.RecdPartFull, tbl_Inv_PO.Discount, tbl_Inv_PO.TotalQty, tbl_Inv_PO.TotalRcd, " & _
                " tbl_Inv_Supplier.SupplierCode, tbl_Inv_Supplier.SupplierName, tbl_Inv_Supplier.Address1, " & _
                " tbl_Inv_Supplier.Address2, tbl_Inv_Supplier.Address3, tbl_Inv_Supplier.TelNo, tbl_Inv_Supplier.FaxNo, " & _
                " tbl_Inv_Supplier.Email, tbl_Inv_Supplier.ContactPerson, tbl_GL_Department.Code as DepartmentCode, " & _
                " tbl_GL_Department.DeptName as DepartmentName " & _
                " FROM tbl_Inv_PO LEFT OUTER JOIN " & _
                " tbl_Inv_Supplier ON tbl_Inv_PO.SuppKey = tbl_Inv_Supplier.PK LEFT OUTER JOIN " & _
                " tbl_GL_Department ON tbl_Inv_PO.DeptKey = tbl_GL_Department.PK " & _
                " WHERE (tbl_Inv_PO.PONumber = '" & sPONum & "') " & _
                " ORDER BY tbl_Inv_PO.PONumber"
        Else
            s = "SELECT TOP 1 tbl_Inv_PO.PK, tbl_Inv_PO.PONumber, tbl_Inv_PO.PODate, tbl_Inv_PO.PRNumber, " & _
                " tbl_Inv_PO.SuppKey, tbl_Inv_PO.DeptKey, tbl_Inv_PO.Terms, tbl_Inv_PO.Remarks, tbl_Inv_PO.TotalCost, " & _
                " tbl_Inv_PO.TotalDiscount, tbl_Inv_PO.TotalNetCost, tbl_Inv_PO.RequestBy, tbl_Inv_PO.CheckedBy, " & _
                " tbl_Inv_PO.ApprovedBy, tbl_Inv_PO.Posted, tbl_Inv_PO.Printed, tbl_Inv_PO.LastModified, " & _
                " tbl_Inv_PO.RecdPartFull, tbl_Inv_PO.Discount, tbl_Inv_PO.TotalQty, tbl_Inv_PO.TotalRcd, " & _
                " tbl_Inv_Supplier.SupplierCode, tbl_Inv_Supplier.SupplierName, tbl_Inv_Supplier.Address1, " & _
                " tbl_Inv_Supplier.Address2, tbl_Inv_Supplier.Address3, tbl_Inv_Supplier.TelNo, tbl_Inv_Supplier.FaxNo, " & _
                " tbl_Inv_Supplier.Email, tbl_Inv_Supplier.ContactPerson, tbl_GL_Department.Code as DepartmentCode, " & _
                " tbl_GL_Department.DeptName as DepartmentName " & _
                " FROM tbl_Inv_PO LEFT OUTER JOIN " & _
                " tbl_Inv_Supplier ON tbl_Inv_PO.SuppKey = tbl_Inv_Supplier.PK LEFT OUTER JOIN " & _
                " tbl_GL_Department ON tbl_Inv_PO.DeptKey = tbl_GL_Department.PK " & _
                " ORDER BY tbl_Inv_PO.PONumber"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_PO.PK, tbl_Inv_PO.PONumber, tbl_Inv_PO.PODate, tbl_Inv_PO.PRNumber, " & _
            " tbl_Inv_PO.SuppKey, tbl_Inv_PO.DeptKey, tbl_Inv_PO.Terms, tbl_Inv_PO.Remarks, tbl_Inv_PO.TotalCost, " & _
            " tbl_Inv_PO.TotalDiscount, tbl_Inv_PO.TotalNetCost, tbl_Inv_PO.RequestBy, tbl_Inv_PO.CheckedBy, " & _
            " tbl_Inv_PO.ApprovedBy, tbl_Inv_PO.Posted, tbl_Inv_PO.Printed, tbl_Inv_PO.LastModified, " & _
            " tbl_Inv_PO.RecdPartFull, tbl_Inv_PO.Discount, tbl_Inv_PO.TotalQty, tbl_Inv_PO.TotalRcd, " & _
            " tbl_Inv_Supplier.SupplierCode, tbl_Inv_Supplier.SupplierName, tbl_Inv_Supplier.Address1, " & _
            " tbl_Inv_Supplier.Address2, tbl_Inv_Supplier.Address3, tbl_Inv_Supplier.TelNo, tbl_Inv_Supplier.FaxNo, " & _
            " tbl_Inv_Supplier.Email, tbl_Inv_Supplier.ContactPerson, tbl_GL_Department.Code as DepartmentCode, " & _
            " tbl_GL_Department.DeptName as DepartmentName " & _
            " FROM tbl_Inv_PO LEFT OUTER JOIN " & _
            " tbl_Inv_Supplier ON tbl_Inv_PO.SuppKey = tbl_Inv_Supplier.PK LEFT OUTER JOIN " & _
            " tbl_GL_Department ON tbl_Inv_PO.DeptKey = tbl_GL_Department.PK " & _
            " ORDER BY tbl_Inv_PO.PONumber"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_PO.PK, tbl_Inv_PO.PONumber, tbl_Inv_PO.PODate, tbl_Inv_PO.PRNumber, " & _
            " tbl_Inv_PO.SuppKey, tbl_Inv_PO.DeptKey, tbl_Inv_PO.Terms, tbl_Inv_PO.Remarks, tbl_Inv_PO.TotalCost, " & _
            " tbl_Inv_PO.TotalDiscount, tbl_Inv_PO.TotalNetCost, tbl_Inv_PO.RequestBy, tbl_Inv_PO.CheckedBy, " & _
            " tbl_Inv_PO.ApprovedBy, tbl_Inv_PO.Posted, tbl_Inv_PO.Printed, tbl_Inv_PO.LastModified, " & _
            " tbl_Inv_PO.RecdPartFull, tbl_Inv_PO.Discount, tbl_Inv_PO.TotalQty, tbl_Inv_PO.TotalRcd, " & _
            " tbl_Inv_Supplier.SupplierCode, tbl_Inv_Supplier.SupplierName, tbl_Inv_Supplier.Address1, " & _
            " tbl_Inv_Supplier.Address2, tbl_Inv_Supplier.Address3, tbl_Inv_Supplier.TelNo, tbl_Inv_Supplier.FaxNo, " & _
            " tbl_Inv_Supplier.Email, tbl_Inv_Supplier.ContactPerson, tbl_GL_Department.Code as DepartmentCode, " & _
            " tbl_GL_Department.DeptName as DepartmentName " & _
            " FROM tbl_Inv_PO LEFT OUTER JOIN " & _
            " tbl_Inv_Supplier ON tbl_Inv_PO.SuppKey = tbl_Inv_Supplier.PK LEFT OUTER JOIN " & _
            " tbl_GL_Department ON tbl_Inv_PO.DeptKey = tbl_GL_Department.PK " & _
            " WHERE (tbl_Inv_PO.PONumber < '" & sPONum & "') " & _
            " ORDER BY tbl_Inv_PO.PONumber DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_PO.PK, tbl_Inv_PO.PONumber, tbl_Inv_PO.PODate, tbl_Inv_PO.PRNumber, " & _
            " tbl_Inv_PO.SuppKey, tbl_Inv_PO.DeptKey, tbl_Inv_PO.Terms, tbl_Inv_PO.Remarks, tbl_Inv_PO.TotalCost, " & _
            " tbl_Inv_PO.TotalDiscount, tbl_Inv_PO.TotalNetCost, tbl_Inv_PO.RequestBy, tbl_Inv_PO.CheckedBy, " & _
            " tbl_Inv_PO.ApprovedBy, tbl_Inv_PO.Posted, tbl_Inv_PO.Printed, tbl_Inv_PO.LastModified, " & _
            " tbl_Inv_PO.RecdPartFull, tbl_Inv_PO.Discount, tbl_Inv_PO.TotalQty, tbl_Inv_PO.TotalRcd, " & _
            " tbl_Inv_Supplier.SupplierCode, tbl_Inv_Supplier.SupplierName, tbl_Inv_Supplier.Address1, " & _
            " tbl_Inv_Supplier.Address2, tbl_Inv_Supplier.Address3, tbl_Inv_Supplier.TelNo, tbl_Inv_Supplier.FaxNo, " & _
            " tbl_Inv_Supplier.Email, tbl_Inv_Supplier.ContactPerson, tbl_GL_Department.Code as DepartmentCode, " & _
            " tbl_GL_Department.DeptName as DepartmentName " & _
            " FROM tbl_Inv_PO LEFT OUTER JOIN " & _
            " tbl_Inv_Supplier ON tbl_Inv_PO.SuppKey = tbl_Inv_Supplier.PK LEFT OUTER JOIN " & _
            " tbl_GL_Department ON tbl_Inv_PO.DeptKey = tbl_GL_Department.PK " & _
            " WHERE (tbl_Inv_PO.PONumber > '" & sPONum & "') " & _
            " ORDER BY tbl_Inv_PO.PONumber "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_PO.PK, tbl_Inv_PO.PONumber, tbl_Inv_PO.PODate, tbl_Inv_PO.PRNumber, " & _
            " tbl_Inv_PO.SuppKey, tbl_Inv_PO.DeptKey, tbl_Inv_PO.Terms, tbl_Inv_PO.Remarks, tbl_Inv_PO.TotalCost, " & _
            " tbl_Inv_PO.TotalDiscount, tbl_Inv_PO.TotalNetCost, tbl_Inv_PO.RequestBy, tbl_Inv_PO.CheckedBy, " & _
            " tbl_Inv_PO.ApprovedBy, tbl_Inv_PO.Posted, tbl_Inv_PO.Printed, tbl_Inv_PO.LastModified, " & _
            " tbl_Inv_PO.RecdPartFull, tbl_Inv_PO.Discount, tbl_Inv_PO.TotalQty, tbl_Inv_PO.TotalRcd, " & _
            " tbl_Inv_Supplier.SupplierCode, tbl_Inv_Supplier.SupplierName, tbl_Inv_Supplier.Address1, " & _
            " tbl_Inv_Supplier.Address2, tbl_Inv_Supplier.Address3, tbl_Inv_Supplier.TelNo, tbl_Inv_Supplier.FaxNo, " & _
            " tbl_Inv_Supplier.Email, tbl_Inv_Supplier.ContactPerson, tbl_GL_Department.Code as DepartmentCode, " & _
            " tbl_GL_Department.DeptName as DepartmentName " & _
            " FROM tbl_Inv_PO LEFT OUTER JOIN " & _
            " tbl_Inv_Supplier ON tbl_Inv_PO.SuppKey = tbl_Inv_Supplier.PK LEFT OUTER JOIN " & _
            " tbl_GL_Department ON tbl_Inv_PO.DeptKey = tbl_GL_Department.PK " & _
            " ORDER BY tbl_Inv_PO.PONumber DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iSupplier = rs!SuppKey
    iDept = rs!DeptKey
    txtPONumber.Text = rs!PONumber
    txtPODate.Text = Format(rs!PODate, "mm/dd/yyyy")
    txtRefNo.Text = rs!PRNumber
    txtTerms.Text = rs!Terms
    txtSuppKey.Text = rs!SuppKey
    txtSuppCode.Text = rs!SupplierCode
    cmbSuppName.Text = rs!SupplierName
    cmbDeptName.Text = UCase(rs!DepartmentName)
    txtAddress.Text = rs!Address1 & " " & rs!Address2 & " " & rs!Address3
    txtTelNo.Text = rs!TelNo
    txtFaxNo.Text = rs!FaxNo
    txtRemarks.Text = rs!Remarks
    txtRequested.Text = rs!RequestBy
    txtChecked.Text = rs!CheckedBy
    txtApproved.Text = rs!ApprovedBy
    txtDisc1.Text = IIf(IsNull(rs!Discount), "", rs!Discount)
    lblTotalCost.Caption = Format(rs!TotalCost, "#,##0.00")
    lblTotalNetCost.Caption = Format(rs!TotalNetCost, "#,##0.00")
    txtDisc1.Text = IIf(IsNull(rs!Discount), "", rs!Discount)
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    imgPosted.Visible = IIf(rs!Posted = 1, True, False)
    Toolbar1.Buttons(19).ToolTipText = IIf(rs!Posted = 1, "UNPOST (F8)", "POST (F8)")
    
    SaveSetting App.EXEName, "PONumber", "PONum", rs!PONumber
    
    t = "SELECT tbl_Inv_PODet.* " & _
        " FROM tbl_Inv_PODet " & _
        " WHERE (POKey = " & rs!PK & ") " & _
        " ORDER BY Line"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    CUSTOMIZE_DETAIL
    If rt.RecordCount > 0 Then
        With lstDetail.ListItems
            .Clear
            While Not rt.EOF
                Set x = .Add()
                x.Text = rt!ItemKey
                x.SubItems(1) = Format(rt!line, "0#")
                x.SubItems(2) = rt!Item_FA
                x.SubItems(3) = IIf(rt!Item_FA = 1, "Items", IIf(rt!Item_FA = 2, "Fixed Asset", ""))
                x.SubItems(4) = Format(rt!Qty, "#,##0.00")
                x.SubItems(5) = rt!Unit
                
                Select Case rt!Item_FA
                    Case 1
                        u = "SELECT ItemCode, ItemDesc " & _
                            " FROM tbl_Inv_Items " & _
                            " WHERE (PK = " & rt!ItemKey & ")"
                    Case 2
                        u = "SELECT Code as ItemCode, " & _
                            " Description as ItemDesc " & _
                            " FROM tbl_FA_Items " & _
                            " WHERE (PK = " & rt!ItemKey & ")"
                End Select
                If ru.State = adStateOpen Then ru.Close
                ru.Open u, ConnOmega
                If ru.RecordCount > 0 Then
                    x.SubItems(6) = ru!ItemCode
                    x.SubItems(7) = ru!ItemDesc
                Else
                    x.SubItems(6) = " "
                    x.SubItems(7) = " "
                End If
                ru.Close
                
                x.SubItems(8) = Format(rt!Cost, "#,##0.00")
                x.SubItems(9) = Format(rt!NetCost, "#,##0.00")
                x.SubItems(10) = Format(rt!TotalNetCost, "#,##0.00")
                x.SubItems(11) = rt!O_Qty
                x.SubItems(12) = rt!O_Cost
                x.SubItems(13) = rt!TotalCost
                rt.MoveNext
            Wend
        End With
    End If
    rt.Close
End If
rs.Close
End Sub


Private Sub PRESS_INSERT()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSLine.Visible = True Then Exit Sub
    If AccessRights("Purchase Order", "Add") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    CLEARTEXT
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_ADDING
    'Me.Caption = "PURCHASE ORDER - NEW"
    s = "SELECT TOP 1 PONumber " & _
        " FROM tbl_Inv_PO " & _
        " ORDER BY PONumber DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        txtPONumber.Text = Format(CDbl(rs!PONumber) + 1, "0000000#")
    End If
    rs.Close
    txtRequested.Text = GetSetting(App.EXEName, "PO_Requested", "PO_Req", "")
    txtChecked.Text = GetSetting(App.EXEName, "PO_Checked", "PO_Check", "")
    txtApproved.Text = GetSetting(App.EXEName, "PO_Approved", "PO_Approve", "")
    txtPODate.Text = Format(FormatDateTime(Date, vbShortDate), "mm/dd/yyyy")
    txtPONumber.SetFocus
ElseIf TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If picSLine.Visible = True Then Exit Sub
    If is_DET_FOCUS = 1 Then
        If TRANS_DETAIL = is_DET_REFRESH Then
            If imgPosted.Visible = False And Toolbar1.Buttons(1).Enabled = True Then
                With lstDetail.ListItems
                    If Trim(.Item(1).SubItems(1)) <> "" Then
                        Set x = .Add()
                        x.Text = ""
                        x.SubItems(1) = Format(.Count, "0#")
                        x.SubItems(2) = " "
                        x.SubItems(3) = " "
                        x.SubItems(4) = " "
                        x.SubItems(5) = " "
                        x.SubItems(6) = " "
                        x.SubItems(7) = " "
                        x.SubItems(8) = " "
                        x.SubItems(9) = " "
                        x.SubItems(10) = " "
                        ROW = .Count
                    Else
                        .Item(1).SubItems(1) = "01"
                        ROW = 1
                    End If
                    CLEAR_DETAIL
                    picBody.Enabled = False
                    picSLine.Visible = True
                    lstDetail.ListItems(ROW).EnsureVisible
                    lstDetail.ListItems(ROW).Selected = True
                    TRANS_DETAIL = is_DET_ADDING
                    TOOLBARFUNC 2
                    cmbType.SetFocus
                    'txtQty.SetFocus
                End With
            End If
        End If
    End If
End If
End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE = is_REFRESH Then
    If Statusbar1.Panels(1).Text = "" Then Exit Sub
    If picSLine.Visible = True Then Exit Sub
        
    If AccessRights("Purchase Order", "Edit") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If imgPosted.Visible = True Then MsgBox "ALREADY POSTED!             ", vbCritical, "Posted": Exit Sub
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_EDITTING
    'Me.Caption = "PURCHASE ORDER - EDIT"
    
ElseIf TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If picSLine.Visible = True Then Exit Sub
    If is_DET_FOCUS = 1 Then
        If imgPosted.Visible = False And Toolbar1.Buttons(3).Enabled = True Then
            With lstDetail.ListItems
                If Trim(.Item(ROW).SubItems(1)) <> "" Then
                    txtQty.Text = .Item(ROW).SubItems(4)
                    t = "SELECT PK, ItemCode, ItemDesc, " & _
                        " Unit, Unit2, Cost, SuppKey" & _
                        " From tbl_Inv_Items " & _
                        " WHERE (ItemCode = '" & Trim(.Item(ROW).SubItems(6)) & "')"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnOmega
                    If rt.RecordCount > 0 Then
                        With cmbUnit
                            .Clear
                            .AddItem rt!Unit
                            If Trim(rt!Unit2) <> "" Then
                                .AddItem rt!Unit2
                            End If
                            If Trim(lstDetail.ListItems.Item(ROW).SubItems(5)) = rt!Unit Then
                                .ListIndex = 0
                            End If
                            If Trim(lstDetail.ListItems.Item(ROW).SubItems(5)) = rt!Unit2 Then
                                .ListIndex = 1
                            End If
                        End With
                    End If
                    rt.Close
                    txtType.Text = .Item(ROW).SubItems(2)
                    cmbType.ListIndex = .Item(ROW).SubItems(2) - 1
                    txtItemCode.Text = .Item(ROW).SubItems(6)
                    txtItemDesc.Text = .Item(ROW).SubItems(7)
                    txtCost.Text = .Item(ROW).SubItems(8)
                    txtNetCost.Text = .Item(ROW).SubItems(9)
                    txtTotalNetCost.Text = .Item(ROW).SubItems(10)
                    
                    txtItemKey1.Text = .Item(ROW).Text
                    txtType1.Text = .Item(ROW).SubItems(2)
                    txtTypeDesc1.Text = .Item(ROW).SubItems(3)
                    txtQty1.Text = .Item(ROW).SubItems(4)
                    txtUnit1.Text = .Item(ROW).SubItems(5)
                    txtItemCode1.Text = .Item(ROW).SubItems(6)
                    txtItemDesc1.Text = .Item(ROW).SubItems(7)
                    txtCost1.Text = .Item(ROW).SubItems(8)
                    txtNetCost1.Text = .Item(ROW).SubItems(9)
                    txtTotalNetCost1.Text = .Item(ROW).SubItems(10)
                    picBody.Enabled = False
                    picSLine.Visible = True
                    TRANS_DETAIL = is_DET_EDITTING
                    TOOLBARFUNC 2
                    cmbType.SetFocus
                    'txtQty.SetFocus
                End If
            End With
        End If
    End If
End If
End Sub

Private Sub PRESS_DELETE()

If TRANSACTIONTYPE = is_REFRESH Then
    If Statusbar1.Panels(1).Text = "" Then Exit Sub
    If picSLine.Visible = True Then Exit Sub
    If AccessRights("Purchase Order", "Delete") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If imgPosted.Visible = True Then MsgBox "ALREADY POSTED!             ", vbCritical, "Posted": Exit Sub
    If MsgBox("ARE YOU SURE TO DELETE THIS RECORD?          ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    On Error GoTo PG:
    ConnOmega.Execute "DELETE FROM tbl_Inv_PO WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    CLEARTEXT
    BROWSER GetSetting(App.EXEName, "PONumber", "PONum", ""), "is_PAGEDOWN"
    If Trim(txtPONumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "PONumber", "PONum", ""), "is_HOME"
            
ElseIf TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If picSLine.Visible = True Then Exit Sub
    If is_DET_FOCUS = 1 Then
        If imgPosted.Visible = False And Toolbar1.Buttons(5).Enabled = True Then
            With lstDetail.ListItems
                If Trim(.Item(ROW).SubItems(1)) <> "" Then
                    If .Count > 1 Then
                        .Remove ROW
                        For i = 1 To .Count
                            .Item(i).SubItems(1) = Format(i, "0#")
                        Next i
                        If ROW > .Count Then
                            ROW = .Count
                        End If
                    Else
                        .Item(1).SubItems(1) = " "
                        .Item(1).SubItems(2) = " "
                        .Item(1).SubItems(3) = " "
                        .Item(1).SubItems(4) = " "
                        .Item(1).SubItems(5) = " "
                        .Item(1).SubItems(6) = " "
                        .Item(1).SubItems(7) = " "
                        .Item(1).SubItems(8) = " "
                        .Item(1).SubItems(9) = " "
                        .Item(1).SubItems(10) = " "
                        ROW = 1
                    End If
                    lstDetail.ListItems(ROW).EnsureVisible
                    lstDetail.ListItems(ROW).Selected = True
                    lstDetail.SetFocus
                End If
            End With
        End If
    End If
End If
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Function CHECK_ITEM(iPK) As Long
s = "SELECT SuppKey" & _
    " From tbl_Inv_Items " & _
    " WHERE (PK = " & iPK & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    CHECK_ITEM = rs!SuppKey
End If
rs.Close
End Function

Private Sub PRESS_F5()
If picSLine.Visible = True Then Exit Sub

If TRANSACTIONTYPE = is_ADDING And _
Toolbar1.Buttons(7).Enabled = True Then
    If Trim(txtPONumber.Text) = "" Then
        MsgBox "Please Supply PO # !            ", vbCritical, "Error..."
        txtPONumber.SetFocus
        HTEXT txtPONumber
        Exit Sub
    End If
    If IsDate(txtPODate.Text) = False Then
        MsgBox "Please Supply A Valid Date!         ", vbCritical, "Error..."
        txtPODate.SetFocus
        HTEXT txtPODate
        Exit Sub
    End If
    If Trim(txtRefNo.Text) = "" Then
        MsgBox "Please Supply Requisition #!            ", vbCritical, "Error..."
        txtRefNo.SetFocus
        HTEXT txtRefNo
        Exit Sub
    End If
    If Trim(txtTerms.Text) = "" Then
        MsgBox "Please Supply Terms!            ", vbCritical, "Error..."
        txtTerms.SetFocus
        HTEXT txtTerms
        Exit Sub
    End If
    If CDbl(iSupplier) = 0 Then
        MsgBox "Please Supply Supplier!         ", vbCritical, "Error..."
        txtSuppCode.SetFocus
        Exit Sub
    End If
    If Trim(txtRequested.Text) = "" Then
        MsgBox "Please Suply Requested Person!          ", vbCritical, "Error..."
        txtRequested.SetFocus
        HTEXT txtRequested
        Exit Sub
    End If
    If Trim(txtChecked.Text) = "" Then
        MsgBox "Please Supply Person Verified this P.O. !           ", vbCritical, "Error..."
        txtChecked.SetFocus
        HTEXT txtChecked
        Exit Sub
    End If
    If Trim(txtApproved.Text) = "" Then
        MsgBox "Please Supply Approving Officer!            ", vbCritical, "Error..."
        txtApproved.SetFocus
        HTEXT txtApproved
        Exit Sub
    End If
    If CDbl(iDept) = 0 Then
        MsgBox "Please Supply Requesting Department!            ", vbCritical, "Error..."
        cmbDeptName.SetFocus
        Exit Sub
    End If
    With lstDetail.ListItems
    If Trim(.Item(1).SubItems(1)) <> "" Then
        For k = 1 To .Count
            lstDetail.ListItems(k).EnsureVisible
            lstDetail.ListItems(k).Selected = True
            If CDbl(.Item(k).SubItems(2)) = 1 Then
                If CHECK_ITEM(.Item(k).Text) <> CLng(txtSuppKey.Text) Then GoTo PX:
            End If
        Next k
    End If
    
    On Error GoTo PG:
    strDisc = IIf(Trim(txtDisc1.Text) = "", "", Replace(Trim(txtDisc1.Text), "%", "") & "%")
    
'    sPONumber = ""
'    s = "SELECT TOP 1 PONumber " & _
'        " FROM tbl_Inv_PO " & _
'        " WHERE (PODate = " & Format(FormatDateTime(txtPODate.Text, vbShortDate), "yyyy") & ") " & _
'        " ORDER BY PONumber DESC"
'    If rs.State = adStateOpen Then rs.Close
'    rs.Open s, ConnOmega
'    If rs.RecordCount > 0 Then
'        sPONumber = Format(CDbl(rs!PONumber) + 1, "0000000#")
'    Else
'        sPONumber = Format(FormatDateTime(txtPODate.Text, vbShortDate), "yyyy") & "0000"
'    End If
'    rs.Close
    
    'sPONumber = Format(CDbl(RETURNTEXTVALUE(txtPONumber)) + 1, "0000000#")
    
    sPONumber = Format(CDbl(RETURNTEXTVALUE(txtPONumber)), "0000000#")
    
    Do
        s = "SELECT tbl_Inv_PO.* " & _
            " FROM tbl_Inv_PO " & _
            " WHERE (PONumber = '" & sPONumber & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If
        rs.Close
        sPONumber = Format(CDbl(sPONumber) + 1, "0000000#")
    Loop
    
    ConnOmega.Execute "INSERT INTO tbl_Inv_PO" & _
                      " (PONumber, PODate, PRNumber, " & _
                      " SuppKey, Terms, Remarks, TotalCost, " & _
                      " RequestBy, CheckedBy, ApprovedBy, " & _
                      " LastModified, DeptKey, Discount, TotalNetCost) " & _
                      " VALUES('" & sPONumber & "', " & _
                      " '" & FormatDateTime(txtPODate.Text, vbShortDate) & "', '" & Trim(txtRefNo.Text) & "', " & _
                      " " & iSupplier & ", '" & FORMATSQL(Trim(txtTerms.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtRemarks.Text)) & "', " & CDbl(lblTotalCost.Caption) & ",  " & _
                      " '" & FORMATSQL(Trim(txtRequested.Text)) & "', '" & FORMATSQL(Trim(txtChecked.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtApproved.Text)) & "', '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                      " " & iDept & ", '" & strDisc & "'," & CDbl(lblTotalNetCost.Caption) & ")"
    s = "SELECT PK" & _
        " FROM tbl_Inv_PO" & _
        " WHERE (PONumber = '" & Trim(txtPONumber.Text) & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        ConnOmega.Execute "DELETE FROM tbl_Inv_PODet WHERE (POKey = " & rs!PK & ")"
        dblTotalQty = 0
        dblTotalCost = 0
        dblTotalNetCost = 0
        If Trim(.Item(1).SubItems(1)) <> "" Then
            For i = 1 To .Count
                If Trim(.Item(i).Text) <> "" Then
                    lstDetail.ListItems(i).EnsureVisible
                    lstDetail.ListItems(i).Selected = True
                    If Trim(strDisc) <> "" Then
                        .Item(i).SubItems(9) = Format(CDbl(.Item(i).SubItems(8)) * ((100 - Val(Trim(strDisc))) / 100), "#,##0.00")
                    Else
                        .Item(i).SubItems(9) = .Item(i).SubItems(8)
                    End If
                    dblTotalQty = dblTotalQty + CDbl(.Item(i).SubItems(11))
                    dblTotalCost = dblTotalCost + (CDbl(.Item(i).SubItems(8)) * CDbl(.Item(i).SubItems(4)))
                    dblTotalNetCost = dblTotalNetCost + (CDbl(.Item(i).SubItems(9)) * CDbl(.Item(i).SubItems(4)))
                    ConnOmega.Execute "INSERT INTO tbl_Inv_PODet" & _
                                      " (POKey, Line, ItemKey, Qty, Unit, Cost, O_Qty, O_Cost, Discount, NetCost, Item_FA) " & _
                                      " VALUES(" & rs!PK & ", " & i & ",  " & _
                                      " " & .Item(i).Text & ", " & CDbl(.Item(i).SubItems(4)) & ", " & _
                                      " '" & CStr(FORMATSQL(Trim((.Item(i).SubItems(5))))) & "', " & _
                                      " " & CDbl(.Item(i).SubItems(8)) & ", " & CDbl(.Item(i).SubItems(11)) & ", " & _
                                      " " & CDbl(.Item(i).SubItems(12)) & ", '" & strDisc & "', " & _
                                      " " & CDbl(.Item(i).SubItems(9)) & ", " & CDbl(.Item(i).SubItems(2)) & ")"
                End If
            Next i
        End If
        ConnOmega.Execute "UPDATE tbl_Inv_PO SET TotalCost = " & CDbl(dblTotalCost) & ", TotalNetCost = " & CDbl(dblTotalNetCost) & ", TotalQty = " & CDbl(dblTotalQty) & " WHERE (PK = " & rs!PK & ")"
    End If
    End With
    rs.Close
    LOCKTEXT True
    txtPONumber.SetFocus
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    SaveSetting App.EXEName, "PO_Requested", "PO_Req", Trim(txtRequested.Text)
    SaveSetting App.EXEName, "PO_Checked", "PO_Check", Trim(txtChecked.Text)
    SaveSetting App.EXEName, "PO_Approved", "PO_Approve", Trim(txtApproved.Text)
    'Me.Caption = "PURCHASE ORDER - BROWSE"
    BROWSER sPONumber, "is_LOAD"
ElseIf TRANSACTIONTYPE = is_EDITTING And _
Toolbar1.Buttons(7).Enabled = True Then
    If Trim(txtPONumber.Text) = "" Then
        MsgBox "Please Supply PO # !            ", vbCritical, "Error..."
        txtPONumber.SetFocus
        HTEXT txtPONumber
        Exit Sub
    End If
    If IsDate(txtPODate.Text) = False Then
        MsgBox "Please Supply A Valid Date!         ", vbCritical, "Error..."
        txtPODate.SetFocus
        HTEXT txtPODate
        Exit Sub
    End If
    If Trim(txtRefNo.Text) = "" Then
        MsgBox "Please Supply Requisition #!            ", vbCritical, "Error..."
        txtRefNo.SetFocus
        HTEXT txtRefNo
        Exit Sub
    End If
    If Trim(txtTerms.Text) = "" Then
        MsgBox "Please Supply Terms!            ", vbCritical, "Error..."
        txtTerms.SetFocus
        HTEXT txtTerms
        Exit Sub
    End If
    If CDbl(iSupplier) = 0 Then
        MsgBox "Please Supply Supplier!         ", vbCritical, "Error..."
        txtSuppCode.SetFocus
        Exit Sub
    End If
    If Trim(txtRequested.Text) = "" Then
        MsgBox "Please Suply Requested Person!          ", vbCritical, "Error..."
        txtRequested.SetFocus
        HTEXT txtRequested
        Exit Sub
    End If
    If Trim(txtChecked.Text) = "" Then
        MsgBox "Please Supply Person Verified this P.O. !           ", vbCritical, "Error..."
        txtChecked.SetFocus
        HTEXT txtChecked
        Exit Sub
    End If
    If Trim(txtApproved.Text) = "" Then
        MsgBox "Please Supply Approving Officer!            ", vbCritical, "Error..."
        txtApproved.SetFocus
        HTEXT txtApproved
        Exit Sub
    End If
    If CDbl(iDept) = 0 Then
        MsgBox "Please Supply Requesting Department!            ", vbCritical, "Error..."
        cmbDeptName.SetFocus
        Exit Sub
    End If
    With lstDetail.ListItems
    If Trim(.Item(1).SubItems(1)) <> "" Then
        For k = 1 To .Count
            lstDetail.ListItems(k).EnsureVisible
            lstDetail.ListItems(k).Selected = True
            If CDbl(.Item(k).SubItems(2)) = 1 Then
                If CHECK_ITEM(.Item(k).Text) <> CLng(txtSuppKey.Text) Then GoTo PX:
            End If
        Next k
    End If
    On Error GoTo PG:
    strDisc = IIf(Trim(txtDisc1.Text) = "", "", Replace(Trim(txtDisc1.Text), "%", "") & "%")
    ConnOmega.Execute "UPDATE tbl_Inv_PO" & _
                      " SET PONumber = '" & Trim(txtPONumber.Text) & "', " & _
                      " PODate = '" & FormatDateTime(txtPODate.Text, vbShortDate) & "', " & _
                      " PRNumber = '" & Trim(txtRefNo.Text) & "', " & _
                      " SuppKey = " & iSupplier & ", " & _
                      " Terms = '" & FORMATSQL(Trim(txtTerms.Text)) & "', " & _
                      " Remarks = '" & FORMATSQL(Trim(txtRemarks.Text)) & "', " & _
                      " TotalCost = " & CDbl(lblTotalCost.Caption) & ", " & _
                      " RequestBy = '" & FORMATSQL(Trim(txtRequested.Text)) & "', " & _
                      " CheckedBy = '" & FORMATSQL(Trim(txtChecked.Text)) & "', " & _
                      " ApprovedBy = '" & FORMATSQL(Trim(txtApproved.Text)) & "', " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                      " DeptKey = " & iDept & ", " & _
                      " Discount = '" & strDisc & "'" & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    ConnOmega.Execute "DELETE FROM tbl_Inv_PODet WHERE (POKey = " & Statusbar1.Panels(1).Text & ")"
    dblTotalQty = 0
    dblTotalCost = 0
    dblTotalNetCost = 0
    If Trim(.Item(1).SubItems(1)) <> "" Then
        For i = 1 To .Count
            If Trim(.Item(i).Text) <> "" Then
                lstDetail.ListItems(i).EnsureVisible
                lstDetail.ListItems(i).Selected = True
                If Trim(strDisc) <> "" Then
                    .Item(i).SubItems(9) = Format(CDbl(.Item(i).SubItems(8)) * ((100 - Val(Trim(strDisc))) / 100), "#,##0.00")
                Else
                    .Item(i).SubItems(9) = .Item(i).SubItems(8)
                End If
                dblTotalQty = dblTotalQty + CDbl(.Item(i).SubItems(11))
                dblTotalCost = dblTotalCost + (CDbl(.Item(i).SubItems(8)) * CDbl(.Item(i).SubItems(4)))
                dblTotalNetCost = dblTotalNetCost + (CDbl(.Item(i).SubItems(9)) * CDbl(.Item(i).SubItems(4)))
                ConnOmega.Execute "INSERT INTO tbl_Inv_PODet" & _
                                  " (POKey, Line, ItemKey, Qty, Unit, Cost, O_Qty, O_Cost, Discount, NetCost, Item_FA) " & _
                                  " VALUES(" & Statusbar1.Panels(1).Text & ", " & i & ",  " & _
                                  " " & .Item(i).Text & ", " & CDbl(.Item(i).SubItems(4)) & ", " & _
                                  " '" & CStr(FORMATSQL(Trim((.Item(i).SubItems(5))))) & "', " & _
                                  " " & CDbl(.Item(i).SubItems(8)) & ", " & CDbl(.Item(i).SubItems(11)) & ", " & _
                                  " " & CDbl(.Item(i).SubItems(12)) & ", '" & strDisc & "', " & _
                                  " " & CDbl(.Item(i).SubItems(9)) & ", " & CDbl(.Item(i).SubItems(2)) & ")"
            End If
        Next i
    End If
    ConnOmega.Execute "UPDATE tbl_Inv_PO SET TotalCost = " & CDbl(dblTotalCost) & ", TotalNetCost = " & CDbl(dblTotalNetCost) & ", TotalQty = " & CDbl(dblTotalQty) & " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    End With
    LOCKTEXT True
    txtPONumber.SetFocus
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    SaveSetting App.EXEName, "PO_Requested", "PO_Req", Trim(txtRequested.Text)
    SaveSetting App.EXEName, "PO_Checked", "PO_Check", Trim(txtChecked.Text)
    SaveSetting App.EXEName, "PO_Approved", "PO_Approve", Trim(txtApproved.Text)
    'Me.Caption = "PURCHASE ORDER - BROWSE"
    BROWSER Trim(txtPONumber.Text), "is_LOAD"
End If
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
Exit Sub
PX:
MsgBox "THIS ITEM NOT BELONG TO THIS SUPPLIER!      " & vbCrLf & _
       "RECORD CANNOT BE SAVED!                     ", vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If picSLine.Visible = True Then Exit Sub
CLEARTEXT
txtPONumber.Locked = False
TOOLBARFUNC 3
TRANSACTIONTYPE = is_FINDING
'Me.Caption = "PURCHASE ORDER - FIND"
txtPONumber.SetFocus
End Sub

Private Sub PRESS_F8()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If picSLine.Visible = True Then Exit Sub
If AccessRights("Purchase Order", "Post") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If imgPosted.Visible = True Then MsgBox "ALREADY POSTED!             ", vbCritical, "Posted": Exit Sub
If MsgBox("ARE YOU SURE YOU WANT TO POST THIS TRANSACTION?          ", vbInformation + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub

s = "SELECT Printed" & _
    " From tbl_Inv_PO " & _
    " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs!Printed = 0 Then
    MsgBox "NOT YET PRINTED!            ", vbCritical, "Printed"
    rs.Close
    Exit Sub
End If
rs.Close
ConnOmega.Execute "UPDATE tbl_Inv_PO SET Posted = 1 WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
BROWSER GetSetting(App.EXEName, "PONumber", "PONum", ""), "is_LOAD"
End Sub

Private Sub PRESS_F9()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("Purchase Order", "Print") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
    
s = "SELECT Printed" & _
    " From tbl_Inv_PO " & _
    " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    If rs!Printed = 1 Then
        If MsgBox("ALREADY PRINTED!             " & vbCrLf & _
                   "" & vbCrLf & _
                   "Print Another Copy?         ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then
            rs.Close
            Exit Sub
        End If
        frmPrinter.LastPage = 1
        frmPrinter.PRINT_TRANSACTION = 1
        frmPrinter.picPageRange.Enabled = False
        frmPrinter.txtCopies.Locked = True
        frmPrinter.picCPI.Enabled = False
        frmPrinter.Show 1
    Else
        frmPrinter.LastPage = 1
        frmPrinter.PRINT_TRANSACTION = 1
        frmPrinter.picPageRange.Enabled = False
        frmPrinter.txtCopies.Locked = True
        frmPrinter.picCPI.Enabled = False
        frmPrinter.Show 1
    End If
End If
If rs.State = adStateOpen Then rs.Close
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    Unload Me
Else
    If is_DET_FOCUS = 1 Then
        With lstDetail.ListItems
        If TRANS_DETAIL = is_ADDING Then
            If .Count > 1 Then
                .Remove .Count
                ROW = .Count
            Else
                .Item(1).Text = " "
                .Item(1).SubItems(1) = " "
                .Item(1).SubItems(2) = " "
                .Item(1).SubItems(3) = " "
                .Item(1).SubItems(4) = " "
                .Item(1).SubItems(5) = " "
                .Item(1).SubItems(6) = " "
                .Item(1).SubItems(7) = " "
                .Item(1).SubItems(8) = " "
                .Item(1).SubItems(9) = " "
                .Item(1).SubItems(10) = " "
                ROW = 1
            End If
            picBody.Enabled = True
            picSLine.Visible = False
            TRANS_DETAIL = is_DET_REFRESH
            lstDetail.ListItems(ROW).EnsureVisible
            lstDetail.ListItems(ROW).Selected = True
            lstDetail.SetFocus
        ElseIf TRANS_DETAIL = is_EDITTING Then
            .Item(ROW).Text = txtItemKey1.Text
            .Item(ROW).SubItems(2) = txtType1.Text
            .Item(ROW).SubItems(3) = txtTypeDesc1.Text
            .Item(ROW).SubItems(4) = txtQty1.Text
            .Item(ROW).SubItems(5) = txtUnit1.Text
            .Item(ROW).SubItems(6) = txtItemCode1.Text
            .Item(ROW).SubItems(7) = txtItemDesc1.Text
            .Item(ROW).SubItems(8) = txtCost1.Text
            .Item(ROW).SubItems(9) = txtNetCost1.Text
            .Item(ROW).SubItems(10) = txtTotalNetCost1.Text
            picBody.Enabled = True
            picSLine.Visible = False
            TRANS_DETAIL = is_DET_REFRESH
            lstDetail.ListItems(ROW).EnsureVisible
            lstDetail.ListItems(ROW).Selected = True
            lstDetail.SetFocus
        Else
            txtPONumber.SetFocus
            is_DET_FOCUS = 0
            If imgPosted.Visible = False Then TOOLBARFUNC 2 Else TOOLBARFUNC 3
        End If
        End With
    Else
        CLEARTEXT
        LOCKTEXT True
        TOOLBARFUNC 1
        TRANSACTIONTYPE = is_REFRESH
        'Me.Caption = "PURCHASE ORDER - BROWSE"
        BROWSER GetSetting(App.EXEName, "PONumber", "PONum", ""), "is_LOAD"
    End If
End If
End Sub


Private Sub CUSTOMIZE_DETAIL()
With lstDetail.ListItems
    .Clear
    Set x = .Add()
    x.Text = " "
    x.SubItems(1) = " "
    x.SubItems(2) = " "
    x.SubItems(3) = " "
    x.SubItems(4) = " "
    x.SubItems(5) = " "
    x.SubItems(6) = " "
    x.SubItems(7) = " "
    x.SubItems(8) = " "
    x.SubItems(9) = " "
    x.SubItems(10) = " "
End With
End Sub

Public Sub CLEARTEXT()
iSupplier = 0
iDept = 0
'locDept = 0
txtPONumber.Text = ""
txtPODate.Text = ""
txtRefNo.Text = ""
txtTerms.Text = ""
txtSuppKey.Text = ""
txtSuppCode.Text = ""
'DataSupplier.BoundText = 0
txtAddress.Text = ""
txtTelNo.Text = ""
txtFaxNo.Text = ""
'DataDept.BoundText = 0
txtRemarks.Text = ""
txtRequested.Text = ""
txtChecked.Text = ""
txtApproved.Text = ""
lblTotalCost.Caption = "0.00"
lblTotalNetCost.Caption = "0.00"
txtDisc1.Text = ""
'txtDisc2.Text = ""
'txtDisc3.Text = ""
txtVAT.Text = ""
cmbSuppName.Text = ""
cmbDeptName.Text = ""
cmbSuppName.ListIndex = -1
cmbDeptName.ListIndex = -1
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
Toolbar1.Buttons(19).Caption = "Post"
Toolbar1.Buttons(19).ToolTipText = "POST (F8)"
imgPosted.Visible = False
CUSTOMIZE_DETAIL
End Sub

Private Sub CLEAR_DETAIL()
cmbType.ListIndex = -1
txtType.Text = ""
txtItemKey.Text = ""
txtOQty.Text = ""
txtOCost.Text = ""
txtQty.Text = ""
txtUnit.Text = ""
cmbUnit.Clear
txtItemCode.Text = ""
txtItemDesc.Text = ""
txtCost.Text = ""
txtTotalCost.Text = ""
End Sub

Public Sub LOCKTEXT(bln As Boolean)
txtPONumber.Locked = bln
txtPODate.Locked = bln
txtRefNo.Locked = bln
txtTerms.Locked = bln
txtSuppCode.Locked = bln
txtAddress.Locked = bln
txtTelNo.Locked = bln
txtFaxNo.Locked = bln
txtRemarks.Locked = bln
txtRequested.Locked = bln
txtChecked.Locked = bln
txtApproved.Locked = bln
txtDisc1.Locked = bln
txtVAT.Locked = bln
cmbSuppName.Locked = bln
cmbDeptName.Locked = bln
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

Private Sub cmbDeptName_Click()
If TRANSACTIONTYPE = is_ADDING Or TRANSACTIONTYPE = is_EDITTING Then
    If cmbDeptName.ListIndex = -1 Then Exit Sub
    iDept = cmbDeptName.ItemData(cmbDeptName.ListIndex)
End If
End Sub

Private Sub cmbDeptName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtDisc1.SetFocus
End Sub

Private Sub cmbSuppName_Click()
If TRANSACTIONTYPE = is_ADDING Or TRANSACTIONTYPE = is_EDITTING Then
    If cmbSuppName.ListIndex = -1 Then Exit Sub
    s = "SELECT PK, SupplierCode, SupplierName, " & _
        " ltrim(rtrim(Address1 + (CASE ltrim(rtrim(Address2)) WHEN '' THEN '' ELSE ',  ' + Address2 END) + (CASE ltrim(rtrim(Address3)) WHEN '' THEN '' ELSE ', ' + Address3 END))) AS Address,  " & _
        " TelNo, FaxNo" & _
        " From tbl_Inv_Supplier " & _
        " WHERE (PK = " & cmbSuppName.ItemData(cmbSuppName.ListIndex) & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        txtSuppKey.Text = rs!PK
        iSupplier = rs!PK
        txtSuppCode.Text = rs!SupplierCode
        txtAddress.Text = rs!Address
        txtTelNo.Text = rs!TelNo
        txtFaxNo.Text = rs!FaxNo
    End If
    rs.Close
End If
End Sub

Private Sub cmbType_Click()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    If cmbType.ListIndex = -1 Then Exit Sub
    txtType.Text = cmbType.ListIndex + 1
    With lstDetail.ListItems
        .Item(ROW).SubItems(3) = cmbType.List(cmbType.ListIndex)
    End With
End If
End Sub

Private Sub cmbType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtQty.SetFocus
End Sub

Private Sub cmbUnit_Click()
If cmbUnit.ListIndex = -1 Then Exit Sub
If RETURNTEXTVALUE(txtType) = 1 Then
    s = "SELECT ConUnit, ConUnit2" & _
        " From tbl_Inv_Items " & _
        " WHERE (ItemCode = '" & Trim(txtItemCode.Text) & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        txtUnit.Text = cmbUnit.Text
        If cmbUnit.ListIndex = 0 Then
            txtOQty.Text = RETURNTEXTVALUE(txtQty)
            txtOCost.Text = RETURNTEXTVALUE(txtCost)
        End If
        If cmbUnit.ListIndex = 1 Then
            txtOQty.Text = RETURNTEXTVALUE(txtQty) * CDbl(rs!ConUnit)
            If RETURNTEXTVALUE(txtCost) > 0 Then
                txtOCost.Text = RETURNTEXTVALUE(txtCost) / rs!ConUnit 'RETURNTEXTVALUE(txtOQty)
            End If
        End If
    End If
    rs.Close
Else
    txtUnit.Text = cmbUnit.Text
End If
End Sub

Private Sub cmbUnit_GotFocus()
'cmbUnit.Text = txtUnit.Text
End Sub

Private Sub cmbUnit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtCost.SetFocus
End If
End Sub


Private Sub Form_Activate()
If TRANSACTIONTYPE = is_REFRESH Then BROWSER GetSetting(App.EXEName, "PONumber", "PONum", ""), "is_LOAD"
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
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "PONumber", "PONum", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "PONumber", "PONum", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "PONumber", "PONum", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "PONumber", "PONum", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
cmbSuppName.Clear
s = "SELECT PK, SupplierCode, SupplierName" & _
    " FROM tbl_Inv_Supplier " & _
    " ORDER BY SupplierName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbSuppName.AddItem rs!SupplierName ' & " - " & rs!SupplierCode
    cmbSuppName.ItemData(cmbSuppName.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close

cmbDeptName.Clear
's = "SELECT PK, DepartmentCode, DepartmentName" & _
    " FROM tbl_Personnel_Department " & _
    " ORDER BY DepartmentName"
s = "SELECT PK, Code as DepartmentCode, " & _
    " DeptName as DepartmentName" & _
    " FROM tbl_GL_Department " & _
    " WHERE (Issuance = 1) " & _
    " ORDER BY DeptName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    'cmbDeptName.AddItem UCase(rs!DepartmentCode & " - " & rs!DepartmentName)
    cmbDeptName.AddItem UCase(rs!DepartmentName)
    cmbDeptName.ItemData(cmbDeptName.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close

With cmbType
    .Clear
    .AddItem "Item"
    .AddItem "Fixed Asset"
    '.AddItem "Job Order"
End With

CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANS_DETAIL = is_DET_REFRESH
BROWSER GetSetting(App.EXEName, "PONumber", "PONum", ""), "is_LOAD"
If Trim(txtPONumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "PONumber", "PONum", ""), "is_HOME"
is_DET_FOCUS = 0
'Me.Caption = "PURCHASE ORDER - BROWSE"

tmp = SetWindowLong(txtPONumber.hwnd, GWL_STYLE, GetWindowLong(txtPONumber.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPODate.hwnd, GWL_STYLE, GetWindowLong(txtPODate.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtRefNo.hwnd, GWL_STYLE, GetWindowLong(txtRefNo.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtTerms.hwnd, GWL_STYLE, GetWindowLong(txtTerms.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSuppCode.hwnd, GWL_STYLE, GetWindowLong(txtSuppCode.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtAddress.hwnd, GWL_STYLE, GetWindowLong(txtAddress.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtTelNo.hwnd, GWL_STYLE, GetWindowLong(txtTelNo.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtFaxNo.hwnd, GWL_STYLE, GetWindowLong(txtFaxNo.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtRemarks.hwnd, GWL_STYLE, GetWindowLong(txtRemarks.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtRequested.hwnd, GWL_STYLE, GetWindowLong(txtRequested.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtChecked.hwnd, GWL_STYLE, GetWindowLong(txtChecked.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtApproved.hwnd, GWL_STYLE, GetWindowLong(txtApproved.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picSLine.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub


Private Sub lstDetail_GotFocus()
If TRANSACTIONTYPE = is_REFRESH Then
    If Statusbar1.Panels(1) <> "" Then
        If AccessRights("Purchase Order", "Edit") = False Then
            TOOLBARFUNC 3
            TRANSACTIONTYPE = is_EDITTING
        Else
            If imgPosted.Visible = True Then
                TOOLBARFUNC 3
                TRANSACTIONTYPE = is_EDITTING
            Else
                If Trim(lstDetail.ListItems.Item(1).SubItems(1)) <> "" Then
                    TOOLBARFUNC 5
                Else
                    TOOLBARFUNC 4
                End If
                LOCKTEXT False
                TRANSACTIONTYPE = is_EDITTING
            End If
            'Me.Caption = "PURCHASE ORDER - EDIT"
            is_DET_FOCUS = 1
            ROW = lstDetail.SelectedItem.Index
            TRANS_DETAIL = is_DET_REFRESH
        End If
    End If
Else
    If Trim(lstDetail.ListItems.Item(1).SubItems(1)) <> "" Then
        TOOLBARFUNC 5
    Else
        TOOLBARFUNC 4
    End If
    is_DET_FOCUS = 1
    ROW = lstDetail.SelectedItem.Index
    TRANS_DETAIL = is_DET_REFRESH
End If
End Sub

Private Sub lstDetail_ItemClick(ByVal Item As MSComctlLib.ListItem)
ROW = lstDetail.SelectedItem.Index
End Sub

Private Sub lstDetail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
    If ROW = 1 Then
        txtFaxNo.SetFocus
'        cmbDept.SetFocus
    End If
ElseIf KeyCode = vbKeyDown Then
    If ROW = lstDetail.ListItems.Count Then
        txtRemarks.SetFocus
    End If
End If
End Sub

Private Sub lstDetail_LostFocus()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If TRANS_DETAIL = is_DET_REFRESH Then
        If imgPosted.Visible = False Then
            TOOLBARFUNC 2
        Else
            TOOLBARFUNC 3
        End If
        is_DET_FOCUS = 0
    End If
End If
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":           PRESS_INSERT
    Case "Edit":          PRESS_F2
    Case "Delete":        PRESS_DELETE
    Case "First"
        Select Case Toolbar1.Buttons(7).Caption
            Case "Save":  PRESS_F5
            Case "First": BROWSER GetSetting(App.EXEName, "PONumber", "PONum", ""), "is_HOME"
        End Select
    Case "Back"
        Select Case Toolbar1.Buttons(9).Caption
            Case "Undo":  PRESS_ESCAPE
            Case "Back":  BROWSER GetSetting(App.EXEName, "PONumber", "PONum", ""), "is_PAGEUP"
        End Select
    Case "Next":          BROWSER GetSetting(App.EXEName, "PONumber", "PONum", ""), "is_PAGEDOWN"
    Case "Last":          BROWSER GetSetting(App.EXEName, "PONumber", "PONum", ""), "is_END"
    Case "Find":          PRESS_F6
    Case "Post":          PRESS_F8
    Case "Print":         PRESS_F9
    Case "Close":         PRESS_ESCAPE
End Select
End Sub

Private Sub txtAddress_GotFocus()
HTEXT txtAddress
End Sub

Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtTelNo.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtSuppCode.SetFocus
End If
End Sub

Private Sub txtApproved_GotFocus()
HTEXT txtApproved
End Sub

Private Sub txtApproved_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtPONumber.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtChecked.SetFocus
End If
End Sub

Private Sub txtChecked_GotFocus()
HTEXT txtChecked
End Sub

Private Sub txtChecked_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtApproved.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtRequested.SetFocus
End If
End Sub

Private Sub txtCost_Change()
lstDetail.ListItems.Item(ROW).SubItems(8) = Format(RETURNTEXTVALUE(txtCost), "#,##0.00")
If RETURNTEXTVALUE(txtItemKey) > 0 Then
    If RETURNTEXTVALUE(txtType) = 1 Then
        s = "SELECT Unit, ConUnit, Unit2, ConUnit2" & _
            " From tbl_Inv_Items " & _
            " WHERE (ItemCode = '" & Trim(txtItemCode.Text) & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            txtUnit.Text = cmbUnit.Text
            If cmbUnit.ListIndex = 0 Then
                txtOQty.Text = RETURNTEXTVALUE(txtQty)
                txtOCost.Text = RETURNTEXTVALUE(txtCost)
            End If
            If cmbUnit.ListIndex = 1 Then
                txtOQty.Text = RETURNTEXTVALUE(txtQty) * CDbl(rs!ConUnit)
                txtOCost.Text = RETURNTEXTVALUE(txtCost) / CDbl(rs!ConUnit)
            End If
        End If
        rs.Close
    Else
        txtOQty.Text = RETURNTEXTVALUE(txtQty)
        txtOCost.Text = RETURNTEXTVALUE(txtCost)
    End If
    txtTotalCost.Text = Format(RETURNTEXTVALUE(txtCost) * RETURNTEXTVALUE(txtQty), "#,##0.00")
    If Trim(txtDisc1.Text) <> "" Then
        txtNetCost.Text = Format(RETURNTEXTVALUE(txtCost) * (100 - CDbl(Val(Trim(txtDisc1.Text)))) / 100, "#,##0.00")
    Else
        txtNetCost.Text = Format(RETURNTEXTVALUE(txtCost), "#,##0.00")
    End If
End If
End Sub

Private Sub txtCost_GotFocus()
HTEXT txtCost
End Sub

Private Sub txtCost_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If cmbType.ListIndex = -1 Then MsgBox "Please Select Type!                    ", vbCritical, "Error...": cmbType.SetFocus: Exit Sub
    If cmbUnit.ListIndex = -1 Then MsgBox "Please Select Unit!                   ", vbCritical, "Error...": cmbUnit.SetFocus: Exit Sub
    picBody.Enabled = True
    picSLine.Visible = False
    TRANS_DETAIL = is_DET_REFRESH
    TOOLBARFUNC 5
    lstDetail.SetFocus
ElseIf KeyCode = vbKeyUp Then
    cmbUnit.SetFocus
'    txtItemCode.SetFocus
End If
End Sub

Private Sub txtDisc1_GotFocus()
txtDisc1.Text = Val(txtDisc1.Text) ' IIf(RETURNTEXTVALUE(txtDisc1) = 0, "", Mid(txtDisc1.Text, 1, Len(txtDisc1.Text) - 1))
txtDisc1.MaxLength = 2
txtDisc1.Text = IIf(RETURNTEXTVALUE(txtDisc1) = 0, "", RETURNTEXTVALUE(txtDisc1))
HTEXT txtDisc1
End Sub

Private Sub txtDisc1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtVAT.SetFocus
ElseIf KeyCode = vbKeyUp Then
    cmbDeptName.SetFocus
End If
End Sub

Private Sub txtDisc1_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtDisc1_LostFocus()
txtDisc1.MaxLength = 3
txtDisc1.Text = IIf(Trim(txtDisc1.Text) = "", "", Replace(Trim(txtDisc1.Text), "%", "") & "%")
End Sub

Private Sub txtFaxNo_GotFocus()
HTEXT txtFaxNo
End Sub

Private Sub txtFaxNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    cmbDeptName.SetFocus
'    cmbDept.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtTelNo.SetFocus
End If
End Sub

Private Sub txtItemCode_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    If cmbType.ListIndex = -1 Then MsgBox "Please Select Type!               ", vbCritical, "Error...": cmbType.SetFocus: Exit Sub
End If
End Sub

Private Sub txtItemCode_GotFocus()
HTEXT txtItemCode
End Sub

Private Sub txtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    cmbUnit.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtQty.SetFocus
End If
If KeyCode = vbKeyInsert Then
    If IsLoaded(frmInvItems) Then frmInvItems.ZOrder 0 Else frmInvItems.Show
    gbl_Item_Module = "Purchase Order"
    With frmInvItems
        .CLEARTEXT
        .TRANSACTIONTYPE = 1
        .TOOLBARFUNC 2
        .LOCKTEXT False
        s = "SELECT PK, SupplierCode, SupplierName " & _
            " From dbo.tbl_Inv_Supplier " & _
            " WHERE (PK = " & iSupplier & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            .iSupplier = rs!PK
            .txtSuppKey.Text = rs!PK
            .txtSuppCode.Text = rs!SupplierCode
            .cmbSuppName.Text = rs!SupplierName
            t = "SELECT TOP 1 ItemCode" & _
                " From tbl_Inv_Items " & _
                " Where (SuppKey = " & rs!PK & ") " & _
                " ORDER BY ItemCode DESC"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                .txtItemCode.Text = CStr(CDbl(rt!ItemCode) + 1)
            Else
                .txtItemCode.Text = CStr(CDbl(rs!SupplierCode)) & "001"
            End If
            rt.Close
        End If
        rs.Close
        .Caption = "Items - New"
        .txtItemDesc.SetFocus
    End With
End If
End Sub

Private Sub txtItemCode_LostFocus()
If picSLine.Visible = False Then Exit Sub
If Trim(txtItemCode.Text) = "" Then Exit Sub
Select Case RETURNTEXTVALUE(txtType)
    Case 1
        s = "SELECT PK, ItemCode, ItemDesc, " & _
            " Unit, Unit2, Cost, SuppKey" & _
            " From tbl_Inv_Items " & _
            " WHERE (ItemCode = '" & Trim(txtItemCode.Text) & "')"
    Case 2
        s = "SELECT PK, Code as ItemCode, " & _
            " Description as ItemDesc, Unit" & _
            " FROM tbl_FA_Items " & _
            " WHERE (Code = '" & Trim(txtItemCode.Text) & "')"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then

    If RETURNTEXTVALUE(txtType) = 1 Then
        If rs!SuppKey <> iSupplier Then
            MsgBox "THIS ITEM NOT BELONG TO THIS SUPPLIER!      ", vbCritical, "Error..."
            If rs.State = adStateOpen Then rs.Close
            txtItemCode.SetFocus
            HTEXT txtItemCode
            Exit Sub
        End If
    End If
    
    txtItemKey.Text = rs!PK
    
    With cmbUnit
        .Clear
        .AddItem rs!Unit
        If RETURNTEXTVALUE(txtType) = 1 Then
            If Trim(rs!Unit2) <> "" Then
                .AddItem rs!Unit2
            End If
        End If
    End With
    
    txtItemCode.Text = rs!ItemCode
    txtItemDesc.Text = rs!ItemDesc
    With lstDetail.ListItems
        .Item(ROW).Text = rs!PK
        .Item(ROW).SubItems(6) = rs!ItemCode
        .Item(ROW).SubItems(7) = rs!ItemDesc
    End With
    
Else
    txtItemKey.Text = "0"
    txtUnit.Text = ""
    txtItemCode.Text = ""
    txtItemDesc.Text = ""
    With lstDetail.ListItems
        .Item(ROW).Text = " "
        .Item(ROW).SubItems(5) = " "
        .Item(ROW).SubItems(6) = " "
        .Item(ROW).SubItems(7) = " "
    End With
End If
rs.Close
End Sub

Private Sub txtItemDesc_GotFocus()
HTEXT txtItemDesc
End Sub

Private Sub txtItemKey_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(ROW).Text = txtItemKey
    End With
End If
End Sub

Private Sub txtNetCost_Change()
lstDetail.ListItems.Item(ROW).SubItems(9) = Format(RETURNTEXTVALUE(txtNetCost), "#,##0.00")
'If RETURNTEXTVALUE(txtItemKey) > 0 Then
'    Dim s As String
'    Dim rs As New ADODB.Recordset
'    s = "SELECT ConUnit, ConUnit2" & _
'        " From tbl_Inv_Items " & _
'        " WHERE (ItemCode = '" & Trim(txtItemCode.Text) & "')"
'    rs.Open s, ConnOmega
'    If rs.RecordCount > 0 Then
'        txtUnit.Text = cmbUnit.Text
'        If cmbUnit.ListIndex = 0 Then
'            txtOQty.Text = RETURNTEXTVALUE(txtQty)
'            txtOCost.Text = RETURNTEXTVALUE(txtCost)
'        End If
'        If cmbUnit.ListIndex = 1 Then
'            txtOQty.Text = RETURNTEXTVALUE(txtQty) * CDbl(rs!ConUnit)
'            txtOCost.Text = RETURNTEXTVALUE(txtCost) / CDbl(rs!ConUnit)
'        End If
'    End If
'    rs.Close
    txtTotalNetCost.Text = Format(RETURNTEXTVALUE(txtNetCost) * RETURNTEXTVALUE(txtQty), "#,##0.00")
'End If
End Sub

Private Sub txtOCost_Change()
lstDetail.ListItems.Item(ROW).SubItems(12) = RETURNTEXTVALUE(txtOCost)
End Sub

Private Sub txtOQty_Change()
lstDetail.ListItems.Item(ROW).SubItems(11) = RETURNTEXTVALUE(txtOQty)
End Sub

Private Sub txtPODate_GotFocus()
HTEXT txtPODate
End Sub

Private Sub txtPODate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtRefNo.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtPONumber.SetFocus
End If
End Sub

Private Sub txtPODate_LostFocus()
If Trim(txtPODate.Text) <> "" Then
    If IsDate(txtPODate.Text) Then
        txtPODate.Text = Format(FormatDateTime(txtPODate.Text, vbShortDate), "mm/dd/yyyy")
    Else
        MsgBox "PLEASE SUPPLY A VALID DATE!             ", vbCritical, "Error..."
        txtPODate.SetFocus
        HTEXT txtPODate
        Exit Sub
    End If
End If
End Sub

Private Sub txtPONumber_GotFocus()
HTEXT txtPONumber
End Sub

Private Sub txtPONumber_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtPODate.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtApproved.SetFocus
End If
End Sub

Private Sub txtPONumber_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtPONumber_LostFocus()
If TRANSACTIONTYPE = is_FINDING Then
    v = "SELECT tbl_Inv_PO.*" & _
        " FROM tbl_Inv_PO " & _
        " WHERE (PONumber = '" & Format(RETURNTEXTVALUE(txtPONumber), "0000000#") & "')"
    If rv.State = adStateOpen Then ru.Close
    rv.Open v, ConnOmega
    If rv.RecordCount > 0 Then
        BROWSER rv!PONumber, "is_LOAD"
        LOCKTEXT True
        TOOLBARFUNC 1
        TRANSACTIONTYPE = is_REFRESH
        'Me.Caption = "PURCHASE ORDER - BROWSE"
    Else
        MsgBox "PO # '" & Format(RETURNTEXTVALUE(txtPONumber), "0000000#") & "' NOT FOUND!               ", vbCritical, "Error..."
        txtPONumber.Text = Format(RETURNTEXTVALUE(txtPONumber), "0000000#")
        txtPONumber.SetFocus
        HTEXT txtPONumber
        rv.Close
        Exit Sub
    End If
    rv.Close
ElseIf TRANSACTIONTYPE = is_ADDING Then
    txtPONumber.Text = Format(RETURNTEXTVALUE(txtPONumber), "0000000#")
End If
End Sub

Private Sub txtQty_Change()
lstDetail.ListItems.Item(ROW).SubItems(4) = Format(RETURNTEXTVALUE(txtQty), "#,##0.00")
If RETURNTEXTVALUE(txtItemKey) > 0 Then
    s = "SELECT ConUnit, ConUnit2" & _
        " From tbl_Inv_Items " & _
        " WHERE (ItemCode = '" & Trim(txtItemCode.Text) & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        txtUnit.Text = cmbUnit.Text
        If cmbUnit.ListIndex = 0 Then
            txtOQty.Text = RETURNTEXTVALUE(txtQty)
            txtOCost.Text = RETURNTEXTVALUE(txtCost)
        End If
        If cmbUnit.ListIndex = 1 Then
            txtOQty.Text = RETURNTEXTVALUE(txtQty) * CDbl(rs!ConUnit)
            txtOCost.Text = RETURNTEXTVALUE(txtCost) / CDbl(rs!ConUnit)
        End If
    End If
    rs.Close
    txtTotalCost.Text = Format(RETURNTEXTVALUE(txtCost) * RETURNTEXTVALUE(txtQty), "#,##0.00")
End If
End Sub

Private Sub txtQty_GotFocus()
HTEXT txtQty
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtItemCode.SetFocus
'    txtUnit.SetFocus
End If
End Sub

Private Sub txtRefNo_GotFocus()
HTEXT txtRefNo
End Sub

Private Sub txtRefNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtTerms.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtPODate.SetFocus
End If
End Sub

Private Sub txtRemarks_GotFocus()
HTEXT txtRemarks
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtRequested.SetFocus
ElseIf KeyCode = vbKeyUp Then
    lstDetail.SetFocus
End If
End Sub

Private Sub txtRequested_GotFocus()
HTEXT txtRequested
End Sub

Private Sub txtRequested_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtChecked.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtRemarks.SetFocus
End If
End Sub

Private Sub txtSuppCode_GotFocus()
HTEXT txtSuppCode
End Sub

Private Sub txtSuppCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    cmbSuppName.SetFocus
'    txtAddress.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtTerms.SetFocus
End If
End Sub

Private Sub txtSuppCode_LostFocus()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If Trim(txtSuppCode.Text) <> "" Then
        s = "SELECT PK, SupplierCode, SupplierName, " & _
            " ltrim(rtrim(Address1 + (CASE ltrim(rtrim(Address2)) WHEN '' THEN '' ELSE ',  ' + Address2 END) + (CASE ltrim(rtrim(Address3)) WHEN '' THEN '' ELSE ', ' + Address3 END))) AS Address,  " & _
            " TelNo, FaxNo" & _
            " From tbl_Inv_Supplier " & _
            " WHERE (SupplierCode = '" & Trim(txtSuppCode.Text) & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            txtSuppKey.Text = rs!PK
            iSupplier = rs!PK
            txtSuppCode.Text = rs!SupplierCode
            cmbSuppName.Text = rs!SupplierName
            txtAddress.Text = rs!Address
            txtTelNo.Text = rs!TelNo
            txtFaxNo.Text = rs!FaxNo
        Else
            txtSuppKey.Text = "0"
            iSupplier = 0
            txtSuppCode.Text = ""
            cmbSuppName.Text = ""
            txtAddress.Text = ""
            txtTelNo.Text = ""
            txtFaxNo.Text = ""
        End If
        rs.Close
    End If
End If
End Sub

Private Sub txtTelNo_GotFocus()
HTEXT txtTelNo
End Sub

Private Sub txtTelNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtFaxNo.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtAddress.SetFocus
End If
End Sub

Private Sub txtTerms_GotFocus()
HTEXT txtTerms
End Sub

Private Sub txtTerms_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtSuppCode.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtRefNo.SetFocus
End If
End Sub

Private Sub txtTotalCost_Change()

With lstDetail.ListItems
    DoEvents
    .Item(ROW).SubItems(13) = Format(RETURNTEXTVALUE(txtTotalCost), "#,##0.00")
    a = 0
    For i = 1 To .Count
        DoEvents
        a = a + IIf(IsNumeric(.Item(i).SubItems(13)), CDbl(.Item(i).SubItems(13)), 0)
    Next i
    lblTotalCost.Caption = Format(a, "#,##0.00")
End With
End Sub

Private Sub txtTotalCost_GotFocus()
HTEXT txtTotalCost
End Sub

Private Sub txtTotalNetCost_Change()
Dim i, a
With lstDetail.ListItems
    DoEvents
    .Item(ROW).SubItems(10) = Format(RETURNTEXTVALUE(txtTotalNetCost), "#,##0.00")
    a = 0
    For i = 1 To .Count
        DoEvents
        a = a + IIf(IsNumeric(.Item(i).SubItems(10)), CDbl(.Item(i).SubItems(10)), 0)
    Next i
    lblTotalNetCost.Caption = Format(a, "#,##0.00")
End With
End Sub

Private Sub txtType_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(ROW).SubItems(2) = txtType.Text
    End With
End If
End Sub

Private Sub txtUnit_Change()
lstDetail.ListItems.Item(ROW).SubItems(5) = txtUnit.Text
End Sub

Private Sub txtUnit_GotFocus()
HTEXT txtUnit
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtItemCode.SetFocus
End If
End Sub

Private Sub txtVat_GotFocus()
HTEXT txtVAT
End Sub

Private Sub txtVAT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    lstDetail.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtDisc1.SetFocus
End If
End Sub

Private Sub txtVAT_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub


