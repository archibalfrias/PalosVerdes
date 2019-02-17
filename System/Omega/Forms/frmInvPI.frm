VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInvPI 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvPI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   13845
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picSLine 
      Height          =   855
      Left            =   120
      TabIndex        =   17
      Top             =   4800
      Visible         =   0   'False
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox txtTotalNetCost1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtTotalCost1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtDescription1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtItemCode1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox cmbType1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtItemKey1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtItemKey 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox cmbUnit1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.ComboBox cmbUnit 
         Height          =   315
         Left            =   6720
         TabIndex        =   50
         Text            =   "cmbUnit"
         Top             =   360
         Width           =   1155
      End
      Begin VB.TextBox txtType1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtType 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   360
         Width           =   1155
      End
      Begin VB.TextBox txtRecd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         TabIndex        =   29
         Top             =   360
         Width           =   1035
      End
      Begin VB.TextBox txtItemCode 
         Height          =   315
         Left            =   2400
         TabIndex        =   28
         Top             =   360
         Width           =   1155
      End
      Begin VB.TextBox txtDescription 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   360
         Width           =   3075
      End
      Begin VB.TextBox txtCost 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7920
         TabIndex        =   26
         Top             =   360
         Width           =   1035
      End
      Begin VB.TextBox txtTotalNetCost 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   360
         Width           =   1275
      End
      Begin VB.TextBox txtNetCost 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9000
         TabIndex        =   24
         Top             =   360
         Width           =   1035
      End
      Begin VB.TextBox txtRecd1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtCost1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtNetCost1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtTotalCost 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtSLRemarks 
         Height          =   315
         Left            =   11400
         TabIndex        =   19
         Top             =   360
         Width           =   1995
      End
      Begin VB.TextBox txtSLRemarks1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "TYPE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "REC'D"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1320
         TabIndex        =   37
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "UNIT"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6840
         TabIndex        =   36
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM CODE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         TabIndex        =   35
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM DESCRIPTION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   34
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "COST"
         Height          =   255
         Left            =   7920
         TabIndex        =   33
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL NETCOST"
         Height          =   255
         Left            =   10080
         TabIndex        =   32
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NETCOST"
         Height          =   255
         Left            =   9000
         TabIndex        =   31
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "REMARKS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   11520
         TabIndex        =   30
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4335
      ScaleWidth      =   13575
      TabIndex        =   1
      Top             =   1200
      Width           =   13575
      Begin VB.ComboBox cmbDeptName 
         Height          =   315
         Left            =   3720
         TabIndex        =   42
         Text            =   "Combo1"
         Top             =   330
         Width           =   5955
      End
      Begin VB.TextBox txtSupplier 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   3720
         TabIndex        =   9
         Top             =   0
         Width           =   5955
      End
      Begin VB.TextBox txtRemarks 
         Height          =   315
         Left            =   3720
         TabIndex        =   8
         Top             =   660
         Width           =   5955
      End
      Begin VB.TextBox txtPINumber 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   0
         Width           =   1155
      End
      Begin VB.TextBox txtPIDate 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   330
         Width           =   1155
      End
      Begin VB.TextBox txtRefNo 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   660
         Width           =   1155
      End
      Begin MSComctlLib.ListView lstDetail 
         Height          =   2505
         Left            =   0
         TabIndex        =   12
         Top             =   1080
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   4419
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
         NumItems        =   16
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
            Text            =   "Type"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Rec'd"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Item Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Item Description"
            Object.Width           =   5997
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Unit"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Cost"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "NetCost"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Total NetCost"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "TotalCost"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "TypeKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "InvQty"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "InvCost"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "InvNetCost"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Remarks"
            Object.Width           =   3440
         EndProperty
      End
      Begin VB.Label lblInvPosted 
         BackStyle       =   0  'Transparent
         Caption         =   "INVOICE POSTED"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   0
         TabIndex        =   99
         Top             =   3720
         Width           =   2775
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "DEPT"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   43
         Top             =   375
         Width           =   975
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
         Left            =   9960
         TabIndex        =   16
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL COST"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8760
         TabIndex        =   15
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL NETCOST"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8760
         TabIndex        =   14
         Top             =   3960
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
         Left            =   9960
         TabIndex        =   13
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "REMARKS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   705
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PI #"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PI DATE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   350
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "REF/PR #"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   680
         Width           =   975
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15600
      TabIndex        =   112
      Top             =   0
      Width           =   15600
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   113
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
            NumButtons      =   26
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
               Caption         =   "GL Acc"
               Key             =   "Accnt"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               Key             =   "Refresh"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         MousePointer    =   99
         MouseIcon       =   "frmInvPI.frx":08CA
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   12300
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   114
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmInvPI.frx":0BE4
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
   Begin RPVGCC.b8Container picAdd 
      Height          =   4095
      Left            =   4680
      TabIndex        =   44
      Top             =   960
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7223
      BackColor       =   15396057
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
         Picture         =   "frmInvPI.frx":12F7
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   3480
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
         Left            =   2280
         Picture         =   "frmInvPI.frx":1969
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   3480
         Width           =   1560
      End
      Begin VB.TextBox txtAdd 
         Height          =   315
         Left            =   120
         TabIndex        =   46
         Top             =   480
         Width           =   4215
      End
      Begin VB.ListBox lstResultAdd 
         Height          =   2595
         Left            =   120
         TabIndex        =   45
         Top             =   840
         Width           =   4215
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   49
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
         Icon            =   "frmInvPI.frx":20C5
         ShadowVisible   =   0   'False
      End
   End
   Begin RPVGCC.b8Container picPost 
      Height          =   1815
      Left            =   5040
      TabIndex        =   100
      Top             =   2160
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3201
      BackColor       =   15396057
      Begin VB.ComboBox cmbLocation 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   600
         Width           =   2535
      End
      Begin VB.CommandButton cmdCancel 
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
         Left            =   2040
         Picture         =   "frmInvPI.frx":265F
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   1080
         Width           =   1560
      End
      Begin VB.CommandButton cmdOK 
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
         Left            =   360
         Picture         =   "frmInvPI.frx":2DBB
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   1080
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar4 
         Height          =   345
         Left            =   40
         TabIndex        =   104
         Top             =   40
         Width           =   3880
         _ExtentX        =   6853
         _ExtentY        =   609
         Caption         =   "Select Location"
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
         Icon            =   "frmInvPI.frx":342D
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "LOCATION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   105
         Top             =   600
         Width           =   975
      End
   End
   Begin RPVGCC.b8Container picGLAddAutoVAT 
      Height          =   2955
      Left            =   4680
      TabIndex        =   106
      Top             =   1440
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5212
      BackColor       =   15396057
      Begin VB.ListBox lstGLAutoVAT 
         Height          =   1425
         Left            =   120
         TabIndex        =   110
         Top             =   840
         Width           =   5295
      End
      Begin VB.TextBox txtGLAddAutoVAT 
         Height          =   315
         Left            =   120
         TabIndex        =   109
         Top             =   480
         Width           =   5295
      End
      Begin VB.CommandButton cmdCancelAutoVAT 
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
         Left            =   2640
         Picture         =   "frmInvPI.frx":39C7
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   2355
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKAutoVAT 
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
         Left            =   960
         Picture         =   "frmInvPI.frx":4123
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   2355
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar5 
         Height          =   345
         Left            =   45
         TabIndex        =   111
         Top             =   45
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   609
         Caption         =   "Enter Debit Account"
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
         Icon            =   "frmInvPI.frx":4795
         ShadowVisible   =   0   'False
      End
   End
   Begin VB.PictureBox picAccDistribution 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   3600
      ScaleHeight     =   3855
      ScaleWidth      =   7335
      TabIndex        =   59
      Top             =   1080
      Visible         =   0   'False
      Width           =   7335
      Begin RPVGCC.b8Container b8Container1 
         Height          =   3615
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6376
         BackColor       =   15396057
         Begin VB.CommandButton cmdPost 
            Caption         =   "P O S T"
            Height          =   495
            Left            =   120
            TabIndex        =   66
            Top             =   3000
            Width           =   1095
         End
         Begin VB.TextBox txtInvNetP 
            Height          =   315
            Left            =   5280
            TabIndex        =   65
            Top             =   720
            Width           =   1875
         End
         Begin VB.TextBox txtInvNumberP 
            Height          =   315
            Left            =   120
            TabIndex        =   64
            Top             =   720
            Width           =   1515
         End
         Begin VB.TextBox txtInvDateP 
            Height          =   315
            Left            =   1680
            TabIndex        =   63
            Top             =   720
            Width           =   1515
         End
         Begin VB.TextBox txtInvGrossP 
            Height          =   315
            Left            =   3240
            TabIndex        =   62
            Top             =   720
            Width           =   1995
         End
         Begin VB.ComboBox cmbBookType 
            Height          =   315
            Left            =   2280
            TabIndex        =   61
            Text            =   "Combo1"
            Top             =   3080
            Width           =   1215
         End
         Begin MSComctlLib.ListView lstAccDistribution 
            Height          =   1815
            Left            =   120
            TabIndex        =   67
            Top             =   1080
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   3201
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
               SubItemIndex    =   1
               Text            =   "Code"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Name"
               Object.Width           =   5821
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Debit"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Credit"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Amount"
               Object.Width           =   0
            EndProperty
         End
         Begin RPVGCC.b8TitleBar b8TitleBar1 
            Height          =   345
            Left            =   45
            TabIndex        =   68
            Top             =   45
            Width           =   7245
            _ExtentX        =   12779
            _ExtentY        =   609
            Caption         =   "Account Distribution"
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
            Icon            =   "frmInvPI.frx":4D2F
         End
         Begin VB.Label lblTotalDebit 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4440
            TabIndex        =   78
            Top             =   3000
            Width           =   1215
         End
         Begin VB.Label lblTotalCredit 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5760
            TabIndex        =   77
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL >>"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3600
            TabIndex        =   76
            Top             =   3000
            Width           =   855
         End
         Begin VB.Label lblBalance 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5760
            TabIndex        =   75
            Top             =   3240
            Width           =   1095
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "BALANCE >>"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3600
            TabIndex        =   74
            Top             =   3240
            Width           =   975
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "INVOICE NET AMOUNT"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5280
            TabIndex        =   73
            Top             =   480
            Width           =   1875
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "INVOICE NUMBER"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "INVOICE DATE"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1680
            TabIndex        =   71
            Top             =   480
            Width           =   1515
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "INVOICE GROSS AMOUNT"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3240
            TabIndex        =   70
            Top             =   480
            Width           =   1995
         End
         Begin VB.Label Label45 
            BackStyle       =   0  'Transparent
            Caption         =   "BOOK TYPE"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1320
            TabIndex        =   69
            Top             =   3120
            Width           =   975
         End
      End
   End
   Begin RPVGCC.b8Container picSearchGLAccount 
      Height          =   2955
      Left            =   4920
      TabIndex        =   79
      Top             =   2040
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5212
      BackColor       =   15396057
      Begin VB.CommandButton cmdOKGLAccount 
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
         Left            =   960
         Picture         =   "frmInvPI.frx":52C9
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   2355
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelGLAccount 
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
         Left            =   2640
         Picture         =   "frmInvPI.frx":593B
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   2355
         Width           =   1560
      End
      Begin VB.TextBox txtSearchGLAccount 
         Height          =   315
         Left            =   120
         TabIndex        =   81
         Top             =   480
         Width           =   5295
      End
      Begin VB.ListBox lstResultGLAccount 
         Height          =   1425
         Left            =   120
         TabIndex        =   80
         Top             =   840
         Width           =   5295
      End
      Begin RPVGCC.b8TitleBar b8TitleBar3 
         Height          =   345
         Left            =   45
         TabIndex        =   84
         Top             =   45
         Width           =   5445
         _ExtentX        =   9604
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
         Icon            =   "frmInvPI.frx":6097
         ShadowVisible   =   0   'False
      End
   End
   Begin VB.PictureBox picADSLine 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   3600
      ScaleHeight     =   855
      ScaleWidth      =   7695
      TabIndex        =   85
      Top             =   1320
      Visible         =   0   'False
      Width           =   7695
      Begin RPVGCC.b8Container picADSLine1 
         Height          =   855
         Left            =   0
         TabIndex        =   86
         Top             =   0
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   1508
         BackColor       =   8438015
         Begin VB.TextBox txtCredit1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   94
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtDebit1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   93
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtAccountName1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   92
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtAccountNo1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   91
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtCredit 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5880
            TabIndex        =   90
            Top             =   360
            Width           =   1275
         End
         Begin VB.TextBox txtDebit 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4560
            TabIndex        =   89
            Top             =   360
            Width           =   1275
         End
         Begin VB.TextBox txtAccountName 
            Height          =   315
            Left            =   1320
            TabIndex        =   88
            Top             =   360
            Width           =   3195
         End
         Begin VB.TextBox txtAccountNo 
            Height          =   315
            Left            =   120
            TabIndex        =   87
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "CREDIT"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5880
            TabIndex        =   98
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DEBIT"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4560
            TabIndex        =   97
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "ACCOUNT NAME"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1320
            TabIndex        =   96
            Top             =   120
            Width           =   3135
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "ACCOUNT #"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   120
            Width           =   975
         End
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   5610
      Width           =   13845
      _ExtentX        =   24421
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17639
            MinWidth        =   17639
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "PRINTED"
            TextSave        =   "PRINTED"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   13320
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPI.frx":6631
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPI.frx":730B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPI.frx":7FE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPI.frx":8CBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPI.frx":9999
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPI.frx":A673
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPI.frx":B34D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPI.frx":C027
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPI.frx":CD01
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPI.frx":D5DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPI.frx":E2B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPI.frx":EF8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPI.frx":FC69
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPI.frx":10943
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPI.frx":1161D
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvPI.frx":122F7
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInvPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim TRANS_DETAIL As Long
Const is_DET_REFRESH = 0
Const is_DET_ADDING = 1
Const is_DET_EDITTING = 2

Dim iFocus      As Long
Dim iRow        As Long
Dim tmp         As Long
Dim iFocusItem  As Long
Dim iSearch     As Long

Dim dR_Inv_Qty      As Double
Dim dR_Inv_Cost     As Double
Dim dR_Inv_NetCost  As Double
Dim isGLCodeFocus   As Long
Dim isGLCodeChange  As Long

Dim iSupplier, iDepartment, sCtrl, iPK, iBookType, Arr, i, l, x, a, b, _
dVATable, dNetVAT, dVAT

Private Sub BROWSER(Ctrl, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If Ctrl <> "" Then
            s = "SELECT TOP 1 tbl_Inv_PI.* " & _
                " FROM tbl_Inv_PI " & _
                " WHERE (CtrlNo = '" & Ctrl & "') " & _
                " ORDER BY CtrlNo"
        Else
            s = "SELECT TOP 1 tbl_Inv_PI.* " & _
                " FROM tbl_Inv_PI " & _
                " ORDER BY CtrlNo"
        End If
    Case "is_HOME"
        If picAdd.Visible = True Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_PI.* " & _
            " FROM tbl_Inv_PI " & _
            " ORDER BY CtrlNo"
    Case "is_PAGEUP"
        If picAdd.Visible = True Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_PI.* " & _
            " FROM tbl_Inv_PI " & _
            " WHERE (CtrlNo < '" & Ctrl & "') " & _
            " ORDER BY CtrlNo DESC"
    Case "is_PAGEDOWN"
        If picAdd.Visible = True Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_PI.* " & _
            " FROM tbl_Inv_PI " & _
            " WHERE (CtrlNo > '" & Ctrl & "') " & _
            " ORDER BY CtrlNo "
    Case "is_END"
        If picAdd.Visible = True Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_PI.* " & _
            " FROM tbl_Inv_PI " & _
            " ORDER BY CtrlNo DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iBookType = IIf(IsNull(rs!BookType), 5, rs!BookType)
    iSupplier = rs!SupplierKey
    iDepartment = rs!DepartmentKey
    txtPINumber.Text = rs!CtrlNo
    txtPIDate.Text = Format(rs!PIDate, "mm/dd/yyyy")
    txtRefNo.Text = rs!Reference
    txtRemarks.Text = rs!Remarks
    
    txtSupplier.Text = ""
    t = "SELECT SupplierCode, SupplierName " & _
        " From tbl_Inv_Supplier " & _
        " WHERE  (PK = " & iSupplier & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtSupplier.Text = rt!SupplierCode & " - " & rt!SupplierName
    End If
    rt.Close
    
    cmbDeptName.Text = ""
    cmbDeptName.ListIndex = -1
    t = "SELECT Code, DeptName " & _
        " From tbl_GL_Department " & _
        " WHERE (PK = " & iDepartment & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        cmbDeptName.Text = UCase(rt!DeptName)
    End If
    rt.Close
    
    lblTotalCost.Caption = Format(rs!TotalCost, "#,##0.00")
    lblTotalNetCost.Caption = Format(rs!TotalNetCost, "#,##0.00")
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    Statusbar1.Panels(3).Text = IIf(rs!Printed = 1, "PRINTED", "")
    imgPosted.Visible = IIf(rs!Posted = 1, True, False)
    lblInvPosted.Visible = IIf(rs!GLPosted = 1, True, False)
    
    t = "SELECT tbl_Inv_PI_Det.* " & _
        " FROM tbl_Inv_PI_Det " & _
        " WHERE (PIKey = " & rs!PK & ") " & _
        " ORDER BY Line"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        With lstDetail.ListItems
            .Clear: a = 0
            While Not rt.EOF
                a = a + 1
                Set x = .Add()
                x.Text = rt!ItemKey
                x.SubItems(1) = Format(a, "0#")
                x.SubItems(2) = IIf(rt!Item_FA = 1, "Items", IIf(rt!Item_FA = 2, "Fixed Asset", ""))     'itemtype
                x.SubItems(3) = Format(rt!Qty, "#,##0.00")   'recd
                
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
                    x.SubItems(4) = ru!ItemCode
                    x.SubItems(5) = ru!ItemDesc
                Else
                    x.SubItems(4) = " "
                    x.SubItems(5) = " "
                End If
                ru.Close
                
                x.SubItems(6) = rt!Unit     'unit
                x.SubItems(7) = Format(rt!Cost, "#,##0.00")     'cost
                x.SubItems(8) = Format(rt!NetCost, "#,##0.00")     'netcost
                x.SubItems(9) = Format(rt!TotalNetCost, "#,##0.00")     'totalnetcost
                x.SubItems(10) = Format(rt!TotalCost, "#,##0.00")    'totalcost
                x.SubItems(11) = rt!Item_FA    'Typekey
                x.SubItems(12) = rt!Inv_Qty    'InvQty
                x.SubItems(13) = rt!Inv_Cost    'InvCost
                x.SubItems(14) = rt!Inv_NetCost    'InvNetCost
                x.SubItems(15) = rt!Remarks    'Remarks
                rt.MoveNext
            Wend
        End With
    Else
        CLEAR_DETAIL
    End If
    rt.Close
    
    SaveSetting App.EXEName, "PURCHASE_INVOICE", "PURC_INV", rs!CtrlNo
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If TRANSACTIONTYPE = is_REFRESH Then

    If picAccDistribution.Visible = True Then
        If picADSLine.Visible = True Then Exit Sub
        'If is_DET_FOCUS = 0 Then Exit Sub
        If iFocus = 0 Then Exit Sub
        With lstAccDistribution.ListItems
            If .Count > 1 Then
                Set x = .Add()
                x.Text = ""
                x.SubItems(1) = " "
                x.SubItems(2) = " "
                x.SubItems(3) = " "
                x.SubItems(4) = " "
                iRow = .Count
            Else
                If Trim(.Item(.Count).SubItems(1)) <> "" Then
                    Set x = .Add()
                    x.Text = ""
                    x.SubItems(1) = " "
                    x.SubItems(2) = " "
                    x.SubItems(3) = " "
                    x.SubItems(4) = " "
                    iRow = .Count
                Else
                    iRow = 1
                End If
            End If
            lstAccDistribution.ListItems(iRow).EnsureVisible
            lstAccDistribution.ListItems(iRow).Selected = True
            TRANS_DETAIL = is_DET_ADDING
            isGLCodeChange = 1
            txtAccountNo.Text = ""
            txtAccountName.Text = ""
            txtDebit.Text = ""
            txtCredit.Text = ""
            picADSLine.Height = 855
            picADSLine.Width = 7335
            picADSLine.ZOrder 0
            picAccDistribution.Enabled = False
            picADSLine.Visible = True
            txtAccountNo.SetFocus
            Exit Sub
        End With
    End If
    
    If picAdd.Visible = True Then Exit Sub
    If AccessRights("Purchase Invoice", "Add") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    iSearch = 1
    picMain.Enabled = False
    picToolbar.Enabled = False
    picAdd.ZOrder 0
    txtAdd.Text = ""
    picAdd.Visible = True
    txtAdd.SetFocus
Else
    If picSLine.Visible = True Then Exit Sub
    If iFocus = 0 Then Exit Sub
    If imgPosted.Visible = True Then Exit Sub
    With lstDetail.ListItems
        If Trim(.Item(1).SubItems(2)) <> "" Then
            Set x = .Add()
            x.Text = ""
            x.SubItems(1) = Format(.Count, "0#")
            x.SubItems(2) = " "     'itemtype
            x.SubItems(3) = " "     'recd
            x.SubItems(4) = " "     'item code
            x.SubItems(5) = " "     'item desc
            x.SubItems(6) = " "     'unit
            x.SubItems(7) = " "     'cost
            x.SubItems(8) = " "     'netcost
            x.SubItems(9) = " "     'totalnetcost
            x.SubItems(10) = " "    'totalcost
            x.SubItems(11) = " "    'Typekey
            x.SubItems(12) = " "    'InvQty
            x.SubItems(13) = " "    'InvCost
            x.SubItems(14) = " "    'InvNetCost
            x.SubItems(15) = " "    'Remarks
            iRow = .Count
            lstDetail.ListItems(iRow).EnsureVisible
            lstDetail.ListItems(iRow).Selected = True
            txtItemKey.Text = ""
'            cmbType.Text = ""
            cmbType.ListIndex = -1
            txtType.Text = ""
            txtRecd.Text = ""
            cmbUnit.Clear
            txtItemCode.Text = ""
            txtDescription.Text = ""
            txtCost.Text = ""
            txtNetCost.Text = ""
            txtTotalNetCost.Text = ""
            txtTotalCost.Text = ""
            txtSLRemarks.Text = ""
            picToolbar.Enabled = False
            picMain.Enabled = False
            picSLine.ZOrder 0
            picSLine.Visible = True
            TRANS_DETAIL = is_DET_ADDING
            cmbType.SetFocus
        Else
            
            .Item(1).SubItems(1) = Format(.Count, "0#")
            .Item(1).SubItems(2) = " "     'itemtype
            .Item(1).SubItems(3) = " "     'recd
            .Item(1).SubItems(4) = " "     'item code
            .Item(1).SubItems(5) = " "     'item desc
            .Item(1).SubItems(6) = " "     'unit
            .Item(1).SubItems(7) = " "     'cost
            .Item(1).SubItems(8) = " "     'netcost
            .Item(1).SubItems(9) = " "     'totalnetcost
            .Item(1).SubItems(10) = " "    'totalcost
            .Item(1).SubItems(11) = " "    'Typekey
            .Item(1).SubItems(12) = " "    'InvQty
            .Item(1).SubItems(13) = " "    'InvCost
            .Item(1).SubItems(14) = " "    'InvNetCost
            .Item(1).SubItems(15) = " "    'Remarks
            iRow = 1
            lstDetail.ListItems(iRow).EnsureVisible
            lstDetail.ListItems(iRow).Selected = True
            txtItemKey.Text = ""
'            cmbType.Text = ""
            cmbType.ListIndex = -1
            txtType.Text = ""
            txtRecd.Text = ""
            cmbUnit.Clear
            txtItemCode.Text = ""
            txtDescription.Text = ""
            txtCost.Text = ""
            txtNetCost.Text = ""
            txtTotalNetCost.Text = ""
            txtTotalCost.Text = ""
            txtSLRemarks.Text = ""
            picToolbar.Enabled = False
            picMain.Enabled = False
            picSLine.ZOrder 0
            picSLine.Visible = True
            TRANS_DETAIL = is_DET_ADDING
            cmbType.SetFocus
            
        End If
    End With
End If
End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE = is_REFRESH Then
    
    If picAccDistribution.Visible = True Then
        If iFocus = 0 Then Exit Sub
        If picADSLine.Visible = True Then Exit Sub
        If lblInvPosted.Visible = True Then Exit Sub
        With lstAccDistribution.ListItems
            If Trim(.Item(iRow).SubItems(1)) = "" Then Exit Sub
            txtAccountNo.Text = .Item(iRow).SubItems(1)
            txtAccountName.Text = .Item(iRow).SubItems(2)
            txtDebit.Text = .Item(iRow).SubItems(3)
            txtCredit.Text = .Item(iRow).SubItems(4)
            txtAccountNo1.Text = .Item(iRow).SubItems(1)
            txtAccountName1.Text = .Item(iRow).SubItems(2)
            txtDebit1.Text = .Item(iRow).SubItems(3)
            txtCredit1.Text = .Item(iRow).SubItems(4)
        End With
        TRANS_DETAIL = is_DET_EDITTING
        isGLCodeChange = 1
        picADSLine.Height = 855
        picADSLine.Width = 7335
        picADSLine.ZOrder 0
        picAccDistribution.Enabled = False
        picADSLine.Visible = True
        txtAccountNo.SetFocus
        Exit Sub
    End If
    
    If Statusbar1.Panels(1).Text = "" Then Exit Sub
    If picAdd.Visible = True Then Exit Sub
    If AccessRights("Purchase Invoice", "Edit") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If imgPosted.Visible = True Then MsgBox "Already Posted!                         ", vbCritical, "Error...": Exit Sub
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_EDITTING
    'Me.Caption = "PURCHASE INVOICE - EDIT"
Else
    If picSLine.Visible = True Then Exit Sub
    If iFocus = 0 Then Exit Sub
    If imgPosted.Visible = True Then Exit Sub
    With lstDetail.ListItems
        txtType.Text = .Item(iRow).SubItems(11)
        txtItemKey.Text = .Item(iRow).Text
        cmbUnit.Clear
        Select Case RETURNTEXTVALUE(txtType)
            Case 1
                t = "SELECT PK, ItemCode, ItemDesc, " & _
                    " Unit, Unit2, Cost, SuppKey" & _
                    " From tbl_Inv_Items " & _
                    " WHERE (PK = " & RETURNTEXTVALUE(txtItemKey) & ")"
            Case 2
                t = "SELECT PK, Code as ItemCode, " & _
                    " Description as ItemDesc, Unit" & _
                    " FROM tbl_FA_Items " & _
                    " WHERE (PK = " & RETURNTEXTVALUE(txtItemKey) & ")"
        End Select
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            With cmbUnit
                .AddItem rt!Unit
                If RETURNTEXTVALUE(txtType) = 1 Then
                    If Trim(rt!Unit2) <> "" Then
                        If Trim(rt!Unit2) <> Trim(rt!Unit) Then
                            .AddItem rt!Unit2
                        End If
                    End If
                End If
            End With
        End If
        rt.Close
        
        cmbType.ListIndex = .Item(iRow).SubItems(11) - 1
        
        txtRecd.Text = .Item(iRow).SubItems(3)
        txtItemCode.Text = .Item(iRow).SubItems(4)
        txtDescription.Text = .Item(iRow).SubItems(5)
        cmbUnit.Text = .Item(iRow).SubItems(6)
        txtCost.Text = .Item(iRow).SubItems(7)
        txtNetCost.Text = .Item(iRow).SubItems(8)
        txtTotalNetCost.Text = .Item(iRow).SubItems(9)
        txtTotalCost.Text = .Item(iRow).SubItems(10)
        txtSLRemarks.Text = .Item(iRow).SubItems(15)
        
        txtItemKey1.Text = .Item(iRow).Text
        txtType1.Text = .Item(iRow).SubItems(11)
        cmbType1.Text = .Item(iRow).SubItems(2)
        txtRecd1.Text = .Item(iRow).SubItems(3)
        txtItemCode1.Text = .Item(iRow).SubItems(4)
        txtDescription1.Text = .Item(iRow).SubItems(5)
        cmbUnit1.Text = .Item(iRow).SubItems(6)
        txtCost1.Text = .Item(iRow).SubItems(7)
        txtNetCost1.Text = .Item(iRow).SubItems(8)
        txtTotalNetCost1.Text = .Item(iRow).SubItems(9)
        txtTotalCost1.Text = .Item(iRow).SubItems(10)
        txtSLRemarks1.Text = .Item(iRow).SubItems(15)
        TRANS_DETAIL = is_DET_EDITTING
        picToolbar.Enabled = False
        picMain.Enabled = False
        picSLine.ZOrder 0
        picSLine.Visible = True
        cmbType.SetFocus
    End With
End If
End Sub

Private Sub PRESS_DELETE()
If TRANSACTIONTYPE = is_REFRESH Then
    
    If picAccDistribution.Visible = True Then
        If picADSLine.Visible = True Then Exit Sub
        If iFocus = 0 Then Exit Sub
        With lstAccDistribution.ListItems
            If .Count > 1 Then
                .Remove iRow
                If CDbl(iRow) > CDbl(.Count) Then iRow = .Count
            Else
                .Item(1).SubItems(1) = " "
                .Item(1).SubItems(2) = " "
                .Item(1).SubItems(3) = " "
                .Item(1).SubItems(4) = " "
                iRow = 1
            End If
            lstAccDistribution.ListItems(iRow).EnsureVisible
            lstAccDistribution.ListItems(iRow).Selected = True
            isGLCodeChange = 1
        End With
        Exit Sub
    End If
    
    If Statusbar1.Panels(1).Text = "" Then Exit Sub
    If picAdd.Visible = True Then Exit Sub
    If AccessRights("Purchase Invoice", "Delete") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If imgPosted.Visible = True Then MsgBox "Already Posted!                         ", vbCritical, "Error...": Exit Sub
    If MsgBox("ARE YOU SURE IN DELETING THIS RECORD?                        ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    On Error GoTo PG:
    ConnOmega.Execute "DELETE FROM tbl_Inv_PI WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    CLEARTEXT
    BROWSER GetSetting(App.EXEName, "PURCHASE_INVOICE", "PURC_INV", ""), "is_PAGEDOWN"
    If Trim(txtPINumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "PURCHASE_INVOICE", "PURC_INV", ""), "is_HOME"
Else
    If picSLine.Visible = True Then Exit Sub
    If iFocus = 0 Then Exit Sub
    If imgPosted.Visible = True Then Exit Sub
    With lstDetail.ListItems
        If .Count > 1 Then
            .Remove iRow
            If CDbl(iRow) > CDbl(.Count) Then
                iRow = .Count
            End If
        Else
            .Item(1).SubItems(1) = " "
            .Item(1).SubItems(2) = " "     'itemtype
            .Item(1).SubItems(3) = " "     'recd
            .Item(1).SubItems(4) = " "     'item code
            .Item(1).SubItems(5) = " "     'item desc
            .Item(1).SubItems(6) = " "     'unit
            .Item(1).SubItems(7) = " "     'cost
            .Item(1).SubItems(8) = " "     'netcost
            .Item(1).SubItems(9) = " "     'totalnetcost
            .Item(1).SubItems(10) = " "    'totalcost
            .Item(1).SubItems(11) = " "    'Typekey
            .Item(1).SubItems(12) = " "    'InvQty
            .Item(1).SubItems(13) = " "    'InvCost
            .Item(1).SubItems(14) = " "    'InvNetCost
            .Item(1).SubItems(15) = " "    'Remarks
            iRow = 1
        End If
        lstDetail.ListItems(iRow).EnsureVisible
        lstDetail.ListItems(iRow).Selected = True
    End With
End If
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()

If picAccDistribution.Visible = True Then
    If picSearchGLAccount.Visible = True Then Exit Sub
    If picADSLine.Visible = True Then Exit Sub
    If iBookType = 0 Then MsgBox "Please Select Book Type!                      ", vbCritical, "Error...": cmbBookType.SetFocus: Exit Sub
    On Error GoTo PH:
    With lstAccDistribution.ListItems
        a = 0
        ConnOmega.Execute "DELETE FROM tbl_Inv_PI_Account_Distribution WHERE (PIKey = " & Statusbar1.Panels(1).Text & ")"
        For i = 1 To .Count
            If Trim(.Item(i).SubItems(1)) <> "" Then
                a = a + 1
                ConnOmega.Execute "INSERT INTO tbl_Inv_PI_Account_Distribution " & _
                                  " (PIKey, Line, AccountCode, Debit, Credit) " & _
                                  " VALUES (" & Statusbar1.Panels(1).Text & ", " & _
                                  " " & a & ", '" & Trim(.Item(i).SubItems(1)) & "', " & _
                                  " " & CDbl(IIf(Trim(.Item(i).SubItems(3)) = "", 0, .Item(i).SubItems(3))) & ", " & _
                                  " " & CDbl(IIf(Trim(.Item(i).SubItems(4)) = "", 0, .Item(i).SubItems(4))) & ")"
            End If
        Next i
        ConnOmega.Execute "UPDATE tbl_Inv_PI SET BookType = " & iBookType & " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    End With
    isGLCodeChange = 0
    
    Exit Sub
End If

If picSLine.Visible = True Then Exit Sub
If IsDate(txtPIDate.Text) = False Then MsgBox "Please Supply a Valid Date!                    ", vbCritical, "Error...": txtPIDate.SetFocus: Exit Sub
If iSupplier = 0 Then MsgBox "Please Supply Supplier!                       ", vbCritical, "Error...": Exit Sub
If iDepartment = 0 Then MsgBox "Pleaes Select Department!                   ", vbCritical, "Error...": cmbDeptName.SetFocus: Exit Sub
If Trim(txtRefNo.Text) = "" Then MsgBox "Please Supply Ref #!                       ", vbCritical, "Error...": txtRefNo.SetFocus: Exit Sub
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    sCtrl = ""
    s = "SELECT TOP 1 CtrlNo " & _
        " FROM tbl_Inv_PI " & _
        " WHERE (Year(PIDate) = " & Format(FormatDateTime(txtPIDate.Text, vbShortDate), "yyyy") & ") " & _
        " ORDER BY CtrlNo DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        sCtrl = Format(CDbl(rs!CtrlNo) + 1, "#######0")
    Else
        sCtrl = Format(FormatDateTime(txtPIDate.Text, vbShortDate), "yyyy") & "0000"
    End If
    rs.Close
    Do
        s = "SELECT tbl_Inv_PI.* " & _
            " FROM tbl_Inv_PI " & _
            " WHERE (CtrlNo = '" & sCtrl & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If
        rs.Close
        sCtrl = Format(CDbl(sCtrl) + 1, "#######0")
    Loop
    
    ConnOmega.Execute "INSERT INTO tbl_Inv_PI " & _
                      " (CtrlNo, PIDate, Reference, SupplierKey, DepartmentKey, Remarks, LastModified, " & _
                      " TotalCost, TotalNetCost) " & _
                      " VALUES ('" & sCtrl & "', '" & FormatDateTime(txtPIDate.Text, vbShortDate) & "', " & _
                      " '" & FORMATSQL(Trim(txtRefNo.Text)) & "', " & iSupplier & ", " & iDepartment & ", " & _
                      " '" & FORMATSQL(Trim(txtRemarks.Text)) & "', '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                      " " & RETURNLABELVALUE(lblTotalCost) & ", " & RETURNLABELVALUE(lblTotalNetCost) & ")"
    iPK = 0
    s = "SELECT PK " & _
        " FROM tbl_Inv_PI " & _
        " WHERE (CtrlNo = '" & sCtrl & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        iPK = rs!PK
    End If
    rs.Close
         
    If CDbl(iPK) > 0 Then
        With lstDetail.ListItems
            a = 0
            For i = 1 To .Count
                If CDbl(IIf(IsNumeric(.Item(i).Text) = False, 0, .Item(i).Text)) <> 0 Then
                    Select Case .Item(i).SubItems(11)
                        Case 1
                            s = "SELECT Unit, ConUnit, Unit2, ConUnit2 " & _
                                " From tbl_Inv_Items " & _
                                " WHERE (ItemCode = '" & .Item(i).SubItems(4) & "')"
                            If rs.State = adStateOpen Then rs.Close
                            rs.Open s, ConnOmega
                            If rs.RecordCount > 0 Then
                                If Trim(.Item(i).SubItems(6)) = Trim(rs!Unit) Then
                                    dR_Inv_Qty = .Item(i).SubItems(3)
                                    dR_Inv_Cost = .Item(i).SubItems(7)
                                    dR_Inv_NetCost = .Item(i).SubItems(8)
                                ElseIf Trim(.Item(i).SubItems(6)) = Trim(rs!Unit2) Then
                                    dR_Inv_Qty = Format(CDbl(.Item(i).SubItems(3)) * CDbl(rs!ConUnit), "#,##0.00")
                                    dR_Inv_Cost = Format(CDbl(.Item(i).SubItems(7)) / CDbl(rs!ConUnit), "#,##0.00")
                                    dR_Inv_NetCost = Format(CDbl(.Item(i).SubItems(8)) / CDbl(rs!ConUnit), "#,##0.00")
                                End If
                            End If
                            rs.Close
                        Case Else
                            dR_Inv_Qty = .Item(i).SubItems(3)
                            dR_Inv_Cost = .Item(i).SubItems(7)
                            dR_Inv_NetCost = .Item(i).SubItems(8)
                    End Select
                    a = a + 1
                    ConnOmega.Execute "INSERT INTO tbl_Inv_PI_Det " & _
                                      " (PIKey, Line, Item_FA, ItemKey, Qty, Unit, Cost, NetCost, Remarks, Inv_Qty, " & _
                                      " Inv_Cost, Inv_NetCost) " & _
                                      " VALUES (" & iPK & ", " & a & ", " & .Item(i).SubItems(11) & ", " & _
                                      " " & .Item(i).Text & ", " & CDbl(.Item(i).SubItems(3)) & ", " & _
                                      " '" & FORMATSQL(.Item(i).SubItems(6)) & "', " & CDbl(.Item(i).SubItems(7)) & ", " & _
                                      " " & CDbl(.Item(i).SubItems(8)) & ", '" & FORMATSQL(.Item(i).SubItems(15)) & "', " & _
                                      " " & CDbl(dR_Inv_Qty) & ", " & CDbl(dR_Inv_Cost) & ", " & CDbl(dR_Inv_NetCost) & ")"
                End If
            Next i
        End With
    End If
    
    
End If
If TRANSACTIONTYPE = is_EDITTING Then
    iPK = Statusbar1.Panels(1).Text
    sCtrl = Trim(txtPINumber.Text)
    ConnOmega.Execute "UPDATE tbl_Inv_PI " & _
                      " SET PIDate = '" & FormatDateTime(txtPIDate.Text, vbShortDate) & "', " & _
                      " Reference = '" & FORMATSQL(Trim(txtRefNo.Text)) & "', " & _
                      " Remarks = '" & FORMATSQL(Trim(txtRemarks.Text)) & "', " & _
                      " TotalCost = " & RETURNLABELVALUE(lblTotalCost) & ", " & _
                      " TotalNetCost = " & RETURNLABELVALUE(lblTotalNetCost) & ", " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & iPK & ")"
    
    If CDbl(iPK) > 0 Then
        ConnOmega.Execute "DELETE FROM tbl_Inv_PI_Det WHERE (PIKey = " & iPK & ")"
        With lstDetail.ListItems
            a = 0
            For i = 1 To .Count
                If CDbl(IIf(IsNumeric(.Item(i).Text) = False, 0, .Item(i).Text)) <> 0 Then
                    Select Case .Item(i).SubItems(11)
                        Case 1
                            s = "SELECT Unit, ConUnit, Unit2, ConUnit2 " & _
                                " From tbl_Inv_Items " & _
                                " WHERE (ItemCode = '" & .Item(i).SubItems(4) & "')"
                            If rs.State = adStateOpen Then rs.Close
                            rs.Open s, ConnOmega
                            If rs.RecordCount > 0 Then
                                If Trim(.Item(i).SubItems(6)) = Trim(rs!Unit) Then
                                    dR_Inv_Qty = .Item(i).SubItems(3)
                                    dR_Inv_Cost = .Item(i).SubItems(7)
                                    dR_Inv_NetCost = .Item(i).SubItems(8)
                                ElseIf Trim(.Item(i).SubItems(6)) = Trim(rs!Unit2) Then
                                    dR_Inv_Qty = Format(CDbl(.Item(i).SubItems(3)) * CDbl(rs!ConUnit), "#,##0.00")
                                    dR_Inv_Cost = Format(CDbl(.Item(i).SubItems(7)) / CDbl(rs!ConUnit), "#,##0.00")
                                    dR_Inv_NetCost = Format(CDbl(.Item(i).SubItems(8)) / CDbl(rs!ConUnit), "#,##0.00")
                                End If
                            End If
                            rs.Close
                        Case Else
                            dR_Inv_Qty = .Item(i).SubItems(3)
                            dR_Inv_Cost = .Item(i).SubItems(7)
                            dR_Inv_NetCost = .Item(i).SubItems(8)
                    End Select
                    a = a + 1
                    ConnOmega.Execute "INSERT INTO tbl_Inv_PI_Det " & _
                                      " (PIKey, Line, Item_FA, ItemKey, Qty, Unit, Cost, NetCost, Remarks, Inv_Qty, " & _
                                      " Inv_Cost, Inv_NetCost) " & _
                                      " VALUES (" & iPK & ", " & a & ", " & .Item(i).SubItems(11) & ", " & _
                                      " " & .Item(i).Text & ", " & CDbl(.Item(i).SubItems(3)) & ", " & _
                                      " '" & FORMATSQL(.Item(i).SubItems(6)) & "', " & CDbl(.Item(i).SubItems(7)) & ", " & _
                                      " " & CDbl(.Item(i).SubItems(8)) & ", '" & FORMATSQL(.Item(i).SubItems(15)) & "', " & _
                                      " " & CDbl(dR_Inv_Qty) & ", " & CDbl(dR_Inv_Cost) & ", " & CDbl(dR_Inv_NetCost) & ")"
                End If
            Next i
        End With
    End If
End If
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANS_DETAIL = is_DET_REFRESH
'Me.Caption = "PURCHASE INVOICE - BROWSE"
BROWSER sCtrl, "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
PH:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If picSLine.Visible = True Then Exit Sub

End Sub

Private Sub PRESS_F7()
If picSLine.Visible = True Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If picSLine.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If imgPosted.Visible = False Then Exit Sub
If AccessRights("Purchase Invoice", "Post GL") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If

txtInvNumberP.Text = txtPINumber.Text ' txtInvNumber.Text
txtInvDateP.Text = txtPIDate.Text  'txtInvDate.Text
txtInvGrossP.Text = lblTotalCost.Caption  'txtInvGross.Text
txtInvNetP.Text = lblTotalNetCost.Caption  'txtInvNet.Text

lstAccDistribution.ListItems.Clear
Set x = lstAccDistribution.ListItems.Add()
x.Text = ""
x.SubItems(1) = " "
x.SubItems(2) = " "
x.SubItems(3) = " "
x.SubItems(4) = " "
x.SubItems(5) = "0"
a = 0: b = 0
s = "SELECT tbl_Inv_PI_Account_Distribution.AccountCode, " & _
    " tbl_GL_Accounts.AccountName, " & _
    " tbl_Inv_PI_Account_Distribution.Debit, " & _
    " tbl_Inv_PI_Account_Distribution.Credit, " & _
    " tbl_Inv_PI_Account_Distribution.Amount " & _
    " FROM tbl_Inv_PI_Account_Distribution LEFT OUTER JOIN " & _
    " tbl_GL_Accounts ON tbl_Inv_PI_Account_Distribution.AccountCode = tbl_GL_Accounts.AccountCode " & _
    " Where (tbl_Inv_PI_Account_Distribution.PIKey = " & Statusbar1.Panels(1).Text & ") " & _
    " ORDER BY tbl_Inv_PI_Account_Distribution.Line"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    lstAccDistribution.ListItems.Clear
    While Not rs.EOF
        Set x = lstAccDistribution.ListItems.Add()
        x.Text = ""
        x.SubItems(1) = rs!AccountCode
        x.SubItems(2) = rs!AccountName
        x.SubItems(3) = IIf(CDbl(rs!Debit) = 0, " ", Format(rs!Debit, "#,##0.00"))
        x.SubItems(4) = IIf(CDbl(rs!Credit) = 0, " ", Format(rs!Credit, "#,##0.00"))
        x.SubItems(5) = rs!Amount
        a = a + CDbl(rs!Debit)
        b = b + CDbl(rs!Credit)
        rs.MoveNext
    Wend
End If
rs.Close
lblTotalDebit.Caption = Format(a, "#,##0.00")
lblTotalCredit.Caption = Format(b, "#,##0.00")

cmbBookType.Clear
s = "SELECT tbl_Acctg_Book.* " & _
    " FROM tbl_Acctg_Book " & _
    " WHERE (ViewInRR = 1) " & _
    " ORDER BY PK"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbBookType.AddItem rs!Abb
    cmbBookType.ItemData(cmbBookType.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close

s = "SELECT tbl_Acctg_Book.* " & _
    " FROM tbl_Acctg_Book " & _
    " WHERE (PK = " & iBookType & ") " & _
    " AND (ViewInRR = 1)"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    cmbBookType.Text = rs!Abb
End If
rs.Close

picToolbar.Enabled = False
picMain.Enabled = False
picAccDistribution.Height = 3615 '3015
picAccDistribution.Width = 7335
picAccDistribution.ZOrder 0
picAccDistribution.Visible = True
lstAccDistribution.SetFocus

End Sub

Private Sub PRESS_F8()
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If picSLine.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If AccessRights("Purchase Invoice", "Post Inv") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If imgPosted.Visible = True Then MsgBox "Already Posted!                         ", vbCritical, "Error...": Exit Sub

cmbLocation.Clear
s = "SELECT tbl_Inv_Location.* " & _
    " FROM tbl_Inv_Location " & _
    " WHERE (Receiving = 1) " & _
    " ORDER BY LocName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbLocation.AddItem rs!LocName
    cmbLocation.ItemData(cmbLocation.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close

picToolbar.Enabled = False
picMain.Enabled = False
picPost.ZOrder 0
picPost.Visible = True
cmbLocation.SetFocus
    
End Sub

Private Sub PRESS_F9()
If picSLine.Visible = True Then Exit Sub
If picSLine.Visible = True Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If picSLine.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
If AccessRights("Purchase Invoice", "Print") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSearchGLAccount.Visible = True Then cmdCancelGLAccount_Click: Exit Sub
    If picADSLine.Visible = True Then
        With lstAccDistribution.ListItems
            If TRANS_DETAIL = is_DET_ADDING Then
                If .Count > 1 Then
                    .Remove .Count
                    iRow = .Count
                Else
                    iRow = 1
                    .Item(iRow).SubItems(1) = " "
                    .Item(iRow).SubItems(2) = " "
                    .Item(iRow).SubItems(3) = " "
                    .Item(iRow).SubItems(4) = " "
                End If
            ElseIf TRANS_DETAIL = is_DET_EDITTING Then
                .Item(iRow).SubItems(1) = txtAccountNo1.Text
                .Item(iRow).SubItems(2) = txtAccountName1.Text
                .Item(iRow).SubItems(3) = txtDebit1.Text
                .Item(iRow).SubItems(4) = txtCredit1.Text
            End If
        End With
        picADSLine.Visible = False
        picAccDistribution.Enabled = True
        lstAccDistribution.SetFocus
        Exit Sub
    End If
    If picAccDistribution.Visible = True Then b8TitleBar2_CLoseClick: Exit Sub
    If picPost.Visible = True Then cmdCancel_Click: Exit Sub
    If picAdd.Visible = True Then cmdCancelAdd_Click: Exit Sub
    Unload Me
Else
    If picAdd.Visible = True Then cmdCancelAdd_Click: Exit Sub
    If picSLine.Visible = True Then
        With lstDetail.ListItems
            If TRANS_DETAIL = is_DET_ADDING Then
                If .Count > 1 Then
                    .Remove .Count
                Else
                    .Item(1).Text = " "
                    .Item(1).SubItems(1) = " "
                    .Item(1).SubItems(2) = " "     'itemtype
                    .Item(1).SubItems(3) = " "     'recd
                    .Item(1).SubItems(4) = " "     'item code
                    .Item(1).SubItems(5) = " "     'item desc
                    .Item(1).SubItems(6) = " "     'unit
                    .Item(1).SubItems(7) = " "     'cost
                    .Item(1).SubItems(8) = " "     'netcost
                    .Item(1).SubItems(9) = " "     'totalnetcost
                    .Item(1).SubItems(10) = " "    'totalcost
                    .Item(1).SubItems(11) = " "    'Typekey
                    .Item(1).SubItems(12) = " "    'InvQty
                    .Item(1).SubItems(13) = " "    'InvCost
                    .Item(1).SubItems(14) = " "    'InvNetCost
                    .Item(1).SubItems(15) = " "    'Remarks
                End If
                iRow = .Count
                picMain.Enabled = True
                picToolbar.Enabled = True
                picSLine.Visible = False
                lstDetail.ListItems(iRow).EnsureVisible
                lstDetail.ListItems(iRow).Selected = True
                lstDetail.SetFocus
                Exit Sub
            End If
            If TRANS_DETAIL = is_DET_EDITTING Then
                .Item(iRow).Text = txtItemKey1.Text
                .Item(iRow).SubItems(11) = txtType1.Text
                .Item(iRow).SubItems(2) = cmbType1.Text
                .Item(iRow).SubItems(3) = txtRecd1.Text
                .Item(iRow).SubItems(4) = txtItemCode1.Text
                .Item(iRow).SubItems(5) = txtDescription1.Text
                .Item(iRow).SubItems(6) = cmbUnit1.Text
                .Item(iRow).SubItems(7) = txtCost1.Text
                .Item(iRow).SubItems(8) = txtNetCost1.Text
                .Item(iRow).SubItems(9) = txtTotalNetCost1.Text
                .Item(iRow).SubItems(10) = txtTotalCost1.Text
                .Item(iRow).SubItems(15) = txtSLRemarks1.Text
                picMain.Enabled = True
                picToolbar.Enabled = True
                picSLine.Visible = False
                lstDetail.ListItems(iRow).EnsureVisible
                lstDetail.ListItems(iRow).Selected = True
                lstDetail.SetFocus
                Exit Sub
            End If
        End With
    End If
    If iFocus = 1 Then
        txtPINumber.SetFocus
        Exit Sub
    End If
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    'Me.Caption = "PURCHASE INVOICE - BROWSE"
    BROWSER GetSetting(App.EXEName, "PURCHASE_INVOICE", "PURC_INV", ""), "is_LOAD"
    If Trim(txtPINumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "PURCHASE_INVOICE", "PURC_INV", ""), "is_HOME"
End If
End Sub

Private Sub CLEARTEXT()
iSupplier = 0
iDepartment = 0
txtPINumber.Text = ""
txtPIDate.Text = ""
txtRefNo.Text = ""
txtSupplier.Text = ""
txtRemarks.Text = ""
cmbDeptName.Text = ""
cmbDeptName.ListIndex = -1
lblTotalCost.Caption = "0.00"
lblTotalNetCost.Caption = "0.00"
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
Statusbar1.Panels(3).Text = ""
imgPosted.Visible = False
lblInvPosted.Visible = False
CLEAR_DETAIL
End Sub

Private Sub CLEAR_DETAIL()
With lstDetail.ListItems
    .Clear
    Set x = .Add()
    x.Text = ""             'ItemKey
    x.SubItems(1) = " "
    x.SubItems(2) = " "     'itemtype
    x.SubItems(3) = " "     'recd
    x.SubItems(4) = " "     'item code
    x.SubItems(5) = " "     'item desc
    x.SubItems(6) = " "     'unit
    x.SubItems(7) = " "     'cost
    x.SubItems(8) = " "     'netcost
    x.SubItems(9) = " "     'totalnetcost
    x.SubItems(10) = " "    'totalcost
    x.SubItems(11) = " "    'Typekey
    x.SubItems(12) = " "    'InvQty
    x.SubItems(13) = " "    'InvCost
    x.SubItems(14) = " "    'InvNetCost
    x.SubItems(15) = " "    'Remarks
End With
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtPINumber.Locked = True
txtPIDate.Locked = bln
txtRefNo.Locked = bln
txtSupplier.Locked = True
txtRemarks.Locked = bln
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
            .Buttons(25).Image = 14
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
            .Buttons(21).Caption = "GL Acc"
            .Buttons(23).Caption = "Refresh"
            .Buttons(25).Caption = "Close"
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
            .Buttons(25).Enabled = True
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
            .Buttons(21).ToolTipText = "ACCOUNT DISTRIBUTION (F7)"
            .Buttons(23).ToolTipText = "REFRESH (F11)"
            .Buttons(25).ToolTipText = "CLOSE (Esc)"
        Case 2      '=== ADD/EDIT ====
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 15
            .Buttons(9).Image = 16
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(21).Image = 12
            .Buttons(23).Image = 13
            .Buttons(25).Image = 14
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
            .Buttons(21).Caption = "GL Acc"
            .Buttons(23).Caption = "Refresh"
            .Buttons(25).Caption = "Close"
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
            .Buttons(25).Enabled = False
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
            .Buttons(25).ToolTipText = ""
        Case 3      '=== FIND ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 4
            .Buttons(9).Image = 16
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(21).Image = 12
            .Buttons(23).Image = 13
            .Buttons(25).Image = 14
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
            .Buttons(21).Caption = "GL Acc"
            .Buttons(23).Caption = "Refresh"
            .Buttons(25).Caption = "Close"
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
            .Buttons(25).Enabled = False
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
            .Buttons(25).ToolTipText = ""
        Case 4      '=== EMPTY DETAIL ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 15
            .Buttons(9).Image = 16
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(21).Image = 12
            .Buttons(23).Image = 13
            .Buttons(25).Image = 14
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
            .Buttons(21).Caption = "GL Acc"
            .Buttons(23).Caption = "Refresh"
            .Buttons(25).Caption = "Close"
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
            .Buttons(25).Enabled = False
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
            .Buttons(25).ToolTipText = ""
        Case 5      '=== NOT EMPTY DETAIL ===
            .Buttons(1).Image = 1
            .Buttons(3).Image = 2
            .Buttons(5).Image = 3
            .Buttons(7).Image = 15
            .Buttons(9).Image = 16
            .Buttons(11).Image = 6
            .Buttons(13).Image = 7
            .Buttons(15).Image = 8
            .Buttons(17).Image = 9
            .Buttons(19).Image = 10
            .Buttons(21).Image = 12
            .Buttons(23).Image = 13
            .Buttons(25).Image = 14
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
            .Buttons(21).Caption = "GL Acc"
            .Buttons(23).Caption = "Refresh"
            .Buttons(25).Caption = "Close"
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
            .Buttons(25).Enabled = False
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
            .Buttons(25).ToolTipText = ""
    End Select
End With
End Sub

Private Sub b8TitleBar1_CLoseClick()
If isGLCodeChange = 1 Then
    If MsgBox("Save Changes!                     ", vbCritical + vbYesNo, "Save") = vbYes Then
        PRESS_F5
    Else
        picToolbar.Enabled = True
        picMain.Enabled = True
        picAccDistribution.Visible = False
        Exit Sub
    End If
End If
picToolbar.Enabled = True
picMain.Enabled = True
picAccDistribution.Visible = False
End Sub

Private Sub b8TitleBar2_CLoseClick()
cmdCancelAdd_Click
End Sub

Private Sub b8TitleBar3_CLoseClick()
cmdCancelGLAccount_Click
End Sub

Private Sub b8TitleBar4_CLoseClick()
cmdCancel_Click
End Sub

Private Sub cmbBookType_Click()
If cmbBookType.ListIndex = -1 Then Exit Sub
iBookType = cmbBookType.ItemData(cmbBookType.ListIndex)
End Sub

Private Sub cmbDeptName_Click()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If cmbDeptName.ListIndex = -1 Then Exit Sub
    iDepartment = cmbDeptName.ItemData(cmbDeptName.ListIndex)
End If
End Sub

Private Sub cmbType_Click()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(2) = cmbType.List(cmbType.ListIndex)
        txtType.Text = cmbType.ListIndex + 1
    End With
End If
End Sub

Private Sub cmbType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtRecd.SetFocus
End Sub

Private Sub cmbUnit_Click()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(6) = cmbUnit.List(cmbUnit.ListIndex)
    End With
End If
End Sub

Private Sub cmbUnit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtCost.SetFocus
End Sub

Private Sub cmdCancel_Click()
picPost.Visible = False
picToolbar.Enabled = True
picMain.Enabled = True
End Sub

Private Sub cmdCancelAdd_Click()
If iSearch = 1 Then
    picToolbar.Enabled = True
    picMain.Enabled = True
    picAdd.Visible = False
ElseIf iSearch = 2 Then
    picAdd.Visible = False
End If
End Sub

Private Sub cmdCancelGLAccount_Click()
picSearchGLAccount.Visible = False
picADSLine.Enabled = True
txtAccountNo.SetFocus
End Sub

Private Sub cmdOK_Click()
If cmbLocation.ListIndex = -1 Then MsgBox "Please Select Location!                          ", vbCritical, "Error...": cmbLocation.SetFocus: Exit Sub
If MsgBox("ARE YOU SURE YOU WANT TO POST THIS TRANSACTION?                          ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
With lstDetail.ListItems
    ' Inventory / FA Lapsing
    For i = 1 To .Count
        'Items
        If CDbl(IIf(IsNumeric(.Item(i).SubItems(11)) = False, 0, .Item(i).SubItems(11))) = 1 Then
            s = "SELECT tbl_Inv_Items.* " & _
                " FROM tbl_Inv_Items " & _
                " WHERE (PK = " & .Item(i).Text & ")"
            If rs.State = adStateOpen Then rs.Close
            rs.Open s, ConnOmega
            If rs.RecordCount > 0 Then
                dVATable = .Item(i).SubItems(14)
                dNetVAT = NET_OF_VAT(FormatDateTime(txtPIDate.Text, vbShortDate), dVATable, rs!PK)
                dVAT = CDbl(dVATable) - CDbl(dNetVAT)
                ConnOmega.Execute "INSERT INTO tbl_Inv_Items_Transaction " & _
                                  " (ItemKey, Cleared, InOut, DocType, DocNumber, DocDate, " & _
                                  " Location, QuantityIn, Cost, NetCost, LogInName, NetVAT) " & _
                                  " VALUES (" & rs!PK & ", 0, 'I', 2, '" & txtPINumber.Text & "', " & _
                                  " '" & FormatDateTime(txtPIDate.Text, vbShortDate) & "', " & _
                                  " " & cmbLocation.ItemData(cmbLocation.ListIndex) & ", " & _
                                  " " & .Item(i).SubItems(12) & ", " & .Item(i).SubItems(13) & ", " & _
                                  " " & .Item(i).SubItems(14) & ", '" & gbl_UserName & "', " & _
                                  " " & CDbl(dNetVAT) & ")"
                
                t = "SELECT tbl_Inv_Items_Cost.* " & _
                    " FROM tbl_Inv_Items_Cost " & _
                    " WHERE (ItemKey = " & rs!PK & ") " & _
                    " AND (EffectDate = '" & FormatDateTime(txtPIDate.Text, vbShortDate) & "')"
                If rt.State = adStateOpen Then rt.Close
                rt.Open t, ConnOmega
                If rt.RecordCount = 0 Then
                    ConnOmega.Execute "INSERT INTO tbl_Inv_Items_Cost " & _
                                      " (ItemKey, EffectDate, Cost) " & _
                                      " VALUES (" & rs!PK & ", " & _
                                      " '" & FormatDateTime(txtPIDate.Text, vbShortDate) & "', " & _
                                      " " & .Item(i).SubItems(14) & ")"
                Else
                    ConnOmega.Execute "UPDATE tbl_Inv_Items_Cost " & _
                                      " SET Cost = " & .Item(i).SubItems(14) & " " & _
                                      " WHERE (ItemKey = " & rs!PK & ") " & _
                                      " AND (EffectDate = '" & FormatDateTime(txtPIDate.Text, vbShortDate) & "')"
                End If
                rt.Close
                
            End If
            rs.Close
        'Fixed Assets
        ElseIf CDbl(IIf(IsNumeric(.Item(i).SubItems(11)) = False, 0, .Item(i).SubItems(11))) = 2 Then
            s = "SELECT tbl_FA_Items.* " & _
                " FROM tbl_FA_Items " & _
                " WHERE (PK = " & .Item(i).Text & ")"
            If rs.State = adStateOpen Then rs.Close
            rs.Open s, ConnOmega
            If rs.RecordCount > 0 Then
                ConnOmega.Execute "INSERT INTO tbl_FA_Items_Lapsing " & _
                                  " (FAKey, DepreciationDate, DepreciationAmount) " & _
                                  " VALUES (" & rs!PK & ", '" & FormatDateTime(txtPIDate.Text, vbShortDate) & "', " & _
                                  " '" & .Item(i).SubItems(9) & "')"
                ConnOmega.Execute "UPDATE tbl_FA_Items " & _
                                  " SET DateAcquired = '" & FormatDateTime(txtPIDate.Text, vbShortDate) & "', " & _
                                  " Cost = '" & .Item(i).SubItems(9) & "' " & _
                                  " WHERE (PK = " & rs!PK & ")"
            End If
            rs.Close
        End If
    Next i

End With
                  
ConnOmega.Execute "UPDATE tbl_Inv_PI " & _
                  " SET Posted = 1, " & _
                  " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                  " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
                  
cmdCancel_Click

BROWSER GetSetting(App.EXEName, "PURCHASE_INVOICE", "PURC_INV", ""), "is_LOAD"

End Sub

Private Sub cmdOKAdd_Click()
If lstResultAdd.ListIndex = -1 Then Exit Sub
If iSearch = 1 Then
    CLEARTEXT
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_ADDING
    'Me.Caption = "PURCHASE INVOICE - NEW"
    iSupplier = lstResultAdd.ItemData(lstResultAdd.ListIndex)
    txtSupplier.Text = lstResultAdd.List(lstResultAdd.ListIndex)
    cmdCancelAdd_Click
    txtPIDate.Text = Format(FormatDateTime(Date, vbShortDate), "mm/dd/yyyy")
    txtPIDate.SetFocus
ElseIf iSearch = 2 Then
    Arr = Split(lstResultAdd.List(lstResultAdd.ListIndex), " - ", -1, 1)
    txtItemCode.Text = Arr(0)
    txtDescription.Text = Arr(1)
    cmdCancelAdd_Click
    txtItemCode.SetFocus
End If
End Sub

Private Sub cmdOKGLAccount_Click()
If lstResultGLAccount.ListIndex = -1 Then Exit Sub
Arr = Split(lstResultGLAccount.List(lstResultGLAccount.ListIndex), " : ", -1, 1)
txtAccountNo.Text = Arr(0)
txtAccountName.Text = Arr(1)
cmdCancelGLAccount_Click
txtDebit.SetFocus
End Sub

Private Sub cmdPost_Click()
If lblInvPosted.Visible = True Then MsgBox "Already Posted!                          ", vbCritical, "Error...": Exit Sub
If RETURNLABELVALUE(lblBalance) <> 0 Then MsgBox "Account Distribution not Balance!                      ", vbCritical, "Error...": Exit Sub
If RETURNTEXTVALUE(txtInvNetP) <> RETURNLABELVALUE(lblTotalDebit) Then MsgBox "Please Check your details!                           ", vbCritical, "Error...": Exit Sub
If iBookType = 0 Then MsgBox "Please Select Book Type!                      ", vbCritical, "Error...": cmbBookType.SetFocus: Exit Sub
With lstAccDistribution.ListItems
    a = 0
    For i = 1 To .Count
        If Trim(.Item(i).SubItems(1)) <> "" Then
            a = a + 1
        End If
    Next i
End With
If CDbl(a) = 0 Then MsgBox "Check Account Distribution!                       ", vbCritical, "Error...": Exit Sub

If MsgBox("ARE YOU SURE IN POSTING THIS INVOICE?                        ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub

On Error GoTo PG:
'-- GL
With lstAccDistribution.ListItems
    Arr = Split(Trim(txtSupplier.Text), " - ", -1, 1)
    For i = 1 To .Count
        If Trim(.Item(i).SubItems(1)) <> "" Then
            ConnOmega.Execute "INSERT INTO tbl_GL_Transaction " & _
                              " (GLCode, DocDate, DocNumber, SupplierCode, SupplierName, InvoiceNumber, InvoiceDate, PayeeKey, " & _
                              " PayeeType, BookType, Debit, Credit) " & _
                              " VALUES ('" & .Item(i).SubItems(1) & "', '" & FormatDateTime(txtPIDate.Text, vbShortDate) & "', " & _
                              " '" & txtPINumber.Text & "', '" & FORMATSQL(CStr(Arr(0))) & "', '" & FORMATSQL(CStr(Arr(1))) & "', " & _
                              " '" & Trim(txtInvNumberP.Text) & "', '" & FormatDateTime(txtInvDateP.Text, vbShortDate) & "', " & _
                              " " & iSupplier & ", 1, " & iBookType & ", " & CDbl(IIf(IsNumeric(.Item(i).SubItems(3)) = False, 0, .Item(i).SubItems(3))) & ", " & _
                              " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(4)) = False, 0, .Item(i).SubItems(4))) & ")"
        
            t = "SELECT tbl_GL_Accounts.* " & _
                " FROM tbl_GL_Accounts " & _
                " WHERE (AccountCode = '" & .Item(i).SubItems(1) & "')"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                If rt!withSL = 1 Then
                    If rt!SupplierKey = 0 Then
                        ConnOmega.Execute "INSERT INTO tbl_Inv_Supplier_SL " & _
                                          " (SupplierKey, GLCode, DocNumber, DocDate, InvoiceNumber, " & _
                                          " InvoiceDate, Description, iType, Reference, Debit, Credit) " & _
                                          " VALUES (" & iSupplier & ", '" & .Item(i).SubItems(1) & "', " & _
                                          " '" & txtPINumber.Text & "', '" & FormatDateTime(txtPIDate.Text, vbShortDate) & "', " & _
                                          " '" & Trim(txtInvNumberP.Text) & "', '" & FormatDateTime(txtInvDateP.Text, vbShortDate) & "', " & _
                                          " 'PURCHASES', " & iBookType & ", '" & "PI# " & Trim(txtPINumber.Text) & ", INV# " & Trim(txtInvNumberP.Text) & "', " & _
                                          " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(3)) = False, 0, .Item(i).SubItems(3))) & ", " & _
                                          " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(4)) = False, 0, .Item(i).SubItems(4))) & ")"
                    Else
                        ConnOmega.Execute "INSERT INTO tbl_Inv_Supplier_SL " & _
                                          " (SupplierKey, GLCode, DocNumber, DocDate, InvoiceNumber, " & _
                                          " InvoiceDate, Description, iType, Reference, Debit, Credit) " & _
                                          " VALUES (" & rt!SupplierKey & ", '" & .Item(i).SubItems(1) & "', " & _
                                          " '" & txtPINumber.Text & "', '" & FormatDateTime(txtPIDate.Text, vbShortDate) & "', " & _
                                          " '" & Trim(txtInvNumberP.Text) & "', '" & FormatDateTime(txtInvDateP.Text, vbShortDate) & "', " & _
                                          " 'PURCHASES', " & iBookType & ", '" & "PI# " & Trim(txtPINumber.Text) & ", INV# " & Trim(txtInvNumberP.Text) & "', " & _
                                          " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(3)) = False, 0, .Item(i).SubItems(3))) & ", " & _
                                          " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(4)) = False, 0, .Item(i).SubItems(4))) & ")"
                    End If
                    
                End If
            End If
            rt.Close
        End If
    Next i
End With

ConnOmega.Execute "UPDATE tbl_Inv_PI " & _
                  " SET GLPosted = 1 " & _
                  " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"

picToolbar.Enabled = True
picMain.Enabled = True
picAccDistribution.Visible = False

BROWSER GetSetting(App.EXEName, "PURCHASE_INVOICE", "PURC_INV", ""), "is_LOAD"

Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub Form_Activate()
If TRANSACTIONTYPE = is_REFRESH Then BROWSER GetSetting(App.EXEName, "PURCHASE_INVOICE", "PURC_INV", ""), "is_LOAD"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert:   PRESS_INSERT
    Case vbKeyF2:       PRESS_F2
    Case vbKeyDelete:   PRESS_DELETE
    Case vbKeyF5:       PRESS_F5
    Case vbKeyF6:       PRESS_F6
    Case vbKeyF7:       PRESS_F7
    Case vbKeyF8:       PRESS_F8
    Case vbKeyF9:       PRESS_F9
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "PURCHASE_INVOICE", "PURC_INV", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "PURCHASE_INVOICE", "PURC_INV", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "PURCHASE_INVOICE", "PURC_INV", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "PURCHASE_INVOICE", "PURC_INV", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
'Me.Caption = "PURCHASE INVOICE - BROWSE"
cmbDeptName.Clear
s = "SELECT PK, Code as DepartmentCode, " & _
    " DeptName as DepartmentName" & _
    " FROM tbl_GL_Department " & _
    " WHERE (Issuance = 1) " & _
    " ORDER BY DeptName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbDeptName.AddItem UCase(rs!DepartmentName)
    cmbDeptName.ItemData(cmbDeptName.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
With cmbType
    .Clear
    .AddItem "Item"
    .AddItem "Fixed Asset"
End With
iFocusItem = 0
iFocus = 0
iRow = 0
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANS_DETAIL = is_DET_REFRESH
BROWSER GetSetting(App.EXEName, "PURCHASE_INVOICE", "PURC_INV", ""), "is_LOAD"
If Trim(txtPINumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "PURCHASE_INVOICE", "PURC_INV", ""), "is_HOME"
tmp = SetWindowLong(txtRemarks.hwnd, GWL_STYLE, GetWindowLong(txtRemarks.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtAdd.hwnd, GWL_STYLE, GetWindowLong(txtAdd.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtRefNo.hwnd, GWL_STYLE, GetWindowLong(txtRefNo.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearchGLAccount.hwnd, GWL_STYLE, GetWindowLong(txtSearchGLAccount.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picSearchGLAccount.Visible = True Then Cancel = -1
If picADSLine.Visible = True Then Cancel = -1
If picAccDistribution.Visible = True Then Cancel = -1
If picSLine.Visible = True Then Cancel = -1
If picAdd.Visible = True Then Cancel = -1
If picPost.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstAccDistribution_GotFocus()
iFocus = 1
iRow = lstAccDistribution.SelectedItem.Index
TRANS_DETAIL = is_DET_REFRESH
End Sub

Private Sub lstAccDistribution_ItemClick(ByVal Item As MSComctlLib.ListItem)
iRow = lstAccDistribution.SelectedItem.Index
End Sub

Private Sub lstAccDistribution_LostFocus()
iFocus = 0
End Sub

Private Sub lstDetail_GotFocus()
TRANS_DETAIL = is_DET_REFRESH
iRow = lstDetail.SelectedItem.Index
iFocus = 1
If imgPosted.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    TRANSACTIONTYPE = is_EDITTING
    'Me.Caption = "PURCHASE INVOICE - EDIT"
End If
If Trim(lstDetail.ListItems.Item(1).Text) <> "" Then
    TOOLBARFUNC 5
Else
    TOOLBARFUNC 4
End If
End Sub

Private Sub lstDetail_ItemClick(ByVal Item As MSComctlLib.ListItem)
iRow = lstDetail.SelectedItem.Index
End Sub

Private Sub lstDetail_LostFocus()
iFocus = 0
End Sub

Private Sub lstResultAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAdd_Click
End Sub

Private Sub lstResultGLAccount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKGLAccount_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "PURCHASE_INVOICE", "PURC_INV", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "PURCHASE_INVOICE", "PURC_INV", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "PURCHASE_INVOICE", "PURC_INV", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "PURCHASE_INVOICE", "PURC_INV", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":   PRESS_F9
    Case "Post":    PRESS_F8
    Case "Accnt":   PRESS_F7
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtAccountName_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstAccDistribution.ListItems
        .Item(iRow).SubItems(2) = Trim(txtAccountName.Text)
    End With
End If
End Sub

Private Sub txtAccountName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF6 Then
    If picADSLine.Visible = False Then Exit Sub
    picADSLine.Enabled = False
    picSearchGLAccount.ZOrder 0
    txtSearchGLAccount.Text = ""
    picSearchGLAccount.Visible = True
    txtSearchGLAccount.SetFocus
ElseIf KeyCode = vbKeyReturn Then
    txtDebit.SetFocus
End If
End Sub

Private Sub txtAccountNo_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstAccDistribution.ListItems
        .Item(iRow).SubItems(1) = Trim(txtAccountNo.Text)
    End With
End If
End Sub

Private Sub txtAccountNo_GotFocus()
isGLCodeFocus = 1
End Sub

Private Sub txtAccountNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If Trim(txtAccountNo.Text) <> "" Then
        s = "SELECT tbl_GL_Accounts.* " & _
            " FROM tbl_GL_Accounts " & _
            " WHERE (AccountCode = '" & Trim(txtAccountNo.Text) & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            txtAccountNo.Text = rs!AccountCode
            txtAccountName.Text = rs!AccountName
        Else
            MsgBox "Account Code '" & Trim(txtAccountNo.Text) & "' not Found!                           ", vbCritical, "Error..."
            rs.Close
            Exit Sub
        End If
        rs.Close
    End If
    txtDebit.SetFocus
End If
If KeyCode = vbKeyF6 Then
    picADSLine.Enabled = False
    txtSearchGLAccount.Text = ""
    picSearchGLAccount.ZOrder 0
    picSearchGLAccount.Visible = True
    txtSearchGLAccount.SetFocus
End If
End Sub

Private Sub txtAccountNo_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtAccountNo_LostFocus()
isGLCodeFocus = 0
End Sub
Private Sub txtAdd_Change()
If Trim(txtAdd.Text) = "" Then lstResultAdd.Clear: Exit Sub
lstResultAdd.Clear
If iSearch = 1 Then
    s = "SELECT PK, SupplierCode, SupplierName " & _
        " From tbl_Inv_Supplier " & _
        " WHERE (SupplierName LIKE '" & FORMATSQL(Trim(txtAdd.Text)) & "%') " & _
        " ORDER BY SupplierName, SupplierCode"
ElseIf iSearch = 2 Then
    Select Case RETURNTEXTVALUE(txtType)
        Case 1
            s = "SELECT PK, ItemCode, ItemDesc, SuppKey, " & _
                " Unit, Unit2, Cost " & _
                " From dbo.tbl_Inv_Items " & _
                " WHERE (ItemDesc LIKE '" & FORMATSQL(Trim(txtAdd.Text)) & "%') " & _
                " AND (SuppKey = " & iSupplier & ") " & _
                " ORDER BY ItemDesc"
        Case 2
            s = "SELECT PK, Code as ItemCode, " & _
                " Description as ItemDesc, Unit" & _
                " FROM tbl_FA_Items " & _
                " WHERE (Description LIKE '" & FORMATSQL(Trim(txtAdd.Text)) & "%') " & _
                " ORDER BY Description"
    End Select
End If
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    If iSearch = 1 Then
        lstResultAdd.AddItem rs!SupplierCode & " - " & rs!SupplierName
    ElseIf iSearch = 2 Then
        lstResultAdd.AddItem rs!ItemCode & " - " & rs!ItemDesc
    End If
    lstResultAdd.ItemData(lstResultAdd.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResultAdd.ListCount Then lstResultAdd.ListIndex = 0
End Sub

Private Sub txtAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultAdd.SetFocus
End Sub

Private Sub txtCost_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(7) = Format(RETURNTEXTVALUE(txtCost), "#,##0.00")
    End With
    txtTotalCost.Text = Format(RETURNTEXTVALUE(txtRecd) * RETURNTEXTVALUE(txtCost), "#,##0.00")
End If
End Sub

Private Sub txtCost_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtNetCost.SetFocus
End Sub

Private Sub txtCost_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtCredit_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstAccDistribution.ListItems
        .Item(iRow).SubItems(4) = IIf(RETURNTEXTVALUE(txtCredit) = 0, " ", Format(RETURNTEXTVALUE(txtCredit), "#,##0.00"))
        .Item(iRow).SubItems(5) = Format(RETURNTEXTVALUE(txtDebit) - RETURNTEXTVALUE(txtCredit), "#,##0.00")
        b = 0
        For i = 1 To .Count
            b = b + CDbl(IIf(IsNumeric(.Item(i).SubItems(4)) = False, 0, .Item(i).SubItems(4)))
        Next i
        lblTotalCredit.Caption = Format(b, "#,##0.00")
    End With
End If
End Sub

Private Sub txtCredit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    picADSLine.Visible = False
    picAccDistribution.Enabled = True
    lstAccDistribution.SetFocus
End If
End Sub

Private Sub txtCredit_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtDebit_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstAccDistribution.ListItems
        .Item(iRow).SubItems(3) = IIf(RETURNTEXTVALUE(txtDebit) = 0, " ", Format(RETURNTEXTVALUE(txtDebit), "#,##0.00"))
        .Item(iRow).SubItems(5) = Format(RETURNTEXTVALUE(txtDebit) - RETURNTEXTVALUE(txtCredit), "#,##0.00")
        b = 0
        For i = 1 To .Count
            b = b + CDbl(IIf(IsNumeric(.Item(i).SubItems(3)) = False, 0, .Item(i).SubItems(3)))
        Next i
        lblTotalDebit.Caption = Format(b, "#,##0.00")
    End With
End If
End Sub

Private Sub txtDebit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtCredit.SetFocus
End Sub

Private Sub txtDebit_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtDescription_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(5) = txtDescription.Text
    End With
End If
End Sub

Private Sub txtItemCode_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(4) = txtItemCode.Text
    End With
End If
End Sub

Private Sub txtItemCode_GotFocus()
iFocusItem = 1
End Sub

Private Sub txtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmbUnit.SetFocus
If KeyCode = vbKeyF6 Then
    iSearch = 2
    picAdd.ZOrder 0
    txtAdd.Text = ""
    picAdd.Visible = True
    txtAdd.SetFocus
End If
If KeyCode = vbKeyInsert Then
    If IsLoaded(frmInvItems) Then frmInvItems.ZOrder 0 Else frmInvItems.Show
    gbl_Item_Module = "Purchase Invoice"
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

Private Sub txtItemCode_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtItemCode_LostFocus()
iFocusItem = 0
If picSLine.Visible = False Then Exit Sub
If Trim(txtItemCode.Text) = "" Then Exit Sub
Select Case RETURNTEXTVALUE(txtType)
    Case 1
        t = "SELECT PK, ItemCode, ItemDesc, " & _
            " Unit, Unit2, Cost, SuppKey" & _
            " From tbl_Inv_Items " & _
            " WHERE (ItemCode = '" & Trim(txtItemCode.Text) & "')"
    Case 2
        t = "SELECT PK, Code as ItemCode, " & _
            " Description as ItemDesc, Unit" & _
            " FROM tbl_FA_Items " & _
            " WHERE (Code = '" & Trim(txtItemCode.Text) & "')"
End Select
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then

    If RETURNTEXTVALUE(txtType) = 1 Then
        If rt!SuppKey <> iSupplier Then
            MsgBox "THIS ITEM NOT BELONG TO THIS SUPPLIER!      ", vbCritical, "Error..."
            rt.Close
            txtItemCode.SetFocus
            HTEXT txtItemCode
            Exit Sub
        End If
    End If
    
    txtItemKey.Text = rt!PK
    
    With cmbUnit
        .Clear
        .AddItem rt!Unit
        If RETURNTEXTVALUE(txtType) = 1 Then
            If Trim(rt!Unit2) <> "" Then
                If Trim(rt!Unit2) <> Trim(rt!Unit) Then
                    .AddItem rt!Unit2
                End If
            End If
        End If
    End With
    
    txtItemCode.Text = rt!ItemCode
    txtDescription.Text = rt!ItemDesc
    With lstDetail.ListItems
        .Item(iRow).Text = rt!PK
        .Item(iRow).SubItems(4) = rt!ItemCode
        .Item(iRow).SubItems(5) = rt!ItemDesc
    End With
    
Else
    txtItemKey.Text = "0"
'    txtUnit.Text = ""
    txtItemCode.Text = ""
    txtDescription.Text = ""
    With lstDetail.ListItems
        .Item(iRow).Text = " "
        .Item(iRow).SubItems(6) = " "
        .Item(iRow).SubItems(4) = " "
        .Item(iRow).SubItems(5) = " "
    End With
End If
rt.Close
End Sub

Private Sub txtItemKey_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).Text = txtItemKey.Text
    End With
End If
End Sub

Private Sub txtNetCost_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(8) = Format(RETURNTEXTVALUE(txtNetCost), "#,##0.00")
    End With
    txtTotalNetCost.Text = Format(RETURNTEXTVALUE(txtRecd) * RETURNTEXTVALUE(txtNetCost), "#,##0.00")
End If
End Sub

Private Sub txtNetCost_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtSLRemarks.SetFocus
End Sub

Private Sub txtNetCost_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtRecd_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(3) = RETURNTEXTVALUE(txtRecd)
    End With
    txtTotalNetCost.Text = Format(RETURNTEXTVALUE(txtRecd) * RETURNTEXTVALUE(txtNetCost), "#,##0.00")
    txtTotalCost.Text = Format(RETURNTEXTVALUE(txtRecd) * RETURNTEXTVALUE(txtCost), "#,##0.00")
End If
End Sub

Private Sub txtRecd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtItemCode.SetFocus
End Sub

Private Sub txtRecd_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtSearchGLAccount_Change()
If Trim(txtSearchGLAccount.Text) = "" Then lstResultGLAccount.Clear: Exit Sub
lstResultGLAccount.Clear
s = "SELECT tbl_GL_Accounts.* " & _
    " FROM tbl_GL_Accounts " & _
    " WHERE (AccountName LIKE '" & FORMATSQL(Trim(txtSearchGLAccount.Text)) & "%') " & _
    " ORDER BY AccountName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultGLAccount.AddItem rs!AccountCode & " : " & rs!AccountName
    rs.MoveNext
Wend
rs.Close
If lstResultGLAccount.ListCount Then lstResultGLAccount.ListIndex = 0
End Sub

Private Sub txtSearchGLAccount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultGLAccount.SetFocus
End Sub

Private Sub txtSLRemarks_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(15) = txtSLRemarks.Text
    End With
End If
End Sub

Private Sub txtSLRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If TRANS_DETAIL = is_DET_ADDING Then
        TRANS_DETAIL = is_DET_REFRESH
        With lstDetail.ListItems
            Set x = .Add()
            x.Text = ""
            x.SubItems(1) = Format(.Count, "0#")
            x.SubItems(2) = " "     'itemtype
            x.SubItems(3) = " "     'recd
            x.SubItems(4) = " "     'item code
            x.SubItems(5) = " "     'item desc
            x.SubItems(6) = " "     'unit
            x.SubItems(7) = " "     'cost
            x.SubItems(8) = " "     'netcost
            x.SubItems(9) = " "     'totalnetcost
            x.SubItems(10) = " "    'totalcost
            x.SubItems(11) = " "    'Typekey
            x.SubItems(12) = " "    'InvQty
            x.SubItems(13) = " "    'InvCost
            x.SubItems(14) = " "    'InvNetCost
            x.SubItems(15) = " "    'Remarks
            iRow = .Count
        End With
        lstDetail.ListItems(iRow).EnsureVisible
        lstDetail.ListItems(iRow).Selected = True
'        cmbType.Text = ""
        cmbType.ListIndex = -1
        txtType.Text = ""
        txtRecd.Text = ""
        cmbUnit.Clear
        txtItemCode.Text = ""
        txtDescription.Text = ""
        txtCost.Text = ""
        txtNetCost.Text = ""
        txtTotalNetCost.Text = ""
        txtSLRemarks.Text = ""
        TRANS_DETAIL = is_DET_ADDING
        cmbType.SetFocus
        
    End If
    If TRANS_DETAIL = is_DET_EDITTING Then
        picSLine.Visible = False
        picToolbar.Enabled = True
        picMain.Enabled = True
        lstDetail.SetFocus
    End If
End If
End Sub

Private Sub txtTotalCost_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(10) = Format(RETURNTEXTVALUE(txtTotalCost), "#,##0.00")
        a = 0
        For i = 1 To .Count
            If CDbl(IIf(IsNumeric(.Item(i).SubItems(10)) = False, 0, .Item(i).SubItems(10))) <> 0 Then
                a = a + CDbl(IIf(IsNumeric(.Item(i).SubItems(10)) = False, 0, .Item(i).SubItems(10)))
            End If
        Next i
        lblTotalCost.Caption = Format(a, "#,##0.00")
    End With
End If
End Sub

Private Sub txtTotalNetCost_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(9) = Format(RETURNTEXTVALUE(txtTotalNetCost), "#,##0.00")
        a = 0
        For i = 1 To .Count
            'ii(cdbl(isnumeric(.Item(i).SubItems(9))=False ,0,.Item(i).SubItems(9)))<>0 then
            If CDbl(IIf(IsNumeric(.Item(i).SubItems(9)) = False, 0, .Item(i).SubItems(9))) <> 0 Then
                a = a + CDbl(IIf(IsNumeric(.Item(i).SubItems(9)) = False, 0, .Item(i).SubItems(9)))
            End If
        Next i
        lblTotalNetCost.Caption = Format(a, "#,##0.00")
    End With
End If
End Sub

Private Sub txtType_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(11) = txtType.Text
    End With
End If
End Sub

'Private Sub txtUnit_Change()
'If TRANS_DETAIL = is_DET_ADDING Or _
'TRANS_DETAIL = is_DET_EDITTING Then
'    With lstDetail.ListItems
'        .Item(iRow).SubItems(4) = txtUnit.Text
'    End With
'End If
'End Sub
