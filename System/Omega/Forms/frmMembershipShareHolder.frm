VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMembershipShareHolder 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMembershipShareHolder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picAdd 
      Height          =   3855
      Left            =   2400
      TabIndex        =   19
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6800
      BackColor       =   15396057
      Begin VB.ListBox lstResultAdd 
         Height          =   2205
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtSearchAdd 
         Height          =   315
         Left            =   120
         TabIndex        =   22
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
         Picture         =   "frmMembershipShareHolder.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3195
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
         Picture         =   "frmMembershipShareHolder.frx":1026
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3195
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   40
         TabIndex        =   24
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
         Icon            =   "frmMembershipShareHolder.frx":1698
         ShadowVisible   =   0   'False
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   840
      ScaleHeight     =   3375
      ScaleWidth      =   7455
      TabIndex        =   1
      Top             =   1200
      Width           =   7455
      Begin MSComctlLib.ListView lstDetails 
         Height          =   1215
         Left            =   0
         TabIndex        =   18
         Top             =   2160
         Width           =   7455
         _ExtentX        =   13150
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "IDNumberKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "MemberKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Member's Name"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C6B8A4&
         Height          =   855
         Left            =   960
         TabIndex        =   13
         Top             =   1200
         Width           =   2535
         Begin VB.TextBox txtSqrMtr 
            Height          =   315
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   16
            Top             =   360
            Width           =   1095
         End
         Begin VB.PictureBox picWithLot 
            BackColor       =   &H00C6B8A4&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   855
            TabIndex        =   14
            Top             =   0
            Width           =   855
            Begin VB.CheckBox chkWithLot 
               BackColor       =   &H00C6B8A4&
               Caption         =   "w/ Lot"
               Height          =   195
               Left            =   0
               TabIndex        =   15
               Top             =   0
               Width           =   750
            End
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Square Meter"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.TextBox txtNoShare 
         Height          =   315
         Left            =   6240
         MaxLength       =   100
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtDateTo 
         Height          =   315
         Left            =   3000
         MaxLength       =   100
         TabIndex        =   10
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtDateFrom 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtShareCertNo 
         Height          =   315
         Left            =   6240
         MaxLength       =   100
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cmbShareType 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox txtShareHolder 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   2
         Top             =   0
         Width           =   6375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Share"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4680
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Share Certificate No"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4680
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Share Type"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Share Holder"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   25
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   26
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
         MouseIcon       =   "frmMembershipShareHolder.frx":1C32
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
      TabIndex        =   0
      Top             =   4755
      Width           =   9090
      _ExtentX        =   16034
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
      Left            =   8400
      Top             =   960
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
            Picture         =   "frmMembershipShareHolder.frx":1F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipShareHolder.frx":2C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipShareHolder.frx":3900
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipShareHolder.frx":45DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipShareHolder.frx":52B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipShareHolder.frx":5F8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipShareHolder.frx":6C68
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipShareHolder.frx":7942
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipShareHolder.frx":861C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipShareHolder.frx":8EF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipShareHolder.frx":9BD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipShareHolder.frx":A8AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipShareHolder.frx":B584
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipShareHolder.frx":C25E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipShareHolder.frx":CF38
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMembershipShareHolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim FocusDetail As Long

Dim TRANSDetail As Long
Const isDetRefresh = 0
Const isDetAdding = 1
Const isDetEditting = 2

Dim iMasterKey, iMemberType, iShareType, x

Private Sub CLEARTEXT()
iMasterKey = 0
iMemberType = 0
iShareType = 0
txtShareHolder.Text = ""
txtShareCertNo.Text = ""
txtNoShare.Text = ""
txtDateFrom.Text = ""
txtDateTo.Text = ""
txtSqrMtr.Text = ""
chkWithLot.Value = 0
cmbShareType.ListIndex = 0
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtShareHolder.Locked = bln
txtShareCertNo.Locked = bln
txtNoShare.Locked = bln
txtDateFrom.Locked = bln
txtDateTo.Locked = bln
txtSqrMtr.Locked = bln
cmbShareType.Locked = bln
picWithLot.Enabled = IIf(bln = True, False, True)
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
cmdCancelAdd_Click
End Sub

Private Sub cmdCancelAdd_Click()
picAdd.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True
End Sub

Private Sub cmdOKAdd_Click()
If lstResultAdd.ListIndex = -1 Then Exit Sub
CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
txtShareHolder.Text = lstResultAdd.List(lstResultAdd.ListIndex)
iMasterKey = lstResultAdd.ItemData(lstResultAdd.ListIndex)
t = "SELECT KeyType, PrimaryKeys " & _
    " From dbo.tbl_Member_Company_Corporate_Keys " & _
    " WHERE (PK = " & iMasterKey & ")"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then
    iMemberType = rt!KeyType
    Select Case iMemberType
        Case 3
            u = "SELECT ID1, ID2 " & _
                " From dbo.tbl_Corporate_Account " & _
                " WHERE (PK = " & rt!PrimaryKeys & ")"
            If ru.State = adStateOpen Then ru.Close
            ru.Open u, ConnOmega
            If ru.RecordCount > 0 Then
                If Trim(ru!ID1) <> "" Then
                    v = "SELECT PK " & _
                        " From dbo.tbl_Share_IDNumber " & _
                        " WHERE (IDNumber = '" & Trim(ru!ID1) & "')"
                    If rv.State = adStateOpen Then rv.Close
                    rv.Open v, ConnOmega
                    If rv.RecordCount > 0 Then
                        Set x = lstDetails.ListItems.Add()
                        x.Text = ""
                        x.SubItems(1) = rv!PK
                        x.SubItems(2) = ru!ID1
                        x.SubItems(3) = "0"
                        x.SubItems(4) = " "
                    End If
                    rv.Close
                End If
                If Trim(ru!ID2) <> "" Then
                    v = "SELECT PK " & _
                        " From dbo.tbl_Share_IDNumber " & _
                        " WHERE (IDNumber = '" & Trim(ru!ID2) & "')"
                    If rv.State = adStateOpen Then rv.Close
                    rv.Open v, ConnOmega
                    If rv.RecordCount > 0 Then
                        Set x = lstDetails.ListItems.Add()
                        x.Text = ""
                        x.SubItems(1) = rv!PK
                        x.SubItems(2) = ru!ID2
                        x.SubItems(3) = "0"
                        x.SubItems(4) = " "
                    End If
                    rv.Close
                End If
            End If
            ru.Close
        Case 2
            u = "SELECT ID1 " & _
                " From dbo.tbl_Company_Account " & _
                " WHERE (PK = " & rt!PrimaryKeys & ")"
            If ru.State = adStateOpen Then ru.Close
            ru.Open u, ConnOmega
            If ru.RecordCount > 0 Then
                If Trim(ru!ID1) <> "" Then
                    v = "SELECT PK " & _
                        " From dbo.tbl_Share_IDNumber " & _
                        " WHERE (IDNumber = '" & Trim(ru!ID1) & "')"
                    If rv.State = adStateOpen Then rv.Close
                    rv.Open v, ConnOmega
                    If rv.RecordCount > 0 Then
                        Set x = lstDetails.ListItems.Add()
                        x.Text = ""
                        x.SubItems(1) = rv!PK
                        x.SubItems(2) = ru!ID1
                        x.SubItems(3) = "0"
                        x.SubItems(4) = " "
                    End If
                    rv.Close
                End If
            End If
            ru.Close
    End Select
End If
rt.Close
cmdCancelAdd_Click
End Sub

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.Height - Me.Height) / 3
Me.Left = (MainForm.Width - Me.Width) / 5
With cmbShareType
    .Clear
    .AddItem "--Select--"
    .ItemData(.NewIndex) = 0
    s = "SELECT tbl_Share_Type.* " & _
        " FROM tbl_Share_Type " & _
        " ORDER BY ShareTypeName "
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        .AddItem rs!ShareTypeName
        .ItemData(.NewIndex) = rs!PK
        rs.MoveNext
    Wend
    rs.Close
End With
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH

End Sub

Private Sub txtSearchAdd_Change()
If Trim(txtSearchAdd.Text) = "" Then lstResultAdd.Clear:  Exit Sub
lstResultAdd.Clear
s = "SELECT PK, (CASE dbo.tbl_Member_Company_Corporate_Keys.KeyType WHEN 1 THEN " & _
    " (SELECT dbo.tbl_Member_Information.LastName + ',  ' + dbo.tbl_Member_Information.FirstName + '  ' + dbo.tbl_Member_Information.MiddleName AS QueryName " & _
    " From dbo.tbl_Member_Information WHERE (dbo.tbl_Member_Information.PK = dbo.tbl_Member_Company_Corporate_Keys.PrimaryKeys)) " & _
    " ELSE (CASE dbo.tbl_Member_Company_Corporate_Keys.KeyType WHEN 2 THEN (SELECT dbo.tbl_Company_Account.Name " & _
    " From dbo.tbl_Company_Account WHERE (dbo.tbl_Company_Account.PK = dbo.tbl_Member_Company_Corporate_Keys.PrimaryKeys)) " & _
    " ELSE (CASE dbo.tbl_Member_Company_Corporate_Keys.KeyType WHEN 3 THEN (SELECT dbo.tbl_Corporate_Account.Name " & _
    " From dbo.tbl_Corporate_Account WHERE (dbo.tbl_Corporate_Account.PK = dbo.tbl_Member_Company_Corporate_Keys.PrimaryKeys)) ELSE '' END) END) END) AS sName " & _
    " From dbo.tbl_Member_Company_Corporate_Keys " & _
    " WHERE ((CASE dbo.tbl_Member_Company_Corporate_Keys.KeyType WHEN 1 THEN " & _
    " (SELECT dbo.tbl_Member_Information.LastName + ',  ' + dbo.tbl_Member_Information.FirstName + '  ' + dbo.tbl_Member_Information.MiddleName AS QueryName " & _
    " From dbo.tbl_Member_Information WHERE (dbo.tbl_Member_Information.PK = dbo.tbl_Member_Company_Corporate_Keys.PrimaryKeys)) " & _
    " ELSE (CASE dbo.tbl_Member_Company_Corporate_Keys.KeyType WHEN 2 THEN (SELECT dbo.tbl_Company_Account.Name " & _
    " From dbo.tbl_Company_Account WHERE (dbo.tbl_Company_Account.PK = dbo.tbl_Member_Company_Corporate_Keys.PrimaryKeys)) " & _
    " ELSE (CASE dbo.tbl_Member_Company_Corporate_Keys.KeyType WHEN 3 THEN (SELECT dbo.tbl_Corporate_Account.Name " & _
    " From dbo.tbl_Corporate_Account WHERE (dbo.tbl_Corporate_Account.PK = dbo.tbl_Member_Company_Corporate_Keys.PrimaryKeys)) ELSE '' END) END) END) LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%') " & _
    " ORDER BY (CASE dbo.tbl_Member_Company_Corporate_Keys.KeyType WHEN 1 THEN " & _
    " (SELECT dbo.tbl_Member_Information.LastName + ',  ' + dbo.tbl_Member_Information.FirstName + '  ' + dbo.tbl_Member_Information.MiddleName AS QueryName " & _
    " From dbo.tbl_Member_Information WHERE (dbo.tbl_Member_Information.PK = dbo.tbl_Member_Company_Corporate_Keys.PrimaryKeys)) " & _
    " ELSE (CASE dbo.tbl_Member_Company_Corporate_Keys.KeyType WHEN 2 THEN (SELECT dbo.tbl_Company_Account.Name " & _
    " From dbo.tbl_Company_Account WHERE (dbo.tbl_Company_Account.PK = dbo.tbl_Member_Company_Corporate_Keys.PrimaryKeys)) " & _
    " ELSE (CASE dbo.tbl_Member_Company_Corporate_Keys.KeyType WHEN 3 THEN (SELECT dbo.tbl_Corporate_Account.Name " & _
    " From dbo.tbl_Corporate_Account WHERE (dbo.tbl_Corporate_Account.PK = dbo.tbl_Member_Company_Corporate_Keys.PrimaryKeys)) ELSE '' END) END) END)"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultAdd.AddItem rs!sName
    lstResultAdd.ItemData(lstResultAdd.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResultAdd.ListCount Then lstResultAdd.ListIndex = 0
End Sub
