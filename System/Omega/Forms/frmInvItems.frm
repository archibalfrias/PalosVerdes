VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInvItems 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   9075
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvItems.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   51
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   52
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
         MouseIcon       =   "frmInvItems.frx":08CA
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
   Begin RPVGCC.b8Container picSearch 
      Height          =   4095
      Left            =   2400
      TabIndex        =   38
      Top             =   960
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7223
      BackColor       =   15396057
      Begin VB.ListBox lstResult 
         Height          =   2595
         Left            =   120
         TabIndex        =   42
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   4215
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
         Left            =   2280
         Picture         =   "frmInvItems.frx":0BE4
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   3480
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
         Left            =   600
         Picture         =   "frmInvItems.frx":1340
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3480
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   43
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
         Icon            =   "frmInvItems.frx":19B2
         ShadowVisible   =   0   'False
      End
   End
   Begin VB.PictureBox picBody 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   4695
      Left            =   960
      ScaleHeight     =   4695
      ScaleWidth      =   6855
      TabIndex        =   15
      Top             =   1200
      Width           =   6855
      Begin VB.PictureBox picNonVAT 
         BackColor       =   &H00C6B8A4&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5520
         ScaleHeight     =   255
         ScaleWidth      =   1215
         TabIndex        =   49
         Top             =   3650
         Width           =   1215
         Begin VB.CheckBox chkNonVAT 
            BackColor       =   &H00C6B8A4&
            Caption         =   "Non VAT"
            Height          =   255
            Left            =   240
            TabIndex        =   50
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.TextBox txtAccountCode 
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   45
         Top             =   4320
         Width           =   1030
      End
      Begin VB.ComboBox cmbAccountName 
         Height          =   315
         Left            =   2520
         TabIndex        =   44
         Text            =   "Combo1"
         Top             =   4320
         Width           =   4215
      End
      Begin VB.TextBox txtType 
         Height          =   315
         Left            =   3000
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtClassKey 
         Height          =   315
         Left            =   720
         TabIndex        =   24
         Top             =   1080
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtSuppKey 
         Height          =   315
         Left            =   720
         TabIndex        =   23
         Top             =   720
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtRemarks 
         Height          =   315
         Left            =   960
         TabIndex        =   11
         Top             =   3960
         Width           =   5775
      End
      Begin VB.TextBox txtSRP 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   960
         TabIndex        =   10
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox txtCost 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   960
         TabIndex        =   9
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtMinQty 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   960
         TabIndex        =   8
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtMaxQty 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   960
         TabIndex        =   7
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtUnit1 
         Height          =   315
         Left            =   960
         TabIndex        =   5
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtSectName 
         Height          =   315
         Left            =   1800
         TabIndex        =   22
         Top             =   1440
         Width           =   4935
      End
      Begin VB.TextBox txtSectCode 
         Height          =   315
         Left            =   960
         TabIndex        =   4
         Top             =   1440
         Width           =   800
      End
      Begin VB.ComboBox cmbClassName 
         Height          =   315
         Left            =   1800
         TabIndex        =   21
         Text            =   "Combo1"
         Top             =   1080
         Width           =   4935
      End
      Begin VB.TextBox txtClassCode 
         Height          =   315
         Left            =   960
         TabIndex        =   3
         Top             =   1080
         Width           =   800
      End
      Begin VB.ComboBox cmbSuppName 
         Height          =   315
         Left            =   1800
         TabIndex        =   20
         Text            =   "Combo1"
         Top             =   720
         Width           =   4935
      End
      Begin VB.TextBox txtSuppCode 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   720
         Width           =   800
      End
      Begin VB.TextBox txtItemDesc 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   5775
      End
      Begin VB.TextBox txtItemCode 
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Top             =   0
         Width           =   1455
      End
      Begin VB.TextBox txtUnit2 
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C6B8A4&
         Caption         =   "Conversion"
         ForeColor       =   &H00000000&
         Height          =   1335
         Left            =   2520
         TabIndex        =   16
         Top             =   1800
         Width           =   4215
         Begin VB.TextBox txtConUnit2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            TabIndex        =   12
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtConUnit1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1440
            TabIndex        =   13
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lblCon1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "UNIT 1"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1440
            TabIndex        =   19
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblCon2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "UNIT 2"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "="
            Height          =   255
            Left            =   1200
            TabIndex        =   17
            Top             =   600
            Width           =   255
         End
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "POSTING GROUP"
         Height          =   255
         Left            =   0
         TabIndex        =   46
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "REMARKS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   37
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "SRP"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   36
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "COST"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   35
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "MIN QTY"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   34
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "MAX QTY"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "UNIT 1 (INV)"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   32
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "SECTION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "CLASS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   30
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM DESC"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM CODE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "UNIT 2 (PO)"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   26
         Top             =   2160
         Width           =   975
      End
   End
   Begin RPVGCC.b8Container picProgress 
      Height          =   975
      Left            =   1920
      TabIndex        =   47
      Top             =   2760
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1720
      BackColor       =   15396057
      Begin VB.Timer TimerExporttoExcel 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   120
         Top             =   480
      End
      Begin VB.PictureBox picProgressBar 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5235
         TabIndex        =   48
         Top             =   120
         Width           =   5295
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   14
      Top             =   6105
      Width           =   9075
      _ExtentX        =   16007
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
            Picture         =   "frmInvItems.frx":1F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvItems.frx":2C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvItems.frx":3900
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvItems.frx":45DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvItems.frx":52B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvItems.frx":5F8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvItems.frx":6C68
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvItems.frx":7942
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvItems.frx":861C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvItems.frx":8EF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvItems.frx":9BD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvItems.frx":A8AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvItems.frx":B584
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvItems.frx":C25E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvItems.frx":CF38
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInvItems"
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

Public iSupplier    As Double

Dim iClass          As Double
Dim tmp             As Long
Dim WorkbookName    As String
Dim iWorkSheet      As Integer


Dim sItemCode, Arr, sFileName, i, RowCnt, ColCnt, strRange, HeaderRow, iNo, sSectCode

Private Sub BROWSER(sCode, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If sCode <> "" Then
            s = "SELECT TOP 1 tbl_Inv_Items.PK, tbl_Inv_Items.ItemCode, tbl_Inv_Items.ItemDesc, " & _
                " tbl_Inv_Items.SuppKey, tbl_Inv_Items.ClassKey, tbl_Inv_Items.Unit, tbl_Inv_Items.Unit2, " & _
                " tbl_Inv_Items.ConUnit, tbl_Inv_Items.ConUnit2, tbl_Inv_Items.Cost, tbl_Inv_Items.SRP, " & _
                " tbl_Inv_Items.MaxQty, tbl_Inv_Items.MinQty, tbl_Inv_Items.Remarks, tbl_Inv_Items.LastModified, " & _
                " tbl_Inv_Supplier.SupplierCode, tbl_Inv_Supplier.SupplierName , tbl_Inv_Class.ClassCode, " & _
                " tbl_Inv_Class.ClassName, tbl_Inv_Section.SectCode, tbl_Inv_Section.SectName, tbl_Inv_Supplier.Type, " & _
                " tbl_Inv_Items.PostingGroup, tbl_Inv_Items.NonVAT " & _
                " FROM tbl_Inv_Items LEFT OUTER JOIN " & _
                " tbl_Inv_Supplier ON tbl_Inv_Items.SuppKey = tbl_Inv_Supplier.PK LEFT OUTER JOIN " & _
                " tbl_Inv_Class ON tbl_Inv_Items.ClassKey = tbl_Inv_Class.PK LEFT OUTER JOIN " & _
                " tbl_Inv_Section ON tbl_Inv_Class.SectKey = tbl_Inv_Section.PK " & _
                " WHERE (tbl_Inv_Items.ItemCode = '" & sCode & "') " & _
                " ORDER BY tbl_Inv_Items.ItemCode"
        Else
            s = "SELECT TOP 1 tbl_Inv_Items.PK, tbl_Inv_Items.ItemCode, tbl_Inv_Items.ItemDesc, " & _
                " tbl_Inv_Items.SuppKey, tbl_Inv_Items.ClassKey, tbl_Inv_Items.Unit, tbl_Inv_Items.Unit2, " & _
                " tbl_Inv_Items.ConUnit, tbl_Inv_Items.ConUnit2, tbl_Inv_Items.Cost, tbl_Inv_Items.SRP, " & _
                " tbl_Inv_Items.MaxQty, tbl_Inv_Items.MinQty, tbl_Inv_Items.Remarks, tbl_Inv_Items.LastModified, " & _
                " tbl_Inv_Supplier.SupplierCode, tbl_Inv_Supplier.SupplierName , tbl_Inv_Class.ClassCode, " & _
                " tbl_Inv_Class.ClassName, tbl_Inv_Section.SectCode, tbl_Inv_Section.SectName, tbl_Inv_Supplier.Type, " & _
                " tbl_Inv_Items.PostingGroup, tbl_Inv_Items.NonVAT " & _
                " FROM tbl_Inv_Items LEFT OUTER JOIN " & _
                " tbl_Inv_Supplier ON tbl_Inv_Items.SuppKey = tbl_Inv_Supplier.PK LEFT OUTER JOIN " & _
                " tbl_Inv_Class ON tbl_Inv_Items.ClassKey = tbl_Inv_Class.PK LEFT OUTER JOIN " & _
                " tbl_Inv_Section ON tbl_Inv_Class.SectKey = tbl_Inv_Section.PK " & _
                " ORDER BY tbl_Inv_Items.ItemCode"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If picProgress.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_Items.PK, tbl_Inv_Items.ItemCode, tbl_Inv_Items.ItemDesc, " & _
            " tbl_Inv_Items.SuppKey, tbl_Inv_Items.ClassKey, tbl_Inv_Items.Unit, tbl_Inv_Items.Unit2, " & _
            " tbl_Inv_Items.ConUnit, tbl_Inv_Items.ConUnit2, tbl_Inv_Items.Cost, tbl_Inv_Items.SRP, " & _
            " tbl_Inv_Items.MaxQty, tbl_Inv_Items.MinQty, tbl_Inv_Items.Remarks, tbl_Inv_Items.LastModified, " & _
            " tbl_Inv_Supplier.SupplierCode, tbl_Inv_Supplier.SupplierName , tbl_Inv_Class.ClassCode, " & _
            " tbl_Inv_Class.ClassName, tbl_Inv_Section.SectCode, tbl_Inv_Section.SectName, tbl_Inv_Supplier.Type, " & _
            " tbl_Inv_Items.PostingGroup, tbl_Inv_Items.NonVAT " & _
            " FROM tbl_Inv_Items LEFT OUTER JOIN " & _
            " tbl_Inv_Supplier ON tbl_Inv_Items.SuppKey = tbl_Inv_Supplier.PK LEFT OUTER JOIN " & _
            " tbl_Inv_Class ON tbl_Inv_Items.ClassKey = tbl_Inv_Class.PK LEFT OUTER JOIN " & _
            " tbl_Inv_Section ON tbl_Inv_Class.SectKey = tbl_Inv_Section.PK " & _
            " ORDER BY tbl_Inv_Items.ItemCode"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If picProgress.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_Items.PK, tbl_Inv_Items.ItemCode, tbl_Inv_Items.ItemDesc, " & _
            " tbl_Inv_Items.SuppKey, tbl_Inv_Items.ClassKey, tbl_Inv_Items.Unit, tbl_Inv_Items.Unit2, " & _
            " tbl_Inv_Items.ConUnit, tbl_Inv_Items.ConUnit2, tbl_Inv_Items.Cost, tbl_Inv_Items.SRP, " & _
            " tbl_Inv_Items.MaxQty, tbl_Inv_Items.MinQty, tbl_Inv_Items.Remarks, tbl_Inv_Items.LastModified, " & _
            " tbl_Inv_Supplier.SupplierCode, tbl_Inv_Supplier.SupplierName , tbl_Inv_Class.ClassCode, " & _
            " tbl_Inv_Class.ClassName, tbl_Inv_Section.SectCode, tbl_Inv_Section.SectName, tbl_Inv_Supplier.Type, " & _
            " tbl_Inv_Items.PostingGroup, tbl_Inv_Items.NonVAT " & _
            " FROM tbl_Inv_Items LEFT OUTER JOIN " & _
            " tbl_Inv_Supplier ON tbl_Inv_Items.SuppKey = tbl_Inv_Supplier.PK LEFT OUTER JOIN " & _
            " tbl_Inv_Class ON tbl_Inv_Items.ClassKey = tbl_Inv_Class.PK LEFT OUTER JOIN " & _
            " tbl_Inv_Section ON tbl_Inv_Class.SectKey = tbl_Inv_Section.PK " & _
            " WHERE (tbl_Inv_Items.ItemCode < '" & sCode & "') " & _
            " ORDER BY tbl_Inv_Items.ItemCode DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If picProgress.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_Items.PK, tbl_Inv_Items.ItemCode, tbl_Inv_Items.ItemDesc, " & _
            " tbl_Inv_Items.SuppKey, tbl_Inv_Items.ClassKey, tbl_Inv_Items.Unit, tbl_Inv_Items.Unit2, " & _
            " tbl_Inv_Items.ConUnit, tbl_Inv_Items.ConUnit2, tbl_Inv_Items.Cost, tbl_Inv_Items.SRP, " & _
            " tbl_Inv_Items.MaxQty, tbl_Inv_Items.MinQty, tbl_Inv_Items.Remarks, tbl_Inv_Items.LastModified, " & _
            " tbl_Inv_Supplier.SupplierCode, tbl_Inv_Supplier.SupplierName , tbl_Inv_Class.ClassCode, " & _
            " tbl_Inv_Class.ClassName, tbl_Inv_Section.SectCode, tbl_Inv_Section.SectName, tbl_Inv_Supplier.Type, " & _
            " tbl_Inv_Items.PostingGroup, tbl_Inv_Items.NonVAT " & _
            " FROM tbl_Inv_Items LEFT OUTER JOIN " & _
            " tbl_Inv_Supplier ON tbl_Inv_Items.SuppKey = tbl_Inv_Supplier.PK LEFT OUTER JOIN " & _
            " tbl_Inv_Class ON tbl_Inv_Items.ClassKey = tbl_Inv_Class.PK LEFT OUTER JOIN " & _
            " tbl_Inv_Section ON tbl_Inv_Class.SectKey = tbl_Inv_Section.PK " & _
            " WHERE (tbl_Inv_Items.ItemCode > '" & sCode & "') " & _
            " ORDER BY tbl_Inv_Items.ItemCode "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If picProgress.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_Items.PK, tbl_Inv_Items.ItemCode, tbl_Inv_Items.ItemDesc, " & _
            " tbl_Inv_Items.SuppKey, tbl_Inv_Items.ClassKey, tbl_Inv_Items.Unit, tbl_Inv_Items.Unit2, " & _
            " tbl_Inv_Items.ConUnit, tbl_Inv_Items.ConUnit2, tbl_Inv_Items.Cost, tbl_Inv_Items.SRP, " & _
            " tbl_Inv_Items.MaxQty, tbl_Inv_Items.MinQty, tbl_Inv_Items.Remarks, tbl_Inv_Items.LastModified, " & _
            " tbl_Inv_Supplier.SupplierCode, tbl_Inv_Supplier.SupplierName , tbl_Inv_Class.ClassCode, " & _
            " tbl_Inv_Class.ClassName, tbl_Inv_Section.SectCode, tbl_Inv_Section.SectName, tbl_Inv_Supplier.Type, " & _
            " tbl_Inv_Items.PostingGroup, tbl_Inv_Items.NonVAT " & _
            " FROM tbl_Inv_Items LEFT OUTER JOIN " & _
            " tbl_Inv_Supplier ON tbl_Inv_Items.SuppKey = tbl_Inv_Supplier.PK LEFT OUTER JOIN " & _
            " tbl_Inv_Class ON tbl_Inv_Items.ClassKey = tbl_Inv_Class.PK LEFT OUTER JOIN " & _
            " tbl_Inv_Section ON tbl_Inv_Class.SectKey = tbl_Inv_Section.PK " & _
            " ORDER BY tbl_Inv_Items.ItemCode DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iSupplier = rs!SuppKey
    iClass = rs!ClassKey
    txtItemCode.Text = rs!ItemCode
    txtItemDesc.Text = rs!ItemDesc
    txtSuppCode.Text = rs!SupplierCode
    cmbSuppName.Text = rs!SupplierName
    txtType.Text = rs!Type
    txtSuppKey.Text = rs!SuppKey
    txtClassKey.Text = rs!ClassKey
    txtClassCode.Text = rs!ClassCode
    cmbClassName.Text = rs!ClassName
    txtSectCode.Text = rs!SectCode
    txtSectName.Text = rs!SectName
    txtUnit1.Text = rs!Unit
    txtUnit2.Text = rs!Unit2
    lblCon2.Caption = IIf(Trim(rs!Unit2) = "", "UNIT 2", rs!Unit2)
    lblCon1.Caption = IIf(Trim(rs!Unit2) = "", "UNIT 1", rs!Unit)
    txtConUnit1.Text = IIf(CDbl(rs!ConUnit) = 0, "", Format(rs!ConUnit, "#,##0.00"))
    txtConUnit2.Text = IIf(CDbl(rs!ConUnit2) = 0, "", Format(rs!ConUnit2, "#,##0.00"))
    txtMaxQty.Text = Format(rs!MaxQty, "#,##0.00")
    txtMinQty.Text = Format(rs!MinQty, "#,##0.00")
    txtCost.Text = Format(rs!Cost, "#,##0.00")
    txtSRP.Text = Format(rs!SRP, "#,##0.00")
    txtRemarks.Text = rs!Remarks
    chkNonVAT.Value = rs!NonVAT
    
    If IsNull(rs!PostingGroup) = True Then
        txtAccountCode.Text = ""
        cmbAccountName.Text = ""
        cmbAccountName.ListIndex = -1
    Else
        t = "SELECT tbl_GL_Accounts.* " & _
            " FROM tbl_GL_Accounts " & _
            " WHERE (AccountCode = '" & rs!PostingGroup & "')"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            txtAccountCode.Text = rt!AccountCode
            cmbAccountName.Text = rt!AccountName
        Else
            txtAccountCode.Text = ""
            cmbAccountName.Text = ""
        End If
        rt.Close
    End If
    
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    SaveSetting App.EXEName, "ItemCode", "ItCode", rs!ItemCode
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picProgress.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If AccessRights("Inventory Items", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
'Me.Caption = "Items - New"
End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If picProgress.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If AccessRights("Inventory Items", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_EDITTING
'Me.Caption = "Items - Edit"
End Sub

Private Sub PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If picProgress.Visible = True Then Exit Sub
If AccessRights("Inventory Items", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If

s = "SELECT TOP 1 tbl_Inv_PODet.* " & _
    " FROM tbl_Inv_PODet " & _
    " WHERE (ItemKey = " & Statusbar1.Panels(1).Text & ") " & _
    " AND (Item_FA = 1)"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    MsgBox "CAN'T BE DELETED.. ITEM CURRENTLY PRESENT ON P.O.!                          ", vbCritical, "Error..."
    rs.Close
    Exit Sub
End If
rs.Close

If MsgBox("ARE YOU SURE IN DELETING THIS RECORD?                        ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_Inv_Items WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "ItemCode", "ItCode", ""), "is_PAGEDOWN"
If Trim(txtItemCode.Text) = "" Then BROWSER GetSetting(App.EXEName, "ItemCode", "ItCode", ""), "is_HOME"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If picProgress.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If Trim(txtItemCode.Text) = "" Then MsgBox "Please Supply ItemCode!                 ", vbCritical, "Error...": txtItemCode.SetFocus: HTEXT txtItemCode: Exit Sub
If Trim(txtItemDesc.Text) = "" Then MsgBox "Please Supply Item Description!         ", vbCritical, "Error...": txtItemDesc.SetFocus: HTEXT txtItemDesc: Exit Sub
If RETURNTEXTVALUE(txtSuppKey) = 0 Then MsgBox "Please Select Supplier!             ", vbCritical, "Error...": txtSuppCode.SetFocus: HTEXT txtSuppCode: Exit Sub
If RETURNTEXTVALUE(txtClassKey) = 0 Then MsgBox "Please Select Classification!           ", vbCritical, "Error...": txtClassCode.SetFocus: HTEXT txtClassCode: Exit Sub
If Trim(txtUnit1.Text) = "" Then MsgBox "Please Supply Unit of Measure!             ", vbCritical, "Error...": txtUnit1.SetFocus: HTEXT txtUnit1: Exit Sub
If Trim(txtUnit2.Text) <> "" Then
    If RETURNTEXTVALUE(txtConUnit2) = 0 Then MsgBox "Please Supply Conversion Unit!      ", vbCritical, "Error...": txtConUnit2.SetFocus: HTEXT txtConUnit2: Exit Sub
    If RETURNTEXTVALUE(txtConUnit1) = 0 Then MsgBox "Please Supply Conversion Unit!      ", vbCritical, "Error...": txtConUnit1.SetFocus: HTEXT txtConUnit1: Exit Sub
End If

If Trim(txtAccountCode.Text) = "" Then MsgBox "Please Supply Account Code!                              ", vbCritical, "Error...": txtAccountCode.SetFocus: Exit Sub

s = "SELECT tbl_GL_Accounts.* " & _
    " FROM tbl_GL_Accounts " & _
    " WHERE (AccountCode = '" & Trim(txtAccountCode.Text) & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount = 0 Then
    MsgBox "'" & Trim(txtAccountCode.Text) & "' not found in GL Account!                        ", vbCritical, "Error...": txtAccountCode.SetFocus: Exit Sub
End If
rs.Close

On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    sItemCode = Trim(txtItemCode.Text)
    Do
        s = "SELECT tbl_Inv_Items.* " & _
            " FROM tbl_Inv_Items " & _
            " WHERE (ItemCode = '" & sItemCode & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If
        rs.Close
        sItemCode = CDbl(sItemCode) + 1
    Loop
    
    ConnOmega.Execute "INSERT INTO tbl_Inv_Items " & _
                      " (ItemCode, ItemDesc, SuppKey, " & _
                      " ClassKey, Unit, MaxQty, MinQty, " & _
                      " Cost, SRP, Remarks, LastModified, Unit2, " & _
                      " ConUnit, ConUnit2, PostingGroup, NonVAT) " & _
                      " VALUES('" & sItemCode & "',  " & _
                      " '" & FORMATSQL(Trim(txtItemDesc.Text)) & "',  " & _
                      " " & txtSuppKey.Text & ", " & _
                      " " & txtClassKey.Text & ", '" & Trim(txtUnit1.Text) & "', " & _
                      " " & RETURNTEXTVALUE(txtMaxQty) & ", " & _
                      " " & RETURNTEXTVALUE(txtMinQty) & ", " & _
                      " " & RETURNTEXTVALUE(txtCost) & ",  " & _
                      " " & RETURNTEXTVALUE(txtSRP) & ", " & _
                      " '" & FORMATSQL(Trim(txtRemarks.Text)) & "', " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                      " '" & Trim(txtUnit2.Text) & "', " & RETURNTEXTVALUE(txtConUnit1) & ", " & _
                      " " & RETURNTEXTVALUE(txtConUnit2) & ", " & _
                      " '" & Trim(txtAccountCode.Text) & "', " & _
                      " " & chkNonVAT.Value & ")"
End If
If TRANSACTIONTYPE = is_EDITTING Then
    sItemCode = Trim(txtItemCode.Text)
    ConnOmega.Execute "UPDATE tbl_Inv_Items" & _
                      " SET ItemCode = '" & sItemCode & "', " & _
                      " ItemDesc = '" & FORMATSQL(Trim(txtItemDesc.Text)) & "', " & _
                      " SuppKey = " & txtSuppKey.Text & ", " & _
                      " ClassKey = " & txtClassKey.Text & ", " & _
                      " Unit = '" & Trim(txtUnit1.Text) & "', " & _
                      " Unit2 = '" & Trim(txtUnit2.Text) & "', " & _
                      " ConUnit = '" & RETURNTEXTVALUE(txtConUnit1) & "', " & _
                      " ConUnit2 = '" & RETURNTEXTVALUE(txtConUnit2) & "', " & _
                      " MaxQty = " & RETURNTEXTVALUE(txtMaxQty) & ", " & _
                      " MinQty = " & RETURNTEXTVALUE(txtMinQty) & ", " & _
                      " Cost = " & RETURNTEXTVALUE(txtCost) & ", " & _
                      " SRP = " & RETURNTEXTVALUE(txtSRP) & ", " & _
                      " Remarks = '" & FORMATSQL(Trim(txtRemarks.Text)) & "', " & _
                      " LastModified ='" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                      " PostingGroup = '" & Trim(txtAccountCode.Text) & "', " & _
                      " NonVAT = " & chkNonVAT.Value & " " & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
End If
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
'Me.Caption = "Items - Browse"
BROWSER sItemCode, "is_LOAD"

Select Case gbl_Item_Module
    Case "Purchase Invoice":
        With frmInvPI
            .txtItemCode.Text = sItemCode
            .txtDescription.Text = Trim(txtItemDesc.Text)
            gbl_Item_Module = ""
            Unload Me
            .ZOrder 0
            .txtItemCode.SetFocus
        End With
    Case "Purchase Order":
        With frmInvPO
            .txtItemCode.Text = sItemCode
            .txtItemDesc.Text = Trim(txtItemDesc.Text)
            gbl_Item_Module = ""
            Unload Me
            .ZOrder 0
            .txtItemCode.SetFocus
        End With
End Select


Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picProgress.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
PopupMenu MainFormPopupF.mnuItemFind, , Toolbar1.Buttons(15).Left, Toolbar1.Buttons(15).Top + Toolbar1.Buttons(15).Height
End Sub

Private Sub PRESS_F9()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picProgress.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub

MainFormPopupF.mnuItemReportSections(0).Caption = ""
For i = 1 To MainFormPopupF.mnuItemReportSections.UBound
    Unload MainFormPopupF.mnuItemReportSections(i)
Next i
i = -1
s = "SELECT tbl_Inv_Section.* " & _
    " FROM tbl_Inv_Section " & _
    " ORDER BY SectName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    i = i + 1
    If i = 0 Then
        MainFormPopupF.mnuItemReportSections(i).Caption = rs!SectCode & " - " & rs!SectName
    Else
        Load MainFormPopupF.mnuItemReportSections(i)
        MainFormPopupF.mnuItemReportSections(i).Caption = rs!SectCode & " - " & rs!SectName
    End If
    rs.MoveNext
Wend
rs.Close
PopupMenu MainFormPopupF.mnuItemReport, , Toolbar1.Buttons(17).Left, Toolbar1.Buttons(17).Top + Toolbar1.Buttons(17).Height
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picProgress.Visible = True Then Exit Sub
    If picSearch.Visible = True Then cmdCancel_Click: Exit Sub
    Unload Me
Else
    Select Case gbl_Item_Module
        Case ""
            CLEARTEXT
            LOCKTEXT True
            TOOLBARFUNC 1
            TRANSACTIONTYPE = is_REFRESH
            'Me.Caption = "Items - Browse"
            BROWSER GetSetting(App.EXEName, "ItemCode", "ItCode", ""), "is_LOAD"
            If Trim(txtItemCode.Text) = "" Then BROWSER GetSetting(App.EXEName, "ItemCode", "ItCode", ""), "is_HOME"
        Case "Purchase Invoice": TRANSACTIONTYPE = is_REFRESH: Unload Me
        Case "Purchase Order": TRANSACTIONTYPE = is_REFRESH: Unload Me
    End Select
End If
End Sub

Public Sub CLEARTEXT()
iSupplier = 0
iClass = 0
txtItemCode.Text = ""
txtItemDesc.Text = ""
txtSuppCode.Text = ""
cmbSuppName.Text = ""
txtType.Text = ""
txtSuppKey.Text = ""
txtClassKey.Text = ""
cmbSuppName.ListIndex = -1
txtClassCode.Text = ""
cmbClassName.Text = ""
cmbClassName.ListIndex = -1
txtSectCode.Text = ""
txtSectName.Text = ""
txtUnit1.Text = ""
txtUnit2.Text = ""
txtConUnit1.Text = ""
txtConUnit2.Text = ""
txtMaxQty.Text = ""
txtMinQty.Text = ""
txtCost.Text = ""
txtSRP.Text = ""
txtRemarks.Text = ""
txtAccountCode.Text = ""
cmbAccountName.Text = ""
cmbAccountName.ListIndex = -1
chkNonVAT.Value = 0
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
End Sub

Public Sub LOCKTEXT(bln As Boolean)
txtItemCode.Locked = True
txtItemDesc.Locked = bln
txtSuppCode.Locked = bln
cmbSuppName.Locked = bln
txtClassCode.Locked = bln
cmbClassName.Locked = bln
txtSectCode.Locked = True
txtSectName.Locked = True
txtUnit1.Locked = bln
txtUnit2.Locked = bln
txtConUnit1.Locked = bln
txtConUnit2.Locked = bln
txtMaxQty.Locked = bln
txtMinQty.Locked = bln
txtCost.Locked = bln
txtSRP.Locked = bln
txtRemarks.Locked = bln
txtAccountCode.Locked = bln
cmbAccountName.Locked = bln
picNonVAT.Enabled = IIf(bln = True, False, True)
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

Private Sub b8TitleBar2_CLoseClick()
cmdCancel_Click
End Sub

Private Sub cmbAccountName_Click()
If cmbAccountName.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    txtAccountCode.Text = ""
    t = "SELECT tbl_GL_Accounts.* " & _
        " FROM tbl_GL_Accounts " & _
        " WHERE (PK = " & cmbAccountName.ItemData(cmbAccountName.ListIndex) & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtAccountCode.Text = rt!AccountCode
    End If
    rt.Close
End If
End Sub

Private Sub cmbClassName_Click()
If cmbClassName.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    s = "SELECT tbl_Inv_Class.PK, tbl_Inv_Class.ClassCode, " & _
        " tbl_Inv_Class.ClassName, tbl_Inv_Class.SectKey, " & _
        " tbl_Inv_Section.SectCode, tbl_Inv_Section.SectName " & _
        " FROM tbl_Inv_Class LEFT OUTER JOIN " & _
        " tbl_Inv_Section ON tbl_Inv_Class.SectKey = tbl_Inv_Section.PK " & _
        " WHERE (tbl_Inv_Class.PK = " & cmbClassName.ItemData(cmbClassName.ListIndex) & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        txtClassKey.Text = rs!PK
        txtClassCode.Text = rs!ClassCode
        txtSectCode.Text = rs!SectCode
        txtSectName.Text = rs!SectName
    End If
    rs.Close
End If
End Sub

Private Sub cmbSuppName_Click()
If cmbSuppName.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Then
    s = "SELECT PK, SupplierCode, " & _
        " SupplierName, Type " & _
        " From tbl_Inv_Supplier " & _
        " WHERE (PK = " & cmbSuppName.ItemData(cmbSuppName.ListIndex) & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        txtSuppKey.Text = rs!PK
        txtSuppCode.Text = rs!SupplierCode
        txtType.Text = rs!Type
        t = "SELECT TOP 1 ItemCode" & _
            " From tbl_Inv_Items " & _
            " Where (SuppKey = " & rs!PK & ") " & _
            " ORDER BY ItemCode DESC"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            txtItemCode.Text = CStr(CDbl(rt!ItemCode) + 1)
        Else
            txtItemCode.Text = CStr(CDbl(rs!SupplierCode)) & "001"
        End If
        rt.Close
    End If
    rs.Close
ElseIf TRANSACTIONTYPE = is_EDITTING Then
    s = "SELECT PK, SupplierCode, " & _
        " SupplierName, Type " & _
        " From tbl_Inv_Supplier " & _
        " WHERE (PK = " & cmbSuppName.ItemData(cmbSuppName.ListIndex) & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        txtSuppKey.Text = rs!PK
        txtSuppCode.Text = rs!SupplierCode
        txtType.Text = rs!Type
        If iSupplier <> CLng(txtSuppKey.Text) Then
            t = "SELECT TOP 1 ItemCode" & _
                " From tbl_Inv_Items " & _
                " Where (SuppKey = " & rs!PK & ") " & _
                " ORDER BY ItemCode DESC"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                txtItemCode.Text = CStr(CDbl(rt!ItemCode) + 1)
            Else
                txtItemCode.Text = CStr(CDbl(rs!SupplierCode)) & "001"
            End If
            rt.Close
        End If
    End If
    rs.Close
End If
End Sub

Private Sub cmdCancel_Click()
picSearch.Visible = False
picBody.Enabled = True
picToolbar.Enabled = True
End Sub

Private Sub cmdOK_Click()
If lstResult.ListIndex = -1 Then Exit Sub
Arr = Split(lstResult.List(lstResult.ListIndex), " - ", -1, 1)
BROWSER CStr(Arr(0)), "is_LOAD"
cmdCancel_Click
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
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "ItemCode", "ItCode", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "ItemCode", "ItCode", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "ItemCode", "ItCode", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "ItemCode", "ItCode", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
'Me.Caption = "Items - Browse"
POPULATE_COMBO "PK", "SupplierName", "tbl_Inv_Supplier", "SupplierName", cmbSuppName
POPULATE_COMBO "PK", "ClassName", "tbl_Inv_Class", "ClassName", cmbClassName
POPULATE_COMBO_EXEMPTION "PK", "AccountName", "tbl_GL_Accounts", "AccountName", "Inventory", "" & 1 & "", cmbAccountName

Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "ItemCode", "ItCode", ""), "is_LOAD"
If Trim(txtItemCode.Text) = "" Then BROWSER GetSetting(App.EXEName, "ItemCode", "ItCode", ""), "is_HOME"

tmp = SetWindowLong(txtItemDesc.hwnd, GWL_STYLE, GetWindowLong(txtItemDesc.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtUnit1.hwnd, GWL_STYLE, GetWindowLong(txtUnit1.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtUnit2.hwnd, GWL_STYLE, GetWindowLong(txtUnit2.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSectCode.hwnd, GWL_STYLE, GetWindowLong(txtSectCode.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSectName.hwnd, GWL_STYLE, GetWindowLong(txtSectName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtRemarks.hwnd, GWL_STYLE, GetWindowLong(txtRemarks.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picProgress.Visible = True Then Cancel = -1
If picSearch.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstResult_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub TimerExporttoExcel_Timer()
TimerExporttoExcel.Enabled = False


MainForm.CommonDialog1.CancelError = True
On Error GoTo ErrorHandler
MainForm.CommonDialog1.DialogTitle = "Save"
MainForm.CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
MainForm.CommonDialog1.ShowSave
sFileName = Trim(MainForm.CommonDialog1.Filename)

On Error GoTo PG:

picProgressBar.BackColor = &HFFFFFF
picProgress.ZOrder 0
picProgress.Visible = True

WorkbookName = sFileName
Set xlsApp = CreateObject("Excel.Application")
xlsApp.Visible = False
xlsApp.Workbooks.Add
xlsApp.Workbooks(1).Sheets(2).Delete
xlsApp.Workbooks(1).Sheets(2).Delete
xlsApp.DisplayAlerts = False

s = "SELECT tbl_Inv_Section.* " & _
    " FROM tbl_Inv_Section"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
For i = 1 To rs.RecordCount
    xlsApp.Workbooks(1).Sheets.Add
Next i
rs.Close
xlsApp.Workbooks(1).Sheets(1).Delete

iWorkSheet = 0: sSectCode = ""
s = "SELECT tbl_Inv_Section.SectCode, tbl_Inv_Section.SectName, " & _
    " tbl_Inv_Class.ClassCode, tbl_Inv_Class.ClassName, " & _
    " tbl_Inv_Items.ItemCode , tbl_Inv_Items.ItemDesc, " & _
    " tbl_Inv_Items.Unit, tbl_Inv_Items.Cost " & _
    " FROM tbl_Inv_Items LEFT OUTER JOIN " & _
    " tbl_Inv_Class ON tbl_Inv_Items.ClassKey = tbl_Inv_Class.PK LEFT OUTER JOIN " & _
    " tbl_Inv_Section ON tbl_Inv_Class.SectKey = tbl_Inv_Section.PK " & _
    " ORDER BY tbl_Inv_Section.SectName, tbl_Inv_Class.ClassCode, tbl_Inv_Items.ItemCode"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    i = i + 1
    If Trim(CStr(sSectCode)) <> Trim(rs!SectCode) Then
        If CStr(sSectCode) <> "" Then
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)
        End If
        sSectCode = rs!SectCode
        iWorkSheet = iWorkSheet + 1
        RowCnt = 0: HeaderRow = 0
        With xlsApp.Workbooks(1).Sheets(iWorkSheet)
            .Activate
            .Name = rs!SectName
            RowCnt = RowCnt + 1
            HeaderRow = HeaderRow + 1
            ColCnt = 0
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = "ITEMCODE"
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 10
            .Range(strRange).Font.Bold = True
            .Range(strRange).HorizontalAlignment = 3
            .Columns(ColCnt).ColumnWidth = 10
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = "ITEM DESCRIPTION"
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 10
            .Range(strRange).Font.Bold = True
            .Range(strRange).HorizontalAlignment = 2
            .Columns(ColCnt).ColumnWidth = 55
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = "UNIT"
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 10
            .Range(strRange).Font.Bold = True
            .Range(strRange).HorizontalAlignment = 3
            .Columns(ColCnt).ColumnWidth = 10
            
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = "COST"
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 10
            .Range(strRange).Font.Bold = True
            .Columns(ColCnt).ColumnWidth = 15
            .Range(strRange).HorizontalAlignment = 4
        End With
    End If
    
    With xlsApp.ActiveWorkbook.Sheets(iWorkSheet)
        RowCnt = RowCnt + 1
        ColCnt = 0
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = rs!ItemCode
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 10
        .Range(strRange).Font.Bold = False

        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = rs!ItemDesc
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 10
        .Range(strRange).Font.Bold = False

        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = rs!Unit
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 10
        .Range(strRange).Font.Bold = False

        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        .Range(strRange).Value = rs!Cost
        .Range(strRange).Font.Name = "Tahoma"
        .Range(strRange).Font.Size = 10
        .Range(strRange).Font.Bold = False
        .Range(strRange).NumberFormat = "#,##0.00"
    End With
    
    UpdateProgress picProgressBar, i / rs.RecordCount
    rs.MoveNext
Wend
rs.Close

xlsApp.ActiveWorkbook.Sheets(iWorkSheet).PageSetup.PrintTitleRows = "$1" & ":$" & CStr(HeaderRow)

SAVING1:
On Error GoTo err_saving1:
If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

xlsApp.Visible = True
picProgress.Visible = False
Exit Sub
err_saving1:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING1:

Exit Sub
ErrorHandler:
Exit Sub

Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "ItemCode", "ItCode", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "ItemCode", "ItCode", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "ItemCode", "ItCode", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "ItemCode", "ItCode", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":   PRESS_F9
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtAccountCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If TRANSACTIONTYPE = is_ADDING Or _
    TRANSACTIONTYPE = is_EDITTING Then
        If Trim(txtAccountCode.Text) = "" Then Exit Sub
        t = "SELECT tbl_GL_Accounts.* " & _
            " FROM tbl_GL_Accounts " & _
            " WHERE (AccountCode = '" & Trim(txtAccountCode.Text) & "')"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            cmbAccountName.Text = rt!AccountName
        End If
        rt.Close
    End If
End If
End Sub

Private Sub txtAccountCode_LostFocus()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If Trim(txtAccountCode.Text) = "" Then Exit Sub
    t = "SELECT tbl_GL_Accounts.* " & _
        " FROM tbl_GL_Accounts " & _
        " WHERE (AccountCode = '" & Trim(txtAccountCode.Text) & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        cmbAccountName.Text = rt!AccountName
    End If
    rt.Close
End If
End Sub

Private Sub txtClassCode_LostFocus()
If Trim(txtClassCode.Text) = "" Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    s = "SELECT tbl_Inv_Class.PK, tbl_Inv_Class.ClassCode, " & _
        " tbl_Inv_Class.ClassName, tbl_Inv_Class.SectKey, " & _
        " tbl_Inv_Section.SectCode, tbl_Inv_Section.SectName " & _
        " FROM tbl_Inv_Class LEFT OUTER JOIN " & _
        " tbl_Inv_Section ON tbl_Inv_Class.SectKey = tbl_Inv_Section.PK " & _
        " WHERE (tbl_Inv_Class.ClassCode = '" & Format(Trim(txtClassCode.Text), "00#") & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        txtClassKey.Text = rs!PK
        txtClassCode.Text = rs!ClassCode
        cmbClassName.Text = rs!ClassName
        txtSectCode.Text = rs!SectCode
        txtSectName.Text = rs!SectName
    Else
        MsgBox "CLASS NOT FOUND!            ", vbCritical, "Error..."
        txtClassKey.Text = ""
        cmbClassName.Text = ""
        txtSectCode.Text = ""
        txtSectName.Text = ""
        txtClassCode.SetFocus
        HTEXT txtClassCode
        Exit Sub
    End If
    rs.Close
End If
End Sub

Private Sub txtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If TRANSACTIONTYPE = is_FINDING Then
        u = "SELECT tbl_Inv_Items.* " & _
            " FROM tbl_Inv_Items " & _
            " WHERE (ItemCode = '" & Trim(txtItemCode.Text) & "')"
        If ru.State = adStateOpen Then ru.Close
        ru.Open u, ConnOmega
        If ru.RecordCount = 0 Then
            MsgBox "'" & Trim(txtItemCode.Text) & "' Not Found!                      ", vbCritical, "Error..."
            Exit Sub
        End If
        ru.Close
        LOCKTEXT True
        TOOLBARFUNC 1
        TRANSACTIONTYPE = is_REFRESH
        BROWSER Trim(txtItemCode.Text), "is_LOAD"
    End If
End If
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstResult.Clear: Exit Sub
lstResult.Clear
s = "SELECT PK, ItemCode, ItemDesc" & _
    " From tbl_Inv_Items " & _
    " WHERE (ItemDesc LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
    " ORDER BY ItemDesc"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResult.AddItem rs!ItemCode & " - " & rs!ItemDesc
    lstResult.ItemData(lstResult.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResult.ListCount Then lstResult.ListIndex = 0
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then lstResult.SetFocus
End Sub

Private Sub txtSuppCode_LostFocus()
If Trim(txtSuppCode.Text) = "" Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Then
    s = "SELECT PK, SupplierCode, " & _
        " SupplierName, Type " & _
        " From tbl_Inv_Supplier " & _
        " WHERE (SupplierCode = '" & Trim(txtSuppCode.Text) & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        txtSuppKey.Text = rs!PK
        txtSuppCode.Text = rs!SupplierCode
        cmbSuppName.Text = rs!SupplierName
        txtType.Text = rs!Type
        t = "SELECT TOP 1 ItemCode" & _
            " From tbl_Inv_Items " & _
            " Where (SuppKey = " & rs!PK & ") " & _
            " ORDER BY ItemCode DESC"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            txtItemCode.Text = CStr(CDbl(rt!ItemCode) + 1)
        Else
            txtItemCode.Text = CStr(CDbl(rs!SupplierCode)) & "001"
        End If
        rt.Close
    Else
        MsgBox "SUPPLIER NOT FOUND!         ", vbCritical, "Error..."
        txtSuppKey.Text = ""
        cmbSuppName.Text = ""
        txtType.Text = ""
        txtItemCode.Text = ""
        txtSuppCode.SetFocus
        HTEXT txtSuppCode
        Exit Sub
    End If
    rs.Close
End If
If TRANSACTIONTYPE = is_EDITTING Then
    s = "SELECT PK, SupplierCode, " & _
        " SupplierName, Type " & _
        " From tbl_Inv_Supplier " & _
        " WHERE (SupplierCode = '" & Trim(txtSuppCode.Text) & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        txtSuppKey.Text = rs!PK
        txtSuppCode.Text = rs!SupplierCode
        cmbSuppName.Text = rs!SupplierName
        txtType.Text = rs!Type
        If iSupplier <> CLng(txtSuppKey.Text) Then
            t = "SELECT TOP 1 ItemCode" & _
                " From tbl_Inv_Items " & _
                " Where (SuppKey = " & rs!PK & ") " & _
                " ORDER BY ItemCode DESC"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                txtItemCode.Text = CStr(CDbl(rt!ItemCode) + 1)
            Else
                txtItemCode.Text = CStr(CDbl(rs!SupplierCode)) & "001"
            End If
            rt.Close
        End If
    Else
        MsgBox "SUPPLIER NOT FOUND!         ", vbCritical, "Error..."
        txtSuppKey.Text = ""
        cmbSuppName.Text = ""
        txtType.Text = ""
        txtItemCode.Text = ""
        txtSuppCode.SetFocus
        HTEXT txtSuppCode
        Exit Sub
    End If
    rs.Close
End If
End Sub

Private Sub txtUnit1_Change()
If Trim(txtUnit2.Text) <> "" Then
    If TRANSACTIONTYPE = is_ADDING Or _
    TRANSACTIONTYPE = is_EDITTING Then
        txtConUnit2.Locked = False
        txtConUnit1.Locked = False
        lblCon2.Caption = Trim(txtUnit2.Text)
        lblCon1.Caption = Trim(txtUnit1.Text)
    End If
Else
    txtConUnit2.Locked = True
    txtConUnit1.Locked = True
    txtConUnit2.Text = ""
    txtConUnit1.Text = ""
    lblCon2.Caption = "UNIT 2"
    lblCon1.Caption = "UNIT 1"
End If
End Sub

Private Sub txtUnit2_Change()
If Trim(txtUnit2.Text) = "" Then
    txtConUnit2.Locked = True
    txtConUnit1.Locked = True
    txtConUnit2.Text = ""
    txtConUnit1.Text = ""
    lblCon2.Caption = "UNIT 2"
    lblCon1.Caption = "UNIT 1"
Else
    If TRANSACTIONTYPE = is_ADDING Or _
    TRANSACTIONTYPE = is_EDITTING Then
        txtConUnit2.Locked = False
        txtConUnit1.Locked = False
        lblCon2.Caption = Trim(txtUnit2.Text)
        lblCon1.Caption = Trim(txtUnit1.Text)
    End If
End If
End Sub
