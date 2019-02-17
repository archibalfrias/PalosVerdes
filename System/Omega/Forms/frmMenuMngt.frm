VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenuMngt 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13665
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMenuMngt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   13665
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12240
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuMngt.frx":1CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuMngt.frx":1DFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuMngt.frx":1F80
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuMngt.frx":229A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuMngt.frx":2653
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuMngt.frx":2AA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuMngt.frx":2EF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuMngt.frx":32AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuMngt.frx":37F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuMngt.frx":3903
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuMngt.frx":3A5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuMngt.frx":3BB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuMngt.frx":40F9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   770
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   15000
      TabIndex        =   13
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   570
         Left            =   0
         TabIndex        =   14
         Top             =   105
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   1005
         ButtonWidth     =   1217
         ButtonHeight    =   1005
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
               ImageIndex      =   9
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Print"
               Key             =   "Print"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               Key             =   "Refresh"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         MouseIcon       =   "frmMenuMngt.frx":431D
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
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   5925
      Width           =   13665
      _ExtentX        =   24104
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
   Begin RPVGCC.b8Container picMSLine 
      Height          =   855
      Left            =   240
      TabIndex        =   15
      Top             =   1320
      Visible         =   0   'False
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox txtItemDesc 
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   360
         Width           =   4395
      End
      Begin VB.TextBox txtTotalDetCost1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4080
         TabIndex        =   27
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtTotalDetCost 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8760
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   360
         Width           =   1155
      End
      Begin VB.TextBox txtDetCost1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3600
         TabIndex        =   25
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtDetCost 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   360
         Width           =   1155
      End
      Begin VB.TextBox txtItemCode 
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1035
      End
      Begin VB.TextBox txtUnit 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7800
         TabIndex        =   21
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txtQty1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtUnit1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtItemCode1 
         Height          =   285
         Left            =   2880
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtItemDesc1 
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtItemKey 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtItemKey1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cost"
         Height          =   255
         Left            =   8760
         TabIndex        =   34
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cost"
         Height          =   255
         Left            =   6600
         TabIndex        =   33
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "ItemCode"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Description"
         Height          =   255
         Left            =   1200
         TabIndex        =   31
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   255
         Left            =   5640
         TabIndex        =   30
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   255
         Left            =   7800
         TabIndex        =   29
         Top             =   120
         Width           =   915
      End
   End
   Begin RPVGCC.b8Container picPSline 
      Height          =   855
      Left            =   600
      TabIndex        =   36
      Top             =   3480
      Visible         =   0   'False
      Width           =   12330
      _ExtentX        =   21749
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox txtSRPTmp 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10080
         TabIndex        =   69
         Top             =   600
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtEmpRate1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   11280
         TabIndex        =   66
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtMemberRate1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10080
         TabIndex        =   65
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtSRP1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8880
         TabIndex        =   64
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtSrvcCharge1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7560
         TabIndex        =   63
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtLocalTax1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6240
         TabIndex        =   62
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtVAT1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5040
         TabIndex        =   61
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtAdjustment1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4080
         TabIndex        =   60
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtMarkUp1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         TabIndex        =   59
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtCost1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   58
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtEffectDate1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   720
         TabIndex        =   57
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtEmpRate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10800
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   360
         Width           =   1395
      End
      Begin VB.TextBox txtMemberRate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9600
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   360
         Width           =   1155
      End
      Begin VB.TextBox txtSRP 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   360
         Width           =   1035
      End
      Begin VB.TextBox txtSrvcCharge 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7080
         TabIndex        =   49
         Top             =   360
         Width           =   1395
      End
      Begin VB.TextBox txtLocalTax 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5760
         TabIndex        =   47
         Top             =   360
         Width           =   1275
      End
      Begin VB.TextBox txtVAT 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4680
         TabIndex        =   45
         Top             =   360
         Width           =   1035
      End
      Begin VB.TextBox txtAdjustment 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3480
         TabIndex        =   43
         Top             =   360
         Width           =   1155
      End
      Begin VB.TextBox txtMarkUp 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2280
         TabIndex        =   41
         Top             =   360
         Width           =   1155
      End
      Begin VB.TextBox txtCost 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   360
         Width           =   1035
      End
      Begin VB.TextBox txtEffectDate 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Rate"
         Height          =   255
         Left            =   10800
         TabIndex        =   56
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Member Rate"
         Height          =   255
         Left            =   9600
         TabIndex        =   54
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SRP"
         Height          =   255
         Left            =   8520
         TabIndex        =   52
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Srvc Charge (%)"
         Height          =   255
         Left            =   7080
         TabIndex        =   50
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Local Tax (%)"
         Height          =   255
         Left            =   5760
         TabIndex        =   48
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "VAT (%)"
         Height          =   255
         Left            =   4680
         TabIndex        =   46
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Adjustment"
         Height          =   255
         Left            =   3480
         TabIndex        =   44
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Markup (%)"
         Height          =   255
         Left            =   2280
         TabIndex        =   42
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cost"
         Height          =   255
         Left            =   1200
         TabIndex        =   40
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Effect Date"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox picBody 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   120
      ScaleHeight     =   4815
      ScaleWidth      =   13335
      TabIndex        =   5
      Top             =   960
      Width           =   13335
      Begin VB.ComboBox cmbCategory 
         Height          =   315
         Left            =   3240
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   0
         Width           =   6975
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2940
         Left            =   10260
         ScaleHeight     =   2910
         ScaleWidth      =   3030
         TabIndex        =   9
         Top             =   0
         Width           =   3060
         Begin VB.Line Line2 
            X1              =   2760
            X2              =   0
            Y1              =   0
            Y2              =   2880
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   3000
            Y1              =   0
            Y2              =   2880
         End
         Begin VB.Image imgPicture 
            Height          =   2910
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   3030
         End
      End
      Begin VB.TextBox txtCode 
         Height          =   315
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         Top             =   0
         Width           =   1095
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   1
         Top             =   355
         Width           =   9015
      End
      Begin MSComctlLib.ListView lstMaterials 
         Height          =   1620
         Left            =   0
         TabIndex        =   3
         Top             =   1080
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   2858
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
            Text            =   "ItemKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ItemCode"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Item Description"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Unit"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Unit Cost"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Total Cost"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView lstPrice 
         Height          =   1425
         Left            =   0
         TabIndex        =   10
         Top             =   3360
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   2514
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
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Line"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Effect Date"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cost"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Mark Up (%)"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Adjustment"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "VAT (%)"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Local Tax (%)"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Srvc Charge (%)"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "SRP"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "Member Rate"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "Employee Rate"
            Object.Width           =   2470
         EndProperty
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL >>"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7920
         TabIndex        =   68
         Top             =   2760
         Width           =   855
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
         Left            =   8760
         TabIndex        =   67
         Top             =   2760
         Width           =   1130
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "PRICING"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   12
         Top             =   3050
         Width           =   13335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "RAW MATERIALS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   11
         Top             =   800
         Width           =   10215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmMenuMngt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim DET_TRANS As Long
Const is_DET_REFRESH = 0
Const is_DET_ADDING = 1
Const is_DET_EDITTING = 2

Dim iCategory, iPK, x, i, j, _
Filename, sCode

Dim dblMarkUp, dblVat, dblSRP, dblService, dblMarkUpSum, _
dblVatSum, dblLocalTax, dblLocalTaxSum, dblServiceSum

Dim dEmployeePer As Double
Dim dMemberPer As Double
Dim dEffectDate As Date

Dim sPicture As String

Dim isFocusMat As Long
Dim isFocusPri As Long
Dim iRow As Long
Dim tmp As Long


Private Sub BROWSER(sCode, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If sCode <> "" Then
            s = "SELECT TOP 1 tbl_Menu.* " & _
                " From tbl_Menu " & _
                " WHERE (Code = '" & sCode & "') " & _
                " ORDER BY Code"
        Else
            s = "SELECT TOP 1 tbl_Menu.* " & _
                " From tbl_Menu " & _
                " ORDER BY Code"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picMSLine.Visible = True Then Exit Sub
        If picPSline.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Menu.* " & _
            " From tbl_Menu " & _
            " ORDER BY Code"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picMSLine.Visible = True Then Exit Sub
        If picPSline.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Menu.* " & _
            " From tbl_Menu " & _
            " WHERE (Code < '" & sCode & "') " & _
            " ORDER BY Code DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picMSLine.Visible = True Then Exit Sub
        If picPSline.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Menu.* " & _
            " From tbl_Menu " & _
            " WHERE (Code > '" & sCode & "') " & _
            " ORDER BY Code"
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picMSLine.Visible = True Then Exit Sub
        If picPSline.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Menu.* " & _
            " From tbl_Menu " & _
            " ORDER BY Code DESC"
    Case Else: Exit Sub
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iCategory = rs!Category
    cmbCategory.Text = ""
    t = "SELECT Category " & _
        " FROM tbl_Menu_Category " & _
        " WHERE (PK = " & rs!Category & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        cmbCategory.Text = rt!Category
    End If
    rt.Close
    
    txtCode.Text = rs!Code
    txtName.Text = rs!Name
    imgPicture.ZOrder 0
    If IsNull(rs!Picture) = False Then
        imgPicture.Picture = LoadPicture(SHOW_IMAGES(rs!PK, 0, "Menu Management"))
    Else
        imgPicture.Picture = LoadPicture("")
    End If
    
    SaveSetting App.EXEName, "MenuManagementCode", "MenuMngtCode", rs!Code
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    
    t = "SELECT tbl_Menu_Raw_Materials.MenuKey, tbl_Menu_Raw_Materials.Line, " & _
        " tbl_Menu_Raw_Materials.ItemKey, tbl_Inv_Items.ItemCode, " & _
        " tbl_Inv_Items.ItemDesc, tbl_Inv_Items.Unit, tbl_Menu_Raw_Materials.Cost, " & _
        " tbl_Menu_Raw_Materials.Qty, tbl_Menu_Raw_Materials.TotalCost " & _
        " FROM tbl_Menu_Raw_Materials LEFT OUTER JOIN " & _
        " tbl_Inv_Items ON tbl_Menu_Raw_Materials.ItemKey = tbl_Inv_Items.PK " & _
        " WHERE (tbl_Menu_Raw_Materials.MenuKey = " & rs!PK & ") " & _
        " ORDER BY tbl_Menu_Raw_Materials.Line"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstMaterials.ListItems.Clear
        i = 0
        While Not rt.EOF
            i = i + 1
            Set x = lstMaterials.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = Format(i, "0#")
            x.SubItems(2) = rt!ItemKey
            x.SubItems(3) = rt!ItemCode
            x.SubItems(4) = rt!ItemDesc
            x.SubItems(5) = rt!Unit
            x.SubItems(6) = Format(rt!Cost, "#,##0.00")
            x.SubItems(7) = rt!Qty
            x.SubItems(8) = Format(rt!TotalCost, "#,##0.00")
            rt.MoveNext
        Wend
    Else
        lstMaterials.ListItems.Clear
        Set x = lstMaterials.ListItems.Add()
        x.Text = ""
        x.SubItems(1) = " "
        x.SubItems(2) = "0"
        x.SubItems(3) = " "
        x.SubItems(4) = " "
        x.SubItems(5) = " "
        x.SubItems(6) = " "
        x.SubItems(7) = " "
        x.SubItems(8) = " "
    End If
    rt.Close
    
    t = "SELECT MenuKey, EffectDate, Cost, MarkUp, Vat, LocalTax, " & _
        " ServiceCharge, Adjusted, SRP, MemberPrice, EmployeePrice " & _
        " From tbl_Menu_Pricing " & _
        " WHERE (MenuKey = " & rs!PK & ") " & _
        " ORDER BY EffectDate DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        i = 0
        lstPrice.ListItems.Clear
        While Not rt.EOF
            i = i + 1
            Set x = lstPrice.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = i
            x.SubItems(2) = Format(rt!EffectDate, "mm/dd/yyyy")
            x.SubItems(3) = Format(rt!Cost, "#,##0.00")
            x.SubItems(4) = Format(rt!MarkUp, "#,##0.00")
            x.SubItems(5) = Format(rt!Adjusted, "#,##0.00")
            x.SubItems(6) = Format(rt!Vat, "#,##0.00")
            x.SubItems(7) = Format(rt!LocalTax, "#,##0.00")
            x.SubItems(8) = Format(rt!ServiceCharge, "#,##0.00")
            x.SubItems(9) = Format(rt!SRP, "#,##0.00")
            x.SubItems(10) = Format(rt!MemberPrice, "#,##0.00")
            x.SubItems(11) = Format(rt!EmployeePrice, "#,##0.00")
            rt.MoveNext
        Wend
    Else
        lstPrice.ListItems.Clear
        Set x = lstPrice.ListItems.Add()
        x.Text = ""
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
        x.SubItems(11) = " "
    End If
    rt.Close
    
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If AccessRights("Menu Management", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If picMSLine.Visible = True Then Exit Sub
If picPSline.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    CLEARTEXT
    LOCKTEXT False
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_ADDING
    'Me.Caption = "Menu Management - New"
Else
    If isFocusMat = 1 Then
        With lstMaterials.ListItems
            txtItemCode.Text = ""
            txtItemKey.Text = "0"
            txtItemDesc.Text = ""
            txtUnit.Text = ""
            txtDetCost.Text = ""
            txtQty.Text = ""
            txtTotalDetCost.Text = ""
            If CDbl(.Item(.Count).SubItems(2)) <> 0 Then
                Set x = .Add()
                x.SubItems(1) = " "
                x.SubItems(2) = "0"
                x.SubItems(3) = " "
                x.SubItems(4) = " "
                x.SubItems(5) = " "
                x.SubItems(6) = " "
                x.SubItems(7) = " "
                x.SubItems(8) = " "
                iRow = .Count
            End If
            lstMaterials.ListItems(iRow).EnsureVisible
            lstMaterials.ListItems(iRow).Selected = True
            DET_TRANS = is_DET_ADDING
            picBody.Enabled = False
            picMSLine.ZOrder 0
            picMSLine.Visible = True
            txtItemCode.SetFocus
        End With
    End If
    If isFocusPri = 1 Then
        If CDbl(lblTotalCost.Caption) = 0 Then MsgBox "Please Supply Material/s First!                  ", vbCritical, "Error...": Exit Sub
        With lstPrice.ListItems
            txtEffectDate.Text = ""
            txtCost.Text = ""
            txtMarkUp.Text = ""
            txtAdjustment.Text = ""
            txtVAT.Text = ""
            txtLocalTax.Text = ""
            txtSrvcCharge.Text = ""
            txtSRP.Text = ""
            txtMemberRate.Text = ""
            txtEmpRate.Text = ""
            If IsDate(.Item(.Count).SubItems(2)) = True Then
                Set x = .Add()
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
                x.SubItems(11) = " "
                iRow = .Count
            End If
            lstPrice.ListItems(iRow).EnsureVisible
            lstPrice.ListItems(iRow).Selected = True
            DET_TRANS = is_DET_ADDING
            picBody.Enabled = False
            picPSline.ZOrder 0
            picPSline.Visible = True
            txtCost.Text = lblTotalCost.Caption
            txtEffectDate.SetFocus
        End With
    End If
End If
End Sub

Private Sub PRESS_F2()
If AccessRights("Menu Management", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If picMSLine.Visible = True Then Exit Sub
If picPSline.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    If Statusbar1.Panels(1).Text = "" Then Exit Sub
    LOCKTEXT False
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_EDITTING
    'Me.Caption = "Menu Management - Edit"
Else
    If isFocusMat = 1 Then
        With lstMaterials.ListItems
            txtItemKey.Text = .Item(iRow).SubItems(2)
            txtItemCode.Text = .Item(iRow).SubItems(3)
            txtItemDesc.Text = .Item(iRow).SubItems(4)
            txtUnit.Text = .Item(iRow).SubItems(5)
            txtDetCost.Text = .Item(iRow).SubItems(6)
            txtQty.Text = .Item(iRow).SubItems(7)
            txtTotalDetCost.Text = .Item(iRow).SubItems(8)
            
            txtItemKey1.Text = .Item(iRow).SubItems(2)
            txtItemCode1.Text = .Item(iRow).SubItems(3)
            txtItemDesc1.Text = .Item(iRow).SubItems(4)
            txtUnit1.Text = .Item(iRow).SubItems(5)
            txtDetCost1.Text = .Item(iRow).SubItems(6)
            txtQty1.Text = .Item(iRow).SubItems(7)
            txtTotalDetCost1.Text = .Item(iRow).SubItems(8)
            
            DET_TRANS = is_DET_EDITTING
            picBody.Enabled = False
            picMSLine.ZOrder 0
            picMSLine.Visible = True
            txtItemCode.SetFocus
        End With
    End If
    If isFocusPri = 1 Then
        If CDbl(lblTotalCost.Caption) = 0 Then MsgBox "Please Supply Material/s First!                  ", vbCritical, "Error...": Exit Sub
        With lstPrice.ListItems
            txtEffectDate.Text = .Item(iRow).SubItems(2)
            txtCost.Text = .Item(iRow).SubItems(3)
            txtMarkUp.Text = .Item(iRow).SubItems(4)
            txtAdjustment.Text = .Item(iRow).SubItems(5)
            txtVAT.Text = .Item(iRow).SubItems(6)
            txtLocalTax.Text = .Item(iRow).SubItems(7)
            txtSrvcCharge.Text = .Item(iRow).SubItems(8)
            txtSRP.Text = .Item(iRow).SubItems(9)
            txtMemberRate.Text = .Item(iRow).SubItems(10)
            txtEmpRate.Text = .Item(iRow).SubItems(11)
            
            txtEffectDate1.Text = .Item(iRow).SubItems(2)
            txtCost1.Text = .Item(iRow).SubItems(3)
            txtMarkUp1.Text = .Item(iRow).SubItems(4)
            txtAdjustment1.Text = .Item(iRow).SubItems(5)
            txtVAT1.Text = .Item(iRow).SubItems(6)
            txtLocalTax1.Text = .Item(iRow).SubItems(7)
            txtSrvcCharge1.Text = .Item(iRow).SubItems(8)
            txtSRP1.Text = .Item(iRow).SubItems(9)
            txtMemberRate1.Text = .Item(iRow).SubItems(10)
            txtEmpRate1.Text = .Item(iRow).SubItems(11)
            
            DET_TRANS = is_DET_EDITTING
            picBody.Enabled = False
            picPSline.ZOrder 0
            picPSline.Visible = True
            txtCost.Text = lblTotalCost.Caption
            txtEffectDate.SetFocus
        End With
    End If
End If
End Sub

Private Sub PRESS_DELETE()
If AccessRights("Menu Management", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If picMSLine.Visible = True Then Exit Sub
If picPSline.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    If Statusbar1.Panels(1).Text = "" Then Exit Sub
    If MsgBox("ARE SURE IN DELETING THIS RECORD?                    ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    On Error GoTo PG:
    ConnOmega.Execute "DELETE FROM tbl_Menu WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    CLEARTEXT
    BROWSER GetSetting(App.EXEName, "MenuManagementCode", "MenuMngtCode", ""), "is_PAGEDOWN"
    If Trim(txtCode.Text) = "" Then BROWSER GetSetting(App.EXEName, "MenuManagementCode", "MenuMngtCode", ""), "is_HOME"
Else
    If isFocusMat = 1 Then
        With lstMaterials.ListItems
            If .Count > 1 Then
                .Remove iRow
                If iRow > .Count Then
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
                .Item(1).SubItems(8) = " "
                iRow = .Count
            End If
            lstMaterials.ListItems(iRow).EnsureVisible
            lstMaterials.ListItems(iRow).Selected = True
        End With
    End If
    If isFocusPri = 1 Then
        With lstPrice.ListItems
            If .Count > 1 Then
                .Remove iRow
                If iRow > .Count Then
                    iRow = .Count
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
                .Item(1).SubItems(11) = " "
                iRow = .Count
            End If
            lstPrice.ListItems(iRow).EnsureVisible
            lstPrice.ListItems(iRow).Selected = True
        End With
    End If
End If
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If picMSLine.Visible = True Then Exit Sub
If picPSline.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
If Trim(txtCode.Text) = "" Then MsgBox "Please Supply Code!                   ", vbCritical, "Error...": txtCode.SetFocus: Exit Sub
If Trim(txtName.Text) = "" Then MsgBox "Please Supply Name!                       ", vbCritical, "Error...": txtName.SetFocus: Exit Sub
If iCategory = 0 Then MsgBox "Please Select Category!                       ", vbCritical, "Error...": cmbCategory.SetFocus: Exit Sub
On Error GoTo PG:

If TRANSACTIONTYPE = is_ADDING Then
    sCode = Trim(txtCode.Text)
    ConnOmega.Execute "INSERT INTO tbl_Menu " & _
                      " (Code, Name, Category, LastModified) " & _
                      " VALUES ('" & Trim(txtCode.Text) & "', " & _
                      " " & _
                      " '" & FORMATSQL(Trim(txtName.Text)) & "', " & iCategory & ", " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
    
    iPK = 0
    s = "SELECT PK " & _
        " FROM tbl_Menu " & _
        " WHERE (Code = '" & Trim(txtCode.Text) & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        iPK = rs!PK
    End If
    rs.Close
    
End If

If TRANSACTIONTYPE = is_EDITTING Then
    sCode = Trim(txtCode.Text)
    iPK = Statusbar1.Panels(1).Text
    ConnOmega.Execute "UPDATE tbl_Menu " & _
                      " SET Code = '" & Trim(txtCode.Text) & "', " & _
                      " Name = '" & FORMATSQL(Trim(txtName.Text)) & "', " & _
                      " Category = " & iCategory & ", " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & iPK & ")"
End If

If CDbl(iPK) <> 0 Then

    If Trim(CStr(sPicture)) <> "" Then
        SAVE_IMAGES iPK, 0, sPicture, "Menu Management"
    End If
    
    ConnOmega.Execute "DELETE FROM tbl_Menu_Raw_Materials WHERE (MenuKey = " & iPK & ")"
    ConnOmega.Execute "DELETE FROM tbl_Menu_Pricing WHERE (MenuKey = " & iPK & ")"
    
    With lstMaterials.ListItems
        j = 0
        For i = 1 To .Count
            If CDbl(.Item(i).SubItems(2)) <> 0 Then
                j = j + 1
                ConnOmega.Execute "INSERT INTO tbl_Menu_Raw_Materials " & _
                                  " (MenuKey, Line, ItemKey, Cost, Qty) " & _
                                  " VALUES (" & iPK & ", " & j & ", " & _
                                  " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2))) & ", " & _
                                  " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(6)) = False, 0, .Item(i).SubItems(6))) & ", " & _
                                  " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(7)) = False, 0, .Item(i).SubItems(7))) & ")"
            End If
        Next i
    End With
    
    With lstPrice.ListItems
        j = 0
        For i = 1 To .Count
            If IsDate(.Item(i).SubItems(2)) = True Then
                j = j + 1
                ConnOmega.Execute "INSERT INTO tbl_Menu_Pricing " & _
                                  " (MenuKey, EffectDate, Cost, MarkUp, Vat, LocalTax, " & _
                                  " ServiceCharge, Adjusted, SRP, MemberPrice, EmployeePrice) " & _
                                  " VALUES (" & iPK & ",'" & FormatDateTime(.Item(i).SubItems(2), vbShortDate) & "', " & _
                                  " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(3)) = False, 0, .Item(i).SubItems(3))) & ", " & _
                                  " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(4)) = False, 0, .Item(i).SubItems(4))) & ", " & _
                                  " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(5)) = False, 0, .Item(i).SubItems(5))) & ", " & _
                                  " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(6)) = False, 0, .Item(i).SubItems(6))) & ", " & _
                                  " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(7)) = False, 0, .Item(i).SubItems(7))) & ", " & _
                                  " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(8)) = False, 0, .Item(i).SubItems(8))) & ", " & _
                                  " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(9)) = False, 0, .Item(i).SubItems(9))) & ", " & _
                                  " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(10)) = False, 0, .Item(i).SubItems(10))) & ", " & _
                                  " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(11)) = False, 0, .Item(i).SubItems(11))) & ")"
            End If
        Next i
    End With
End If

sPicture = ""
LOCKTEXT True
TOOLBARFUNC 0
TRANSACTIONTYPE = is_REFRESH
DET_TRANS = is_DET_REFRESH
'Me.Caption = "Menu Management - Browse"
BROWSER sCode, "is_LOAD"
txtCode.SetFocus
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()

End Sub

Private Sub PRESS_F11()

End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    Unload Me
Else
    If picMSLine.Visible = True Then
        With lstMaterials.ListItems
            If DET_TRANS = is_DET_ADDING Then
                If .Count > 1 Then
                    .Remove iRow
                    iRow = .Count
                Else
                    .Item(1).SubItems(1) = " "
                    .Item(1).SubItems(2) = "0"
                    .Item(1).SubItems(3) = " "
                    .Item(1).SubItems(4) = " "
                    .Item(1).SubItems(5) = " "
                    .Item(1).SubItems(6) = " "
                    .Item(1).SubItems(7) = " "
                    .Item(1).SubItems(8) = " "
                    iRow = .Count
                End If
            End If
            If DET_TRANS = is_DET_EDITTING Then
                .Item(iRow).SubItems(2) = txtItemKey1.Text
                .Item(iRow).SubItems(3) = txtItemCode1.Text
                .Item(iRow).SubItems(4) = txtItemDesc1.Text
                .Item(iRow).SubItems(5) = txtUnit1.Text
                .Item(iRow).SubItems(6) = txtDetCost1.Text
                .Item(iRow).SubItems(7) = txtQty1.Text
                .Item(iRow).SubItems(8) = txtTotalDetCost1.Text
            End If
        End With
        picMSLine.Visible = False
        picBody.Enabled = True
        lstPrice.SetFocus
        lstPrice.ListItems(iRow).EnsureVisible
        lstPrice.ListItems(iRow).Selected = True
        Exit Sub
    End If
    
    If picPSline.Visible = True Then
        With lstPrice.ListItems
            If DET_TRANS = is_DET_ADDING Then
                If .Count > 1 Then
                    .Remove iRow
                    iRow = .Count
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
                    .Item(1).SubItems(11) = " "
                    iRow = .Count
                End If
            End If
            If DET_TRANS = is_DET_EDITTING Then
                .Item(iRow).SubItems(2) = txtEffectDate1.Text
                .Item(iRow).SubItems(3) = txtCost1.Text
                .Item(iRow).SubItems(4) = txtMarkUp1.Text
                .Item(iRow).SubItems(5) = txtAdjustment1.Text
                .Item(iRow).SubItems(6) = txtVAT1.Text
                .Item(iRow).SubItems(7) = txtLocalTax1.Text
                .Item(iRow).SubItems(8) = txtSrvcCharge1.Text
                .Item(iRow).SubItems(9) = txtSRP1.Text
                .Item(iRow).SubItems(10) = txtMemberRate1.Text
                .Item(iRow).SubItems(11) = txtEmpRate1.Text
            End If
        End With
        picPSline.Visible = False
        picBody.Enabled = True
        lstPrice.SetFocus
        lstPrice.ListItems(iRow).EnsureVisible
        lstPrice.ListItems(iRow).Selected = True
        Exit Sub
    End If
    
    If isFocusMat = 1 Or isFocusPri = 1 Then
        txtCode.SetFocus
        DET_TRANS = is_DET_REFRESH
        Exit Sub
    End If
    
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 0
    TRANSACTIONTYPE = is_REFRESH
    DET_TRANS = is_DET_REFRESH
    'Me.Caption = "Menu Management - Browse"
    
    BROWSER GetSetting(App.EXEName, "MenuManagementCode", "MenuMngtCode", ""), "is_LOAD"
    If Trim(txtCode.Text) = "" Then BROWSER GetSetting(App.EXEName, "MenuManagementCode", "MenuMngtCode", ""), "is_HOME"
End If
End Sub

Private Sub CLEARTEXT()
sPicture = ""
iCategory = 0
txtCode.Text = ""
txtName.Text = ""
cmbCategory.Text = ""
cmbCategory.ListIndex = -1
lblTotalCost.Caption = "0.00"
imgPicture.Picture = LoadPicture("")
lstMaterials.ListItems.Clear
Set x = lstMaterials.ListItems.Add()
x.Text = ""
x.SubItems(1) = " "
x.SubItems(2) = "0"
x.SubItems(3) = " "
x.SubItems(4) = " "
x.SubItems(5) = " "
x.SubItems(6) = " "
x.SubItems(7) = " "
x.SubItems(8) = " "
lstPrice.ListItems.Clear
Set x = lstPrice.ListItems.Add()
x.Text = ""
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
x.SubItems(11) = " "
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtCode.Locked = bln
txtName.Locked = bln
cmbCategory.Locked = bln
End Sub


Private Sub TOOLBARFUNC(iSel As Long)
Set Toolbar1.ImageList = ImageList1
With Toolbar1.Buttons
    Select Case iSel
        Case 0      'Refresh
        
            .Item(1).Image = 1
            .Item(3).Image = 2
            .Item(5).Image = 3
            .Item(11).Image = 6
            .Item(13).Image = 7
            .Item(15).Image = 9
            .Item(17).Image = 8
            .Item(19).Image = 10
            .Item(21).Image = 11
            .Item(1).Enabled = True
            .Item(3).Enabled = True
            .Item(5).Enabled = True
            .Item(7).Image = 4
            .Item(7).Caption = "First"
            .Item(9).Image = 5
            .Item(9).Caption = "Back"
            .Item(7).Enabled = True
            .Item(9).Enabled = True
            .Item(11).Enabled = True
            .Item(13).Enabled = True
            .Item(15).Enabled = True
            .Item(17).Enabled = True
            .Item(19).Enabled = True
            .Item(21).Enabled = True
            .Item(1).ToolTipText = "NEW (Ins)"
            .Item(3).ToolTipText = "EDIT (F2)"
            .Item(5).ToolTipText = "DELETE (Del)"
            .Item(7).ToolTipText = "FIRST (Home)"
            .Item(9).ToolTipText = "BACK (PgUp)"
            .Item(11).ToolTipText = "NEXT (PgDown)"
            .Item(13).ToolTipText = "LAST (End)"
            .Item(15).ToolTipText = "FIND (F6)"
            .Item(17).ToolTipText = "PRINT (F9)"
            .Item(19).ToolTipText = "REFRESH (F8)"
            .Item(21).ToolTipText = "CLOSE (Esc)"
        
        Case 1      'Add/Edit
        
            .Item(1).Image = 1
            .Item(3).Image = 2
            .Item(5).Image = 3
            .Item(11).Image = 6
            .Item(13).Image = 7
            .Item(15).Image = 9
            .Item(17).Image = 8
            .Item(19).Image = 10
            .Item(21).Image = 11
            .Item(1).Enabled = False
            .Item(3).Enabled = False
            .Item(5).Enabled = False
            .Item(7).Image = 12
            .Item(7).Caption = "Save"
            .Item(9).Image = 13
            .Item(9).Caption = "Undo"
            .Item(7).Enabled = True
            .Item(9).Enabled = True
            .Item(11).Enabled = False
            .Item(13).Enabled = False
            .Item(15).Enabled = False
            .Item(17).Enabled = False
            .Item(19).Enabled = False
            .Item(21).Enabled = False
            
            .Item(1).ToolTipText = ""
            .Item(3).ToolTipText = ""
            .Item(5).ToolTipText = ""
            .Item(7).ToolTipText = "SAVE (F5)"
            .Item(9).ToolTipText = "UNDO (Esc)"
            .Item(11).ToolTipText = ""
            .Item(13).ToolTipText = ""
            .Item(15).ToolTipText = ""
            .Item(17).ToolTipText = ""
            .Item(19).ToolTipText = ""
            .Item(21).ToolTipText = ""
        
        Case 2      'Find
        
            .Item(1).Image = 1
            .Item(3).Image = 2
            .Item(5).Image = 3
            .Item(11).Image = 6
            .Item(13).Image = 7
            .Item(15).Image = 9
            .Item(17).Image = 8
            .Item(19).Image = 10
            .Item(21).Image = 11
            .Item(1).Enabled = False
            .Item(3).Enabled = False
            .Item(5).Enabled = False
            .Item(7).Image = 4
            .Item(7).Caption = "First"
            .Item(9).Image = 13
            .Item(9).Caption = "Undo"
            .Item(7).Enabled = False
            .Item(9).Enabled = True
            .Item(11).Enabled = False
            .Item(13).Enabled = False
            .Item(15).Enabled = False
            .Item(17).Enabled = False
            .Item(19).Enabled = False
            .Item(21).Enabled = False
            
            .Item(1).ToolTipText = ""
            .Item(3).ToolTipText = ""
            .Item(5).ToolTipText = ""
            .Item(7).ToolTipText = ""
            .Item(9).ToolTipText = "UNDO (Esc)"
            .Item(11).ToolTipText = ""
            .Item(13).ToolTipText = ""
            .Item(15).ToolTipText = ""
            .Item(17).ToolTipText = ""
            .Item(19).ToolTipText = ""
            .Item(21).ToolTipText = ""
            
        Case 3          'Detail Empty
            
            .Item(1).Image = 1
            .Item(3).Image = 2
            .Item(5).Image = 3
            .Item(11).Image = 6
            .Item(13).Image = 7
            .Item(15).Image = 9
            .Item(17).Image = 8
            .Item(19).Image = 10
            .Item(21).Image = 11
            .Item(1).Enabled = True
            .Item(3).Enabled = False
            .Item(5).Enabled = False
            .Item(7).Image = 12
            .Item(7).Caption = "Save"
            .Item(9).Image = 13
            .Item(9).Caption = "Undo"
            .Item(7).Enabled = True
            .Item(9).Enabled = True
            .Item(11).Enabled = False
            .Item(13).Enabled = False
            .Item(15).Enabled = False
            .Item(17).Enabled = False
            .Item(19).Enabled = False
            .Item(21).Enabled = False
            .Item(1).ToolTipText = "NEW (Ins)"
            .Item(3).ToolTipText = ""
            .Item(5).ToolTipText = ""
            .Item(7).ToolTipText = "SAVE (F5)"
            .Item(9).ToolTipText = "UNDO (Esc)"
            .Item(11).ToolTipText = ""
            .Item(13).ToolTipText = ""
            .Item(15).ToolTipText = ""
            .Item(17).ToolTipText = ""
            .Item(19).ToolTipText = ""
            .Item(21).ToolTipText = ""
        
        Case 4          'Detail not Empty
            
            .Item(1).Image = 1
            .Item(3).Image = 2
            .Item(5).Image = 3
            .Item(11).Image = 6
            .Item(13).Image = 7
            .Item(15).Image = 9
            .Item(17).Image = 8
            .Item(19).Image = 10
            .Item(21).Image = 11
            .Item(1).Enabled = True
            .Item(3).Enabled = True
            .Item(5).Enabled = True
            .Item(7).Image = 12
            .Item(7).Caption = "Save"
            .Item(9).Image = 13
            .Item(9).Caption = "Undo"
            .Item(7).Enabled = True
            .Item(9).Enabled = True
            .Item(11).Enabled = False
            .Item(13).Enabled = False
            .Item(15).Enabled = False
            .Item(17).Enabled = False
            .Item(19).Enabled = False
            .Item(21).Enabled = False
            .Item(1).ToolTipText = "NEW (Ins)"
            .Item(3).ToolTipText = "EDIT (F2)"
            .Item(5).ToolTipText = "DELETE (Del)"
            .Item(7).ToolTipText = "SAVE (F5)"
            .Item(9).ToolTipText = "UNDO (Esc)"
            .Item(11).ToolTipText = ""
            .Item(13).ToolTipText = ""
            .Item(15).ToolTipText = ""
            .Item(17).ToolTipText = ""
            .Item(19).ToolTipText = ""
            .Item(21).ToolTipText = ""
            
    End Select
End With
End Sub

Private Sub cmbCategory_Click()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
iCategory = cmbCategory.ItemData(cmbCategory.ListIndex)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert:   PRESS_INSERT
    Case vbKeyF2:       PRESS_F2
    Case vbKeyDelete:   PRESS_DELETE
    Case vbKeyF5:       PRESS_F5
    Case vbKeyF6:       PRESS_F6
    Case vbKeyF11:      PRESS_F11
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "MenuManagementCode", "MenuMngtCode", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "MenuManagementCode", "MenuMngtCode", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "MenuManagementCode", "MenuMngtCode", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "MenuManagementCode", "MenuMngtCode", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
'Me.Caption = "Menu Management - Browse"
Me.Top = (MainForm.Height - Me.Height) / 10
Me.Left = (MainForm.Width - Me.Width) / 5

cmbCategory.Clear
s = "SELECT tbl_Menu_Category.* " & _
    " FROM tbl_Menu_Category " & _
    " ORDER BY Category"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbCategory.AddItem rs!Category
    cmbCategory.ItemData(cmbCategory.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close

isFocusMat = 0
isFocusPri = 0
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 0
TRANSACTIONTYPE = is_REFRESH
DET_TRANS = is_DET_REFRESH
BROWSER GetSetting(App.EXEName, "MenuManagementCode", "MenuMngtCode", ""), "is_LOAD"
If Trim(txtCode.Text) = "" Then BROWSER GetSetting(App.EXEName, "MenuManagementCode", "MenuMngtCode", ""), "is_HOME"

tmp = SetWindowLong(txtName.hwnd, GWL_STYLE, GetWindowLong(txtName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub


Private Sub Form_Unload(Cancel As Integer)
If picMSLine.Visible = True Then Cancel = -1
If picPSline.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub imgPicture_DblClick()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
MainForm.CommonDialog1.CancelError = True
On Error GoTo ErrorHandler
MainForm.CommonDialog1.Filter = "Image Files|*.JPG;*.JPEG;*.JPE;*.BMP;*.RLE;*.DIB;*.GIF;*.PNG;*.TIF;*.TIFF"
MainForm.CommonDialog1.ShowOpen
Filename = Trim(MainForm.CommonDialog1.Filename)
If ((FileLen(Filename) \ 1024) + 1) > 50 Then
    MsgBox "Image is too large please reduce the size to 50kb or below!          ", vbCritical, "Error..."
    Exit Sub
End If
sPicture = Filename
imgPicture.ZOrder 0
imgPicture.Picture = LoadPicture(Filename)
Exit Sub
ErrorHandler:
Exit Sub
End Sub

Private Sub lstMaterials_GotFocus()
isFocusMat = 1
DET_TRANS = is_DET_REFRESH
iRow = lstMaterials.SelectedItem.Index
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
With lstMaterials.ListItems
    If CDbl(.Item(iRow).SubItems(2)) = 0 Then
        TOOLBARFUNC 3
    Else
        TOOLBARFUNC 4
    End If
End With
End Sub

Private Sub lstMaterials_ItemClick(ByVal Item As MSComctlLib.ListItem)
iRow = lstMaterials.SelectedItem.Index
End Sub

Private Sub lstMaterials_LostFocus()
isFocusMat = 0
End Sub

Private Sub lstPrice_GotFocus()
isFocusPri = 1
DET_TRANS = is_DET_REFRESH
iRow = lstPrice.SelectedItem.Index
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
With lstPrice.ListItems
    If IsDate(.Item(iRow).SubItems(2)) = False Then
        TOOLBARFUNC 3
    Else
        TOOLBARFUNC 4
    End If
End With
End Sub

Private Sub lstPrice_ItemClick(ByVal Item As MSComctlLib.ListItem)
iRow = lstPrice.SelectedItem.Index
End Sub

Private Sub lstPrice_LostFocus()
isFocusPri = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "MenuManagementCode", "MenuMngtCode", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "MenuManagementCode", "MenuMngtCode", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "MenuManagementCode", "MenuMngtCode", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "MenuManagementCode", "MenuMngtCode", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":   'PRESS_F9
    Case "Refresh": PRESS_F11
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtAdjustment_Change()
If DET_TRANS = is_DET_ADDING Or _
DET_TRANS = is_DET_EDITTING Then
    txtSRP.Text = Format(RETURNTEXTVALUE(txtSRPTmp) + RETURNTEXTVALUE(txtAdjustment), "#,##0.00")
    With lstPrice.ListItems
        .Item(iRow).SubItems(5) = txtAdjustment.Text
    End With
End If
End Sub

Private Sub txtAdjustment_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtVAT.SetFocus
End Sub

Private Sub txtAdjustment_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtCost_Change()
If DET_TRANS = is_DET_ADDING Or _
DET_TRANS = is_DET_EDITTING Then

    dblMarkUp = RETURNTEXTVALUE(txtMarkUp)
    dblVat = RETURNTEXTVALUE(txtVAT)
    dblLocalTax = RETURNTEXTVALUE(txtLocalTax)
    dblService = RETURNTEXTVALUE(txtSrvcCharge)
    
    dblMarkUpSum = Format(RETURNTEXTVALUE(txtCost) * IIf(CDbl(dblMarkUp) <> 0, CDbl(dblMarkUp) / 100, 1), "#,##0.00")
    dblVatSum = Format(CDbl(dblMarkUpSum) * IIf(CDbl(dblVat) > 0, CDbl(dblVat) / 100, 0))
    dblLocalTaxSum = Format(CDbl(dblMarkUpSum) * IIf(CDbl(dblLocalTax) > 0, CDbl(dblLocalTax) / 100, 0))
    dblServiceSum = Format(CDbl(dblMarkUpSum) * IIf(CDbl(dblService) > 0, CDbl(dblService) / 100, 0))
    
    dblSRP = Format(CDbl(dblMarkUpSum) + CDbl(dblVatSum) + CDbl(dblLocalTaxSum) + CDbl(dblServiceSum), "#,##0.00")
    'dblSRP = Format(RETURNTEXTVALUE(txtCost) + CDbl(dblMarkUpSum) + CDbl(dblVatSum) + CDbl(dblLocalTaxSum), "#,##0.00")
    txtSRPTmp.Text = dblSRP
    
    dEmployeePer = 0
    If IsDate(txtEffectDate.Text) = False Then
        dEffectDate = FormatDateTime(Date, vbShortDate)
    Else
        dEffectDate = FormatDateTime(txtEffectDate.Text, vbShortDate)
    End If
    
    s = "SELECT tbl_Menu_PricingRate.* " & _
        " FROM tbl_Menu_PricingRate " & _
        " WHERE (EffectDate <= '" & FormatDateTime(dEffectDate, vbShortDate) & "')" & _
        " ORDER BY EffectDate DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        dEmployeePer = rs!Employee
    End If
    rs.Close
            
    dblMarkUpSum = RETURNTEXTVALUE(txtCost) + CDbl(RETURNTEXTVALUE(txtCost) * CDbl(Format((100 - dEmployeePer) / 100, "#,##0.00")))
    dblVatSum = CDbl(dblMarkUpSum) * IIf(RETURNTEXTVALUE(txtVAT) <> 0, RETURNTEXTVALUE(txtVAT) / 100, 0)
    dblLocalTaxSum = Format(CDbl(dblMarkUpSum) * IIf(CDbl(dblLocalTax) > 0, CDbl(dblLocalTax) / 100, 0))
    
    txtEmpRate.Text = Format(CDbl(dblMarkUpSum) + CDbl(dblVatSum), "#,##0.00")
    
    With lstPrice.ListItems
        .Item(iRow).SubItems(3) = txtCost.Text
    End With
End If
End Sub

Private Sub txtCost_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtDetCost_Change()
If DET_TRANS = is_DET_ADDING Or _
DET_TRANS = is_DET_EDITTING Then
    With lstMaterials.ListItems
        .Item(iRow).SubItems(6) = txtDetCost.Text
    End With
End If
End Sub

Private Sub txtDetCost_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtEffectDate_Change()
If DET_TRANS = is_DET_ADDING Or _
DET_TRANS = is_DET_EDITTING Then
    With lstPrice.ListItems
        If IsDate(txtEffectDate.Text) = True Then
            .Item(iRow).SubItems(2) = Format(FormatDateTime(txtEffectDate.Text, vbShortDate), "mm/dd/yyyy")
        End If
    End With
End If
End Sub

Private Sub txtEffectDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtMarkUp.SetFocus
End Sub

Private Sub txtEffectDate_LostFocus()
If IsDate(txtEffectDate.Text) = True Then
    txtEffectDate.Text = Format(FormatDateTime(txtEffectDate.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub

Private Sub txtEmpRate_Change()
If DET_TRANS = is_DET_ADDING Or _
DET_TRANS = is_DET_EDITTING Then
    With lstPrice.ListItems
        .Item(iRow).SubItems(11) = txtEmpRate.Text
    End With
End If
End Sub

Private Sub txtEmpRate_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtItemCode_Change()
If DET_TRANS = is_DET_ADDING Or _
DET_TRANS = is_DET_EDITTING Then
    With lstMaterials.ListItems
        .Item(iRow).SubItems(3) = txtItemCode.Text
    End With
End If
End Sub

Private Sub txtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtQty.SetFocus
End Sub

Private Sub txtItemCode_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtItemCode_LostFocus()
If Trim(txtItemCode.Text) = "" Then Exit Sub
s = "SELECT tbl_Inv_Items.* " & _
    " FROM tbl_Inv_Items " & _
    " WHERE (ItemCode = '" & txtItemCode.Text & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtItemKey.Text = rs!PK
    txtItemCode.Text = rs!ItemCode
    txtItemDesc.Text = rs!ItemDesc
    txtUnit.Text = rs!Unit
    txtDetCost.Text = Format(CDbl(rs!Cost) / CDbl(rs!ConUnit), "#,##0.00")
Else
    MsgBox "No Record Found!                ", vbCritical, "Error..."
    rs.Close
    Exit Sub
End If
rs.Close
End Sub

Private Sub txtItemDesc_Change()
If DET_TRANS = is_DET_ADDING Or _
DET_TRANS = is_DET_EDITTING Then
    With lstMaterials.ListItems
        .Item(iRow).SubItems(4) = txtItemDesc.Text
    End With
End If
End Sub

Private Sub txtItemKey_Change()
If DET_TRANS = is_DET_ADDING Or _
DET_TRANS = is_DET_EDITTING Then
    With lstMaterials.ListItems
        .Item(iRow).SubItems(2) = txtItemKey.Text
    End With
End If
End Sub

Private Sub txtLocalTax_Change()
If DET_TRANS = is_DET_ADDING Or _
DET_TRANS = is_DET_EDITTING Then
    dblMarkUp = RETURNTEXTVALUE(txtMarkUp)
    dblVat = RETURNTEXTVALUE(txtVAT)
    dblLocalTax = RETURNTEXTVALUE(txtLocalTax)
    dblService = RETURNTEXTVALUE(txtSrvcCharge)
    
    dblMarkUpSum = Format(RETURNTEXTVALUE(txtCost) * IIf(CDbl(dblMarkUp) <> 0, CDbl(dblMarkUp) / 100, 1), "#,##0.00")
    dblVatSum = Format(CDbl(dblMarkUpSum) * IIf(CDbl(dblVat) > 0, CDbl(dblVat) / 100, 0))
    dblLocalTaxSum = Format(CDbl(dblMarkUpSum) * IIf(CDbl(dblLocalTax) > 0, CDbl(dblLocalTax) / 100, 0))
    dblServiceSum = Format(CDbl(dblMarkUpSum) * IIf(CDbl(dblService) > 0, CDbl(dblService) / 100, 0))
    
    dblSRP = Format(CDbl(dblMarkUpSum) + CDbl(dblVatSum) + CDbl(dblLocalTaxSum) + CDbl(dblServiceSum), "#,##0.00")
    'dblSRP = Format(CDbl(dblMarkUpSum) + CDbl(dblVatSum) + CDbl(dblLocalTaxSum), "#,##0.00")
    'dblSRP = Format(RETURNTEXTVALUE(txtCost) + CDbl(dblMarkUpSum) + CDbl(dblVatSum) + CDbl(dblLocalTaxSum), "#,##0.00")
    txtSRPTmp.Text = dblSRP
    
    dEmployeePer = 0
    If IsDate(txtEffectDate.Text) = False Then
        dEffectDate = FormatDateTime(Date, vbShortDate)
    Else
        dEffectDate = FormatDateTime(txtEffectDate.Text, vbShortDate)
    End If
    
    s = "SELECT tbl_Menu_PricingRate.* " & _
        " FROM tbl_Menu_PricingRate " & _
        " WHERE (EffectDate <= '" & FormatDateTime(dEffectDate, vbShortDate) & "')" & _
        " ORDER BY EffectDate DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        dEmployeePer = rs!Employee
    End If
    rs.Close
            
    dblMarkUpSum = RETURNTEXTVALUE(txtCost) + CDbl(RETURNTEXTVALUE(txtCost) * CDbl(Format((100 - dEmployeePer) / 100, "#,##0.00")))
    'dblMarkUpSum = RETURNTEXTVALUE(txtCost) + CDbl(RETURNTEXTVALUE(txtCost) * 0.5)
    dblVatSum = CDbl(dblMarkUpSum) * IIf(RETURNTEXTVALUE(txtVAT) <> 0, RETURNTEXTVALUE(txtVAT) / 100, 0)
    dblLocalTaxSum = Format(CDbl(dblMarkUpSum) * IIf(CDbl(dblLocalTax) > 0, CDbl(dblLocalTax) / 100, 0))
    
    txtEmpRate.Text = Format(CDbl(dblMarkUpSum) + CDbl(dblVatSum), "#,##0.00")
    With lstPrice.ListItems
        .Item(iRow).SubItems(7) = txtLocalTax.Text
    End With
End If
End Sub

Private Sub txtLocalTax_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtSrvcCharge.SetFocus
End Sub

Private Sub txtLocalTax_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtMarkUp_Change()
If DET_TRANS = is_DET_ADDING Or _
DET_TRANS = is_DET_EDITTING Then

    dblMarkUp = RETURNTEXTVALUE(txtMarkUp)
    dblVat = RETURNTEXTVALUE(txtVAT)
    dblLocalTax = RETURNTEXTVALUE(txtLocalTax)
    dblService = RETURNTEXTVALUE(txtSrvcCharge)
    
    dblMarkUpSum = Format(RETURNTEXTVALUE(txtCost) * IIf(CDbl(dblMarkUp) <> 0, CDbl(dblMarkUp) / 100, 1), "#,##0.00")
    dblVatSum = Format(CDbl(dblMarkUpSum) * IIf(CDbl(dblVat) > 0, CDbl(dblVat) / 100, 0))
    dblLocalTaxSum = Format(CDbl(dblMarkUpSum) * IIf(CDbl(dblLocalTax) > 0, CDbl(dblLocalTax) / 100, 0))
    dblServiceSum = Format(CDbl(dblMarkUpSum) * IIf(CDbl(dblService) > 0, CDbl(dblService) / 100, 0))
    
    dblSRP = Format(CDbl(dblMarkUpSum) + CDbl(dblVatSum) + CDbl(dblLocalTaxSum) + CDbl(dblServiceSum), "#,##0.00")
    'dblSRP = Format(RETURNTEXTVALUE(txtCost) + CDbl(dblMarkUpSum) + CDbl(dblVatSum) + CDbl(dblLocalTaxSum), "#,##0.00")
    'dblSRP = Format(CDbl(dblMarkUpSum) + CDbl(dblVatSum) + CDbl(dblLocalTaxSum), "#,##0.00")
    txtSRPTmp.Text = dblSRP
    
    dEmployeePer = 0
    If IsDate(txtEffectDate.Text) = False Then
        dEffectDate = FormatDateTime(Date, vbShortDate)
    Else
        dEffectDate = FormatDateTime(txtEffectDate.Text, vbShortDate)
    End If
    
    s = "SELECT tbl_Menu_PricingRate.* " & _
        " FROM tbl_Menu_PricingRate " & _
        " WHERE (EffectDate <= '" & FormatDateTime(dEffectDate, vbShortDate) & "')" & _
        " ORDER BY EffectDate DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        dEmployeePer = rs!Employee
    End If
    rs.Close
            
    dblMarkUpSum = RETURNTEXTVALUE(txtCost) + CDbl(RETURNTEXTVALUE(txtCost) * CDbl(Format((100 - dEmployeePer) / 100, "#,##0.00")))
    
    'dblMarkUpSum = RETURNTEXTVALUE(txtCost) + CDbl(RETURNTEXTVALUE(txtCost) * 0.5)
    dblVatSum = CDbl(dblMarkUpSum) * IIf(RETURNTEXTVALUE(txtVAT) <> 0, RETURNTEXTVALUE(txtVAT) / 100, 0)
    dblLocalTaxSum = Format(CDbl(dblMarkUpSum) * IIf(CDbl(dblLocalTax) > 0, CDbl(dblLocalTax) / 100, 0))
    
    txtEmpRate.Text = Format(CDbl(dblMarkUpSum) + CDbl(dblVatSum), "#,##0.00")

    With lstPrice.ListItems
        .Item(iRow).SubItems(4) = txtMarkUp.Text
    End With
End If
End Sub

Private Sub txtMarkUp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtAdjustment.SetFocus
End Sub

Private Sub txtMarkUp_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtMemberRate_Change()
If DET_TRANS = is_DET_ADDING Or _
DET_TRANS = is_DET_EDITTING Then
    With lstPrice.ListItems
        .Item(iRow).SubItems(10) = txtMemberRate.Text
    End With
End If
End Sub

Private Sub txtMemberRate_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtQty_Change()
If DET_TRANS = is_DET_ADDING Or _
DET_TRANS = is_DET_EDITTING Then
    With lstMaterials.ListItems
        .Item(iRow).SubItems(7) = txtQty.Text
    End With
    txtTotalDetCost.Text = Format(RETURNTEXTVALUE(txtQty) * RETURNTEXTVALUE(txtDetCost), "#,##0.00")
End If
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    picMSLine.Visible = False
    picBody.Enabled = True
    lstMaterials.SetFocus
End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtSRP_Change()
If DET_TRANS = is_DET_ADDING Or _
DET_TRANS = is_DET_EDITTING Then

    dMemberPer = 0
    If IsDate(txtEffectDate.Text) = False Then
        dEffectDate = FormatDateTime(Date, vbShortDate)
    Else
        dEffectDate = FormatDateTime(txtEffectDate.Text, vbShortDate)
    End If
    
    t = "SELECT tbl_Menu_PricingRate.* " & _
        " FROM tbl_Menu_PricingRate " & _
        " WHERE (EffectDate <= '" & FormatDateTime(dEffectDate, vbShortDate) & "')" & _
        " ORDER BY EffectDate DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        dMemberPer = rt!Member
    End If
    rt.Close
    
    txtMemberRate.Text = Format(RETURNTEXTVALUE(txtSRP) * CDbl((100 - dMemberPer) / 100), "#,##0.00")
    'txtMemberRate.Text = Format(RETURNTEXTVALUE(txtSRP) * 0.9, "#,##0.00")
    
    With lstPrice.ListItems
        .Item(iRow).SubItems(9) = txtSRP.Text
    End With
    
End If
End Sub

Private Sub txtSRP_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtSRPTmp_Change()
txtSRP.Text = Format(RETURNTEXTVALUE(txtSRPTmp) + RETURNTEXTVALUE(txtAdjustment), "#,##0.00")
End Sub

Private Sub txtSrvcCharge_Change()
If DET_TRANS = is_DET_ADDING Or _
DET_TRANS = is_DET_EDITTING Then

    dblMarkUp = RETURNTEXTVALUE(txtMarkUp)
    dblVat = RETURNTEXTVALUE(txtVAT)
    dblLocalTax = RETURNTEXTVALUE(txtLocalTax)
    dblService = RETURNTEXTVALUE(txtSrvcCharge)
    
    dblMarkUpSum = Format(RETURNTEXTVALUE(txtCost) * IIf(CDbl(dblMarkUp) <> 0, CDbl(dblMarkUp) / 100, 1), "#,##0.00")
    dblVatSum = Format(CDbl(dblMarkUpSum) * IIf(CDbl(dblVat) > 0, CDbl(dblVat) / 100, 0))
    dblLocalTaxSum = Format(CDbl(dblMarkUpSum) * IIf(CDbl(dblLocalTax) > 0, CDbl(dblLocalTax) / 100, 0))
    dblServiceSum = Format(CDbl(dblMarkUpSum) * IIf(CDbl(dblService) > 0, CDbl(dblService) / 100, 0))
    
    dblSRP = Format(CDbl(dblMarkUpSum) + CDbl(dblVatSum) + CDbl(dblLocalTaxSum) + CDbl(dblServiceSum), "#,##0.00")
    'dblSRP = Format(CDbl(dblMarkUpSum) + CDbl(dblVatSum) + CDbl(dblLocalTaxSum), "#,##0.00")
    'dblSRP = Format(RETURNTEXTVALUE(txtCost) + CDbl(dblMarkUpSum) + CDbl(dblVatSum) + CDbl(dblLocalTaxSum), "#,##0.00")
    txtSRPTmp.Text = dblSRP
    
    dEmployeePer = 0
    If IsDate(txtEffectDate.Text) = False Then
        dEffectDate = FormatDateTime(Date, vbShortDate)
    Else
        dEffectDate = FormatDateTime(txtEffectDate.Text, vbShortDate)
    End If
    
    s = "SELECT tbl_Menu_PricingRate.* " & _
        " FROM tbl_Menu_PricingRate " & _
        " WHERE (EffectDate <= '" & FormatDateTime(dEffectDate, vbShortDate) & "')" & _
        " ORDER BY EffectDate DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        dEmployeePer = rs!Employee
    End If
    rs.Close
            
    dblMarkUpSum = RETURNTEXTVALUE(txtCost) + CDbl(RETURNTEXTVALUE(txtCost) * CDbl(Format((100 - dEmployeePer) / 100, "#,##0.00")))
    'dblMarkUpSum = RETURNTEXTVALUE(txtCost) + CDbl(RETURNTEXTVALUE(txtCost) * 0.5)
    dblVatSum = CDbl(dblMarkUpSum) * IIf(RETURNTEXTVALUE(txtVAT) <> 0, RETURNTEXTVALUE(txtVAT) / 100, 0)
    dblLocalTaxSum = Format(CDbl(dblMarkUpSum) * IIf(CDbl(dblLocalTax) > 0, CDbl(dblLocalTax) / 100, 0))
    
    txtEmpRate.Text = Format(CDbl(dblMarkUpSum) + CDbl(dblVatSum), "#,##0.00")
    With lstPrice.ListItems
        .Item(iRow).SubItems(8) = txtSrvcCharge.Text
    End With
End If
End Sub

Private Sub txtSrvcCharge_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If IsDate(txtEffectDate.Text) = False Then MsgBox "Please Supply a Valid Date!                ", vbCritical, "Error...": txtEffectDate.SetFocus: Exit Sub
    picBody.Enabled = True
    picPSline.Visible = False
    lstPrice.SetFocus
End If
End Sub

Private Sub txtSrvcCharge_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtTotalDetCost_Change()
If DET_TRANS = is_DET_ADDING Or _
DET_TRANS = is_DET_EDITTING Then
    With lstMaterials.ListItems
        .Item(iRow).SubItems(8) = txtTotalDetCost.Text
        j = 0
        For i = 1 To .Count
            j = j + CDbl(IIf(IsNumeric(.Item(i).SubItems(8)) = False, 0, .Item(i).SubItems(8)))
        Next i
        lblTotalCost.Caption = Format(j, "#,##0.00")
    End With
End If
End Sub

Private Sub txtUnit_Change()
If DET_TRANS = is_DET_ADDING Or _
DET_TRANS = is_DET_EDITTING Then
    With lstMaterials.ListItems
        .Item(iRow).SubItems(5) = txtUnit.Text
    End With
End If
End Sub

Private Sub txtVAT_Change()
If DET_TRANS = is_DET_ADDING Or _
DET_TRANS = is_DET_EDITTING Then
    dblMarkUp = RETURNTEXTVALUE(txtMarkUp)
    dblVat = RETURNTEXTVALUE(txtVAT)
    dblLocalTax = RETURNTEXTVALUE(txtLocalTax)
    dblService = RETURNTEXTVALUE(txtSrvcCharge)
    
    dblMarkUpSum = Format(RETURNTEXTVALUE(txtCost) * IIf(CDbl(dblMarkUp) <> 0, CDbl(dblMarkUp) / 100, 1), "#,##0.00")
    dblVatSum = Format(CDbl(dblMarkUpSum) * IIf(CDbl(dblVat) > 0, CDbl(dblVat) / 100, 0))
    dblLocalTaxSum = Format(CDbl(dblMarkUpSum) * IIf(CDbl(dblLocalTax) > 0, CDbl(dblLocalTax) / 100, 0))
    dblServiceSum = Format(CDbl(dblMarkUpSum) * IIf(CDbl(dblService) > 0, CDbl(dblService) / 100, 0))
    
    dblSRP = Format(CDbl(dblMarkUpSum) + CDbl(dblVatSum) + CDbl(dblLocalTaxSum) + CDbl(dblServiceSum), "#,##0.00")
    'dblSRP = Format(CDbl(dblMarkUpSum) + CDbl(dblVatSum) + CDbl(dblLocalTaxSum), "#,##0.00")
    'dblSRP = Format(RETURNTEXTVALUE(txtCost) + CDbl(dblMarkUpSum) + CDbl(dblVatSum) + CDbl(dblLocalTaxSum), "#,##0.00")
    txtSRPTmp.Text = dblSRP
    
    dEmployeePer = 0
    If IsDate(txtEffectDate.Text) = False Then
        dEffectDate = FormatDateTime(Date, vbShortDate)
    Else
        dEffectDate = FormatDateTime(txtEffectDate.Text, vbShortDate)
    End If
    
    s = "SELECT tbl_Menu_PricingRate.* " & _
        " FROM tbl_Menu_PricingRate " & _
        " WHERE (EffectDate <= '" & FormatDateTime(dEffectDate, vbShortDate) & "')" & _
        " ORDER BY EffectDate DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        dEmployeePer = rs!Employee
    End If
    rs.Close
            
    dblMarkUpSum = RETURNTEXTVALUE(txtCost) + CDbl(RETURNTEXTVALUE(txtCost) * CDbl(Format((100 - dEmployeePer) / 100, "#,##0.00")))
    'dblMarkUpSum = RETURNTEXTVALUE(txtCost) + CDbl(RETURNTEXTVALUE(txtCost) * 0.5)
    dblVatSum = CDbl(dblMarkUpSum) * IIf(RETURNTEXTVALUE(txtVAT) <> 0, RETURNTEXTVALUE(txtVAT) / 100, 0)
    dblLocalTaxSum = Format(CDbl(dblMarkUpSum) * IIf(CDbl(dblLocalTax) > 0, CDbl(dblLocalTax) / 100, 0))
    
    txtEmpRate.Text = Format(CDbl(dblMarkUpSum) + CDbl(dblVatSum), "#,##0.00")
    With lstPrice.ListItems
        .Item(iRow).SubItems(6) = txtVAT.Text
    End With
End If
End Sub

Private Sub txtVAT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtLocalTax.SetFocus
End Sub

Private Sub txtVAT_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub


