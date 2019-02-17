VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInvRR 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvRR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   14970
   ShowInTaskbar   =   0   'False
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
            Picture         =   "frmInvRR.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvRR.frx":15A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvRR.frx":227E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvRR.frx":2F58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvRR.frx":3C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvRR.frx":490C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvRR.frx":55E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvRR.frx":62C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvRR.frx":6F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvRR.frx":7874
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvRR.frx":854E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvRR.frx":9228
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvRR.frx":9F02
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvRR.frx":ABDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvRR.frx":B8B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvRR.frx":C590
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15600
      TabIndex        =   145
      Top             =   0
      Width           =   15600
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   146
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
         MouseIcon       =   "frmInvRR.frx":D26A
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   13380
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   147
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmInvRR.frx":D584
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
   Begin RPVGCC.b8Container picSLine 
      Height          =   855
      Left            =   240
      TabIndex        =   45
      Top             =   5040
      Visible         =   0   'False
      Width           =   14625
      _ExtentX        =   25797
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox txtSLRemarks1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   117
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtSLRemarks 
         Height          =   315
         Left            =   12480
         TabIndex        =   115
         Top             =   360
         Width           =   1995
      End
      Begin VB.TextBox txtTotalCost 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   12000
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtNetCost1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtCost1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtRecd1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtNetCost 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   360
         Width           =   1035
      End
      Begin VB.TextBox txtTotalNetCost 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   11160
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   360
         Width           =   1275
      End
      Begin VB.TextBox txtCost 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9000
         TabIndex        =   58
         Top             =   360
         Width           =   1035
      End
      Begin VB.TextBox txtDescription 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   5880
         TabIndex        =   56
         Top             =   360
         Width           =   3075
      End
      Begin VB.TextBox txtItemCode 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   4680
         TabIndex        =   54
         Top             =   360
         Width           =   1155
      End
      Begin VB.TextBox txtUnit 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   3480
         TabIndex        =   52
         Top             =   360
         Width           =   1155
      End
      Begin VB.TextBox txtRecd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2400
         TabIndex        =   50
         Top             =   360
         Width           =   1035
      End
      Begin VB.TextBox txtOrdd 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1320
         TabIndex        =   48
         Top             =   360
         Width           =   1035
      End
      Begin VB.TextBox txtType 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "REMARKS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   12600
         TabIndex        =   116
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NETCOST"
         Height          =   255
         Left            =   10080
         TabIndex        =   63
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL NETCOST"
         Height          =   255
         Left            =   11160
         TabIndex        =   62
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "COST"
         Height          =   255
         Left            =   9000
         TabIndex        =   61
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM DESCRIPTION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5880
         TabIndex        =   57
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM CODE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4680
         TabIndex        =   55
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "UNIT"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3480
         TabIndex        =   53
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "REC'D"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         TabIndex        =   51
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ORD'D"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1320
         TabIndex        =   49
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "TYPE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   120
         Width           =   1095
      End
   End
   Begin RPVGCC.b8Container picAddRR 
      Height          =   2535
      Left            =   5520
      TabIndex        =   129
      Top             =   2040
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4471
      BackColor       =   15396057
      Begin VB.TextBox txtRRDateAdd 
         Height          =   315
         Left            =   1560
         TabIndex        =   137
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtRRNumAdd 
         Height          =   315
         Left            =   1560
         TabIndex        =   135
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancelAddRR 
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
         Picture         =   "frmInvRR.frx":DC97
         Style           =   1  'Graphical
         TabIndex        =   132
         Top             =   1800
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKAddRR 
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
         Picture         =   "frmInvRR.frx":E3F3
         Style           =   1  'Graphical
         TabIndex        =   131
         Top             =   1800
         Width           =   1560
      End
      Begin VB.TextBox txtPONumAdd 
         Height          =   315
         Left            =   1560
         TabIndex        =   130
         Top             =   600
         Width           =   1575
      End
      Begin RPVGCC.b8TitleBar b8TitleBar5 
         Height          =   345
         Left            =   40
         TabIndex        =   133
         Top             =   40
         Width           =   3880
         _ExtentX        =   6853
         _ExtentY        =   609
         Caption         =   "Supply PO Number"
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
         Icon            =   "frmInvRR.frx":EA65
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "RR DATE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   138
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "RR  NUMBER"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   136
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "PO  NUMBER"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   134
         Top             =   600
         Width           =   975
      End
   End
   Begin RPVGCC.b8Container picPost 
      Height          =   1815
      Left            =   5520
      TabIndex        =   68
      Top             =   2400
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3201
      BackColor       =   15396057
      Begin VB.TextBox txtRRNoPosting 
         Height          =   315
         Left            =   1200
         TabIndex        =   118
         Top             =   4200
         Width           =   2535
      End
      Begin VB.TextBox txtRRDatePosting 
         Height          =   315
         Left            =   1200
         TabIndex        =   75
         Top             =   3840
         Width           =   2535
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
         Picture         =   "frmInvRR.frx":EFFF
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   1080
         Width           =   1560
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
         Picture         =   "frmInvRR.frx":F671
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   1080
         Width           =   1560
      End
      Begin VB.ComboBox cmbLocation 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   600
         Width           =   2535
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   40
         TabIndex        =   69
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
         Icon            =   "frmInvRR.frx":FDCD
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "RR NUMBER"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   119
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "RR DATE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "LOCATION"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   71
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.PictureBox picAccDistribution 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   3720
      ScaleHeight     =   3855
      ScaleWidth      =   7335
      TabIndex        =   77
      Top             =   1560
      Visible         =   0   'False
      Width           =   7335
      Begin RPVGCC.b8Container b8Container1 
         Height          =   3615
         Left            =   0
         TabIndex        =   78
         Top             =   0
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6376
         BackColor       =   15396057
         Begin VB.ComboBox cmbBookType 
            Height          =   315
            Left            =   2280
            TabIndex        =   120
            Text            =   "Combo1"
            Top             =   3080
            Width           =   1215
         End
         Begin VB.TextBox txtInvGrossP 
            Height          =   315
            Left            =   3240
            TabIndex        =   110
            Top             =   720
            Width           =   1995
         End
         Begin VB.TextBox txtInvDateP 
            Height          =   315
            Left            =   1680
            TabIndex        =   109
            Top             =   720
            Width           =   1515
         End
         Begin VB.TextBox txtInvNumberP 
            Height          =   315
            Left            =   120
            TabIndex        =   108
            Top             =   720
            Width           =   1515
         End
         Begin VB.TextBox txtInvNetP 
            Height          =   315
            Left            =   5280
            TabIndex        =   107
            Top             =   720
            Width           =   1875
         End
         Begin VB.CommandButton cmdPost 
            Caption         =   "P O S T"
            Height          =   495
            Left            =   120
            TabIndex        =   106
            Top             =   3000
            Width           =   1095
         End
         Begin MSComctlLib.ListView lstAccDistribution 
            Height          =   1815
            Left            =   120
            TabIndex        =   80
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
         Begin RPVGCC.b8TitleBar b8TitleBar2 
            Height          =   345
            Left            =   45
            TabIndex        =   79
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
            Icon            =   "frmInvRR.frx":10367
         End
         Begin VB.Label Label45 
            BackStyle       =   0  'Transparent
            Caption         =   "BOOK TYPE"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1320
            TabIndex        =   121
            Top             =   3120
            Width           =   975
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "INVOICE GROSS AMOUNT"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3240
            TabIndex        =   114
            Top             =   480
            Width           =   1995
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "INVOICE DATE"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1680
            TabIndex        =   113
            Top             =   480
            Width           =   1515
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "INVOICE NUMBER"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "INVOICE NET AMOUNT"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5280
            TabIndex        =   111
            Top             =   480
            Width           =   1875
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "BALANCE >>"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3600
            TabIndex        =   85
            Top             =   3240
            Width           =   975
         End
         Begin VB.Label lblBalance 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5760
            TabIndex        =   84
            Top             =   3240
            Width           =   1095
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL >>"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3600
            TabIndex        =   83
            Top             =   3000
            Width           =   855
         End
         Begin VB.Label lblTotalCredit 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5760
            TabIndex        =   82
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label lblTotalDebit 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4440
            TabIndex        =   81
            Top             =   3000
            Width           =   1215
         End
      End
   End
   Begin RPVGCC.b8Container picGLAddAutoVAT 
      Height          =   3315
      Left            =   4800
      TabIndex        =   122
      Top             =   2160
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5847
      BackColor       =   15396057
      Begin VB.CheckBox chkWithVAT 
         BackColor       =   &H00EAECD9&
         Caption         =   "With VAT"
         Height          =   195
         Left            =   2040
         TabIndex        =   128
         Top             =   2400
         Width           =   1335
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
         Picture         =   "frmInvRR.frx":10901
         Style           =   1  'Graphical
         TabIndex        =   126
         Top             =   2715
         Width           =   1560
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
         Picture         =   "frmInvRR.frx":10F73
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   2715
         Width           =   1560
      End
      Begin VB.TextBox txtGLAddAutoVAT 
         Height          =   315
         Left            =   120
         TabIndex        =   124
         Top             =   480
         Width           =   5295
      End
      Begin VB.ListBox lstGLAutoVAT 
         Height          =   1425
         Left            =   120
         TabIndex        =   123
         Top             =   840
         Width           =   5295
      End
      Begin RPVGCC.b8TitleBar b8TitleBar4 
         Height          =   345
         Left            =   45
         TabIndex        =   127
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
         Icon            =   "frmInvRR.frx":116CF
         ShadowVisible   =   0   'False
      End
   End
   Begin VB.PictureBox picADSLine 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   3720
      ScaleHeight     =   855
      ScaleWidth      =   7695
      TabIndex        =   92
      Top             =   1800
      Visible         =   0   'False
      Width           =   7695
      Begin RPVGCC.b8Container picADSLine1 
         Height          =   855
         Left            =   0
         TabIndex        =   93
         Top             =   0
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   1508
         BackColor       =   8438015
         Begin VB.TextBox txtAccountNo 
            Height          =   315
            Left            =   120
            TabIndex        =   101
            Top             =   360
            Width           =   1155
         End
         Begin VB.TextBox txtAccountName 
            Height          =   315
            Left            =   1320
            TabIndex        =   100
            Top             =   360
            Width           =   3195
         End
         Begin VB.TextBox txtDebit 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4560
            TabIndex        =   99
            Top             =   360
            Width           =   1275
         End
         Begin VB.TextBox txtCredit 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5880
            TabIndex        =   98
            Top             =   360
            Width           =   1275
         End
         Begin VB.TextBox txtAccountNo1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   97
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtAccountName1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   96
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtDebit1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   95
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
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
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "ACCOUNT #"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "ACCOUNT NAME"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1320
            TabIndex        =   104
            Top             =   120
            Width           =   3135
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DEBIT"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4560
            TabIndex        =   103
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "CREDIT"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5880
            TabIndex        =   102
            Top             =   120
            Width           =   1215
         End
      End
   End
   Begin RPVGCC.b8Container picSearchGLAccount 
      Height          =   2955
      Left            =   5040
      TabIndex        =   86
      Top             =   2520
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5212
      BackColor       =   15396057
      Begin VB.ListBox lstResultGLAccount 
         Height          =   1425
         Left            =   120
         TabIndex        =   90
         Top             =   840
         Width           =   5295
      End
      Begin VB.TextBox txtSearchGLAccount 
         Height          =   315
         Left            =   120
         TabIndex        =   89
         Top             =   480
         Width           =   5295
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
         Picture         =   "frmInvRR.frx":11C69
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   2355
         Width           =   1560
      End
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
         Picture         =   "frmInvRR.frx":123C5
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   2355
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar3 
         Height          =   345
         Left            =   45
         TabIndex        =   91
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
         Icon            =   "frmInvRR.frx":12A37
         ShadowVisible   =   0   'False
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   34
      Top             =   6825
      Width           =   14970
      _ExtentX        =   26405
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
   Begin VB.PictureBox picBody 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   240
      ScaleHeight     =   5415
      ScaleWidth      =   14535
      TabIndex        =   0
      Top             =   1200
      Width           =   14535
      Begin VB.TextBox txtDeptHead 
         Height          =   315
         Left            =   3600
         TabIndex        =   143
         Top             =   5040
         Width           =   1995
      End
      Begin VB.TextBox txtPurchaser 
         Height          =   315
         Left            =   1800
         TabIndex        =   141
         Top             =   5040
         Width           =   1755
      End
      Begin VB.TextBox txtStockClerk 
         Height          =   315
         Left            =   0
         TabIndex        =   139
         Top             =   5040
         Width           =   1755
      End
      Begin VB.TextBox txtRRTime 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   3720
         TabIndex        =   43
         Top             =   660
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.TextBox txtSupplier 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   6960
         TabIndex        =   42
         Top             =   0
         Width           =   5655
      End
      Begin VB.TextBox txtDept 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   6960
         TabIndex        =   41
         Top             =   990
         Width           =   5655
      End
      Begin VB.TextBox txtRRDate 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   3720
         TabIndex        =   38
         Top             =   330
         Width           =   1155
      End
      Begin VB.TextBox txtRRNumber 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   3720
         TabIndex        =   37
         Text            =   "2011"
         Top             =   0
         Width           =   1155
      End
      Begin VB.TextBox txtInvNet 
         Height          =   315
         Left            =   5640
         TabIndex        =   35
         Top             =   4440
         Width           =   1875
      End
      Begin VB.TextBox txtVat 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   10320
         TabIndex        =   13
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtDisc 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   6960
         TabIndex        =   12
         Top             =   1320
         Width           =   2130
      End
      Begin VB.TextBox txtRemarks 
         Height          =   315
         Left            =   840
         TabIndex        =   11
         Top             =   3840
         Width           =   6680
      End
      Begin VB.TextBox txtPONumber 
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Top             =   0
         Width           =   1155
      End
      Begin VB.TextBox txtPODate 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   330
         Width           =   1155
      End
      Begin VB.TextBox txtRefNo 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   660
         Width           =   1155
      End
      Begin VB.TextBox txtTerms 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   990
         Width           =   3435
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   6960
         TabIndex        =   6
         Top             =   330
         Width           =   5655
      End
      Begin VB.TextBox txtTelNo 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   6960
         TabIndex        =   5
         Top             =   660
         Width           =   2115
      End
      Begin VB.TextBox txtFaxNo 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   10320
         TabIndex        =   4
         Top             =   660
         Width           =   2295
      End
      Begin VB.TextBox txtInvNumber 
         Height          =   315
         Left            =   0
         TabIndex        =   3
         Top             =   4440
         Width           =   1755
      End
      Begin VB.TextBox txtInvDate 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   4440
         Width           =   1755
      End
      Begin VB.TextBox txtInvGross 
         Height          =   315
         Left            =   3600
         TabIndex        =   1
         Top             =   4440
         Width           =   1995
      End
      Begin MSComctlLib.ListView lstDetail 
         Height          =   2025
         Left            =   0
         TabIndex        =   14
         Top             =   1680
         Width           =   14535
         _ExtentX        =   25638
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
         NumItems        =   17
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
            Text            =   "Ord'd"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Rec'd"
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
            Object.Width           =   5997
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
            Text            =   "TotalCost"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "TypeKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "InvQty"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "InvCost"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "InvNetCost"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Remarks"
            Object.Width           =   3440
         EndProperty
      End
      Begin VB.Label Label51 
         BackStyle       =   0  'Transparent
         Caption         =   "DEPARTMENT HEAD"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   144
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "PURCHASER"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   142
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "STOCK CLERK"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   140
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label lblInvPosted 
         Alignment       =   1  'Right Justify
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
         Left            =   9480
         TabIndex        =   76
         Top             =   4440
         Width           =   2775
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "RR TIME"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   44
         Top             =   680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "RR DATE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   40
         Top             =   345
         Width           =   975
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "RR #"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   39
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "INVOICE NET AMOUNT"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5640
         TabIndex        =   36
         Top             =   4200
         Width           =   1875
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
         Left            =   10920
         TabIndex        =   33
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL NETCOST"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9720
         TabIndex        =   32
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "VAT"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9360
         TabIndex        =   31
         Top             =   1335
         Width           =   975
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "DISCOUNT"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6120
         TabIndex        =   30
         Top             =   1335
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "REMARKS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "DEPT"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6120
         TabIndex        =   28
         Top             =   1005
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6120
         TabIndex        =   27
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PO #"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PO DATE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   25
         Top             =   350
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "REF/PR #"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   24
         Top             =   680
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Terms of Payment"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   23
         Top             =   1000
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6120
         TabIndex        =   22
         Top             =   345
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "TEL NO"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6120
         TabIndex        =   21
         Top             =   675
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "FAX NO"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9360
         TabIndex        =   20
         Top             =   675
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "INVOICE NUMBER"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "INVOICE DATE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   4200
         Width           =   1755
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "INVOICE GROSS AMOUNT"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   17
         Top             =   4200
         Width           =   1995
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL COST"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9720
         TabIndex        =   16
         Top             =   3840
         Width           =   1095
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
         Left            =   10920
         TabIndex        =   15
         Top             =   3840
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmInvRR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TRANSACTIONTYPE  As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2
Const is_FINDING = 3

Public iSupplier


Dim TRANS_DETAIL As Long
Const is_DET_REFRESH = 0
Const is_DET_ADDING = 1
Const is_DET_EDITTING = 2

Dim is_DET_FOCUS    As Long
Dim ROW             As Long
Dim isGLCodeFocus   As Long
Dim isGLCodeChange  As Long
Dim tmp             As Long

Dim iGLDepartment   As Long
Dim dR_Inv_Qty      As Double
Dim dR_Inv_Cost     As Double
Dim dR_Inv_NetCost  As Double
Dim iRRNumberAuto   As Boolean

Dim i, a, b, x, sSupplierCode, sRRNumber, Arr, dAPTrade, dAmount, iBookType, _
dVATable, dNetVAT, dVAT, dTotalRcd, iDept, strDisc

Public Sub BROWSER(sPONum, isWant As String)
Select Case isWant
    Case "is_LOAD"
        If sPONum <> "" Then
            s = "SELECT TOP 1 tbl_Inv_RR.PK, tbl_Inv_RR.PONumber, tbl_Inv_RR.PODate, tbl_Inv_RR.PRNumber, " & _
                " tbl_Inv_RR.SuppKey, tbl_Inv_RR.DeptKey, tbl_Inv_RR.Terms, tbl_Inv_RR.Remarks, tbl_Inv_RR.TotalCost, " & _
                " tbl_Inv_RR.TotalDiscount, tbl_Inv_RR.TotalNetCost, tbl_Inv_RR.RequestBy, tbl_Inv_RR.CheckedBy, " & _
                " tbl_Inv_RR.ApprovedBy, tbl_Inv_RR.Posted, tbl_Inv_RR.Printed, tbl_Inv_RR.LastModified, " & _
                " tbl_Inv_RR.RecdPartFull, tbl_Inv_RR.Discount, tbl_Inv_RR.TotalQty, tbl_Inv_RR.TotalRcd, " & _
                " tbl_Inv_Supplier.SupplierCode, tbl_Inv_Supplier.SupplierName, tbl_Inv_Supplier.Address1, " & _
                " tbl_Inv_Supplier.Address2, tbl_Inv_Supplier.Address3, tbl_Inv_Supplier.TelNo, tbl_Inv_Supplier.FaxNo, " & _
                " tbl_Inv_Supplier.Email, tbl_Inv_Supplier.ContactPerson, tbl_GL_Department.Code as DepartmentCode, " & _
                " tbl_GL_Department.DeptName as DepartmentName, tbl_Inv_RR.RRNumber, tbl_Inv_RR.RRDateTime, " & _
                " tbl_Inv_RR.RRPosted, tbl_Inv_RR.InvNumber, tbl_Inv_RR.InvDate, tbl_Inv_RR.InvGrossAmt, " & _
                " tbl_Inv_RR.InvNetAmt, tbl_Inv_RR.RRPrinted, tbl_Inv_RR.LastModifiedRR, tbl_Inv_RR.TotalCostRecd, " & _
                " tbl_Inv_RR.TotalNetCostRecd, tbl_Inv_RR.RRInvPosted, tbl_Inv_RR.BookType, tbl_Inv_RR.StockClerk, " & _
                " tbl_Inv_RR.Purchaser, tbl_Inv_RR.DeptHead " & _
                " FROM tbl_Inv_RR LEFT OUTER JOIN " & _
                " tbl_Inv_Supplier ON tbl_Inv_RR.SuppKey = tbl_Inv_Supplier.PK LEFT OUTER JOIN " & _
                " tbl_GL_Department ON tbl_Inv_RR.DeptKey = tbl_GL_Department.PK " & _
                " WHERE (tbl_Inv_RR.RRNumber = '" & sPONum & "') " & _
                " ORDER BY tbl_Inv_RR.RRNumber"
        Else
            s = "SELECT TOP 1 tbl_Inv_RR.PK, tbl_Inv_RR.PONumber, tbl_Inv_RR.PODate, tbl_Inv_RR.PRNumber, " & _
                " tbl_Inv_RR.SuppKey, tbl_Inv_RR.DeptKey, tbl_Inv_RR.Terms, tbl_Inv_RR.Remarks, tbl_Inv_RR.TotalCost, " & _
                " tbl_Inv_RR.TotalDiscount, tbl_Inv_RR.TotalNetCost, tbl_Inv_RR.RequestBy, tbl_Inv_RR.CheckedBy, " & _
                " tbl_Inv_RR.ApprovedBy, tbl_Inv_RR.Posted, tbl_Inv_RR.Printed, tbl_Inv_RR.LastModified, " & _
                " tbl_Inv_RR.RecdPartFull, tbl_Inv_RR.Discount, tbl_Inv_RR.TotalQty, tbl_Inv_RR.TotalRcd, " & _
                " tbl_Inv_Supplier.SupplierCode, tbl_Inv_Supplier.SupplierName, tbl_Inv_Supplier.Address1, " & _
                " tbl_Inv_Supplier.Address2, tbl_Inv_Supplier.Address3, tbl_Inv_Supplier.TelNo, tbl_Inv_Supplier.FaxNo, " & _
                " tbl_Inv_Supplier.Email, tbl_Inv_Supplier.ContactPerson, tbl_GL_Department.Code as DepartmentCode, " & _
                " tbl_GL_Department.DeptName as DepartmentName, tbl_Inv_RR.RRNumber, tbl_Inv_RR.RRDateTime, " & _
                " tbl_Inv_RR.RRPosted, tbl_Inv_RR.InvNumber, tbl_Inv_RR.InvDate, tbl_Inv_RR.InvGrossAmt, " & _
                " tbl_Inv_RR.InvNetAmt, tbl_Inv_RR.RRPrinted, tbl_Inv_RR.LastModifiedRR, tbl_Inv_RR.TotalCostRecd, " & _
                " tbl_Inv_RR.TotalNetCostRecd, tbl_Inv_RR.RRInvPosted, tbl_Inv_RR.BookType, tbl_Inv_RR.StockClerk, " & _
                " tbl_Inv_RR.Purchaser, tbl_Inv_RR.DeptHead " & _
                " FROM tbl_Inv_RR LEFT OUTER JOIN " & _
                " tbl_Inv_Supplier ON tbl_Inv_RR.SuppKey = tbl_Inv_Supplier.PK LEFT OUTER JOIN " & _
                " tbl_GL_Department ON tbl_Inv_RR.DeptKey = tbl_GL_Department.PK " & _
                " ORDER BY tbl_Inv_RR.RRNumber"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        If picPost.Visible = True Then Exit Sub
        If picAccDistribution.Visible = True Then Exit Sub
        If picAddRR.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_RR.PK, tbl_Inv_RR.PONumber, tbl_Inv_RR.PODate, tbl_Inv_RR.PRNumber, " & _
            " tbl_Inv_RR.SuppKey, tbl_Inv_RR.DeptKey, tbl_Inv_RR.Terms, tbl_Inv_RR.Remarks, tbl_Inv_RR.TotalCost, " & _
            " tbl_Inv_RR.TotalDiscount, tbl_Inv_RR.TotalNetCost, tbl_Inv_RR.RequestBy, tbl_Inv_RR.CheckedBy, " & _
            " tbl_Inv_RR.ApprovedBy, tbl_Inv_RR.Posted, tbl_Inv_RR.Printed, tbl_Inv_RR.LastModified, " & _
            " tbl_Inv_RR.RecdPartFull, tbl_Inv_RR.Discount, tbl_Inv_RR.TotalQty, tbl_Inv_RR.TotalRcd, " & _
            " tbl_Inv_Supplier.SupplierCode, tbl_Inv_Supplier.SupplierName, tbl_Inv_Supplier.Address1, " & _
            " tbl_Inv_Supplier.Address2, tbl_Inv_Supplier.Address3, tbl_Inv_Supplier.TelNo, tbl_Inv_Supplier.FaxNo, " & _
            " tbl_Inv_Supplier.Email, tbl_Inv_Supplier.ContactPerson, tbl_GL_Department.Code as DepartmentCode, " & _
            " tbl_GL_Department.DeptName as DepartmentName, tbl_Inv_RR.RRNumber, tbl_Inv_RR.RRDateTime, " & _
            " tbl_Inv_RR.RRPosted, tbl_Inv_RR.InvNumber, tbl_Inv_RR.InvDate, tbl_Inv_RR.InvGrossAmt, " & _
            " tbl_Inv_RR.InvNetAmt, tbl_Inv_RR.RRPrinted, tbl_Inv_RR.LastModifiedRR, tbl_Inv_RR.TotalCostRecd, " & _
            " tbl_Inv_RR.TotalNetCostRecd, tbl_Inv_RR.RRInvPosted, tbl_Inv_RR.BookType, tbl_Inv_RR.StockClerk, " & _
            " tbl_Inv_RR.Purchaser, tbl_Inv_RR.DeptHead " & _
            " FROM tbl_Inv_RR LEFT OUTER JOIN " & _
            " tbl_Inv_Supplier ON tbl_Inv_RR.SuppKey = tbl_Inv_Supplier.PK LEFT OUTER JOIN " & _
            " tbl_GL_Department ON tbl_Inv_RR.DeptKey = tbl_GL_Department.PK " & _
            " ORDER BY tbl_Inv_RR.RRNumber"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        If picPost.Visible = True Then Exit Sub
        If picAccDistribution.Visible = True Then Exit Sub
        If picAddRR.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_RR.PK, tbl_Inv_RR.PONumber, tbl_Inv_RR.PODate, tbl_Inv_RR.PRNumber, " & _
            " tbl_Inv_RR.SuppKey, tbl_Inv_RR.DeptKey, tbl_Inv_RR.Terms, tbl_Inv_RR.Remarks, tbl_Inv_RR.TotalCost, " & _
            " tbl_Inv_RR.TotalDiscount, tbl_Inv_RR.TotalNetCost, tbl_Inv_RR.RequestBy, tbl_Inv_RR.CheckedBy, " & _
            " tbl_Inv_RR.ApprovedBy, tbl_Inv_RR.Posted, tbl_Inv_RR.Printed, tbl_Inv_RR.LastModified, " & _
            " tbl_Inv_RR.RecdPartFull, tbl_Inv_RR.Discount, tbl_Inv_RR.TotalQty, tbl_Inv_RR.TotalRcd, " & _
            " tbl_Inv_Supplier.SupplierCode, tbl_Inv_Supplier.SupplierName, tbl_Inv_Supplier.Address1, " & _
            " tbl_Inv_Supplier.Address2, tbl_Inv_Supplier.Address3, tbl_Inv_Supplier.TelNo, tbl_Inv_Supplier.FaxNo, " & _
            " tbl_Inv_Supplier.Email, tbl_Inv_Supplier.ContactPerson, tbl_GL_Department.Code as DepartmentCode, " & _
            " tbl_GL_Department.DeptName as DepartmentName, tbl_Inv_RR.RRNumber, tbl_Inv_RR.RRDateTime, " & _
            " tbl_Inv_RR.RRPosted, tbl_Inv_RR.InvNumber, tbl_Inv_RR.InvDate, tbl_Inv_RR.InvGrossAmt, tbl_Inv_RR.InvGrossAmt, " & _
            " tbl_Inv_RR.InvNetAmt, tbl_Inv_RR.RRPrinted, tbl_Inv_RR.LastModifiedRR, tbl_Inv_RR.TotalCostRecd, " & _
            " tbl_Inv_RR.TotalNetCostRecd, tbl_Inv_RR.RRInvPosted, tbl_Inv_RR.BookType, tbl_Inv_RR.StockClerk, " & _
            " tbl_Inv_RR.Purchaser, tbl_Inv_RR.DeptHead " & _
            " FROM tbl_Inv_RR LEFT OUTER JOIN " & _
            " tbl_Inv_Supplier ON tbl_Inv_RR.SuppKey = tbl_Inv_Supplier.PK LEFT OUTER JOIN " & _
            " tbl_GL_Department ON tbl_Inv_RR.DeptKey = tbl_GL_Department.PK " & _
            " WHERE (tbl_Inv_RR.RRNumber < '" & sPONum & "') " & _
            " ORDER BY tbl_Inv_RR.RRNumber DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        If picPost.Visible = True Then Exit Sub
        If picAccDistribution.Visible = True Then Exit Sub
        If picAddRR.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_RR.PK, tbl_Inv_RR.PONumber, tbl_Inv_RR.PODate, tbl_Inv_RR.PRNumber, " & _
            " tbl_Inv_RR.SuppKey, tbl_Inv_RR.DeptKey, tbl_Inv_RR.Terms, tbl_Inv_RR.Remarks, tbl_Inv_RR.TotalCost, " & _
            " tbl_Inv_RR.TotalDiscount, tbl_Inv_RR.TotalNetCost, tbl_Inv_RR.RequestBy, tbl_Inv_RR.CheckedBy, " & _
            " tbl_Inv_RR.ApprovedBy, tbl_Inv_RR.Posted, tbl_Inv_RR.Printed, tbl_Inv_RR.LastModified, " & _
            " tbl_Inv_RR.RecdPartFull, tbl_Inv_RR.Discount, tbl_Inv_RR.TotalQty, tbl_Inv_RR.TotalRcd, " & _
            " tbl_Inv_Supplier.SupplierCode, tbl_Inv_Supplier.SupplierName, tbl_Inv_Supplier.Address1, " & _
            " tbl_Inv_Supplier.Address2, tbl_Inv_Supplier.Address3, tbl_Inv_Supplier.TelNo, tbl_Inv_Supplier.FaxNo, " & _
            " tbl_Inv_Supplier.Email, tbl_Inv_Supplier.ContactPerson, tbl_GL_Department.Code as DepartmentCode, " & _
            " tbl_GL_Department.DeptName as DepartmentName, tbl_Inv_RR.RRNumber, tbl_Inv_RR.RRDateTime, " & _
            " tbl_Inv_RR.RRPosted, tbl_Inv_RR.InvNumber, tbl_Inv_RR.InvDate, tbl_Inv_RR.InvGrossAmt, tbl_Inv_RR.InvGrossAmt, " & _
            " tbl_Inv_RR.InvNetAmt, tbl_Inv_RR.RRPrinted, tbl_Inv_RR.LastModifiedRR, tbl_Inv_RR.TotalCostRecd, " & _
            " tbl_Inv_RR.TotalNetCostRecd, tbl_Inv_RR.RRInvPosted, tbl_Inv_RR.BookType, tbl_Inv_RR.StockClerk, " & _
            " tbl_Inv_RR.Purchaser, tbl_Inv_RR.DeptHead " & _
            " FROM tbl_Inv_RR LEFT OUTER JOIN " & _
            " tbl_Inv_Supplier ON tbl_Inv_RR.SuppKey = tbl_Inv_Supplier.PK LEFT OUTER JOIN " & _
            " tbl_GL_Department ON tbl_Inv_RR.DeptKey = tbl_GL_Department.PK " & _
            " WHERE (tbl_Inv_RR.RRNumber > '" & sPONum & "') " & _
            " ORDER BY tbl_Inv_RR.RRNumber "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        If picPost.Visible = True Then Exit Sub
        If picAccDistribution.Visible = True Then Exit Sub
        If picAddRR.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_RR.PK, tbl_Inv_RR.PONumber, tbl_Inv_RR.PODate, tbl_Inv_RR.PRNumber, " & _
            " tbl_Inv_RR.SuppKey, tbl_Inv_RR.DeptKey, tbl_Inv_RR.Terms, tbl_Inv_RR.Remarks, tbl_Inv_RR.TotalCost, " & _
            " tbl_Inv_RR.TotalDiscount, tbl_Inv_RR.TotalNetCost, tbl_Inv_RR.RequestBy, tbl_Inv_RR.CheckedBy, " & _
            " tbl_Inv_RR.ApprovedBy, tbl_Inv_RR.Posted, tbl_Inv_RR.Printed, tbl_Inv_RR.LastModified, " & _
            " tbl_Inv_RR.RecdPartFull, tbl_Inv_RR.Discount, tbl_Inv_RR.TotalQty, tbl_Inv_RR.TotalRcd, " & _
            " tbl_Inv_Supplier.SupplierCode, tbl_Inv_Supplier.SupplierName, tbl_Inv_Supplier.Address1, " & _
            " tbl_Inv_Supplier.Address2, tbl_Inv_Supplier.Address3, tbl_Inv_Supplier.TelNo, tbl_Inv_Supplier.FaxNo, " & _
            " tbl_Inv_Supplier.Email, tbl_Inv_Supplier.ContactPerson, tbl_GL_Department.Code as DepartmentCode, " & _
            " tbl_GL_Department.DeptName as DepartmentName, tbl_Inv_RR.RRNumber, tbl_Inv_RR.RRDateTime, " & _
            " tbl_Inv_RR.RRPosted, tbl_Inv_RR.InvNumber, tbl_Inv_RR.InvDate, tbl_Inv_RR.InvGrossAmt, tbl_Inv_RR.InvGrossAmt, " & _
            " tbl_Inv_RR.InvNetAmt, tbl_Inv_RR.RRPrinted, tbl_Inv_RR.LastModifiedRR, tbl_Inv_RR.TotalCostRecd, " & _
            " tbl_Inv_RR.TotalNetCostRecd, tbl_Inv_RR.RRInvPosted, tbl_Inv_RR.BookType, tbl_Inv_RR.StockClerk, " & _
            " tbl_Inv_RR.Purchaser, tbl_Inv_RR.DeptHead " & _
            " FROM tbl_Inv_RR LEFT OUTER JOIN " & _
            " tbl_Inv_Supplier ON tbl_Inv_RR.SuppKey = tbl_Inv_Supplier.PK LEFT OUTER JOIN " & _
            " tbl_GL_Department ON tbl_Inv_RR.DeptKey = tbl_GL_Department.PK " & _
            " ORDER BY tbl_Inv_RR.RRNumber DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iBookType = IIf(IsNull(rs!BookType), 2, rs!BookType)
    iSupplier = rs!SuppKey
    iDept = rs!DeptKey
    sSupplierCode = rs!SupplierCode
    txtPONumber.Text = rs!PONumber
    txtPODate.Text = Format(rs!PODate, "mm/dd/yyyy")
    txtRefNo.Text = rs!PRNumber
    txtTerms.Text = rs!Terms
    txtSupplier.Text = rs!SupplierCode & " - " & rs!SupplierName
    txtAddress.Text = rs!Address1 & " " & rs!Address2 & " " & rs!Address3
    txtTelNo.Text = rs!TelNo
    txtFaxNo.Text = rs!FaxNo
    txtRemarks.Text = rs!Remarks
    txtDept.Text = UCase(rs!DepartmentName)
    txtDisc.Text = IIf(IsNull(rs!Discount), "", rs!Discount)
    
    txtRRNumber.Text = IIf(IsNull(rs!RRNumber), "", rs!RRNumber)
    txtRRDate.Text = Format(rs!RRDateTime, "mm/dd/yyyy")
    txtRRTime.Text = ""
    'If IsNull(rs!RRDateTime) = False Then
    '    txtRRDate.Text = Format(rs!RRDateTime, "mm/dd/yyyy")
    '    txtRRTime.Text = Format(rs!RRDateTime, "hh:mm:ss AM/PM")
    'End If
    
    lblTotalCost.Caption = Format(IIf(CDbl(rs!TotalCostRecd) = 0, rs!TotalCost, rs!TotalCostRecd), "#,##0.00")
    lblTotalNetCost.Caption = Format(IIf(CDbl(rs!TotalCostRecd) = 0, rs!TotalCostRecd, rs!TotalCostRecd), "#,##0.00")
    'lblTotalNetCost.Caption = Format(IIf(CDbl(rs!TotalNetCostRecd) = 0, rs!TotalNetCost, rs!TotalNetCostRecd), "#,##0.00")
    
    txtInvNumber.Text = IIf(IsNull(rs!InvNumber), "", rs!InvNumber)
    txtInvDate.Text = ""
    If IsNull(rs!InvDate) = False Then
        txtInvDate.Text = Format(rs!InvDate, "mm/dd/yyyy")
    End If
    txtInvGross.Text = Format(rs!InvGrossAmt, "#,##0.00")
    txtInvNet.Text = Format(rs!InvNetAmt, "#,##0.00")
    
    txtStockClerk.Text = rs!StockClerk
    txtPurchaser.Text = rs!Purchaser
    txtDeptHead.Text = rs!DeptHead
    
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModifiedRR), IIf(IsNull(rs!LastModified), "", rs!LastModified), rs!LastModifiedRR)
    Statusbar1.Panels(3).Text = IIf(rs!RRPrinted = 1, "PRINTED", "UNPRINTED")
    
    imgPosted.Visible = IIf(rs!RRPosted = 1, True, False)
    lblInvPosted.Visible = IIf(rs!RRInvPosted = 1, True, False)
    Toolbar1.Buttons(19).ToolTipText = IIf(lblInvPosted.Visible = True, "POST (F8)", IIf(rs!RRPosted = 1, "UNPOST (F8)", "POST (F8)"))
    
    SaveSetting App.EXEName, "PONumberRR", "PONumRR", IIf(IsNull(rs!RRNumber), "", rs!RRNumber) 'rs!PONumber
    
    CUSTOMIZE_DETAIL
    t = "SELECT tbl_Inv_RRDet.* " & _
        " FROM tbl_Inv_RRDet " & _
        " WHERE (RRKey = " & rs!PK & ") " & _
        " ORDER BY Line"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        With lstDetail.ListItems
            .Clear
            While Not rt.EOF
                Set x = .Add()
                x.Text = rt!ItemKey
                x.SubItems(1) = Format(rt!line, "0#")
                x.SubItems(2) = IIf(rt!Item_FA = 1, "Items", IIf(rt!Item_FA = 2, "Fixed Asset", ""))
                x.SubItems(3) = Format(rt!Qty, "#,##0.00")
                x.SubItems(4) = Format(rt!RecdQty, "#,##0.00")
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
                
                x.SubItems(8) = Format(IIf(CDbl(rt!CostRecd) = 0, rt!Cost, rt!CostRecd), "#,##0.00")
                x.SubItems(9) = Format(IIf(CDbl(rt!NetCostRecd) = 0, rt!NetCost, rt!NetCostRecd), "#,##0.00")
                x.SubItems(10) = Format(IIf(CDbl(rt!TotalNetCostRecd) = 0, rt!TotalNetCost, rt!TotalNetCostRecd), "#,##0.00")
                x.SubItems(11) = IIf(CDbl(rt!TotalCostRecd) = 0, rt!TotalCost, rt!TotalCostRecd)
                x.SubItems(12) = rt!Item_FA
                x.SubItems(13) = rt!R_Inv_Qty
                x.SubItems(14) = rt!R_Inv_Cost
                x.SubItems(15) = rt!R_Inv_NetCost
                x.SubItems(16) = IIf(IsNull(rt!Remarks), "", rt!Remarks)
                rt.MoveNext
            Wend
        End With
    End If
    rt.Close
    
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If picADSLine.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    If AccessRights("Receiving Report", "Add") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    picBody.Enabled = False
    picToolbar.Enabled = False
    picAddRR.ZOrder 0
    txtPONumAdd.Text = ""
    txtRRNumAdd.Text = ""
    picAddRR.Visible = True
    txtPONumAdd.SetFocus
    Exit Sub
End If
If picAccDistribution.Visible = True Then
    If is_DET_FOCUS = 0 Then Exit Sub
    With lstAccDistribution.ListItems
        If .Count > 1 Then
            Set x = .Add()
            x.Text = ""
            x.SubItems(1) = " "
            x.SubItems(2) = " "
            x.SubItems(3) = " "
            x.SubItems(4) = " "
            ROW = .Count
        Else
            If Trim(.Item(.Count).SubItems(1)) <> "" Then
                Set x = .Add()
                x.Text = ""
                x.SubItems(1) = " "
                x.SubItems(2) = " "
                x.SubItems(3) = " "
                x.SubItems(4) = " "
                ROW = .Count
            Else
                ROW = 1
            End If
        End If
        lstAccDistribution.ListItems(ROW).EnsureVisible
        lstAccDistribution.ListItems(ROW).Selected = True
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
    End With
End If
End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE = is_REFRESH Then
    'If Statusbar1.Panels(1).Text = "" Then Exit Sub
    If picSLine.Visible = True Then Exit Sub
    If picPost.Visible = True Then Exit Sub
    If picAddRR.Visible = True Then Exit Sub
    If picAccDistribution.Visible = True Then
        If is_DET_FOCUS = 0 Then Exit Sub
        If picADSLine.Visible = True Then Exit Sub
        If lblInvPosted.Visible = True Then Exit Sub
        With lstAccDistribution.ListItems
            If Trim(.Item(ROW).SubItems(1)) = "" Then Exit Sub
            txtAccountNo.Text = .Item(ROW).SubItems(1)
            txtAccountName.Text = .Item(ROW).SubItems(2)
            txtDebit.Text = .Item(ROW).SubItems(3)
            txtCredit.Text = .Item(ROW).SubItems(4)
            txtAccountNo1.Text = .Item(ROW).SubItems(1)
            txtAccountName1.Text = .Item(ROW).SubItems(2)
            txtDebit1.Text = .Item(ROW).SubItems(3)
            txtCredit1.Text = .Item(ROW).SubItems(4)
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
    
    If AccessRights("Receiving Report", "Edit") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If imgPosted.Visible = True Then MsgBox "ALREADY POSTED!             ", vbCritical, "Posted": Exit Sub
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_EDITTING
    'Me.Caption = "RECEIVING REPORT - EDIT"

ElseIf TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If picSLine.Visible = True Then Exit Sub
    If is_DET_FOCUS = 0 Then Exit Sub
    If imgPosted.Visible = True Then Exit Sub
    If Toolbar1.Buttons(3).Enabled = False Then Exit Sub
    If picAccDistribution.Visible = True Then
        
        Exit Sub
    End If
    With lstDetail.ListItems
        If Trim(.Item(ROW).SubItems(1)) = "" Then Exit Sub
        txtType.Text = .Item(ROW).SubItems(2)
        txtOrdd.Text = .Item(ROW).SubItems(3)
        txtRecd.Text = .Item(ROW).SubItems(4)
        txtUnit.Text = .Item(ROW).SubItems(5)
        txtItemCode.Text = .Item(ROW).SubItems(6)
        txtDescription.Text = .Item(ROW).SubItems(7)
        txtCost.Text = .Item(ROW).SubItems(8)
        txtNetCost.Text = .Item(ROW).SubItems(9)
        txtTotalNetCost.Text = .Item(ROW).SubItems(10)
        
        txtRecd1.Text = .Item(ROW).SubItems(4)
        txtCost1.Text = .Item(ROW).SubItems(8)
        txtNetCost1.Text = .Item(ROW).SubItems(9)
        
        picBody.Enabled = False
        picSLine.ZOrder 0
        picSLine.Visible = True
        TRANS_DETAIL = is_DET_EDITTING
        TOOLBARFUNC 2
        txtRecd.SetFocus
    End With
End If
End Sub

Private Sub PRESS_DELETE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picAccDistribution.Visible = True Then
        If is_DET_FOCUS = 0 Then Exit Sub
        If picADSLine.Visible = True Then Exit Sub
        If lblInvPosted.Visible = True Then Exit Sub
        With lstAccDistribution.ListItems
            If .Count > 1 Then
                .Remove ROW
                If CDbl(ROW) > CDbl(.Count) Then ROW = .Count
            Else
                .Item(1).SubItems(1) = " "
                .Item(1).SubItems(2) = " "
                .Item(1).SubItems(3) = " "
                .Item(1).SubItems(4) = " "
                ROW = 1
            End If
            lstAccDistribution.ListItems(ROW).EnsureVisible
            lstAccDistribution.ListItems(ROW).Selected = True
            isGLCodeChange = 1
        End With
        Exit Sub
    End If
    
    If picSLine.Visible = True Then Exit Sub
    If picPost.Visible = True Then Exit Sub
    If picAddRR.Visible = True Then Exit Sub
    If Statusbar1.Panels(1).Text = "" Then Exit Sub
    
    If AccessRights("Receiving Report", "Delete") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If imgPosted.Visible = True Then MsgBox "ALREADY POSTED!             ", vbCritical, "Posted": Exit Sub
    If MsgBox("ARE YOU SURE IN DELETING THIS RECORD?                    ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    On Error GoTo PG:
    ConnOmega.Execute "DELETE FROM tbl_Inv_RR WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    CLEARTEXT
    BROWSER GetSetting(App.EXEName, "PONumberRR", "PONumRR", ""), "is_PAGEDOWN"
    If Trim(txtPONumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "PONumberRR", "PONumRR", ""), "is_HOME"
    
'ElseIf TRANSACTIONTYPE = is_ADDING Or _
'TRANSACTIONTYPE = is_EDITTING Then

End If
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If picAddRR.Visible = True Then Exit Sub
If picAccDistribution.Visible = True Then
    If picSearchGLAccount.Visible = True Then Exit Sub
    If picADSLine.Visible = True Then Exit Sub
    If iBookType = 0 Then MsgBox "Please Select Book Type!                      ", vbCritical, "Error...": cmbBookType.SetFocus: Exit Sub
    On Error GoTo PH:
    With lstAccDistribution.ListItems
        a = 0
        ConnOmega.Execute "DELETE FROM tbl_Inv_RR_Account_Distribution WHERE (POKey = " & Statusbar1.Panels(1).Text & ")"
        For i = 1 To .Count
            If Trim(.Item(i).SubItems(1)) <> "" Then
                a = a + 1
                ConnOmega.Execute "INSERT INTO tbl_Inv_RR_Account_Distribution " & _
                                  " (POKey, Line, AccountCode, Debit, Credit) " & _
                                  " VALUES (" & Statusbar1.Panels(1).Text & ", " & _
                                  " " & a & ", '" & Trim(.Item(i).SubItems(1)) & "', " & _
                                  " " & CDbl(IIf(Trim(.Item(i).SubItems(3)) = "", 0, .Item(i).SubItems(3))) & ", " & _
                                  " " & CDbl(IIf(Trim(.Item(i).SubItems(4)) = "", 0, .Item(i).SubItems(4))) & ")"
            End If
        Next i
        ConnOmega.Execute "UPDATE tbl_Inv_RR SET BookType = " & iBookType & " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    End With
    isGLCodeChange = 0
    'b8TitleBar2_CLoseClick
    Exit Sub
End If

If Toolbar1.Buttons(7).Caption <> "Save" Then Exit Sub
If IsDate(txtRRDate.Text) = False Then MsgBox "Please Supply RR Date!                             ", vbCritical, "Error...": txtRRDate.SetFocus: Exit Sub
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    sRRNumber = Format(CDbl(RETURNTEXTVALUE(txtRRNumber)), "0000000#")
    
    Do
        s = "SELECT tbl_Inv_RR.* " & _
            " FROM tbl_Inv_RR " & _
            " WHERE (RRNumber = '" & sRRNumber & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If
        rs.Close
        sRRNumber = Format(CDbl(sRRNumber) + 1, "0000000#")
    Loop
    
    ConnOmega.Execute "INSERT INTO tbl_Inv_RR_Numbers " & _
                      " (RRNumber) " & _
                      " VALUES('" & sRRNumber & "')"
    
    ConnOmega.Execute "INSERT INTO tbl_Inv_RR" & _
                      " (RRNumber, PONumber, PODate, PRNumber, " & _
                      " SuppKey, Terms, Remarks, TotalCostRecd, " & _
                      " LastModifiedRR, DeptKey, Discount, TotalNetCostRecd, " & _
                      " InvNumber, InvGrossAmt, InvNetAmt, RRDateTime, StockClerk, " & _
                      " Purchaser, DeptHead) " & _
                      " VALUES('" & sRRNumber & "', '" & Trim(txtPONumber.Text) & "', " & _
                      " '" & FormatDateTime(txtPODate.Text, vbShortDate) & "', '" & Trim(txtRefNo.Text) & "', " & _
                      " " & iSupplier & ", '" & FORMATSQL(Trim(txtTerms.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtRemarks.Text)) & "', " & CDbl(lblTotalCost.Caption) & ",  " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                      " " & iDept & ", '" & Trim(txtDisc.Text) & "'," & CDbl(lblTotalNetCost.Caption) & ", " & _
                      " '" & FORMATSQL(Trim(txtInvNumber.Text)) & "', " & RETURNTEXTVALUE(txtInvGross) & ", " & _
                      " " & RETURNTEXTVALUE(txtInvNet) & ", '" & FormatDateTime(txtRRDate.Text, vbShortDate) & "', " & _
                      " '" & FORMATSQL(Trim(txtStockClerk.Text)) & "', '" & FORMATSQL(Trim(txtPurchaser.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtDeptHead.Text)) & "')"
    
    s = "SELECT PK" & _
        " FROM tbl_Inv_RR" & _
        " WHERE (RRNumber = '" & sRRNumber & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        If IsDate(txtInvDate.Text) = True Then
            ConnOmega.Execute "UPDATE tbl_Inv_RR " & _
                              " SET InvDate = '" & FormatDateTime(txtInvDate.Text, vbShortDate) & "' " & _
                              " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
        End If
        ConnOmega.Execute "DELETE FROM tbl_Inv_PODet WHERE (POKey = " & rs!PK & ")"
        dTotalRcd = 0: a = 0 'dblTotalCost = 0: dblTotalNetCost = 0: a = 0
        With lstDetail.ListItems
            If Trim(.Item(1).SubItems(1)) <> "" Then
                For i = 1 To .Count
                    If Trim(.Item(i).Text) <> "" Then
                        a = a + 1
                        dTotalRcd = dTotalRcd + CDbl(.Item(i).SubItems(4))
                        Select Case .Item(i).SubItems(12)
                            Case 1
                                t = "SELECT Unit, ConUnit, Unit2, ConUnit2 " & _
                                    " From tbl_Inv_Items " & _
                                    " WHERE (ItemCode = '" & .Item(i).SubItems(6) & "')"
                                If rt.State = adStateOpen Then rt.Close
                                rt.Open t, ConnOmega
                                If rt.RecordCount > 0 Then
                                    If Trim(.Item(i).SubItems(5)) = Trim(rt!Unit) Then
                                        dR_Inv_Qty = .Item(i).SubItems(4)
                                        dR_Inv_Cost = .Item(i).SubItems(8)
                                        dR_Inv_NetCost = .Item(i).SubItems(9)
                                    ElseIf Trim(.Item(i).SubItems(5)) = Trim(rt!Unit2) Then
                                        dR_Inv_Qty = Format(CDbl(.Item(i).SubItems(4)) * CDbl(rt!ConUnit), "#,##0.00")
                                        dR_Inv_Cost = Format(CDbl(.Item(i).SubItems(8)) / CDbl(rt!ConUnit), "#,##0.00")
                                        dR_Inv_NetCost = Format(CDbl(.Item(i).SubItems(9)) / CDbl(rt!ConUnit), "#,##0.00")
                                    End If
                                End If
                                rt.Close
                            Case Else
                                dR_Inv_Qty = .Item(i).SubItems(4)
                                dR_Inv_Cost = .Item(i).SubItems(8)
                                dR_Inv_NetCost = .Item(i).SubItems(9)
                        End Select
                        
                        ConnOmega.Execute "INSERT INTO tbl_Inv_RRDet" & _
                                          " (RRKey, Line, ItemKey, Qty, Unit, RecdQty, CostRecd, NetCostRecd, " & _
                                          " TotalNetCostRecd, TotalCostRecd, R_Inv_Qty, R_Inv_Cost, R_Inv_NetCost, " & _
                                          " Remarks, Item_FA) " & _
                                          " VALUES (" & rs!PK & ", " & a & ", " & .Item(i).Text & ", " & _
                                          " " & CDbl(.Item(i).SubItems(3)) & ", '" & .Item(i).SubItems(5) & "', " & _
                                          " " & CDbl(.Item(i).SubItems(4)) & ", " & CDbl(.Item(i).SubItems(8)) & ", " & _
                                          " " & CDbl(.Item(i).SubItems(9)) & ", " & CDbl(.Item(i).SubItems(10)) & ", " & _
                                          " " & CDbl(.Item(i).SubItems(11)) & ", " & CDbl(dR_Inv_Qty) & ", " & _
                                          " " & CDbl(dR_Inv_Cost) & ", " & CDbl(dR_Inv_NetCost) & ", " & _
                                          " '" & FORMATSQL(Trim(.Item(i).SubItems(16))) & "', " & .Item(i).SubItems(12) & ")"
                        
                    End If
                Next i
            End If
        End With
        ConnOmega.Execute "UPDATE tbl_Inv_RR SET TotalRcd = " & CDbl(dTotalRcd) & " WHERE (PK = " & rs!PK & ")"
    End If
    rs.Close
    
End If
If TRANSACTIONTYPE = is_EDITTING Then
    sRRNumber = Trim(txtRRNumber.Text)
    ConnOmega.Execute "UPDATE tbl_Inv_RR " & _
                      " SET TotalCostRecd = " & CDbl(lblTotalCost.Caption) & ", " & _
                      " TotalNetCostRecd = " & CDbl(lblTotalNetCost.Caption) & ", " & _
                      " InvNumber = '" & FORMATSQL(Trim(txtInvNumber.Text)) & "', " & _
                      " InvGrossAmt = " & RETURNTEXTVALUE(txtInvGross) & ", " & _
                      " InvNetAmt = " & RETURNTEXTVALUE(txtInvNet) & ", " & _
                      " LastModifiedRR = '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                      " StockClerk = '" & FORMATSQL(Trim(txtStockClerk.Text)) & "', " & _
                      " Purchaser = '" & FORMATSQL(Trim(txtPurchaser.Text)) & "', " & _
                      " DeptHead = '" & FORMATSQL(Trim(txtDeptHead.Text)) & "' " & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    If IsDate(txtInvDate.Text) = True Then
        ConnOmega.Execute "UPDATE tbl_Inv_RR " & _
                          " SET InvDate = '" & FormatDateTime(txtInvDate.Text, vbShortDate) & "', " & _
                          " LastModifiedRR = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                          " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    End If
    With lstDetail.ListItems
        For i = 1 To .Count
            If CDbl(IIf(IsNumeric(.Item(i).SubItems(1)) = False, 0, .Item(i).SubItems(1))) > 0 Then
                Select Case .Item(i).SubItems(12)
                    Case 1
                        s = "SELECT Unit, ConUnit, Unit2, ConUnit2 " & _
                            " From tbl_Inv_Items " & _
                            " WHERE (ItemCode = '" & .Item(i).SubItems(6) & "')"
                        If rs.State = adStateOpen Then rs.Close
                        rs.Open s, ConnOmega
                        If rs.RecordCount > 0 Then
                            If Trim(.Item(i).SubItems(5)) = Trim(rs!Unit) Then
                                dR_Inv_Qty = .Item(i).SubItems(4)
                                dR_Inv_Cost = .Item(i).SubItems(8)
                                dR_Inv_NetCost = .Item(i).SubItems(9)
                            ElseIf Trim(.Item(i).SubItems(5)) = Trim(rs!Unit2) Then
                                dR_Inv_Qty = Format(CDbl(.Item(i).SubItems(4)) * CDbl(rs!ConUnit), "#,##0.00")
                                dR_Inv_Cost = Format(CDbl(.Item(i).SubItems(8)) / CDbl(rs!ConUnit), "#,##0.00")
                                dR_Inv_NetCost = Format(CDbl(.Item(i).SubItems(9)) / CDbl(rs!ConUnit), "#,##0.00")
                            End If
                        End If
                        rs.Close
                    Case Else
                        dR_Inv_Qty = .Item(i).SubItems(4)
                        dR_Inv_Cost = .Item(i).SubItems(8)
                        dR_Inv_NetCost = .Item(i).SubItems(9)
                End Select
                
                ConnOmega.Execute "UPDATE tbl_Inv_RRDet " & _
                                  " SET RecdQty = " & CDbl(.Item(i).SubItems(4)) & ", " & _
                                  " CostRecd  = " & CDbl(.Item(i).SubItems(8)) & ", " & _
                                  " NetCostRecd  = " & CDbl(.Item(i).SubItems(9)) & ", " & _
                                  " TotalNetCostRecd = " & CDbl(.Item(i).SubItems(10)) & ", " & _
                                  " TotalCostRecd = " & CDbl(.Item(i).SubItems(11)) & ", " & _
                                  " R_Inv_Qty = " & CDbl(dR_Inv_Qty) & ", " & _
                                  " R_Inv_Cost = " & CDbl(dR_Inv_Cost) & ", " & _
                                  " R_Inv_NetCost = " & CDbl(dR_Inv_NetCost) & ", " & _
                                  " Remarks = '" & FORMATSQL(Trim(.Item(i).SubItems(16))) & "' " & _
                                  " WHERE (RRKey = " & Statusbar1.Panels(1).Text & ") " & _
                                  " AND (Line = " & CDbl(.Item(i).SubItems(1)) & ")"
            End If
        Next i
    End With
End If

SaveSetting App.EXEName, "RRStockClerk", "RRStockClerk", Trim(txtStockClerk.Text)
SaveSetting App.EXEName, "RRPurchaser", "RRPurchaser", Trim(txtPurchaser.Text)
SaveSetting App.EXEName, "RRDeptHead", "RRDeptHead", Trim(txtDeptHead.Text)

CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER sRRNumber, "is_LOAD"
'Me.Caption = "RECEIVING REPORT - BROWSE"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
PH:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSLine.Visible = True Then Exit Sub
If picPost.Visible = True Then Exit Sub
If picAccDistribution.Visible = True Then Exit Sub
If picAddRR.Visible = True Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub

PopupMenu MainFormPopupF.mnuRRFind, , 5500, 500

'CLEARTEXT
'TOOLBARFUNC 3
'TRANSACTIONTYPE = is_FINDING
'Me.Caption = "RECEIVING REPORT - FIND"
'txtPONumber.Locked = False
'txtPONumber.SetFocus
End Sub

Private Sub PRESS_F7()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picAddRR.Visible = True Then Exit Sub
If picSLine.Visible = True Then Exit Sub
If picAccDistribution.Visible = True Then Exit Sub
If picPost.Visible = True Then Exit Sub
If imgPosted.Visible = False Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("Receiving Report", "Post Inv") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If

s = "SELECT tbl_Inv_RR_Account_Distribution.AccountCode, " & _
    " tbl_GL_Accounts.AccountName, " & _
    " tbl_Inv_RR_Account_Distribution.Debit, " & _
    " tbl_Inv_RR_Account_Distribution.Credit, " & _
    " tbl_Inv_RR_Account_Distribution.Amount " & _
    " FROM tbl_Inv_RR_Account_Distribution LEFT OUTER JOIN " & _
    " tbl_GL_Accounts ON tbl_Inv_RR_Account_Distribution.AccountCode = tbl_GL_Accounts.AccountCode " & _
    " Where (tbl_Inv_RR_Account_Distribution.POKey = " & Statusbar1.Panels(1).Text & ") " & _
    " ORDER BY tbl_Inv_RR_Account_Distribution.Line"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount = 0 Then
    picToolbar.Enabled = False
    picBody.Enabled = False
    picGLAddAutoVAT.ZOrder 0
    chkWithVAT.Value = 1
    txtGLAddAutoVAT.Text = ""
    picGLAddAutoVAT.Visible = True
    txtGLAddAutoVAT.SetFocus
    rs.Close
    Exit Sub
End If
rs.Close

txtInvNumberP.Text = txtInvNumber.Text
txtInvDateP.Text = txtInvDate.Text
txtInvGrossP.Text = txtInvGross.Text
txtInvNetP.Text = txtInvNet.Text

lstAccDistribution.ListItems.Clear
Set x = lstAccDistribution.ListItems.Add()
x.Text = ""
x.SubItems(1) = " "
x.SubItems(2) = " "
x.SubItems(3) = " "
x.SubItems(4) = " "
x.SubItems(5) = "0"
a = 0: b = 0
s = "SELECT tbl_Inv_RR_Account_Distribution.AccountCode, " & _
    " tbl_GL_Accounts.AccountName, " & _
    " tbl_Inv_RR_Account_Distribution.Debit, " & _
    " tbl_Inv_RR_Account_Distribution.Credit, " & _
    " tbl_Inv_RR_Account_Distribution.Amount " & _
    " FROM tbl_Inv_RR_Account_Distribution LEFT OUTER JOIN " & _
    " tbl_GL_Accounts ON tbl_Inv_RR_Account_Distribution.AccountCode = tbl_GL_Accounts.AccountCode " & _
    " Where (tbl_Inv_RR_Account_Distribution.POKey = " & Statusbar1.Panels(1).Text & ") " & _
    " ORDER BY tbl_Inv_RR_Account_Distribution.Line"
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
picBody.Enabled = False
picAccDistribution.Height = 3615 '3015
picAccDistribution.Width = 7335
picAccDistribution.ZOrder 0
picAccDistribution.Visible = True
lstAccDistribution.SetFocus
End Sub

Private Sub PRESS_F8()
If picAddRR.Visible = True Then Exit Sub
With frmInvRR
    If .Statusbar1.Panels(1).Text = "" Then Exit Sub
    If .TRANSACTIONTYPE <> 0 Then Exit Sub
    If .picSLine.Visible = True Then Exit Sub
    If .picPost.Visible = True Then Exit Sub
    If AccessRights("Receiving Report", "Post Rcd") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If Trim(.txtInvNumber.Text) = "" Then MsgBox "Please Supply Invoice Number!              ", vbCritical, "Error...": .txtInvNumber.SetFocus: Exit Sub
    If IsDate(.txtInvDate.Text) = False Then MsgBox "Please Supply a Valid Invoice Date!                 ", vbCritical, "Error...": .txtInvDate.SetFocus: Exit Sub
    If RETURNTEXTVALUE(.txtInvGross) <= 0 Then MsgBox "Please Supply a Valid Amount!               ", vbCritical, "Error...": .txtInvGross.SetFocus: Exit Sub
    If RETURNTEXTVALUE(.txtInvNet) <= 0 Then MsgBox "Please Supply a Valid Amount!               ", vbCritical, "Error...": .txtInvNet.SetFocus: Exit Sub
    .BROWSER GetSetting(App.EXEName, "PONumberRR", "PONumRR", ""), "is_LOAD"
    If .imgPosted.Visible = True Then MsgBox "ALREADY POSTED!                     ", vbCritical, "Error...": Exit Sub
    
    .cmbLocation.Clear
    s = "SELECT tbl_Inv_Location.* " & _
        " FROM tbl_Inv_Location " & _
        " WHERE (Receiving = 1) " & _
        " ORDER BY LocName"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        .cmbLocation.AddItem rs!LocName
        .cmbLocation.ItemData(.cmbLocation.NewIndex) = rs!PK
        rs.MoveNext
    Wend
    rs.Close
    .txtRRDatePosting.Text = Format(Now, "mm/dd/yyyy")
    .txtRRNoPosting.Text = ""
    .picToolbar.Enabled = False
    .picBody.Enabled = False
    .picPost.ZOrder 0
    .picPost.Visible = True
    .cmbLocation.SetFocus

End With

'If Statusbar1.Panels(1).Text = "" Then Exit Sub
'If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
'If picSLine.Visible = True Then Exit Sub
'If picPost.Visible = True Then Exit Sub
'If Trim(txtInvNumber.Text) = "" Then MsgBox "Please Supply Invoice Number!              ", vbCritical, "Error...": txtInvNumber.SetFocus: Exit Sub
'If IsDate(txtInvDate.Text) = False Then MsgBox "Please Supply a Valid Invoice Date!                 ", vbCritical, "Error...": txtInvDate.SetFocus: Exit Sub
'If RETURNTEXTVALUE(txtInvGross) <= 0 Then MsgBox "Please Supply a Valid Amount!               ", vbCritical, "Error...": txtInvGross.SetFocus: Exit Sub
'If RETURNTEXTVALUE(txtInvNet) <= 0 Then MsgBox "Please Supply a Valid Amount!               ", vbCritical, "Error...": txtInvNet.SetFocus: Exit Sub
'BROWSER GetSetting(App.EXEName, "PONumberRR", "PONumRR", ""), "is_LOAD"
'If imgPosted.Visible = True Then MsgBox "ALREADY POSTED!                     ", vbCritical, "Error...": Exit Sub

'PopupMenu MainFormPopupF.mnuRRPosting, , Toolbar1.Buttons(19).Left, 500
'Exit Sub

'cmbLocation.Clear
's = "SELECT tbl_Inv_Location.* " & _
'    " FROM tbl_Inv_Location " & _
'    " ORDER BY LocName"
'If rs.State = adStateOpen Then rs.Close
'rs.Open s, ConnOmega
'While Not rs.EOF
'    cmbLocation.AddItem rs!LocName
'    cmbLocation.ItemData(cmbLocation.NewIndex) = rs!PK
'    rs.MoveNext
'Wend
'rs.Close
'txtRRDatePosting.Text = Format(Now, "mm/dd/yyyy")
'picToolbar.Enabled = False
'picBody.Enabled = False
'picPost.ZOrder 0
'picPost.Visible = True
'cmbLocation.SetFocus
End Sub

Private Sub PRESS_F9()
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picAddRR.Visible = True Then Exit Sub
If picSLine.Visible = True Then Exit Sub
If picPost.Visible = True Then Exit Sub
If picAccDistribution.Visible = True Then Exit Sub

If AccessRights("Receiving Report", "Print") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If

If RETURNLABELVALUE(lblTotalNetCost) <= 0 Then MsgBox "No Item/s to be Received!                      ", vbCritical, "Error...": Exit Sub

frmPrinter.LastPage = 1
frmPrinter.PRINT_TRANSACTION = 2
frmPrinter.picPageRange.Enabled = False
frmPrinter.txtCopies.Locked = True
frmPrinter.picCPI.Enabled = False
frmPrinter.Show 1

End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picGLAddAutoVAT.Visible = True Then cmdCancelAutoVAT_Click: Exit Sub
    If picSearchGLAccount.Visible = True Then cmdCancelGLAccount_Click: Exit Sub
    If picAddRR.Visible = True Then cmdCancelAddRR_Click: Exit Sub
    If picADSLine.Visible = True Then
        With lstAccDistribution.ListItems
            If TRANS_DETAIL = is_DET_ADDING Then
                If .Count > 1 Then
                    .Remove .Count
                    ROW = .Count
                Else
                    ROW = 1
                    .Item(ROW).SubItems(1) = " "
                    .Item(ROW).SubItems(2) = " "
                    .Item(ROW).SubItems(3) = " "
                    .Item(ROW).SubItems(4) = " "
                End If
            ElseIf TRANS_DETAIL = is_DET_EDITTING Then
                .Item(ROW).SubItems(1) = txtAccountNo1.Text
                .Item(ROW).SubItems(2) = txtAccountName1.Text
                .Item(ROW).SubItems(3) = txtDebit1.Text
                .Item(ROW).SubItems(4) = txtCredit1.Text
            End If
        End With
        picADSLine.Visible = False
        picAccDistribution.Enabled = True
        lstAccDistribution.SetFocus
        Exit Sub
    End If
    If picAccDistribution.Visible = True Then b8TitleBar2_CLoseClick: Exit Sub
    If picPost.Visible = True Then cmdCancel_Click: Exit Sub
    Unload Me
Else
    If picSLine.Visible = True Then
        With lstDetail.ListItems
            .Item(ROW).SubItems(4) = txtRecd1.Text
            .Item(ROW).SubItems(8) = txtCost1.Text
            .Item(ROW).SubItems(9) = txtNetCost1.Text
            .Item(ROW).SubItems(16) = txtSLRemarks1.Text
        End With
        picSLine.Visible = False
        picBody.Enabled = True
        lstDetail.SetFocus
        Exit Sub
    End If
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    TRANS_DETAIL = is_DET_REFRESH
    BROWSER GetSetting(App.EXEName, "PONumberRR", "PONumRR", ""), "is_LOAD"
    If Trim(txtPONumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "PONumberRR", "PONumRR", ""), "is_HOME"
    'Me.Caption = "RECEIVING REPORT - BROWSE"
    txtPONumber.SetFocus
End If
End Sub

Public Sub CLEARTEXT()
iSupplier = 0
iDept = 0
txtPONumber.Text = ""
txtPODate.Text = ""
txtRefNo.Text = ""
txtTerms.Text = ""
txtRRNumber.Text = ""
txtRRDate.Text = ""
txtRRTime.Text = ""
txtSupplier.Text = ""
txtAddress.Text = ""
txtTelNo.Text = ""
txtFaxNo.Text = ""
txtDept.Text = ""
txtDisc.Text = ""
txtVAT.Text = ""
txtRemarks.Text = ""
txtInvNumber.Text = ""
txtInvDate.Text = ""
txtInvGross.Text = ""
txtInvNet.Text = ""
txtStockClerk.Text = ""
txtPurchaser.Text = ""
txtDeptHead.Text = ""
lblTotalCost.Caption = "0.00"
lblTotalNetCost.Caption = "0.00"
imgPosted.Visible = False
lblInvPosted.Visible = False
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
Statusbar1.Panels(3).Text = ""
CUSTOMIZE_DETAIL
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
    x.SubItems(11) = " "
    x.SubItems(16) = " "
End With
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtInvNumberP.Locked = True
txtInvDateP.Locked = True
txtInvGrossP.Locked = True
txtInvNetP.Locked = True
txtPONumber.Locked = bln
txtAccountName.Locked = True
txtPODate.Locked = True
txtRefNo.Locked = True
txtTerms.Locked = True
txtRRNumber.Locked = True
txtRRNumber.BackColor = &HE0E0E0
txtPONumber.BackColor = &HE0E0E0
txtRRDate.Locked = True
txtRRTime.Locked = True
txtSupplier.Locked = True
txtAddress.Locked = True
txtTelNo.Locked = True
txtFaxNo.Locked = True
txtDept.Locked = True
txtDisc.Locked = True
txtVAT.Locked = True
txtRemarks.Locked = True
txtInvNumber.Locked = bln
txtInvDate.Locked = bln
txtInvGross.Locked = bln
txtInvNet.Locked = bln
txtStockClerk.Locked = bln
txtPurchaser.Locked = bln
txtDeptHead.Locked = bln
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
cmdCancel_Click
End Sub

Private Sub b8TitleBar2_CLoseClick()
If isGLCodeChange = 1 Then
    If MsgBox("Save Changes!                     ", vbCritical + vbYesNo, "Save") = vbYes Then
        PRESS_F5
    Else
        picToolbar.Enabled = True
        picBody.Enabled = True
        picAccDistribution.Visible = False
        Exit Sub
    End If
End If
picToolbar.Enabled = True
picBody.Enabled = True
picAccDistribution.Visible = False
End Sub

Private Sub b8TitleBar3_CLoseClick()
cmdCancelGLAccount_Click
End Sub

Private Sub b8TitleBar4_CLoseClick()
cmdCancelAutoVAT_Click
End Sub

Private Sub b8TitleBar5_CLoseClick()
cmdCancelAddRR_Click
End Sub

Private Sub cmbBookType_Click()
If cmbBookType.ListIndex = -1 Then Exit Sub
iBookType = cmbBookType.ItemData(cmbBookType.ListIndex)
End Sub

Private Sub cmbLocation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtRRDatePosting.SetFocus
End Sub

Private Sub cmdCancel_Click()
picPost.Visible = False
picToolbar.Enabled = True
picBody.Enabled = True
End Sub

Private Sub cmdCancelAddRR_Click()
picBody.Enabled = True
picToolbar.Enabled = True
picAddRR.Visible = False
End Sub

Private Sub cmdCancelAutoVAT_Click()
picBody.Enabled = True
picToolbar.Enabled = True
picGLAddAutoVAT.Visible = False
End Sub

Private Sub cmdCancelGLAccount_Click()
picSearchGLAccount.Visible = False
picADSLine.Enabled = True
txtAccountNo.SetFocus
End Sub

Private Sub cmdOK_Click()
If cmbLocation.ListIndex = -1 Then MsgBox "Please Select Location!                          ", vbCritical, "Error...": cmbLocation.SetFocus: Exit Sub
If IsDate(txtRRDatePosting.Text) = False Then MsgBox "Please Supply a Valid Date!                           ", vbCritical, "Error...": txtRRDatePosting.SetFocus: Exit Sub
'iRRNumberAuto = False
's = "SELECT TOP 1 tbl_Inv_RR_Number_Auto.* " & _
'    " FROM tbl_Inv_RR_Number_Auto " & _
'    " WHERE (EffectDate <= '" & FormatDateTime(txtRRDatePosting.Text, vbShortDate) & "') " & _
'    " ORDER BY EffectDate DESC"
'If rs.State = adStateOpen Then rs.Close
'rs.Open s, ConnOmega
'If rs.RecordCount > 0 Then
'    iRRNumberAuto = IIf(rs!AutomaticRRNumber = 1, True, False)
'End If
'rs.Close
'If iRRNumberAuto = False Then
'    If Trim(txtRRNoPosting.Text) = "" Then MsgBox "Please Supply RR Number!                   ", vbCritical, "Error...": txtRRNoPosting.SetFocus: Exit Sub
'    If IsNumeric(txtRRNoPosting.Text) = False Then MsgBox "Please Supply Numeric Value!                   ", vbCritical, "Error...": txtRRNoPosting.SetFocus: Exit Sub
'    s = "SELECT tbl_Inv_RR_Numbers.* " & _
'        " FROM tbl_Inv_RR_Numbers " & _
'        " WHERE (RRNumber = '" & Trim(txtRRNoPosting.Text) & "')"
'    If rs.State = adStateOpen Then rs.Close
'    rs.Open s, ConnOmega
'    If rs.RecordCount > 0 Then
'        MsgBox "Found Duplicate RR Number!                      ", vbCritical, "Error..."
'        rs.Close
'        txtRRNoPosting.SetFocus
'        Exit Sub
'    End If
'    rs.Close
'End If

If MsgBox("ARE YOU SURE YOU WANT TO POST THIS TRANSACTION?                          ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
With lstDetail.ListItems
    
    
'    sRRNumber = ""
'    If iRRNumberAuto = True Then
'        s = "SELECT tbl_Inv_RR_Number.* " & _
'            " FROM tbl_Inv_RR_Number " & _
'            " WHERE (RRYear = " & Year(Now) & ")"
'        If rs.State = adStateOpen Then rs.Close
'        rs.Open s, ConnOmega
'        If rs.RecordCount > 0 Then
'            sRRNumber = rs!RRNumber
'        Else
'            sRRNumber = Format(Year(FormatDateTime(txtRRDatePosting.Text, vbShortDate)), "000#") & "0000"
'        End If
'        rs.Close
'
'        Do
'            's = "SELECT tbl_Inv_Items_Transaction.* " & _
'                " FROM tbl_Inv_Items_Transaction " & _
'                " WHERE (DocType = 2) " & _
'                " AND (DocNumber = '" & sRRNumber & "')"
'            s = "SELECT tbl_Inv_RR_Numbers.* " & _
'                " FROM tbl_Inv_RR_Numbers " & _
'                " WHERE (RRNumber = '" & sRRNumber & "')"
'            If rs.State = adStateOpen Then rs.Close
'            rs.Open s, ConnOmega
'            If rs.RecordCount = 0 Then
'                rs.Close
'                Exit Do
'            End If
'            rs.Close
'            sRRNumber = Format(CDbl(sRRNumber) + 1, "0000000#")
'        Loop
'
'    Else
'
'        sRRNumber = Format(RETURNTEXTVALUE(txtRRNumber), "0000000#")
'
'    End If
    
    sRRNumber = Format(RETURNTEXTVALUE(txtRRNumber), "0000000#")
    ' Inventory / FA Lapsing
    For i = 1 To .Count
        'Items
        If CDbl(IIf(IsNumeric(.Item(i).SubItems(12)) = False, 0, .Item(i).SubItems(12))) = 1 Then
            s = "SELECT tbl_Inv_Items.* " & _
                " FROM tbl_Inv_Items " & _
                " WHERE (PK = " & .Item(i).Text & ")"
            If rs.State = adStateOpen Then rs.Close
            rs.Open s, ConnOmega
            If rs.RecordCount > 0 Then
                dVATable = .Item(i).SubItems(15)
                dNetVAT = NET_OF_VAT(FormatDateTime(txtRRDatePosting.Text, vbShortDate), dVATable, rs!PK)
                dVAT = CDbl(dVATable) - CDbl(dNetVAT)
                ConnOmega.Execute "INSERT INTO tbl_Inv_Items_Transaction " & _
                                  " (ItemKey, Cleared, InOut, DocType, DocNumber, DocDate, " & _
                                  " Location, QuantityIn, Cost, NetCost, LogInName, NetVAT) " & _
                                  " VALUES (" & rs!PK & ", 0, 'I', 2, '" & sRRNumber & "', " & _
                                  " '" & FormatDateTime(txtRRDatePosting.Text, vbShortDate) & "', " & _
                                  " " & cmbLocation.ItemData(cmbLocation.ListIndex) & ", " & _
                                  " " & .Item(i).SubItems(13) & ", " & .Item(i).SubItems(14) & ", " & _
                                  " " & .Item(i).SubItems(15) & ", '" & gbl_UserName & "', " & _
                                  " " & CDbl(dNetVAT) & ")"
                
                t = "SELECT tbl_Inv_Items_Cost.* " & _
                    " FROM tbl_Inv_Items_Cost " & _
                    " WHERE (ItemKey = " & rs!PK & ") " & _
                    " AND (EffectDate = '" & FormatDateTime(txtRRDatePosting.Text, vbShortDate) & "')"
                If rt.State = adStateOpen Then rt.Close
                rt.Open t, ConnOmega
                If rt.RecordCount = 0 Then
                    ConnOmega.Execute "INSERT INTO tbl_Inv_Items_Cost " & _
                                      " (ItemKey, EffectDate, Cost) " & _
                                      " VALUES (" & rs!PK & ", " & _
                                      " '" & FormatDateTime(txtRRDatePosting.Text, vbShortDate) & "', " & _
                                      " " & .Item(i).SubItems(15) & ")"
                Else
                    ConnOmega.Execute "UPDATE tbl_Inv_Items_Cost " & _
                                      " SET Cost = " & .Item(i).SubItems(15) & " " & _
                                      " WHERE (ItemKey = " & rs!PK & ") " & _
                                      " AND (EffectDate = '" & FormatDateTime(txtRRDatePosting.Text, vbShortDate) & "')"
                End If
                rt.Close
                
            End If
            rs.Close
        'Fixed Assets
        ElseIf CDbl(IIf(IsNumeric(.Item(i).SubItems(12)) = False, 0, .Item(i).SubItems(12))) = 2 Then
            s = "SELECT tbl_FA_Items.* " & _
                " FROM tbl_FA_Items " & _
                " WHERE (PK = " & .Item(i).Text & ")"
            If rs.State = adStateOpen Then rs.Close
            rs.Open s, ConnOmega
            If rs.RecordCount > 0 Then
                ConnOmega.Execute "INSERT INTO tbl_FA_Items_Lapsing " & _
                                  " (FAKey, DepreciationDate, DepreciationAmount) " & _
                                  " VALUES (" & rs!PK & ", '" & FormatDateTime(txtRRDatePosting.Text, vbShortDate) & "', " & _
                                  " '" & .Item(i).SubItems(10) & "')"
                ConnOmega.Execute "UPDATE tbl_FA_Items " & _
                                  " SET DateAcquired = '" & FormatDateTime(txtRRDatePosting.Text, vbShortDate) & "', " & _
                                  " Cost = '" & .Item(i).SubItems(10) & "' " & _
                                  " WHERE (PK = " & rs!PK & ")"
            End If
            rs.Close
        End If
    Next i
    
           
'    Arr = Split(Trim(txtSupplier.Text), " - ", -1, 1)
'    s = "SELECT SupplierCode, SupplierName, Type " & _
'        " FROM tbl_Inv_Supplier " & _
'        " WHERE (SupplierCode = '" & FORMATSQL(CStr(Arr(0))) & "')"
'    If rs.State = adStateOpen Then rs.Close
'    rs.Open s, ConnOmega
'    If rs.RecordCount > 0 Then
'        If rs!Type = 1 Then
'
'            dAPTrade = 0: dVat = 0
'            dAPTrade = Format(RETURNTEXTVALUE(txtInvNet) / 1.12, "#,##0.00")
'            dVat = RETURNTEXTVALUE(txtInvNet) - CDbl(dAPTrade)
'
'            '---- G L
'            ' A/P Trade
'            ConnOmega.Execute "INSERT INTO tbl_GL_Transaction " & _
'                              " (GLCode, DocDate, DocNumber, InvoiceNumber, InvoiceDate, " & _
'                              " SupplierCode, SupplierName, Particulars, Amount) " & _
'                              " VALUES ('201201', '" & FormatDateTime(txtRRDatePosting.Text, vbShortDate) & "', " & _
'                              " '" & sRRNumber & "', '" & FORMATSQL(Trim(txtInvNumber.Text)) & "', " & _
'                              " '" & FormatDateTime(txtInvDate.Text, vbShortDate) & "', " & _
'                              " '" & FORMATSQL(CStr(Arr(0))) & "', '" & FORMATSQL(CStr(Arr(1))) & "', " & _
'                              " 'PURCHASES', " & CDbl(dAPTrade) & ")"
'            ' Vat
'            ConnOmega.Execute "INSERT INTO tbl_GL_Transaction " & _
'                              " (GLCode, DocDate, DocNumber, InvoiceNumber, InvoiceDate, " & _
'                              " SupplierCode, SupplierName, Particulars, Amount) " & _
'                              " VALUES ('201224', '" & FormatDateTime(txtRRDatePosting.Text, vbShortDate) & "', " & _
'                              " '" & sRRNumber & "', '" & FORMATSQL(Trim(txtInvNumber.Text)) & "', " & _
'                              " '" & FormatDateTime(txtInvDate.Text, vbShortDate) & "', " & _
'                              " '" & FORMATSQL(CStr(Arr(0))) & "', '" & FORMATSQL(CStr(Arr(1))) & "', " & _
'                              " 'INPUT TAX', " & CDbl(dVat) & ")"
'        End If
'    End If
'    rs.Close

End With

'ConnOmega.Execute "INSERT INTO tbl_Inv_RR_Numbers " & _
'                  " (RRNumber) " & _
'                  " VALUES('" & sRRNumber & "')"
                  
'ConnOmega.Execute "UPDATE tbl_Inv_RR " & _
                  " SET RRPosted = 1, RRNumber = '" & sRRNumber & "', " & _
                  " RRDateTime = '" & FormatDateTime(txtRRDatePosting.Text, vbShortDate) & "', " & _
                  " LastModifiedRR = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                  " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
                 
ConnOmega.Execute "UPDATE tbl_Inv_RR " & _
                  " SET RRPosted = 1, " & _
                  " LastModifiedRR = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                  " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
                  

'If iRRNumberAuto = True Then
'    ConnOmega.Execute "UPDATE tbl_Inv_RR_Number " & _
'                      " SET RRNumber = '" & sRRNumber & "' " & _
'                      " WHERE (RRYear = " & Year(Now) & ")"
'End If

cmdCancel_Click

BROWSER GetSetting(App.EXEName, "PONumberRR", "PONumRR", ""), "is_LOAD"

End Sub

Private Sub cmdOKAddRR_Click()
If RETURNTEXTVALUE(txtPONumAdd) = 0 Then Exit Sub
If RETURNTEXTVALUE(txtRRNumAdd) = 0 Then Exit Sub
If IsDate(txtRRDateAdd.Text) = False Then Exit Sub

s = "SELECT tbl_Inv_PO.* " & _
    " FROM tbl_Inv_PO " & _
    " WHERE (PONumber = '" & Format(RETURNTEXTVALUE(txtPONumAdd), "0000000#") & "') " & _
    " AND (Posted = 1)"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount = 0 Then
    MsgBox "'" & Format(RETURNTEXTVALUE(txtPONumAdd), "0000000#") & "' Not Found!                   ", vbCritical, "Error..."
    txtPONumAdd.SetFocus
    Exit Sub
Else
    CLEARTEXT
    t = "SELECT TOP 1 tbl_Inv_PO.PK, tbl_Inv_PO.PONumber, tbl_Inv_PO.PODate, tbl_Inv_PO.PRNumber, " & _
        " tbl_Inv_PO.SuppKey, tbl_Inv_PO.DeptKey, tbl_Inv_PO.Terms, tbl_Inv_PO.Remarks, tbl_Inv_PO.TotalCost, " & _
        " tbl_Inv_PO.TotalDiscount, tbl_Inv_PO.TotalNetCost, tbl_Inv_PO.RequestBy, tbl_Inv_PO.CheckedBy, " & _
        " tbl_Inv_PO.ApprovedBy, tbl_Inv_PO.Posted, tbl_Inv_PO.Printed, tbl_Inv_PO.LastModified, " & _
        " tbl_Inv_PO.RecdPartFull, tbl_Inv_PO.Discount, tbl_Inv_PO.TotalQty, tbl_Inv_PO.TotalRcd, " & _
        " tbl_Inv_Supplier.SupplierCode, tbl_Inv_Supplier.SupplierName, tbl_Inv_Supplier.Address1, " & _
        " tbl_Inv_Supplier.Address2, tbl_Inv_Supplier.Address3, tbl_Inv_Supplier.TelNo, tbl_Inv_Supplier.FaxNo, " & _
        " tbl_Inv_Supplier.Email, tbl_Inv_Supplier.ContactPerson, tbl_GL_Department.Code as DepartmentCode, " & _
        " tbl_GL_Department.DeptName as DepartmentName, tbl_Inv_PO.RRNumber, tbl_Inv_PO.RRDateTime, " & _
        " tbl_Inv_PO.RRPosted, tbl_Inv_PO.InvNumber, tbl_Inv_PO.InvDate, tbl_Inv_PO.InvGrossAmt, " & _
        " tbl_Inv_PO.InvNetAmt, tbl_Inv_PO.RRPrinted, tbl_Inv_PO.LastModifiedRR, tbl_Inv_PO.TotalCostRecd, " & _
        " tbl_Inv_PO.TotalNetCostRecd, tbl_Inv_PO.RRInvPosted, tbl_Inv_PO.BookType " & _
        " FROM tbl_Inv_PO LEFT OUTER JOIN " & _
        " tbl_Inv_Supplier ON tbl_Inv_PO.SuppKey = tbl_Inv_Supplier.PK LEFT OUTER JOIN " & _
        " tbl_GL_Department ON tbl_Inv_PO.DeptKey = tbl_GL_Department.PK " & _
        " WHERE (tbl_Inv_PO.PK = " & rs!PK & ") " & _
        " ORDER BY tbl_Inv_PO.PONumber"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        iBookType = IIf(IsNull(rt!BookType), 2, rt!BookType)
        iSupplier = rt!SuppKey
        iDept = rt!DeptKey
        sSupplierCode = rt!SupplierCode
        txtPONumber.Text = rt!PONumber
        txtPODate.Text = Format(rt!PODate, "mm/dd/yyyy")
        txtRefNo.Text = rt!PRNumber
        txtTerms.Text = rt!Terms
        txtSupplier.Text = rt!SupplierCode & " - " & rt!SupplierName
        txtAddress.Text = rt!Address1 & " " & rt!Address2 & " " & rt!Address3
        txtTelNo.Text = rt!TelNo
        txtFaxNo.Text = rt!FaxNo
        txtRemarks.Text = rt!Remarks
        txtDept.Text = UCase(rt!DepartmentName)
        txtDisc.Text = IIf(IsNull(rt!Discount), "", rt!Discount)
    
        txtRRNumber.Text = Format(RETURNTEXTVALUE(txtRRNumAdd), "0000000#")
        txtRRDate.Text = Format(FormatDateTime(txtRRDateAdd.Text, vbShortDate), "mm/dd/yyyy")
        txtRRTime.Text = Format(Date, "hh:mm:ss AM/PM")
        
        CUSTOMIZE_DETAIL
        v = "SELECT tbl_Inv_PODet.* " & _
            " FROM tbl_Inv_PODet " & _
            " WHERE (POKey = " & rt!PK & ") " & _
            " ORDER BY Line"
        If rv.State = adStateOpen Then rv.Close
        rv.Open v, ConnOmega
        If rv.RecordCount > 0 Then
            With lstDetail.ListItems
                .Clear
                While Not rv.EOF
                    Set x = .Add()
                    x.Text = rv!ItemKey
                    x.SubItems(1) = Format(rv!line, "0#")
                    x.SubItems(2) = IIf(rv!Item_FA = 1, "Items", IIf(rv!Item_FA = 2, "Fixed Asset", ""))
                    x.SubItems(3) = Format(rv!Qty, "#,##0.00")
                    x.SubItems(4) = Format(rv!RecdQty, "#,##0.00")
                    x.SubItems(5) = rv!Unit
                    Select Case rv!Item_FA
                        Case 1
                            u = "SELECT ItemCode, ItemDesc " & _
                                " FROM tbl_Inv_Items " & _
                                " WHERE (PK = " & rv!ItemKey & ")"
                        Case 2
                            u = "SELECT Code as ItemCode, " & _
                                " Description as ItemDesc " & _
                                " FROM tbl_FA_Items " & _
                                " WHERE (PK = " & rv!ItemKey & ")"
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
                    
                    x.SubItems(8) = Format(IIf(CDbl(rv!CostRecd) = 0, rv!Cost, rv!CostRecd), "#,##0.00")
                    x.SubItems(9) = Format(IIf(CDbl(rv!NetCostRecd) = 0, rv!NetCost, rv!NetCostRecd), "#,##0.00")
                    x.SubItems(10) = Format(rv!TotalNetCostRecd, "#,##0.00") 'Format(IIf(CDbl(rv!TotalNetCostRecd) = 0, rv!TotalNetCost, rv!TotalNetCostRecd), "#,##0.00")
                    x.SubItems(11) = rv!TotalCostRecd 'IIf(CDbl(rv!TotalCostRecd) = 0, rv!TotalCost, rv!TotalCostRecd)
                    x.SubItems(12) = rv!Item_FA
                    x.SubItems(13) = rv!R_Inv_Qty
                    x.SubItems(14) = rv!R_Inv_Cost
                    x.SubItems(15) = rv!R_Inv_NetCost
                    x.SubItems(16) = IIf(IsNull(rv!Remarks), "", rv!Remarks)
                    rv.MoveNext
                Wend
            End With
        End If
        rv.Close
        
    End If
    rt.Close
    
    txtStockClerk.Text = GetSetting(App.EXEName, "RRStockClerk", "RRStockClerk", "")
    txtPurchaser.Text = GetSetting(App.EXEName, "RRPurchaser", "RRPurchaser", "")
    txtDeptHead.Text = GetSetting(App.EXEName, "RRDeptHead", "RRDeptHead", "")
    
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_ADDING
    'Me.Caption = "RECEIVING REPORT - ADD"
    cmdCancelAddRR_Click
End If
rs.Close
End Sub

Private Sub cmdOKAutoVAT_Click()
If lstGLAutoVAT.ListIndex = -1 Then Exit Sub
'Dim dVAT
'dVAT = Format(RETURNTEXTVALUE(txtInvNet) / gbl_VAT, "#,##0.00")
dVAT = NET_OF_VAT(FormatDateTime(txtRRDate.Text, vbShortDate), RETURNTEXTVALUE(txtInvNet))
Arr = Split(lstGLAutoVAT.List(lstGLAutoVAT.ListIndex), " : ", -1, 1)
a = 0: b = 0
With lstAccDistribution.ListItems
    .Clear
    Set x = .Add()
    x.Text = ""
    x.SubItems(1) = Arr(0)
    x.SubItems(2) = Arr(1)
    If chkWithVAT.Value = 1 Then
        x.SubItems(3) = Format(dVAT, "#,##0.00")
        x.SubItems(4) = " "
        x.SubItems(5) = dVAT
        a = a + CDbl(dVAT)
        b = b + 0
        
        Set x = .Add()
        x.Text = ""
        x.SubItems(1) = "201228"
        x.SubItems(2) = "VAT PAYABLE"
        x.SubItems(3) = Format(RETURNTEXTVALUE(txtInvNet) - CDbl(dVAT), "#,##0.00")
        x.SubItems(4) = " "
        x.SubItems(5) = RETURNTEXTVALUE(txtInvNet) - CDbl(dVAT)
        a = a + CDbl(Format(RETURNTEXTVALUE(txtInvNet) - CDbl(dVAT), "#,##0.00"))
        b = b + 0
    Else
        x.SubItems(3) = Format(RETURNTEXTVALUE(txtInvNet), "#,##0.00")
        x.SubItems(4) = " "
        x.SubItems(5) = RETURNTEXTVALUE(txtInvNet)
        a = a + CDbl(RETURNTEXTVALUE(txtInvNet))
        b = b + 0
    End If
    
    
    Set x = .Add()
    x.Text = ""
    x.SubItems(1) = "201201"
    x.SubItems(2) = "AP - TRADE"
    x.SubItems(3) = " "
    x.SubItems(4) = Format(RETURNTEXTVALUE(txtInvNet), "#,##0.00")
    x.SubItems(5) = RETURNTEXTVALUE(txtInvNet) * -1
    a = a + 0
    b = b + CDbl(Format(RETURNTEXTVALUE(txtInvNet), "#,##0.00"))
End With

txtInvNumberP.Text = txtInvNumber.Text
txtInvDateP.Text = txtInvDate.Text
txtInvGrossP.Text = txtInvGross.Text
txtInvNetP.Text = txtInvNet.Text

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

isGLCodeChange = 1

picGLAddAutoVAT.Visible = False
picAccDistribution.Height = 3615 '3015
picAccDistribution.Width = 7335
picAccDistribution.ZOrder 0
picAccDistribution.Visible = True
lstAccDistribution.SetFocus

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

If isGLCodeChange = 1 Then PRESS_F5

If MsgBox("ARE YOU SURE IN POSTING THIS INVOICE?                        ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub

On Error GoTo PG:
'-- SL
'ConnOmega.Execute "INSERT INTO tbl_Inv_Supplier_SL " & _
                  " (SupplierKey, GLCode, DocNumber, DocDate, InvoiceNumber, InvoiceDate, Description, Amount, iType, Reference) " & _
                  " VALUES (" & iSupplier & ", '" & sSupplierCode & "', '" & txtRRNumber.Text & "', '" & FormatDateTime(txtRRDate.Text, vbShortDate) & "', " & _
                  " '" & Trim(txtInvNumberP.Text) & "', '" & FormatDateTime(txtInvDateP.Text, vbShortDate) & "', 'PURCHASES', " & _
                  " " & RETURNTEXTVALUE(txtInvNetP) * -1 & ", 2, '" & "RR# " & Trim(txtRRNumber.Text) & ", INV# " & Trim(txtInvNumberP.Text) & "')"
'-- GL
With lstAccDistribution.ListItems
    Arr = Split(Trim(txtSupplier.Text), " - ", -1, 1)
    For i = 1 To .Count
        If Trim(.Item(i).SubItems(1)) <> "" Then
            'dAmount = CDbl(.Item(i).SubItems(5))
            ConnOmega.Execute "INSERT INTO tbl_GL_Transaction " & _
                              " (GLCode, DocDate, DocNumber, SupplierCode, SupplierName, InvoiceNumber, InvoiceDate, PayeeKey, " & _
                              " PayeeType, BookType, Debit, Credit) " & _
                              " VALUES ('" & .Item(i).SubItems(1) & "', '" & FormatDateTime(txtRRDate.Text, vbShortDate) & "', " & _
                              " '" & txtRRNumber.Text & "', '" & FORMATSQL(CStr(Arr(0))) & "', '" & FORMATSQL(CStr(Arr(1))) & "', " & _
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
                                          " '" & txtRRNumber.Text & "', '" & FormatDateTime(txtRRDate.Text, vbShortDate) & "', " & _
                                          " '" & Trim(txtInvNumberP.Text) & "', '" & FormatDateTime(txtInvDateP.Text, vbShortDate) & "', " & _
                                          " 'PURCHASES', " & iBookType & ", '" & "RR# " & Trim(txtRRNumber.Text) & ", INV# " & Trim(txtInvNumberP.Text) & "', " & _
                                          " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(3)) = False, 0, .Item(i).SubItems(3))) & ", " & _
                                          " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(4)) = False, 0, .Item(i).SubItems(4))) & ")"
                    Else
                        ConnOmega.Execute "INSERT INTO tbl_Inv_Supplier_SL " & _
                                          " (SupplierKey, GLCode, DocNumber, DocDate, InvoiceNumber, " & _
                                          " InvoiceDate, Description, iType, Reference, Debit, Credit) " & _
                                          " VALUES (" & rt!SupplierKey & ", '" & .Item(i).SubItems(1) & "', " & _
                                          " '" & txtRRNumber.Text & "', '" & FormatDateTime(txtRRDate.Text, vbShortDate) & "', " & _
                                          " '" & Trim(txtInvNumberP.Text) & "', '" & FormatDateTime(txtInvDateP.Text, vbShortDate) & "', " & _
                                          " 'PURCHASES', " & iBookType & ", '" & "RR# " & Trim(txtRRNumber.Text) & ", INV# " & Trim(txtInvNumberP.Text) & "', " & _
                                          " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(3)) = False, 0, .Item(i).SubItems(3))) & ", " & _
                                          " " & CDbl(IIf(IsNumeric(.Item(i).SubItems(4)) = False, 0, .Item(i).SubItems(4))) & ")"
                    End If
                    'ConnOmega.Execute "INSERT INTO tbl_Inv_Supplier_SL " & _
                                      " (SupplierKey, GLCode, DocNumber, DocDate, " & _
                                      " Description, CheckNumber, Reference, iType, Amount) " & _
                                      " VALUES (" & iSupplier & ", '" & .Item(i).SubItems(1) & "', " & _
                                      " '" & Trim(txtCVNumber.Text) & "', " & _
                                      " '" & FormatDateTime(txtCVDate.Text, vbShortDate) & "', " & _
                                      " '', '" & Trim(txtCheckNumber.Text) & "', " & _
                                      " '" & "RR# " & Trim(txtRRNumber.Text) & ", INV# " & Trim(txtInvNumberP.Text) & "', " & _
                                      " 2, " & CDbl(dAmount) & ")"
                End If
            End If
            rt.Close
        End If
    Next i
End With

ConnOmega.Execute "UPDATE tbl_Inv_RR " & _
                  " SET RRInvPosted = 1 " & _
                  " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"

picToolbar.Enabled = True
picBody.Enabled = True
picAccDistribution.Visible = False

BROWSER GetSetting(App.EXEName, "PONumberRR", "PONumRR", ""), "is_LOAD"

Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub Form_Activate()
If TRANSACTIONTYPE = is_REFRESH Then BROWSER GetSetting(App.EXEName, "PONumberRR", "PONumRR", ""), "is_LOAD"
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
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "PONumberRR", "PONumRR", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "PONumberRR", "PONumRR", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "PONumberRR", "PONumRR", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "PONumberRR", "PONumRR", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANS_DETAIL = is_DET_REFRESH
BROWSER GetSetting(App.EXEName, "PONumberRR", "PONumRR", ""), "is_LOAD"
If Trim(txtPONumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "PONumberRR", "PONumRR", ""), "is_HOME"
is_DET_FOCUS = 0
isGLCodeFocus = 0
isGLCodeChange = 0
'Me.Caption = "RECEIVING REPORT - BROWSE"
tmp = SetWindowLong(txtSearchGLAccount.hwnd, GWL_STYLE, GetWindowLong(txtSearchGLAccount.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtGLAddAutoVAT.hwnd, GWL_STYLE, GetWindowLong(txtGLAddAutoVAT.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtStockClerk.hwnd, GWL_STYLE, GetWindowLong(txtStockClerk.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPurchaser.hwnd, GWL_STYLE, GetWindowLong(txtPurchaser.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtDeptHead.hwnd, GWL_STYLE, GetWindowLong(txtDeptHead.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSLRemarks.hwnd, GWL_STYLE, GetWindowLong(txtSLRemarks.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picGLAddAutoVAT.Visible = True Then Cancel = -1
If picADSLine.Visible = True Then Cancel = -1
If picAccDistribution.Visible = True Then Cancel = -1
If picSLine.Visible = True Then Cancel = -1
If picPost.Visible = True Then Cancel = -1
If picAccDistribution.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lblTotalCredit_Change()
lblBalance.Caption = Format(RETURNLABELVALUE(lblTotalDebit) - RETURNLABELVALUE(lblTotalCredit), "#,##0.00")
End Sub

Private Sub lblTotalDebit_Change()
lblBalance.Caption = Format(RETURNLABELVALUE(lblTotalDebit) - RETURNLABELVALUE(lblTotalCredit), "#,##0.00")
End Sub

Private Sub lstAccDistribution_GotFocus()
is_DET_FOCUS = 1
ROW = lstAccDistribution.SelectedItem.Index
TRANS_DETAIL = is_DET_REFRESH
End Sub

Private Sub lstAccDistribution_ItemClick(ByVal Item As MSComctlLib.ListItem)
ROW = lstAccDistribution.SelectedItem.Index
End Sub

Private Sub lstAccDistribution_LostFocus()
is_DET_FOCUS = 0
End Sub

Private Sub lstDetail_GotFocus()
If TRANSACTIONTYPE = is_REFRESH Then
    If Statusbar1.Panels(1) <> "" Then
        If AccessRights("Receiving Report", "Edit") = False Then
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
            'Me.Caption = "RECEIVING REPORT - EDIT"
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

Private Sub lstDetail_LostFocus()
is_DET_FOCUS = 0
End Sub



Private Sub lstGLAutoVAT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAutoVAT_Click
End Sub

Private Sub lstResultGLAccount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKGLAccount_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "PONumberRR", "PONumRR", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "PONumberRR", "PONumRR", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "PONumberRR", "PONumRR", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "PONumberRR", "PONumRR", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Post":    PRESS_F8
    Case "Accnt":   PRESS_F7
    Case "Print":   PRESS_F9
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtAccountName_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstAccDistribution.ListItems
        .Item(ROW).SubItems(2) = Trim(txtAccountName.Text)
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
        .Item(ROW).SubItems(1) = Trim(txtAccountNo.Text)
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

Private Sub txtCost_Change()
If TRANS_DETAIL = is_DET_EDITTING Then
    lstDetail.ListItems.Item(ROW).SubItems(8) = Format(RETURNTEXTVALUE(txtCost), "#,##0.00")
    txtTotalCost.Text = Format(RETURNTEXTVALUE(txtRecd) * RETURNTEXTVALUE(txtCost), "#,##0.00")
    If Trim(txtDisc.Text) <> "" Then
        txtNetCost.Text = RETURNTEXTVALUE(txtCost) * (100 - CDbl(Val(Trim(txtDisc.Text)))) 'Format(RETURNTEXTVALUE(txtCost) * (100 - CDbl(Val(Trim(txtDisc.Text)))) / 100, "#,##0.000")
    Else
        txtNetCost.Text = RETURNTEXTVALUE(txtCost) 'Format(RETURNTEXTVALUE(txtCost), "#,##0.000")
    End If
End If
End Sub

Private Sub txtCost_GotFocus()
HTEXT txtCost
End Sub

Private Sub txtCost_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtNetCost.SetFocus
    
'    picBody.Enabled = True
'    picSLine.Visible = False
'    TRANS_DETAIL = is_DET_REFRESH
'    TOOLBARFUNC 5
'    lstDetail.SetFocus
'End If
End Sub

Private Sub txtCredit_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstAccDistribution.ListItems
        .Item(ROW).SubItems(4) = IIf(RETURNTEXTVALUE(txtCredit) = 0, " ", Format(RETURNTEXTVALUE(txtCredit), "#,##0.00"))
        .Item(ROW).SubItems(5) = Format(RETURNTEXTVALUE(txtDebit) - RETURNTEXTVALUE(txtCredit), "#,##0.00")
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
        .Item(ROW).SubItems(3) = IIf(RETURNTEXTVALUE(txtDebit) = 0, " ", Format(RETURNTEXTVALUE(txtDebit), "#,##0.00"))
        .Item(ROW).SubItems(5) = Format(RETURNTEXTVALUE(txtDebit) - RETURNTEXTVALUE(txtCredit), "#,##0.00")
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

Private Sub txtGLAddAutoVAT_Change()
If Trim(txtGLAddAutoVAT.Text) = "" Then lstGLAutoVAT.Clear: Exit Sub
lstGLAutoVAT.Clear
s = "SELECT tbl_GL_Accounts.* " & _
    " FROM tbl_GL_Accounts " & _
    " WHERE (AccountName LIKE '" & FORMATSQL(Trim(txtGLAddAutoVAT.Text)) & "%') " & _
    " ORDER BY AccountName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstGLAutoVAT.AddItem rs!AccountCode & " : " & rs!AccountName
    rs.MoveNext
Wend
rs.Close
If lstGLAutoVAT.ListCount Then lstGLAutoVAT.ListIndex = 0
End Sub

Private Sub txtGLAddAutoVAT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstGLAutoVAT.SetFocus
End Sub

Private Sub txtInvDate_GotFocus()
HTEXT txtInvDate
End Sub

Private Sub txtInvDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtInvGross.SetFocus
End Sub

Private Sub txtInvGross_Change()
'If TRANSACTIONTYPE = is_ADDING Or _
'TRANSACTIONTYPE = is_EDITTING Then
'    If RETURNTEXTVALUE(txtInvNet) = 0 Then
'        txtInvNet.Text = Format(RETURNTEXTVALUE(txtInvGross), "#,##0.00")
'    End If
'End If
End Sub

Private Sub txtInvGross_GotFocus()
HTEXT txtInvGross
End Sub

Private Sub txtInvGross_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtInvNet.SetFocus
End Sub

Private Sub txtInvGross_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtInvGross_LostFocus()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If RETURNTEXTVALUE(txtInvNet) = 0 Then
        txtInvNet.Text = Format(RETURNTEXTVALUE(txtInvGross), "#,##0.00")
    End If
End If
End Sub

Private Sub txtInvNet_Change()
'If TRANSACTIONTYPE = is_ADDING Or _
'TRANSACTIONTYPE = is_EDITTING Then
'    If RETURNTEXTVALUE(txtInvGross) = 0 Then
'        txtInvGross.Text = Format(RETURNTEXTVALUE(txtInvNet), "#,##0.00")
'    End If
'End If
End Sub

Private Sub txtInvNet_GotFocus()
HTEXT txtInvNet
End Sub

Private Sub txtInvNet_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtStockClerk.SetFocus
End Sub

Private Sub txtInvNet_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtInvNet_LostFocus()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If RETURNTEXTVALUE(txtInvGross) = 0 Then
        txtInvGross.Text = Format(RETURNTEXTVALUE(txtInvNet), "#,##0.00")
    End If
End If
End Sub

Private Sub txtInvNumber_GotFocus()
HTEXT txtPONumber
End Sub

Private Sub txtInvNumber_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtInvDate.SetFocus
End Sub

Private Sub txtNetCost_Change()
If TRANS_DETAIL = is_DET_EDITTING Then
    lstDetail.ListItems.Item(ROW).SubItems(9) = Format(RETURNTEXTVALUE(txtCost), "#,##0.00")
    txtTotalNetCost.Text = Format(RETURNTEXTVALUE(txtRecd) * RETURNTEXTVALUE(txtNetCost), "#,##0.00")
End If
End Sub

Private Sub txtNetCost_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtSLRemarks.SetFocus
End Sub

Private Sub txtPONumAdd_GotFocus()
HTEXT txtPONumAdd
End Sub

Private Sub txtPONumAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtRRNumAdd.SetFocus
End Sub

Private Sub txtPONumAdd_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtPONumAdd_LostFocus()
txtPONumAdd.Text = Format(RETURNTEXTVALUE(txtPONumAdd), "0000000#")
End Sub

Private Sub txtPONumber_GotFocus()
HTEXT txtPONumber
End Sub

Private Sub txtPONumber_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If TRANSACTIONTYPE = is_FINDING Then
        s = "SELECT tbl_Inv_RR.* " & _
            " FROM tbl_Inv_RR " & _
            " WHERE (PONumber = '" & Format(RETURNTEXTVALUE(txtPONumber), "0000000#") & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            MsgBox "'" & Format(RETURNTEXTVALUE(txtPONumber), "0000000#") & "' Not Found!                   ", vbCritical, "Error..."
            txtPONumber.SetFocus
            Exit Sub
        End If
        rs.Close
        
        LOCKTEXT True
        TOOLBARFUNC 1
        TRANSACTIONTYPE = is_REFRESH
        'Me.Caption = "RECEIVING REPORT - BROWSE"
        BROWSER Format(RETURNTEXTVALUE(txtPONumber), "0000000#"), "is_LOAD"
        
    End If
End If
End Sub

Private Sub txtPONumber_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub


Private Sub txtPurchaser_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtDeptHead.SetFocus
End Sub

Private Sub txtRecd_Change()
If TRANS_DETAIL = is_DET_EDITTING Then
    lstDetail.ListItems.Item(ROW).SubItems(4) = Format(RETURNTEXTVALUE(txtRecd), "#,##0.00")
    txtTotalCost.Text = Format(RETURNTEXTVALUE(txtRecd) * RETURNTEXTVALUE(txtCost), "#,##0.00")
    txtTotalNetCost.Text = Format(RETURNTEXTVALUE(txtRecd) * RETURNTEXTVALUE(txtNetCost), "#,##0.00")
End If
End Sub

Private Sub txtRecd_GotFocus()
HTEXT txtRecd
End Sub

Private Sub txtRecd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtCost.SetFocus
End Sub
Private Sub txtRRDateAdd_GotFocus()
HTEXT txtRRDateAdd
End Sub

Private Sub txtRRDateAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAddRR_Click
End Sub

Private Sub txtRRDateAdd_LostFocus()
If IsDate(txtRRDateAdd.Text) Then
    txtRRDateAdd.Text = Format(FormatDateTime(txtRRDateAdd.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub

Private Sub txtRRDatePosting_GotFocus()
HTEXT txtRRDatePosting
End Sub

Private Sub txtRRDatePosting_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtRRNoPosting.SetFocus
End Sub

Private Sub txtRRDatePosting_LostFocus()
If IsDate(txtRRDatePosting.Text) = True Then
    txtRRDatePosting.Text = Format(FormatDateTime(txtRRDatePosting.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub


Private Sub txtRRNoPosting_GotFocus()
HTEXT txtRRNoPosting
End Sub

Private Sub txtRRNoPosting_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub txtRRNoPosting_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtRRNumAdd_GotFocus()
HTEXT txtRRNumAdd
End Sub

Private Sub txtRRNumAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtRRDateAdd.SetFocus
End Sub

Private Sub txtRRNumAdd_KeyPress(KeyAscii As Integer)
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
If TRANS_DETAIL = is_DET_EDITTING Then
    lstDetail.ListItems.Item(ROW).SubItems(16) = Trim(txtSLRemarks.Text)
End If
End Sub

Private Sub txtSLRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    picBody.Enabled = True
    picSLine.Visible = False
    TRANS_DETAIL = is_DET_REFRESH
    TOOLBARFUNC 5
    lstDetail.SetFocus
End If
End Sub

Private Sub txtStockClerk_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtPurchaser.SetFocus
End Sub

Private Sub txtTotalCost_Change()
With lstDetail.ListItems
    DoEvents
    .Item(ROW).SubItems(11) = Format(RETURNTEXTVALUE(txtTotalCost), "#,##0.00")
    a = 0
    For i = 1 To .Count
        DoEvents
        a = a + IIf(IsNumeric(.Item(i).SubItems(11)), CDbl(.Item(i).SubItems(11)), 0)
    Next i
    lblTotalCost.Caption = Format(a, "#,##0.00")
End With
End Sub

Private Sub txtTotalNetCost_Change()
With lstDetail.ListItems
    DoEvents
    .Item(ROW).SubItems(10) = Format(RETURNTEXTVALUE(txtTotalNetCost), "#,##0.00")
    b = 0
    For i = 1 To .Count
        DoEvents
        b = b + IIf(IsNumeric(.Item(i).SubItems(10)), CDbl(.Item(i).SubItems(10)), 0)
    Next i
    lblTotalNetCost.Caption = Format(b, "#,##0.00")
End With
End Sub
