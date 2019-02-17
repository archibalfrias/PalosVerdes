VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{76880EFA-2CCC-4791-B35E-F6A7359CAFDD}#1.0#0"; "prjXTab.ocx"
Begin VB.Form aUserAccount 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14310
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "aUserAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   14310
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   19995
      TabIndex        =   146
      Top             =   0
      Width           =   20000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   147
         Top             =   105
         Width           =   14880
         _ExtentX        =   26247
         _ExtentY        =   1429
         ButtonWidth     =   1058
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   20
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
               Caption         =   "Reset"
               Key             =   "Reset"
               ImageKey        =   "IMG12"
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
               ImageKey        =   "IMG13"
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         MousePointer    =   99
         MouseIcon       =   "aUserAccount.frx":0CCA
         Begin VB.TextBox txtPassword01 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   9120
            Locked          =   -1  'True
            TabIndex        =   148
            Top             =   240
            Visible         =   0   'False
            Width           =   2175
         End
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   20000
         Y1              =   1005
         Y2              =   1005
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   20000
         Y1              =   910
         Y2              =   910
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   20000
         Y1              =   90
         Y2              =   90
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   13440
      Top             =   4680
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
            Picture         =   "aUserAccount.frx":0FE4
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aUserAccount.frx":1CBE
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aUserAccount.frx":2998
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aUserAccount.frx":3672
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aUserAccount.frx":434C
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aUserAccount.frx":5026
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aUserAccount.frx":5D00
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aUserAccount.frx":69DA
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aUserAccount.frx":76B4
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aUserAccount.frx":7F8E
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aUserAccount.frx":8C68
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aUserAccount.frx":9942
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aUserAccount.frx":A61C
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aUserAccount.frx":B2F6
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aUserAccount.frx":BFD0
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picAdministrator 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   7560
      ScaleHeight     =   375
      ScaleWidth      =   1815
      TabIndex        =   29
      Top             =   600
      Width           =   1815
      Begin VB.CheckBox chkAdministrator 
         Caption         =   "Administrator"
         Height          =   195
         Left            =   10
         TabIndex        =   30
         Top             =   10
         Width           =   1335
      End
   End
   Begin VB.TextBox txtUserNameFind 
      Height          =   315
      Left            =   8520
      TabIndex        =   18
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox picBody 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   7335
      Left            =   120
      ScaleHeight     =   7335
      ScaleWidth      =   14055
      TabIndex        =   3
      Top             =   1200
      Width           =   14055
      Begin VB.TextBox txtCompleteName 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   31
         Top             =   360
         Width           =   8535
      End
      Begin VB.TextBox txtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   5520
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtUserName 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   0
         Width           =   1695
      End
      Begin prjXTab.XTab XTab1 
         Height          =   6495
         Left            =   0
         TabIndex        =   4
         Top             =   840
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   11456
         TabCount        =   7
         TabCaption(0)   =   "Human Resource"
         TabContCtrlCnt(0)=   1
         Tab(0)ContCtrlCap(1)=   "picPersonnel"
         TabCaption(1)   =   "Membership"
         TabContCtrlCnt(1)=   1
         Tab(1)ContCtrlCap(1)=   "picMembership"
         TabCaption(2)   =   "Golf Operation"
         TabContCtrlCnt(2)=   1
         Tab(2)ContCtrlCap(1)=   "picGolfOperation"
         TabCaption(3)   =   "Golf Scoring"
         TabContCtrlCnt(3)=   1
         Tab(3)ContCtrlCap(1)=   "picGolfScoring"
         TabCaption(4)   =   "Food and Beverage"
         TabContCtrlCnt(4)=   1
         Tab(4)ContCtrlCap(1)=   "picFnB"
         TabCaption(5)   =   "Finance and Controllership"
         TabContCtrlCnt(5)=   1
         Tab(5)ContCtrlCap(1)=   "picFinance"
         TabCaption(6)   =   "Utility"
         TabContCtrlCnt(6)=   1
         Tab(6)ContCtrlCap(1)=   "picUtility"
         TabTheme        =   2
         ActiveTabBackStartColor=   16777215
         ActiveTabBackEndColor=   13023396
         InActiveTabBackStartColor=   16777215
         InActiveTabBackEndColor=   -2147483634
         InActiveTabForeColor=   -2147483631
         BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OuterBorderColor=   -2147483628
         BottomRightInnerBorderColor=   -2147483633
         TabStripBackColor=   16777215
         DisabledTabBackColor=   -2147483633
         DisabledTabForeColor=   -2147483627
         Begin VB.PictureBox picFnB 
            BackColor       =   &H00C6B8A4&
            BorderStyle     =   0  'None
            Height          =   5775
            Left            =   -74760
            ScaleHeight     =   5775
            ScaleWidth      =   13455
            TabIndex        =   132
            Top             =   480
            Width           =   13455
            Begin VB.Frame Frame57 
               BackColor       =   &H00C6B8A4&
               Caption         =   "LOCATION"
               ForeColor       =   &H000000FF&
               Height          =   1575
               Left            =   0
               TabIndex        =   133
               Top             =   0
               Width           =   1575
               Begin MSComctlLib.ListView lstFnBLocation 
                  Height          =   1170
                  Left            =   120
                  TabIndex        =   134
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2064
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
         End
         Begin VB.PictureBox picFinance 
            BackColor       =   &H00C6B8A4&
            BorderStyle     =   0  'None
            Height          =   6015
            Left            =   -74880
            ScaleHeight     =   6015
            ScaleWidth      =   13095
            TabIndex        =   77
            Top             =   360
            Width           =   13095
            Begin VB.Frame Frame50 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Purchase Invoice"
               ForeColor       =   &H000000FF&
               Height          =   2175
               Left            =   11520
               TabIndex        =   116
               Top             =   120
               Width           =   1575
               Begin MSComctlLib.ListView lstPI 
                  Height          =   1890
                  Left            =   120
                  TabIndex        =   117
                  Top             =   240
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   3334
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1764
                  EndProperty
               End
            End
            Begin VB.Frame Frame49 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Acknowledgement"
               ForeColor       =   &H000000FF&
               Height          =   1695
               Left            =   3120
               TabIndex        =   114
               Top             =   4320
               Width           =   1575
               Begin MSComctlLib.ListView lstAckReceipt 
                  Height          =   1290
                  Left            =   120
                  TabIndex        =   115
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2275
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame48 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Official Receipt"
               ForeColor       =   &H000000FF&
               Height          =   1695
               Left            =   0
               TabIndex        =   112
               Top             =   4320
               Width           =   1455
               Begin MSComctlLib.ListView lstOR 
                  Height          =   1290
                  Left            =   120
                  TabIndex        =   113
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2275
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame17 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Item's Information"
               ForeColor       =   &H000000FF&
               Height          =   2175
               Left            =   5040
               TabIndex        =   110
               Top             =   120
               Width           =   1575
               Begin MSComctlLib.ListView lstItemInfo 
                  Height          =   1050
                  Left            =   120
                  TabIndex        =   111
                  Top             =   240
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   1852
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame19 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Supplier"
               ForeColor       =   &H000000FF&
               Height          =   2175
               Left            =   3360
               TabIndex        =   108
               Top             =   120
               Width           =   1575
               Begin MSComctlLib.ListView lstSupplier 
                  Height          =   1050
                  Left            =   120
                  TabIndex        =   109
                  Top             =   240
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   1852
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame20 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Classification"
               ForeColor       =   &H000000FF&
               Height          =   2175
               Left            =   1680
               TabIndex        =   106
               Top             =   120
               Width           =   1575
               Begin MSComctlLib.ListView lstClassification 
                  Height          =   1050
                  Left            =   120
                  TabIndex        =   107
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   1852
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame21 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Section"
               ForeColor       =   &H000000FF&
               Height          =   2175
               Left            =   0
               TabIndex        =   104
               Top             =   120
               Width           =   1575
               Begin MSComctlLib.ListView lstSection 
                  Height          =   1050
                  Left            =   120
                  TabIndex        =   105
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   1852
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame31 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Purchase Order"
               ForeColor       =   &H000000FF&
               Height          =   2175
               Left            =   8280
               TabIndex        =   102
               Top             =   120
               Width           =   1455
               Begin MSComctlLib.ListView lstPO 
                  Height          =   1530
                  Left            =   120
                  TabIndex        =   103
                  Top             =   240
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   2699
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame32 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Menu Management"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   4800
               TabIndex        =   100
               Top             =   2400
               Width           =   1575
               Begin MSComctlLib.ListView lstMenuMngt 
                  Height          =   1530
                  Left            =   120
                  TabIndex        =   101
                  Top             =   240
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   2699
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame33 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Fixed Asset"
               ForeColor       =   &H000000FF&
               Height          =   2175
               Left            =   6720
               TabIndex        =   98
               Top             =   120
               Width           =   1455
               Begin MSComctlLib.ListView lstFixedAsset 
                  Height          =   1050
                  Left            =   120
                  TabIndex        =   99
                  Top             =   240
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   1852
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame34 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Receiving Report"
               ForeColor       =   &H000000FF&
               Height          =   2175
               Left            =   9840
               TabIndex        =   96
               Top             =   120
               Width           =   1575
               Begin MSComctlLib.ListView lstRR 
                  Height          =   1890
                  Left            =   120
                  TabIndex        =   97
                  Top             =   240
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   3334
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1764
                  EndProperty
               End
            End
            Begin VB.Frame Frame35 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Stock Transfer"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   0
               TabIndex        =   94
               Top             =   2400
               Width           =   1455
               Begin MSComctlLib.ListView lstST 
                  Height          =   1530
                  Left            =   120
                  TabIndex        =   95
                  Top             =   240
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   2699
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame36 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Stock Adjustment"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   1560
               TabIndex        =   92
               Top             =   2400
               Width           =   1455
               Begin MSComctlLib.ListView lstSA 
                  Height          =   1530
                  Left            =   120
                  TabIndex        =   93
                  Top             =   240
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   2699
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame37 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Stock Issuance"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   3120
               TabIndex        =   90
               Top             =   2400
               Width           =   1575
               Begin MSComctlLib.ListView lstSI 
                  Height          =   1530
                  Left            =   120
                  TabIndex        =   91
                  Top             =   240
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   2699
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1764
                  EndProperty
               End
            End
            Begin VB.Frame Frame47 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Chart of Accounts"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   6480
               TabIndex        =   88
               Top             =   2400
               Width           =   1935
               Begin MSComctlLib.ListView lstChartofAccounts 
                  Height          =   1290
                  Left            =   120
                  TabIndex        =   89
                  Top             =   240
                  Width           =   1755
                  _ExtentX        =   3096
                  _ExtentY        =   2275
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   2470
                  EndProperty
               End
            End
            Begin VB.Frame Frame41 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Charge Invoice"
               ForeColor       =   &H000000FF&
               Height          =   1695
               Left            =   1560
               TabIndex        =   86
               Top             =   4320
               Width           =   1455
               Begin MSComctlLib.ListView lstChargeInvoice 
                  Height          =   1290
                  Left            =   120
                  TabIndex        =   87
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2275
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame40 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Petty Cash"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   11520
               TabIndex        =   84
               Top             =   4200
               Visible         =   0   'False
               Width           =   1575
               Begin MSComctlLib.ListView lstPettyCash 
                  Height          =   1530
                  Left            =   120
                  TabIndex        =   85
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2699
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame43 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Check Voucher"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   8520
               TabIndex        =   82
               Top             =   2400
               Width           =   1335
               Begin MSComctlLib.ListView lstCheckVoucher 
                  Height          =   1290
                  Left            =   120
                  TabIndex        =   83
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2275
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame42 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Debit/Credit Memo"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   11520
               TabIndex        =   80
               Top             =   2400
               Width           =   1575
               Begin MSComctlLib.ListView lstDRCRMemo 
                  Height          =   1530
                  Left            =   120
                  TabIndex        =   81
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2699
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame38 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Journal Voucher"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   9960
               TabIndex        =   78
               Top             =   2400
               Width           =   1455
               Begin MSComctlLib.ListView lstJournalVoucher 
                  Height          =   1290
                  Left            =   120
                  TabIndex        =   79
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2275
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
         End
         Begin VB.PictureBox picUtility 
            BackColor       =   &H00C6B8A4&
            BorderStyle     =   0  'None
            Height          =   3975
            Left            =   -74880
            ScaleHeight     =   3975
            ScaleWidth      =   9855
            TabIndex        =   68
            Top             =   360
            Width           =   9855
            Begin VB.Frame Frame8 
               BackColor       =   &H00C6B8A4&
               Caption         =   "DEPARTMENT"
               ForeColor       =   &H000000FF&
               Height          =   1695
               Left            =   3120
               TabIndex        =   75
               Top             =   120
               Width           =   1455
               Begin MSComctlLib.ListView lstDept 
                  Height          =   1050
                  Left            =   120
                  TabIndex        =   76
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   1852
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame30 
               BackColor       =   &H00C6B8A4&
               Caption         =   "BACK UP DATABASE"
               ForeColor       =   &H000000FF&
               Height          =   1695
               Left            =   4680
               TabIndex        =   73
               Top             =   120
               Width           =   1695
               Begin MSComctlLib.ListView lstBackUp 
                  Height          =   1410
                  Left            =   120
                  TabIndex        =   74
                  Top             =   240
                  Width           =   1515
                  _ExtentX        =   2672
                  _ExtentY        =   2487
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1764
                  EndProperty
               End
            End
            Begin VB.Frame Frame18 
               BackColor       =   &H00C6B8A4&
               Caption         =   "USER ACCOUNT"
               ForeColor       =   &H000000FF&
               Height          =   1695
               Left            =   0
               TabIndex        =   71
               Top             =   120
               Width           =   1455
               Begin MSComctlLib.ListView lstUserRights 
                  Height          =   1170
                  Left            =   120
                  TabIndex        =   72
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2064
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame5 
               BackColor       =   &H00C6B8A4&
               Caption         =   "COMPANY INFO"
               ForeColor       =   &H000000FF&
               Height          =   1695
               Left            =   1560
               TabIndex        =   69
               Top             =   120
               Width           =   1455
               Begin MSComctlLib.ListView lstCompany 
                  Height          =   1170
                  Left            =   120
                  TabIndex        =   70
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2064
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
         End
         Begin VB.PictureBox picGolfScoring 
            BackColor       =   &H00C6B8A4&
            BorderStyle     =   0  'None
            Height          =   5415
            Left            =   -74880
            ScaleHeight     =   5415
            ScaleWidth      =   12975
            TabIndex        =   59
            Top             =   480
            Width           =   12975
            Begin VB.Frame Frame12 
               BackColor       =   &H00C6B8A4&
               Caption         =   "SCORE CARD"
               ForeColor       =   &H000000FF&
               Height          =   1575
               Left            =   5280
               TabIndex        =   66
               Top             =   0
               Width           =   1575
               Begin MSComctlLib.ListView lstScoreCard 
                  Height          =   1170
                  Left            =   120
                  TabIndex        =   67
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2064
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame11 
               BackColor       =   &H00C6B8A4&
               Caption         =   "TEAM SETUP"
               ForeColor       =   &H000000FF&
               Height          =   1575
               Left            =   3600
               TabIndex        =   64
               Top             =   0
               Width           =   1575
               Begin MSComctlLib.ListView lstTeam 
                  Height          =   1170
                  Left            =   120
                  TabIndex        =   65
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2064
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame9 
               BackColor       =   &H00C6B8A4&
               Caption         =   "PLAYER SETUP"
               ForeColor       =   &H000000FF&
               Height          =   1575
               Left            =   1920
               TabIndex        =   62
               Top             =   0
               Width           =   1575
               Begin MSComctlLib.ListView lstPlayer 
                  Height          =   1170
                  Left            =   120
                  TabIndex        =   63
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2064
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame6 
               BackColor       =   &H00C6B8A4&
               Caption         =   "TOURNAMENT SETUP"
               ForeColor       =   &H000000FF&
               Height          =   1575
               Left            =   0
               TabIndex        =   60
               Top             =   0
               Width           =   1815
               Begin MSComctlLib.ListView lstTournament 
                  Height          =   1170
                  Left            =   120
                  TabIndex        =   61
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2064
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
         End
         Begin VB.PictureBox picMembership 
            BackColor       =   &H00C6B8A4&
            BorderStyle     =   0  'None
            Height          =   6015
            Left            =   -74880
            ScaleHeight     =   6015
            ScaleWidth      =   12975
            TabIndex        =   24
            Top             =   360
            Width           =   12975
            Begin VB.Frame Frame24 
               BackColor       =   &H00C6B8A4&
               Caption         =   "SHARE's ID"
               ForeColor       =   &H000000FF&
               Height          =   1455
               Left            =   4680
               TabIndex        =   47
               Top             =   120
               Width           =   1455
               Begin MSComctlLib.ListView lstShareID 
                  Height          =   1170
                  Left            =   120
                  TabIndex        =   48
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2064
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame29 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Member Action"
               ForeColor       =   &H000000FF&
               Height          =   1455
               Left            =   3120
               TabIndex        =   49
               Top             =   120
               Width           =   1455
               Begin MSComctlLib.ListView lstMemberAction 
                  Height          =   1050
                  Left            =   120
                  TabIndex        =   50
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   1852
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame26 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Corporate/Company Accnt"
               ForeColor       =   &H000000FF&
               Height          =   1455
               Left            =   6240
               TabIndex        =   51
               Top             =   120
               Width           =   2175
               Begin MSComctlLib.ListView lstCorporateAccnt 
                  Height          =   1050
                  Left            =   120
                  TabIndex        =   52
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   1852
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame25 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Member ID No."
               ForeColor       =   &H000000FF&
               Height          =   1455
               Left            =   1560
               TabIndex        =   53
               Top             =   120
               Width           =   1455
               Begin MSComctlLib.ListView lstMemberID 
                  Height          =   1050
                  Left            =   120
                  TabIndex        =   54
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   1852
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame23 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Member Info"
               ForeColor       =   &H000000FF&
               Height          =   1455
               Left            =   0
               TabIndex        =   55
               Top             =   120
               Width           =   1455
               Begin MSComctlLib.ListView lstMemberInfo 
                  Height          =   1050
                  Left            =   120
                  TabIndex        =   56
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   1852
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
         End
         Begin VB.PictureBox picGolfOperation 
            BackColor       =   &H00C6B8A4&
            BorderStyle     =   0  'None
            Height          =   6015
            Left            =   -74880
            ScaleHeight     =   6015
            ScaleWidth      =   13695
            TabIndex        =   23
            Top             =   360
            Width           =   13695
            Begin VB.Frame Frame60 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Pro Shop (Items)"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   0
               TabIndex        =   135
               Top             =   1920
               Width           =   7335
               Begin VB.Frame Frame63 
                  BackColor       =   &H00C6B8A4&
                  Caption         =   "Item Type"
                  ForeColor       =   &H000000FF&
                  Height          =   1455
                  Left            =   5880
                  TabIndex        =   144
                  Top             =   240
                  Width           =   1335
                  Begin MSComctlLib.ListView lstProShopItemsItemType 
                     Height          =   1170
                     Left            =   120
                     TabIndex        =   145
                     Top             =   240
                     Width           =   1155
                     _ExtentX        =   2037
                     _ExtentY        =   2064
                     View            =   3
                     LabelEdit       =   1
                     LabelWrap       =   0   'False
                     HideSelection   =   -1  'True
                     HideColumnHeaders=   -1  'True
                     Checkboxes      =   -1  'True
                     _Version        =   393217
                     ForeColor       =   -2147483640
                     BackColor       =   13023396
                     Appearance      =   0
                     NumItems        =   2
                     BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        Text            =   "0"
                        Object.Width           =   529
                     EndProperty
                     BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        SubItemIndex    =   1
                        Object.Width           =   1411
                     EndProperty
                  End
               End
               Begin VB.Frame Frame62 
                  BackColor       =   &H00C6B8A4&
                  Caption         =   "Sizes"
                  ForeColor       =   &H000000FF&
                  Height          =   1455
                  Left            =   3000
                  TabIndex        =   142
                  Top             =   240
                  Width           =   1335
                  Begin MSComctlLib.ListView lstProShopItemsSize 
                     Height          =   1170
                     Left            =   120
                     TabIndex        =   143
                     Top             =   240
                     Width           =   1155
                     _ExtentX        =   2037
                     _ExtentY        =   2064
                     View            =   3
                     LabelEdit       =   1
                     LabelWrap       =   0   'False
                     HideSelection   =   -1  'True
                     HideColumnHeaders=   -1  'True
                     Checkboxes      =   -1  'True
                     _Version        =   393217
                     ForeColor       =   -2147483640
                     BackColor       =   13023396
                     Appearance      =   0
                     NumItems        =   2
                     BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        Text            =   "0"
                        Object.Width           =   529
                     EndProperty
                     BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        SubItemIndex    =   1
                        Object.Width           =   1411
                     EndProperty
                  End
               End
               Begin VB.Frame Frame61 
                  BackColor       =   &H00C6B8A4&
                  Caption         =   "Color"
                  ForeColor       =   &H000000FF&
                  Height          =   1455
                  Left            =   4440
                  TabIndex        =   140
                  Top             =   240
                  Width           =   1335
                  Begin MSComctlLib.ListView lstProShopItemsColor 
                     Height          =   1170
                     Left            =   120
                     TabIndex        =   141
                     Top             =   240
                     Width           =   1155
                     _ExtentX        =   2037
                     _ExtentY        =   2064
                     View            =   3
                     LabelEdit       =   1
                     LabelWrap       =   0   'False
                     HideSelection   =   -1  'True
                     HideColumnHeaders=   -1  'True
                     Checkboxes      =   -1  'True
                     _Version        =   393217
                     ForeColor       =   -2147483640
                     BackColor       =   13023396
                     Appearance      =   0
                     NumItems        =   2
                     BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        Text            =   "0"
                        Object.Width           =   529
                     EndProperty
                     BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        SubItemIndex    =   1
                        Object.Width           =   1411
                     EndProperty
                  End
               End
               Begin VB.Frame Frame58 
                  BackColor       =   &H00C6B8A4&
                  Caption         =   "Brand"
                  ForeColor       =   &H000000FF&
                  Height          =   1455
                  Left            =   120
                  TabIndex        =   138
                  Top             =   240
                  Width           =   1335
                  Begin MSComctlLib.ListView lstProShopItemsBrand 
                     Height          =   1170
                     Left            =   120
                     TabIndex        =   139
                     Top             =   240
                     Width           =   1155
                     _ExtentX        =   2037
                     _ExtentY        =   2064
                     View            =   3
                     LabelEdit       =   1
                     LabelWrap       =   0   'False
                     HideSelection   =   -1  'True
                     HideColumnHeaders=   -1  'True
                     Checkboxes      =   -1  'True
                     _Version        =   393217
                     ForeColor       =   -2147483640
                     BackColor       =   13023396
                     Appearance      =   0
                     NumItems        =   2
                     BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        Text            =   "0"
                        Object.Width           =   529
                     EndProperty
                     BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        SubItemIndex    =   1
                        Object.Width           =   1411
                     EndProperty
                  End
               End
               Begin VB.Frame Frame59 
                  BackColor       =   &H00C6B8A4&
                  Caption         =   "Model"
                  ForeColor       =   &H000000FF&
                  Height          =   1455
                  Left            =   1560
                  TabIndex        =   136
                  Top             =   240
                  Width           =   1335
                  Begin MSComctlLib.ListView lstProShopItemsModel 
                     Height          =   1170
                     Left            =   120
                     TabIndex        =   137
                     Top             =   240
                     Width           =   1155
                     _ExtentX        =   2037
                     _ExtentY        =   2064
                     View            =   3
                     LabelEdit       =   1
                     LabelWrap       =   0   'False
                     HideSelection   =   -1  'True
                     HideColumnHeaders=   -1  'True
                     Checkboxes      =   -1  'True
                     _Version        =   393217
                     ForeColor       =   -2147483640
                     BackColor       =   13023396
                     Appearance      =   0
                     NumItems        =   2
                     BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        Text            =   "0"
                        Object.Width           =   529
                     EndProperty
                     BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        SubItemIndex    =   1
                        Object.Width           =   1411
                     EndProperty
                  End
               End
            End
            Begin VB.Frame Frame56 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Pro Shop (Items)"
               ForeColor       =   &H000000FF&
               Height          =   1695
               Left            =   7560
               TabIndex        =   130
               Top             =   120
               Width           =   1455
               Begin MSComctlLib.ListView lstProShopItems 
                  Height          =   1290
                  Left            =   120
                  TabIndex        =   131
                  Top             =   240
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   2275
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1587
                  EndProperty
               End
            End
            Begin VB.Frame Frame55 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Golf Cart"
               ForeColor       =   &H000000FF&
               Height          =   1695
               Left            =   12240
               TabIndex        =   128
               Top             =   120
               Width           =   1455
               Begin MSComctlLib.ListView lstGolfCartOP 
                  Height          =   1290
                  Left            =   120
                  TabIndex        =   129
                  Top             =   240
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   2275
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1587
                  EndProperty
               End
            End
            Begin VB.Frame Frame54 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Locker Room"
               ForeColor       =   &H000000FF&
               Height          =   1695
               Left            =   10680
               TabIndex        =   126
               Top             =   120
               Width           =   1455
               Begin MSComctlLib.ListView lstLockerRoom 
                  Height          =   1290
                  Left            =   120
                  TabIndex        =   127
                  Top             =   240
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   2275
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1587
                  EndProperty
               End
            End
            Begin VB.Frame Frame53 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Driving Range"
               ForeColor       =   &H000000FF&
               Height          =   1695
               Left            =   9120
               TabIndex        =   124
               Top             =   120
               Width           =   1455
               Begin MSComctlLib.ListView lstDrivingRange 
                  Height          =   1290
                  Left            =   120
                  TabIndex        =   125
                  Top             =   240
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   2275
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1587
                  EndProperty
               End
            End
            Begin VB.Frame Frame52 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Pro Shop"
               ForeColor       =   &H000000FF&
               Height          =   1695
               Left            =   6000
               TabIndex        =   122
               Top             =   120
               Width           =   1455
               Begin MSComctlLib.ListView lstProShop 
                  Height          =   1290
                  Left            =   120
                  TabIndex        =   123
                  Top             =   240
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   2275
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1587
                  EndProperty
               End
            End
            Begin VB.Frame Frame51 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Registration"
               ForeColor       =   &H000000FF&
               Height          =   1695
               Left            =   4440
               TabIndex        =   120
               Top             =   120
               Width           =   1455
               Begin MSComctlLib.ListView lstRegistration 
                  Height          =   1290
                  Left            =   120
                  TabIndex        =   121
                  Top             =   240
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   2275
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1587
                  EndProperty
               End
            End
            Begin VB.Frame Frame44 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Bag Drop"
               ForeColor       =   &H000000FF&
               Height          =   1695
               Left            =   2880
               TabIndex        =   118
               Top             =   120
               Width           =   1455
               Begin MSComctlLib.ListView lstBagDrop 
                  Height          =   1290
                  Left            =   120
                  TabIndex        =   119
                  Top             =   240
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   2275
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1587
                  EndProperty
               End
            End
            Begin VB.Frame Frame28 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Caddy Info"
               ForeColor       =   &H000000FF&
               Height          =   1695
               Left            =   1440
               TabIndex        =   57
               Top             =   120
               Width           =   1335
               Begin MSComctlLib.ListView lstCaddyInfo 
                  Height          =   1170
                  Left            =   120
                  TabIndex        =   58
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2064
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame27 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Golf Cart Info"
               ForeColor       =   &H000000FF&
               Height          =   1695
               Left            =   0
               TabIndex        =   27
               Top             =   120
               Width           =   1335
               Begin MSComctlLib.ListView lstGolfCart 
                  Height          =   1170
                  Left            =   120
                  TabIndex        =   28
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2064
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
         End
         Begin VB.PictureBox picPersonnel 
            BackColor       =   &H00C6B8A4&
            BorderStyle     =   0  'None
            Height          =   6015
            Left            =   120
            ScaleHeight     =   6015
            ScaleWidth      =   13095
            TabIndex        =   7
            Top             =   360
            Width           =   13095
            Begin VB.Frame Frame68 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Manual Payment"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   7200
               TabIndex        =   157
               Top             =   2040
               Width           =   1455
               Begin MSComctlLib.ListView lstManualPayment 
                  Height          =   1530
                  Left            =   120
                  TabIndex        =   158
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2699
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame67 
               BackColor       =   &H00C6B8A4&
               Caption         =   "PagIbig ADD'L Contribution"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   8880
               TabIndex        =   155
               Top             =   4200
               Width           =   2175
               Begin MSComctlLib.ListView lstPagIbigAddContri 
                  Height          =   1050
                  Left            =   120
                  TabIndex        =   156
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   1852
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame66 
               BackColor       =   &H00C6B8A4&
               Caption         =   "Perfect Days (Daily)"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   11160
               TabIndex        =   153
               Top             =   4200
               Width           =   1695
               Begin MSComctlLib.ListView lstPerfectDaysDaily 
                  Height          =   1410
                  Left            =   120
                  TabIndex        =   154
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2487
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame65 
               BackColor       =   &H00C6B8A4&
               Caption         =   "for DEDUCTION"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   5640
               TabIndex        =   151
               Top             =   2040
               Width           =   1455
               Begin MSComctlLib.ListView lstforDeduction 
                  Height          =   1530
                  Left            =   120
                  TabIndex        =   152
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2699
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame64 
               BackColor       =   &H00C6B8A4&
               Caption         =   "HOURS"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   2640
               TabIndex        =   149
               Top             =   2040
               Width           =   1455
               Begin MSComctlLib.ListView lstHours 
                  Height          =   1530
                  Left            =   120
                  TabIndex        =   150
                  Top             =   240
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   2699
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame15 
               BackColor       =   &H00C6B8A4&
               Caption         =   "GOV'T TABLES"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   0
               TabIndex        =   45
               Top             =   4200
               Width           =   3255
               Begin MSComctlLib.ListView lstGovtTables 
                  Height          =   1290
                  Left            =   120
                  TabIndex        =   46
                  Top             =   240
                  Width           =   2415
                  _ExtentX        =   4260
                  _ExtentY        =   2275
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   3704
                  EndProperty
               End
            End
            Begin VB.Frame Frame13 
               BackColor       =   &H00C6B8A4&
               Caption         =   "SRVC CHRG SETUP"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   7200
               TabIndex        =   43
               Top             =   4200
               Width           =   1575
               Begin MSComctlLib.ListView lstServiceChargeSetup 
                  Height          =   1050
                  Left            =   120
                  TabIndex        =   44
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   1852
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame7 
               BackColor       =   &H00C6B8A4&
               Caption         =   "EMP. STATUS"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   11640
               TabIndex        =   41
               Top             =   120
               Width           =   1455
               Begin MSComctlLib.ListView lstEmpStatus 
                  Height          =   1050
                  Left            =   120
                  TabIndex        =   42
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   1852
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame10 
               BackColor       =   &H00C6B8A4&
               Caption         =   "POSITION"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   9960
               TabIndex        =   39
               Top             =   120
               Width           =   1575
               Begin MSComctlLib.ListView lstPosition 
                  Height          =   1050
                  Left            =   120
                  TabIndex        =   40
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   1852
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame46 
               BackColor       =   &H00C6B8A4&
               Caption         =   "MORTUARY"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   7080
               TabIndex        =   37
               Top             =   120
               Width           =   1335
               Begin MSComctlLib.ListView lstMortuary 
                  Height          =   1530
                  Left            =   120
                  TabIndex        =   38
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2699
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame frameAllowance 
               BackColor       =   &H00C6B8A4&
               Caption         =   "ALLOWANCE"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   5400
               TabIndex        =   35
               Top             =   120
               Width           =   1575
               Begin MSComctlLib.ListView lstAllowance 
                  Height          =   1410
                  Left            =   120
                  TabIndex        =   36
                  Top             =   240
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   2487
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1764
                  EndProperty
               End
            End
            Begin VB.Frame Frame45 
               BackColor       =   &H00C6B8A4&
               Caption         =   "DEDUCTION"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   4200
               TabIndex        =   33
               Top             =   2040
               Width           =   1335
               Begin MSComctlLib.ListView lstDeduction 
                  Height          =   1530
                  Left            =   120
                  TabIndex        =   34
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2699
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame22 
               BackColor       =   &H00C6B8A4&
               Caption         =   "SERVICE CHARGE SUMM"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   5040
               TabIndex        =   25
               Top             =   4200
               Width           =   2055
               Begin MSComctlLib.ListView lstServiceChargeSumm 
                  Height          =   1290
                  Left            =   120
                  TabIndex        =   26
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2275
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame16 
               BackColor       =   &H00C6B8A4&
               Caption         =   "ABSENT/UNDERTIME EMPLOYEE"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   0
               TabIndex        =   21
               Top             =   2040
               Width           =   2535
               Begin MSComctlLib.ListView lstAbsentUndertimeEmployee 
                  Height          =   1530
                  Left            =   120
                  TabIndex        =   22
                  Top             =   240
                  Width           =   2235
                  _ExtentX        =   3942
                  _ExtentY        =   2699
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame14 
               BackColor       =   &H00C6B8A4&
               Caption         =   "SERVICE CHARGE"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   3360
               TabIndex        =   19
               Top             =   4200
               Width           =   1575
               Begin MSComctlLib.ListView lstServiceCharge 
                  Height          =   1530
                  Left            =   120
                  TabIndex        =   20
                  Top             =   240
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   2699
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1464
                  EndProperty
               End
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00C6B8A4&
               Caption         =   "COMPENSATION"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   9120
               TabIndex        =   16
               Top             =   2040
               Width           =   2055
               Begin MSComctlLib.ListView lstCompensation 
                  Height          =   1530
                  Left            =   120
                  TabIndex        =   17
                  Top             =   240
                  Width           =   1845
                  _ExtentX        =   3254
                  _ExtentY        =   2699
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   2646
                  EndProperty
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00C6B8A4&
               Caption         =   "LOANS"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   8520
               TabIndex        =   14
               Top             =   120
               Width           =   1335
               Begin MSComctlLib.ListView lstLoans 
                  Height          =   1530
                  Left            =   120
                  TabIndex        =   15
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2699
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00C6B8A4&
               Caption         =   "PERSONNEL INFO"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   0
               TabIndex        =   12
               Top             =   120
               Width           =   1575
               Begin MSComctlLib.ListView lstPersonnalInfo 
                  Height          =   1170
                  Left            =   120
                  TabIndex        =   13
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2064
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00C6B8A4&
               Caption         =   "ACTION MEMO"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   3360
               TabIndex        =   10
               Top             =   120
               Width           =   1935
               Begin MSComctlLib.ListView lstActionMemo 
                  Height          =   1530
                  Left            =   120
                  TabIndex        =   11
                  Top             =   240
                  Width           =   1755
                  _ExtentX        =   3096
                  _ExtentY        =   2699
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   2470
                  EndProperty
               End
            End
            Begin VB.Frame Frame39 
               BackColor       =   &H00C6B8A4&
               Caption         =   "ASSIGN ID"
               ForeColor       =   &H000000FF&
               Height          =   1815
               Left            =   1680
               TabIndex        =   8
               Top             =   120
               Width           =   1575
               Begin MSComctlLib.ListView lstIDNumber 
                  Height          =   1170
                  Left            =   120
                  TabIndex        =   9
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   2064
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   -1  'True
                  HideColumnHeaders=   -1  'True
                  Checkboxes      =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   13023396
                  Appearance      =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "0"
                     Object.Width           =   529
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   1411
                  EndProperty
               End
            End
         End
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "COMPLETE NAME"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD"
         Height          =   255
         Left            =   4560
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "USERNAME"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   855
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   8700
      Width           =   14310
      _ExtentX        =   25241
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   19756
            MinWidth        =   19756
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3616
            MinWidth        =   3616
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
End
Attribute VB_Name = "aUserAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strPersonnelInfo, strPersonnelID, strPersonnelAction, strPersonnelDept, strPersonnelStatus, strPersonnelPost, _
strPersonnelGovt, strPersonnelLoans, strPersonnelCompensation, strUserAccount, strCompany, strTournamentInfo, strPlayerInfo, _
strTeamInfo, strScoreCard, strServiceChargeSetup, strServiceCharge, strAbsentUndertimeEmployee, strSuspended, strLeave, strSections, _
strClassification, strSupplier, strItemInfo, strServiceChargeSumm, strMemberInfo, strShareID, strMemberIDNumber, strCorporateAccount, _
strGolfCart, strCaddyInfo, strAllowance, strMemberAction, strBackup, strPurchaseOrder, strMenuMngt, strFixedAsset, strReceivingReport, _
strStockTransfer, strStockAdjustment, strStockIssuance, strCheckVoucher, strPersonnelDeduction, strMortuary, strChartOfAccounts, strJournalVoucher, _
strOfficialReceipt, strChargeInvoice, strAcknowledgementReceipt, strDebitCreditMemo, strPurchaseInvoice, strBagDrop, strRegistration, _
strProShop, strDrivingRange, strLockerRoom, strGolfCartOP, strFnBLocation, strProShopItemsBrand, strProShopItemsModel, _
strProShopItemsSizes, strProShopItemsColor, strProShopItemsItemType, strProShopItems, strPersonnelHours, strPersonnelForDeduction, _
strPersonnelSetUpPerfectDays, strPersonnelPagIbigAddContri, strPersonnelManualPayment

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2
Const is_FINDING = 3

Dim ShiftDown, AltDown, CtrlDown
Const vbShiftMask = 1
Const vbCtrlMask = 2
Const vbAltMask = 4

Dim i, x, ArrAccount, strUserName


Private Function BROWSER(strUser, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If strUser <> "" Then
            s = "SELECT TOP 1 tbl_Users_Account.* " & _
                " FROM tbl_Users_Account " & _
                " WHERE (UserName = '" & FORMATSQL(CStr(strUser)) & "') " & _
                " ORDER BY UserName"
        Else
            s = "SELECT TOP 1 tbl_Users_Account.* " & _
                " FROM tbl_Users_Account " & _
                " ORDER BY UserName"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Users_Account.* " & _
            " FROM tbl_Users_Account " & _
            " ORDER BY UserName"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Users_Account.* " & _
            " FROM tbl_Users_Account " & _
            " WHERE (UserName < '" & FORMATSQL(CStr(strUser)) & "') " & _
            " ORDER BY UserName DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Users_Account.* " & _
            " FROM tbl_Users_Account " & _
            " WHERE (UserName > '" & FORMATSQL(CStr(strUser)) & "') " & _
            " ORDER BY UserName "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Users_Account.* " & _
            " FROM tbl_Users_Account " & _
            " ORDER BY UserName DESC"
    Case Else: Exit Function
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtUserName.Text = rs!UserName
    txtPassword.Text = EncryptDecrypt(CStr(rs!Password))
    txtPassword01.Text = EncryptDecrypt(CStr(rs!Password))
    txtCompleteName.Text = rs!CompleteName
    chkAdministrator.Value = rs!Admin
    iAdmin = rs!Admin
    
    ArrAccount = Split(rs!PersonnelInfo, "/", -1, 1)
    For i = 1 To lstPersonnalInfo.ListItems.Count
        lstPersonnalInfo.ListItems.Item(i).Checked = IIf(CDbl(ArrAccount(i - 1)) = 1, True, False)
    Next i
    ArrAccount = Split(rs!PersonnelID, "/", -1, 1)
    For i = 1 To lstIDNumber.ListItems.Count
        lstIDNumber.ListItems.Item(i).Checked = IIf(CDbl(ArrAccount(i - 1)) = 1, True, False)
    Next i
    ArrAccount = Split(rs!PersonnelAction, "/", -1, 1)
    For i = 1 To lstActionMemo.ListItems.Count
        lstActionMemo.ListItems.Item(i).Checked = IIf(CDbl(ArrAccount(i - 1)) = 1, True, False)
    Next i
    ArrAccount = Split(rs!PersonnelDept, "/", -1, 1)
    For i = 1 To lstDept.ListItems.Count
        lstDept.ListItems.Item(i).Checked = IIf(CDbl(ArrAccount(i - 1)) = 1, True, False)
    Next i
    ArrAccount = Split(rs!PersonnelStatus, "/", -1, 1)
    For i = 1 To lstEmpStatus.ListItems.Count
        lstEmpStatus.ListItems.Item(i).Checked = IIf(CDbl(ArrAccount(i - 1)) = 1, True, False)
    Next i
    ArrAccount = Split(rs!PersonnelPost, "/", -1, 1)
    For i = 1 To lstPosition.ListItems.Count
        lstPosition.ListItems.Item(i).Checked = IIf(CDbl(ArrAccount(i - 1)) = 1, True, False)
    Next i
    ArrAccount = Split(rs!PersonnelGovt, "/", -1, 1)
    For i = 1 To lstGovtTables.ListItems.Count
        lstGovtTables.ListItems.Item(i).Checked = IIf(CDbl(ArrAccount(i - 1)) = 1, True, False)
    Next i
    ArrAccount = Split(rs!PersonnelLoan, "/", -1, 1)
    For i = 1 To lstLoans.ListItems.Count
        lstLoans.ListItems.Item(i).Checked = IIf(CDbl(ArrAccount(i - 1)) = 1, True, False)
    Next i
    ArrAccount = Split(rs!PersonnelCompensation, "/", -1, 1)
    For i = 1 To lstCompensation.ListItems.Count
        lstCompensation.ListItems.Item(i).Checked = IIf(CDbl(ArrAccount(i - 1)) = 1, True, False)
    Next i
    ArrAccount = Split(rs!UserAccount, "/", -1, 1)
    For i = 1 To lstUserRights.ListItems.Count
        lstUserRights.ListItems.Item(i).Checked = IIf(CDbl(ArrAccount(i - 1)) = 1, True, False)
    Next i
    ArrAccount = Split(rs!CompanyInfo, "/", -1, 1)
    For i = 1 To lstCompany.ListItems.Count
        lstCompany.ListItems.Item(i).Checked = IIf(CDbl(ArrAccount(i - 1)) = 1, True, False)
    Next i
    ArrAccount = Split(rs!ScoringTournamentInfo, "/", -1, 1)
    For i = 1 To lstTournament.ListItems.Count
        lstTournament.ListItems.Item(i).Checked = IIf(CDbl(ArrAccount(i - 1)) = 1, True, False)
    Next i
    ArrAccount = Split(rs!ScoringPlayerInfo, "/", -1, 1)
    For i = 1 To lstPlayer.ListItems.Count
        lstPlayer.ListItems.Item(i).Checked = IIf(CDbl(ArrAccount(i - 1)) = 1, True, False)
    Next i
    ArrAccount = Split(rs!ScoringTeamInfo, "/", -1, 1)
    For i = 1 To lstTeam.ListItems.Count
        lstTeam.ListItems.Item(i).Checked = IIf(CDbl(ArrAccount(i - 1)) = 1, True, False)
    Next i
    ArrAccount = Split(rs!ScoringScoreCard, "/", -1, 1)
    For i = 1 To lstScoreCard.ListItems.Count
        lstScoreCard.ListItems.Item(i).Checked = IIf(CDbl(ArrAccount(i - 1)) = 1, True, False)
    Next i
    ArrAccount = Split(rs!ServiceChargeSetup, "/", -1, 1)
    For i = 1 To lstServiceChargeSetup.ListItems.Count
        lstServiceChargeSetup.ListItems.Item(i).Checked = IIf(CDbl(ArrAccount(i - 1)) = 1, True, False)
    Next i
    ArrAccount = Split(rs!ServiceCharge, "/", -1, 1)
    For i = 1 To lstServiceCharge.ListItems.Count
        lstServiceCharge.ListItems.Item(i).Checked = IIf(CDbl(ArrAccount(i - 1)) = 1, True, False)
    Next i
    ArrAccount = Split(rs!AbsentUndertime, "/", -1, 1)
    For i = 1 To lstAbsentUndertimeEmployee.ListItems.Count
        lstAbsentUndertimeEmployee.ListItems.Item(i).Checked = IIf(CDbl(ArrAccount(i - 1)) = 1, True, False)
    Next i
    ArrAccount = Split(rs!ServiceChargeSumm, "/", -1, 1)
    For i = 1 To lstServiceChargeSumm.ListItems.Count
        lstServiceChargeSumm.ListItems.Item(i).Checked = IIf(CDbl(ArrAccount(i - 1)) = 1, True, False)
    Next i
    ArrAccount = Split(rs!Sections, "/", -1, 1)
    For i = 1 To lstSection.ListItems.Count
        lstSection.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!Classification, "/", -1, 1)
    For i = 1 To lstClassification.ListItems.Count
        lstClassification.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!Supplier, "/", -1, 1)
    For i = 1 To lstSupplier.ListItems.Count
        lstSupplier.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!ItemInfo, "/", -1, 1)
    For i = 1 To lstItemInfo.ListItems.Count
        lstItemInfo.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!MemberInfo, "/", -1, 1)
    For i = 1 To lstMemberInfo.ListItems.Count
        lstMemberInfo.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!ShareID, "/", -1, 1)
    For i = 1 To lstShareID.ListItems.Count
        lstShareID.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!MemberIDNumber, "/", -1, 1)
    For i = 1 To lstMemberID.ListItems.Count
        lstMemberID.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!CorporateAccount, "/", -1, 1)
    For i = 1 To lstCorporateAccnt.ListItems.Count
        lstCorporateAccnt.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!GolfCart, "/", -1, 1)
    For i = 1 To lstGolfCart.ListItems.Count
        lstGolfCart.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!CaddyInfo, "/", -1, 1)
    For i = 1 To lstCaddyInfo.ListItems.Count
        lstCaddyInfo.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!Allowance, "/", -1, 1)
    For i = 1 To lstAllowance.ListItems.Count
        lstAllowance.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!MemberAction, "/", -1, 1)
    For i = 1 To lstMemberAction.ListItems.Count
        lstMemberAction.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!AllowBackup, "/", -1, 1)
    For i = 1 To lstBackUp.ListItems.Count
        lstBackUp.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!PurchaseOrder, "/", -1, 1)
    For i = 1 To lstPO.ListItems.Count
        lstPO.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!MenuMngt, "/", -1, 1)
    For i = 1 To lstMenuMngt.ListItems.Count
        lstMenuMngt.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!FixedAsset, "/", -1, 1)
    For i = 1 To lstFixedAsset.ListItems.Count
        lstFixedAsset.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!ReceivingReport, "/", -1, 1)
    For i = 1 To lstRR.ListItems.Count
        lstRR.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!StockTransfer, "/", -1, 1)
    For i = 1 To lstST.ListItems.Count
        lstST.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!StockAdjustment, "/", -1, 1)
    For i = 1 To lstSA.ListItems.Count
        lstSA.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!StockIssuance, "/", -1, 1)
    For i = 1 To lstSI.ListItems.Count
        lstSI.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!CheckVoucher, "/", -1, 1)
    For i = 1 To lstCheckVoucher.ListItems.Count
        lstCheckVoucher.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!PersonnelDeduction, "/", -1, 1)
    For i = 1 To lstDeduction.ListItems.Count
        lstDeduction.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!Mortuary, "/", -1, 1)
    For i = 1 To lstMortuary.ListItems.Count
        lstMortuary.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!ChartOfAccounts, "/", -1, 1)
    For i = 1 To lstChartofAccounts.ListItems.Count
        lstChartofAccounts.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!JournalVoucher, "/", -1, 1)
    For i = 1 To lstJournalVoucher.ListItems.Count
        lstJournalVoucher.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!OfficialReceipt, "/", -1, 1)
    For i = 1 To lstOR.ListItems.Count
        lstOR.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!ChargeInvoice, "/", -1, 1)
    For i = 1 To lstChargeInvoice.ListItems.Count
        lstChargeInvoice.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!AcknowledgementReceipt, "/", -1, 1)
    For i = 1 To lstAckReceipt.ListItems.Count
        lstAckReceipt.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!DebitCreditMemo, "/", -1, 1)
    For i = 1 To lstDRCRMemo.ListItems.Count
        lstDRCRMemo.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!PurchaseInvoice, "/", -1, 1)
    For i = 1 To lstPI.ListItems.Count
        lstPI.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!BagDrop, "/", -1, 1)
    For i = 1 To lstBagDrop.ListItems.Count
        lstBagDrop.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!Registration, "/", -1, 1)
    For i = 1 To lstRegistration.ListItems.Count
        lstRegistration.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!ProShop, "/", -1, 1)
    For i = 1 To lstProShop.ListItems.Count
        lstProShop.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!DrivingRange, "/", -1, 1)
    For i = 1 To lstDrivingRange.ListItems.Count
        lstDrivingRange.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!LockerRoom, "/", -1, 1)
    For i = 1 To lstLockerRoom.ListItems.Count
        lstLockerRoom.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!GolfCartOP, "/", -1, 1)
    For i = 1 To lstGolfCartOP.ListItems.Count
        lstGolfCartOP.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!ProShopItems, "/", -1, 1)
    For i = 1 To lstProShopItems.ListItems.Count
        lstProShopItems.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!FnBLocation, "/", -1, 1)
    For i = 1 To lstFnBLocation.ListItems.Count
        lstFnBLocation.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!ProShopItemsBrand, "/", -1, 1)
    For i = 1 To lstProShopItemsBrand.ListItems.Count
        lstProShopItemsBrand.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!ProShopItemsModel, "/", -1, 1)
    For i = 1 To lstProShopItemsModel.ListItems.Count
        lstProShopItemsModel.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!ProShopItemsSizes, "/", -1, 1)
    For i = 1 To lstProShopItemsSize.ListItems.Count
        lstProShopItemsSize.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!ProShopItemsColor, "/", -1, 1)
    For i = 1 To lstProShopItemsColor.ListItems.Count
        lstProShopItemsColor.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!ProShopItemsItemType, "/", -1, 1)
    For i = 1 To lstProShopItemsItemType.ListItems.Count
        lstProShopItemsItemType.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!PersonnelHours, "/", -1, 1)
    For i = 1 To lstHours.ListItems.Count
        lstHours.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!PersonnelForDeduction, "/", -1, 1)
    For i = 1 To lstforDeduction.ListItems.Count
        lstforDeduction.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!PersonnelSetUpPerfectDays, "/", -1, 1)
    For i = 1 To lstPerfectDaysDaily.ListItems.Count
        lstPerfectDaysDaily.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!PersonnelPagIbigAddContri, "/", -1, 1)
    For i = 1 To lstPagIbigAddContri.ListItems.Count
        lstPagIbigAddContri.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    ArrAccount = Split(rs!PersonnelManualPayment, "/", -1, 1)
    For i = 1 To lstManualPayment.ListItems.Count
        lstManualPayment.ListItems.Item(i).Checked = IIf(ArrAccount(i - 1) = 1, True, False)
    Next i
    StatusBar1.Panels(1).Text = rs!PK
    StatusBar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    
    SaveSetting App.EXEName, "UserAccountName", "UserAccntName", rs!UserName
    
End If
rs.Close
End Function

Private Function PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If AccessRights("User's Account", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
'Me.Caption = "User Account - New"
txtPassword.Text = sDefaultPW
txtUserName.SetFocus
End Function

Private Function PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If StatusBar1.Panels(1).Text = "" Then Exit Function
If AccessRights("User's Account", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_EDITTING
'Me.Caption = "User Account - Edit"
End Function

Private Function PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If StatusBar1.Panels(1).Text = "" Then Exit Function
If AccessRights("User's Account", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
If CDbl(iAdmin) = 1 Then
    If AccessRights("User's Account", "Admin") = False Then
        MsgBox "This is an Administrator Account!             " & vbCrLf & vbCrLf & "You can't Delete it.                   ", vbCritical, "Error..."
        Exit Function
    End If
End If
If MsgBox("ARE YOU SURE IN DELETING THIS ACCOUNT?                  ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Function
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_Users_Account WHERE (PK = " & StatusBar1.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "UserAccountName", "UserAccntName", ""), "is_PAGEDOWN"
If Trim(txtUserName.Text) = "" Then BROWSER GetSetting(App.EXEName, "UserAccountName", "UserAccntName", ""), "is_HOME"
Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Function
End Function

Private Function PRESS_F5()
If Trim(txtUserName.Text) = "" Then MsgBox "Please Supply Username!                   ", vbCritical, "Error...": txtUserName.SetFocus: Exit Function
If Trim(txtPassword.Text) = "" Then MsgBox "Please Supply Password!                   ", vbCritical, "Error...": txtPassword.SetFocus: Exit Function
strPersonnelInfo = ""
For i = 1 To lstPersonnalInfo.ListItems.Count
    strPersonnelInfo = strPersonnelInfo & "/" & IIf(lstPersonnalInfo.ListItems.Item(i).Checked = True, 1, 0)
Next i
strPersonnelInfo = Mid(strPersonnelInfo, 2, Len(strPersonnelInfo))
strPersonnelID = ""
For i = 1 To lstIDNumber.ListItems.Count
    strPersonnelID = strPersonnelID & "/" & IIf(lstIDNumber.ListItems.Item(i).Checked = True, 1, 0)
Next i
strPersonnelID = Mid(strPersonnelID, 2, Len(strPersonnelID))
strPersonnelAction = ""
For i = 1 To lstActionMemo.ListItems.Count
    strPersonnelAction = strPersonnelAction & "/" & IIf(lstActionMemo.ListItems.Item(i).Checked = True, 1, 0)
Next i
strPersonnelAction = Mid(strPersonnelAction, 2, Len(strPersonnelAction))
strPersonnelDept = ""
For i = 1 To lstDept.ListItems.Count
    strPersonnelDept = strPersonnelDept & "/" & IIf(lstDept.ListItems.Item(i).Checked = True, 1, 0)
Next i
strPersonnelDept = Mid(strPersonnelDept, 2, Len(strPersonnelDept))
strPersonnelStatus = ""
For i = 1 To lstEmpStatus.ListItems.Count
    strPersonnelStatus = strPersonnelStatus & "/" & IIf(lstEmpStatus.ListItems.Item(i).Checked = True, 1, 0)
Next i
strPersonnelStatus = Mid(strPersonnelStatus, 2, Len(strPersonnelStatus))
strPersonnelPost = ""
For i = 1 To lstPosition.ListItems.Count
    strPersonnelPost = strPersonnelPost & "/" & IIf(lstPosition.ListItems.Item(i).Checked = True, 1, 0)
Next i
strPersonnelPost = Mid(strPersonnelPost, 2, Len(strPersonnelPost))
strPersonnelGovt = ""
For i = 1 To lstGovtTables.ListItems.Count
    strPersonnelGovt = strPersonnelGovt & "/" & IIf(lstGovtTables.ListItems.Item(i).Checked = True, 1, 0)
Next i
strPersonnelGovt = Mid(strPersonnelGovt, 2, Len(strPersonnelGovt))
strPersonnelLoans = ""
For i = 1 To lstLoans.ListItems.Count
    strPersonnelLoans = strPersonnelLoans & "/" & IIf(lstLoans.ListItems.Item(i).Checked = True, 1, 0)
Next i
strPersonnelLoans = Mid(strPersonnelLoans, 2, Len(strPersonnelLoans))
strPersonnelCompensation = ""
For i = 1 To lstCompensation.ListItems.Count
    strPersonnelCompensation = strPersonnelCompensation & "/" & IIf(lstCompensation.ListItems.Item(i).Checked = True, 1, 0)
Next i
strPersonnelCompensation = Mid(strPersonnelCompensation, 2, Len(strPersonnelCompensation))
strUserAccount = ""
For i = 1 To lstUserRights.ListItems.Count
    strUserAccount = strUserAccount & "/" & IIf(lstUserRights.ListItems.Item(i).Checked = True, 1, 0)
Next i
strUserAccount = Mid(strUserAccount, 2, Len(strUserAccount))
strCompany = ""
For i = 1 To lstCompany.ListItems.Count
    strCompany = strCompany & "/" & IIf(lstCompany.ListItems.Item(i).Checked = True, 1, 0)
Next i
strCompany = Mid(strCompany, 2, Len(strCompany))
strTournamentInfo = ""
For i = 1 To lstTournament.ListItems.Count
    strTournamentInfo = strTournamentInfo & "/" & IIf(lstTournament.ListItems.Item(i).Checked = True, 1, 0)
Next i
strTournamentInfo = Mid(strTournamentInfo, 2, Len(strTournamentInfo))
strPlayerInfo = ""
For i = 1 To lstPlayer.ListItems.Count
    strPlayerInfo = strPlayerInfo & "/" & IIf(lstPlayer.ListItems.Item(i).Checked = True, 1, 0)
Next i
strPlayerInfo = Mid(strPlayerInfo, 2, Len(strPlayerInfo))
strTeamInfo = ""
For i = 1 To lstTeam.ListItems.Count
    strTeamInfo = strTeamInfo & "/" & IIf(lstTeam.ListItems.Item(i).Checked = True, 1, 0)
Next i
strTeamInfo = Mid(strTeamInfo, 2, Len(strPlayerInfo))
strScoreCard = ""
For i = 1 To lstScoreCard.ListItems.Count
    strScoreCard = strScoreCard & "/" & IIf(lstScoreCard.ListItems.Item(i).Checked = True, 1, 0)
Next i
strScoreCard = Mid(strScoreCard, 2, Len(strScoreCard))
strServiceChargeSetup = ""
For i = 1 To lstServiceChargeSetup.ListItems.Count
    strServiceChargeSetup = strServiceChargeSetup & "/" & IIf(lstServiceChargeSetup.ListItems.Item(i).Checked = True, 1, 0)
Next i
strServiceChargeSetup = Mid(strServiceChargeSetup, 2, Len(strServiceChargeSetup))
strServiceCharge = ""
For i = 1 To lstServiceCharge.ListItems.Count
    strServiceCharge = strServiceCharge & "/" & IIf(lstServiceCharge.ListItems.Item(i).Checked = True, 1, 0)
Next i
strServiceCharge = Mid(strServiceCharge, 2, Len(strServiceCharge))
strAbsentUndertimeEmployee = ""
For i = 1 To lstAbsentUndertimeEmployee.ListItems.Count
    strAbsentUndertimeEmployee = strAbsentUndertimeEmployee & "/" & IIf(lstAbsentUndertimeEmployee.ListItems.Item(i).Checked = True, 1, 0)
Next i
strAbsentUndertimeEmployee = Mid(strAbsentUndertimeEmployee, 2, Len(strAbsentUndertimeEmployee))
strSections = ""
For i = 1 To lstSection.ListItems.Count
    strSections = strSections & "/" & IIf(lstSection.ListItems.Item(i).Checked = True, 1, 0)
Next i
strSections = Mid(strSections, 2, Len(strSections))
strClassification = ""
For i = 1 To lstClassification.ListItems.Count
    strClassification = strClassification & "/" & IIf(lstClassification.ListItems.Item(i).Checked = True, 1, 0)
Next i
strClassification = Mid(strClassification, 2, Len(strClassification))
strSupplier = ""
For i = 1 To lstSupplier.ListItems.Count
    strSupplier = strSupplier & "/" & IIf(lstSupplier.ListItems.Item(i).Checked = True, 1, 0)
Next i
strSupplier = Mid(strSupplier, 2, Len(strSupplier))
strItemInfo = ""
For i = 1 To lstItemInfo.ListItems.Count
    strItemInfo = strItemInfo & "/" & IIf(lstItemInfo.ListItems.Item(i).Checked = True, 1, 0)
Next i
strItemInfo = Mid(strItemInfo, 2, Len(strItemInfo))
strServiceChargeSumm = ""
For i = 1 To lstServiceChargeSumm.ListItems.Count
    strServiceChargeSumm = strServiceChargeSumm & "/" & IIf(lstServiceChargeSumm.ListItems.Item(i).Checked = True, 1, 0)
Next i
strServiceChargeSumm = Mid(strServiceChargeSumm, 2, Len(strServiceChargeSumm))
strMemberInfo = ""
For i = 1 To lstMemberInfo.ListItems.Count
    strMemberInfo = strMemberInfo & "/" & IIf(lstMemberInfo.ListItems.Item(i).Checked = True, 1, 0)
Next i
strMemberInfo = Mid(strMemberInfo, 2, Len(strMemberInfo))
strShareID = ""
For i = 1 To lstShareID.ListItems.Count
    strShareID = strShareID & "/" & IIf(lstShareID.ListItems.Item(i).Checked = True, 1, 0)
Next i
strShareID = Mid(strShareID, 2, Len(strShareID))
strMemberIDNumber = ""
For i = 1 To lstMemberID.ListItems.Count
    strMemberIDNumber = strMemberIDNumber & "/" & IIf(lstMemberID.ListItems.Item(i).Checked = True, 1, 0)
Next i
strMemberIDNumber = Mid(strMemberIDNumber, 2, Len(strMemberIDNumber))
strCorporateAccount = ""
For i = 1 To lstCorporateAccnt.ListItems.Count
    strCorporateAccount = strCorporateAccount & "/" & IIf(lstCorporateAccnt.ListItems.Item(i).Checked = True, 1, 0)
Next i
strCorporateAccount = Mid(strCorporateAccount, 2, Len(strCorporateAccount))
strGolfCart = ""
For i = 1 To lstGolfCart.ListItems.Count
    strGolfCart = strGolfCart & "/" & IIf(lstGolfCart.ListItems.Item(i).Checked = True, 1, 0)
Next i
strGolfCart = Mid(strGolfCart, 2, Len(strGolfCart))
strCaddyInfo = ""
For i = 1 To lstCaddyInfo.ListItems.Count
    strCaddyInfo = strCaddyInfo & "/" & IIf(lstCaddyInfo.ListItems.Item(i).Checked = True, 1, 0)
Next i
strCaddyInfo = Mid(strCaddyInfo, 2, Len(strCaddyInfo))
strAllowance = ""
For i = 1 To lstAllowance.ListItems.Count
    strAllowance = strAllowance & "/" & IIf(lstAllowance.ListItems.Item(i).Checked = True, 1, 0)
Next i
strAllowance = Mid(strAllowance, 2, Len(strAllowance))
strMemberAction = ""
For i = 1 To lstMemberAction.ListItems.Count
    strMemberAction = strMemberAction & "/" & IIf(lstMemberAction.ListItems.Item(i).Checked = True, 1, 0)
Next i
strMemberAction = Mid(strMemberAction, 2, Len(strMemberAction))

strBackup = ""
For i = 1 To lstBackUp.ListItems.Count
    strBackup = strBackup & "/" & IIf(lstBackUp.ListItems.Item(i).Checked = True, 1, 0)
Next i
strBackup = Mid(strBackup, 2, Len(strBackup))

strPurchaseOrder = ""
For i = 1 To lstPO.ListItems.Count
    strPurchaseOrder = strPurchaseOrder & "/" & IIf(lstPO.ListItems.Item(i).Checked = True, 1, 0)
Next i
strPurchaseOrder = Mid(strPurchaseOrder, 2, Len(strPurchaseOrder))

strMenuMngt = ""
For i = 1 To lstMenuMngt.ListItems.Count
    strMenuMngt = strMenuMngt & "/" & IIf(lstMenuMngt.ListItems.Item(i).Checked = True, 1, 0)
Next i
strMenuMngt = Mid(strMenuMngt, 2, Len(strMenuMngt))
strFixedAsset = ""
For i = 1 To lstFixedAsset.ListItems.Count
    strFixedAsset = strFixedAsset & "/" & IIf(lstFixedAsset.ListItems.Item(i).Checked = True, 1, 0)
Next i
strFixedAsset = Mid(strFixedAsset, 2, Len(strFixedAsset))
strReceivingReport = ""
For i = 1 To lstRR.ListItems.Count
    strReceivingReport = strReceivingReport & "/" & IIf(lstRR.ListItems.Item(i).Checked = True, 1, 0)
Next i
strReceivingReport = Mid(strReceivingReport, 2, Len(strReceivingReport))
strStockTransfer = ""
For i = 1 To lstST.ListItems.Count
    strStockTransfer = strStockTransfer & "/" & IIf(lstST.ListItems.Item(i).Checked = True, 1, 0)
Next i
strStockTransfer = Mid(strStockTransfer, 2, Len(strStockTransfer))
strStockAdjustment = ""
For i = 1 To lstSA.ListItems.Count
    strStockAdjustment = strStockAdjustment & "/" & IIf(lstSA.ListItems.Item(i).Checked = True, 1, 0)
Next i
strStockAdjustment = Mid(strStockAdjustment, 2, Len(strStockAdjustment))
strStockIssuance = ""
For i = 1 To lstSI.ListItems.Count
    strStockIssuance = strStockIssuance & "/" & IIf(lstSI.ListItems.Item(i).Checked = True, 1, 0)
Next i
strStockIssuance = Mid(strStockIssuance, 2, Len(strStockIssuance))
strCheckVoucher = ""
For i = 1 To lstCheckVoucher.ListItems.Count
    strCheckVoucher = strCheckVoucher & "/" & IIf(lstCheckVoucher.ListItems.Item(i).Checked = True, 1, 0)
Next i
strCheckVoucher = Mid(strCheckVoucher, 2, Len(strCheckVoucher))
strPersonnelDeduction = ""
For i = 1 To lstDeduction.ListItems.Count
    strPersonnelDeduction = strPersonnelDeduction & "/" & IIf(lstDeduction.ListItems.Item(i).Checked = True, 1, 0)
Next i
strPersonnelDeduction = Mid(strPersonnelDeduction, 2, Len(strPersonnelDeduction))
strMortuary = ""
For i = 1 To lstMortuary.ListItems.Count
    strMortuary = strMortuary & "/" & IIf(lstMortuary.ListItems.Item(i).Checked = True, 1, 0)
Next i
strMortuary = Mid(strMortuary, 2, Len(strMortuary))
strChartOfAccounts = ""
For i = 1 To lstChartofAccounts.ListItems.Count
    strChartOfAccounts = strChartOfAccounts & "/" & IIf(lstChartofAccounts.ListItems.Item(i).Checked = True, 1, 0)
Next i
strChartOfAccounts = Mid(strChartOfAccounts, 2, Len(strChartOfAccounts))
strJournalVoucher = ""
For i = 1 To lstJournalVoucher.ListItems.Count
    strJournalVoucher = strJournalVoucher & "/" & IIf(lstJournalVoucher.ListItems.Item(i).Checked = True, 1, 0)
Next i
strJournalVoucher = Mid(strJournalVoucher, 2, Len(strJournalVoucher))
strOfficialReceipt = ""
For i = 1 To lstOR.ListItems.Count
    strOfficialReceipt = strOfficialReceipt & "/" & IIf(lstOR.ListItems.Item(i).Checked = True, 1, 0)
Next i
strOfficialReceipt = Mid(strOfficialReceipt, 2, Len(strOfficialReceipt))
strChargeInvoice = ""
For i = 1 To lstChargeInvoice.ListItems.Count
    strChargeInvoice = strChargeInvoice & "/" & IIf(lstChargeInvoice.ListItems.Item(i).Checked = True, 1, 0)
Next i
strChargeInvoice = Mid(strChargeInvoice, 2, Len(strChargeInvoice))
strAcknowledgementReceipt = ""
For i = 1 To lstAckReceipt.ListItems.Count
    strAcknowledgementReceipt = strAcknowledgementReceipt & "/" & IIf(lstAckReceipt.ListItems.Item(i).Checked = True, 1, 0)
Next i
strAcknowledgementReceipt = Mid(strAcknowledgementReceipt, 2, Len(strAcknowledgementReceipt))
strDebitCreditMemo = ""
For i = 1 To lstDRCRMemo.ListItems.Count
    strDebitCreditMemo = strDebitCreditMemo & "/" & IIf(lstDRCRMemo.ListItems.Item(i).Checked = True, 1, 0)
Next i
strDebitCreditMemo = Mid(strDebitCreditMemo, 2, Len(strDebitCreditMemo))
strPurchaseInvoice = ""
For i = 1 To lstPI.ListItems.Count
    strPurchaseInvoice = strPurchaseInvoice & "/" & IIf(lstPI.ListItems.Item(i).Checked = True, 1, 0)
Next i
strPurchaseInvoice = Mid(strPurchaseInvoice, 2, Len(strPurchaseInvoice))
strBagDrop = ""
For i = 1 To lstBagDrop.ListItems.Count
    strBagDrop = strBagDrop & "/" & IIf(lstBagDrop.ListItems.Item(i).Checked = True, 1, 0)
Next i
strBagDrop = Mid(strBagDrop, 2, Len(strBagDrop))
strRegistration = ""
For i = 1 To lstRegistration.ListItems.Count
    strRegistration = strRegistration & "/" & IIf(lstRegistration.ListItems.Item(i).Checked = True, 1, 0)
Next i
strRegistration = Mid(strRegistration, 2, Len(strRegistration))
strProShop = ""
For i = 1 To lstProShop.ListItems.Count
    strProShop = strProShop & "/" & IIf(lstProShop.ListItems.Item(i).Checked = True, 1, 0)
Next i
strProShop = Mid(strProShop, 2, Len(strProShop))
strDrivingRange = ""
For i = 1 To lstDrivingRange.ListItems.Count
    strDrivingRange = strDrivingRange & "/" & IIf(lstDrivingRange.ListItems.Item(i).Checked = True, 1, 0)
Next i
strDrivingRange = Mid(strDrivingRange, 2, Len(strDrivingRange))
strLockerRoom = ""
For i = 1 To lstLockerRoom.ListItems.Count
    strLockerRoom = strLockerRoom & "/" & IIf(lstLockerRoom.ListItems.Item(i).Checked = True, 1, 0)
Next i
strLockerRoom = Mid(strLockerRoom, 2, Len(strLockerRoom))
strGolfCartOP = ""
For i = 1 To lstGolfCartOP.ListItems.Count
    strGolfCartOP = strGolfCartOP & "/" & IIf(lstGolfCartOP.ListItems.Item(i).Checked = True, 1, 0)
Next i
strGolfCartOP = Mid(strGolfCartOP, 2, Len(strGolfCartOP))
strFnBLocation = ""
For i = 1 To lstFnBLocation.ListItems.Count
    strFnBLocation = strFnBLocation & "/" & IIf(lstFnBLocation.ListItems.Item(i).Checked = True, 1, 0)
Next i
strFnBLocation = Mid(strFnBLocation, 2, Len(strFnBLocation))
strProShopItems = ""
For i = 1 To lstProShopItems.ListItems.Count
    strProShopItems = strProShopItems & "/" & IIf(lstProShopItems.ListItems.Item(i).Checked = True, 1, 0)
Next i
strProShopItems = Mid(strProShopItems, 2, Len(strProShopItems))
strProShopItemsBrand = ""
For i = 1 To lstProShopItemsBrand.ListItems.Count
    strProShopItemsBrand = strProShopItemsBrand & "/" & IIf(lstProShopItemsBrand.ListItems.Item(i).Checked = True, 1, 0)
Next i
strProShopItemsBrand = Mid(strProShopItemsBrand, 2, Len(strProShopItemsBrand))
strProShopItemsModel = ""
For i = 1 To lstProShopItemsModel.ListItems.Count
    strProShopItemsModel = strProShopItemsModel & "/" & IIf(lstProShopItemsModel.ListItems.Item(i).Checked = True, 1, 0)
Next i
strProShopItemsModel = Mid(strProShopItemsModel, 2, Len(strProShopItemsModel))
strProShopItemsSizes = ""
For i = 1 To lstProShopItemsSize.ListItems.Count
    strProShopItemsSizes = strProShopItemsSizes & "/" & IIf(lstProShopItemsSize.ListItems.Item(i).Checked = True, 1, 0)
Next i
strProShopItemsSizes = Mid(strProShopItemsSizes, 2, Len(strProShopItemsSizes))
strProShopItemsColor = ""
For i = 1 To lstProShopItemsColor.ListItems.Count
    strProShopItemsColor = strProShopItemsColor & "/" & IIf(lstProShopItemsColor.ListItems.Item(i).Checked = True, 1, 0)
Next i
strProShopItemsColor = Mid(strProShopItemsColor, 2, Len(strProShopItemsColor))
strProShopItemsItemType = ""
For i = 1 To lstProShopItemsItemType.ListItems.Count
    strProShopItemsItemType = strProShopItemsItemType & "/" & IIf(lstProShopItemsItemType.ListItems.Item(i).Checked = True, 1, 0)
Next i
strProShopItemsItemType = Mid(strProShopItemsItemType, 2, Len(strProShopItemsItemType))
strPersonnelHours = ""
For i = 1 To lstHours.ListItems.Count
    strPersonnelHours = strPersonnelHours & "/" & IIf(lstHours.ListItems.Item(i).Checked = True, 1, 0)
Next i
strPersonnelHours = Mid(strPersonnelHours, 2, Len(strPersonnelHours))
strPersonnelForDeduction = ""
For i = 1 To lstforDeduction.ListItems.Count
    strPersonnelForDeduction = strPersonnelForDeduction & "/" & IIf(lstforDeduction.ListItems.Item(i).Checked = True, 1, 0)
Next i
strPersonnelForDeduction = Mid(strPersonnelForDeduction, 2, Len(strPersonnelForDeduction))

strPersonnelSetUpPerfectDays = ""
For i = 1 To lstPerfectDaysDaily.ListItems.Count
    strPersonnelSetUpPerfectDays = strPersonnelSetUpPerfectDays & "/" & IIf(lstPerfectDaysDaily.ListItems.Item(i).Checked = True, 1, 0)
Next i
strPersonnelSetUpPerfectDays = Mid(strPersonnelSetUpPerfectDays, 2, Len(strPersonnelSetUpPerfectDays))

strPersonnelPagIbigAddContri = ""
For i = 1 To lstPagIbigAddContri.ListItems.Count
    strPersonnelPagIbigAddContri = strPersonnelPagIbigAddContri & "/" & IIf(lstPagIbigAddContri.ListItems.Item(i).Checked = True, 1, 0)
Next i
strPersonnelPagIbigAddContri = Mid(strPersonnelPagIbigAddContri, 2, Len(strPersonnelPagIbigAddContri))

strPersonnelManualPayment = ""
For i = 1 To lstManualPayment.ListItems.Count
    strPersonnelManualPayment = strPersonnelManualPayment & "/" & IIf(lstManualPayment.ListItems.Item(i).Checked = True, 1, 0)
Next i
strPersonnelManualPayment = Mid(strPersonnelManualPayment, 2, Len(strPersonnelManualPayment))

strUserName = Replace(FORMATSQL(Trim(txtUserName.Text)), " ", "")

'On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    
    ConnOmega.Execute "INSERT INTO tbl_Users_Account " & _
                      " (UserName, Password, PersonnelInfo, PersonnelID, PersonnelAction, PersonnelDept, PersonnelStatus, PersonnelPost, PersonnelGovt, PersonnelLoan, " & _
                      " PersonnelCompensation, UserAccount, CompanyInfo, LastModified, ScoringTournamentInfo, ScoringPlayerInfo, ScoringTeamInfo, ScoringScoreCard, " & _
                      " ServiceChargeSetup, ServiceCharge, Sections , Classification, Supplier, ItemInfo, ServiceChargeSumm, MemberInfo, ShareID, MemberIDNumber, " & _
                      " CorporateAccount, GolfCart, CaddyInfo, Admin, Allowance, MemberAction, AllowBackup, CompleteName, PurchaseOrder, MenuMngt, FixedAsset, " & _
                      " ReceivingReport, StockTransfer, StockAdjustment, StockIssuance, CheckVoucher, PersonnelDeduction, Mortuary, ChartOfAccounts, JournalVoucher, " & _
                      " OfficialReceipt, ChargeInvoice, AcknowledgementReceipt, DebitCreditMemo, PurchaseInvoice, BagDrop, Registration, ProShop, DrivingRange, LockerRoom, " & _
                      " GolfCartOP, ProShopItems, ProItemsBrand, ProShopItemsBrand, ProShopItemsModel, ProShopItemsSizes, ProShopItemsColor, ProShopItemsItemType, PersonnelHours, " & _
                      " PersonnelForDeduction, PersonnelSetUpPerfectDays, PersonnelPagIbigAddContri, PersonnelManualPayment) " & _
                      " VALUES ('" & strUserName & "', '" & EncryptDecrypt(FORMATSQL(Trim(txtPassword.Text))) & "', '" & strPersonnelInfo & "', " & _
                      " '" & strPersonnelID & "', '" & strPersonnelAction & "', '" & strPersonnelDept & "', '" & strPersonnelStatus & "', '" & strPersonnelPost & "', " & _
                      " '" & strPersonnelGovt & "', '" & strPersonnelLoans & "', '" & strPersonnelCompensation & "', '" & strUserAccount & "', '" & strCompany & "', " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "', '" & strTournamentInfo & "', '" & strPlayerInfo & "', '" & strTeamInfo & "', '" & strScoreCard & "', " & _
                      " '" & strServiceChargeSetup & "', '" & strServiceCharge & "', '" & strSections & "', '" & strClassification & "', '" & strSupplier & "', '" & strItemInfo & "', " & _
                      " '" & strServiceChargeSumm & "', '" & strMemberInfo & "', '" & strShareID & "', '" & strMemberIDNumber & "', '" & strCorporateAccount & "', " & _
                      " '" & strGolfCart & "', '" & strCaddyInfo & "', " & chkAdministrator.Value & ", '" & strAllowance & "', '" & strMemberAction & "', '" & strBackup & "', " & _
                      " '" & FORMATSQL(Trim(txtCompleteName.Text)) & "', '" & strPurchaseOrder & "', '" & strMenuMngt & "', '" & strFixedAsset & "', '" & strReceivingReport & "', " & _
                      " '" & strStockTransfer & "', '" & strStockAdjustment & "', '" & strStockIssuance & "', '" & strCheckVoucher & "', '" & strPersonnelDeduction & "', " & _
                      " '" & strMortuary & "', '" & strChartOfAccounts & "', '" & strJournalVoucher & "', '" & strOfficialReceipt & "', '" & strChargeInvoice & "', '" & strAcknowledgementReceipt & "', " & _
                      " '" & strDebitCreditMemo & "', '" & strPurchaseInvoice & "', '" & strBagDrop & "', '" & strRegistration & "', '" & strProShop & "', '" & strDrivingRange & "', " & _
                      " '" & strLockerRoom & "', '" & strGolfCartOP & "', '" & strFnBLocation & "', '" & strProShopItems & "', '" & strProShopItemsBrand & "', '" & strProShopItemsModel & "', " & _
                      " '" & strProShopItemsSizes & "', '" & strProShopItemsColor & "', '" & strProShopItemsItemType & "', '" & strPersonnelHours & "', '" & strPersonnelForDeduction & "', " & _
                      " '" & strPersonnelSetUpPerfectDays & "', '" & strPersonnelPagIbigAddContri & "', '" & strPersonnelManualPayment & "')"

End If
If TRANSACTIONTYPE = is_EDITTING Then
    
    ConnOmega.Execute "UPDATE tbl_Users_Account " & _
                      " SET UserName = '" & strUserName & "', Password = '" & EncryptDecrypt(FORMATSQL(Trim(txtPassword.Text))) & "', PersonnelInfo = '" & strPersonnelInfo & "', PersonnelID = '" & strPersonnelID & "', PersonnelAction = '" & strPersonnelAction & "', " & _
                      " PersonnelDept = '" & strPersonnelDept & "', PersonnelStatus = '" & strPersonnelStatus & "', PersonnelPost = '" & strPersonnelPost & "', " & _
                      " PersonnelGovt = '" & strPersonnelGovt & "', PersonnelLoan = '" & strPersonnelLoans & "', PersonnelCompensation = '" & strPersonnelCompensation & "', " & _
                      " UserAccount = '" & strUserAccount & "', CompanyInfo = '" & strCompany & "', LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                      " ScoringTournamentInfo = '" & strTournamentInfo & "', ScoringPlayerInfo = '" & strPlayerInfo & "', ScoringTeamInfo = '" & strTeamInfo & "', " & _
                      " ScoringScoreCard = '" & strScoreCard & "', ServiceChargeSetup = '" & strServiceChargeSetup & "', ServiceCharge = '" & strServiceCharge & "', " & _
                      " AbsentUndertime = '" & strAbsentUndertimeEmployee & "', Sections = '" & strSections & "', Classification = '" & strClassification & "', " & _
                      " Supplier = '" & strSupplier & "', ItemInfo = '" & strItemInfo & "', ServiceChargeSumm = '" & strServiceChargeSumm & "', " & _
                      " MemberInfo = '" & strMemberInfo & "', ShareID = '" & strShareID & "', MemberIDNumber = '" & strMemberIDNumber & "', " & _
                      " CorporateAccount = '" & strCorporateAccount & "', GolfCart = '" & strGolfCart & "', CaddyInfo = '" & strCaddyInfo & "', " & _
                      " Admin = " & chkAdministrator.Value & ", Allowance = '" & strAllowance & "', MemberAction = '" & strMemberAction & "', AllowBackup = '" & strBackup & "', " & _
                      " CompleteName = '" & FORMATSQL(Trim(txtCompleteName.Text)) & "', PurchaseOrder = '" & strPurchaseOrder & "', MenuMngt = '" & strMenuMngt & "', " & _
                      " FixedAsset = '" & strFixedAsset & "', ReceivingReport = '" & strReceivingReport & "', StockTransfer = '" & strStockTransfer & "', " & _
                      " StockAdjustment = '" & strStockAdjustment & "', StockIssuance = '" & strStockIssuance & "', CheckVoucher = '" & strCheckVoucher & "', " & _
                      " PersonnelDeduction = '" & strPersonnelDeduction & "', Mortuary = '" & strMortuary & "', ChartOfAccounts = '" & strChartOfAccounts & "', " & _
                      " JournalVoucher = '" & strJournalVoucher & "', OfficialReceipt = '" & strOfficialReceipt & "', ChargeInvoice = '" & strChargeInvoice & "', " & _
                      " AcknowledgementReceipt = '" & strAcknowledgementReceipt & "', DebitCreditMemo = '" & strDebitCreditMemo & "', PurchaseInvoice = '" & strPurchaseInvoice & "', " & _
                      " BagDrop = '" & strBagDrop & "', Registration = '" & strRegistration & "', ProShop = '" & strProShop & "', DrivingRange = '" & strDrivingRange & "', " & _
                      " LockerRoom = '" & strLockerRoom & "', GolfCartOP = '" & strGolfCartOP & "', FnBLocation = '" & strFnBLocation & "',ProShopItems = '" & strProShopItems & "',  " & _
                      " ProShopItemsBrand = '" & strProShopItemsBrand & "', ProShopItemsModel = '" & strProShopItemsModel & "', ProShopItemsSizes = '" & strProShopItemsSizes & "', " & _
                      " ProShopItemsColor = '" & strProShopItemsColor & "', ProShopItemsItemType = '" & strProShopItemsItemType & "', PersonnelHours = '" & strPersonnelHours & "', " & _
                      " PersonnelForDeduction = '" & strPersonnelForDeduction & "', PersonnelSetUpPerfectDays = '" & strPersonnelSetUpPerfectDays & "', " & _
                      " PersonnelPagIbigAddContri = '" & strPersonnelPagIbigAddContri & "', PersonnelManualPayment = '" & strPersonnelManualPayment & "' " & _
                      " WHERE (PK = " & StatusBar1.Panels(1).Text & ")"
 
End If
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
'Me.Caption = "User Account - Browse"
BROWSER strUserName, "is_LOAD"
Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Function
End Function

Private Function PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
CLEARTEXT
TOOLBARFUNC 3
TRANSACTIONTYPE = is_FINDING
'Me.Caption = "User Account - Find"
picBody.Enabled = False
txtUserNameFind.Text = ""
txtUserNameFind.ZOrder 0
txtUserNameFind.Move picBody.Left + txtUserName.Left, picBody.Top + txtUserName.Top, txtUserName.Width, txtUserName.Height
txtUserNameFind.Visible = True
txtUserNameFind.SetFocus
End Function

Private Function PRESS_F8()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If StatusBar1.Panels(1).Text = "" Then Exit Function
If AccessRights("User's Account", "Admin") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
If AccessRights("User's Account", "Admin") = True Then
    If CDbl(iAdmin) = 1 Then
        If CStr(gbl_UserName) <> Trim(txtUserName.Text) Then
            MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
                   "ACCESS DENIED!                                      ", vbCritical, "Alert"
            Exit Function
        End If
    End If
End If
If MsgBox("RESETING TO DEFAULT PASSWORD                         " & vbCrLf & "CONTINUE?                     ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Function
ConnOmega.Execute "UPDATE tbl_Users_Account SET Password = '" & EncryptDecrypt(FORMATSQL(Trim(CStr(sDefaultPW)))) & "' WHERE (PK = " & StatusBar1.Panels(1).Text & ")"
gbl_Password = sDefaultPW
txtPassword.Text = sDefaultPW
txtPassword01.Text = sDefaultPW
MsgBox "SUCCESSFULLY RESET PASSWORD!                    ", vbInformation, "Reset"
End Function

Private Function PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    Unload Me
Else
    If TRANSACTIONTYPE = is_FINDING Then
        txtUserNameFind.Visible = False
        picBody.Enabled = True
    End If
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    'Me.Caption = "User Account - Browse"
    BROWSER GetSetting(App.EXEName, "UserAccountName", "UserAccntName", ""), "is_LOAD"
    If Trim(txtUserName.Text) = "" Then BROWSER GetSetting(App.EXEName, "UserAccountName", "UserAccntName", ""), "is_HOME"
End If
End Function

Private Sub ShowObjectInStatusBar(ByVal bShowObject As Boolean)
Dim tRC As RECT
    SendMessageAny StatusBar1.hwnd, SB_GETRECT, 2, tRC
    With tRC
        .Top = (.Top * Screen.TwipsPerPixelY)
        .Left = (.Left * Screen.TwipsPerPixelX)
        .Bottom = (.Bottom * Screen.TwipsPerPixelY) - .Top
        .Right = (.Right * Screen.TwipsPerPixelX) - .Left
    End With
    With picAdministrator

        SetParent .hwnd, StatusBar1.hwnd
        .Move tRC.Left + 40, tRC.Top + 30, tRC.Right - 80, tRC.Bottom - 80
        .Visible = True
    End With
End Sub


Private Function CLEARTEXT()
txtUserName.Text = ""
txtPassword.Text = ""
txtCompleteName.Text = ""

chkAdministrator.Value = 0
StatusBar1.Panels(1).Text = ""
StatusBar1.Panels(2).Text = ""

With lstManualPayment.ListItems
    .Clear
    For i = 1 To 6
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POST"
            Case 6: x.SubItems(1) = "UNPOST"
        End Select
    Next i
End With

With lstPagIbigAddContri.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstPerfectDaysDaily.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstforDeduction.ListItems
    .Clear
    For i = 1 To 6
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POST"
            Case 6: x.SubItems(1) = "UNPOST"
        End Select
    Next i
End With

With lstHours.ListItems
    .Clear
    For i = 1 To 6
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POST"
            Case 6: x.SubItems(1) = "UNPOST"
        End Select
    Next i
End With

With lstProShopItemsItemType.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstProShopItemsColor.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstProShopItemsSize.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstProShopItemsModel.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstProShopItemsBrand.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstFnBLocation.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstGolfCartOP.ListItems
    .Clear
    For i = 1 To 5
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POSTING"
        End Select
    Next i
End With


With lstLockerRoom.ListItems
    .Clear
    For i = 1 To 5
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POSTING"
        End Select
    Next i
End With

With lstDrivingRange.ListItems
    .Clear
    For i = 1 To 5
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POSTING"
        End Select
    Next i
End With

With lstProShop.ListItems
    .Clear
    For i = 1 To 5
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POSTING"
        End Select
    Next i
End With

With lstProShopItems.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstRegistration.ListItems
    .Clear
    For i = 1 To 5
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POSTING"
        End Select
    Next i
End With

With lstBagDrop.ListItems
    .Clear
    For i = 1 To 5
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POSTING"
        End Select
    Next i
End With

With lstChartofAccounts.ListItems
    .Clear
    For i = 1 To 5
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POSTING RANGE"
        End Select
    Next i
End With

With lstMortuary.ListItems
    .Clear
    For i = 1 To 6
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POST"
            Case 6: x.SubItems(1) = "UNPOST"
        End Select
    Next i
End With

With lstDeduction.ListItems
    .Clear
    For i = 1 To 6
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POST"
            Case 6: x.SubItems(1) = "UNPOST"
        End Select
    Next i
End With

With lstDRCRMemo.ListItems
    .Clear
    For i = 1 To 5
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POST"
        End Select
    Next i
End With

With lstOR.ListItems
    .Clear
    For i = 1 To 5
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POST"
        End Select
    Next i
End With

With lstAckReceipt.ListItems
    .Clear
    For i = 1 To 5
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POST"
        End Select
    Next i
End With

With lstCheckVoucher.ListItems
    .Clear
    For i = 1 To 5
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POST"
        End Select
    Next i
End With

With lstJournalVoucher.ListItems
    .Clear
    For i = 1 To 5
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POST"
        End Select
    Next i
End With

With lstPettyCash.ListItems
    .Clear
    For i = 1 To 5
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POST"
        End Select
    Next i
End With

With lstChargeInvoice.ListItems
    .Clear
    For i = 1 To 5
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POST"
        End Select
    Next i
End With

With lstSI.ListItems
    .Clear
    For i = 1 To 6
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POST"
            Case 6: x.SubItems(1) = "POST INV"
        End Select
    Next i
End With

With lstSA.ListItems
    .Clear
    For i = 1 To 5
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POST"
        End Select
    Next i
End With

With lstST.ListItems
    .Clear
    For i = 1 To 5
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POST"
        End Select
    Next i
End With

With lstRR.ListItems
    .Clear
    For i = 1 To 7
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POST INV"
            Case 6: x.SubItems(1) = "POST GL"
            Case 7: x.SubItems(1) = "PRINT"
        End Select
    Next i
End With

With lstPI.ListItems
    .Clear
    For i = 1 To 7
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "PRINT"
            Case 6: x.SubItems(1) = "POST INV"
            Case 7: x.SubItems(1) = "POST GL"
        End Select
    Next i
End With

With lstFixedAsset.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstMenuMngt.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstPO.ListItems
    .Clear
    For i = 1 To 6
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "PRINT"
            Case 6: x.SubItems(1) = "POST"
        End Select
    Next i
End With

With lstBackUp.ListItems
    .Clear
    For i = 1 To 1
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "BACKUP"
'            Case 2: x.SubItems(1) = "ADD"
'            Case 3: x.SubItems(1) = "EDIT"
'            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstMemberAction.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstAllowance.ListItems
    .Clear
    For i = 1 To 5
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "GENERATE"
        End Select
    Next i
End With

With lstCaddyInfo.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstGolfCart.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstCorporateAccnt.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstMemberID.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstShareID.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstMemberInfo.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstAbsentUndertimeEmployee.ListItems
    .Clear
    For i = 1 To 6
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POST"
            Case 6: x.SubItems(1) = "UNPOST"
        End Select
    Next i
End With

With lstServiceCharge.ListItems
    .Clear
    For i = 1 To 6
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POST"
            Case 6: x.SubItems(1) = "UNPOST"
        End Select
    Next i
End With

With lstServiceChargeSumm.ListItems
    .Clear
    For i = 1 To 5
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POST"
        End Select
    Next i
End With

With lstServiceChargeSetup.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstTournament.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstPlayer.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstTeam.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstScoreCard.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstPersonnalInfo.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstIDNumber.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstActionMemo.ListItems
    .Clear
    For i = 1 To 6
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "PRINT"
            Case 6: x.SubItems(1) = "SUPERVISORY"
        End Select
    Next i
End With

With lstGovtTables.ListItems
    .Clear
    For i = 1 To 5
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "SSS"
            Case 2: x.SubItems(1) = "PHIL HEALTH"
            Case 3: x.SubItems(1) = "PAG - IBIG"
            Case 4: x.SubItems(1) = "TAXABLE INCOME"
            Case 5: x.SubItems(1) = "PERSONNAL EXEMPTIONS"
        End Select
    Next i
End With


With lstLoans.ListItems
    .Clear
    For i = 1 To 6
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "POST"
            Case 6: x.SubItems(1) = "UNPOST"
        End Select
    Next i
End With

With lstCompensation.ListItems
    .Clear
    For i = 1 To 6
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
            Case 5: x.SubItems(1) = "SUPERVISORY"
            Case 6: x.SubItems(1) = "LOCKED PAYROLL"
        End Select
    Next i
End With

With lstDept.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstEmpStatus.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstPosition.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstUserRights.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

With lstCompany.ListItems
    .Clear
    For i = 1 To 2
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "EDIT"
        End Select
    Next i
End With

With lstSection.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With
With lstClassification.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With
With lstSupplier.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With
With lstItemInfo.ListItems
    .Clear
    For i = 1 To 4
        Set x = .Add()
        x.Text = ""
        Select Case i
            Case 1: x.SubItems(1) = "OPEN"
            Case 2: x.SubItems(1) = "ADD"
            Case 3: x.SubItems(1) = "EDIT"
            Case 4: x.SubItems(1) = "DELETE"
        End Select
    Next i
End With

End Function

Private Function LOCKTEXT(bln As Boolean)
txtUserName.Locked = bln
txtCompleteName.Locked = bln
picPersonnel.Enabled = IIf(bln = True, False, True)
picMembership.Enabled = IIf(bln = True, False, True)
'picScoring.Enabled = IIf(bln = True, False, True)
picGolfScoring.Enabled = IIf(bln = True, False, True)
picUtility.Enabled = IIf(bln = True, False, True)
picGolfOperation.Enabled = IIf(bln = True, False, True)
picFinance.Enabled = IIf(bln = True, False, True)
picFnB.Enabled = IIf(bln = True, False, True)
txtPassword.Locked = True
End Function

Private Function TOOLBARFUNC(intSel As Integer)
With Toolbar1
    Select Case intSel
        Case 1      'REFRESH
            .Buttons(1).Enabled = True
            .Buttons(3).Enabled = True
            .Buttons(5).Enabled = True
            .Buttons(7).Enabled = True
            .Buttons(7).Caption = "First"
            .Buttons(7).Image = 4
            .Buttons(9).Enabled = True
            .Buttons(9).Caption = "Back"
            .Buttons(9).Image = 5
            .Buttons(11).Enabled = True
            .Buttons(13).Enabled = True
            .Buttons(15).Enabled = True
            .Buttons(17).Enabled = True
            .Buttons(19).Enabled = True
'            .Buttons(21).Enabled = True
        Case 2      'ADD/EDIT
            .Buttons(1).Enabled = False
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = False
            .Buttons(7).Enabled = True
            .Buttons(7).Caption = "Save"
            .Buttons(7).Image = 14
            .Buttons(9).Enabled = True
            .Buttons(9).Caption = "Undo"
            .Buttons(9).Image = 15
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
            .Buttons(19).Enabled = False
'            .Buttons(21).Enabled = False
        Case 3      'FIND
            .Buttons(1).Enabled = False
            .Buttons(3).Enabled = False
            .Buttons(5).Enabled = False
            .Buttons(7).Enabled = False
            .Buttons(7).Caption = "First"
            .Buttons(7).Image = 4
            .Buttons(9).Enabled = True
            .Buttons(9).Caption = "Undo"
            .Buttons(9).Image = 15
            .Buttons(11).Enabled = False
            .Buttons(13).Enabled = False
            .Buttons(15).Enabled = False
            .Buttons(17).Enabled = False
            .Buttons(19).Enabled = False
'            .Buttons(21).Enabled = False
    End Select
End With
End Function

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
frameAllowance.Visible = AccessRights("User's Account", "Admin")
chkAdministrator.Visible = AccessRights("User's Account", "Admin")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert:   PRESS_INSERT
    Case vbKeyF2:       PRESS_F2
    Case vbKeyDelete:   PRESS_DELETE
    Case vbKeyF5:       PRESS_F5
    Case vbKeyF6:       PRESS_F6
    Case vbKeyF8:       PRESS_F8
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "UserAccountName", "UserAccntName", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "UserAccountName", "UserAccntName", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "UserAccountName", "UserAccntName", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "UserAccountName", "UserAccntName", ""), "is_END"
    Case vbKeyEscape:   PRESS_ESCAPE
End Select

ShiftDown = (Shift And vbShiftMask) > 0
AltDown = (Shift And vbAltMask) > 0
CtrlDown = (Shift And vbCtrlMask) > 0
If CtrlDown And AltDown And _
ShiftDown And KeyCode = vbKeyA Then
    If AccessRights("User's Account", "Admin") = True Then
        If txtPassword01.Visible = False Then
            txtPassword01.Visible = True
        Else
            txtPassword01.Visible = False
        End If
    End If
End If

End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.Height - Me.Height) / 6
Me.Left = (MainForm.Width - Me.Width) / 6

ShowObjectInStatusBar True


frameAllowance.Visible = AccessRights("User's Account", "Admin")
chkAdministrator.Visible = AccessRights("User's Account", "Admin")

XTab1.ActiveTab = 0
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "UserAccountName", "UserAccntName", ""), "is_LOAD"
If Trim(txtUserName.Text) = "" Then BROWSER GetSetting(App.EXEName, "UserAccountName", "UserAccntName", ""), "is_HOME"
Dim tmp As Long
tmp = SetWindowLong(txtUserName.hwnd, GWL_STYLE, GetWindowLong(txtUserName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPassword.hwnd, GWL_STYLE, GetWindowLong(txtPassword.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtUserNameFind.hwnd, GWL_STYLE, GetWindowLong(txtUserNameFind.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtCompleteName.hwnd, GWL_STYLE, GetWindowLong(txtCompleteName.hwnd, GWL_STYLE) Or ES_UPPERCASE)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "UserAccountName", "UserAccntName", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "UserAccountName", "UserAccntName", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "UserAccountName", "UserAccntName", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "UserAccountName", "UserAccntName", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Reset":   PRESS_F8
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtCompleteName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then txtUserName.SetFocus
End Sub

Private Sub txtPassword_GotFocus()
HTEXT txtPassword
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
    txtUserName.SetFocus
End If
End Sub

Private Sub txtUserName_GotFocus()
HTEXT txtUserName
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtCompleteName.SetFocus
'    txtPassword.SetFocus
End If
End Sub

