VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{76880EFA-2CCC-4791-B35E-F6A7359CAFDD}#1.0#0"; "prjXTab.ocx"
Begin VB.Form frmPersonnelInformation 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10365
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPersonnelInformation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   120
      ScaleHeight     =   5055
      ScaleWidth      =   10095
      TabIndex        =   1
      Top             =   1200
      Width           =   10095
      Begin prjXTab.XTab XTab1 
         Height          =   5055
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   8916
         TabCount        =   5
         TabCaption(0)   =   "Personnal Information"
         TabContCtrlCnt(0)=   1
         Tab(0)ContCtrlCap(1)=   "picPersonnal"
         TabCaption(1)   =   "Family Background"
         TabContCtrlCnt(1)=   1
         Tab(1)ContCtrlCap(1)=   "picFamily"
         TabCaption(2)   =   "Education"
         TabContCtrlCnt(2)=   1
         Tab(2)ContCtrlCap(1)=   "picEducation"
         TabCaption(3)   =   "Miscellaneous"
         TabContCtrlCnt(3)=   1
         Tab(3)ContCtrlCap(1)=   "picEmployment"
         TabCaption(4)   =   "Picture"
         TabContCtrlCnt(4)=   1
         Tab(4)ContCtrlCap(1)=   "picPicture"
         ActiveTab       =   1
         TabTheme        =   1
         ActiveTabBackStartColor=   16514555
         ActiveTabBackEndColor=   13023396
         InActiveTabBackStartColor=   16777215
         InActiveTabBackEndColor=   15397104
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
         OuterBorderColor=   13023396
         DisabledTabBackColor=   -2147483633
         DisabledTabForeColor=   10526880
         Begin VB.PictureBox picPicture 
            BackColor       =   &H00C6B8A4&
            BorderStyle     =   0  'None
            Height          =   4455
            Left            =   -74880
            ScaleHeight     =   4455
            ScaleWidth      =   9975
            TabIndex        =   127
            Top             =   480
            Width           =   9975
            Begin VB.PictureBox picTmpInfo 
               BackColor       =   &H00C6B8A4&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4095
               Left            =   120
               ScaleHeight     =   4095
               ScaleWidth      =   5295
               TabIndex        =   129
               Top             =   360
               Width           =   5295
               Begin VB.TextBox txtTmpBloodType 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  ForeColor       =   &H000000FF&
                  Height          =   315
                  Left            =   3600
                  TabIndex        =   237
                  Top             =   3240
                  Width           =   1695
               End
               Begin VB.TextBox txtTmpStatus 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  ForeColor       =   &H000000FF&
                  Height          =   315
                  Left            =   3600
                  TabIndex        =   219
                  Top             =   3600
                  Width           =   1695
               End
               Begin VB.TextBox txtTmpIDNumber 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  ForeColor       =   &H000000FF&
                  Height          =   315
                  Left            =   840
                  TabIndex        =   217
                  Top             =   3600
                  Width           =   1695
               End
               Begin VB.TextBox txtPicturePath 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   160
                  Top             =   1440
                  Visible         =   0   'False
                  Width           =   435
               End
               Begin VB.TextBox txtTmpDriverLic 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  ForeColor       =   &H000000FF&
                  Height          =   315
                  Left            =   840
                  TabIndex        =   159
                  Top             =   3240
                  Width           =   1695
               End
               Begin VB.TextBox txtTmpPagibig 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  ForeColor       =   &H000000FF&
                  Height          =   315
                  Left            =   3600
                  TabIndex        =   157
                  Top             =   2880
                  Width           =   1695
               End
               Begin VB.TextBox txtTmpPHIC 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  ForeColor       =   &H000000FF&
                  Height          =   315
                  Left            =   3600
                  TabIndex        =   155
                  Top             =   2520
                  Width           =   1695
               End
               Begin VB.TextBox txtTmpTIN 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  ForeColor       =   &H000000FF&
                  Height          =   315
                  Left            =   840
                  TabIndex        =   153
                  Top             =   2880
                  Width           =   1695
               End
               Begin VB.TextBox txtTmpSSS 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  ForeColor       =   &H000000FF&
                  Height          =   315
                  Left            =   840
                  TabIndex        =   151
                  Top             =   2520
                  Width           =   1695
               End
               Begin VB.TextBox txtTmpContact 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  ForeColor       =   &H000000FF&
                  Height          =   315
                  Left            =   840
                  TabIndex        =   149
                  Top             =   1800
                  Width           =   4455
               End
               Begin VB.TextBox txtTmpWeight 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  ForeColor       =   &H000000FF&
                  Height          =   315
                  Left            =   3600
                  TabIndex        =   147
                  Top             =   2160
                  Width           =   1695
               End
               Begin VB.TextBox txtTmpHeight 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  ForeColor       =   &H000000FF&
                  Height          =   315
                  Left            =   840
                  TabIndex        =   145
                  Top             =   2160
                  Width           =   1695
               End
               Begin VB.TextBox txtTmpAddress 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  ForeColor       =   &H000000FF&
                  Height          =   675
                  Left            =   840
                  MultiLine       =   -1  'True
                  TabIndex        =   143
                  Top             =   1080
                  Width           =   4455
               End
               Begin VB.TextBox txtTmpCivilStatus 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  ForeColor       =   &H000000FF&
                  Height          =   315
                  Left            =   3600
                  TabIndex        =   139
                  Top             =   720
                  Width           =   1695
               End
               Begin VB.TextBox txtTmpAge 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  ForeColor       =   &H000000FF&
                  Height          =   315
                  Left            =   3600
                  TabIndex        =   137
                  Top             =   360
                  Width           =   1695
               End
               Begin VB.TextBox txtTmpGender 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  ForeColor       =   &H000000FF&
                  Height          =   315
                  Left            =   840
                  TabIndex        =   135
                  Top             =   720
                  Width           =   1695
               End
               Begin VB.TextBox txtTmpBDay 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  ForeColor       =   &H000000FF&
                  Height          =   315
                  Left            =   840
                  TabIndex        =   133
                  Top             =   360
                  Width           =   1695
               End
               Begin VB.TextBox txtTmpFullName 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C6B8A4&
                  ForeColor       =   &H00FF0000&
                  Height          =   315
                  Left            =   840
                  TabIndex        =   131
                  Text            =   "ARCHIE"
                  Top             =   0
                  Width           =   4455
               End
               Begin VB.Label Label91 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Blood Type"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   238
                  Top             =   3240
                  Width           =   1095
               End
               Begin VB.Label Label88 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Status"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   220
                  Top             =   3600
                  Width           =   1095
               End
               Begin VB.Label Label87 
                  BackStyle       =   0  'Transparent
                  Caption         =   "ID Number"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   218
                  Top             =   3600
                  Width           =   1095
               End
               Begin VB.Label Label68 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Driver Lic"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   158
                  Top             =   3240
                  Width           =   1095
               End
               Begin VB.Label Label67 
                  BackStyle       =   0  'Transparent
                  Caption         =   "PagIbig"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   156
                  Top             =   2880
                  Width           =   1095
               End
               Begin VB.Label Label66 
                  BackStyle       =   0  'Transparent
                  Caption         =   "PhilHealth"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   154
                  Top             =   2520
                  Width           =   1095
               End
               Begin VB.Label Label65 
                  BackStyle       =   0  'Transparent
                  Caption         =   "T I N"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   152
                  Top             =   2880
                  Width           =   1095
               End
               Begin VB.Label Label64 
                  BackStyle       =   0  'Transparent
                  Caption         =   "SSS #"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   150
                  Top             =   2520
                  Width           =   1095
               End
               Begin VB.Label Label63 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Contact #"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   148
                  Top             =   1800
                  Width           =   1095
               End
               Begin VB.Label Label62 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Weight"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   146
                  Top             =   2160
                  Width           =   1095
               End
               Begin VB.Label Label61 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Height"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   144
                  Top             =   2160
                  Width           =   1095
               End
               Begin VB.Label Label60 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   142
                  Top             =   1080
                  Width           =   1095
               End
               Begin VB.Label Label58 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Civil Status"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   138
                  Top             =   720
                  Width           =   1095
               End
               Begin VB.Label Label57 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Age"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   136
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.Label Label56 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Gender"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   134
                  Top             =   720
                  Width           =   1095
               End
               Begin VB.Label Label55 
                  BackStyle       =   0  'Transparent
                  Caption         =   "BirthDate"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   132
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.Label Label54 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Name"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   130
                  Top             =   0
                  Width           =   1095
               End
            End
            Begin VB.PictureBox Picture1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   4335
               Left            =   5520
               ScaleHeight     =   4305
               ScaleWidth      =   4305
               TabIndex        =   128
               Top             =   0
               Width           =   4335
               Begin VB.Image imgPicture 
                  Height          =   4305
                  Left            =   0
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   4305
               End
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00404040&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00404040&
               Height          =   4335
               Left            =   5600
               Top             =   80
               Width           =   4335
            End
         End
         Begin VB.PictureBox picEmployment 
            BackColor       =   &H00C6B8A4&
            BorderStyle     =   0  'None
            Height          =   4335
            Left            =   -74880
            ScaleHeight     =   4335
            ScaleWidth      =   9855
            TabIndex        =   104
            Top             =   480
            Width           =   9855
            Begin VB.TextBox txtEmegencyRelation 
               Height          =   315
               Left            =   7440
               TabIndex        =   124
               Top             =   3600
               Width           =   2415
            End
            Begin VB.TextBox txtEmegencyAddress 
               Height          =   315
               Left            =   840
               TabIndex        =   122
               Top             =   3960
               Width           =   5055
            End
            Begin VB.TextBox txtEmegencyContact 
               Height          =   315
               Left            =   7440
               TabIndex        =   120
               Top             =   3960
               Width           =   2415
            End
            Begin VB.TextBox txtEmegencyName 
               Height          =   315
               Left            =   840
               TabIndex        =   118
               Top             =   3600
               Width           =   5055
            End
            Begin VB.TextBox txtRelativeCompanyAddress 
               Height          =   315
               Left            =   7200
               TabIndex        =   114
               Top             =   2880
               Width           =   2655
            End
            Begin VB.TextBox txtRelativeCompanyContact 
               Height          =   315
               Left            =   5880
               TabIndex        =   113
               Top             =   2880
               Width           =   1280
            End
            Begin VB.TextBox txtRelativeCompanyName 
               Height          =   315
               Left            =   3360
               TabIndex        =   112
               Top             =   2880
               Width           =   2480
            End
            Begin VB.TextBox txtNotRelatedAddress 
               Height          =   315
               Left            =   7200
               TabIndex        =   110
               Top             =   2520
               Width           =   2655
            End
            Begin VB.TextBox txtNotRelatedContact 
               Height          =   315
               Left            =   5880
               TabIndex        =   109
               Top             =   2520
               Width           =   1280
            End
            Begin VB.TextBox txtNotRelatedName 
               Height          =   315
               Left            =   3360
               TabIndex        =   108
               Top             =   2520
               Width           =   2480
            End
            Begin MSComctlLib.ListView lstEmployment 
               Height          =   1820
               Left            =   0
               TabIndex        =   106
               Top             =   240
               Width           =   9855
               _ExtentX        =   17383
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
               NumItems        =   7
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
                  Text            =   "Company"
                  Object.Width           =   4410
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Position"
                  Object.Width           =   2646
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   4
                  Text            =   "Salary"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Inclusive Date"
                  Object.Width           =   2646
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "Address"
                  Object.Width           =   5203
               EndProperty
            End
            Begin VB.Label Label52 
               BackStyle       =   0  'Transparent
               Caption         =   "Relation"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   6120
               TabIndex        =   123
               Top             =   3600
               Width           =   1215
            End
            Begin VB.Label Label51 
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   121
               Top             =   3960
               Width           =   855
            End
            Begin VB.Label Label50 
               BackStyle       =   0  'Transparent
               Caption         =   "Contact Number"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   6120
               TabIndex        =   119
               Top             =   3960
               Width           =   1215
            End
            Begin VB.Label Label49 
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   117
               Top             =   3600
               Width           =   855
            End
            Begin VB.Label Label48 
               Alignment       =   2  'Center
               BackColor       =   &H00404040&
               Caption         =   "PERSON TO BE CONTACTED IN CASE OF EMERGENCY"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C6B8A4&
               Height          =   220
               Left            =   0
               TabIndex        =   116
               Top             =   3280
               Width           =   9855
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               BackColor       =   &H00404040&
               Caption         =   "REFERENCES"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C6B8A4&
               Height          =   220
               Left            =   0
               TabIndex        =   115
               Top             =   2158
               Width           =   9855
            End
            Begin VB.Label Label47 
               BackStyle       =   0  'Transparent
               Caption         =   "(Name of relative connected to this company)"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   111
               Top             =   2880
               Width           =   3375
            End
            Begin VB.Label Label46 
               BackStyle       =   0  'Transparent
               Caption         =   "(Name must be not related to you)"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   107
               Top             =   2520
               Width           =   2535
            End
            Begin VB.Label Label44 
               Alignment       =   2  'Center
               BackColor       =   &H00404040&
               Caption         =   "EMPLOYMENT HISTORY"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C6B8A4&
               Height          =   220
               Left            =   0
               TabIndex        =   105
               Top             =   0
               Width           =   9855
            End
         End
         Begin VB.PictureBox picEducation 
            BackColor       =   &H00C6B8A4&
            BorderStyle     =   0  'None
            Height          =   4335
            Left            =   -74880
            ScaleHeight     =   4335
            ScaleWidth      =   9855
            TabIndex        =   73
            Top             =   480
            Width           =   9855
            Begin VB.TextBox txtOrgsClubs 
               Height          =   315
               Left            =   960
               TabIndex        =   103
               Top             =   3960
               Width           =   8890
            End
            Begin MSComctlLib.ListView lstTraining 
               Height          =   1815
               Left            =   960
               TabIndex        =   100
               Top             =   2060
               Width           =   8895
               _ExtentX        =   15690
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
               NumItems        =   5
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
                  Text            =   "Title"
                  Object.Width           =   6174
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Inclusive Date"
                  Object.Width           =   3528
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Venue"
                  Object.Width           =   5468
               EndProperty
            End
            Begin VB.TextBox txtSkills 
               Height          =   315
               Left            =   960
               TabIndex        =   99
               Top             =   1680
               Width           =   8890
            End
            Begin VB.TextBox txtPostName 
               Height          =   315
               Left            =   960
               TabIndex        =   96
               Top             =   1320
               Width           =   2715
            End
            Begin VB.TextBox txtPostInclusiveDate 
               Height          =   315
               Left            =   3720
               TabIndex        =   95
               Text            =   "11/24/1979"
               Top             =   1320
               Width           =   1215
            End
            Begin VB.TextBox txtPostCourse 
               Height          =   315
               Left            =   4980
               TabIndex        =   94
               Top             =   1320
               Width           =   1815
            End
            Begin VB.TextBox txtPostAddress 
               Height          =   315
               Left            =   6840
               TabIndex        =   93
               Top             =   1320
               Width           =   3015
            End
            Begin VB.TextBox txtElemName 
               Height          =   315
               Left            =   960
               TabIndex        =   85
               Top             =   240
               Width           =   2715
            End
            Begin VB.TextBox txtElemInclusiveDate 
               Height          =   315
               Left            =   3720
               TabIndex        =   84
               Text            =   "1986-1987"
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox txtElemCourse 
               Height          =   315
               Left            =   4980
               TabIndex        =   83
               Top             =   240
               Width           =   1815
            End
            Begin VB.TextBox txtElemAddress 
               Height          =   315
               Left            =   6840
               TabIndex        =   82
               Top             =   240
               Width           =   3015
            End
            Begin VB.TextBox txtHiSchoolName 
               Height          =   315
               Left            =   960
               TabIndex        =   81
               Top             =   600
               Width           =   2715
            End
            Begin VB.TextBox txtHiSchoolInclusiveDate 
               Height          =   315
               Left            =   3720
               TabIndex        =   80
               Text            =   "11/24/1979"
               Top             =   600
               Width           =   1215
            End
            Begin VB.TextBox txtHiSchoolCourse 
               Height          =   315
               Left            =   4980
               TabIndex        =   79
               Top             =   600
               Width           =   1815
            End
            Begin VB.TextBox txtHiSchoolAddress 
               Height          =   315
               Left            =   6840
               TabIndex        =   78
               Top             =   600
               Width           =   3015
            End
            Begin VB.TextBox txtCollegeName 
               Height          =   315
               Left            =   960
               TabIndex        =   77
               Top             =   960
               Width           =   2715
            End
            Begin VB.TextBox txtCollegeInclusiveDate 
               Height          =   315
               Left            =   3720
               TabIndex        =   76
               Text            =   "11/24/1979"
               Top             =   960
               Width           =   1215
            End
            Begin VB.TextBox txtCollegeCourse 
               Height          =   315
               Left            =   4980
               TabIndex        =   75
               Top             =   960
               Width           =   1815
            End
            Begin VB.TextBox txtCollegeAddress 
               Height          =   315
               Left            =   6840
               TabIndex        =   74
               Top             =   960
               Width           =   3015
            End
            Begin VB.Label Label43 
               BackStyle       =   0  'Transparent
               Caption         =   "Orgs/Clubs"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   102
               Top             =   3960
               Width           =   1095
            End
            Begin VB.Label Label42 
               BackStyle       =   0  'Transparent
               Caption         =   "Trainings / Seminars"
               ForeColor       =   &H00000000&
               Height          =   735
               Left            =   0
               TabIndex        =   101
               Top             =   2040
               Width           =   855
            End
            Begin VB.Label Label41 
               BackStyle       =   0  'Transparent
               Caption         =   "Skills"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   98
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label Label40 
               BackStyle       =   0  'Transparent
               Caption         =   "Post Studies"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   97
               Top             =   1350
               Width           =   1095
            End
            Begin VB.Label Label39 
               BackStyle       =   0  'Transparent
               Caption         =   "Elementary"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   92
               Top             =   270
               Width           =   1095
            End
            Begin VB.Label Label38 
               BackStyle       =   0  'Transparent
               Caption         =   "School"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   960
               TabIndex        =   91
               Top             =   0
               Width           =   1095
            End
            Begin VB.Label Label37 
               BackStyle       =   0  'Transparent
               Caption         =   "Inclusive Date"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   3720
               TabIndex        =   90
               Top             =   0
               Width           =   1215
            End
            Begin VB.Label Label36 
               BackStyle       =   0  'Transparent
               Caption         =   "Course"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5040
               TabIndex        =   89
               Top             =   0
               Width           =   1095
            End
            Begin VB.Label Label35 
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   6840
               TabIndex        =   88
               Top             =   0
               Width           =   1095
            End
            Begin VB.Label Label34 
               BackStyle       =   0  'Transparent
               Caption         =   "High School"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   87
               Top             =   630
               Width           =   1095
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "College"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   86
               Top             =   990
               Width           =   1095
            End
         End
         Begin VB.PictureBox picFamily 
            BackColor       =   &H00C6B8A4&
            BorderStyle     =   0  'None
            Height          =   4455
            Left            =   120
            ScaleHeight     =   4455
            ScaleWidth      =   9855
            TabIndex        =   49
            Top             =   480
            Width           =   9855
            Begin MSComctlLib.ListView lstBroSis 
               Height          =   1620
               Left            =   1200
               TabIndex        =   69
               Top             =   1320
               Width           =   8655
               _ExtentX        =   15266
               _ExtentY        =   2858
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   0   'False
               HideSelection   =   -1  'True
               HideColumnHeaders=   -1  'True
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
                  Text            =   "Line"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "FullName"
                  Object.Width           =   4674
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "BirthDate"
                  Object.Width           =   1941
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Occupation"
                  Object.Width           =   3175
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Address"
                  Object.Width           =   4939
               EndProperty
            End
            Begin VB.TextBox txtMotherAddress 
               Height          =   315
               Left            =   6840
               TabIndex        =   68
               Top             =   960
               Width           =   3015
            End
            Begin VB.TextBox txtMotherOccupation 
               Height          =   315
               Left            =   4980
               TabIndex        =   67
               Top             =   960
               Width           =   1815
            End
            Begin VB.TextBox txtMotherBDay 
               Height          =   315
               Left            =   3960
               TabIndex        =   66
               Text            =   "11/24/1979"
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox txtMotherName 
               Height          =   315
               Left            =   1200
               TabIndex        =   64
               Top             =   960
               Width           =   2720
            End
            Begin VB.TextBox txtFatherAddress 
               Height          =   315
               Left            =   6840
               TabIndex        =   63
               Top             =   600
               Width           =   3015
            End
            Begin VB.TextBox txtFatherOccupation 
               Height          =   315
               Left            =   4980
               TabIndex        =   62
               Top             =   600
               Width           =   1815
            End
            Begin VB.TextBox txtFatherBDay 
               Height          =   315
               Left            =   3960
               TabIndex        =   61
               Text            =   "11/24/1979"
               Top             =   600
               Width           =   975
            End
            Begin VB.TextBox txtFatherName 
               Height          =   315
               Left            =   1200
               TabIndex        =   59
               Top             =   600
               Width           =   2720
            End
            Begin VB.TextBox txtSpouseAddress 
               Height          =   315
               Left            =   6840
               TabIndex        =   54
               Top             =   240
               Width           =   3015
            End
            Begin VB.TextBox txtSpouseOccupation 
               Height          =   315
               Left            =   4980
               TabIndex        =   53
               Top             =   240
               Width           =   1815
            End
            Begin VB.TextBox txtSpouseBDay 
               Height          =   315
               Left            =   3960
               TabIndex        =   52
               Text            =   "11/24/1979"
               Top             =   240
               Width           =   975
            End
            Begin VB.TextBox txtSpouseName 
               Height          =   315
               Left            =   1200
               TabIndex        =   50
               Top             =   240
               Width           =   2720
            End
            Begin MSComctlLib.ListView lstChildren 
               Height          =   1410
               Left            =   1200
               TabIndex        =   72
               Top             =   3000
               Width           =   8655
               _ExtentX        =   15266
               _ExtentY        =   2487
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   0   'False
               HideSelection   =   -1  'True
               HideColumnHeaders=   -1  'True
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
                  Text            =   "Line"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "FullName"
                  Object.Width           =   4674
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "BirthDate"
                  Object.Width           =   1941
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Occupation"
                  Object.Width           =   3175
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Address"
                  Object.Width           =   4939
               EndProperty
            End
            Begin VB.Label Label32 
               BackStyle       =   0  'Transparent
               Caption         =   "Children's Name"
               ForeColor       =   &H00000000&
               Height          =   615
               Left            =   0
               TabIndex        =   71
               Top             =   3000
               Width           =   1215
            End
            Begin VB.Label Label31 
               BackStyle       =   0  'Transparent
               Caption         =   "Brother's and Sister's Name"
               ForeColor       =   &H00000000&
               Height          =   615
               Left            =   0
               TabIndex        =   70
               Top             =   1320
               Width           =   1095
            End
            Begin VB.Label Label30 
               BackStyle       =   0  'Transparent
               Caption         =   "Mother's Name"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   65
               Top             =   990
               Width           =   1095
            End
            Begin VB.Label Label29 
               BackStyle       =   0  'Transparent
               Caption         =   "Father's Name"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   60
               Top             =   630
               Width           =   1095
            End
            Begin VB.Label Label28 
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   6840
               TabIndex        =   58
               Top             =   0
               Width           =   1095
            End
            Begin VB.Label Label27 
               BackStyle       =   0  'Transparent
               Caption         =   "Occupation"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5040
               TabIndex        =   57
               Top             =   0
               Width           =   1095
            End
            Begin VB.Label Label26 
               BackStyle       =   0  'Transparent
               Caption         =   "Birth Date"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   3960
               TabIndex        =   56
               Top             =   0
               Width           =   975
            End
            Begin VB.Label Label25 
               BackStyle       =   0  'Transparent
               Caption         =   "Full Name"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   1320
               TabIndex        =   55
               Top             =   0
               Width           =   1095
            End
            Begin VB.Label Label24 
               BackStyle       =   0  'Transparent
               Caption         =   "Spouse Name"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   51
               Top             =   270
               Width           =   1095
            End
         End
         Begin VB.PictureBox picPersonnal 
            BackColor       =   &H00C6B8A4&
            BorderStyle     =   0  'None
            Height          =   4575
            Left            =   -74880
            ScaleHeight     =   4575
            ScaleWidth      =   9855
            TabIndex        =   3
            Top             =   480
            Width           =   9855
            Begin VB.TextBox txtBloodType 
               Height          =   315
               Left            =   6480
               TabIndex        =   236
               Top             =   2040
               Width           =   3375
            End
            Begin VB.ComboBox cmbTaxStatus 
               Height          =   315
               Left            =   6480
               Style           =   2  'Dropdown List
               TabIndex        =   215
               Top             =   3840
               Width           =   1335
            End
            Begin VB.TextBox txtNoDependent 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   9240
               TabIndex        =   213
               Top             =   3840
               Width           =   615
            End
            Begin VB.TextBox txtContactNumber 
               Height          =   315
               Left            =   6480
               TabIndex        =   140
               Top             =   2760
               Width           =   3375
            End
            Begin VB.ComboBox cmbGender 
               Height          =   315
               Left            =   1440
               Style           =   2  'Dropdown List
               TabIndex        =   125
               Top             =   2400
               Width           =   3495
            End
            Begin VB.ComboBox cmbIDNumber 
               Height          =   315
               Left            =   6480
               Style           =   2  'Dropdown List
               TabIndex        =   48
               Top             =   4200
               Width           =   3375
            End
            Begin VB.TextBox txtPagIbigNumber 
               Height          =   315
               Left            =   1440
               TabIndex        =   46
               Top             =   4200
               Width           =   3495
            End
            Begin VB.TextBox txtPHICNumber 
               Height          =   315
               Left            =   1440
               TabIndex        =   44
               Top             =   3840
               Width           =   3495
            End
            Begin VB.TextBox txtDriverLicense 
               Height          =   315
               Left            =   6480
               TabIndex        =   42
               Top             =   3120
               Width           =   3375
            End
            Begin VB.TextBox txtTin 
               Height          =   315
               Left            =   6480
               TabIndex        =   40
               Top             =   3480
               Width           =   3375
            End
            Begin VB.TextBox txtSSSNumber 
               Height          =   315
               Left            =   1440
               TabIndex        =   38
               Top             =   3480
               Width           =   3495
            End
            Begin VB.TextBox txtNationality 
               Height          =   315
               Left            =   6480
               TabIndex        =   36
               Top             =   2400
               Width           =   3375
            End
            Begin VB.TextBox txtWeight 
               Height          =   315
               Left            =   8640
               TabIndex        =   34
               Top             =   1680
               Width           =   1215
            End
            Begin VB.TextBox txtHeight 
               Height          =   315
               Left            =   6480
               TabIndex        =   32
               Top             =   1680
               Width           =   1215
            End
            Begin VB.TextBox txtDateMarriage 
               Height          =   315
               Left            =   6480
               TabIndex        =   30
               Top             =   1320
               Width           =   3375
            End
            Begin VB.ComboBox cmbCivilStatus 
               Height          =   315
               Left            =   6480
               Style           =   2  'Dropdown List
               TabIndex        =   28
               Top             =   960
               Width           =   3375
            End
            Begin VB.ComboBox cmbLivingParents 
               Height          =   315
               Left            =   1440
               Style           =   2  'Dropdown List
               TabIndex        =   27
               Top             =   1320
               Width           =   3495
            End
            Begin VB.TextBox txtReligion 
               Height          =   315
               Left            =   1440
               TabIndex        =   24
               Top             =   3120
               Width           =   3495
            End
            Begin VB.TextBox txtBirthPlace 
               Height          =   315
               Left            =   1440
               TabIndex        =   22
               Top             =   2760
               Width           =   3495
            End
            Begin VB.TextBox txtAge 
               Height          =   315
               Left            =   3840
               TabIndex        =   20
               Top             =   2040
               Width           =   1095
            End
            Begin VB.ComboBox cmbRent 
               Height          =   315
               Left            =   1440
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   1680
               Width           =   3495
            End
            Begin VB.ComboBox cmbOwnedHouse 
               Height          =   315
               Left            =   1440
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   960
               Width           =   3495
            End
            Begin VB.TextBox txtBirthDate 
               Height          =   315
               Left            =   1440
               TabIndex        =   16
               Top             =   2040
               Width           =   1695
            End
            Begin VB.TextBox txtPresentAddress 
               Height          =   315
               Left            =   1440
               TabIndex        =   8
               Top             =   600
               Width           =   8415
            End
            Begin VB.TextBox txtMiddleName 
               Height          =   315
               Left            =   7200
               MaxLength       =   100
               TabIndex        =   7
               Top             =   0
               Width           =   2655
            End
            Begin VB.TextBox txtFirstName 
               Height          =   315
               Left            =   4320
               MaxLength       =   100
               TabIndex        =   6
               Top             =   0
               Width           =   2775
            End
            Begin VB.TextBox txtLastName 
               Height          =   315
               Left            =   1440
               MaxLength       =   100
               TabIndex        =   4
               Top             =   0
               Width           =   2775
            End
            Begin VB.Label Label90 
               BackStyle       =   0  'Transparent
               Caption         =   "Blood Type"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5040
               TabIndex        =   235
               Top             =   2040
               Width           =   1575
            End
            Begin VB.Label Label86 
               BackStyle       =   0  'Transparent
               Caption         =   "# Of Dependent"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   7920
               TabIndex        =   216
               Top             =   3885
               Width           =   1215
            End
            Begin VB.Label Label85 
               BackStyle       =   0  'Transparent
               Caption         =   "Tax Status"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5040
               TabIndex        =   214
               Top             =   3840
               Width           =   1575
            End
            Begin VB.Label Label59 
               BackStyle       =   0  'Transparent
               Caption         =   "Contact Number"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5040
               TabIndex        =   141
               Top             =   2760
               Width           =   1575
            End
            Begin VB.Label Label53 
               BackStyle       =   0  'Transparent
               Caption         =   "Gender"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   126
               Top             =   2400
               Width           =   1575
            End
            Begin VB.Label Label23 
               BackStyle       =   0  'Transparent
               Caption         =   "ID Number"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5040
               TabIndex        =   47
               Top             =   4200
               Width           =   1575
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "Pag Ibig #"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   45
               Top             =   4200
               Width           =   1575
            End
            Begin VB.Label Label21 
               BackStyle       =   0  'Transparent
               Caption         =   "Phil Health #"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   43
               Top             =   3840
               Width           =   1575
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "Driver's License"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5040
               TabIndex        =   41
               Top             =   3120
               Width           =   1575
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "T I N"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5040
               TabIndex        =   39
               Top             =   3480
               Width           =   1575
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "SSS #"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   37
               Top             =   3480
               Width           =   1575
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "Nationality"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5040
               TabIndex        =   35
               Top             =   2400
               Width           =   1575
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Weight (kg)"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   7800
               TabIndex        =   33
               Top             =   1725
               Width           =   855
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Height (inches)"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5040
               TabIndex        =   31
               Top             =   1680
               Width           =   1575
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Date of Marriage"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5040
               TabIndex        =   29
               Top             =   1320
               Width           =   1575
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Civil Status"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5040
               TabIndex        =   26
               Top             =   960
               Width           =   1575
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Living w/ Parents?"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   25
               Top             =   1320
               Width           =   1575
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Religion"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   23
               Top             =   3120
               Width           =   1575
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Birth Place"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   21
               Top             =   2760
               Width           =   1575
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Age"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   3360
               TabIndex        =   19
               Top             =   2040
               Width           =   375
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Birth Date "
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   15
               Top             =   2040
               Width           =   1575
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Rent?"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   14
               Top             =   1680
               Width           =   1575
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Owned House?"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   13
               Top             =   960
               Width           =   1575
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Middle Name"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   7200
               TabIndex        =   12
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Given Name"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   4320
               TabIndex        =   11
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Last Name"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   1440
               TabIndex        =   10
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Present Address"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   9
               Top             =   600
               Width           =   1575
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   5
               Top             =   30
               Width           =   1095
            End
         End
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   239
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   240
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
         MouseIcon       =   "frmPersonnelInformation.frx":0CCA
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11520
      Top             =   1680
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
            Picture         =   "frmPersonnelInformation.frx":0FE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInformation.frx":1CBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInformation.frx":2998
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInformation.frx":3672
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInformation.frx":434C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInformation.frx":5026
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInformation.frx":5D00
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInformation.frx":69DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInformation.frx":76B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInformation.frx":7F8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInformation.frx":8C68
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInformation.frx":9942
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInformation.frx":A61C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInformation.frx":B2F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInformation.frx":BFD0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RPVGCC.b8Container picProgressReport 
      Height          =   975
      Left            =   2160
      TabIndex        =   227
      Top             =   3120
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1720
      BackColor       =   15266266
      Begin VB.Timer TimerProfile 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   3240
         Top             =   720
      End
      Begin VB.Timer TimerHistory 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   3720
         Top             =   720
      End
      Begin VB.Timer TimerAlphaActive 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   4200
         Top             =   720
      End
      Begin VB.Timer TimerHeadCount 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   4680
         Top             =   720
      End
      Begin VB.Timer TimerInactive 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   5160
         Top             =   720
      End
      Begin VB.Timer TimerActive 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   5640
         Top             =   720
      End
      Begin VB.PictureBox picProgress 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5955
         TabIndex        =   228
         Top             =   120
         Width           =   6015
      End
   End
   Begin RPVGCC.b8Container picAlphalist 
      Height          =   1815
      Left            =   3360
      TabIndex        =   229
      Top             =   2640
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3201
      BackColor       =   15396057
      Begin VB.TextBox txtAsof 
         Height          =   315
         Left            =   1320
         TabIndex        =   232
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancelAlphalist 
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
         Left            =   1920
         Picture         =   "frmPersonnelInformation.frx":CCAA
         Style           =   1  'Graphical
         TabIndex        =   231
         Top             =   1080
         Width           =   1560
      End
      Begin VB.CommandButton cmdOKAlphalist 
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
         Left            =   240
         Picture         =   "frmPersonnelInformation.frx":D406
         Style           =   1  'Graphical
         TabIndex        =   230
         Top             =   1080
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   45
         TabIndex        =   233
         Top             =   45
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   609
         Caption         =   "Alpha List"
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
         Icon            =   "frmPersonnelInformation.frx":DA78
         ShadowVisible   =   0   'False
      End
      Begin VB.Label Label89 
         BackStyle       =   0  'Transparent
         Caption         =   "As of"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   234
         Top             =   600
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   6435
      Width           =   10365
      _ExtentX        =   18283
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
   Begin RPVGCC.b8Container picSearch 
      Height          =   4095
      Left            =   3120
      TabIndex        =   221
      Top             =   1320
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7223
      BackColor       =   15396057
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
         Picture         =   "frmPersonnelInformation.frx":E012
         Style           =   1  'Graphical
         TabIndex        =   225
         Top             =   3480
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
         Left            =   2280
         Picture         =   "frmPersonnelInformation.frx":E684
         Style           =   1  'Graphical
         TabIndex        =   224
         Top             =   3480
         Width           =   1560
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   223
         Top             =   480
         Width           =   4215
      End
      Begin VB.ListBox lstResult 
         Height          =   2595
         Left            =   120
         TabIndex        =   222
         Top             =   840
         Width           =   4215
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   226
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
         Icon            =   "frmPersonnelInformation.frx":EDE0
         ShadowVisible   =   0   'False
      End
   End
   Begin RPVGCC.b8Container picChildrenSLine 
      Height          =   855
      Left            =   1320
      TabIndex        =   170
      Top             =   3840
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox txtChildAddress1 
         Height          =   315
         Left            =   2520
         TabIndex        =   186
         Top             =   0
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.TextBox txtChildOccupation1 
         Height          =   315
         Left            =   2400
         TabIndex        =   185
         Top             =   0
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.TextBox txtChildBDay1 
         Height          =   315
         Left            =   2280
         TabIndex        =   184
         Top             =   0
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.TextBox txtChildName1 
         Height          =   315
         Left            =   2160
         TabIndex        =   183
         Top             =   0
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.TextBox txtChildAddress 
         Height          =   315
         Left            =   5760
         TabIndex        =   174
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtChildOccupation 
         Height          =   315
         Left            =   3900
         TabIndex        =   173
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtChildBDay 
         Height          =   315
         Left            =   2880
         TabIndex        =   172
         Text            =   "11/24/1979"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtChildName 
         Height          =   315
         Left            =   120
         TabIndex        =   171
         Top             =   360
         Width           =   2720
      End
      Begin VB.Label Label76 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5760
         TabIndex        =   178
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label75 
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3960
         TabIndex        =   177
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label74 
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2880
         TabIndex        =   176
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label73 
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   175
         Top             =   120
         Width           =   1095
      End
   End
   Begin RPVGCC.b8Container picTrainingsSLine 
      Height          =   855
      Left            =   1080
      TabIndex        =   187
      Top             =   2880
      Visible         =   0   'False
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox txtTrainingVenue1 
         Height          =   315
         Left            =   1920
         TabIndex        =   196
         Top             =   0
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.TextBox txtTrainingDates1 
         Height          =   315
         Left            =   1800
         TabIndex        =   195
         Top             =   0
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.TextBox txtTrainingTitle1 
         Height          =   315
         Left            =   1680
         TabIndex        =   194
         Top             =   0
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.TextBox txtTrainingVenue 
         Height          =   315
         Left            =   5640
         TabIndex        =   190
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox txtTrainingDates 
         Height          =   315
         Left            =   3600
         TabIndex        =   189
         Text            =   "1986-1987"
         Top             =   360
         Width           =   1980
      End
      Begin VB.TextBox txtTrainingTitle 
         Height          =   315
         Left            =   120
         TabIndex        =   188
         Top             =   360
         Width           =   3435
      End
      Begin VB.Label Label79 
         BackStyle       =   0  'Transparent
         Caption         =   "Venue"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5640
         TabIndex        =   193
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label78 
         BackStyle       =   0  'Transparent
         Caption         =   "Inclusive Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   192
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label77 
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   191
         Top             =   120
         Width           =   1095
      End
   End
   Begin RPVGCC.b8Container picBrotherSisterSLine 
      Height          =   855
      Left            =   1320
      TabIndex        =   161
      Top             =   2160
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox txtBroSisAddress1 
         Height          =   315
         Left            =   2520
         TabIndex        =   182
         Top             =   0
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.TextBox txtBroSisOccupation1 
         Height          =   315
         Left            =   2400
         TabIndex        =   181
         Top             =   0
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.TextBox txtBroSisBDay1 
         Height          =   315
         Left            =   2280
         TabIndex        =   180
         Top             =   0
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.TextBox txtBroSisName1 
         Height          =   315
         Left            =   2160
         TabIndex        =   179
         Top             =   0
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.TextBox txtBroSisName 
         Height          =   315
         Left            =   120
         TabIndex        =   165
         Top             =   360
         Width           =   2720
      End
      Begin VB.TextBox txtBroSisBDay 
         Height          =   315
         Left            =   2880
         TabIndex        =   164
         Text            =   "11/24/1979"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtBroSisOccupation 
         Height          =   315
         Left            =   3900
         TabIndex        =   163
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtBroSisAddress 
         Height          =   315
         Left            =   5760
         TabIndex        =   162
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label72 
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   169
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label71 
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2880
         TabIndex        =   168
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label70 
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3960
         TabIndex        =   167
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label69 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5760
         TabIndex        =   166
         Top             =   120
         Width           =   1095
      End
   End
   Begin RPVGCC.b8Container picEmploymentSLine 
      Height          =   855
      Left            =   120
      TabIndex        =   197
      Top             =   1080
      Visible         =   0   'False
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox txtEmploymentAddress1 
         Height          =   315
         Left            =   2040
         TabIndex        =   212
         Top             =   0
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.TextBox txtEmploymentDates1 
         Height          =   315
         Left            =   1920
         TabIndex        =   211
         Top             =   0
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.TextBox txtEmploymentSalary1 
         Height          =   315
         Left            =   1800
         TabIndex        =   210
         Top             =   0
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.TextBox txtEmploymentPosition1 
         Height          =   315
         Left            =   1680
         TabIndex        =   209
         Top             =   0
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.TextBox txtEmploymentCompany1 
         Height          =   315
         Left            =   1560
         TabIndex        =   208
         Top             =   0
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.TextBox txtEmploymentDates 
         Height          =   315
         Left            =   5280
         TabIndex        =   202
         Top             =   360
         Width           =   1515
      End
      Begin VB.TextBox txtEmploymentAddress 
         Height          =   315
         Left            =   6840
         TabIndex        =   201
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox txtEmploymentCompany 
         Height          =   315
         Left            =   120
         TabIndex        =   200
         Top             =   360
         Width           =   2475
      End
      Begin VB.TextBox txtEmploymentPosition 
         Height          =   315
         Left            =   2640
         TabIndex        =   199
         Top             =   360
         Width           =   1515
      End
      Begin VB.TextBox txtEmploymentSalary 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4200
         TabIndex        =   198
         Top             =   360
         Width           =   1030
      End
      Begin VB.Label Label84 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6840
         TabIndex        =   207
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label83 
         BackStyle       =   0  'Transparent
         Caption         =   "Inclusive Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5280
         TabIndex        =   206
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label82 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Salary"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4200
         TabIndex        =   205
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label81 
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2640
         TabIndex        =   204
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label80 
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   203
         Top             =   120
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmPersonnelInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ListFocus   As Long
Dim ListRow     As Long

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Public SearchType As Long
Const is_LName = 10
Const is_FName = 11
Const is_MName = 12

Dim ListTrans As Long
Const is_LstRefresh = 0
Const is_LstAdding = 1
Const is_LstEditting = 2

Dim Filename, FileNameTmp, x

Dim FileName_xls As String
Dim tmp As Long

Dim i, TableName, Columns, j, k, RowCnt, ColCnt, strRange
Dim WorkbookName As String
Dim iWorkSheet As Integer
Dim ProfilePK, strLastName, strFirstName, strMiddleName, strPresentAddress, intOwnedHouse, intRented, intGender, intLivingWParents, intCivilStatus, dtmBirthDate, _
strBirthPlace, strReligion, dblHeight, dblWeight, strNationality, strContactNumber, strSSSNumber, strPHICNumber, strHDMFNumber, strTIN, strDriverLicense, _
strSpouseName, strSpouseOccupation, strSpouseAddress, strFatherName, strFatherOccupation, strFatherAddress, strMotherName, strMotherOccupation, strMotherAddress, _
strSkills, strOrganizationClubs, strRefName, strRefContact, strRefAddress, strRefCompName, strRefCompContact, strRefCompAddress, strEmergencyName, _
strEmergencyRelation, strEmergencyAddress, strEmergencyContact, strLastModified, strFullName

Private Function BROWSER(strFullName, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If strFullName <> "" Then
            s = "SELECT TOP 1 tbl_Personnel_Information.* " & _
                " FROM tbl_Personnel_Information " & _
                " WHERE (LastName + ',  ' + FirstName + '  ' + MiddleName = '" & FORMATSQL(CStr(strFullName)) & "') " & _
                " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
        Else
            s = "SELECT TOP 1 tbl_Personnel_Information.* " & _
                " FROM tbl_Personnel_Information " & _
                " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
        End If
    Case "is_FIND"
        s = "SELECT TOP 1 tbl_Personnel_Information.* " & _
            " FROM tbl_Personnel_Information " & _
            " WHERE (PK = " & strFullName & ") " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        If picAlphalist.Visible = True Then Exit Function
        If picSearch.Visible = True Then Exit Function
        s = "SELECT TOP 1 tbl_Personnel_Information.* " & _
            " FROM tbl_Personnel_Information " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        If picAlphalist.Visible = True Then Exit Function
        If picSearch.Visible = True Then Exit Function
        s = "SELECT TOP 1 tbl_Personnel_Information.* " & _
            " FROM tbl_Personnel_Information " & _
            " WHERE (LastName + ',  ' + FirstName + '  ' + MiddleName < '" & FORMATSQL(CStr(strFullName)) & "') " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        If picAlphalist.Visible = True Then Exit Function
        If picSearch.Visible = True Then Exit Function
        s = "SELECT TOP 1 tbl_Personnel_Information.* " & _
            " FROM tbl_Personnel_Information " & _
            " WHERE (LastName + ',  ' + FirstName + '  ' + MiddleName > '" & FORMATSQL(CStr(strFullName)) & "') " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        If picAlphalist.Visible = True Then Exit Function
        If picSearch.Visible = True Then Exit Function
        s = "SELECT TOP 1 tbl_Personnel_Information.* " & _
            " FROM tbl_Personnel_Information " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName DESC"
    Case Else: Exit Function
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtLastName.Text = rs!LastName
    txtFirstName.Text = rs!FirstName
    txtMiddleName.Text = rs!MiddleName
    txtPresentAddress.Text = rs!PresentAddress
    txtBirthDate.Text = Format(rs!BirthDate, "mm/dd/yyyy")
    txtAge.Text = Get_Age(FormatDateTime(rs!BirthDate, vbShortDate), FormatDateTime(Date, vbShortDate))
    txtBirthPlace.Text = rs!BirthPlace
    txtReligion.Text = rs!Religion
    
    If IsNull(rs!DateMarriage) = True Then
        txtDateMarriage.Text = ""
    Else
        If DateValue(rs!DateMarriage) = DateValue(CDate("01/01/1900")) Then
            txtDateMarriage.Text = ""
        Else
            txtDateMarriage.Text = Format(rs!DateMarriage, "mm/dd/yyyy")
        End If
    End If
    
    txtHeight.Text = IIf(CDbl(rs!Height) = 0, "", rs!Height)
    txtWeight.Text = IIf(CDbl(rs!Weight) = 0, "", rs!Weight)
    txtBloodType.Text = rs!BloodType
    txtNationality.Text = rs!Nationality
    txtContactNumber.Text = rs!ContactNumber
    txtSSSNumber.Text = rs!SSSNumber
    txtTIN.Text = rs!TIN
    cmbTaxStatus.ListIndex = rs!TaxStatus - 1
    txtNoDependent.Text = rs!NoDependent
    txtDriverLicense.Text = rs!DriverLicense
    txtPHICNumber.Text = rs!PHICNumber
    txtPagIbigNumber.Text = rs!HDMFNumber
    txtSpouseName.Text = rs!SpouseName
    
    If IsNull(rs!SpouseBDay) = True Then
        txtSpouseBDay.Text = ""
    Else
        If DateValue(rs!SpouseBDay) = DateValue(CDate("01/01/1900")) Then
            txtSpouseBDay.Text = ""
        Else
            txtSpouseBDay.Text = Format(rs!SpouseBDay, "mm/dd/yyyy")
        End If
    End If
    
    txtSpouseOccupation.Text = rs!SpouseOccupation
    txtSpouseAddress.Text = rs!SpouseAddress
    txtFatherName.Text = rs!FatherName
    
    If IsNull(rs!FatherBDay) = True Then
        txtFatherBDay.Text = ""
    Else
        If DateValue(rs!FatherBDay) = DateValue(CDate("01/01/1900")) Then
            txtFatherBDay.Text = ""
        Else
            txtFatherBDay.Text = Format(rs!FatherBDay, "mm/dd/yyyy")
        End If
    End If
    
    txtFatherOccupation.Text = rs!FatherOccupation
    txtFatherAddress.Text = rs!FatherAddress
    txtMotherName.Text = rs!MotherName
    
    If IsNull(rs!MotherBDay) = True Then
        txtMotherBDay.Text = ""
    Else
        If DateValue(rs!MotherBDay) = DateValue(CDate("01/01/1900")) Then
            txtMotherBDay.Text = ""
        Else
            txtMotherBDay.Text = Format(rs!MotherBDay, "mm/dd/yyyy")
        End If
    End If
    
    txtMotherOccupation.Text = rs!MotherOccupation
    txtMotherAddress.Text = rs!MotherAddress
    
    txtElemName.Text = ""
    txtElemInclusiveDate.Text = ""
    txtElemCourse.Text = ""
    txtElemAddress.Text = ""
    txtHiSchoolName.Text = ""
    txtHiSchoolInclusiveDate.Text = ""
    txtHiSchoolCourse.Text = ""
    txtHiSchoolAddress.Text = ""
    txtCollegeName.Text = ""
    txtCollegeInclusiveDate.Text = ""
    txtCollegeCourse.Text = ""
    txtCollegeAddress.Text = ""
    txtPostName.Text = ""
    txtPostInclusiveDate.Text = ""
    txtPostCourse.Text = ""
    txtPostAddress.Text = ""
    
    t = "SELECT tbl_Personnel_Education.* " & _
        " FROM tbl_Personnel_Education " & _
        " WHERE (ProfileKey = " & rs!PK & ") " & _
        " ORDER BY Line "
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        Select Case rt!line
            Case 1
                txtElemName.Text = rt!SchoolName
                txtElemInclusiveDate.Text = rt!InclusiveDate
                txtElemCourse.Text = rt!Course
                txtElemAddress.Text = rt!Address
            Case 2
                txtHiSchoolName.Text = rt!SchoolName
                txtHiSchoolInclusiveDate.Text = rt!InclusiveDate
                txtHiSchoolCourse.Text = rt!Course
                txtHiSchoolAddress.Text = rt!Address
            Case 3
                txtCollegeName.Text = rt!SchoolName
                txtCollegeInclusiveDate.Text = rt!InclusiveDate
                txtCollegeCourse.Text = rt!Course
                txtCollegeAddress.Text = rt!Address
            Case 4
                txtPostName.Text = rt!SchoolName
                txtPostInclusiveDate.Text = rt!InclusiveDate
                txtPostCourse.Text = rt!Course
                txtPostAddress.Text = rt!Address
        End Select
        rt.MoveNext
    Wend
    rt.Close
    
    txtSkills.Text = rs!Skills
    txtOrgsClubs.Text = rs!OrganizationClubs
    txtNotRelatedName.Text = rs!RefName
    txtNotRelatedContact.Text = rs!RefContact
    txtNotRelatedAddress.Text = rs!RefAddress
    txtRelativeCompanyName.Text = rs!RefCompName
    txtRelativeCompanyContact.Text = rs!RefCompContact
    txtRelativeCompanyAddress.Text = rs!RefCompAddress
    txtEmegencyName.Text = rs!EmergencyName
    txtEmegencyRelation.Text = rs!EmergencyRelation
    txtEmegencyAddress.Text = rs!EmergencyAddress
    txtEmegencyContact.Text = rs!EmergencyContact
    
    cmbGender.ListIndex = rs!Gender - 1
    cmbOwnedHouse.ListIndex = rs!OwnedHouse - 1
    cmbRent.ListIndex = rs!Rented - 1
    cmbLivingParents.ListIndex = rs!LivingWParents - 1
    cmbCivilStatus.ListIndex = rs!CivilStatus - 1
    
    cmbIDNumber.Clear
    t = "SELECT IDNumber " & _
        " FROM tbl_Personnel_IDNumber " & _
        " WHERE (ProfileKey = " & rs!PK & ") " & _
        " ORDER BY PK DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        cmbIDNumber.AddItem rt!IDNumber
        rt.MoveNext
    Wend
    rt.Close
    If cmbIDNumber.ListCount Then cmbIDNumber.ListIndex = 0
    
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", "Last Modified : " & rs!LastModified)
    
    txtTmpFullName.Text = rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
    txtTmpBDay.Text = Format(rs!BirthDate, "mm/dd/yyyy")
    txtTmpAge.Text = Get_Age(FormatDateTime(rs!BirthDate, vbShortDate), FormatDateTime(Date, vbShortDate))
    txtTmpGender.Text = IIf(rs!Gender = 0, "MALE", "FEMALE")
    txtTmpCivilStatus.Text = IIf(rs!CivilStatus = 1, "SINGLE", IIf(rs!CivilStatus = 2, "MARRIED", IIf(rs!CivilStatus = 3, "WIDOWED", IIf(rs!CivilStatus = 4, "WIDOWER", ""))))
    txtTmpAddress.Text = rs!PresentAddress
    txtTmpContact.Text = rs!ContactNumber
    txtTmpHeight.Text = IIf(CDbl(rs!Height) = 0, "", rs!Height)
    txtTmpWeight.Text = IIf(CDbl(rs!Weight) = 0, "", rs!Weight)
    txtTmpSSS.Text = rs!SSSNumber
    txtTmpPHIC.Text = rs!PHICNumber
    txtTmpPagibig.Text = rs!HDMFNumber
    txtTmpTIN.Text = rs!TIN
    txtTmpDriverLic.Text = rs!DriverLicense
    txtTmpBloodType.Text = rs!BloodType
    txtTmpIDNumber.Text = ""
    txtTmpStatus.Text = ""
    
    t = "SELECT TOP 1 IDNumber, " & _
        " ISNULL ((SELECT TOP 1 tbl_Personnel_EmploymentStatus.Active " & _
        " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
        " tbl_Personnel_EmploymentStatus ON tbl_Personnel_Action.EmpStatus = tbl_Personnel_EmploymentStatus.PK " & _
        " WHERE (tbl_Personnel_Action.EmpPK = tbl_Personnel_IDNumber.PK) " & _
        " AND (tbl_Personnel_Action.EffectivityDate <= '" & FormatDateTime(Date, vbShortDate) & "') " & _
        " ORDER BY tbl_Personnel_Action.EffectivityDate DESC), 0) AS Status " & _
        " From tbl_Personnel_IDNumber " & _
        " Where (ProfileKey = " & rs!PK & ") " & _
        " ORDER BY PK DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtTmpIDNumber.Text = rt!IDNumber
        txtTmpStatus.Text = IIf(rt!Status = 1, "ACTIVE", "INACTIVE")
    End If
    rt.Close
    
    imgPicture.Picture = LoadPicture(SHOW_IMAGES(rs!PK, 0, "Employee Profile"))
    
    lstBroSis.ListItems.Clear
    Set x = lstBroSis.ListItems.Add()
    x.Text = "": x.SubItems(1) = "": x.SubItems(2) = " ": x.SubItems(3) = " ": x.SubItems(4) = " ": x.SubItems(5) = " "
    t = "SELECT Line, BrotherSisterName, BrotherSisterBDay, " & _
        " BrotherSisterOccupation, BrotherSisterAddress " & _
        " FROM tbl_Personnel_BrotherSister " & _
        " WHERE (ProfileKey = " & rs!PK & ") " & _
        " ORDER BY Line"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstBroSis.ListItems.Clear
        While Not rt.EOF
            Set x = lstBroSis.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rt!line
            x.SubItems(2) = rt!BrotherSisterName
            If IsNull(rt!BrotherSisterBDay) = True Then
                x.SubItems(3) = " "
            Else
                If DateValue(rt!BrotherSisterBDay) = DateValue(CDate("01/01/1900")) Then
                    x.SubItems(3) = " "
                Else
                    x.SubItems(3) = Format(rt!BrotherSisterBDay, "mm/dd/yyyy")
                End If
            End If
            x.SubItems(4) = rt!BrotherSisterOccupation
            x.SubItems(5) = rt!BrotherSisterAddress
            rt.MoveNext
        Wend
    End If
    rt.Close
    
    lstChildren.ListItems.Clear
    Set x = lstChildren.ListItems.Add()
    x.Text = "": x.SubItems(1) = "": x.SubItems(2) = " ": x.SubItems(3) = " ": x.SubItems(4) = " ": x.SubItems(5) = " "
    t = "SELECT Line, ChildName, ChildBDay, " & _
        " ChildOccupation, ChildAddress " & _
        " FROM tbl_Personnel_Children " & _
        " WHERE (ProfileKey = " & rs!PK & ") " & _
        " ORDER BY Line"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstChildren.ListItems.Clear
        While Not rt.EOF
            Set x = lstChildren.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rt!line
            x.SubItems(2) = rt!ChildName
            If IsNull(rt!ChildBDay) = True Then
                x.SubItems(3) = " "
            Else
                If DateValue(rt!ChildBDay) = DateValue(CDate("01/01/1900")) Then
                    x.SubItems(3) = " "
                Else
                    x.SubItems(3) = Format(rt!ChildBDay, "mm/dd/yyyy")
                End If
            End If
            x.SubItems(4) = rt!ChildOccupation
            x.SubItems(5) = rt!ChildAddress
            rt.MoveNext
        Wend
    End If
    rt.Close
    
    lstTraining.ListItems.Clear
    Set x = lstTraining.ListItems.Add()
    x.Text = "": x.SubItems(1) = "": x.SubItems(2) = " ": x.SubItems(3) = " ": x.SubItems(4) = " "
    t = "SELECT Line, Title, InclusiveDate, " & _
        " Venue" & _
        " FROM tbl_Personnel_Training " & _
        " WHERE (ProfileKey = " & rs!PK & ") " & _
        " ORDER BY Line"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstTraining.ListItems.Clear
        While Not rt.EOF
            Set x = lstTraining.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rt!line
            x.SubItems(2) = rt!Title
            x.SubItems(3) = rt!InclusiveDate
            x.SubItems(4) = rt!Venue
            rt.MoveNext
        Wend
    End If
    rt.Close
    
    lstEmployment.ListItems.Clear
    Set x = lstEmployment.ListItems.Add()
    x.Text = "": x.SubItems(1) = "": x.SubItems(2) = " ": x.SubItems(3) = " ": x.SubItems(4) = " ": x.SubItems(5) = " ": x.SubItems(6) = " "
    t = "SELECT Line, Company, Positions, " & _
        " Salary, InclusiveDate, Address " & _
        " FROM tbl_Personnel_Employment " & _
        " WHERE (ProfileKey = " & rs!PK & ") " & _
        " ORDER BY Line"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstEmployment.ListItems.Clear
        While Not rt.EOF
            Set x = lstEmployment.ListItems.Add()
            x.Text = ""
            x.SubItems(1) = rt!line
            x.SubItems(2) = rt!Company
            x.SubItems(3) = rt!Positions
            x.SubItems(4) = Format(rt!Salary, "#,##0.00")
            x.SubItems(5) = rt!InclusiveDate
            x.SubItems(6) = rt!Address
            rt.MoveNext
        Wend
    End If
    rt.Close
    
    SaveSetting App.EXEName, "ProfileInformation", "ProfileInfo", rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
        
End If
rs.Close
End Function


Private Function PRESS_INSERT()
If TRANSACTIONTYPE = is_REFRESH Then
    If AccessRights("Personnel Information", "Add") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Function
    End If
    CLEARTEXT
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_ADDING
    XTab1.ActiveTab = 0
    txtLastName.SetFocus
ElseIf TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If picBrotherSisterSLine.Visible = True Then Exit Function
    If picChildrenSLine.Visible = True Then Exit Function
    If picTrainingsSLine.Visible = True Then Exit Function
    If picEmploymentSLine.Visible = True Then Exit Function
    Select Case ListFocus
        Case 1
            picBrotherSisterSLine.ZOrder 0
            txtBroSisName.Text = "": txtBroSisBDay.Text = ""
            txtBroSisOccupation.Text = "": txtBroSisAddress.Text = ""
            With lstBroSis.ListItems
                If Trim(.Item(1).SubItems(2)) <> "" Then
                    Set x = .Add()
                    x.Text = ""
                    x.SubItems(1) = " "
                    x.SubItems(2) = " "
                    x.SubItems(3) = " "
                    x.SubItems(4) = " "
                    x.SubItems(5) = " "
                    ListRow = .Count
                End If
            End With
            lstBroSis.ListItems(ListRow).EnsureVisible
            lstBroSis.ListItems(ListRow).Selected = True
            picToolbar.Enabled = False
            picMain.Enabled = False
            picBrotherSisterSLine.Visible = True
            ListTrans = is_LstAdding
            txtBroSisName.SetFocus
        Case 2
            picChildrenSLine.ZOrder 0
            txtChildName.Text = "": txtChildBDay.Text = ""
            txtChildOccupation.Text = "": txtChildAddress.Text = ""
            With lstChildren.ListItems
                If Trim(.Item(1).SubItems(2)) <> "" Then
                    Set x = .Add()
                    x.Text = ""
                    x.SubItems(1) = " "
                    x.SubItems(2) = " "
                    x.SubItems(3) = " "
                    x.SubItems(4) = " "
                    x.SubItems(5) = " "
                    ListRow = .Count
                End If
            End With
            lstChildren.ListItems(ListRow).EnsureVisible
            lstChildren.ListItems(ListRow).Selected = True
            picToolbar.Enabled = False
            picMain.Enabled = False
            picChildrenSLine.Visible = True
            ListTrans = is_LstAdding
            txtChildName.SetFocus
        Case 3
            picTrainingsSLine.ZOrder 0
            txtTrainingTitle.Text = "": txtTrainingDates.Text = ""
            txtTrainingVenue.Text = ""
            With lstTraining.ListItems
                If Trim(.Item(1).SubItems(2)) <> "" Then
                    Set x = .Add()
                    x.Text = ""
                    x.SubItems(1) = " "
                    x.SubItems(2) = " "
                    x.SubItems(3) = " "
                    x.SubItems(4) = " "
                    ListRow = .Count
                End If
            End With
            lstTraining.ListItems(ListRow).EnsureVisible
            lstTraining.ListItems(ListRow).Selected = True
            picToolbar.Enabled = False
            picMain.Enabled = False
            picTrainingsSLine.Visible = True
            ListTrans = is_LstAdding
            txtTrainingTitle.SetFocus
        Case 4
            picEmploymentSLine.ZOrder 0
            txtEmploymentCompany.Text = "": txtEmploymentPosition.Text = ""
            txtEmploymentSalary.Text = "": txtEmploymentDates.Text = ""
            txtEmploymentAddress.Text = ""
            With lstEmployment.ListItems
                If Trim(.Item(1).SubItems(2)) <> "" Then
                    Set x = .Add()
                    x.Text = ""
                    x.SubItems(1) = " "
                    x.SubItems(2) = " "
                    x.SubItems(3) = " "
                    x.SubItems(4) = " "
                    x.SubItems(5) = " "
                    x.SubItems(6) = " "
                    ListRow = .Count
                End If
            End With
            lstEmployment.ListItems(ListRow).EnsureVisible
            lstEmployment.ListItems(ListRow).Selected = True
            picToolbar.Enabled = False
            picMain.Enabled = False
            picEmploymentSLine.Visible = True
            ListTrans = is_LstAdding
            txtEmploymentCompany.SetFocus
        Case Else: Exit Function
    End Select
End If
End Function

Private Function PRESS_F2()
If TRANSACTIONTYPE = is_REFRESH Then
    If AccessRights("Personnel Information", "Edit") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Function
    End If
    If Statusbar1.Panels(1).Text = "" Then Exit Function
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_EDITTING
ElseIf TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If picBrotherSisterSLine.Visible = True Then Exit Function
    If picChildrenSLine.Visible = True Then Exit Function
    If picTrainingsSLine.Visible = True Then Exit Function
    If picEmploymentSLine.Visible = True Then Exit Function
    Select Case ListFocus
        Case 1
            picBrotherSisterSLine.ZOrder 0
            With lstBroSis.ListItems
                txtBroSisName.Text = .Item(ListRow).SubItems(2)
                txtBroSisBDay.Text = .Item(ListRow).SubItems(3)
                txtBroSisOccupation.Text = .Item(ListRow).SubItems(4)
                txtBroSisAddress.Text = .Item(ListRow).SubItems(5)
                
                txtBroSisName1.Text = .Item(ListRow).SubItems(2)
                txtBroSisBDay1.Text = .Item(ListRow).SubItems(3)
                txtBroSisOccupation1.Text = .Item(ListRow).SubItems(4)
                txtBroSisAddress1.Text = .Item(ListRow).SubItems(5)
            End With
            picToolbar.Enabled = False
            picMain.Enabled = False
            picBrotherSisterSLine.Visible = True
            ListTrans = is_LstEditting
            txtBroSisName.SetFocus
        Case 2
            picChildrenSLine.ZOrder 0
            With lstChildren.ListItems
                txtChildName.Text = .Item(ListRow).SubItems(2)
                txtChildBDay.Text = .Item(ListRow).SubItems(3)
                txtChildOccupation.Text = .Item(ListRow).SubItems(4)
                txtChildAddress.Text = .Item(ListRow).SubItems(5)
                
                txtChildName1.Text = .Item(ListRow).SubItems(2)
                txtChildBDay1.Text = .Item(ListRow).SubItems(3)
                txtChildOccupation1.Text = .Item(ListRow).SubItems(4)
                txtChildAddress1.Text = .Item(ListRow).SubItems(5)
            End With
            picToolbar.Enabled = False
            picMain.Enabled = False
            picChildrenSLine.Visible = True
            ListTrans = is_LstEditting
            txtChildName.SetFocus
        Case 3
            picTrainingsSLine.ZOrder 0
            With lstTraining.ListItems
                txtTrainingTitle.Text = .Item(ListRow).SubItems(2)
                txtTrainingDates.Text = .Item(ListRow).SubItems(3)
                txtTrainingVenue.Text = .Item(ListRow).SubItems(4)
                
                txtTrainingTitle1.Text = .Item(ListRow).SubItems(2)
                txtTrainingDates1.Text = .Item(ListRow).SubItems(3)
                txtTrainingVenue1.Text = .Item(ListRow).SubItems(4)
            End With
            picToolbar.Enabled = False
            picMain.Enabled = False
            picTrainingsSLine.Visible = True
            ListTrans = is_LstEditting
            txtTrainingTitle.SetFocus
        Case 4
            picEmploymentSLine.ZOrder 0
            With lstEmployment.ListItems
                txtEmploymentCompany.Text = .Item(ListRow).SubItems(2)
                txtEmploymentPosition.Text = .Item(ListRow).SubItems(3)
                txtEmploymentSalary.Text = .Item(ListRow).SubItems(4)
                txtEmploymentDates.Text = .Item(ListRow).SubItems(5)
                txtEmploymentAddress.Text = .Item(ListRow).SubItems(6)
                
                txtEmploymentCompany1.Text = .Item(ListRow).SubItems(2)
                txtEmploymentPosition1.Text = .Item(ListRow).SubItems(3)
                txtEmploymentSalary1.Text = .Item(ListRow).SubItems(4)
                txtEmploymentDates1.Text = .Item(ListRow).SubItems(5)
                txtEmploymentAddress1.Text = .Item(ListRow).SubItems(6)
            End With
            picToolbar.Enabled = False
            picMain.Enabled = False
            picTrainingsSLine.Visible = True
            ListTrans = is_LstEditting
            txtTrainingTitle.SetFocus
        Case Else: Exit Function
    End Select
End If
End Function

Private Function PRESS_DELETE()
If TRANSACTIONTYPE = is_REFRESH Then
    If AccessRights("Personnel Information", "Delete") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Function
    End If
    If Statusbar1.Panels(1).Text = "" Then Exit Function
    If MsgBox("ARE YOU SURE IN DELETING THIS RECORD?                       ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Function
    On Error GoTo PG:
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Information WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    CLEARTEXT
    BROWSER GetSetting(App.EXEName, "ProfileInformation", "ProfileInfo", ""), "is_PAGEDOWN"
    If Trim(txtTmpFullName.Text) = "" Then BROWSER GetSetting(App.EXEName, "ProfileInformation", "ProfileInfo", ""), "is_HOME"
ElseIf TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If picBrotherSisterSLine.Visible = True Then Exit Function
    If picChildrenSLine.Visible = True Then Exit Function
    If picTrainingsSLine.Visible = True Then Exit Function
    If picEmploymentSLine.Visible = True Then Exit Function
    Select Case ListFocus
        Case 1
            With lstBroSis.ListItems
                If .Count > 1 Then
                    .Remove ListRow
                    If ListRow > .Count Then
                        ListRow = .Count
                    End If
                Else
                    .Item(1).SubItems(1) = " "
                    .Item(1).SubItems(2) = " "
                    .Item(1).SubItems(3) = " "
                    .Item(1).SubItems(4) = " "
                    .Item(1).SubItems(5) = " "
                End If
            End With
            lstBroSis.ListItems(ListRow).EnsureVisible
            lstBroSis.ListItems(ListRow).Selected = True
        Case 2
            With lstChildren.ListItems
                If .Count > 1 Then
                    .Remove ListRow
                    If ListRow > .Count Then
                        ListRow = .Count
                    End If
                Else
                    .Item(1).SubItems(1) = " "
                    .Item(1).SubItems(2) = " "
                    .Item(1).SubItems(3) = " "
                    .Item(1).SubItems(4) = " "
                    .Item(1).SubItems(5) = " "
                End If
            End With
            lstChildren.ListItems(ListRow).EnsureVisible
            lstChildren.ListItems(ListRow).Selected = True
        Case 3
            With lstTraining.ListItems
                If .Count > 1 Then
                    .Remove ListRow
                    If ListRow > .Count Then
                        ListRow = .Count
                    End If
                Else
                    .Item(1).SubItems(1) = " "
                    .Item(1).SubItems(2) = " "
                    .Item(1).SubItems(3) = " "
                    .Item(1).SubItems(4) = " "
                End If
            End With
            lstTraining.ListItems(ListRow).EnsureVisible
            lstTraining.ListItems(ListRow).Selected = True
        Case 4
            With lstEmployment.ListItems
                If .Count > 1 Then
                    .Remove ListRow
                    If ListRow > .Count Then
                        ListRow = .Count
                    End If
                Else
                    .Item(1).SubItems(1) = " "
                    .Item(1).SubItems(2) = " "
                    .Item(1).SubItems(3) = " "
                    .Item(1).SubItems(4) = " "
                    .Item(1).SubItems(5) = " "
                    .Item(1).SubItems(6) = " "
                End If
            End With
            lstEmployment.ListItems(ListRow).EnsureVisible
            lstEmployment.ListItems(ListRow).Selected = True
        Case Else: Exit Function
    End Select
End If
Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Function
End Function

Private Function PRESS_F5()
If Trim(txtLastName.Text) = "" Then MsgBox "Please Supply Last Name!                      ", vbCritical, "Error...": XTab1.ActiveTab = 0: txtLastName.SetFocus: HTEXT txtLastName: Exit Function
If Trim(txtFirstName.Text) = "" Then MsgBox "Please Supply Given Name!                      ", vbCritical, "Error...": XTab1.ActiveTab = 0: txtFirstName.SetFocus: HTEXT txtFirstName: Exit Function
If Trim(txtMiddleName.Text) = "" Then MsgBox "Please Supply Middle Name!                      ", vbCritical, "Error...": XTab1.ActiveTab = 0: txtMiddleName.SetFocus: HTEXT txtMiddleName: Exit Function
If cmbOwnedHouse.ListIndex = -1 Then MsgBox "Please select if house is owned!                   ", vbCritical, "Error...": XTab1.ActiveTab = 0: cmbOwnedHouse.SetFocus: Exit Function
If cmbRent.ListIndex = -1 Then MsgBox "Please select if your are rented!                         ", vbCritical, "Error...": XTab1.ActiveTab = 0: cmbRent.SetFocus: Exit Function
If cmbGender.ListIndex = -1 Then MsgBox "Please select Gender!                           ", vbCritical, "Error...": XTab1.ActiveTab = 0: cmbGender.SetFocus: Exit Function
If cmbLivingParents.ListIndex = -1 Then MsgBox "Please select if living with parents!                        ", vbCritical, "Error...": XTab1.ActiveTab = 0: cmbLivingParents.SetFocus: Exit Function
If cmbCivilStatus.ListIndex = -1 Then MsgBox "Please select Civil Status!                    ", vbCritical, "Error...": XTab1.ActiveTab = 0: cmbCivilStatus.SetFocus: Exit Function
If IsDate(txtBirthDate.Text) = False Then MsgBox "Please Supply Birth Date!                       ", vbCritical, "Error...": XTab1.ActiveTab = 0: txtBirthDate.SetFocus: HTEXT txtBirthDate: Exit Function

strLastName = FORMATSQL(Trim(txtLastName.Text))
strFirstName = FORMATSQL(Trim(txtFirstName.Text))
strMiddleName = FORMATSQL(Trim(txtMiddleName.Text))
strPresentAddress = FORMATSQL(Trim(txtPresentAddress.Text))
intOwnedHouse = cmbOwnedHouse.ListIndex + 1
intRented = cmbRent.ListIndex + 1
intGender = cmbGender.ListIndex + 1
intLivingWParents = cmbLivingParents.ListIndex + 1
intCivilStatus = cmbCivilStatus.ListIndex + 1
dtmBirthDate = FormatDateTime(txtBirthDate.Text, vbShortDate)
strBirthPlace = FORMATSQL(Trim(txtBirthPlace.Text))
strReligion = FORMATSQL(Trim(txtReligion.Text))
dblHeight = RETURNTEXTVALUE(txtHeight)
dblWeight = RETURNTEXTVALUE(txtWeight)
strNationality = FORMATSQL(Trim(txtNationality.Text))
strContactNumber = FORMATSQL(Trim(txtContactNumber.Text))
strSSSNumber = FORMATSQL(Trim(txtSSSNumber.Text))
strPHICNumber = FORMATSQL(Trim(txtPHICNumber.Text))
strHDMFNumber = FORMATSQL(Trim(txtPagIbigNumber.Text))
strTIN = FORMATSQL(Trim(txtTIN.Text))
strDriverLicense = FORMATSQL(Trim(txtDriverLicense.Text))
strSpouseName = FORMATSQL(Trim(txtSpouseName.Text))
strSpouseOccupation = FORMATSQL(Trim(txtSpouseOccupation.Text))
strSpouseAddress = FORMATSQL(Trim(txtSpouseAddress.Text))
strFatherName = FORMATSQL(Trim(txtFatherName.Text))
strFatherOccupation = FORMATSQL(Trim(txtFatherOccupation.Text))
strFatherAddress = FORMATSQL(Trim(txtFatherAddress.Text))
strMotherName = FORMATSQL(Trim(txtMotherName.Text))
strMotherOccupation = FORMATSQL(Trim(txtMotherOccupation.Text))
strMotherAddress = FORMATSQL(Trim(txtMotherAddress.Text))
strSkills = FORMATSQL(Trim(txtSkills.Text))
strOrganizationClubs = FORMATSQL(Trim(txtOrgsClubs.Text))
strRefName = FORMATSQL(Trim(txtNotRelatedName.Text))
strRefContact = FORMATSQL(Trim(txtNotRelatedContact.Text))
strRefAddress = FORMATSQL(Trim(txtNotRelatedAddress.Text))
strRefCompName = FORMATSQL(Trim(txtRelativeCompanyName.Text))
strRefCompContact = FORMATSQL(Trim(txtRelativeCompanyContact.Text))
strRefCompAddress = FORMATSQL(Trim(txtRelativeCompanyAddress.Text))
strEmergencyName = FORMATSQL(Trim(txtEmegencyName.Text))
strEmergencyRelation = FORMATSQL(Trim(txtEmegencyRelation.Text))
strEmergencyAddress = FORMATSQL(Trim(txtEmegencyAddress.Text))
strEmergencyContact = FORMATSQL(Trim(txtEmegencyContact.Text))
strLastModified = CStr(Now) & " - " & gbl_CompleteName
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Information " & _
                      " (LastName, FirstName, MiddleName, PresentAddress, OwnedHouse, Rented, Gender, LivingWParents, CivilStatus, BirthDate, BirthPlace, Religion, " & _
                      " Height, Weight, Nationality, ContactNumber, SSSNumber, PHICNumber, HDMFNumber, TIN, DriverLicense, SpouseName, SpouseOccupation, " & _
                      " SpouseAddress, FatherName, FatherOccupation, FatherAddress, MotherName, MotherOccupation, MotherAddress, Skills, OrganizationClubs, RefName, " & _
                      " RefContact, RefAddress, RefCompName, RefCompContact, RefCompAddress, EmergencyName, EmergencyRelation, EmergencyAddress, EmergencyContact, " & _
                      " LastModified, NoDependent, TaxStatus, BloodType) " & _
                      " VALUES ('" & strLastName & "', '" & strFirstName & "', '" & strMiddleName & "', '" & strPresentAddress & "', " & intOwnedHouse & ", " & intRented & ", " & _
                      " " & intGender & ", " & intLivingWParents & ", " & intCivilStatus & ", '" & dtmBirthDate & "', '" & strBirthPlace & "', '" & strReligion & "', " & dblHeight & ", " & _
                      " " & dblWeight & ", '" & strNationality & "', '" & strContactNumber & "', '" & strSSSNumber & "', '" & strPHICNumber & "', '" & strHDMFNumber & "', '" & strTIN & "', " & _
                      " '" & strDriverLicense & "', '" & strSpouseName & "', '" & strSpouseOccupation & "', '" & strSpouseAddress & "', '" & strFatherName & "', '" & strFatherOccupation & "', " & _
                      " '" & strFatherAddress & "', '" & strMotherName & "', '" & strMotherOccupation & "', '" & strMotherAddress & "', '" & strSkills & "', '" & strOrganizationClubs & "', " & _
                      " '" & strRefName & "', '" & strRefContact & "', '" & strRefAddress & "', '" & strRefCompName & "', '" & strRefCompContact & "', '" & strRefCompAddress & "', " & _
                      " '" & strEmergencyName & "', '" & strEmergencyRelation & "', '" & strEmergencyAddress & "', '" & strEmergencyContact & "', '" & strLastModified & "', " & _
                      " " & RETURNTEXTVALUE(txtNoDependent) & ", " & cmbTaxStatus.ListIndex + 1 & ", '" & FORMATSQL(txtBloodType.Text) & "')"
    
    strFullName = Trim(txtLastName.Text) & ",  " & Trim(txtFirstName.Text) & "  " & Trim(txtMiddleName.Text)
    ProfilePK = 0
    s = "SELECT PK " & _
        " FROM tbl_Personnel_Information " & _
        " WHERE (LastName + ',  ' + FirstName + '  ' + MiddleName = '" & FORMATSQL(CStr(strFullName)) & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        ProfilePK = rs!PK
    End If
    rs.Close
    
    If CDbl(ProfilePK) > 0 Then
        
        If IsDate(txtDateMarriage.Text) = False Then
            ConnOmega.Execute "UPDATE tbl_Personnel_Information SET DateMarriage = Null WHERE (PK = " & ProfilePK & ")"
        Else
            ConnOmega.Execute "UPDATE tbl_Personnel_Information SET DateMarriage = '" & FormatDateTime(txtDateMarriage.Text, vbShortDate) & "' WHERE (PK = " & ProfilePK & ")"
        End If
        If IsDate(txtSpouseBDay.Text) = False Then
            ConnOmega.Execute "UPDATE tbl_Personnel_Information SET SpouseBDay = Null WHERE (PK = " & ProfilePK & ")"
        Else
            ConnOmega.Execute "UPDATE tbl_Personnel_Information SET SpouseBDay = '" & FormatDateTime(txtSpouseBDay.Text, vbShortDate) & "' WHERE (PK = " & ProfilePK & ")"
        End If
        If IsDate(txtFatherBDay.Text) = False Then
            ConnOmega.Execute "UPDATE tbl_Personnel_Information SET FatherBDay = Null WHERE (PK = " & ProfilePK & ")"
        Else
            ConnOmega.Execute "UPDATE tbl_Personnel_Information SET FatherBDay = '" & FormatDateTime(txtFatherBDay.Text, vbShortDate) & "' WHERE (PK = " & ProfilePK & ")"
        End If
        If IsDate(txtMotherBDay.Text) = False Then
            ConnOmega.Execute "UPDATE tbl_Personnel_Information SET MotherBDay = Null WHERE (PK = " & ProfilePK & ")"
        Else
            ConnOmega.Execute "UPDATE tbl_Personnel_Information SET MotherBDay = '" & FormatDateTime(txtMotherBDay.Text, vbShortDate) & "' WHERE (PK = " & ProfilePK & ")"
        End If
        
        If Trim(txtPicturePath.Text) <> "" Then
            
            SAVE_IMAGES ProfilePK, 0, Trim(txtPicturePath.Text), "Employee Profile"
            
        End If
        
        ConnOmega.Execute "DELETE FROM tbl_Personnel_BrotherSister WHERE(ProfileKey =" & ProfilePK & ")"
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Children WHERE(ProfileKey =" & ProfilePK & ")"
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Education WHERE(ProfileKey =" & ProfilePK & ")"
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Employment WHERE(ProfileKey =" & ProfilePK & ")"
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Training WHERE(ProfileKey =" & ProfilePK & ")"
        
        '=== Brother / Sister
        j = 0
        With lstBroSis.ListItems
            For i = 1 To .Count
                If Trim(.Item(i).SubItems(2)) <> "" Then
                    j = j + 1
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_BrotherSister " & _
                                      " (ProfileKey, Line, BrotherSisterName, BrotherSisterOccupation, BrotherSisterAddress) " & _
                                      " VALUES (" & ProfilePK & ", " & j & ", '" & FORMATSQL(Trim(.Item(i).SubItems(2))) & "', " & _
                                      " '" & FORMATSQL(Trim(.Item(i).SubItems(4))) & "', '" & FORMATSQL(Trim(.Item(i).SubItems(5))) & "')"
                    If IsDate(.Item(i).SubItems(3)) = False Then
                        ConnOmega.Execute "UPDATE tbl_Personnel_BrotherSister " & _
                                          " SET BrotherSisterBDay = Null " & _
                                          " WHERE (ProfileKey = " & ProfilePK & ") " & _
                                          " AND (Line = " & j & ")"
                    Else
                        ConnOmega.Execute "UPDATE tbl_Personnel_BrotherSister " & _
                                          " SET BrotherSisterBDay = '" & FormatDateTime(.Item(i).SubItems(3), vbShortDate) & "' " & _
                                          " WHERE (ProfileKey = " & ProfilePK & ") " & _
                                          " AND (Line = " & j & ")"
                    End If
                End If
            Next i
        End With
        
        '=== Childrens
        j = 0
        With lstChildren.ListItems
            For i = 1 To .Count
                If Trim(.Item(i).SubItems(2)) <> "" Then
                    j = j + 1
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Children " & _
                                      " (ProfileKey, Line, ChildName, ChildOccupation, ChildAddress) " & _
                                      " VALUES (" & ProfilePK & ", " & j & ", '" & FORMATSQL(Trim(.Item(i).SubItems(2))) & "', " & _
                                      " '" & FORMATSQL(Trim(.Item(i).SubItems(4))) & "', '" & FORMATSQL(Trim(.Item(i).SubItems(5))) & "')"
                    If IsDate(.Item(i).SubItems(3)) = False Then
                        ConnOmega.Execute "UPDATE tbl_Personnel_Children " & _
                                          " SET ChildBDay = Null " & _
                                          " WHERE (ProfileKey = " & ProfilePK & ") " & _
                                          " AND (Line = " & j & ")"
                    Else
                        ConnOmega.Execute "UPDATE tbl_Personnel_Children " & _
                                          " SET ChildBDay = '" & FormatDateTime(.Item(i).SubItems(3), vbShortDate) & "' " & _
                                          " WHERE (ProfileKey = " & ProfilePK & ") " & _
                                          " AND (Line = " & j & ")"
                    End If
                End If
            Next i
        End With
        
        '=== Trainings
        j = 0
        With lstTraining.ListItems
            For i = 1 To .Count
                If Trim(.Item(i).SubItems(2)) <> "" Then
                    j = j + 1
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Training " & _
                                      " (ProfileKey, Line, Title, InclusiveDate, Venue) " & _
                                      " VALUES (" & ProfilePK & ", " & j & ", '" & FORMATSQL(Trim(.Item(i).SubItems(2))) & "', " & _
                                      " '" & FORMATSQL(Trim(.Item(i).SubItems(3))) & "', '" & FORMATSQL(Trim(.Item(i).SubItems(4))) & "')"
                End If
            Next i
        End With
        
        '=== Education
        For i = 1 To 4
            Select Case i
                Case 1
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Education " & _
                                      " (ProfileKey, Line, SchoolName, InclusiveDate, Course, Address) " & _
                                      " VALUES (" & ProfilePK & ", " & i & ", '" & FORMATSQL(Trim(txtElemName.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtElemInclusiveDate.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtElemCourse.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtElemAddress.Text)) & "')"
                Case 2
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Education " & _
                                      " (ProfileKey, Line, SchoolName, InclusiveDate, Course, Address) " & _
                                      " VALUES (" & ProfilePK & ", " & i & ", '" & FORMATSQL(Trim(txtHiSchoolName.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtHiSchoolInclusiveDate.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtHiSchoolCourse.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtHiSchoolAddress.Text)) & "')"
                Case 3
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Education " & _
                                      " (ProfileKey, Line, SchoolName, InclusiveDate, Course, Address) " & _
                                      " VALUES (" & ProfilePK & ", " & i & ", '" & FORMATSQL(Trim(txtCollegeName.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtCollegeInclusiveDate.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtCollegeCourse.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtCollegeAddress.Text)) & "')"
                Case 4
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Education " & _
                                      " (ProfileKey, Line, SchoolName, InclusiveDate, Course, Address) " & _
                                      " VALUES (" & ProfilePK & ", " & i & ", '" & FORMATSQL(Trim(txtPostName.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtPostInclusiveDate.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtPostCourse.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtPostAddress.Text)) & "')"
            End Select
        Next i
        
        '=== Employment
        j = 0
        With lstEmployment.ListItems
            For i = 1 To .Count
                If Trim(.Item(i).SubItems(2)) <> "" Then
                    j = j + 1
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Employment " & _
                                      " (ProfileKey, Line, Company, Positions, Salary, InclusiveDate, Address) " & _
                                      " VALUES (" & ProfilePK & ", " & j & ", '" & FORMATSQL(Trim(.Item(i).SubItems(2))) & "', " & _
                                      " '" & FORMATSQL(Trim(.Item(i).SubItems(3))) & "', " & CDbl(IIf(Trim(Trim(.Item(i).SubItems(4))) = "", "0", Trim(.Item(i).SubItems(4)))) & ", " & _
                                      " '" & FORMATSQL(Trim(.Item(i).SubItems(5))) & "', '" & FORMATSQL(Trim(.Item(i).SubItems(6))) & "')"
                End If
            Next i
        End With
    End If
    
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER strFullName, "is_LOAD"
    
    If MsgBox("SAVED!                   " & vbCrLf & vbCrLf & "ADD ANOTHER PROFILE?                 ", vbInformation + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Function
    PRESS_INSERT
    
End If
If TRANSACTIONTYPE = is_EDITTING Then
    ProfilePK = Statusbar1.Panels(1).Text
    
    ConnOmega.Execute "UPDATE tbl_Personnel_Information " & _
                      " SET LastName = '" & strLastName & "', FirstName = '" & strFirstName & "', MiddleName = '" & strMiddleName & "', " & _
                      " PresentAddress = '" & strPresentAddress & "', OwnedHouse = " & intOwnedHouse & ", Rented = " & intRented & ", " & _
                      " Gender = " & intGender & ", LivingWParents = " & intLivingWParents & ", CivilStatus = " & intCivilStatus & ", " & _
                      " BirthDate = '" & dtmBirthDate & "', BirthPlace = '" & strBirthPlace & "', Religion = '" & strReligion & "', " & _
                      " Height = " & dblHeight & ", Weight = " & dblWeight & ", Nationality = '" & strNationality & "', " & _
                      " ContactNumber = '" & strContactNumber & "', SSSNumber = '" & strSSSNumber & "', PHICNumber = '" & strPHICNumber & "', " & _
                      " HDMFNumber = '" & strHDMFNumber & "', TIN = '" & strTIN & "', DriverLicense = '" & strDriverLicense & "', " & _
                      " SpouseName = '" & strSpouseName & "', SpouseOccupation = '" & strSpouseOccupation & "', SpouseAddress = '" & strSpouseAddress & "', " & _
                      " FatherName = '" & strFatherName & "', FatherOccupation = '" & strFatherOccupation & "', FatherAddress = '" & strFatherAddress & "', " & _
                      " MotherName = '" & strMotherName & "', MotherOccupation = '" & strMotherOccupation & "', MotherAddress = '" & strMotherAddress & "', " & _
                      " Skills = '" & strSkills & "', OrganizationClubs = '" & strOrganizationClubs & "', RefName = '" & strRefName & "', " & _
                      " RefContact = '" & strRefContact & "', RefAddress = '" & strRefAddress & "', RefCompName = '" & strRefCompName & "', " & _
                      " RefCompContact = '" & strRefCompContact & "', RefCompAddress = '" & strRefCompAddress & "', EmergencyName = '" & strEmergencyName & "', " & _
                      " EmergencyRelation = '" & strEmergencyRelation & "', EmergencyAddress = '" & strEmergencyAddress & "', " & _
                      " EmergencyContact = '" & strEmergencyContact & "', LastModified = '" & strLastModified & "', " & _
                      " NoDependent = " & RETURNTEXTVALUE(txtNoDependent) & ", TaxStatus = " & cmbTaxStatus.ListIndex + 1 & ", " & _
                      " BloodType = '" & FORMATSQL(txtBloodType.Text) & "' " & _
                      " WHERE (PK = " & ProfilePK & ")"
    
    strFullName = Trim(txtLastName.Text) & ",  " & Trim(txtFirstName.Text) & "  " & Trim(txtMiddleName.Text)
    
    If CDbl(ProfilePK) > 0 Then
        
        If IsDate(txtDateMarriage.Text) = False Then
            ConnOmega.Execute "UPDATE tbl_Personnel_Information SET DateMarriage = Null WHERE (PK = " & ProfilePK & ")"
        Else
            ConnOmega.Execute "UPDATE tbl_Personnel_Information SET DateMarriage = '" & FormatDateTime(txtDateMarriage.Text, vbShortDate) & "' WHERE (PK = " & ProfilePK & ")"
        End If
        If IsDate(txtSpouseBDay.Text) = False Then
            ConnOmega.Execute "UPDATE tbl_Personnel_Information SET SpouseBDay = Null WHERE (PK = " & ProfilePK & ")"
        Else
            ConnOmega.Execute "UPDATE tbl_Personnel_Information SET SpouseBDay = '" & FormatDateTime(txtSpouseBDay.Text, vbShortDate) & "' WHERE (PK = " & ProfilePK & ")"
        End If
        If IsDate(txtFatherBDay.Text) = False Then
            ConnOmega.Execute "UPDATE tbl_Personnel_Information SET FatherBDay = Null WHERE (PK = " & ProfilePK & ")"
        Else
            ConnOmega.Execute "UPDATE tbl_Personnel_Information SET FatherBDay = '" & FormatDateTime(txtFatherBDay.Text, vbShortDate) & "' WHERE (PK = " & ProfilePK & ")"
        End If
        If IsDate(txtMotherBDay.Text) = False Then
            ConnOmega.Execute "UPDATE tbl_Personnel_Information SET MotherBDay = Null WHERE (PK = " & ProfilePK & ")"
        Else
            ConnOmega.Execute "UPDATE tbl_Personnel_Information SET MotherBDay = '" & FormatDateTime(txtMotherBDay.Text, vbShortDate) & "' WHERE (PK = " & ProfilePK & ")"
        End If
        
        If Trim(txtPicturePath.Text) <> "" Then
            
            SAVE_IMAGES ProfilePK, 0, Trim(txtPicturePath.Text), "Employee Profile"
            
        End If
        
        ConnOmega.Execute "DELETE FROM tbl_Personnel_BrotherSister WHERE(ProfileKey =" & ProfilePK & ")"
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Children WHERE(ProfileKey =" & ProfilePK & ")"
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Education WHERE(ProfileKey =" & ProfilePK & ")"
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Employment WHERE(ProfileKey =" & ProfilePK & ")"
        ConnOmega.Execute "DELETE FROM tbl_Personnel_Training WHERE(ProfileKey =" & ProfilePK & ")"
        
        '=== Brother / Sister
        j = 0
        With lstBroSis.ListItems
            For i = 1 To .Count
                If Trim(.Item(i).SubItems(2)) <> "" Then
                    j = j + 1
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_BrotherSister " & _
                                      " (ProfileKey, Line, BrotherSisterName, BrotherSisterOccupation, BrotherSisterAddress) " & _
                                      " VALUES (" & ProfilePK & ", " & j & ", '" & FORMATSQL(Trim(.Item(i).SubItems(2))) & "', " & _
                                      " '" & FORMATSQL(Trim(.Item(i).SubItems(4))) & "', '" & FORMATSQL(Trim(.Item(i).SubItems(5))) & "')"
                    If IsDate(.Item(i).SubItems(3)) = False Then
                        ConnOmega.Execute "UPDATE tbl_Personnel_BrotherSister " & _
                                          " SET BrotherSisterBDay = Null " & _
                                          " WHERE (ProfileKey = " & ProfilePK & ") " & _
                                          " AND (Line = " & j & ")"
                    Else
                        ConnOmega.Execute "UPDATE tbl_Personnel_BrotherSister " & _
                                          " SET BrotherSisterBDay = '" & FormatDateTime(.Item(i).SubItems(3), vbShortDate) & "' " & _
                                          " WHERE (ProfileKey = " & ProfilePK & ") " & _
                                          " AND (Line = " & j & ")"
                    End If
                End If
            Next i
        End With
        
        '=== Childrens
        j = 0
        With lstChildren.ListItems
            For i = 1 To .Count
                If Trim(.Item(i).SubItems(2)) <> "" Then
                    j = j + 1
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Children " & _
                                      " (ProfileKey, Line, ChildName, ChildOccupation, ChildAddress) " & _
                                      " VALUES (" & ProfilePK & ", " & j & ", '" & FORMATSQL(Trim(.Item(i).SubItems(2))) & "', " & _
                                      " '" & FORMATSQL(Trim(.Item(i).SubItems(4))) & "', '" & FORMATSQL(Trim(.Item(i).SubItems(5))) & "')"
                    If IsDate(.Item(i).SubItems(3)) = False Then
                        ConnOmega.Execute "UPDATE tbl_Personnel_Children " & _
                                          " SET ChildBDay = Null " & _
                                          " WHERE (ProfileKey = " & ProfilePK & ") " & _
                                          " AND (Line = " & j & ")"
                    Else
                        ConnOmega.Execute "UPDATE tbl_Personnel_Children " & _
                                          " SET ChildBDay = '" & FormatDateTime(.Item(i).SubItems(3), vbShortDate) & "' " & _
                                          " WHERE (ProfileKey = " & ProfilePK & ") " & _
                                          " AND (Line = " & j & ")"
                    End If
                End If
            Next i
        End With
        
        '=== Trainings
        j = 0
        With lstTraining.ListItems
            For i = 1 To .Count
                If Trim(.Item(i).SubItems(2)) <> "" Then
                    j = j + 1
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Training " & _
                                      " (ProfileKey, Line, Title, InclusiveDate, Venue) " & _
                                      " VALUES (" & ProfilePK & ", " & j & ", '" & FORMATSQL(Trim(.Item(i).SubItems(2))) & "', " & _
                                      " '" & FORMATSQL(Trim(.Item(i).SubItems(3))) & "', '" & FORMATSQL(Trim(.Item(i).SubItems(4))) & "')"
                End If
            Next i
        End With
        
        '=== Education
        For i = 1 To 4
            Select Case i
                Case 1
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Education " & _
                                      " (ProfileKey, Line, SchoolName, InclusiveDate, Course, Address) " & _
                                      " VALUES (" & ProfilePK & ", " & i & ", '" & FORMATSQL(Trim(txtElemName.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtElemInclusiveDate.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtElemCourse.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtElemAddress.Text)) & "')"
                Case 2
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Education " & _
                                      " (ProfileKey, Line, SchoolName, InclusiveDate, Course, Address) " & _
                                      " VALUES (" & ProfilePK & ", " & i & ", '" & FORMATSQL(Trim(txtHiSchoolName.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtHiSchoolInclusiveDate.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtHiSchoolCourse.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtHiSchoolAddress.Text)) & "')"
                Case 3
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Education " & _
                                      " (ProfileKey, Line, SchoolName, InclusiveDate, Course, Address) " & _
                                      " VALUES (" & ProfilePK & ", " & i & ", '" & FORMATSQL(Trim(txtCollegeName.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtCollegeInclusiveDate.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtCollegeCourse.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtCollegeAddress.Text)) & "')"
                Case 4
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Education " & _
                                      " (ProfileKey, Line, SchoolName, InclusiveDate, Course, Address) " & _
                                      " VALUES (" & ProfilePK & ", " & i & ", '" & FORMATSQL(Trim(txtPostName.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtPostInclusiveDate.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtPostCourse.Text)) & "', " & _
                                      " '" & FORMATSQL(Trim(txtPostAddress.Text)) & "')"
            End Select
        Next i
        
        '=== Employment
        j = 0
        With lstEmployment.ListItems
            For i = 1 To .Count
                If Trim(.Item(i).SubItems(2)) <> "" Then
                    j = j + 1
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Employment " & _
                                      " (ProfileKey, Line, Company, Positions, Salary, InclusiveDate, Address) " & _
                                      " VALUES (" & ProfilePK & ", " & j & ", '" & FORMATSQL(Trim(.Item(i).SubItems(2))) & "', " & _
                                      " '" & FORMATSQL(Trim(.Item(i).SubItems(3))) & "', " & CDbl(IIf(Trim(Trim(.Item(i).SubItems(4))) = "", "0", Trim(.Item(i).SubItems(4)))) & ", " & _
                                      " '" & FORMATSQL(Trim(.Item(i).SubItems(5))) & "', '" & FORMATSQL(Trim(.Item(i).SubItems(6))) & "')"
                End If
            Next i
        End With
    End If
    
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER strFullName, "is_LOAD"
    
End If

Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Function
End Function

Private Function PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
PopupMenu MainFormPopupF.ProfileSearch, , Toolbar1.Buttons(15).Left, Toolbar1.Buttons(15).Top + Toolbar1.Buttons(15).Height
End Function

Private Function PRESS_F9()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If Statusbar1.Panels(1).Text = "" Then Exit Function
PopupMenu MainFormPopupF.ProfilePrint, , Toolbar1.Buttons(17).Left, Toolbar1.Buttons(17).Top + Toolbar1.Buttons(17).Height
End Function

Private Function PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSearch.Visible = True Then cmdCancel_Click: Exit Function
    If picAlphalist.Visible = True Then cmdCancelAlphalist_Click: Exit Function
    Unload Me
Else
    If picBrotherSisterSLine.Visible = True Then
        If ListTrans = is_LstAdding Then
            With lstBroSis.ListItems
                If .Count > 1 Then
                    .Remove .Count
                Else
                    .Item(1).SubItems(1) = " "
                    .Item(1).SubItems(2) = " "
                    .Item(1).SubItems(3) = " "
                    .Item(1).SubItems(4) = " "
                    .Item(1).SubItems(5) = " "
                End If
                ListRow = .Count
            End With
            picBrotherSisterSLine.Visible = False
            picMain.Enabled = True
            picToolbar.Enabled = True
            lstBroSis.ListItems(ListRow).EnsureVisible
            lstBroSis.ListItems(ListRow).Selected = True
            lstBroSis.SetFocus
            Exit Function
        End If
        
        If ListTrans = is_LstEditting Then
            With lstBroSis.ListItems
                .Item(ListRow).SubItems(2) = txtBroSisName1.Text
                .Item(ListRow).SubItems(3) = txtBroSisBDay1.Text
                .Item(ListRow).SubItems(4) = txtBroSisOccupation1.Text
                .Item(ListRow).SubItems(5) = txtBroSisAddress1.Text
            End With
            picBrotherSisterSLine.Visible = False
            picMain.Enabled = True
            picToolbar.Enabled = True
            lstBroSis.ListItems(ListRow).EnsureVisible
            lstBroSis.ListItems(ListRow).Selected = True
            lstBroSis.SetFocus
            Exit Function
        End If
    End If
    
    If picChildrenSLine.Visible = True Then
        If ListTrans = is_LstAdding Then
            With lstChildren.ListItems
                If .Count > 1 Then
                    .Remove .Count
                Else
                    .Item(1).SubItems(1) = " "
                    .Item(1).SubItems(2) = " "
                    .Item(1).SubItems(3) = " "
                    .Item(1).SubItems(4) = " "
                    .Item(1).SubItems(5) = " "
                End If
                ListRow = .Count
            End With
            picChildrenSLine.Visible = False
            picMain.Enabled = True
            picToolbar.Enabled = True
            lstChildren.ListItems(ListRow).EnsureVisible
            lstChildren.ListItems(ListRow).Selected = True
            lstChildren.SetFocus
            Exit Function
        End If
        
        If ListTrans = is_LstEditting Then
            With lstChildren.ListItems
                .Item(ListRow).SubItems(2) = txtChildName1.Text
                .Item(ListRow).SubItems(3) = txtChildBDay1.Text
                .Item(ListRow).SubItems(4) = txtChildOccupation1.Text
                .Item(ListRow).SubItems(5) = txtChildAddress1.Text
            End With
            picChildrenSLine.Visible = False
            picMain.Enabled = True
            picToolbar.Enabled = True
            lstChildren.ListItems(ListRow).EnsureVisible
            lstChildren.ListItems(ListRow).Selected = True
            lstChildren.SetFocus
            Exit Function
        End If
    End If
    
    If picTrainingsSLine.Visible = True Then
        If ListTrans = is_LstAdding Then
            With lstTraining.ListItems
                If .Count > 1 Then
                    .Remove .Count
                Else
                    .Item(1).SubItems(1) = " "
                    .Item(1).SubItems(2) = " "
                    .Item(1).SubItems(3) = " "
                    .Item(1).SubItems(4) = " "
                End If
                ListRow = .Count
            End With
            picTrainingsSLine.Visible = False
            picMain.Enabled = True
            picToolbar.Enabled = True
            lstTraining.ListItems(ListRow).EnsureVisible
            lstTraining.ListItems(ListRow).Selected = True
            lstTraining.SetFocus
            Exit Function
        End If
        
        If ListTrans = is_LstEditting Then
            With lstTraining.ListItems
                .Item(ListRow).SubItems(2) = txtTrainingTitle1.Text
                .Item(ListRow).SubItems(3) = txtTrainingDates1.Text
                .Item(ListRow).SubItems(4) = txtTrainingVenue1.Text
            End With
            picTrainingsSLine.Visible = False
            picMain.Enabled = True
            picToolbar.Enabled = True
            lstTraining.ListItems(ListRow).EnsureVisible
            lstTraining.ListItems(ListRow).Selected = True
            lstTraining.SetFocus
            Exit Function
        End If
    End If
    
    If picEmploymentSLine.Visible = True Then
        If ListTrans = is_LstAdding Then
            With lstEmployment.ListItems
                If .Count > 1 Then
                    .Remove .Count
                Else
                    .Item(1).SubItems(1) = " "
                    .Item(1).SubItems(2) = " "
                    .Item(1).SubItems(3) = " "
                    .Item(1).SubItems(4) = " "
                    .Item(1).SubItems(5) = " "
                    .Item(1).SubItems(6) = " "
                End If
                ListRow = .Count
            End With
            picEmploymentSLine.Visible = False
            picMain.Enabled = True
            picToolbar.Enabled = True
            lstEmployment.ListItems(ListRow).EnsureVisible
            lstEmployment.ListItems(ListRow).Selected = True
            lstEmployment.SetFocus
            Exit Function
        End If
        
        If ListTrans = is_LstEditting Then
            With lstEmployment.ListItems
                .Item(ListRow).SubItems(2) = txtEmploymentCompany1.Text
                .Item(ListRow).SubItems(3) = txtEmploymentPosition1.Text
                .Item(ListRow).SubItems(4) = txtEmploymentSalary1.Text
                .Item(ListRow).SubItems(5) = txtEmploymentDates1.Text
                .Item(ListRow).SubItems(6) = txtEmploymentAddress1.Text
            End With
            picEmploymentSLine.Visible = False
            picMain.Enabled = True
            picToolbar.Enabled = True
            lstEmployment.ListItems(ListRow).EnsureVisible
            lstEmployment.ListItems(ListRow).Selected = True
            lstEmployment.SetFocus
            Exit Function
        End If
    End If
    
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    ListTrans = is_LstRefresh
    BROWSER GetSetting(App.EXEName, "ProfileInformation", "ProfileInfo", ""), "is_LOAD"
    If Trim(txtTmpFullName.Text) = "" Then BROWSER GetSetting(App.EXEName, "ProfileInformation", "ProfileInfo", ""), "is_HOME"
End If
End Function

Private Function CLEARTEXT()
txtLastName.Text = ""
txtFirstName.Text = ""
txtMiddleName.Text = ""
txtPresentAddress.Text = ""
txtBirthDate.Text = ""
txtAge.Text = ""
txtBirthPlace.Text = ""
txtReligion.Text = ""
txtDateMarriage.Text = ""
txtHeight.Text = ""
txtWeight.Text = ""
txtNationality.Text = ""
txtContactNumber.Text = ""
txtSSSNumber.Text = ""
txtTIN.Text = ""
txtDriverLicense.Text = ""
txtPHICNumber.Text = ""
txtPagIbigNumber.Text = ""
txtSpouseName.Text = ""
txtSpouseBDay.Text = ""
txtSpouseOccupation.Text = ""
txtSpouseAddress.Text = ""
txtFatherName.Text = ""
txtFatherBDay.Text = ""
txtFatherOccupation.Text = ""
txtFatherAddress.Text = ""
txtMotherName.Text = ""
txtMotherBDay.Text = ""
txtMotherOccupation.Text = ""
txtMotherAddress.Text = ""
txtElemName.Text = ""
txtElemInclusiveDate.Text = ""
txtElemCourse.Text = ""
txtElemAddress.Text = ""
txtHiSchoolName.Text = ""
txtHiSchoolInclusiveDate.Text = ""
txtHiSchoolCourse.Text = ""
txtHiSchoolAddress.Text = ""
txtCollegeName.Text = ""
txtCollegeInclusiveDate.Text = ""
txtCollegeCourse.Text = ""
txtCollegeAddress.Text = ""
txtPostName.Text = ""
txtPostInclusiveDate.Text = ""
txtPostCourse.Text = ""
txtPostAddress.Text = ""
txtSkills.Text = ""
txtOrgsClubs.Text = ""
txtNotRelatedName.Text = ""
txtNotRelatedContact.Text = ""
txtNotRelatedAddress.Text = ""
txtRelativeCompanyName.Text = ""
txtRelativeCompanyContact.Text = ""
txtRelativeCompanyAddress.Text = ""
txtEmegencyName.Text = ""
txtEmegencyRelation.Text = ""
txtEmegencyAddress.Text = ""
txtEmegencyContact.Text = ""
txtNoDependent.Text = ""
txtBloodType.Text = ""

cmbGender.ListIndex = -1
cmbOwnedHouse.ListIndex = -1
cmbRent.ListIndex = -1
cmbLivingParents.ListIndex = -1
cmbCivilStatus.ListIndex = -1
cmbIDNumber.Clear
cmbTaxStatus.ListIndex = -1

Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""

txtTmpFullName.Text = ""
txtTmpBDay.Text = ""
txtTmpAge.Text = ""
txtTmpGender.Text = ""
txtTmpCivilStatus.Text = ""
txtTmpAddress.Text = ""
txtTmpContact.Text = ""
txtTmpHeight.Text = ""
txtTmpWeight.Text = ""
txtTmpSSS.Text = ""
txtTmpPHIC.Text = ""
txtTmpPagibig.Text = ""
txtTmpTIN.Text = ""
txtTmpDriverLic.Text = ""
txtPicturePath.Text = ""
txtTmpIDNumber.Text = ""
txtTmpStatus.Text = ""
txtTmpBloodType.Text = ""

imgPicture.Picture = LoadPicture("")

lstBroSis.ListItems.Clear
Set x = lstBroSis.ListItems.Add()
x.Text = "": x.SubItems(1) = "": x.SubItems(2) = " ": x.SubItems(3) = " ": x.SubItems(4) = " ": x.SubItems(5) = " "

lstChildren.ListItems.Clear
Set x = lstChildren.ListItems.Add()
x.Text = "": x.SubItems(1) = "": x.SubItems(2) = " ": x.SubItems(3) = " ": x.SubItems(4) = " ": x.SubItems(5) = " "

lstTraining.ListItems.Clear
Set x = lstTraining.ListItems.Add()
x.Text = "": x.SubItems(1) = "": x.SubItems(2) = " ": x.SubItems(3) = " ": x.SubItems(4) = " "

lstEmployment.ListItems.Clear
Set x = lstEmployment.ListItems.Add()
x.Text = "": x.SubItems(1) = "": x.SubItems(2) = " ": x.SubItems(3) = " ": x.SubItems(4) = " ": x.SubItems(5) = " ": x.SubItems(6) = " "

End Function

Private Function LOCKTEXT(bln As Boolean)
If bln Then
    txtLastName.Locked = True
    txtFirstName.Locked = True
    txtMiddleName.Locked = True
    txtPresentAddress.Locked = True
    txtBirthDate.Locked = True
    txtAge.Locked = True
    txtBirthPlace.Locked = True
    txtReligion.Locked = True
    txtDateMarriage.Locked = True
    txtHeight.Locked = True
    txtWeight.Locked = True
    txtNationality.Locked = True
    txtBloodType.Locked = True
    txtContactNumber.Locked = True
    txtSSSNumber.Locked = True
    txtTIN.Locked = True
    txtDriverLicense.Locked = True
    txtPHICNumber.Locked = True
    txtPagIbigNumber.Locked = True
    txtSpouseName.Locked = True
    txtSpouseBDay.Locked = True
    txtSpouseOccupation.Locked = True
    txtSpouseAddress.Locked = True
    txtFatherName.Locked = True
    txtFatherBDay.Locked = True
    txtFatherOccupation.Locked = True
    txtFatherAddress.Locked = True
    txtMotherName.Locked = True
    txtMotherBDay.Locked = True
    txtMotherOccupation.Locked = True
    txtMotherAddress.Locked = True
    txtElemName.Locked = True
    txtElemInclusiveDate.Locked = True
    txtElemCourse.Locked = True
    txtElemAddress.Locked = True
    txtHiSchoolName.Locked = True
    txtHiSchoolInclusiveDate.Locked = True
    txtHiSchoolCourse.Locked = True
    txtHiSchoolAddress.Locked = True
    txtCollegeName.Locked = True
    txtCollegeInclusiveDate.Locked = True
    txtCollegeCourse.Locked = True
    txtCollegeAddress.Locked = True
    txtPostName.Locked = True
    txtPostInclusiveDate.Locked = True
    txtPostCourse.Locked = True
    txtPostAddress.Locked = True
    txtSkills.Locked = True
    txtOrgsClubs.Locked = True
    txtNotRelatedName.Locked = True
    txtNotRelatedContact.Locked = True
    txtNotRelatedAddress.Locked = True
    txtRelativeCompanyName.Locked = True
    txtRelativeCompanyContact.Locked = True
    txtRelativeCompanyAddress.Locked = True
    txtEmegencyName.Locked = True
    txtEmegencyRelation.Locked = True
    txtEmegencyAddress.Locked = True
    txtEmegencyContact.Locked = True
    txtNoDependent.Locked = True
    
    cmbGender.Locked = True
    cmbOwnedHouse.Locked = True
    cmbRent.Locked = True
    cmbLivingParents.Locked = True
    cmbCivilStatus.Locked = True
    cmbIDNumber.Locked = True
    cmbTaxStatus.Locked = True
Else
    txtLastName.Locked = False
    txtFirstName.Locked = False
    txtMiddleName.Locked = False
    txtPresentAddress.Locked = False
    txtBirthDate.Locked = False
    'txtAge.Locked = False
    txtBirthPlace.Locked = False
    txtReligion.Locked = False
    txtDateMarriage.Locked = False
    txtHeight.Locked = False
    txtWeight.Locked = False
    txtBloodType.Locked = False
    txtNationality.Locked = False
    txtContactNumber.Locked = False
    txtSSSNumber.Locked = False
    txtTIN.Locked = False
    txtDriverLicense.Locked = False
    txtPHICNumber.Locked = False
    txtPagIbigNumber.Locked = False
    txtSpouseName.Locked = False
    txtSpouseBDay.Locked = False
    txtSpouseOccupation.Locked = False
    txtSpouseAddress.Locked = False
    txtFatherName.Locked = False
    txtFatherBDay.Locked = False
    txtFatherOccupation.Locked = False
    txtFatherAddress.Locked = False
    txtMotherName.Locked = False
    txtMotherBDay.Locked = False
    txtMotherOccupation.Locked = False
    txtMotherAddress.Locked = False
    txtElemName.Locked = False
    txtElemInclusiveDate.Locked = False
    txtElemCourse.Locked = False
    txtElemAddress.Locked = False
    txtHiSchoolName.Locked = False
    txtHiSchoolInclusiveDate.Locked = False
    txtHiSchoolCourse.Locked = False
    txtHiSchoolAddress.Locked = False
    txtCollegeName.Locked = False
    txtCollegeInclusiveDate.Locked = False
    txtCollegeCourse.Locked = False
    txtCollegeAddress.Locked = False
    txtPostName.Locked = False
    txtPostInclusiveDate.Locked = False
    txtPostCourse.Locked = False
    txtPostAddress.Locked = False
    txtSkills.Locked = False
    txtOrgsClubs.Locked = False
    txtNotRelatedName.Locked = False
    txtNotRelatedContact.Locked = False
    txtNotRelatedAddress.Locked = False
    txtRelativeCompanyName.Locked = False
    txtRelativeCompanyContact.Locked = False
    txtRelativeCompanyAddress.Locked = False
    txtEmegencyName.Locked = False
    txtEmegencyRelation.Locked = False
    txtEmegencyAddress.Locked = False
    txtEmegencyContact.Locked = False
    txtNoDependent.Locked = False
    
    cmbGender.Locked = False
    cmbOwnedHouse.Locked = False
    cmbRent.Locked = False
    cmbLivingParents.Locked = False
    cmbCivilStatus.Locked = False
    cmbTaxStatus.Locked = False
End If
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
cmdCancelAlphalist_Click
End Sub

Private Sub b8TitleBar2_CLoseClick()
cmdCancel_Click
End Sub

Private Sub cmbCivilStatus_Click()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    txtTmpCivilStatus.Text = IIf(cmbCivilStatus.ListIndex = 0, "SINGLE", IIf(cmbCivilStatus.ListIndex = 1, "MARRIED", IIf(cmbCivilStatus.ListIndex = 2, "WIDOWED", IIf(cmbCivilStatus.ListIndex = 3, "WIDOWER", ""))))
    If cmbCivilStatus.ListIndex = 0 Then txtDateMarriage.Text = ""
End If
End Sub

Private Sub cmbCivilStatus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtDateMarriage.SetFocus
End If
End Sub

Private Sub cmbGender_Click()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    txtTmpGender.Text = IIf(cmbGender.ListIndex = 0, "MALE", IIf(cmbGender.ListIndex = 1, "FEMALE", ""))
End If
End Sub


Private Sub cmbGender_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtBirthPlace.SetFocus
End If
End Sub

Private Sub cmbLivingParents_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
'    cmbCivilStatus.SetFocus
    cmbRent.SetFocus
End If
End Sub

Private Sub cmbOwnedHouse_Click()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If cmbOwnedHouse.ListIndex = 1 Then
        cmbRent.ListIndex = 0
        cmbLivingParents.ListIndex = 0
    End If
End If
End Sub

Private Sub cmbOwnedHouse_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
'    cmbRent.SetFocus
    cmbLivingParents.SetFocus
End If
End Sub

Private Sub cmbOwnedHouse_LostFocus()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If cmbOwnedHouse.ListIndex = 1 Then
        cmbRent.ListIndex = 0
        cmbLivingParents.ListIndex = 0
    End If
End If
End Sub

Private Sub cmbRent_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtBirthDate.SetFocus
End If
End Sub

Private Sub cmbTaxStatus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtNoDependent.SetFocus
End If
End Sub

Private Sub cmdCancel_Click()
picToolbar.Enabled = True
picMain.Enabled = True
picSearch.Visible = False
End Sub

Private Sub cmdCancelAlphalist_Click()
picToolbar.Enabled = True
picMain.Enabled = True
picAlphalist.Visible = False
End Sub

Private Sub cmdOK_Click()
If lstResult.ListIndex = -1 Then Exit Sub
BROWSER lstResult.ItemData(lstResult.ListIndex), "is_FIND"
cmdCancel_Click
End Sub

Private Sub cmdOKAlphalist_Click()
If IsDate(txtAsOf.Text) = False Then MsgBox "Please Supply a Valid Date!                  ", vbCritical, "Error...": txtAsOf.SetFocus: Exit Sub

MainForm.CommonDialog1.CancelError = True
On Error GoTo ErrorHandler
MainForm.CommonDialog1.DialogTitle = "Save"
MainForm.CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
MainForm.CommonDialog1.ShowSave
FileName_xls = Trim(MainForm.CommonDialog1.Filename)



txtAsOf.Text = Format(FormatDateTime(txtAsOf.Text, vbShortDate), "mm/dd/yyyy")
picAlphalist.Visible = False
picToolbar.Enabled = False
picMain.Enabled = False
picProgress.BackColor = &HFFFFFF
picProgressReport.ZOrder 0
picProgressReport.Visible = True
DoEvents
TimerAlphaActive.Enabled = True

Exit Sub
ErrorHandler:
Exit Sub
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
    Case vbKeyF11:      BROWSER GetSetting(App.EXEName, "ProfileInformation", "ProfileInfo", ""), "is_LOAD"
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "ProfileInformation", "ProfileInfo", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "ProfileInformation", "ProfileInfo", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "ProfileInformation", "ProfileInfo", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "ProfileInformation", "ProfileInfo", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
'Me.Caption = "Personnel Information"
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption

Me.Top = (MainForm.Height - Me.Height) / 20
Me.Left = (MainForm.Width - Me.Width) / 5

txtTmpFullName.ForeColor = 16711680
txtTmpBDay.ForeColor = 16711680
txtTmpAge.ForeColor = 16711680
txtTmpGender.ForeColor = 16711680
txtTmpCivilStatus.ForeColor = 16711680
txtTmpAddress.ForeColor = 16711680
txtTmpContact.ForeColor = 16711680
txtTmpHeight.ForeColor = 16711680
txtTmpWeight.ForeColor = 16711680
txtTmpSSS.ForeColor = 16711680
txtTmpPHIC.ForeColor = 16711680
txtTmpPagibig.ForeColor = 16711680
txtTmpTIN.ForeColor = 16711680
txtTmpDriverLic.ForeColor = 16711680
txtTmpIDNumber.ForeColor = 16711680
txtTmpStatus.ForeColor = 16711680
txtTmpBloodType.ForeColor = 16711680

With cmbOwnedHouse
    .AddItem "NO"
    .AddItem "YES"
End With
With cmbRent
    .AddItem "NO"
    .AddItem "YES"
End With
With cmbLivingParents
    .AddItem "NO"
    .AddItem "YES"
End With
With cmbGender
    .AddItem "MALE"
    .AddItem "FEMALE"
End With
With cmbCivilStatus
    .Clear
    .AddItem "SINGLE"
    .AddItem "MARRIED"
    .AddItem "WIDOWED"
    .AddItem "WIDOWER"
End With

With cmbTaxStatus
    .Clear
    s = "SELECT tbl_Personnel_TaxStatus.* " & _
        " FROM tbl_Personnel_TaxStatus " & _
        " ORDER BY PK"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        .AddItem rs!TaxStatus
        .ItemData(.NewIndex) = rs!PK
        rs.MoveNext
    Wend
    rs.Close
End With

ListFocus = 0
XTab1.ActiveTab = 4

CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
ListTrans = is_LstRefresh
BROWSER GetSetting(App.EXEName, "ProfileInformation", "ProfileInfo", ""), "is_LOAD"
If Trim(txtTmpFullName.Text) = "" Then BROWSER GetSetting(App.EXEName, "ProfileInformation", "ProfileInfo", ""), "is_HOME"

tmp = SetWindowLong(txtLastName.hwnd, GWL_STYLE, GetWindowLong(txtLastName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtFirstName.hwnd, GWL_STYLE, GetWindowLong(txtFirstName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtMiddleName.hwnd, GWL_STYLE, GetWindowLong(txtMiddleName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPresentAddress.hwnd, GWL_STYLE, GetWindowLong(txtPresentAddress.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtBirthDate.hwnd, GWL_STYLE, GetWindowLong(txtBirthDate.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtBirthPlace.hwnd, GWL_STYLE, GetWindowLong(txtBirthPlace.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtReligion.hwnd, GWL_STYLE, GetWindowLong(txtReligion.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtDateMarriage.hwnd, GWL_STYLE, GetWindowLong(txtDateMarriage.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtNationality.hwnd, GWL_STYLE, GetWindowLong(txtNationality.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtBloodType.hwnd, GWL_STYLE, GetWindowLong(txtBloodType.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtContactNumber.hwnd, GWL_STYLE, GetWindowLong(txtContactNumber.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSSSNumber.hwnd, GWL_STYLE, GetWindowLong(txtSSSNumber.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtTIN.hwnd, GWL_STYLE, GetWindowLong(txtTIN.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtDriverLicense.hwnd, GWL_STYLE, GetWindowLong(txtDriverLicense.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPHICNumber.hwnd, GWL_STYLE, GetWindowLong(txtPHICNumber.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPagIbigNumber.hwnd, GWL_STYLE, GetWindowLong(txtPagIbigNumber.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSpouseName.hwnd, GWL_STYLE, GetWindowLong(txtSpouseName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSpouseBDay.hwnd, GWL_STYLE, GetWindowLong(txtSpouseBDay.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSpouseOccupation.hwnd, GWL_STYLE, GetWindowLong(txtSpouseOccupation.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSpouseAddress.hwnd, GWL_STYLE, GetWindowLong(txtSpouseAddress.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtFatherName.hwnd, GWL_STYLE, GetWindowLong(txtFatherName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtFatherBDay.hwnd, GWL_STYLE, GetWindowLong(txtFatherBDay.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtFatherOccupation.hwnd, GWL_STYLE, GetWindowLong(txtFatherOccupation.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtFatherAddress.hwnd, GWL_STYLE, GetWindowLong(txtFatherAddress.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtMotherName.hwnd, GWL_STYLE, GetWindowLong(txtMotherName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtMotherBDay.hwnd, GWL_STYLE, GetWindowLong(txtMotherBDay.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtMotherOccupation.hwnd, GWL_STYLE, GetWindowLong(txtMotherOccupation.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtMotherAddress.hwnd, GWL_STYLE, GetWindowLong(txtMotherAddress.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtElemName.hwnd, GWL_STYLE, GetWindowLong(txtElemName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtElemInclusiveDate.hwnd, GWL_STYLE, GetWindowLong(txtElemInclusiveDate.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtElemCourse.hwnd, GWL_STYLE, GetWindowLong(txtElemCourse.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtElemAddress.hwnd, GWL_STYLE, GetWindowLong(txtElemAddress.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtHiSchoolName.hwnd, GWL_STYLE, GetWindowLong(txtHiSchoolName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtHiSchoolInclusiveDate.hwnd, GWL_STYLE, GetWindowLong(txtHiSchoolInclusiveDate.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtHiSchoolCourse.hwnd, GWL_STYLE, GetWindowLong(txtHiSchoolCourse.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtHiSchoolAddress.hwnd, GWL_STYLE, GetWindowLong(txtHiSchoolAddress.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtCollegeName.hwnd, GWL_STYLE, GetWindowLong(txtCollegeName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtCollegeInclusiveDate.hwnd, GWL_STYLE, GetWindowLong(txtCollegeInclusiveDate.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtCollegeCourse.hwnd, GWL_STYLE, GetWindowLong(txtCollegeCourse.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtCollegeAddress.hwnd, GWL_STYLE, GetWindowLong(txtCollegeAddress.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPostName.hwnd, GWL_STYLE, GetWindowLong(txtPostName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPostInclusiveDate.hwnd, GWL_STYLE, GetWindowLong(txtPostInclusiveDate.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPostCourse.hwnd, GWL_STYLE, GetWindowLong(txtPostCourse.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtPostAddress.hwnd, GWL_STYLE, GetWindowLong(txtPostAddress.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSkills.hwnd, GWL_STYLE, GetWindowLong(txtSkills.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtOrgsClubs.hwnd, GWL_STYLE, GetWindowLong(txtOrgsClubs.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtNotRelatedName.hwnd, GWL_STYLE, GetWindowLong(txtNotRelatedName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtNotRelatedContact.hwnd, GWL_STYLE, GetWindowLong(txtNotRelatedContact.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtNotRelatedAddress.hwnd, GWL_STYLE, GetWindowLong(txtNotRelatedAddress.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtRelativeCompanyName.hwnd, GWL_STYLE, GetWindowLong(txtRelativeCompanyName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtRelativeCompanyContact.hwnd, GWL_STYLE, GetWindowLong(txtRelativeCompanyContact.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtRelativeCompanyAddress.hwnd, GWL_STYLE, GetWindowLong(txtRelativeCompanyAddress.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtEmegencyName.hwnd, GWL_STYLE, GetWindowLong(txtEmegencyName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtEmegencyRelation.hwnd, GWL_STYLE, GetWindowLong(txtEmegencyRelation.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtEmegencyAddress.hwnd, GWL_STYLE, GetWindowLong(txtEmegencyAddress.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtEmegencyContact.hwnd, GWL_STYLE, GetWindowLong(txtEmegencyContact.hwnd, GWL_STYLE) Or ES_UPPERCASE)

tmp = SetWindowLong(txtBroSisName.hwnd, GWL_STYLE, GetWindowLong(txtBroSisName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtBroSisBDay.hwnd, GWL_STYLE, GetWindowLong(txtBroSisBDay.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtBroSisOccupation.hwnd, GWL_STYLE, GetWindowLong(txtBroSisOccupation.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtBroSisAddress.hwnd, GWL_STYLE, GetWindowLong(txtBroSisAddress.hwnd, GWL_STYLE) Or ES_UPPERCASE)

tmp = SetWindowLong(txtChildName.hwnd, GWL_STYLE, GetWindowLong(txtChildName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtChildBDay.hwnd, GWL_STYLE, GetWindowLong(txtChildBDay.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtChildOccupation.hwnd, GWL_STYLE, GetWindowLong(txtChildOccupation.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtChildAddress.hwnd, GWL_STYLE, GetWindowLong(txtChildAddress.hwnd, GWL_STYLE) Or ES_UPPERCASE)

tmp = SetWindowLong(txtTrainingTitle.hwnd, GWL_STYLE, GetWindowLong(txtTrainingTitle.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtTrainingDates.hwnd, GWL_STYLE, GetWindowLong(txtTrainingDates.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtTrainingVenue.hwnd, GWL_STYLE, GetWindowLong(txtTrainingVenue.hwnd, GWL_STYLE) Or ES_UPPERCASE)

tmp = SetWindowLong(txtEmploymentCompany.hwnd, GWL_STYLE, GetWindowLong(txtEmploymentCompany.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtEmploymentPosition.hwnd, GWL_STYLE, GetWindowLong(txtEmploymentPosition.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtEmploymentDates.hwnd, GWL_STYLE, GetWindowLong(txtEmploymentDates.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtEmploymentAddress.hwnd, GWL_STYLE, GetWindowLong(txtEmploymentAddress.hwnd, GWL_STYLE) Or ES_UPPERCASE)

tmp = SetWindowLong(txtSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picSearch.Visible = True Then Cancel = -1
If picBrotherSisterSLine.Visible = True Then Cancel = -1
If picChildrenSLine.Visible = True Then Cancel = -1
If picTrainingsSLine.Visible = True Then Cancel = -1
If picEmploymentSLine.Visible = True Then Cancel = -1
If picAlphalist.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub imgPicture_DblClick()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    MainForm.CommonDialog1.CancelError = True
    On Error GoTo ErrorHandler
'    Mainform.CommonDialog1.Filter = "Image Files (*.jpg)|*.jpg"
'    Mainform.CommonDialog1.Filter = "JPG|*.JPG;*.JPEG;*.JPE|BMP|*.BMP;*.RLE;*.DIB|GIF|*.GIF|PNG|*.PNG|TIFF|*.TIF;*.TIFF"
    MainForm.CommonDialog1.Filter = "Image Files|*.JPG;*.JPEG;*.JPE;*.BMP;*.RLE;*.DIB;*.GIF;*.PNG;*.TIF;*.TIFF"
    MainForm.CommonDialog1.ShowOpen
    Filename = Trim(MainForm.CommonDialog1.Filename)
'    If ((FileLen(Filename) \ 1024) + 1) > 50 Then
'        MsgBox "Image is too large please reduce the size to 50kb or below!          ", vbCritical, "Error..."
'        Exit Sub
'    End If
'    MsgBox CDbl(IMAGEFILESIZE(Date))
    If ((FileLen(Filename) \ 1024) + 1) > CDbl(IMAGEFILESIZE(Date)) Then
        MsgBox "Image is too large please reduce the size to " & IMAGEFILESIZE(Date) & "kb or below!          ", vbCritical, "Error..."
        Exit Sub
    End If
    txtPicturePath.Text = Filename
    imgPicture.Picture = LoadPicture(Filename)
End If
Exit Sub
ErrorHandler:
Exit Sub
End Sub

Private Sub lstBroSis_GotFocus()
ListFocus = 1
ListTrans = is_LstRefresh
ListRow = lstBroSis.SelectedItem.Index
End Sub

Private Sub lstBroSis_ItemClick(ByVal Item As MSComctlLib.ListItem)
ListRow = lstBroSis.SelectedItem.Index
End Sub

Private Sub lstBroSis_LostFocus()
ListFocus = 0
End Sub

Private Sub lstChildren_GotFocus()
ListFocus = 2
ListTrans = is_LstRefresh
ListRow = lstChildren.SelectedItem.Index
End Sub

Private Sub lstChildren_ItemClick(ByVal Item As MSComctlLib.ListItem)
ListRow = lstChildren.SelectedItem.Index
End Sub

Private Sub lstChildren_LostFocus()
ListFocus = 0
End Sub

Private Sub lstEmployment_GotFocus()
ListFocus = 4
ListTrans = is_LstRefresh
ListRow = lstEmployment.SelectedItem.Index
End Sub

Private Sub lstEmployment_ItemClick(ByVal Item As MSComctlLib.ListItem)
ListRow = lstEmployment.SelectedItem.Index
End Sub

Private Sub lstEmployment_LostFocus()
ListFocus = 0
End Sub

Private Sub lstResult_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub lstTraining_GotFocus()
ListFocus = 3
ListTrans = is_LstRefresh
ListRow = lstTraining.SelectedItem.Index
End Sub

Private Sub lstTraining_ItemClick(ByVal Item As MSComctlLib.ListItem)
ListRow = lstTraining.SelectedItem.Index
End Sub

Private Sub lstTraining_LostFocus()
ListFocus = 0
End Sub

Private Sub TimerActive_Timer()
TimerActive.Enabled = False
i = 0
picProgress.BackColor = &HFFFFFF
ConnOmega.Execute "DELETE FROM tbl_Personnel_Active_Inactive_Report WHERE (LogInName = '" & gbl_UserName & "')"
s = "sp_Personnel_Active_Inactive_Report (1,'" & FormatDateTime(Date, vbShortDate) & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount = 0 Then Exit Sub
DoEvents
picToolbar.Enabled = False
picMain.Enabled = False
picProgressReport.ZOrder 0
picProgressReport.Visible = True
While Not rs.EOF
    DoEvents
    i = i + 1
    
    t = "SELECT tbl_Personnel_Action.Division, tbl_Personnel_Action.Dept, tbl_Personnel_Department.DepartmentName, " & _
        " tbl_Personnel_Action.EmpStatus, tbl_Personnel_EmploymentStatus.StatusName, tbl_Personnel_Action.Positions, " & _
        " tbl_Personnel_Position.PositionName , tbl_Personnel_Action.EffectivityDate, tbl_Personnel_Action.Remarks " & _
        " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
        " tbl_Personnel_Position ON tbl_Personnel_Action.Positions = tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_Department ON tbl_Personnel_Action.Dept = tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_EmploymentStatus ON tbl_Personnel_Action.EmpStatus = tbl_Personnel_EmploymentStatus.PK " & _
        " WHERE (tbl_Personnel_Action.PK = " & rs!ActionMemo & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Active_Inactive_Report " & _
                          " (LogInName, Division, DivisionName, Department, DepartmentName, StatusKey, StatusName, " & _
                          " PositionKey, PositionName, EmpKey, IDNumber, EmployeeName) " & _
                          " VALUES ('" & gbl_UserName & "', " & rt!Division & ", '" & IIf(rt!Division = 1, "CLUB HOUSE", "MAINTENANCE") & "', " & _
                          " " & rt!Dept & ", '" & FORMATSQL(rt!DepartmentName) & "', " & rt!EmpStatus & ", " & _
                          " '" & FORMATSQL(rt!StatusName) & "', " & rt!Positions & ", '" & FORMATSQL(rt!PositionName) & "', " & _
                          " " & rs!PK & ", '" & FORMATSQL(rs!IDNumber) & "', '" & FORMATSQL(rs!EmployeeName) & "')"
    End If
    rt.Close
    
    UpdateProgress picProgress, i / rs.RecordCount
    rs.MoveNext
Wend
rs.Close
picProgressReport.Visible = False
picToolbar.Enabled = True
picMain.Enabled = True

s = "SELECT tbl_Personnel_Active_Inactive_Report.* " & _
    " FROM tbl_Personnel_Active_Inactive_Report " & _
    " WHERE (LogInName = '" & gbl_UserName & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
rs.Requery
rs.Close

frmCrystalReportViewer.PRINT_ACTIVE_EMPLOYEE gbl_CompanyName, gbl_UserName
If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show

End Sub

Private Sub TimerAlphaActive_Timer()
TimerAlphaActive.Enabled = False

WorkbookName = CStr(FileName_xls)

DoEvents
'picToolbar.Enabled = False
'picMain.Enabled = False
'picProgressReport.ZOrder 0
'picProgressReport.Visible = True

Screen.MousePointer = vbHourglass

iWorkSheet = 1
Set xlsApp = CreateObject("Excel.Application")
xlsApp.Visible = False
xlsApp.Workbooks.Add
xlsApp.DisplayAlerts = False
xlsApp.Workbooks(1).Sheets(2).Delete
xlsApp.Workbooks(1).Sheets(2).Delete
xlsApp.Workbooks(1).Sheets(iWorkSheet).Activate
xlsApp.Workbooks(1).Sheets(iWorkSheet).Name = "Alphalist"

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = gbl_CompanyName
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 12
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = gbl_CompanyAddress1
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 9
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = gbl_CompanyAddress2
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 9
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True


RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "As of " & Format(txtAsOf.Text, "mmmm dd, yyyy")
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 9
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True


RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 9
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False


i = 0
picProgress.BackColor = &HFFFFFF
s = "sp_Personnel_Alphalist(1, '" & FormatDateTime(txtAsOf.Text, vbShortDate) & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    RowCnt = RowCnt + 1
    ColCnt = 0
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
    
    For k = 1 To rs.Fields.Count
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs.Fields(k - 1).Name
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
    Next k
    
    While Not rs.EOF
        DoEvents
        i = i + 1
        RowCnt = RowCnt + 1
        ColCnt = 0
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = i
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
        For k = 1 To rs.Fields.Count
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            If IsNumeric(rs.Fields(k - 1).Value) Then
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "@"
            End If
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs.Fields(k - 1).Value
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            If IsDate(rs.Fields(k - 1).Value) Then
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "mm/dd/yyyy"
            End If
            
        Next k
        
        UpdateProgress picProgress, i / rs.RecordCount
        
        rs.MoveNext
    Wend
End If
rs.Close

SAVING:
On Error GoTo err_saving:
If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

xlsApp.Visible = True

Screen.MousePointer = vbDefault

picProgress.BackColor = &HFFFFFF
picProgressReport.Visible = False
picToolbar.Enabled = True
picMain.Enabled = True

Exit Sub
err_saving:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING:

Exit Sub
ErrorHandler:
Screen.MousePointer = vbDefault
Exit Sub
End Sub

Private Sub TimerHeadCount_Timer()
TimerHeadCount.Enabled = False
i = 0
picProgress.BackColor = &HFFFFFF
ConnOmega.Execute "DELETE FROM tbl_Personnel_HeadCount WHERE (LogInName = '" & gbl_UserName & "')"
s = "sp_Personnel_Active_Inactive_Report (1,'" & FormatDateTime(Date, vbShortDate) & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount = 0 Then Exit Sub

DoEvents
picToolbar.Enabled = False
picMain.Enabled = False
picProgressReport.ZOrder 0
picProgressReport.Visible = True
TableName = "tmp_" & gbl_UserName & "_Personnel_HeadCount"

Columns = ""
Columns = Columns & "|DivisionKey:int:NOT NULL:DEFAULT(0)"
Columns = Columns & "|DepartmentKey:int:NOT NULL:DEFAULT(0)"
Columns = Columns & "|StatusKey:int:NOT NULL:DEFAULT(0)"
Columns = Columns & "|EmployeeKey:int:NOT NULL:DEFAULT(0)"
CreateTable gbl_Database, TableName, Columns

While Not rs.EOF
    DoEvents
    i = i + 1
    t = "SELECT tbl_Personnel_Action.Division, tbl_Personnel_Action.Dept, tbl_Personnel_Department.DepartmentName, " & _
        " tbl_Personnel_Action.EmpStatus, tbl_Personnel_EmploymentStatus.StatusName, tbl_Personnel_Action.Positions, " & _
        " tbl_Personnel_Position.PositionName , tbl_Personnel_Action.EffectivityDate, tbl_Personnel_Action.Remarks " & _
        " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
        " tbl_Personnel_Position ON tbl_Personnel_Action.Positions = tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_Department ON tbl_Personnel_Action.Dept = tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_EmploymentStatus ON tbl_Personnel_Action.EmpStatus = tbl_Personnel_EmploymentStatus.PK " & _
        " WHERE (tbl_Personnel_Action.PK = " & rs!ActionMemo & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        ConnOmega.Execute "INSERT INTO " & TableName & " " & _
                          " (DivisionKey, DepartmentKey, StatusKey, EmployeeKey) " & _
                          " VALUES (" & rt!Division & ", " & rt!Dept & ", " & _
                          " " & rt!EmpStatus & ", " & rs!PK & ")"
    End If
    rt.Close
    
    UpdateProgress picProgress, i / rs.RecordCount
    rs.MoveNext
Wend
rs.Close

s = "SELECT " & TableName & ".DivisionKey, " & TableName & ".DepartmentKey, " & _
    " tbl_Personnel_Department.DepartmentName, " & TableName & ".StatusKey, " & _
    " tbl_Personnel_EmploymentStatus.StatusName, COUNT(" & TableName & ".EmployeeKey) AS HeadCount " & _
    " FROM " & TableName & " LEFT OUTER JOIN " & _
    " tbl_Personnel_EmploymentStatus ON " & _
    " " & TableName & ".StatusKey = tbl_Personnel_EmploymentStatus.PK LEFT OUTER JOIN " & _
    " tbl_Personnel_Department ON " & TableName & ".DepartmentKey = tbl_Personnel_Department.PK " & _
    " GROUP BY " & TableName & ".DivisionKey, " & TableName & ".DepartmentKey, " & _
    " " & TableName & ".StatusKey, tbl_Personnel_Department.DepartmentName, " & _
    " tbl_Personnel_EmploymentStatus.StatusName " & _
    " ORDER BY " & TableName & ".DivisionKey, " & TableName & ".DepartmentKey, " & _
    " tbl_Personnel_EmploymentStatus.StatusName "
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    ConnOmega.Execute "INSERT INTO tbl_Personnel_HeadCount " & _
                      " (LogInName, DivKey, DivName, DeptKey, DeptName, StatusKey, StatusName, EmpCount) " & _
                      " VALUES ('" & gbl_UserName & "', " & rs!DivisionKey & ", " & _
                      " '" & IIf(rs!DivisionKey = 1, "CLUB HOUSE", "MAINTENANCE") & "', " & _
                      " " & rs!DepartmentKey & ", '" & FORMATSQL(rs!DepartmentName) & "', " & _
                      " " & rs!StatusKey & ", '" & FORMATSQL(rs!StatusName) & "', " & _
                      " " & rs!HeadCount & ")"
    rs.MoveNext
Wend
rs.Close

picProgressReport.Visible = False
picToolbar.Enabled = True
picMain.Enabled = True


s = "SELECT tbl_Personnel_HeadCount.* " & _
    " FROM tbl_Personnel_HeadCount " & _
    " WHERE (LogInName = '" & gbl_UserName & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
rs.Requery
rs.Close

frmCrystalReportViewer.PRINT_EMPLOYEE_HEADCOUNT gbl_CompanyName, gbl_UserName
If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show

End Sub

Private Sub TimerHistory_Timer()
TimerHistory.Enabled = False

End Sub

Private Sub TimerInactive_Timer()
TimerInactive.Enabled = False
i = 0
picProgress.BackColor = &HFFFFFF
ConnOmega.Execute "DELETE FROM tbl_Personnel_Active_Inactive_Report WHERE (LogInName = '" & gbl_UserName & "')"
s = "sp_Personnel_Active_Inactive_Report (2,'" & FormatDateTime(Date, vbShortDate) & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount = 0 Then Exit Sub
DoEvents
picToolbar.Enabled = False
picMain.Enabled = False
picProgressReport.ZOrder 0
picProgressReport.Visible = True
While Not rs.EOF
    DoEvents
    i = i + 1
    
    t = "SELECT tbl_Personnel_Action.Division, tbl_Personnel_Action.Dept, tbl_Personnel_Department.DepartmentName, " & _
        " tbl_Personnel_Action.EmpStatus, tbl_Personnel_EmploymentStatus.StatusName, tbl_Personnel_Action.Positions, " & _
        " tbl_Personnel_Position.PositionName , tbl_Personnel_Action.EffectivityDate, tbl_Personnel_Action.Remarks " & _
        " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
        " tbl_Personnel_Position ON tbl_Personnel_Action.Positions = tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_Department ON tbl_Personnel_Action.Dept = tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_EmploymentStatus ON tbl_Personnel_Action.EmpStatus = tbl_Personnel_EmploymentStatus.PK " & _
        " WHERE (tbl_Personnel_Action.PK = " & rs!ActionMemo & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Active_Inactive_Report " & _
                          " (LogInName, Division, DivisionName, Department, DepartmentName, StatusKey, StatusName, " & _
                          " PositionKey, PositionName, EmpKey, IDNumber, EmployeeName, EffecDate, Reason) " & _
                          " VALUES ('" & gbl_UserName & "', " & rt!Division & ", '" & IIf(rt!Division = 1, "CLUB HOUSE", "MAINTENANCE") & "', " & _
                          " " & rt!Dept & ", '" & FORMATSQL(rt!DepartmentName) & "', " & rt!EmpStatus & ", " & _
                          " '" & FORMATSQL(rt!StatusName) & "', " & rt!Positions & ", '" & FORMATSQL(rt!PositionName) & "', " & _
                          " " & rs!PK & ", '" & FORMATSQL(rs!IDNumber) & "', '" & FORMATSQL(rs!EmployeeName) & "', " & _
                          " '" & FormatDateTime(rt!EffectivityDate, vbShortDate) & "', '" & FORMATSQL(rt!Remarks) & "')"
    End If
    rt.Close
    
    UpdateProgress picProgress, i / rs.RecordCount
    rs.MoveNext
Wend
rs.Close
picProgressReport.Visible = False
picToolbar.Enabled = True
picMain.Enabled = True

s = "SELECT tbl_Personnel_Active_Inactive_Report.* " & _
    " FROM tbl_Personnel_Active_Inactive_Report " & _
    " WHERE (LogInName = '" & gbl_UserName & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
rs.Requery
rs.Close

frmCrystalReportViewer.PRINT_INACTIVE_EMPLOYEE gbl_CompanyName, gbl_UserName
If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "ProfileInformation", "ProfileInfo", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "ProfileInformation", "ProfileInfo", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "ProfileInformation", "ProfileInfo", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "ProfileInformation", "ProfileInfo", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":   PRESS_F9
    Case "Refresh": BROWSER GetSetting(App.EXEName, "ProfileInformation", "ProfileInfo", ""), "is_LOAD"
    Case "Close":   PRESS_ESCAPE
    Case Else: Exit Sub
End Select
End Sub

Private Sub txtAge_GotFocus()
HTEXT txtAge
End Sub

Private Sub txtAsof_GotFocus()
HTEXT txtAsOf
End Sub

Private Sub txtAsof_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAlphalist_Click
End Sub

Private Sub txtBirthDate_Change()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If IsDate(txtBirthDate.Text) = True Then
        txtTmpBDay = Format(FormatDateTime(txtBirthDate.Text, vbShortDate), "mm/dd/yyyy")
        txtAge.Text = Get_Age(FormatDateTime(txtBirthDate.Text, vbShortDate), FormatDateTime(Date, vbShortDate))
        txtTmpAge.Text = Get_Age(FormatDateTime(txtBirthDate.Text, vbShortDate), FormatDateTime(Date, vbShortDate))
    Else
        txtTmpBDay.Text = ""
        txtAge.Text = ""
        txtTmpAge.Text = ""
    End If
End If
End Sub

Private Sub txtBirthDate_GotFocus()
HTEXT txtBirthDate
End Sub

Private Sub txtBirthDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    cmbGender.SetFocus
ElseIf KeyCode = vbKeyUp Then
    cmbRent.SetFocus
End If
End Sub

Private Sub txtBirthPlace_GotFocus()
HTEXT txtBirthPlace
End Sub

Private Sub txtBirthPlace_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtReligion.SetFocus
ElseIf KeyCode = vbKeyUp Then
    cmbGender.SetFocus
End If
End Sub

Private Sub txtBloodType_Change()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    txtTmpBloodType.Text = txtBloodType.Text
End If
End Sub

Private Sub txtBroSisAddress_Change()
If ListTrans = is_LstAdding Or _
ListTrans = is_LstEditting Then
    With lstBroSis.ListItems
        .Item(ListRow).SubItems(5) = Trim(txtBroSisAddress.Text)
    End With
End If
End Sub

Private Sub txtBroSisAddress_GotFocus()
HTEXT txtBroSisAddress
End Sub

Private Sub txtBroSisAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    picBrotherSisterSLine.Visible = False
    picMain.Enabled = True
    picToolbar.Enabled = True
    lstBroSis.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtBroSisOccupation.SetFocus
End If
End Sub

Private Sub txtBroSisBDay_Change()
If ListTrans = is_LstAdding Or _
ListTrans = is_LstEditting Then
    With lstBroSis.ListItems
        If IsDate(txtBroSisBDay.Text) = False Then
            .Item(ListRow).SubItems(3) = " "
        Else
            .Item(ListRow).SubItems(3) = Format(FormatDateTime(txtBroSisBDay.Text, vbShortDate), "mm/dd/yyyy")
        End If
    End With
End If
End Sub

Private Sub txtBroSisBDay_GotFocus()
HTEXT txtBroSisBDay
End Sub

Private Sub txtBroSisBDay_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtBroSisOccupation.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtBroSisName.SetFocus
End If
End Sub

Private Sub txtBroSisBDay_LostFocus()
If IsDate(txtBroSisBDay.Text) = True Then
    txtBroSisBDay.Text = Format(FormatDateTime(txtBroSisBDay.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub

Private Sub txtBroSisName_Change()
If ListTrans = is_LstAdding Or _
ListTrans = is_LstEditting Then
    With lstBroSis.ListItems
        .Item(ListRow).SubItems(2) = Trim(txtBroSisName.Text)
    End With
End If
End Sub

Private Sub txtBroSisName_GotFocus()
HTEXT txtBroSisName
End Sub

Private Sub txtBroSisName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtBroSisBDay.SetFocus
End If
End Sub

Private Sub txtBroSisOccupation_Change()
If ListTrans = is_LstAdding Or _
ListTrans = is_LstEditting Then
    With lstBroSis.ListItems
        .Item(ListRow).SubItems(4) = Trim(txtBroSisOccupation.Text)
    End With
End If
End Sub

Private Sub txtBroSisOccupation_GotFocus()
HTEXT txtBroSisOccupation
End Sub

Private Sub txtBroSisOccupation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtBroSisAddress.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtBroSisBDay.SetFocus
End If
End Sub

Private Sub txtChildAddress_Change()
If ListTrans = is_LstAdding Or _
ListTrans = is_LstEditting Then
    With lstChildren.ListItems
        .Item(ListRow).SubItems(5) = Trim(txtChildAddress.Text)
    End With
End If
End Sub

Private Sub txtChildAddress_GotFocus()
HTEXT txtChildAddress
End Sub

Private Sub txtChildAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    picChildrenSLine.Visible = False
    picMain.Enabled = True
    picToolbar.Enabled = True
    lstChildren.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtChildOccupation.SetFocus
End If
End Sub

Private Sub txtChildBDay_Change()
If ListTrans = is_LstAdding Or _
ListTrans = is_LstEditting Then
    With lstChildren.ListItems
        If IsDate(txtChildBDay.Text) = False Then
            .Item(ListRow).SubItems(3) = " "
        Else
            .Item(ListRow).SubItems(3) = Format(FormatDateTime(txtChildBDay.Text, vbShortDate), "mm/dd/yyyy")
        End If
    End With
End If
End Sub

Private Sub txtChildBDay_GotFocus()
HTEXT txtChildBDay
End Sub

Private Sub txtChildBDay_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtChildOccupation.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtChildName.SetFocus
End If
End Sub

Private Sub txtChildBDay_LostFocus()
If IsDate(txtChildBDay.Text) = True Then
    txtChildBDay.Text = Format(FormatDateTime(txtChildBDay.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub

Private Sub txtChildName_Change()
If ListTrans = is_LstAdding Or _
ListTrans = is_LstEditting Then
    With lstChildren.ListItems
        .Item(ListRow).SubItems(2) = Trim(txtChildName.Text)
    End With
End If
End Sub

Private Sub txtChildName_GotFocus()
HTEXT txtChildName
End Sub

Private Sub txtChildName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtChildBDay.SetFocus
End If
End Sub

Private Sub txtChildOccupation_Change()
If ListTrans = is_LstAdding Or _
ListTrans = is_LstEditting Then
    With lstChildren.ListItems
        .Item(ListRow).SubItems(4) = Trim(txtChildOccupation.Text)
    End With
End If
End Sub

Private Sub txtChildOccupation_GotFocus()
HTEXT txtChildOccupation
End Sub

Private Sub txtChildOccupation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtChildAddress.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtChildBDay.SetFocus
End If
End Sub

Private Sub txtCollegeAddress_GotFocus()
HTEXT txtCollegeAddress
End Sub

Private Sub txtCollegeAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtPostName.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtCollegeCourse.SetFocus
End If
End Sub

Private Sub txtCollegeCourse_GotFocus()
HTEXT txtCollegeCourse
End Sub

Private Sub txtCollegeCourse_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtCollegeAddress.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtCollegeInclusiveDate.SetFocus
End If
End Sub

Private Sub txtCollegeInclusiveDate_GotFocus()
HTEXT txtCollegeInclusiveDate
End Sub

Private Sub txtCollegeInclusiveDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtCollegeCourse.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtCollegeName.SetFocus
End If
End Sub

Private Sub txtCollegeName_GotFocus()
HTEXT txtCollegeName
End Sub

Private Sub txtCollegeName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtCollegeInclusiveDate.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtHiSchoolAddress.SetFocus
End If
End Sub

Private Sub txtContactNumber_Change()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    txtTmpContact.Text = Trim(txtContactNumber.Text)
End If
End Sub

Private Sub txtContactNumber_GotFocus()
HTEXT txtContactNumber
End Sub

Private Sub txtContactNumber_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtDriverLicense.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtNationality.SetFocus
End If
End Sub

Private Sub txtDateMarriage_GotFocus()
HTEXT txtDateMarriage
End Sub

Private Sub txtDateMarriage_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtHeight.SetFocus
ElseIf KeyCode = vbKeyUp Then
    cmbCivilStatus.SetFocus
End If
End Sub

Private Sub txtDateMarriage_LostFocus()
If IsDate(txtDateMarriage.Text) = True Then
    txtDateMarriage.Text = Format(FormatDateTime(txtDateMarriage.Text, vbShortDate), "mm/dd/yyyy")
Else
    txtDateMarriage.Text = ""
End If
End Sub

Private Sub txtDriverLicense_Change()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    txtTmpDriverLic.Text = Trim(txtDriverLicense.Text)
End If
End Sub

Private Sub txtDriverLicense_GotFocus()
HTEXT txtDriverLicense
End Sub

Private Sub txtDriverLicense_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtTIN.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtContactNumber.SetFocus
End If
End Sub

Private Sub txtElemAddress_GotFocus()
HTEXT txtElemAddress
End Sub

Private Sub txtElemAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtHiSchoolName.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtElemCourse.SetFocus
End If
End Sub

Private Sub txtElemCourse_GotFocus()
HTEXT txtElemCourse
End Sub

Private Sub txtElemCourse_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtElemAddress.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtElemInclusiveDate.SetFocus
End If
End Sub

Private Sub txtElemInclusiveDate_GotFocus()
HTEXT txtElemInclusiveDate
End Sub

Private Sub txtElemInclusiveDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtElemCourse.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtElemName.SetFocus
End If
End Sub

Private Sub txtElemName_GotFocus()
HTEXT txtElemName
End Sub

Private Sub txtElemName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtElemInclusiveDate.SetFocus
ElseIf KeyCode = vbKeyUp Then

End If
End Sub

Private Sub txtEmegencyAddress_GotFocus()
HTEXT txtEmegencyAddress
End Sub

Private Sub txtEmegencyAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtEmegencyContact.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtEmegencyRelation.SetFocus
End If
End Sub

Private Sub txtEmegencyContact_GotFocus()
HTEXT txtEmegencyContact
End Sub

Private Sub txtEmegencyContact_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    
ElseIf KeyCode = vbKeyDown Then
    
ElseIf KeyCode = vbKeyUp Then
    txtEmegencyAddress.SetFocus
End If
End Sub

Private Sub txtEmegencyName_GotFocus()
HTEXT txtEmegencyName
End Sub

Private Sub txtEmegencyName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtEmegencyRelation.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtRelativeCompanyAddress.SetFocus
End If
End Sub

Private Sub txtEmegencyRelation_GotFocus()
HTEXT txtEmegencyRelation
End Sub

Private Sub txtEmegencyRelation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtEmegencyAddress.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtEmegencyName.SetFocus
End If
End Sub

Private Sub txtEmploymentAddress_Change()
If ListTrans = is_LstAdding Or _
ListTrans = is_LstEditting Then
    With lstEmployment.ListItems
        .Item(ListRow).SubItems(6) = Trim(txtEmploymentAddress.Text)
    End With
End If
End Sub

Private Sub txtEmploymentAddress_GotFocus()
HTEXT txtEmploymentAddress
End Sub

Private Sub txtEmploymentAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    picEmploymentSLine.Visible = False
    picMain.Enabled = True
    picToolbar.Enabled = True
    lstEmployment.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtEmploymentDates.SetFocus
End If
End Sub

Private Sub txtEmploymentCompany_Change()
If ListTrans = is_LstAdding Or _
ListTrans = is_LstEditting Then
    With lstEmployment.ListItems
        .Item(ListRow).SubItems(2) = Trim(txtEmploymentCompany.Text)
    End With
End If
End Sub

Private Sub txtEmploymentCompany_GotFocus()
HTEXT txtEmploymentCompany
End Sub

Private Sub txtEmploymentCompany_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtEmploymentPosition.SetFocus
End If
End Sub

Private Sub txtEmploymentDates_Change()
If ListTrans = is_LstAdding Or _
ListTrans = is_LstEditting Then
    With lstEmployment.ListItems
        .Item(ListRow).SubItems(5) = Trim(txtEmploymentDates.Text)
    End With
End If
End Sub

Private Sub txtEmploymentDates_GotFocus()
HTEXT txtEmploymentDates
End Sub

Private Sub txtEmploymentDates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtEmploymentAddress.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtEmploymentSalary.SetFocus
End If
End Sub

Private Sub txtEmploymentPosition_Change()
If ListTrans = is_LstAdding Or _
ListTrans = is_LstEditting Then
    With lstEmployment.ListItems
        .Item(ListRow).SubItems(3) = Trim(txtEmploymentPosition.Text)
    End With
End If
End Sub

Private Sub txtEmploymentPosition_GotFocus()
HTEXT txtEmploymentPosition
End Sub

Private Sub txtEmploymentPosition_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtEmploymentSalary.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtEmploymentCompany.SetFocus
End If
End Sub

Private Sub txtEmploymentSalary_Change()
If ListTrans = is_LstAdding Or _
ListTrans = is_LstEditting Then
    With lstEmployment.ListItems
        .Item(ListRow).SubItems(4) = Format(RETURNTEXTVALUE(txtEmploymentSalary), "#,##0.00")
    End With
End If
End Sub

Private Sub txtEmploymentSalary_GotFocus()
HTEXT txtEmploymentSalary
End Sub

Private Sub txtEmploymentSalary_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtEmploymentDates.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtEmploymentPosition.SetFocus
End If
End Sub

Private Sub txtEmploymentSalary_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtFatherAddress_GotFocus()
HTEXT txtFatherAddress
End Sub

Private Sub txtFatherAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtMotherName.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtFatherOccupation.SetFocus
End If
End Sub

Private Sub txtFatherBDay_GotFocus()
HTEXT txtFatherBDay
End Sub

Private Sub txtFatherBDay_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtFatherOccupation.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtFatherName.SetFocus
End If
End Sub

Private Sub txtFatherName_GotFocus()
HTEXT txtFatherName
End Sub

Private Sub txtFatherName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtFatherBDay.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtSpouseAddress.SetFocus
End If
End Sub

Private Sub txtFatherOccupation_GotFocus()
HTEXT txtFatherOccupation
End Sub

Private Sub txtFatherOccupation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtFatherAddress.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtFatherBDay.SetFocus
End If
End Sub

Private Sub txtFirstName_Change()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    txtTmpFullName.Text = Trim(txtLastName.Text) & _
                          IIf(Trim(txtFirstName.Text) = "", "", ",  " & Trim(txtFirstName.Text)) & _
                          IIf(Trim(txtMiddleName.Text) = "", "", "  " & Trim(txtMiddleName.Text))
End If
End Sub

Private Sub txtFirstName_GotFocus()
HTEXT txtFirstName
End Sub

Private Sub txtFirstName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtMiddleName.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtLastName.SetFocus
End If
End Sub

Private Sub txtHeight_GotFocus()
HTEXT txtHeight
End Sub

Private Sub txtHeight_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtWeight.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtDateMarriage.SetFocus
End If
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtHiSchoolAddress_GotFocus()
HTEXT txtHiSchoolAddress
End Sub

Private Sub txtHiSchoolAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtCollegeName.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtHiSchoolCourse.SetFocus
End If
End Sub

Private Sub txtHiSchoolCourse_GotFocus()
HTEXT txtHiSchoolCourse
End Sub

Private Sub txtHiSchoolCourse_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtHiSchoolAddress.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtHiSchoolInclusiveDate.SetFocus
End If
End Sub

Private Sub txtHiSchoolInclusiveDate_GotFocus()
HTEXT txtHiSchoolInclusiveDate
End Sub

Private Sub txtHiSchoolInclusiveDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtHiSchoolCourse.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtHiSchoolName.SetFocus
End If
End Sub

Private Sub txtHiSchoolName_GotFocus()
HTEXT txtHiSchoolName
End Sub

Private Sub txtHiSchoolName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtHiSchoolInclusiveDate.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtElemAddress.SetFocus
End If
End Sub

Private Sub txtLastName_Change()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    txtTmpFullName.Text = Trim(txtLastName.Text) & _
                          IIf(Trim(txtFirstName.Text) = "", "", ",  " & Trim(txtFirstName.Text)) & _
                          IIf(Trim(txtMiddleName.Text) = "", "", "  " & Trim(txtMiddleName.Text))
End If
End Sub

Private Sub txtLastName_GotFocus()
HTEXT txtLastName
End Sub

Private Sub txtLastName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtFirstName.SetFocus
ElseIf KeyCode = vbKeyUp Then

End If
End Sub

Private Sub txtMiddleName_Change()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    txtTmpFullName.Text = Trim(txtLastName.Text) & _
                          IIf(Trim(txtFirstName.Text) = "", "", ",  " & Trim(txtFirstName.Text)) & _
                          IIf(Trim(txtMiddleName.Text) = "", "", "  " & Trim(txtMiddleName.Text))
End If
End Sub

Private Sub txtMiddleName_GotFocus()
HTEXT txtFirstName
End Sub

Private Sub txtMiddleName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtPresentAddress.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtFirstName.SetFocus
End If
End Sub

Private Sub txtMotherAddress_GotFocus()
HTEXT txtMotherAddress
End Sub

Private Sub txtMotherAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    lstBroSis.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtMotherOccupation.SetFocus
End If
End Sub

Private Sub txtMotherBDay_GotFocus()
HTEXT txtMotherBDay
End Sub

Private Sub txtMotherBDay_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtMotherOccupation.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtMotherName.SetFocus
End If
End Sub

Private Sub txtMotherName_GotFocus()
HTEXT txtMotherName
End Sub

Private Sub txtMotherName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtMotherBDay.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtFatherAddress.SetFocus
End If
End Sub

Private Sub txtMotherOccupation_GotFocus()
HTEXT txtMotherOccupation
End Sub

Private Sub txtMotherOccupation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtMotherAddress.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtMotherBDay.SetFocus
End If
End Sub

Private Sub txtNationality_GotFocus()
HTEXT txtNationality
End Sub

Private Sub txtNationality_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtContactNumber.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtWeight.SetFocus
End If
End Sub

Private Sub txtNoDependent_GotFocus()
HTEXT txtNoDependent
End Sub

Private Sub txtNoDependent_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    cmbIDNumber.SetFocus
ElseIf KeyCode = vbKeyUp Then
    cmbTaxStatus.SetFocus
End If
End Sub

Private Sub txtNoDependent_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtNotRelatedAddress_GotFocus()
HTEXT txtNotRelatedAddress
End Sub

Private Sub txtNotRelatedAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtRelativeCompanyName.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtNotRelatedContact.SetFocus
End If
End Sub

Private Sub txtNotRelatedContact_GotFocus()
HTEXT txtNotRelatedContact
End Sub

Private Sub txtNotRelatedContact_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtNotRelatedAddress.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtNotRelatedName.SetFocus
End If
End Sub

Private Sub txtNotRelatedName_GotFocus()
HTEXT txtNotRelatedName
End Sub

Private Sub txtNotRelatedName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtNotRelatedContact.SetFocus
ElseIf KeyCode = vbKeyUp Then
    lstEmployment.SetFocus
End If
End Sub

Private Sub txtOrgsClubs_GotFocus()
HTEXT txtOrgsClubs
End Sub

Private Sub txtOrgsClubs_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then

ElseIf KeyCode = vbKeyUp Then

End If
End Sub

Private Sub txtPagIbigNumber_Change()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    txtTmpPagibig.Text = Trim(txtPagIbigNumber.Text)
End If
End Sub

Private Sub txtPagIbigNumber_GotFocus()
HTEXT txtPagIbigNumber
End Sub

Private Sub txtPagIbigNumber_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
'    cmbLivingParents.SetFocus
    cmbCivilStatus.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtPHICNumber.SetFocus
End If
End Sub

Private Sub txtPHICNumber_Change()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    txtTmpPHIC.Text = Trim(txtPHICNumber.Text)
End If
End Sub

Private Sub txtPHICNumber_GotFocus()
HTEXT txtPHICNumber
End Sub

Private Sub txtPHICNumber_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtPagIbigNumber.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtSSSNumber.SetFocus
End If
End Sub

Private Sub txtPostAddress_GotFocus()
HTEXT txtPostAddress
End Sub

Private Sub txtPostAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtSkills.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtPostCourse.SetFocus
End If
End Sub

Private Sub txtPostCourse_GotFocus()
HTEXT txtPostCourse
End Sub

Private Sub txtPostCourse_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtPostAddress.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtPostInclusiveDate.SetFocus
End If
End Sub

Private Sub txtPostInclusiveDate_GotFocus()
HTEXT txtPostInclusiveDate
End Sub

Private Sub txtPostInclusiveDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtPostCourse.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtPostName.SetFocus
End If
End Sub

Private Sub txtPostName_GotFocus()
HTEXT txtPostName
End Sub

Private Sub txtPostName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtPostInclusiveDate.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtCollegeAddress.SetFocus
End If
End Sub

Private Sub txtPresentAddress_Change()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    txtTmpAddress.Text = Trim(txtPresentAddress.Text)
End If
End Sub

Private Sub txtPresentAddress_GotFocus()
HTEXT txtPresentAddress
End Sub

Private Sub txtPresentAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    cmbOwnedHouse.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtMiddleName.SetFocus
End If
End Sub

Private Sub txtRelativeCompanyAddress_GotFocus()
HTEXT txtRelativeCompanyAddress
End Sub

Private Sub txtRelativeCompanyAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtEmegencyName.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtRelativeCompanyContact.SetFocus
End If
End Sub

Private Sub txtRelativeCompanyContact_GotFocus()
HTEXT txtRelativeCompanyContact
End Sub

Private Sub txtRelativeCompanyContact_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtRelativeCompanyAddress.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtRelativeCompanyName.SetFocus
End If
End Sub

Private Sub txtRelativeCompanyName_GotFocus()
HTEXT txtRelativeCompanyName
End Sub

Private Sub txtRelativeCompanyName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then

ElseIf KeyCode = vbKeyUp Then

End If
End Sub

Private Sub txtReligion_GotFocus()
HTEXT txtReligion
End Sub

Private Sub txtReligion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtSSSNumber.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtBirthPlace.SetFocus
End If
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstResult.Clear: Exit Sub
lstResult.Clear
Select Case SearchType
    Case is_LName
        s = "SELECT PK, LastName + ',  ' + FirstName + '  ' + MiddleName AS EmployeeName " & _
            " From tbl_Personnel_Information " & _
            " WHERE (LastName LIKE '" & Trim(txtSearch.Text) & "%') " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
    Case is_FName
        s = "SELECT PK, LastName + ',  ' + FirstName + '  ' + MiddleName AS EmployeeName " & _
            " From tbl_Personnel_Information " & _
            " WHERE (FirstName LIKE '" & Trim(txtSearch.Text) & "%') " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
    Case is_MName
        s = "SELECT PK, LastName + ',  ' + FirstName + '  ' + MiddleName AS EmployeeName " & _
            " From tbl_Personnel_Information " & _
            " WHERE (MiddleName LIKE '" & Trim(txtSearch.Text) & "%') " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResult.AddItem rs!EmployeeName
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
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then lstResult.SetFocus
End Sub

Private Sub txtSkills_GotFocus()
HTEXT txtSkills
End Sub

Private Sub txtSkills_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    lstTraining.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtPostAddress.SetFocus
End If
End Sub

Private Sub txtSpouseAddress_GotFocus()
HTEXT txtSpouseAddress
End Sub

Private Sub txtSpouseAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtFatherName.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtSpouseOccupation.SetFocus
End If
End Sub

Private Sub txtSpouseBDay_GotFocus()
HTEXT txtSpouseBDay
End Sub

Private Sub txtSpouseBDay_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtSpouseOccupation.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtSpouseName.SetFocus
End If
End Sub

Private Sub txtSpouseName_GotFocus()
HTEXT txtSpouseName
End Sub

Private Sub txtSpouseName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtSpouseBDay.SetFocus
ElseIf KeyCode = vbKeyUp Then

End If
End Sub

Private Sub txtSpouseOccupation_GotFocus()
HTEXT txtSpouseOccupation
End Sub

Private Sub txtSpouseOccupation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtSpouseAddress.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtSpouseBDay.SetFocus
End If
End Sub

Private Sub txtSSSNumber_Change()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    txtTmpSSS.Text = Trim(txtSSSNumber.Text)
End If
End Sub

Private Sub txtSSSNumber_GotFocus()
HTEXT txtSSSNumber
End Sub

Private Sub txtSSSNumber_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtPHICNumber.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtReligion.SetFocus
End If
End Sub

Private Sub txtTin_Change()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    txtTmpTIN.Text = Trim(txtTIN.Text)
End If
End Sub

Private Sub txtTIN_GotFocus()
HTEXT txtTIN
End Sub

Private Sub txtTIN_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    cmbTaxStatus.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtDriverLicense.SetFocus
End If
End Sub

Private Sub txtTrainingDates_Change()
If ListTrans = is_LstAdding Or _
ListTrans = is_LstEditting Then
    With lstTraining.ListItems
        .Item(ListRow).SubItems(3) = Trim(txtTrainingDates.Text)
    End With
End If
End Sub

Private Sub txtTrainingDates_GotFocus()
HTEXT txtTrainingDates
End Sub

Private Sub txtTrainingDates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtTrainingVenue.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtTrainingTitle.SetFocus
End If
End Sub

Private Sub txtTrainingTitle_Change()
If ListTrans = is_LstAdding Or _
ListTrans = is_LstEditting Then
    With lstTraining.ListItems
        .Item(ListRow).SubItems(2) = Trim(txtTrainingTitle.Text)
    End With
End If
End Sub

Private Sub txtTrainingTitle_GotFocus()
HTEXT txtTrainingTitle
End Sub

Private Sub txtTrainingTitle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtTrainingDates.SetFocus
End If
End Sub

Private Sub txtTrainingVenue_Change()
If ListTrans = is_LstAdding Or _
ListTrans = is_LstEditting Then
    With lstTraining.ListItems
        .Item(ListRow).SubItems(4) = Trim(txtTrainingVenue.Text)
    End With
End If
End Sub

Private Sub txtTrainingVenue_GotFocus()
HTEXT txtTrainingVenue
End Sub

Private Sub txtTrainingVenue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    picTrainingsSLine.Visible = False
    picMain.Enabled = True
    picToolbar.Enabled = True
    lstTraining.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtTrainingDates.SetFocus
End If
End Sub

Private Sub txtWeight_GotFocus()
HTEXT txtWeight
End Sub

Private Sub txtWeight_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtNationality.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtHeight.SetFocus
End If
End Sub

Private Sub txtWeight_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub XTab1_BeforeTabSwitch(ByVal iNewActiveTab As Integer, bCancel As Boolean)
Select Case iNewActiveTab
    Case 0: txtLastName.SetFocus
    Case 1: txtSpouseName.SetFocus
    Case 2: txtElemName.SetFocus
    Case 3: lstEmployment.SetFocus
    Case Else: Exit Sub
End Select
End Sub

