VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPersonnelInfo 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPersonnelInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   11490
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8880
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
            Picture         =   "frmPersonnelInfo.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInfo.frx":048C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInfo.frx":0610
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInfo.frx":092A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInfo.frx":0CE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInfo.frx":1135
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInfo.frx":1587
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInfo.frx":193F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInfo.frx":1E81
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInfo.frx":1F93
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInfo.frx":20ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInfo.frx":2247
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelInfo.frx":2789
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   120
      ScaleHeight     =   5295
      ScaleWidth      =   11175
      TabIndex        =   3
      Top             =   840
      Width           =   11175
      Begin MSMask.MaskEdBox txtID_1 
         Height          =   315
         Left            =   2880
         TabIndex        =   66
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   11
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "######-####"
         PromptChar      =   " "
      End
      Begin VB.ComboBox cmbDivision 
         Height          =   315
         ItemData        =   "frmPersonnelInfo.frx":29AD
         Left            =   1560
         List            =   "frmPersonnelInfo.frx":29B7
         TabIndex        =   36
         Top             =   1650
         Width           =   3975
      End
      Begin VB.ComboBox cmbTaxStatus 
         Height          =   315
         ItemData        =   "frmPersonnelInfo.frx":29D4
         Left            =   6840
         List            =   "frmPersonnelInfo.frx":29DE
         TabIndex        =   35
         Top             =   0
         Width           =   1575
      End
      Begin VB.TextBox txtNoChild 
         Height          =   315
         Left            =   6840
         TabIndex        =   34
         Top             =   330
         Width           =   1575
      End
      Begin VB.TextBox txtParentsAdd 
         Height          =   315
         Left            =   6840
         TabIndex        =   33
         Top             =   2970
         Width           =   4335
      End
      Begin VB.TextBox txtParents 
         Height          =   315
         Left            =   6840
         TabIndex        =   32
         Top             =   2640
         Width           =   4335
      End
      Begin VB.TextBox txtPostCode 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2640
         Width           =   555
      End
      Begin VB.TextBox txtEmpStatusCode 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   2310
         Width           =   555
      End
      Begin VB.TextBox txtDeptCode 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1980
         Width           =   555
      End
      Begin VB.ComboBox cmbSex 
         Height          =   315
         ItemData        =   "frmPersonnelInfo.frx":29F0
         Left            =   1560
         List            =   "frmPersonnelInfo.frx":29FA
         TabIndex        =   28
         Top             =   3630
         Width           =   1455
      End
      Begin VB.TextBox txtCelNo 
         Height          =   315
         Left            =   6840
         TabIndex        =   27
         Top             =   660
         Width           =   1575
      End
      Begin VB.TextBox txtPost 
         Height          =   315
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   2640
         Width           =   3405
      End
      Begin VB.TextBox txtDept 
         Height          =   315
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1980
         Width           =   3405
      End
      Begin VB.TextBox txtEmpStatus 
         Height          =   315
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2310
         Width           =   3405
      End
      Begin VB.ComboBox cmbStatus 
         Height          =   315
         ItemData        =   "frmPersonnelInfo.frx":2A0C
         Left            =   3960
         List            =   "frmPersonnelInfo.frx":2A1C
         TabIndex        =   23
         Top             =   3630
         Width           =   1575
      End
      Begin VB.TextBox txtBDay 
         Height          =   315
         Left            =   1560
         TabIndex        =   22
         Text            =   "11/27/1794"
         Top             =   2970
         Width           =   1455
      End
      Begin VB.TextBox txtBPlace 
         Height          =   315
         Left            =   1560
         TabIndex        =   21
         Top             =   3300
         Width           =   3975
      End
      Begin VB.TextBox txtLicense 
         Height          =   315
         Left            =   6840
         TabIndex        =   20
         Top             =   990
         Width           =   1575
      End
      Begin VB.TextBox txtContact 
         Height          =   315
         Left            =   6840
         TabIndex        =   19
         Top             =   3300
         Width           =   4335
      End
      Begin VB.TextBox txtSSS 
         Height          =   315
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtPHIC 
         Height          =   315
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1650
         Width           =   1575
      End
      Begin VB.TextBox txtPagIbig 
         Height          =   315
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1980
         Width           =   1575
      End
      Begin VB.TextBox txtTIN 
         Height          =   315
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2310
         Width           =   1575
      End
      Begin VB.TextBox txtProvAdd 
         Height          =   315
         Left            =   1560
         TabIndex        =   14
         Top             =   4950
         Width           =   3975
      End
      Begin VB.TextBox txtAddress 
         Height          =   315
         Left            =   1560
         TabIndex        =   13
         Top             =   4620
         Width           =   3975
      End
      Begin VB.TextBox txtDateHired 
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox txtMName 
         Height          =   315
         Left            =   1560
         TabIndex        =   11
         Top             =   990
         Width           =   3975
      End
      Begin VB.TextBox txtFName 
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   660
         Width           =   3975
      End
      Begin VB.TextBox txtLName 
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   330
         Width           =   3975
      End
      Begin VB.TextBox txtID 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Text            =   "000000-0000"
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   8520
         ScaleHeight     =   2505
         ScaleWidth      =   2625
         TabIndex        =   7
         Top             =   0
         Width           =   2655
         Begin VB.Image imgPicture 
            Height          =   2505
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2625
         End
      End
      Begin VB.TextBox txtDateMarried 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   4290
         Width           =   3975
      End
      Begin VB.TextBox txtSpouseName 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   3960
         Width           =   3975
      End
      Begin VB.TextBox txtAge 
         Height          =   315
         Left            =   3960
         TabIndex        =   4
         Top             =   2970
         Width           =   1575
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   65
         Top             =   1695
         Width           =   1095
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Status"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5640
         TabIndex        =   64
         Top             =   30
         Width           =   975
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Of Dependent"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5640
         TabIndex        =   63
         Top             =   280
         Width           =   1095
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Parents Address"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5640
         TabIndex        =   62
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Parents Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5640
         TabIndex        =   61
         Top             =   2685
         Width           =   1455
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender (Sex)"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   60
         Top             =   3660
         Width           =   1095
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Tel/Cel No"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5640
         TabIndex        =   59
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   58
         Top             =   2685
         Width           =   1095
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   57
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Employment Status"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   56
         Top             =   2370
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Civil Status"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3120
         TabIndex        =   55
         Top             =   3660
         Width           =   855
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   54
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Place"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   53
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Lisence No"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5640
         TabIndex        =   52
         Top             =   1065
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Person"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5640
         TabIndex        =   51
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "SSS No"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5640
         TabIndex        =   50
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Phil Health"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5640
         TabIndex        =   49
         Top             =   1695
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Pag Ibig"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5640
         TabIndex        =   48
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "TIN"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5640
         TabIndex        =   47
         Top             =   2370
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Provincial Address"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   46
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   45
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Hired"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   44
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   43
         Top             =   1065
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   42
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   41
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Number"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   40
         Top             =   30
         Width           =   1095
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Place/Date of Marriage"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   39
         Top             =   4240
         Width           =   1455
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Spouse NAme"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   38
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3120
         TabIndex        =   37
         Top             =   3000
         Width           =   495
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   15000
      TabIndex        =   0
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   540
         Left            =   0
         TabIndex        =   1
         Top             =   100
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   953
         ButtonWidth     =   1191
         ButtonHeight    =   953
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
               Caption         =   "Print"
               Key             =   "Print"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Find"
               Key             =   "Find"
               ImageIndex      =   9
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
         MouseIcon       =   "frmPersonnelInfo.frx":2A45
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   15000
         Y1              =   650
         Y2              =   650
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
         Y1              =   720
         Y2              =   720
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   6285
      Width           =   11490
      _ExtentX        =   20267
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
End
Attribute VB_Name = "frmPersonnelInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ImagePath As String
Public ImageChange As Long
Dim locEmployeePK As Long
Dim locDept As Long
Dim locEmpStatus As Long
Dim locTaxStatus As Long
Dim locPosition As Long

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2
Const is_FINDING = 3

Private Function BROWSER(strID, strAction As String)
Dim Array1
Dim s As String
Dim rs As New ADODB.Recordset
Select Case strAction
    Case "is_LOAD"
        If Trim(strID) <> "" Then
            s = "SELECT TOP 1 tbl_PersonnelProfile.*" & _
                " From tbl_PersonnelProfile " & _
                " WHERE (LName + ',  ' + FName +'  ' + MName + ' - ' + IDNumber = '" & strID & "')" & _
                " ORDER BY LName + ',  ' + FName +'  ' + MName + ' - ' + IDNumber"
        Else
            s = "SELECT TOP 1 tbl_PersonnelProfile.*" & _
                " From tbl_PersonnelProfile " & _
                " ORDER BY LName + ',  ' + FName +'  ' + MName + ' - ' + IDNumber"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_PersonnelProfile.*" & _
            " From tbl_PersonnelProfile " & _
            " ORDER BY LName + ',  ' + FName +'  ' + MName + ' - ' + IDNumber"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_PersonnelProfile.*" & _
            " From tbl_PersonnelProfile " & _
            " WHERE (LName + ',  ' + FName +'  ' + MName + ' - ' + IDNumber < '" & strID & "')" & _
            " ORDER BY LName + ',  ' + FName +'  ' + MName + ' - ' + IDNumber DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_PersonnelProfile.*" & _
            " From tbl_PersonnelProfile " & _
            " WHERE (LName + ',  ' + FName +'  ' + MName + ' - ' + IDNumber > '" & strID & "')" & _
            " ORDER BY LName + ',  ' + FName +'  ' + MName + ' - ' + IDNumber "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_PersonnelProfile.*" & _
            " From tbl_PersonnelProfile " & _
            " ORDER BY LName + ',  ' + FName +'  ' + MName + ' - ' + IDNumber DESC"
    Case Else: Exit Function
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    locDept = rs!Dept
    locEmpStatus = rs!EmpStatus
    locPosition = rs!Positions
    locTaxStatus = rs!TaxStatus
    txtID.Text = rs!IDNumber
    txtLName.Text = rs!LName
    txtFName.Text = rs!FName
    txtMName.Text = rs!MName
    txtDateHired.Text = Format(rs!DHired, "mm/dd/yyyy")
    cmbDivision.ListIndex = rs!Division - 1
    If DEPT_NAME(rs!Dept) <> "" Then
        Array1 = Split(DEPT_NAME(rs!Dept), ";", -1, 1)
        txtDeptCode.Text = CStr(Array1(0))
        txtDept.Text = CStr(Array1(1))
    Else
        txtDeptCode.Text = "000"
        txtDept.Text = "ON PROCESS"
    End If
    If EMP_STATUS(rs!EmpStatus) <> "" Then
        Array1 = Split(EMP_STATUS(rs!EmpStatus), ";", -1, 1)
        txtEmpStatusCode.Text = CStr(Array1(0))
        txtEmpStatus.Text = CStr(Array1(1))
    Else
        txtEmpStatusCode.Text = "000"
        txtEmpStatus.Text = "ON PROCESS"
    End If
    If POSITION_NAME(rs!Positions) <> "" Then
        Array1 = Split(POSITION_NAME(rs!Positions), ";", -1, 1)
        txtPostCode.Text = CStr(Array1(0))
        txtPost.Text = CStr(Array1(1))
    Else
        txtPostCode.Text = "000"
        txtPost.Text = "ON PROCESS"
    End If
    cmbTaxStatus.ListIndex = rs!TaxStatus - 1
    txtBDay.Text = Format(rs!BDate, "mm/dd/yyyy")
    txtBPlace.Text = rs!BPlace
    txtContact.Text = rs!ContactPerson
    txtAddress.Text = rs!Address
    txtProvAdd.Text = rs!ProvAdd
    txtCelNo.Text = rs!CelNo
    txtLicense.Text = rs!License
    txtSSS.Text = rs!SSS
    txtPHIC.Text = rs!PHIC
    txtPagIbig.Text = rs!PagIbig
    txtTin.Text = rs!TIN
    cmbSex.ListIndex = rs!Sex - 1
    cmbStatus.ListIndex = rs!Status - 1
    txtParents.Text = rs!Parents
    txtParentsAdd.Text = rs!Parents_Add
    txtSpouseName.Text = rs!SpouseName
    txtDateMarried.Text = IIf(DateValue(CDate(rs!DateMarried)) = DateValue(CDate("01/01/1900")), "", Format(rs!DateMarried, "mm/dd/yyyy"))
    txtAge.Text = rs!Age
    txtNoChild.Text = rs!NoChildren
    StatusBar.Panels(1).Text = rs!PK
    StatusBar.Panels(2).Text = "LAST MODIFIED BY : " & rs!LastModified
    
    SHOW_EMPLOYEE_PHOTO rs!PK, imgPicture, ImagePath
    
    SaveSetting App.EXEName, "PersonnelInfo", "PerInfo", rs!LName & ",  " & rs!FName & "  " & rs!MName & " - " & rs!IDNumber
End If
rs.Close
End Function
Private Function PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If AccessRights("Personnel Information", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
TRANSACTIONTYPE = is_ADDING
TOOLBARBUTTON False
CLEARTEXT
LOCKTEXT False
txtDept.Text = "ON PROCESS"
txtDeptCode.Text = "000"
txtEmpStatus.Text = "ON PROCESS"
txtEmpStatusCode.Text = "000"
txtPost.Text = "ON PROCESS"
txtPostCode.Text = "000"
txtSSS.Text = "ON PROCESS"
txtPHIC.Text = "ON PROCESS"
txtPagIbig.Text = "ON PROCESS"
txtTin.Text = "ON PROCESS"
txtID_1.Visible = True
txtID_1.Move txtID.Left, txtID.Top, txtID.Width
txtID_1.Text = "000000-0000"
txtID_1.SetFocus
HTEXT txtID_1
Me.Caption = "Personnel Information - New"
End Function

Private Function PRESS_F2()
Dim strID
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If StatusBar.Panels(1).Text = "" Then Exit Function
'gbl_MODULE = "PERSONNEL INFORMATION"
'gbl_MODULE_Action = "EDIT"
'USER_ACCESS_RIGHTS
'If ACCESS_RIGHTS = False Then Exit Function
If AccessRights("Personnel Information", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
TRANSACTIONTYPE = is_EDITTING
TOOLBARBUTTON False
LOCKTEXT False
txtID_1.Move txtID.Left, txtID.Top, txtID.Width
txtID_1.Visible = True
strID = Split(Trim(txtID.Text), "-", -1, 1)
txtID_1.Text = Format(strID(0), "mmddyy") & "-" & Format(strID(1), "000#")
txtID_1.SetFocus
HTEXT txtID_1
Me.Caption = "Personnel Information - Edit"
End Function

Private Function PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If StatusBar.Panels(1).Text = "" Then Exit Function
If AccessRights("Personnel Information", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
If MsgBox("ARE YOU SURE TO DELETE THIS RECORD?          ", vbInformation + vbYesNo, "CONFIRMATION") = vbNo Then Exit Function
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_PersonnelProfile WHERE (PK = " & StatusBar.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "PersonnelInfo", "PerInfo", ""), "is_PAGEDOWN"
If StatusBar.Panels(1).Text = "" Then BROWSER GetSetting(App.EXEName, "PersonnelInfo", "PerInfo", ""), "is_HOME"

Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Function
End Function

Public Function SAVEPROFILE(strID, strLName, strFName, strMName, _
dtmHired, intDept, intEmpStatus, intPost, dtmBDay, strBPlace, _
strContact, strAdd, strProvAdd, intSex, intStatus, strCelNo, _
strLicense, strSSS, strPHIC, strPagIbig, strTIN, strParents, _
strParents_Add, strSpouseName, dtmDateMarried, intNoChildren, _
intAge, intDivision, intTaxStatus, strLastModified)
Dim s As String
s = "INSERT INTO tbl_PersonnelProfile" & _
    " (IDNumber, LName, FName, MName, DHired, Dept, EmpStatus, " & _
    " Positions, BDate, BPlace, ContactPerson, Address, ProvAdd, " & _
    " Sex, Status, CelNo, License, SSS, PHIC, PagIbig, TIN, " & _
    " Parents, Parents_Add, SpouseName, DateMarried, NoChildren, " & _
    " Age, Division, TaxStatus, LastModified) " & _
    " VALUES ('" & strID & "', '" & strLName & "', '" & strFName & "', " & _
    " '" & strMName & "', '" & CDate(dtmHired) & "', " & intDept & ", " & _
    " " & intEmpStatus & ", " & intPost & ", '" & CDate(dtmBDay) & "', " & _
    " '" & strBPlace & "', '" & strContact & "', '" & strAdd & "', " & _
    " '" & strProvAdd & "', " & intSex & ", " & intStatus & ", '" & strCelNo & "'," & _
    " '" & strLicense & "', '" & strSSS & "', '" & strPHIC & "', '" & strPagIbig & "', " & _
    " '" & strTIN & "', '" & strParents & "', '" & strParents_Add & "', " & _
    " '" & strSpouseName & "', '" & CDate(dtmDateMarried) & "', " & intNoChildren & ", " & _
    " " & intAge & ", " & intDivision & "," & intTaxStatus & ",'" & strLastModified & "')"
ConnOmega.Execute s, , -1
End Function

Public Function UPDATEPROFILE(intPK, strID, strLName, strFName, strMName, _
dtmHired, intDept, intEmpStatus, intPost, dtmBDay, strBPlace, _
strContact, strAdd, strProvAdd, intSex, intStatus, strCelNo, _
strLicense, strSSS, strPHIC, strPagIbig, strTIN, strParents, _
strParents_Add, strSpouseName, dtmDateMarried, intNoChildren, _
intAge, intDivision, intTaxStatus, strLastModified)
Dim s As String
s = "UPDATE tbl_PersonnelProfile" & _
    " SET IDNumber ='" & strID & "', LName ='" & strLName & "', " & _
    " FName ='" & strFName & "', MName ='" & strMName & "', " & _
    " DHired ='" & CDate(dtmHired) & "', Dept =" & intDept & ", " & _
    " EmpStatus =" & intEmpStatus & ", Positions =" & intPost & ", " & _
    " BDate ='" & CDate(dtmBDay) & "', BPlace ='" & strBPlace & "', " & _
    " ContactPerson ='" & strContact & "', Address ='" & strAdd & "', " & _
    " ProvAdd ='" & strProvAdd & "', Sex =" & intSex & ", " & _
    " Status =" & intStatus & ", CelNo ='" & strCelNo & "', " & _
    " License ='" & strLicense & "', SSS ='" & strSSS & "', " & _
    " PHIC ='" & strPHIC & "', PagIbig ='" & strPagIbig & "', " & _
    " TIN ='" & strTIN & "', LastModified ='" & strLastModified & "', " & _
    " Parents = '" & strParents & "',Parents_Add = '" & strParents_Add & "', " & _
    " SpouseName = '" & strSpouseName & "', DateMarried = '" & CDate(dtmDateMarried) & "', " & _
    " NoChildren = " & intNoChildren & ", Age = " & intAge & ", " & _
    " Division = " & intDivision & ", TaxStatus = " & intTaxStatus & " " & _
    " WHERE (PK = " & intPK & ")"
ConnOmega.Execute s, , -1
End Function

Private Function PRESS_F5()
Dim myStream As New ADODB.Stream
Dim s As String
Dim rs As New ADODB.Recordset
If TRANSACTIONTYPE = is_ADDING Then
    On Error GoTo PG:
    SAVEPROFILE Trim(txtID.Text), Trim(txtLName.Text), _
        Trim(txtFName.Text), Trim(txtMName.Text), _
        Trim(txtDateHired.Text), locDept, locEmpStatus, _
        locPosition, Trim(txtBDay.Text), Trim(txtBPlace.Text), _
        Trim(txtContact.Text), Trim(txtAddress.Text), _
        Trim(txtProvAdd.Text), cmbSex.ListIndex + 1, _
        cmbStatus.ListIndex + 1, Trim(txtCelNo.Text), _
        Trim(txtLicense.Text), Trim(txtSSS.Text), _
        Trim(txtPHIC.Text), Trim(txtPagIbig.Text), _
        Trim(txtTin.Text), Trim(txtParents.Text), _
        Trim(txtParentsAdd.Text), Trim(txtSpouseName.Text), _
        IIf(Trim(txtDateMarried.Text) = "", "01/01/1900", Trim(txtDateMarried.Text)), _
        IIf(Trim(txtNoChild.Text) = "", 0, Trim(txtNoChild.Text)), _
        IIf(Trim(txtAge.Text) = "", 0, Trim(txtAge.Text)), _
        cmbDivision.ListIndex + 1, locTaxStatus, CStr(Now) & " - " & gbl_UserName
    If ImageChange = 1 Then
        
        
        s = "SELECT TOP 1 tbl_PersonnelProfile.*" & _
            " From tbl_PersonnelProfile " & _
            " WHERE (LName + ',  ' + FName +'  ' + MName + ' - ' + IDNumber = '" & FORMATSQL(Trim(txtLName.Text) & ",  " & Trim(txtFName.Text) & "  " & Trim(txtMName.Text) & " - " & Trim(txtID.Text)) & "')" & _
            " ORDER BY LName + ',  ' + FName +'  ' + MName + ' - ' + IDNumber"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            SAVE_EMPLOYEE_PHOTO rs!PK, ImagePath
        End If
        rs.Close
        ImageChange = 0
        
    End If
    BROWSER Trim(txtLName.Text) & ",  " & Trim(txtFName.Text) & "  " & Trim(txtMName.Text) & " - " & Trim(txtID.Text), "is_LOAD"
    LOCKTEXT True
    TOOLBARBUTTON True
    TRANSACTIONTYPE = is_REFRESH
    Me.Caption = "Personnel Information - Browse"
ElseIf TRANSACTIONTYPE = is_EDITTING Then
'    On Error GoTo PG:
    UPDATEPROFILE StatusBar.Panels(1).Text, Trim(txtID.Text), Trim(txtLName.Text), _
        Trim(txtFName.Text), Trim(txtMName.Text), _
        Trim(txtDateHired.Text), locDept, locEmpStatus, _
        locPosition, Trim(txtBDay.Text), Trim(txtBPlace.Text), _
        Trim(txtContact.Text), Trim(txtAddress.Text), _
        Trim(txtProvAdd.Text), cmbSex.ListIndex + 1, _
        cmbStatus.ListIndex + 1, Trim(txtCelNo.Text), _
        Trim(txtLicense.Text), Trim(txtSSS.Text), _
        Trim(txtPHIC.Text), Trim(txtPagIbig.Text), _
        Trim(txtTin.Text), Trim(txtParents.Text), _
        Trim(txtParentsAdd.Text), Trim(txtSpouseName.Text), _
        IIf(Trim(txtDateMarried.Text) = "", "01/01/1900", Trim(txtDateMarried.Text)), _
        IIf(Trim(txtNoChild.Text) = "", 0, Trim(txtNoChild.Text)), _
        IIf(Trim(txtAge.Text) = "", 0, Trim(txtAge.Text)), _
        cmbDivision.ListIndex + 1, locTaxStatus, CStr(Now) & " - " & gbl_UserName
    If ImageChange = 1 Then
        SAVE_EMPLOYEE_PHOTO StatusBar.Panels(1).Text, ImagePath
        ImageChange = 0
    End If
    BROWSER Trim(txtLName.Text) & ",  " & Trim(txtFName.Text) & "  " & Trim(txtMName.Text) & " - " & Trim(txtID.Text), "is_LOAD"
    LOCKTEXT True
    TOOLBARBUTTON True
    TRANSACTIONTYPE = is_REFRESH
    Me.Caption = "Personnel Information - Browse"
End If
Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error Saving"
Exit Function
End Function

Private Function PRESS_F6()
If TRANSACTIONTYPE = is_REFRESH Then
'    PopupMenu frmPopUpMenu.mnuFindEmployee, , 5000, 500
End If
End Function

Private Function PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    Unload Me
Else
    TRANSACTIONTYPE = is_REFRESH
    LOCKTEXT True
    TOOLBARBUTTON True
    txtID_1.Visible = False
    Me.Caption = "Personnel Information - Browse"
    BROWSER GetSetting(App.EXEName, "PersonnelInfo", "PerInfo", ""), "is_LOAD"
End If
End Function

Private Function CLEARTEXT()
ImagePath = ""
locDept = 0
locEmpStatus = 0
locPosition = 0
locTaxStatus = 0
txtAge.Text = ""
txtID.Text = ""
txtLName.Text = ""
txtFName.Text = ""
txtMName.Text = ""
txtDateHired.Text = ""
cmbDivision.Text = ""
cmbDivision.ListIndex = -1
cmbTaxStatus.Text = ""
cmbTaxStatus.ListIndex = -1
txtDept.Text = ""
txtDeptCode.Text = ""
txtEmpStatus.Text = ""
txtEmpStatusCode.Text = ""
txtPost.Text = ""
txtPostCode.Text = ""
txtBDay.Text = ""
txtBPlace.Text = ""
txtContact.Text = ""
txtAddress.Text = ""
txtProvAdd.Text = ""
cmbSex.Text = ""
cmbStatus.Text = ""
txtCelNo.Text = ""
txtLicense.Text = ""
txtSSS.Text = ""
txtPHIC.Text = ""
txtPagIbig.Text = ""
txtTin.Text = ""
cmbSex.ListIndex = -1
cmbStatus.ListIndex = -1
txtParents.Text = ""
txtParentsAdd.Text = ""
txtSpouseName.Text = ""
txtDateMarried.Text = ""
txtNoChild.Text = ""
StatusBar.Panels(1).Text = ""
StatusBar.Panels(2).Text = ""
imgPicture.Picture = LoadPicture("")
'imgPicture.Visible = False
'imgLogo.Visible = True
End Function

Private Function LOCKTEXT(bln As Boolean)
If bln Then
    txtID.Locked = True
    txtLName.Locked = True
    txtFName.Locked = True
    txtMName.Locked = True
    txtDateHired.Locked = True
    cmbDivision.Locked = True
    txtDept.Locked = True
    txtEmpStatus.Locked = True
    cmbTaxStatus.Locked = True
    txtPost.Locked = True
    txtBDay.Locked = True
    txtBPlace.Locked = True
    txtContact.Locked = True
    txtAddress.Locked = True
    txtProvAdd.Locked = True
    cmbSex.Locked = True
    cmbStatus.Locked = True
    txtCelNo.Locked = True
    txtLicense.Locked = True
    txtSSS.Locked = True
    txtPHIC.Locked = True
    txtPagIbig.Locked = True
    txtTin.Locked = True
    txtParents.Locked = True
    txtParentsAdd.Locked = True
    txtSpouseName.Locked = True
    txtDateMarried.Locked = True
    txtNoChild.Locked = True
    txtAge.Locked = True
Else
    txtID.Locked = False
    txtLName.Locked = False
    txtFName.Locked = False
    txtMName.Locked = False
    txtDateHired.Locked = False
    cmbDivision.Locked = False
    txtDept.Locked = True
    txtEmpStatus.Locked = True
    cmbTaxStatus.Locked = False
    txtPost.Locked = True
    txtBDay.Locked = False
    txtBPlace.Locked = False
    txtContact.Locked = False
    txtAddress.Locked = False
    txtProvAdd.Locked = False
    cmbSex.Locked = False
    cmbStatus.Locked = False
    txtCelNo.Locked = False
    txtLicense.Locked = False
    txtSSS.Locked = True
    txtPHIC.Locked = True
    txtPagIbig.Locked = True
    txtTin.Locked = True
    txtParents.Locked = False
    txtParentsAdd.Locked = False
    txtSpouseName.Locked = False
    txtDateMarried.Locked = False
    txtNoChild.Locked = False
    txtAge.Locked = False
End If
End Function

Private Sub TOOLBARBUTTON(blnTag As Boolean)
Set Toolbar1.ImageList = ImageList1
With Toolbar1.Buttons
    If blnTag Then
        .Item(1).Image = 1
        .Item(3).Image = 2
        .Item(5).Image = 3
        .Item(11).Image = 6
        .Item(13).Image = 7
        .Item(15).Image = 8
        .Item(17).Image = 9
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
    Else
        .Item(1).Image = 1
        .Item(3).Image = 2
        .Item(5).Image = 3
        .Item(11).Image = 6
        .Item(13).Image = 7
        .Item(15).Image = 8
        .Item(17).Image = 9
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
    End If
End With
End Sub


Private Sub cmbDivision_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtDept.SetFocus
End If
End Sub

Private Sub cmbSex_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    cmbStatus.SetFocus
'ElseIf KeyCode = vbKeyDown Then
'    If cmbSex.Locked = True Then
'        txtSpouseName.SetFocus
'    End If
'ElseIf KeyCode = vbKeyUp Then
'    txtAge.SetFocus
End If
End Sub

Private Sub cmbStatus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtSpouseName.SetFocus
'ElseIf KeyCode = vbKeyDown Then
'    txtBPlace.SetFocus
'ElseIf KeyCode = vbKeyUp Then
'    cmbSex.SetFocus
End If
End Sub

Private Sub cmbTaxStatus_Click()
If cmbTaxStatus.ListIndex > -1 Then
    locTaxStatus = cmbTaxStatus.ItemData(cmbTaxStatus.ListIndex)
End If
End Sub

Private Sub cmbTaxStatus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtNoChild.SetFocus
End If
End Sub

Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
Dim a As String
Dim s As String
Dim rs As New ADODB.Recordset
s = "SELECT PK, IDNumber, LName, FName, MName, " & _
    " DHired, Division, Dept, EmpStatus, TaxStatus, " & _
    " Positions, BDate, BPlace, ContactPerson, " & _
    " Address, ProvAdd, CelNo, License, SSS, PHIC, " & _
    " PagIbig, TIN, Sex, Status, Parents, Parents_Add, " & _
    " SpouseName, DateMarried, NoChildren, Age, " & _
    " CurrentAction, LastModified " & _
    " FROM PersonnelProfile"
rs.Open s, ConnOmega
While Not rs.EOF
    DoEvents
    a = "INSERT INTO tbl_PersonnelProfile " & _
        " (PK, IDNumber, LName, FName, MName, " & _
        " DHired, Division, Dept, EmpStatus, TaxStatus, " & _
        " Positions, BDate, BPlace, ContactPerson, " & _
        " Address, ProvAdd, CelNo, License, SSS, PHIC, " & _
        " PagIbig, TIN, Sex, Status, Parents, Parents_Add, " & _
        " SpouseName, DateMarried, NoChildren, Age, " & _
        " CurrentAction, LastModified) " & _
        " VALUES (" & rs!PK & ", '" & rs!IDNumber & "', '" & rs!LName & "', " & _
        " '" & rs!FName & "', '" & rs!MName & "', '" & rs!DHired & "', " & rs!Division & ", " & _
        " " & rs!Dept & ", " & rs!EmpStatus & ", " & rs!TaxStatus & ", " & _
        " " & rs!Positions & ", '" & rs!BDate & "', '" & Replace(rs!BPlace, "'", "''") & "', " & _
        " '" & rs!ContactPerson & "', '" & Replace(rs!Address, "'", "''") & "', " & _
        " '" & Replace(rs!ProvAdd, "'", "''") & "', '" & rs!CelNo & "', '" & rs!License & "', " & _
        " '" & rs!SSS & "', '" & rs!PHIC & "', '" & rs!PagIbig & "', '" & rs!TIN & "', " & _
        " " & rs!Sex & ", " & rs!Status & ", '" & rs!Parents & "', " & _
        " '" & Replace(rs!Parents_Add, "'", "''") & "', '" & rs!SpouseName & "', " & _
        " '" & rs!DateMarried & "', " & rs!NoChildren & ", " & rs!Age & ", " & _
        " " & rs!CurrentAction & ", '" & rs!LastModified & "')"
    ConnOmega.Execute a, , -1
    rs.MoveNext
Wend
rs.Close
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert:   PRESS_INSERT
    Case vbKeyF2:       PRESS_F2
    Case vbKeyDelete:   PRESS_DELETE
    Case vbKeyF5:       PRESS_F5
    Case vbKeyF6:       PRESS_F6
    Case vbKeyF9:
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "PersonnelInfo", "PerInfo", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "PersonnelInfo", "PerInfo", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "PersonnelInfo", "PerInfo", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "PersonnelInfo", "PerInfo", ""), "is_END"
    Case vbKeyEscape:   PRESS_ESCAPE
End Select
End Sub

Private Sub Form_Load()
DoEvents
KeyPreview = True
Me.Top = (Mainform.Height - Me.Height) / 20
Me.Left = (Mainform.Width - Me.Width) / 3
POPULATE_COMBO "PK", "TaxStatus", "tbl_Personnel_TaxStatus", "PK", cmbTaxStatus
With cmbSex
    .Clear
    .AddItem "MALE"
    .AddItem "FEMALE"
End With
With cmbStatus
    .Clear
    .AddItem "SINGLE"
    .AddItem "MARRIED"
    .AddItem "WIDOWED"
    .AddItem "WIDOWER"
End With
TOOLBARBUTTON True
LOCKTEXT True
ImageChange = 0

BROWSER GetSetting(App.EXEName, "PersonnelInfo", "PerInfo", ""), "is_LOAD"
If Trim(txtLName.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelInfo", "PerInfo", ""), "is_HOME"
Me.Caption = "Personnel Information - Browse"
'SETFIELDSLOAD Mainform.txtPersonnelIndex.Text
Dim tmp As Long
tmp = SetWindowLong(txtID.hWnd, GWL_STYLE, GetWindowLong(txtID.hWnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtLName.hWnd, GWL_STYLE, GetWindowLong(txtLName.hWnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtFName.hWnd, GWL_STYLE, GetWindowLong(txtFName.hWnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtMName.hWnd, GWL_STYLE, GetWindowLong(txtMName.hWnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtDateHired.hWnd, GWL_STYLE, GetWindowLong(txtDateHired.hWnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtBDay.hWnd, GWL_STYLE, GetWindowLong(txtBDay.hWnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtBPlace.hWnd, GWL_STYLE, GetWindowLong(txtBPlace.hWnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtContact.hWnd, GWL_STYLE, GetWindowLong(txtContact.hWnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtAddress.hWnd, GWL_STYLE, GetWindowLong(txtAddress.hWnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtProvAdd.hWnd, GWL_STYLE, GetWindowLong(txtProvAdd.hWnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtCelNo.hWnd, GWL_STYLE, GetWindowLong(txtCelNo.hWnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtLicense.hWnd, GWL_STYLE, GetWindowLong(txtLicense.hWnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtParents.hWnd, GWL_STYLE, GetWindowLong(txtParents.hWnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtParentsAdd.hWnd, GWL_STYLE, GetWindowLong(txtParentsAdd.hWnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSpouseName.hWnd, GWL_STYLE, GetWindowLong(txtSpouseName.hWnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSpouseName.hWnd, GWL_STYLE, GetWindowLong(txtSpouseName.hWnd, GWL_STYLE) Or ES_UPPERCASE)
'On Error Resume Next
'Me.Picture = LoadPicture(App.Path & "\images\new-6.jpg")
'picToolbar.Picture = LoadPicture(App.Path & "\images\new-6.jpg")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If TRANSACTIONTYPE <> is_REFRESH Then
    Cancel = -1
End If
End Sub



Private Sub imgPicture_DblClick()
Dim FileName
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    Mainform.CommonDialog1.CancelError = True
    On Error GoTo ErrorHandler
        Mainform.CommonDialog1.Filter = "Image Files (*.jpg)|*.jpg"
        Mainform.CommonDialog1.ShowOpen
        FileName = Trim(Mainform.CommonDialog1.FileName)
        If ((FileLen(FileName) \ 1024) + 1) <= 66 Then
            ImagePath = FileName
            imgPicture.Picture = LoadPicture(FileName)
            imgPicture.Visible = True
            ImageChange = 1
        Else
            MsgBox "Image is too large please reduce the size to 65kb or below!          ", vbCritical, "Error..."
            Exit Sub
        End If
End If
Exit Sub
ErrorHandler:
Exit Sub
End Sub

Private Sub imgPicture_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    imgPicture.ToolTipText = "Double Click Here to Insert Picture!"
Else
    imgPicture.ToolTipText = ""
End If
End Sub

Private Sub Picture3_DblClick()
Dim FileName
'Dim DataFile            As Integer
'Dim Fl                  As Long

If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    Mainform.CommonDialog1.CancelError = True
    On Error GoTo ErrorHandler
        Mainform.CommonDialog1.Filter = "Image Files (*.jpg)|*.jpg"
        Mainform.CommonDialog1.ShowOpen
        FileName = Trim(Mainform.CommonDialog1.FileName)
        If ((FileLen(FileName) \ 1024) + 1) <= 15 Then
'        Open FileName For Binary Access Read As DataFile
'        Fl = LOF(DataFile)
'        If Fl = 0 Then Exit Sub
'        If Fl <= 15000 Then
            ImagePath = FileName
            imgPicture.Picture = LoadPicture(FileName)
'            imgLogo.Visible = False
            imgPicture.Visible = True
            ImageChange = 1
        Else
            MsgBox "Image is too large please reduce the size to 15k or below!          ", vbCritical, "Error..."
            Exit Sub
        End If
End If
Exit Sub
ErrorHandler:
Exit Sub
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":           PRESS_INSERT
    Case "Edit":          PRESS_F2
    Case "Delete":        PRESS_DELETE
    Case "First"
        Select Case Toolbar1.Buttons(7).Caption
            Case "Save":  PRESS_F5
            Case "First": BROWSER GetSetting(App.EXEName, "PersonnelInfo", "PerInfo", ""), "is_HOME"
        End Select
    Case "Back"
        Select Case Toolbar1.Buttons(9).Caption
            Case "Undo":  PRESS_ESCAPE
            Case "Back":  BROWSER GetSetting(App.EXEName, "PersonnelInfo", "PerInfo", ""), "is_PAGEUP"
        End Select
    Case "Next":          BROWSER GetSetting(App.EXEName, "PersonnelInfo", "PerInfo", ""), "is_PAGEDOWN"
    Case "Last":          BROWSER GetSetting(App.EXEName, "PersonnelInfo", "PerInfo", ""), "is_END"
    Case "Find":          PRESS_F6
    Case "Print":
    Case "Close":         PRESS_ESCAPE
End Select
End Sub

Private Sub txtAddress_GotFocus()
HTEXT txtAddress
End Sub

Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtProvAdd.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtDateMarried.SetFocus
End If
End Sub

Private Sub txtAge_GotFocus()
HTEXT txtAge
End Sub

Private Sub txtAge_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtBPlace.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtBDay.SetFocus
End If
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtBDay_Change()
Dim intStart, intEnd
If Trim(txtBDay.Text) <> "" Then
    If IsDate(Trim(txtBDay.Text)) Then
        intStart = Year(FormatDateTime(Trim(txtBDay.Text), vbShortDate))
        intEnd = Year(Date)
        txtAge.Text = intEnd - intStart  'DateDiff("y", Date, FormatDateTime(Trim(txtBDay.Text), vbShortDate), vbSunday, vbFirstJan1)
    End If
End If
End Sub

Private Sub txtBDay_GotFocus()
HTEXT txtBDay
End Sub

Private Sub txtBDay_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtAge.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtPost.SetFocus
End If
End Sub


Private Sub txtBDay_LostFocus()
If IsDate(Trim(txtBDay.Text)) Then
    txtBDay.Text = Format(FormatDateTime(Trim(txtBDay.Text), vbShortDate), "mm/dd/yyyy")
Else
    MsgBox "PLEASE SUPPLY A VALID DATE!     ", vbCritical, "Error"
    txtBDay.SetFocus
    HTEXT txtBDay
End If
End Sub

Private Sub txtBPlace_GotFocus()
HTEXT txtBPlace
End Sub

Private Sub txtBPlace_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    cmbSex.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtAge.SetFocus
End If
End Sub


Private Sub txtCelNo_GotFocus()
HTEXT txtCelNo
End Sub

Private Sub txtCelNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtLicense.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtNoChild.SetFocus
End If
End Sub

Private Sub txtContact_GotFocus()
HTEXT txtContact
End Sub

Private Sub txtContact_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If TRANSACTIONTYPE = is_REFRESH Then
        txtID.SetFocus
    Else
        PRESS_F5
    End If
ElseIf KeyCode = vbKeyDown Then
    cmbStatus.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtParentsAdd.SetFocus
End If
End Sub

Private Sub txtDateHired_GotFocus()
HTEXT txtDateHired
End Sub

Private Sub txtDateHired_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    cmbDivision.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtMName.SetFocus
End If
End Sub

Private Sub txtDateHired_LostFocus()
If IsDate(Trim(txtDateHired.Text)) Then
    txtDateHired.Text = Format(FormatDateTime(Trim(txtDateHired.Text), vbShortDate), "mm/dd/yyyy")
Else
    MsgBox "PLEASE SUPPLY A VALID DATE!     ", vbCritical, "Error"
    txtDateHired.SetFocus
    HTEXT txtDateHired
End If
End Sub

Private Sub txtDateMarried_GotFocus()
HTEXT txtDateMarried
End Sub

Private Sub txtDateMarried_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtAddress.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtSpouseName.SetFocus
End If
End Sub

Private Sub txtDateMarried_LostFocus()
If Trim(txtDateMarried.Text) <> "" Then
    If IsDate(Trim(txtDateMarried.Text)) Then
        txtDateMarried.Text = Format(FormatDateTime(Trim(txtDateMarried.Text), vbShortDate), "mm/dd/yyyy")
    Else
        MsgBox "PLEASE SUPPLY A VALID DATE!     ", vbCritical, "Error"
        txtDateMarried.SetFocus
        HTEXT txtDateMarried
    End If
End If
End Sub

Private Sub txtDept_GotFocus()
HTEXT txtDept
End Sub

Private Sub txtDept_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtEmpStatus.SetFocus
ElseIf KeyCode = vbKeyUp Then
    cmbDivision.SetFocus
End If
End Sub

Private Sub txtDeptCode_GotFocus()
HTEXT txtDeptCode
End Sub

Private Sub txtEmpStatus_GotFocus()
HTEXT txtEmpStatus
End Sub

Private Sub txtEmpStatus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtPost.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtDept.SetFocus
End If
End Sub

Private Sub txtEmpStatusCode_GotFocus()
HTEXT txtEmpStatusCode
End Sub

Private Sub txtFName_GotFocus()
HTEXT txtFName
End Sub

Private Sub txtFName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtMName.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtLName.SetFocus
End If
End Sub

Private Sub txtID_1_GotFocus()
HTEXT txtID_1
End Sub

Private Sub txtID_1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtLName.SetFocus
End If
End Sub

Private Sub txtID_1_LostFocus()
Dim Array1, strDate, strSeries
If TRANSACTIONTYPE <> is_REFRESH Then
    Array1 = Split(txtID_1.Text, "-", -1, 1)
    If Trim(Array1(0)) <> "" Then
        strDate = Format(Array1(0), "00000#")
    Else
        strDate = "000000"
    End If
    If Trim(Array1(1)) <> "" Then
        strSeries = Format(Array1(1), "000#")
    Else
        strSeries = "0000"
    End If
    txtID.Text = strDate & "-" & strSeries 'txtID_1.Text
    txtID_1.Visible = False
End If
End Sub

Private Sub txtID_GotFocus()
HTEXT txtID
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    txtID_1.Visible = True
    txtID_1.Move txtID.Left, txtID.Top, txtID.Width
    If Trim(txtID.Text) = "" Then
        txtID_1.Text = "000000-0000"
    Else
        txtID_1.Text = txtID.Text
    End If
    txtID_1.SetFocus
End If
End Sub

Private Sub txtID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtLName.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtContact.SetFocus
End If
End Sub

Private Sub txtLicense_GotFocus()
HTEXT txtLicense
End Sub

Private Sub txtLicense_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtSSS.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtCelNo.SetFocus
End If
End Sub

Private Sub txtLName_GotFocus()
HTEXT txtLName
End Sub

Private Sub txtLName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtFName.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtID.SetFocus
End If
End Sub

Private Sub txtMName_GotFocus()
HTEXT txtMName
End Sub

Private Sub txtMName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtDateHired.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtFName.SetFocus
End If
End Sub

Private Sub txtNoChild_GotFocus()
HTEXT txtNoChild
End Sub

Private Sub txtNoChild_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtCelNo.SetFocus
ElseIf KeyCode = vbKeyUp Then
    cmbTaxStatus.SetFocus
End If
End Sub

Private Sub txtNoChild_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtPagIbig_GotFocus()
HTEXT txtPagIbig
End Sub

Private Sub txtPagIbig_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtTin.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtPHIC.SetFocus
End If
End Sub

Private Sub txtParents_GotFocus()
HTEXT txtParents
End Sub

Private Sub txtParents_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtParentsAdd.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtTin.SetFocus
End If
End Sub

Private Sub txtParentsAdd_GotFocus()
HTEXT txtParentsAdd
End Sub

Private Sub txtParentsAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtContact.SetFocus
ElseIf KeyCode = vbKeyDown Then
    txtContact.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtParents.SetFocus
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

Private Sub txtPost_GotFocus()
HTEXT txtPost
End Sub

Private Sub txtPost_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtBDay.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtEmpStatus.SetFocus
End If
End Sub

Private Sub txtPostCode_GotFocus()
HTEXT txtPostCode
End Sub

Private Sub txtProvAdd_GotFocus()
HTEXT txtProvAdd
End Sub

Private Sub txtProvAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    cmbTaxStatus.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtAddress.SetFocus
End If
End Sub

Private Sub txtSpouseName_GotFocus()
HTEXT txtSpouseName
End Sub

Private Sub txtSpouseName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtDateMarried.SetFocus
ElseIf KeyCode = vbKeyDown Then
    cmbStatus.SetFocus
End If
End Sub

Private Sub txtSSS_GotFocus()
HTEXT txtSSS
End Sub

Private Sub txtSSS_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or _
KeyCode = vbKeyDown Then
    txtPHIC.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtLicense.SetFocus
End If
End Sub

Private Sub txtTIN_GotFocus()
HTEXT txtTin
End Sub

Private Sub txtTIN_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtParents.SetFocus
ElseIf KeyCode = vbKeyDown Then
    txtParents.SetFocus
ElseIf KeyCode = vbKeyUp Then
    txtPagIbig.SetFocus
End If
End Sub




