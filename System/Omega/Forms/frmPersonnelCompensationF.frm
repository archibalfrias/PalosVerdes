VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelCompensationF 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10770
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPersonnelCompensationF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   10770
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9720
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483648
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationF.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationF.frx":09CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationF.frx":0B50
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationF.frx":0E6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationF.frx":1223
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationF.frx":1675
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationF.frx":1AC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationF.frx":1E7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationF.frx":1F91
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationF.frx":24D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationF.frx":262D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelCompensationF.frx":2B6F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   120
      ScaleHeight     =   6015
      ScaleWidth      =   10455
      TabIndex        =   3
      Top             =   840
      Width           =   10455
      Begin VB.PictureBox picEarning 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F1DA&
         ForeColor       =   &H80000008&
         Height          =   4455
         Left            =   0
         ScaleHeight     =   4425
         ScaleWidth      =   4065
         TabIndex        =   16
         Top             =   1320
         Width           =   4095
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   8040
         ScaleHeight     =   1905
         ScaleWidth      =   2265
         TabIndex        =   15
         Top             =   0
         Width           =   2295
         Begin VB.Image imgPicture 
            Height          =   1905
            Left            =   120
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2265
         End
      End
      Begin VB.PictureBox picEmployee 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F1DA&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   0
         ScaleHeight     =   1185
         ScaleWidth      =   7905
         TabIndex        =   4
         Top             =   0
         Width           =   7935
         Begin VB.TextBox txtPayrollPeriod 
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   435
            Width           =   2895
         End
         Begin VB.TextBox txtDivName 
            Height          =   315
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   435
            Width           =   2655
         End
         Begin VB.TextBox txtPost 
            Height          =   315
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   750
            Width           =   2655
         End
         Begin VB.TextBox txtDept 
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   750
            Width           =   2895
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   120
            Width           =   6615
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "DIVISION"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4200
            TabIndex        =   14
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "PERIOD"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "POSITION"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4200
            TabIndex        =   12
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "DEPARTMENT"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "EMPLOYEE"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   4455
         Left            =   120
         Top             =   1440
         Width           =   4095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   1935
         Left            =   8115
         Top             =   90
         Width           =   2295
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   1215
         Left            =   80
         Top             =   75
         Width           =   7935
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
         Height          =   570
         Left            =   0
         TabIndex        =   1
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
            NumButtons      =   18
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
               Caption         =   "Close"
               Key             =   "Close"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   15000
         Y1              =   720
         Y2              =   720
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
         Y1              =   650
         Y2              =   650
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   6990
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11994
            MinWidth        =   11994
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPersonnelCompensationF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim s                   As String
Dim rs                  As New ADODB.Recordset
Dim t                   As String
Dim rt                  As New ADODB.Recordset
Dim u                   As String
Dim ru                  As New ADODB.Recordset


Dim dTotalEarning       As Double
Dim dNetEarning         As Double
Dim iEmployeeKey
Dim iCompensationKey
