VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServiceChargeSummary 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11505
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServiceChargeSummary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picGenerateServiceCharge 
      Height          =   1815
      Left            =   3960
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3201
      BackColor       =   15396057
      Begin VB.TextBox txtDateToG 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmdOKGenerate 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   240
         Picture         =   "frmServiceChargeSummary.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1080
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelGenerate 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   1920
         Picture         =   "frmServiceChargeSummary.frx":1FF4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1080
         Width           =   1560
      End
      Begin VB.TextBox txtDateFromG 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   240
         MaxLength       =   50
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   45
         TabIndex        =   3
         Top             =   45
         Width           =   3650
         _ExtentX        =   6429
         _ExtentY        =   609
         Caption         =   "Generate Service Charge"
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
         Icon            =   "frmServiceChargeSummary.frx":2750
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   27
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   28
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
               Caption         =   "Print"
               Key             =   "Print"
               ImageKey        =   "IMG9"
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   " Post   "
               Key             =   "Post"
               ImageKey        =   "IMG10"
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               Key             =   "Refresh"
               ImageKey        =   "IMG12"
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
               ImageKey        =   "IMG13"
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         MousePointer    =   99
         MouseIcon       =   "frmServiceChargeSummary.frx":2CEA
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   9900
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   29
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmServiceChargeSummary.frx":3004
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
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2400
      ScaleHeight     =   1455
      ScaleWidth      =   6495
      TabIndex        =   1
      Top             =   1440
      Width           =   6495
      Begin VB.TextBox txtRankNFile 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   21
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtManagerial 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   19
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtForCompany 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   17
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtTotalSC 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   15
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtDateTo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   13
         Text            =   "11/20/2010"
         Top             =   0
         Width           =   1695
      End
      Begin VB.TextBox txtDateFrom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   11
         Text            =   "5/21/2010"
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Rank In File"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   22
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Managerial"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "For Company"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   18
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Service Charge"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Date To"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   14
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1320
      TabIndex        =   26
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1320
      TabIndex        =   25
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin RPVGCC.b8Container picProgress 
      Height          =   855
      Left            =   3360
      TabIndex        =   23
      Top             =   1800
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1508
      BackColor       =   15396057
      Begin VB.PictureBox picProgressBar 
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
         Height          =   615
         Left            =   120
         ScaleHeight     =   555
         ScaleWidth      =   4995
         TabIndex        =   24
         Top             =   120
         Width           =   5055
      End
   End
   Begin RPVGCC.b8Container picGenProgress 
      Height          =   1335
      Left            =   3360
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2355
      BackColor       =   15396057
      Begin VB.PictureBox picGenSubProgressBar 
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
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   4995
         TabIndex        =   10
         Top             =   720
         Width           =   5055
      End
      Begin VB.PictureBox picGenProgressBar 
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
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   4995
         TabIndex        =   9
         Top             =   120
         Width           =   5055
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   3555
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1940
            MinWidth        =   1940
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   26458
            MinWidth        =   26458
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9720
      Top             =   1560
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
            Picture         =   "frmServiceChargeSummary.frx":3717
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSummary.frx":43F1
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSummary.frx":50CB
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSummary.frx":5DA5
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSummary.frx":6A7F
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSummary.frx":7759
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSummary.frx":8433
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSummary.frx":910D
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSummary.frx":9DE7
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSummary.frx":A6C1
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSummary.frx":B39B
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSummary.frx":C075
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSummary.frx":CD4F
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSummary.frx":DA29
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeSummary.frx":E703
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmServiceChargeSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim iTotalSC        As Double
Dim iForCompany     As Double
Dim iNetSC          As Double
Dim iMngrCnt        As Double
Dim iRankCnt        As Double
Dim iMngrSC         As Double
Dim iRankSC         As Double
Dim iTotalCntDays   As Double
Dim iTotalCntPerDay As Double
Dim iDate           As Date
Dim iDateStart      As Date
Dim iDateEnd        As Date
Dim iMasterKey      As Double
Dim iMngrRate       As Double
Dim iTotalCntMngr   As Double
Dim iTotalCntRnF    As Double
Dim iPositionLevel  As Long
Dim iDayDuty        As Double
Dim iShare          As Double

Dim iWorkSheet      As Integer
Dim Filename        As String
Dim WorkbookName    As String

Dim i, j, ColTop, RowTop, ColCount, RowCount, strRange, a, b, C, ForCompanyMaster, iAmt, dComputeRange, _
k, l, TableName, Columns, sDetails1, sDetails2, iMonth, dtmDate, iDayFrom, iDayTo, iCat, iTotalRecord, _
Arr1, Arr2, sTotalShare, dTotalShare, TableName2, sFieldName, sFieldNameArr


Private Function BROWSER(dFrom, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If dFrom <> "" Then
            s = "SELECT TOP 1 tbl_Service_Charge_Summary.* " & _
                " FROM tbl_Service_Charge_Summary " & _
                " WHERE (DateFrom = '" & FormatDateTime(dFrom, vbShortDate) & "') " & _
                " ORDER BY DateFrom"
        Else
            s = "SELECT TOP 1 tbl_Service_Charge_Summary.* " & _
                " FROM tbl_Service_Charge_Summary " & _
                " ORDER BY DateFrom"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Service_Charge_Summary.* " & _
            " FROM tbl_Service_Charge_Summary " & _
            " ORDER BY DateFrom"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Service_Charge_Summary.* " & _
            " FROM tbl_Service_Charge_Summary " & _
            " WHERE (DateFrom < '" & FormatDateTime(dFrom, vbShortDate) & "') " & _
            " ORDER BY DateFrom DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Service_Charge_Summary.* " & _
            " FROM tbl_Service_Charge_Summary " & _
            " WHERE (DateFrom > '" & FormatDateTime(dFrom, vbShortDate) & "') " & _
            " ORDER BY DateFrom "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "SELECT TOP 1 tbl_Service_Charge_Summary.* " & _
            " FROM tbl_Service_Charge_Summary " & _
            " ORDER BY DateFrom DESC"
    Case Else: Exit Function
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtDateFrom.Text = Format(rs!DateFrom, "mm/dd/yyyy")
    txtDateTo.Text = Format(rs!DateTo, "mm/dd/yyyy")
    txtTotalSC.Text = Format(rs!TotalSC, "#,##0.00")
    txtForCompany.Text = Format(rs!ForCompany, "#,##0.00")
    txtManagerial.Text = Format(rs!Managerial, "#,##0.00")
    txtRankNFile.Text = Format(rs!RankInFile, "#,##0.00")
    StatusBar1.Panels(1).Text = rs!PK
    StatusBar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    imgPosted.Visible = IIf(rs!Locked = 1, True, False)
    
    SaveSetting App.EXEName, "ServiceChargeSumm", "ServChrgeSumm", Format(rs!DateFrom, "mm/dd/yyyy")
    
End If
rs.Close
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
        Case 6      '=== NOT EMPTY DETAIL NAME ===
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

Private Function PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If picGenerateServiceCharge.Visible = True Then Exit Function
If picGenProgress.Visible = True Then Exit Function
If picProgress.Visible = True Then Exit Function
If AccessRights("Service Charge Summary", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
picGenerateServiceCharge.ZOrder 0
txtDateFromG.Text = ""
txtDateToG.Text = ""
picGenerateServiceCharge.Visible = True
txtDateFromG.SetFocus
TRANSACTIONTYPE = is_ADDING
End Function

Private Function PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If picGenerateServiceCharge.Visible = True Then Exit Function
If picGenProgress.Visible = True Then Exit Function
If picProgress.Visible = True Then Exit Function
If StatusBar1.Panels(1).Text = "" Then Exit Function
If imgPosted.Visible = True Then MsgBox "Already Posted!                   ", vbCritical, "Error...": Exit Function
If AccessRights("Service Charge Summary", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
If MsgBox("ARE YOU SURE IN REGENERATING THIS SERVICE CHARGE SUMMARY?                ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Function
txtDateFromG.Text = txtDateFrom.Text
txtDateToG.Text = txtDateTo.Text
TRANSACTIONTYPE = is_EDITTING
cmdOKGenerate_Click
End Function

Private Function PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If picGenerateServiceCharge.Visible = True Then Exit Function
If picGenProgress.Visible = True Then Exit Function
If picProgress.Visible = True Then Exit Function
If StatusBar1.Panels(1).Text = "" Then Exit Function
If imgPosted.Visible = True Then MsgBox "Already Posted!                   ", vbCritical, "Error...": Exit Function
If AccessRights("Service Charge Summary", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
If MsgBox("ARE YOU SURE IN DELETING THIS SERVICE CHARGE SUMMARY?                ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Function
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_Service_Charge_Summary WHERE (PK = " & StatusBar1.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "ServiceChargeSumm", "ServChrgeSumm", ""), "is_PAGEDOWN"
If Trim(txtDateFrom.Text) = "" Then BROWSER GetSetting(App.EXEName, "ServiceChargeSumm", "ServChrgeSumm", ""), "is_HOME"
Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Function
End Function

Private Function PRESS_F8()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If picGenerateServiceCharge.Visible = True Then Exit Function
If picGenProgress.Visible = True Then Exit Function
If picProgress.Visible = True Then Exit Function
If StatusBar1.Panels(1).Text = "" Then Exit Function
If imgPosted.Visible = True Then MsgBox "Already Posted!                   ", vbCritical, "Error...": Exit Function
If AccessRights("Service Charge Summary", "Post") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
If MsgBox("ARE YOU SURE IN POSTING THIS SERVICE CHARGE SUMMARY?                ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Function
On Error GoTo PG:
ConnOmega.Execute "UPDATE tbl_Service_Charge_Summary SET Locked = 1 WHERE (PK = " & StatusBar1.Panels(1).Text & ")"
BROWSER GetSetting(App.EXEName, "ServiceChargeSumm", "ServChrgeSumm", ""), "is_LOAD"
Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Function
End Function

Private Function PRESS_F9()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If picGenerateServiceCharge.Visible = True Then Exit Function
If picGenProgress.Visible = True Then Exit Function
If picProgress.Visible = True Then Exit Function
If StatusBar1.Panels(1).Text = "" Then Exit Function
Command2_Click
End Function

Private Function PRESS_ESCAPE()
If picGenerateServiceCharge.Visible = True Then cmdCancelGenerate_Click: Exit Function
If picGenProgress.Visible = True Then Exit Function
If picProgress.Visible = True Then Exit Function
If TRANSACTIONTYPE = is_REFRESH Then Unload Me
End Function

Private Function CLEARTEXT()
txtDateFrom.Text = ""
txtDateTo.Text = ""
txtTotalSC.Text = ""
txtForCompany.Text = ""
txtManagerial.Text = ""
txtRankNFile.Text = ""
StatusBar1.Panels(1).Text = ""
StatusBar1.Panels(2).Text = ""
imgPosted.Visible = False
End Function

Private Sub b8TitleBar1_CLoseClick()
cmdCancelGenerate_Click
End Sub

Private Sub cmdCancelGenerate_Click()
picGenerateServiceCharge.Visible = False
picToolbar.Enabled = True
picMain.Enabled = True
TRANSACTIONTYPE = is_REFRESH
End Sub

Private Sub cmdOKGenerate_Click()
If IsDate(txtDateFromG.Text) = False Then MsgBox "Please Supply a Valid Date!               ", vbCritical, "Error...": txtDateFromG.SetFocus: Exit Sub
txtDateFromG.Text = Format(FormatDateTime(txtDateFromG.Text, vbShortDate), "mm/dd/yyyy")
s = "SELECT tbl_Service_Charge_CutOff.* " & _
    " From tbl_Service_Charge_CutOff " & _
    " WHERE (MonthFrom = " & Month(FormatDateTime(txtDateFromG.Text, vbShortDate)) & ") " & _
    " AND (DayFrom = " & Day(FormatDateTime(txtDateFromG.Text, vbShortDate)) & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtDateToG.Text = Format(DateSerial(Year(FormatDateTime(txtDateFromG.Text, vbShortDate)) + rs!YearTo, rs!MonthTo, rs!DayTo), "mm/dd/yyyy")
Else
    MsgBox "Invalid CutOff!                     ", vbCritical, "Error..."
    txtDateFromG.Text = ""
    txtDateToG.Text = ""
    rs.Close
    Exit Sub
End If
rs.Close

'MsgBox "success"
'Exit Sub

iTotalSC = 0: iForCompany = 0: iNetSC = 0
iMngrCnt = 0: iRankCnt = 0: iMngrSC = 0
iRankSC = 0: iTotalCntPerDay = 0

iTotalCntDays = DateDiff("d", FormatDateTime(txtDateFromG.Text, vbShortDate), FormatDateTime(txtDateToG.Text, vbShortDate))
iDateStart = FormatDateTime(txtDateFromG.Text, vbShortDate)
iDateEnd = FormatDateTime(txtDateToG.Text, vbShortDate)

picGenProgress.ZOrder 0
picGenProgressBar.BackColor = &HFFFFFF
picGenSubProgressBar.BackColor = &HFFFFFF
picGenerateServiceCharge.Visible = False
picGenProgress.Visible = True

s = "SELECT PK " & _
    " FROM tbl_Service_Charge_Summary " & _
    " WHERE (DateFrom = '" & FormatDateTime(iDateStart, vbShortDate) & "') " & _
    " AND (DateTo = '" & FormatDateTime(iDateEnd, vbShortDate) & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    ConnOmega.Execute "DELETE FROM tbl_Service_Charge_Summary " & _
                      " WHERE (PK = " & rs!PK & ")"
End If
rs.Close

ConnOmega.Execute "INSERT INTO tbl_Service_Charge_Summary " & _
                  " (DateFrom, DateTo) " & _
                  " VALUES ('" & FormatDateTime(iDateStart, vbShortDate) & "', " & _
                  " '" & FormatDateTime(iDateEnd, vbShortDate) & "')"

s = "SELECT PK " & _
    " FROM tbl_Service_Charge_Summary " & _
    " WHERE (DateFrom = '" & FormatDateTime(iDateStart, vbShortDate) & "') " & _
    " AND (DateTo = '" & FormatDateTime(iDateEnd, vbShortDate) & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iMasterKey = rs!PK
End If
rs.Close

For a = 0 To iTotalCntDays
    iDate = DateAdd("d", a, iDateStart)
    iTotalSC = 0: iForCompany = 0: iNetSC = 0
    iMngrCnt = 0: iRankCnt = 0: iMngrSC = 0
    iRankSC = 0: iTotalCntPerDay = 0
    ForCompanyMaster = 0: iMngrRate = 0
    picGenSubProgressBar.BackColor = &HFFFFFF
    
    's = "SELECT ServiceCharge " & _
        " From tbl_Service_Charge_Detail " & _
        " WHERE (sDate = '" & FormatDateTime(iDate, vbShortDate) & "')"
    s = "SELECT dbo.tbl_Service_Charge_Detail.ServiceCharge " & _
        " FROM  dbo.tbl_Service_Charge_Detail LEFT OUTER JOIN " & _
        " dbo.tbl_Service_Charge ON dbo.tbl_Service_Charge_Detail.MasterKey = dbo.tbl_Service_Charge.PK " & _
        " WHERE (dbo.tbl_Service_Charge_Detail.sDate = '" & FormatDateTime(iDate, vbShortDate) & "') " & _
        " AND (dbo.tbl_Service_Charge.Posted = 1)"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        iTotalSC = rs!ServiceCharge
    Else
        GoTo GONEXT:
    End If
    rs.Close
    
    If CDbl(iTotalSC) = 0 Then GoTo GONEXT:
    
    '== For Company Cnt
    s = "SELECT TOP 1 PK, SupervisoryPerc " & _
        " From tbl_Service_Charge_Setup " & _
        " WHERE (EffectDate <= '" & FormatDateTime(iDate, vbShortDate) & "') " & _
        " ORDER BY EffectDate DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        ForCompanyMaster = rs!PK
        iMngrRate = rs!SupervisoryPerc
    End If
    rs.Close
    s = "SELECT tbl_Service_Charge_SetupDetail.* " & _
        " From tbl_Service_Charge_SetupDetail " & _
        " WHERE (MasterKey = " & ForCompanyMaster & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    iTotalCntPerDay = iTotalCntPerDay + rs.RecordCount
    rs.Close
    '== For Employee Cnt
    s = "sp_Service_Charge_GenerationV2('" & FormatDateTime(iDate, vbShortDate) & "', '" & FormatDateTime(iDateEnd, vbShortDate) & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    iTotalCntPerDay = iTotalCntPerDay + rs.RecordCount
    rs.Close
    
    '== For Employee Cnt
    s = "sp_Service_Charge_GenerationV2('" & FormatDateTime(iDate, vbShortDate) & "', '" & FormatDateTime(iDateEnd, vbShortDate) & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    iTotalCntPerDay = iTotalCntPerDay + rs.RecordCount
    rs.Close
    
    ConnOmega.Execute "INSERT INTO tbl_Service_Charge_Daily " & _
                      " (SummKey, iDate, TotalSC) " & _
                      " VALUES (" & iMasterKey & ", " & _
                      " '" & FormatDateTime(iDate, vbShortDate) & "', " & _
                      " " & CDbl(iTotalSC) & ")"
    
    ConnOmega.Execute "UPDATE tbl_Service_Charge_Summary " & _
                      " SET TotalSC = TotalSC + " & CDbl(iTotalSC) & " " & _
                      " WHERE (PK = " & iMasterKey & ")"
    
    b = 0
    s = "SELECT tbl_Service_Charge_SetupDetail.* " & _
        " From tbl_Service_Charge_SetupDetail " & _
        " WHERE (MasterKey = " & ForCompanyMaster & ")"
    If rc.State = adStateOpen Then rc.Close
    rc.Open s, ConnOmega
    While Not rc.EOF
        DoEvents
        b = b + 1
        iAmt = Format(CDbl(iTotalSC) * CDbl(rc!RatePerc), "#,##0.00")
        iForCompany = iForCompany + CDbl(iAmt)
        ConnOmega.Execute "INSERT INTO tbl_Service_Charge_ForCompany " & _
                          " (SummKey, iDate, ForCompany, Amount) " & _
                          " VALUES (" & iMasterKey & ", '" & FormatDateTime(iDate, vbShortDate) & "', " & _
                          " '" & FORMATSQL(rc!ForCompany) & "', " & CDbl(iAmt) & ")"
                      
        UpdateProgress_Caption rc!ForCompany, picGenSubProgressBar, b / iTotalCntPerDay
        rc.MoveNext
    Wend
    rc.Close
    
    ConnOmega.Execute "UPDATE tbl_Service_Charge_Daily " & _
                      " SET ForCompany = " & CDbl(iForCompany) & " " & _
                      " WHERE (SummKey = " & iMasterKey & ") " & _
                      " AND (iDate = '" & FormatDateTime(iDate, vbShortDate) & "')"
    
    ConnOmega.Execute "UPDATE tbl_Service_Charge_Summary " & _
                      " SET ForCompany = ForCompany + " & CDbl(iForCompany) & " " & _
                      " WHERE (PK = " & iMasterKey & ")"
                      
    iNetSC = CDbl(iTotalSC) - CDbl(iForCompany)
    iMngrSC = Format(CDbl(iNetSC) * CDbl(iMngrRate), "#,##0.00")
    iRankSC = CDbl(iNetSC) - CDbl(iMngrSC)
    
    iPositionLevel = 0
    iTotalCntMngr = 0: iTotalCntRnF = 0
    
    '== For Employee
    s = "sp_Service_Charge_GenerationV2('" & FormatDateTime(iDate, vbShortDate) & "', '" & FormatDateTime(iDateEnd, vbShortDate) & "')"
    If rc.State = adStateOpen Then rc.Close
    rc.Open s, ConnOmega
    While Not rc.EOF
        DoEvents
        b = b + 1
        Select Case rc!PositionLevel
            Case 1
                't = "SELECT tbl_Absent_Employee_Detail.TotalHours " & _
                    " FROM tbl_Absent_Employee LEFT OUTER JOIN " & _
                    " tbl_Absent_Employee_Detail ON tbl_Absent_Employee.PK = tbl_Absent_Employee_Detail.MasterKey " & _
                    " WHERE (tbl_Absent_Employee.DateApplied = '" & FormatDateTime(iDate, vbShortDate) & "') " & _
                    " AND (tbl_Absent_Employee_Detail.EmpKey = " & rc!PK & ")"
                t = "SELECT dbo.tbl_Personnel_AbsentLateUndertime_Details.TotalHours " & _
                    " FROM  dbo.tbl_Personnel_AbsentLateUndertime_Details LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_AbsentLateUndertime ON dbo.tbl_Personnel_AbsentLateUndertime_Details.MasterKey = dbo.tbl_Personnel_AbsentLateUndertime.PK " & _
                    " WHERE (dbo.tbl_Personnel_AbsentLateUndertime_Details.EmployeeKey = " & rc!PK & ") " & _
                    " AND (dbo.tbl_Personnel_AbsentLateUndertime.DateApplied = '" & FormatDateTime(iDate, vbShortDate) & "') " & _
                    " AND (dbo.tbl_Personnel_AbsentLateUndertime.Posted = 1)"
                If rt.State = adStateOpen Then rt.Close
                rt.Open t, ConnOmega
                If rt.RecordCount > 0 Then
                    iTotalCntRnF = iTotalCntRnF + Format(((8 - CDbl(rt!TotalHours)) / 8), "#,##0.0000")
                Else
                    iTotalCntRnF = iTotalCntRnF + 1
                End If
                rt.Close
                
            Case 2
                't = "SELECT tbl_Absent_Employee_Detail.TotalHours " & _
                    " FROM tbl_Absent_Employee LEFT OUTER JOIN " & _
                    " tbl_Absent_Employee_Detail ON tbl_Absent_Employee.PK = tbl_Absent_Employee_Detail.MasterKey " & _
                    " WHERE (tbl_Absent_Employee.DateApplied = '" & FormatDateTime(iDate, vbShortDate) & "') " & _
                    " AND (tbl_Absent_Employee_Detail.EmpKey = " & rc!PK & ")"
                t = "SELECT dbo.tbl_Personnel_AbsentLateUndertime_Details.TotalHours " & _
                    " FROM  dbo.tbl_Personnel_AbsentLateUndertime_Details LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_AbsentLateUndertime ON dbo.tbl_Personnel_AbsentLateUndertime_Details.MasterKey = dbo.tbl_Personnel_AbsentLateUndertime.PK " & _
                    " WHERE (dbo.tbl_Personnel_AbsentLateUndertime_Details.EmployeeKey = " & rc!PK & ") " & _
                    " AND (dbo.tbl_Personnel_AbsentLateUndertime.DateApplied = '" & FormatDateTime(iDate, vbShortDate) & "') " & _
                    " AND (dbo.tbl_Personnel_AbsentLateUndertime.Posted = 1)"
                If rt.State = adStateOpen Then rt.Close
                rt.Open t, ConnOmega
                If rt.RecordCount > 0 Then
                    iTotalCntMngr = iTotalCntMngr + Format(((8 - CDbl(rt!TotalHours)) / 8), "#,##0.0000")
                Else
                    iTotalCntMngr = iTotalCntMngr + 1
                End If
                rt.Close
                
        End Select
        
        'UpdateProgress_Caption "Calculating Managerial and Rank in File", picGenSubProgressBar, b / iTotalCntPerDay
        UpdateProgress_Caption rc!EmployeeName, picGenSubProgressBar, b / iTotalCntPerDay
        rc.MoveNext
    Wend
    rc.Close
    
    s = "sp_Service_Charge_GenerationV2('" & FormatDateTime(iDate, vbShortDate) & "', '" & FormatDateTime(iDateEnd, vbShortDate) & "')"
    If rc.State = adStateOpen Then rc.Close
    rc.Open s, ConnOmega
    While Not rc.EOF
        DoEvents
        b = b + 1
        iDayDuty = 0
        Select Case rc!PositionLevel
            Case 1
                't = "SELECT tbl_Absent_Employee_Detail.TotalHours " & _
                    " FROM tbl_Absent_Employee LEFT OUTER JOIN " & _
                    " tbl_Absent_Employee_Detail ON tbl_Absent_Employee.PK = tbl_Absent_Employee_Detail.MasterKey " & _
                    " WHERE (tbl_Absent_Employee.DateApplied = '" & FormatDateTime(iDate, vbShortDate) & "') " & _
                    " AND (tbl_Absent_Employee_Detail.EmpKey = " & rc!PK & ")"
                t = "SELECT dbo.tbl_Personnel_AbsentLateUndertime_Details.TotalHours " & _
                    " FROM  dbo.tbl_Personnel_AbsentLateUndertime_Details LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_AbsentLateUndertime ON dbo.tbl_Personnel_AbsentLateUndertime_Details.MasterKey = dbo.tbl_Personnel_AbsentLateUndertime.PK " & _
                    " WHERE (dbo.tbl_Personnel_AbsentLateUndertime_Details.EmployeeKey = " & rc!PK & ") " & _
                    " AND (dbo.tbl_Personnel_AbsentLateUndertime.DateApplied = '" & FormatDateTime(iDate, vbShortDate) & "') " & _
                    " AND (dbo.tbl_Personnel_AbsentLateUndertime.Posted = 1)"
                If rt.State = adStateOpen Then rt.Close
                rt.Open t, ConnOmega
                If rt.RecordCount > 0 Then
                    iDayDuty = Format(((8 - CDbl(rt!TotalHours)) / 8), "#,##0.0000")
                Else
                    iDayDuty = 1
                End If
                rt.Close
                
                iShare = Format((CDbl(iRankSC) * CDbl(iDayDuty)) / CDbl(iTotalCntRnF), "#,##0.00")
                
                ConnOmega.Execute "INSERT INTO tbl_Service_Charge_Employee " & _
                                  " (SummKey, EmpPK, iDate, iCategory, DayDuty, iShare) " & _
                                  " VALUES (" & iMasterKey & ", " & rc!PK & ", " & _
                                  " '" & FormatDateTime(iDate, vbShortDate) & "', " & _
                                  " " & rc!PositionLevel & ", " & CDbl(iDayDuty) & ", " & CDbl(iShare) & " )"
                
                ConnOmega.Execute "UPDATE tbl_Service_Charge_Daily " & _
                                  " SET RanInFile = RanInFile + " & CDbl(iShare) & " " & _
                                  " WHERE (SummKey = " & iMasterKey & ") " & _
                                  " AND (iDate = '" & FormatDateTime(iDate, vbShortDate) & "')"
                
                ConnOmega.Execute "UPDATE tbl_Service_Charge_Summary " & _
                                  " SET RankInFile = RankInFile + " & CDbl(iShare) & " " & _
                                  " WHERE (PK = " & iMasterKey & ")"
            Case 2
                't = "SELECT tbl_Absent_Employee_Detail.TotalHours " & _
                    " FROM tbl_Absent_Employee LEFT OUTER JOIN " & _
                    " tbl_Absent_Employee_Detail ON tbl_Absent_Employee.PK = tbl_Absent_Employee_Detail.MasterKey " & _
                    " WHERE (tbl_Absent_Employee.DateApplied = '" & FormatDateTime(iDate, vbShortDate) & "') " & _
                    " AND (tbl_Absent_Employee_Detail.EmpKey = " & rc!PK & ")"
                t = "SELECT dbo.tbl_Personnel_AbsentLateUndertime_Details.TotalHours " & _
                    " FROM  dbo.tbl_Personnel_AbsentLateUndertime_Details LEFT OUTER JOIN " & _
                    " dbo.tbl_Personnel_AbsentLateUndertime ON dbo.tbl_Personnel_AbsentLateUndertime_Details.MasterKey = dbo.tbl_Personnel_AbsentLateUndertime.PK " & _
                    " WHERE (dbo.tbl_Personnel_AbsentLateUndertime_Details.EmployeeKey = " & rc!PK & ") " & _
                    " AND (dbo.tbl_Personnel_AbsentLateUndertime.DateApplied = '" & FormatDateTime(iDate, vbShortDate) & "') " & _
                    " AND (dbo.tbl_Personnel_AbsentLateUndertime.Posted = 1)"
                If rt.State = adStateOpen Then rt.Close
                rt.Open t, ConnOmega
                If rt.RecordCount > 0 Then
                    iDayDuty = Format(((8 - CDbl(rt!TotalHours)) / 8), "#,##0.0000")
                Else
                    iDayDuty = 1
                End If
                rt.Close
                
                iShare = Format((CDbl(iMngrSC) * CDbl(iDayDuty)) / CDbl(iTotalCntMngr), "#,##0.00")
                
                ConnOmega.Execute "INSERT INTO tbl_Service_Charge_Employee " & _
                                  " (SummKey, EmpPK, iDate, iCategory, DayDuty, iShare) " & _
                                  " VALUES (" & iMasterKey & ", " & rc!PK & ", " & _
                                  " '" & FormatDateTime(iDate, vbShortDate) & "', " & _
                                  " " & rc!PositionLevel & ", " & CDbl(iDayDuty) & ", " & CDbl(iShare) & " )"
                                  
                ConnOmega.Execute "UPDATE tbl_Service_Charge_Daily " & _
                                  " SET Managerial = Managerial + " & CDbl(iShare) & " " & _
                                  " WHERE (SummKey = " & iMasterKey & ") " & _
                                  " AND (iDate = '" & FormatDateTime(iDate, vbShortDate) & "')"
                                  
                ConnOmega.Execute "UPDATE tbl_Service_Charge_Summary " & _
                                  " SET Managerial = Managerial + " & CDbl(iShare) & " " & _
                                  " WHERE (PK = " & iMasterKey & ")"
        End Select
        '"Calculating Managerial and Rank in File"
        UpdateProgress_Caption rc!EmployeeName, picGenSubProgressBar, b / iTotalCntPerDay
        'UpdateProgress_Caption "Calculating Managerial and Rank in File", picGenSubProgressBar, b / iTotalCntPerDay
        rc.MoveNext
    Wend
    rc.Close
    
GONEXT:
    UpdateProgress_Caption Format(iDate, "dd-mmm-yyyy"), picGenProgressBar, a / iTotalCntDays
Next a

ConnOmega.Execute "UPDATE tbl_Service_Charge_Summary " & _
                  " SET LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                  " WHERE (PK = " & iMasterKey & ")"

picGenProgress.Visible = False
picToolbar.Enabled = True
picMain.Enabled = True

TRANSACTIONTYPE = is_REFRESH
BROWSER FormatDateTime(txtDateFromG.Text, vbShortDate), "is_LOAD"

End Sub

Private Sub Command1_Click()

MainForm.CommonDialog1.CancelError = True
On Error GoTo ErrorHandler
MainForm.CommonDialog1.DialogTitle = "Save"
MainForm.CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
MainForm.CommonDialog1.ShowSave
Filename = Trim(MainForm.CommonDialog1.Filename)

WorkbookName = CStr(Filename)

Screen.MousePointer = vbHourglass

Set xlsApp = CreateObject("Excel.Application")
xlsApp.Visible = False
iWorkSheet = 1
xlsApp.Workbooks.Add
xlsApp.DisplayAlerts = False
xlsApp.Workbooks(1).Sheets(2).Delete
xlsApp.Workbooks(1).Sheets(2).Delete

xlsApp.Workbooks(1).Sheets(iWorkSheet).Activate
xlsApp.Workbooks(1).Sheets(iWorkSheet).Name = "Computation"

ColTop = 0: RowTop = 0
ColCount = 0: RowCount = 0

RowCount = RowCount + 1
j = 0
j = j + 1
strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = gbl_CompanyName
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCount = RowCount + 1
strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Service Charge"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCount = RowCount + 1
strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "From " & Format(txtDateFrom.Text, "mmmm dd, yyyy") & " to " & Format(txtDateTo.Text, "mmmm dd, yyyy")
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCount = RowCount + 1
strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCount = RowCount + 1
strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Date"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

j = j + 1
strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Total"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

s = "SELECT TOP 1 PK " & _
    " From tbl_Service_Charge_Setup " & _
    " WHERE (EffectDate <= '" & FormatDateTime(txtDateFrom.Text, vbShortDate) & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    t = "SELECT ForCompany, Rate " & _
        " From tbl_Service_Charge_SetupDetail " & _
        " Where (MasterKey = " & rs!PK & ") " & _
        " ORDER BY ForCompany"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        j = j + 1
        strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rt!ForCompany
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
        rt.MoveNext
    Wend
    rt.Close
End If
rs.Close

j = j + 1
strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Net"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

j = j + 1
strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Rank In File"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

j = j + 1
strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Managerial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCount = RowCount + 1
j = 0
j = j + 1
strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

j = j + 1
strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Service Charge"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

s = "SELECT TOP 1 PK " & _
    " From tbl_Service_Charge_Setup " & _
    " WHERE (EffectDate <= '" & FormatDateTime(txtDateFrom.Text, vbShortDate) & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    t = "SELECT ForCompany, Rate " & _
        " From tbl_Service_Charge_SetupDetail " & _
        " Where (MasterKey = " & rs!PK & ") " & _
        " ORDER BY ForCompany"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        j = j + 1
        strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = CStr(rt!Rate) & "%"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
        rt.MoveNext
    Wend
    rt.Close
End If
rs.Close

j = j + 1
strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Service Charge"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

j = j + 1
strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "85%"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

j = j + 1
strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "15%"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True


RowCount = RowCount + 1
j = 0
j = j + 1
strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

s = "SELECT iDate, TotalSC, Managerial, RanInFile " & _
    " From tbl_Service_Charge_Daily " & _
    " WHERE (iDate >= '" & FormatDateTime(txtDateFrom.Text, vbShortDate) & "') " & _
    " AND (iDate <= '" & FormatDateTime(txtDateTo.Text, vbShortDate) & "') " & _
    " ORDER BY iDate"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    RowCount = RowCount + 1
    j = 0
    dComputeRange = ""
    j = j + 1
    strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = Format(rs!iDate, "dd-mmm-yy")
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    
    j = j + 1
    strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = Format(rs!TotalSC, "#,##0.00")
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    dComputeRange = dComputeRange & "=" & strRange
    t = "SELECT Amount " & _
        " From tbl_Service_Charge_ForCompany " & _
        " WHERE (iDate = '" & FormatDateTime(rs!iDate, vbShortDate) & "') " & _
        " ORDER BY ForCompany"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        
        j = j + 1
        strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = Format(rt!Amount, "#,##0.00")
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
        dComputeRange = dComputeRange & "-" & strRange
        rt.MoveNext
    Wend
    rt.Close
    
    j = j + 1
    strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = dComputeRange
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    
    j = j + 1
    strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = Format(rs!RanInFile, "#,##0.00")
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    
    j = j + 1
    strRange = (Chr$(IIf(CDbl(j) > 26, 64 + 1, 64) + j)) & CStr(RowCount)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = Format(rs!Managerial, "#,##0.00")
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    
    rs.MoveNext
Wend
rs.Close



Screen.MousePointer = vbDefault

If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

xlsApp.Visible = True
Exit Sub
ErrorHandler:
Exit Sub
End Sub

Private Sub Command2_Click()


MainForm.CommonDialog1.CancelError = True

On Error GoTo ErrorHandler
MainForm.CommonDialog1.DialogTitle = "Save"
MainForm.CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
MainForm.CommonDialog1.ShowSave
Filename = Trim(MainForm.CommonDialog1.Filename)

WorkbookName = CStr(Filename)

Screen.MousePointer = vbHourglass
On Error GoTo PG:
TableName = "tmp_" & gbl_UserName & "_ServiceCharge_Report"

'GoTo EXCEL:

Columns = ""
Columns = Columns & "|EmployeeName:varchar:(255):NOT NULL:DEFAULT('')"
Columns = Columns & "|iCategory:int:NOT NULL:DEFAULT(0)"
Columns = Columns & "|AccountNo:varchar:(255):NOT NULL:DEFAULT('')"
Columns = Columns & "|iMonth:int:NOT NULL:DEFAULT(0)"
Columns = Columns & "|iYear:int:NOT NULL:DEFAULT(0)"
Columns = Columns & "|sDetails1:varchar:(2000):NOT NULL:DEFAULT('')"
Columns = Columns & "|sDetails2:varchar:(2000):NOT NULL:DEFAULT('')"
CreateTable gbl_Database, TableName, Columns
picProgressBar.BackColor = &HFFFFFF
picProgress.ZOrder 0
picProgress.Visible = True
i = 0

's = "sp_Service_Charge_Report_01('" & FormatDateTime(txtDateFrom.Text, vbShortDate) & "', '" & FormatDateTime(txtDateTo.Text, vbShortDate) & "')"
s = "SELECT tbl_Service_Charge_Employee.EmpPK, " & _
    " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
    " tbl_Service_Charge_Employee.iCategory, MONTH(tbl_Service_Charge_Employee.iDate) AS iMonth, " & _
    " YEAR(tbl_Service_Charge_Employee.iDate) AS iYear, tbl_Personnel_IDNumber.AccountNumber " & _
    " FROM tbl_Service_Charge_Employee LEFT OUTER JOIN " & _
    " tbl_Personnel_IDNumber ON tbl_Service_Charge_Employee.EmpPK = tbl_Personnel_IDNumber.PK LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
    " WHERE (tbl_Service_Charge_Employee.iDate >= '" & FormatDateTime(txtDateFrom.Text, vbShortDate) & "') " & _
    " AND (tbl_Service_Charge_Employee.iDate <= '" & FormatDateTime(txtDateTo.Text, vbShortDate) & "') " & _
    " AND (tbl_Service_Charge_Employee.SummKey = " & StatusBar1.Panels(1).Text & ") " & _
    " GROUP BY tbl_Service_Charge_Employee.EmpPK, " & _
    " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName, " & _
    " Month (tbl_Service_Charge_Employee.iDate), tbl_Service_Charge_Employee.iCategory, YEAR(tbl_Service_Charge_Employee.iDate), tbl_Personnel_IDNumber.AccountNumber " & _
    " ORDER BY YEAR(tbl_Service_Charge_Employee.iDate), MONTH(tbl_Service_Charge_Employee.iDate), " & _
    " tbl_Service_Charge_Employee.iCategory DESC, " & _
    " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName + '  ' + tbl_Personnel_Information.MiddleName"
If rc.State = adStateOpen Then rc.Close
rc.Open s, ConnOmega
While Not rc.EOF
    DoEvents
    i = i + 1
    If CDbl(Month(FormatDateTime(txtDateFrom.Text, vbShortDate))) = CDbl(rc!iMonth) Then
        iDayFrom = Day(FormatDateTime(txtDateFrom.Text, vbShortDate))
        iDayTo = Day(DateSerial(Year(FormatDateTime(txtDateFrom.Text, vbShortDate)), Month(FormatDateTime(txtDateFrom.Text, vbShortDate)) + 1, 0))
    Else
        If CDbl(Month(FormatDateTime(txtDateTo.Text, vbShortDate))) = CDbl(rc!iMonth) Then
            iDayFrom = 1
            iDayTo = Day(FormatDateTime(txtDateTo.Text, vbShortDate))
        Else
            iDayFrom = 1
            iDayTo = Day(DateSerial(rc!iYear, rc!iMonth + 1, 0))
        End If
    End If
    
    sDetails1 = "": sDetails2 = ""
    For j = iDayFrom To iDayTo
        dtmDate = DateSerial(rc!iYear, rc!iMonth, j)
        If CDbl(j) >= 1 And CDbl(j) <= 15 Then
            't = "sp_Service_Charge_Report_02 (" & rs!EmpPK & ", '" & FormatDateTime(dtmDate, vbShortDate) & "')"
            t = "SELECT EmpPK, iDate, DayDuty, iShare, " & _
                "(SELECT TotalSC " & _
                " From tbl_Service_Charge_Daily " & _
                " WHERE (iDate = tbl_Service_Charge_Employee.iDate) " & _
                " AND (SummKey = tbl_Service_Charge_Employee.SummKey)) AS TotalSC " & _
                " From tbl_Service_Charge_Employee " & _
                " WHERE (EmpPK = " & rc!EmpPK & ") " & _
                " AND (iDate = '" & FormatDateTime(dtmDate, vbShortDate) & "') " & _
                " AND (SummKey = " & StatusBar1.Panels(1).Text & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                sDetails1 = sDetails1 & "|" & Format(rt!iDate, "dd") & _
                                      "\" & Format(rt!TotalSC, "#,##0.00") & _
                                      "\" & Format(rt!DayDuty, "#0.00") & _
                                      "\" & Format(rt!iShare, "#,##0.00")
            Else
                sDetails1 = sDetails1 & "|" & Format(dtmDate, "dd") & _
                                      "\" & Format(0, "#,##0.00") & _
                                      "\" & Format(0, "#0.00") & _
                                      "\" & Format(0, "#,##0.00")
            End If
            rt.Close
        ElseIf CDbl(j) >= 16 Then
            't = "sp_Service_Charge_Report_02 (" & rs!EmpPK & ", '" & FormatDateTime(dtmDate, vbShortDate) & "')"
            t = "SELECT EmpPK, iDate, DayDuty, iShare, " & _
                "(SELECT TotalSC " & _
                " From tbl_Service_Charge_Daily " & _
                " WHERE (iDate = tbl_Service_Charge_Employee.iDate) " & _
                " AND (SummKey = tbl_Service_Charge_Employee.SummKey)) AS TotalSC " & _
                " From tbl_Service_Charge_Employee " & _
                " WHERE (EmpPK = " & rc!EmpPK & ") " & _
                " AND (iDate = '" & FormatDateTime(dtmDate, vbShortDate) & "') " & _
                " AND (SummKey = " & StatusBar1.Panels(1).Text & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                sDetails2 = sDetails2 & "|" & Format(rt!iDate, "dd") & _
                                      "\" & Format(rt!TotalSC, "#,##0.00") & _
                                      "\" & Format(rt!DayDuty, "#0.00") & _
                                      "\" & Format(rt!iShare, "#,##0.00")
            Else
                sDetails2 = sDetails2 & "|" & Format(dtmDate, "dd") & _
                                      "\" & Format(0, "#,##0.00") & _
                                      "\" & Format(0, "#0.00") & _
                                      "\" & Format(0, "#,##0.00")
            End If
            rt.Close
        End If
        
    Next j
    
    ConnOmega.Execute "INSERT INTO " & TableName & " " & _
                      " (EmployeeName, iCategory, AccountNo, iMonth, iYear, sDetails1, sDetails2) " & _
                      " VALUES ('" & rc!EmployeeName & "', " & _
                      " " & rc!iCategory & ", " & _
                      " '" & Replace(Trim(rc!AccountNumber), "-", "") & "', " & _
                      " " & rc!iMonth & ", " & _
                      " " & rc!iYear & ", " & _
                      " '" & Mid(sDetails1, 2, Len(sDetails1)) & "', " & _
                      " '" & Mid(sDetails2, 2, Len(sDetails2)) & "')"
    
    UpdateProgress_Caption "Generating Report", picProgressBar, i / rc.RecordCount
    rc.MoveNext
Wend
rc.Close

EXCEL:

'== Exporting to Excel
picProgressBar.BackColor = &HFFFFFF
iCat = 0: j = 0: k = 0: iMonth = 0
DoEvents

Set xlsApp = CreateObject("Excel.Application")
xlsApp.Visible = False
xlsApp.Workbooks.Add
xlsApp.DisplayAlerts = False
If xlsApp.Workbooks(1).Sheets.Count = 3 Then
    xlsApp.Workbooks(1).Sheets(2).Delete
    xlsApp.Workbooks(1).Sheets(2).Delete
End If

ColTop = 0: RowTop = 0
ColCount = 0: RowCount = 0
iTotalRecord = 0

s = "SELECT * FROM " & TableName & ""
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
iTotalRecord = iTotalRecord + rs.RecordCount
rs.Close

s = "SELECT iCategory, EmployeeName " & _
    " From " & TableName & " " & _
    " GROUP BY iCategory, EmployeeName " & _
    " ORDER BY iCategory DESC, EmployeeName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
iTotalRecord = iTotalRecord + rs.RecordCount
rs.Close

s = "SELECT iDate, TotalSC, Managerial, RanInFile " & _
    " From tbl_Service_Charge_Daily " & _
    " WHERE (iDate >= '" & FormatDateTime(txtDateFrom.Text, vbShortDate) & "') " & _
    " AND (iDate <= '" & FormatDateTime(txtDateTo.Text, vbShortDate) & "') " & _
    " AND (SummKey = " & StatusBar1.Panels(1).Text & ") " & _
    " ORDER BY iDate"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
iTotalRecord = iTotalRecord + rs.RecordCount
rs.Close

TableName2 = "tmp_" & gbl_UserName & "_ServiceCharge_Report_Det"
Columns = ""
Columns = Columns & "|EmployeeName:varchar:(255):NOT NULL:DEFAULT('')"
Columns = Columns & "|iCategory:int:NOT NULL:DEFAULT(0)"
Columns = Columns & "|AccountNo:varchar:(255):NOT NULL:DEFAULT('')"

s = "SELECT iYear, iMonth " & _
    " From " & TableName & " " & _
    " GROUP BY iYear, iMonth " & _
    " ORDER BY iYear, iMonth"
If rc.State = adStateOpen Then rc.Close
rc.Open s, ConnOmega
While Not rc.EOF
    Columns = Columns & "|" & IIf(CDbl(rc!iMonth) = 1, "Jan", _
                              IIf(CDbl(rc!iMonth) = 2, "Feb", _
                              IIf(CDbl(rc!iMonth) = 3, "Mar", _
                              IIf(CDbl(rc!iMonth) = 4, "Apr", _
                              IIf(CDbl(rc!iMonth) = 5, "May", _
                              IIf(CDbl(rc!iMonth) = 6, "Jun", _
                              IIf(CDbl(rc!iMonth) = 7, "Jul", _
                              IIf(CDbl(rc!iMonth) = 8, "Aug", _
                              IIf(CDbl(rc!iMonth) = 9, "Sep", _
                              IIf(CDbl(rc!iMonth) = 10, "Oct", _
                              IIf(CDbl(rc!iMonth) = 11, "Nov", _
                              IIf(CDbl(rc!iMonth) = 12, "Dec", "")))))))))))) & _
                              "_" & rc!iYear & ":float:NOT NULL:DEFAULT(0)"
    rc.MoveNext
Wend
rc.Close

CreateTable gbl_Database, TableName2, Columns

s = "SELECT iYear, iMonth " & _
    " From " & TableName & " " & _
    " GROUP BY iYear, iMonth " & _
    " ORDER BY iYear, iMonth"
If rc.State = adStateOpen Then rc.Close
rc.Open s, ConnOmega
For i = 1 To rc.RecordCount
    xlsApp.Workbooks(1).Sheets.Add
Next i
xlsApp.Workbooks(1).Sheets.Add

i = 0
While Not rc.EOF
    DoEvents
    iWorkSheet = iWorkSheet + 1
    xlsApp.Workbooks(1).Sheets(iWorkSheet).Activate
    xlsApp.Workbooks(1).Sheets(iWorkSheet).Name = IIf(CDbl(rc!iMonth) = 1, "Jan", _
                                                  IIf(CDbl(rc!iMonth) = 2, "Feb", _
                                                  IIf(CDbl(rc!iMonth) = 3, "Mar", _
                                                  IIf(CDbl(rc!iMonth) = 4, "Apr", _
                                                  IIf(CDbl(rc!iMonth) = 5, "May", _
                                                  IIf(CDbl(rc!iMonth) = 6, "Jun", _
                                                  IIf(CDbl(rc!iMonth) = 7, "Jul", _
                                                  IIf(CDbl(rc!iMonth) = 8, "Aug", _
                                                  IIf(CDbl(rc!iMonth) = 9, "Sep", _
                                                  IIf(CDbl(rc!iMonth) = 10, "Oct", _
                                                  IIf(CDbl(rc!iMonth) = 11, "Nov", _
                                                  IIf(CDbl(rc!iMonth) = 12, "Dec", "")))))))))))) & _
                                                  Right(rc!iYear, 2)
    RowCount = 0
    RowCount = RowCount + 1
    j = 0
    j = j + 1
    strRange = EXCEL_RANGE(j, RowCount)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = gbl_CompanyName
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
    
    RowCount = RowCount + 1
    strRange = EXCEL_RANGE(j, RowCount)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Service Charge"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
    
    If CDbl(Month(FormatDateTime(txtDateFrom.Text, vbShortDate))) = CDbl(rc!iMonth) Then
        iDayFrom = Day(FormatDateTime(txtDateFrom.Text, vbShortDate))
        iDayTo = Day(DateSerial(Year(FormatDateTime(txtDateFrom.Text, vbShortDate)), Month(FormatDateTime(txtDateFrom.Text, vbShortDate)) + 1, 0))
    Else
        If CDbl(Month(FormatDateTime(txtDateTo.Text, vbShortDate))) = CDbl(rc!iMonth) Then
            iDayFrom = 1
            iDayTo = Day(FormatDateTime(txtDateTo.Text, vbShortDate))
        Else
            iDayFrom = 1
            iDayTo = Day(DateSerial(rc!iYear, rc!iMonth + 1, 0))
        End If
    End If
    
    RowCount = RowCount + 1
    strRange = EXCEL_RANGE(j, RowCount)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = IIf(CDbl(rc!iMonth) = 1, "Jan", _
                                                                     IIf(CDbl(rc!iMonth) = 2, "Feb", _
                                                                     IIf(CDbl(rc!iMonth) = 3, "Mar", _
                                                                     IIf(CDbl(rc!iMonth) = 4, "Apr", _
                                                                     IIf(CDbl(rc!iMonth) = 5, "May", _
                                                                     IIf(CDbl(rc!iMonth) = 6, "Jun", _
                                                                     IIf(CDbl(rc!iMonth) = 7, "Jul", _
                                                                     IIf(CDbl(rc!iMonth) = 8, "Aug", _
                                                                     IIf(CDbl(rc!iMonth) = 9, "Sep", _
                                                                     IIf(CDbl(rc!iMonth) = 10, "Oct", _
                                                                     IIf(CDbl(rc!iMonth) = 11, "Nov", _
                                                                     IIf(CDbl(rc!iMonth) = 12, "Dec", "")))))))))))) & _
                                                                     " " & Format(iDayFrom, "0#") & " - " & Format(iDayTo, "0#") & ", " & rc!iYear
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
    
    RowCount = RowCount + 1
    strRange = EXCEL_RANGE(j, RowCount)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
    
    For k = iDayFrom To iDayTo
        j = j + 1
        strRange = EXCEL_RANGE(j, RowCount)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
        
        j = j + 1
        strRange = EXCEL_RANGE(j, RowCount)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "@"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = Format(k, "0#")
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3
        
        j = j + 1
        strRange = EXCEL_RANGE(j, RowCount)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
    Next k
    
    RowCount = RowCount + 1
    j = 0
    j = j + 1
    strRange = EXCEL_RANGE(j, RowCount)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Employee Name"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(j).ColumnWidth = 30
    
    For k = iDayFrom To iDayTo
        j = j + 1
        strRange = EXCEL_RANGE(j, RowCount)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Total SC"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3
        
        j = j + 1
        strRange = EXCEL_RANGE(j, RowCount)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "No of Hours"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3
        
        j = j + 1
        strRange = EXCEL_RANGE(j, RowCount)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Share"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3
    Next k
    
    j = j + 1
    strRange = EXCEL_RANGE(j, RowCount)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Total Share"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3
    
    t = "SELECT * " & _
        " FROM " & TableName & " " & _
        " WHERE (iYear = " & rc!iYear & ") " & _
        " AND (iMonth = " & rc!iMonth & ")" & _
        " ORDER BY iCategory DESC, EmployeeName"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        DoEvents
        i = i + 1
        
        If CDbl(iCat) <> CDbl(rt!iCategory) Then
            If CDbl(iCat) <> 0 Then
                RowCount = RowCount + 1
                j = 0
                j = j + 1
                strRange = EXCEL_RANGE(j, RowCount)
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            End If
            iCat = CDbl(rt!iCategory)
        End If
        sTotalShare = "="
        dTotalShare = 0
        RowCount = RowCount + 1
        j = 0
        j = j + 1
        strRange = EXCEL_RANGE(j, RowCount)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rt!EmployeeName
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
        
        If Trim(rt!sDetails1) <> "" Then
            Arr1 = Split(rt!sDetails1, "|", -1, 1)
            For k = 0 To UBound(Arr1)
                Arr2 = Split(CStr(Arr1(k)), "\", -1, 1)
                For l = 1 To UBound(Arr2)
                    j = j + 1
                    strRange = EXCEL_RANGE(j, RowCount)
                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = IIf(CDbl(Arr2(l)) = 0, "", Arr2(l))
                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
                    If CDbl(l) = 2 Then
                        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3
                    End If
                    If CDbl(l) = UBound(Arr2) Then
                        sTotalShare = sTotalShare & strRange & "+"
                        dTotalShare = dTotalShare + CDbl(Arr2(l))
                    End If
                Next l
            Next k
        End If
        
        If Trim(rt!sDetails2) <> "" Then
            Arr1 = Split(rt!sDetails2, "|", -1, 1)
            For k = 0 To UBound(Arr1)
                Arr2 = Split(CStr(Arr1(k)), "\", -1, 1)
                For l = 1 To UBound(Arr2)
                    j = j + 1
                    strRange = EXCEL_RANGE(j, RowCount)
                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = IIf(CDbl(Arr2(l)) = 0, "", Arr2(l))
                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
                    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
                    If CDbl(l) = 2 Then
                        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3
                    End If
                    If CDbl(l) = UBound(Arr2) Then
                        sTotalShare = sTotalShare & strRange & "+"
                        dTotalShare = dTotalShare + CDbl(Arr2(l))
                    End If
                Next l
            Next k
        End If
        
        u = "SELECT * " & _
            " FROM " & TableName2 & " " & _
            " WHERE (EmployeeName = '" & FORMATSQL(CStr(rt!EmployeeName)) & "')"
        If ru.State = adStateOpen Then ru.Close
        ru.Open u, ConnOmega
        If ru.RecordCount = 0 Then
            ConnOmega.Execute "INSERT INTO " & TableName2 & " " & _
                              " (EmployeeName, iCategory, AccountNo, " & IIf(CDbl(rc!iMonth) = 1, "Jan", _
                              IIf(CDbl(rc!iMonth) = 2, "Feb", _
                              IIf(CDbl(rc!iMonth) = 3, "Mar", _
                              IIf(CDbl(rc!iMonth) = 4, "Apr", _
                              IIf(CDbl(rc!iMonth) = 5, "May", _
                              IIf(CDbl(rc!iMonth) = 6, "Jun", _
                              IIf(CDbl(rc!iMonth) = 7, "Jul", _
                              IIf(CDbl(rc!iMonth) = 8, "Aug", _
                              IIf(CDbl(rc!iMonth) = 9, "Sep", _
                              IIf(CDbl(rc!iMonth) = 10, "Oct", _
                              IIf(CDbl(rc!iMonth) = 11, "Nov", _
                              IIf(CDbl(rc!iMonth) = 12, "Dec", "")))))))))))) & _
                              "_" & rc!iYear & ") " & _
                              " VALUES ('" & FORMATSQL(CStr(rt!EmployeeName)) & "', " & _
                              " " & rt!iCategory & ", '" & Replace(Trim(rt!AccountNo), "-", "") & "', " & _
                              " " & CDbl(dTotalShare) & ")"
        Else
            ConnOmega.Execute "UPDATE " & TableName2 & " " & _
                              " SET " & IIf(CDbl(rc!iMonth) = 1, "Jan", _
                              IIf(CDbl(rc!iMonth) = 2, "Feb", _
                              IIf(CDbl(rc!iMonth) = 3, "Mar", _
                              IIf(CDbl(rc!iMonth) = 4, "Apr", _
                              IIf(CDbl(rc!iMonth) = 5, "May", _
                              IIf(CDbl(rc!iMonth) = 6, "Jun", _
                              IIf(CDbl(rc!iMonth) = 7, "Jul", _
                              IIf(CDbl(rc!iMonth) = 8, "Aug", _
                              IIf(CDbl(rc!iMonth) = 9, "Sep", _
                              IIf(CDbl(rc!iMonth) = 10, "Oct", _
                              IIf(CDbl(rc!iMonth) = 11, "Nov", _
                              IIf(CDbl(rc!iMonth) = 12, "Dec", "")))))))))))) & _
                              "_" & rc!iYear & " = " & CDbl(dTotalShare) & ", " & _
                              " iCategory = " & rt!iCategory & ", " & _
                              " AccountNo = '" & Replace(Trim(rt!AccountNo), "-", "") & "' " & _
                              " WHERE (EmployeeName = '" & FORMATSQL(CStr(rt!EmployeeName)) & "')"
        End If
        ru.Close
        
        j = j + 1
        strRange = EXCEL_RANGE(j, RowCount)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "#,##0.00"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = Mid(sTotalShare, 1, Len(sTotalShare) - 1)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
        
        UpdateProgress_Caption "Exporting Report to Excel", picProgressBar, i / iTotalRecord
        rt.MoveNext
    Wend
    rt.Close
    
    rc.MoveNext
Wend
rc.Close

iWorkSheet = iWorkSheet + 1
xlsApp.Workbooks(1).Sheets(iWorkSheet).Activate
xlsApp.Workbooks(1).Sheets(iWorkSheet).Name = "Summ"
iCat = 0
RowCount = 0
RowCount = RowCount + 1
j = 0
j = j + 1
strRange = EXCEL_RANGE(j, RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = gbl_CompanyName
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCount = RowCount + 1
strRange = EXCEL_RANGE(j, RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Service Charge"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCount = RowCount + 1
strRange = EXCEL_RANGE(j, RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "From " & Format(txtDateFrom.Text, "mmmm dd, yyyy") & " to " & Format(txtDateTo.Text, "mmmm dd, yyyy")
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCount = RowCount + 1
strRange = EXCEL_RANGE(j, RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCount = RowCount + 1
strRange = EXCEL_RANGE(j, RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Employee Name"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(j).ColumnWidth = 30

j = j + 1
strRange = EXCEL_RANGE(j, RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Account Number"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Columns(j).ColumnWidth = 15
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3

s = "SELECT * " & _
    " FROM " & TableName2 & " " & _
    " ORDER BY iCategory DESC, EmployeeName"
If rc.State = adStateOpen Then rc.Close
rc.Open s, ConnOmega
For k = 0 To rc.Fields.Count
    If CDbl(k) > 4 Then
        sFieldNameArr = Split(CStr(rc.Fields(k - 1).Name), "_", -1, 1)
        sFieldName = sFieldNameArr(0) & " " & sFieldNameArr(1)
        j = j + 1
        strRange = EXCEL_RANGE(j, RowCount)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "@"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = sFieldName
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
    End If
Next k

j = j + 1
strRange = EXCEL_RANGE(j, RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Total SC"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4

While Not rc.EOF
    DoEvents
    i = i + 1
    If CDbl(iCat) <> CDbl(rc!iCategory) Then
        If CDbl(iCat) <> 0 Then
            RowCount = RowCount + 1
            j = 0
            j = j + 1
            strRange = EXCEL_RANGE(j, RowCount)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
        End If
        iCat = CDbl(rc!iCategory)
    End If
    sTotalShare = "="
    RowCount = RowCount + 1
    j = 0
    j = j + 1
    strRange = EXCEL_RANGE(j, RowCount)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rc!EmployeeName
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
    
    j = j + 1
    strRange = EXCEL_RANGE(j, RowCount)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "0000-0000-00"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rc!AccountNo
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3
    
    For k = 0 To rc.Fields.Count
        If CDbl(k) > 4 Then
            j = j + 1
            strRange = EXCEL_RANGE(j, RowCount)
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "#,##0.00"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rc.Fields(k - 1).Value
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            sTotalShare = sTotalShare & strRange & "+"
        End If
    Next k
    
    j = j + 1
    strRange = EXCEL_RANGE(j, RowCount)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "#,##0.00"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = Mid(sTotalShare, 1, Len(sTotalShare) - 1)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
    
    UpdateProgress_Caption "Exporting Report to Excel", picProgressBar, i / iTotalRecord
    rc.MoveNext
Wend
rc.Close

iWorkSheet = iWorkSheet + 1
xlsApp.Workbooks(1).Sheets(iWorkSheet).Activate
xlsApp.Workbooks(1).Sheets(iWorkSheet).Name = "Daily"
iCat = 0
RowCount = 0
RowCount = RowCount + 1
j = 0
j = j + 1
strRange = EXCEL_RANGE(j, RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = gbl_CompanyName
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCount = RowCount + 1
strRange = EXCEL_RANGE(j, RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Service Charge"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCount = RowCount + 1
strRange = EXCEL_RANGE(j, RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "From " & Format(txtDateFrom.Text, "mmmm dd, yyyy") & " to " & Format(txtDateTo.Text, "mmmm dd, yyyy")
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCount = RowCount + 1
strRange = EXCEL_RANGE(j, RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCount = RowCount + 1
j = 0
j = j + 1
strRange = EXCEL_RANGE(j, RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3

j = j + 1
strRange = EXCEL_RANGE(j, RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Total"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3

s = "SELECT TOP 1 PK " & _
    " From tbl_Service_Charge_Setup " & _
    " WHERE (EffectDate <= '" & FormatDateTime(txtDateFrom.Text, vbShortDate) & "')"
If rc.State = adStateOpen Then rc.Close
rc.Open s, ConnOmega
If rc.RecordCount > 0 Then
    t = "SELECT ForCompany, Rate " & _
        " From tbl_Service_Charge_SetupDetail " & _
        " Where (MasterKey = " & rc!PK & ") " & _
        " ORDER BY ForCompany"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        j = j + 1
        strRange = EXCEL_RANGE(j, RowCount)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rt!ForCompany
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3
        rt.MoveNext
    Wend
    rt.Close
End If
rc.Close

j = j + 1
strRange = EXCEL_RANGE(j, RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Net"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3

j = j + 1
strRange = EXCEL_RANGE(j, RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Managerial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3

j = j + 1
strRange = EXCEL_RANGE(j, RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Rank In File"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3

RowCount = RowCount + 1
j = 0
j = j + 1
strRange = EXCEL_RANGE(j, RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "Date"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3

j = j + 1
strRange = EXCEL_RANGE(j, RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "SC"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3

s = "SELECT TOP 1 PK " & _
    " From tbl_Service_Charge_Setup " & _
    " WHERE (EffectDate <= '" & FormatDateTime(txtDateFrom.Text, vbShortDate) & "')"
If rc.State = adStateOpen Then rc.Close
rc.Open s, ConnOmega
If rc.RecordCount > 0 Then
    t = "SELECT ForCompany, Rate " & _
        " From tbl_Service_Charge_SetupDetail " & _
        " Where (MasterKey = " & rc!PK & ") " & _
        " ORDER BY ForCompany"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        j = j + 1
        strRange = EXCEL_RANGE(j, RowCount)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = CStr(rt!Rate) & "%"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3
        rt.MoveNext
    Wend
    rt.Close
End If
rc.Close

j = j + 1
strRange = EXCEL_RANGE(j, RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "SC"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3

j = j + 1
strRange = EXCEL_RANGE(j, RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "15%"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3

j = j + 1
strRange = EXCEL_RANGE(j, RowCount)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "85%"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 3

iMonth = 0
s = "SELECT iDate, TotalSC, Managerial, RanInFile " & _
    " From tbl_Service_Charge_Daily " & _
    " WHERE (iDate >= '" & FormatDateTime(txtDateFrom.Text, vbShortDate) & "') " & _
    " AND (iDate <= '" & FormatDateTime(txtDateTo.Text, vbShortDate) & "') " & _
    " AND (SummKey = " & StatusBar1.Panels(1).Text & ") " & _
    " ORDER BY iDate"
If rc.State = adStateOpen Then rc.Close
rc.Open s, ConnOmega
While Not rc.EOF
    DoEvents
    i = i + 1
    
    If CDbl(iMonth) <> CDbl(Month(rc!iDate)) Then
        RowCount = RowCount + 1
        j = 0
        j = j + 1
        strRange = EXCEL_RANGE(j, RowCount)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
        iMonth = Month(rc!iDate)
    End If
    
    RowCount = RowCount + 1
    dComputeRange = "="
    j = 0
    j = j + 1
    strRange = EXCEL_RANGE(j, RowCount)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "dd-mmm-yy"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rc!iDate
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4

    j = j + 1
    strRange = EXCEL_RANGE(j, RowCount)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "#,##0.00"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rc!TotalSC
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
    
    dComputeRange = dComputeRange & strRange & "-"
    t = "SELECT Amount " & _
        " From tbl_Service_Charge_ForCompany " & _
        " WHERE (iDate = '" & FormatDateTime(rc!iDate, vbShortDate) & "') " & _
        " AND (SummKey = " & StatusBar1.Panels(1).Text & ") " & _
        " ORDER BY ForCompany"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        j = j + 1
        strRange = EXCEL_RANGE(j, RowCount)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "#,##0.00"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rt!Amount
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
        dComputeRange = dComputeRange & strRange & "-"
        rt.MoveNext
    Wend
    rt.Close
    
    j = j + 1
    strRange = EXCEL_RANGE(j, RowCount)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "#,##0.00"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = Mid(dComputeRange, 1, Len(dComputeRange) - 1)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
    
    j = j + 1
    strRange = EXCEL_RANGE(j, RowCount)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "#,##0.00"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rc!Managerial
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
    
    j = j + 1
    strRange = EXCEL_RANGE(j, RowCount)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "#,##0.00"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rc!RanInFile
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 8
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).HorizontalAlignment = 4
    
    UpdateProgress_Caption "Exporting Report to Excel", picProgressBar, i / iTotalRecord
    rc.MoveNext
Wend
rc.Close

SAVING:
On Error GoTo ErrSaving:
If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

xlsApp.Visible = True

picProgress.Visible = False
Screen.MousePointer = vbDefault

Exit Sub
ErrSaving:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File is currently open!           ", vbCritical, "Error..."
GoTo SAVING:

Exit Sub
PG:
Screen.MousePointer = vbDefault
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub

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
    Case vbKeyF8:       PRESS_F8
    Case vbKeyF9:       PRESS_F9
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "ServiceChargeSumm", "ServChrgeSumm", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "ServiceChargeSumm", "ServChrgeSumm", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "ServiceChargeSumm", "ServChrgeSumm", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "ServiceChargeSumm", "ServChrgeSumm", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
CLEARTEXT
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "ServiceChargeSumm", "ServChrgeSumm", ""), "is_LOAD"
If Trim(txtDateFrom.Text) = "" Then BROWSER GetSetting(App.EXEName, "ServiceChargeSumm", "ServChrgeSumm", ""), "is_HOME"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picGenerateServiceCharge.Visible = True Then Cancel = -1
If picGenProgress.Visible = True Then Cancel = -1
If picProgress.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
        Case "Refresh"
            'ToDo: Add 'Refresh' button code.
            MsgBox "Add 'Refresh' button code."
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   BROWSER GetSetting(App.EXEName, "ServiceChargeSumm", "ServChrgeSumm", ""), "is_HOME"
    Case "Back":    BROWSER GetSetting(App.EXEName, "ServiceChargeSumm", "ServChrgeSumm", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "ServiceChargeSumm", "ServChrgeSumm", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "ServiceChargeSumm", "ServChrgeSumm", ""), "is_END"
    Case "Find":
    Case "Print":   PRESS_F9
    Case "Post":    PRESS_F8
    Case "Close":   PRESS_ESCAPE
    Case Else: Exit Sub
End Select
End Sub

Private Sub txtDateFromG_GotFocus()
HTEXT txtDateFromG
End Sub

Private Sub txtDateFromG_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtDateToG.SetFocus
End Sub

Private Sub txtDateFromG_LostFocus()
If IsDate(txtDateFromG.Text) = True Then
    txtDateFromG.Text = Format(FormatDateTime(txtDateFromG.Text, vbShortDate), "mm/dd/yyyy")
    s = "SELECT tbl_Service_Charge_CutOff.* " & _
        " From tbl_Service_Charge_CutOff " & _
        " WHERE (MonthFrom = " & Month(FormatDateTime(txtDateFromG.Text, vbShortDate)) & ") " & _
        " AND (DayFrom = " & Day(FormatDateTime(txtDateFromG.Text, vbShortDate)) & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        txtDateToG.Text = Format(DateSerial(Year(FormatDateTime(txtDateFromG.Text, vbShortDate)) + rs!YearTo, rs!MonthTo, rs!DayTo), "mm/dd/yyyy")
    Else
        txtDateToG.Text = ""
    End If
    rs.Close
End If
End Sub

Private Sub txtDateToG_GotFocus()
HTEXT txtDateToG
End Sub

Private Sub txtDateToG_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKGenerate_Click
End Sub
