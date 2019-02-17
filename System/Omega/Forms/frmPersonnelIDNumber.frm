VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPersonnelIDNumber 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9105
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPersonnelIDNumber.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   37
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   38
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
         MouseIcon       =   "frmPersonnelIDNumber.frx":0CCA
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
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   960
      ScaleHeight     =   3015
      ScaleWidth      =   6855
      TabIndex        =   1
      Top             =   1200
      Width           =   6855
      Begin VB.TextBox txtATMNumber 
         Height          =   315
         Left            =   1680
         TabIndex        =   28
         Top             =   2640
         Width           =   5175
      End
      Begin MSMask.MaskEdBox txtID 
         Height          =   315
         Left            =   3120
         TabIndex        =   19
         Top             =   480
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
      Begin VB.TextBox txtDateHired 
         Height          =   315
         Left            =   5520
         TabIndex        =   20
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtRemarks 
         Height          =   315
         Left            =   1680
         TabIndex        =   17
         Top             =   2280
         Width           =   5175
      End
      Begin VB.TextBox txtDivision 
         Height          =   315
         Left            =   1680
         TabIndex        =   15
         Top             =   840
         Width           =   5175
      End
      Begin VB.TextBox txtPost 
         Height          =   315
         Left            =   2265
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1560
         Width           =   4590
      End
      Begin VB.TextBox txtPostCode 
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1560
         Width           =   555
      End
      Begin VB.TextBox txtStatus 
         Height          =   315
         Left            =   2265
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1920
         Width           =   4590
      End
      Begin VB.TextBox txtStatusCode 
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1920
         Width           =   555
      End
      Begin VB.TextBox txtDept 
         Height          =   315
         Left            =   2270
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1200
         Width           =   4590
      End
      Begin VB.TextBox txtDeptCode 
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   555
      End
      Begin VB.TextBox txtFullName 
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   120
         Width           =   5175
      End
      Begin VB.TextBox txtIDNumber 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Text            =   "000000-0000"
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "ATM Number"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Hired"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4560
         TabIndex        =   21
         Top             =   510
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Employment Status"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Number"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   510
         Width           =   1095
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   4680
      Width           =   9105
      _ExtentX        =   16060
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
   Begin RPVGCC.b8Container picSearch1 
      Height          =   4095
      Left            =   2280
      TabIndex        =   30
      Top             =   360
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   7223
      BackColor       =   15396057
      Begin VB.ListBox lstResult2 
         Height          =   840
         Left            =   120
         TabIndex        =   36
         Top             =   2520
         Width           =   3735
      End
      Begin VB.CommandButton cmdOK1 
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
         Picture         =   "frmPersonnelIDNumber.frx":0FE4
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3480
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancel1 
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
         Picture         =   "frmPersonnelIDNumber.frx":1656
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3480
         Width           =   1560
      End
      Begin VB.TextBox txtSearch1 
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   3735
      End
      Begin VB.ListBox lstResult1 
         Height          =   1620
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Width           =   3735
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   40
         TabIndex        =   35
         Top             =   40
         Width           =   3890
         _ExtentX        =   6853
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
         Icon            =   "frmPersonnelIDNumber.frx":1DB2
         ShadowVisible   =   0   'False
      End
   End
   Begin RPVGCC.b8Container picSearch 
      Height          =   3735
      Left            =   2280
      TabIndex        =   22
      Top             =   480
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   6588
      BackColor       =   15396057
      Begin VB.ListBox lstResult 
         Height          =   2205
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   3735
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
         Picture         =   "frmPersonnelIDNumber.frx":234C
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3120
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
         Picture         =   "frmPersonnelIDNumber.frx":2AA8
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3120
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   40
         TabIndex        =   27
         Top             =   40
         Width           =   3890
         _ExtentX        =   6853
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
         Icon            =   "frmPersonnelIDNumber.frx":311A
         ShadowVisible   =   0   'False
      End
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
            Picture         =   "frmPersonnelIDNumber.frx":36B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelIDNumber.frx":438E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelIDNumber.frx":5068
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelIDNumber.frx":5D42
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelIDNumber.frx":6A1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelIDNumber.frx":76F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelIDNumber.frx":83D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelIDNumber.frx":90AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelIDNumber.frx":9D84
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelIDNumber.frx":A65E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelIDNumber.frx":B338
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelIDNumber.frx":C012
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelIDNumber.frx":CCEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelIDNumber.frx":D9C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelIDNumber.frx":E6A0
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPersonnelIDNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ProfileKey As Double

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Public SearchType As Long
Dim tmp As Long

Dim Arr

Private Function BROWSER(IDNum, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If IDNum <> "" Then
            s = "sp_Personnel_IDNumber_BrowseV2('" & IDNum & "', 0)"
        Else
            s = "sp_Personnel_IDNumber_BrowseV2('" & IDNum & "', 1)"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "sp_Personnel_IDNumber_BrowseV2('" & IDNum & "', 1)"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "sp_Personnel_IDNumber_BrowseV2('" & IDNum & "', 2)"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "sp_Personnel_IDNumber_BrowseV2('" & IDNum & "', 3)"
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
        s = "sp_Personnel_IDNumber_BrowseV2('" & IDNum & "', 4)"
    Case Else: Exit Function
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    ProfileKey = rs!ProfileKey
    txtFullName.Text = rs!EmployeeName
    txtIDNumber.Text = rs!IDNumber
    txtDateHired.Text = Format(rs!DateHired, "mm/dd/yyyy")
    txtDivision.Text = IIf(IsNull(rs!Division), "", rs!Division) 'IIf(IsNull(rs!Division), "", IIf(rs!Division = 1, "CLUB HOUSE", IIf(rs!Division = 2, "MAINTENANCE", "")))
    txtDeptCode.Text = IIf(IsNull(rs!DeptCode), "", rs!DeptCode)
    txtDept.Text = IIf(IsNull(rs!DeptName), "", rs!DeptName)
    txtPostCode.Text = IIf(IsNull(rs!PostCode), "", rs!PostCode)
    txtPost.Text = IIf(IsNull(rs!PostName), "", rs!PostName)
    txtStatusCode.Text = IIf(IsNull(rs!StatusCode), "", rs!StatusCode)
    txtStatus.Text = IIf(IsNull(rs!StatusName), "", rs!StatusName)
    txtRemarks.Text = ""
    txtATMNumber.Text = rs!AccountNumber
    StatusBar1.Panels(1).Text = rs!PK
    StatusBar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", "Last Modified by : " & rs!LastModified)
    
    SaveSetting App.EXEName, "IDNumber", "ISNum", rs!IDNumber
End If
rs.Close
End Function

Private Function PRESS_INSERT()
If picSearch.Visible = True Then Exit Function
If picSearch1.Visible = True Then Exit Function
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If AccessRights("Personnel ID Number", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
SearchType = 0
picSearch.ZOrder 0
picMain.Enabled = False
picToolbar.Enabled = False
txtSearch.Text = ""
picSearch.Visible = True
txtSearch.SetFocus
End Function

Private Function PRESS_F2()
If picSearch.Visible = True Then Exit Function
If picSearch1.Visible = True Then Exit Function
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If AccessRights("Personnel ID Number", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
txtFullName.SetFocus
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_EDITTING
txtIDNumber.SetFocus
End Function

Private Function PRESS_DELETE()
If picSearch.Visible = True Then Exit Function
If picSearch1.Visible = True Then Exit Function
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
If AccessRights("Personnel ID Number", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Function
End If
If MsgBox("ARE YOU SURE IN DELETING THIS ID?                ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Function
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_Personnel_IDNumber WHERE (PK = " & StatusBar1.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "IDNumber", "ISNum", ""), "is_PAGEDOWN"
If Trim(txtIDNumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "IDNumber", "ISNum", ""), "is_HOME"
Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Function
End Function

Private Function PRESS_F5()
If picSearch.Visible = True Then Exit Function
If txtIDNumber.Text = "" Then MsgBox "Please Supply ID Number!               ", vbCritical, "Error...": txtIDNumber.SetFocus: HTEXT txtIDNumber: Exit Function
If ProfileKey = 0 Then MsgBox "Please Supply Employee!                  ", vbCritical, "Error...": Exit Function
If IsDate(txtDateHired.Text) = False Then MsgBox "Please Supply a Valid Date!                 ", vbCritical, "Error...": txtDateHired.SetFocus: HTEXT txtDateHired: Exit Function
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    ConnOmega.Execute "INSERT INTO tbl_Personnel_IDNumber " & _
                      " (ProfileKey, IDNumber, DateHired, AccountNumber, LastModified) " & _
                      " VALUES (" & ProfileKey & ", '" & Trim(txtIDNumber.Text) & "', " & _
                      " '" & FormatDateTime(txtDateHired.Text, vbShortDate) & "', " & _
                      " '" & Replace(Trim(txtATMNumber.Text), "-", "") & "', '" & CStr(Now) & " - " & gbl_CompleteName & "')"
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER Trim(txtIDNumber.Text), "is_LOAD"
End If
If TRANSACTIONTYPE = is_EDITTING Then
    ConnOmega.Execute "UPDATE tbl_Personnel_IDNumber " & _
                      " SET IDNumber = '" & Trim(txtIDNumber.Text) & "', " & _
                      " DateHired = '" & FormatDateTime(txtDateHired.Text, vbShortDate) & "', " & _
                      " AccountNumber = '" & Replace(Trim(txtATMNumber.Text), "-", "") & "', " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & StatusBar1.Panels(1).Text & ")"
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER Trim(txtIDNumber.Text), "is_LOAD"
End If
Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Function
End Function

Private Function PRESS_F6()
If picSearch.Visible = True Then Exit Function
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function

PopupMenu MainFormPopupF.mnuIDSearch, , Toolbar1.Buttons(15).Left, Toolbar1.Buttons(15).Top + Toolbar1.Buttons(15).Height

'SearchType = 1
'picSearch.ZOrder 0
'picMain.Enabled = False
'picToolbar.Enabled = False
'txtSearch.Text = ""
'picSearch.Visible = True
'txtSearch.SetFocus

End Function

Private Function PRESS_F9()
If picSearch.Visible = True Then Exit Function
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
End Function

Private Function PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSearch.Visible = True Then cmdCancel_Click: Exit Function
    If picSearch1.Visible = True Then cmdCancel1_Click: Exit Function
    Unload Me
Else
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "IDNumber", "ISNum", ""), "is_LOAD"
End If
End Function

Private Function CLEARTEXT()
ProfileKey = 0
txtFullName.Text = ""
txtIDNumber.Text = ""
txtDateHired.Text = ""
txtDivision.Text = ""
txtDeptCode.Text = ""
txtDept.Text = ""
txtPostCode.Text = ""
txtPost.Text = ""
txtStatusCode.Text = ""
txtStatus.Text = ""
txtRemarks.Text = ""
txtATMNumber.Text = ""
txtID.Visible = False
StatusBar1.Panels(1).Text = ""
StatusBar1.Panels(2).Text = ""
End Function

Private Function LOCKTEXT(bln As Boolean)
If bln Then
    txtFullName.Locked = True
    txtIDNumber.Locked = True
    txtDateHired.Locked = True
    txtDivision.Locked = True
    txtDeptCode.Locked = True
    txtDept.Locked = True
    txtPostCode.Locked = True
    txtPost.Locked = True
    txtStatusCode.Locked = True
    txtStatus.Locked = True
    txtRemarks.Locked = True
    txtATMNumber.Locked = True
Else
    txtFullName.Locked = True
    txtIDNumber.Locked = True
    txtDateHired.Locked = False
    txtDivision.Locked = True
    txtDeptCode.Locked = True
    txtDept.Locked = True
    txtPostCode.Locked = True
    txtPost.Locked = True
    txtStatusCode.Locked = True
    txtStatus.Locked = True
    txtRemarks.Locked = True
    txtATMNumber.Locked = False
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
cmdCancel1_Click
End Sub

Private Sub b8TitleBar2_CLoseClick()
cmdCancel_Click
End Sub

Private Sub cmdCancel_Click()
picMain.Enabled = True
picToolbar.Enabled = True
picSearch.Visible = False
End Sub

Private Sub cmdCancel1_Click()
picToolbar.Enabled = True
picMain.Enabled = True
picSearch1.Visible = False
End Sub

Private Sub cmdOK_Click()

If lstResult.ListIndex = -1 Then Exit Sub
Select Case SearchType
    Case 0
        CLEARTEXT
        ProfileKey = lstResult.ItemData(lstResult.ListIndex)
        txtFullName.Text = lstResult.List(lstResult.ListIndex)
        cmdCancel_Click
        LOCKTEXT False
        TOOLBARFUNC 2
        TRANSACTIONTYPE = is_ADDING
        txtIDNumber.SetFocus
    Case 1
        Arr = Split(lstResult.List(lstResult.ListIndex), " - ", -1, 1)
        BROWSER Arr(0), "is_LOAD"
        cmdCancel_Click
    Case Else: Exit Sub
End Select
End Sub

Private Sub cmdOK1_Click()
If lstResult2.ListIndex = -1 Then Exit Sub
BROWSER lstResult2.List(lstResult2.ListIndex), "is_LOAD"
cmdCancel1_Click
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
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "IDNumber", "ISNum", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "IDNumber", "ISNum", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "IDNumber", "ISNum", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "IDNumber", "ISNum", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Top = (MainForm.Height - Me.Height) / 4
Me.Left = (MainForm.Width - Me.Width) / 5
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "IDNumber", "ISNum", ""), "is_LOAD"
If Trim(txtIDNumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "IDNumber", "ISNum", ""), "is_HOME"

tmp = SetWindowLong(txtSearch1.hwnd, GWL_STYLE, GetWindowLong(txtSearch1.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picSearch.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstResult_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub lstResult1_Click()
If lstResult1.ListIndex = -1 Then lstResult2.Clear: Exit Sub
lstResult2.Clear
s = "SELECT IDNumber " & _
    " From tbl_Personnel_IDNumber " & _
    " Where (ProfileKey = " & lstResult1.ItemData(lstResult1.ListIndex) & ") " & _
    " ORDER BY IDNumber"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResult2.AddItem rs!IDNumber
    rs.MoveNext
Wend
rs.Close
If lstResult2.ListCount Then lstResult2.ListIndex = 0
End Sub

Private Sub lstResult1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResult2.SetFocus
End Sub

Private Sub lstResult2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOK1_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "IDNumber", "ISNum", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "IDNumber", "ISNum", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "IDNumber", "ISNum", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "IDNumber", "ISNum", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":   PRESS_F9
    Case "Close":   PRESS_ESCAPE
    Case Else: Exit Sub
End Select
End Sub

Private Sub txtATMNumber_GotFocus()
HTEXT txtATMNumber
End Sub

Private Sub txtATMNumber_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtATMNumber_LostFocus()
txtATMNumber.Text = Replace(Trim(txtATMNumber.Text), "-", "")
End Sub

Private Sub txtDateHired_GotFocus()
HTEXT txtDateHired
End Sub

Private Sub txtDateHired_LostFocus()
If IsDate(txtDateHired.Text) = True Then
    txtDateHired.Text = Format(FormatDateTime(txtDateHired.Text, vbShortDate), "mm/dd/yyyy")
End If
End Sub

Private Sub txtID_Change()
Arr = Split(txtID.Text, "-", -1, 1)
txtIDNumber.Text = Format(IIf(Trim(Arr(0)) = "", 0, Arr(0)), "00000#") & "-" & Format(IIf(Trim(Arr(1)) = "", 0, Arr(1)), "000#")
End Sub

Private Sub txtID_GotFocus()
HTEXT txtID
End Sub

Private Sub txtID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtID.Visible = False: txtDateHired.SetFocus
End Sub

Private Sub txtID_LostFocus()
txtID.Visible = False
End Sub

Private Sub txtIDNumber_GotFocus()
HTEXT txtIDNumber
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If Trim(txtIDNumber.Text) = "" Then
        txtID.Text = "      -    "
    Else
        txtID.Text = Trim(txtIDNumber.Text)
    End If
    txtID.ZOrder 0
    txtID.Move txtIDNumber.Left, txtIDNumber.Top, txtIDNumber.Width, txtIDNumber.Height
    txtID.Visible = True
    txtID.SetFocus
End If
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstResult.Clear: Exit Sub
lstResult.Clear
Select Case SearchType
    Case 0  ' add
        s = "SELECT PK, LastName + ',  ' + FirstName + '  ' + MiddleName AS EmployeeName " & _
            " From tbl_Personnel_Information " & _
            " WHERE (LastName LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
    Case 1  ' id number
        s = "SELECT tbl_Personnel_IDNumber.PK, " & _
            " tbl_Personnel_IDNumber.IDNumber + ' - ' + tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName " & _
            " + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName " & _
            " FROM tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
            " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
            " WHERE (tbl_Personnel_IDNumber.IDNumber LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
            " ORDER BY tbl_Personnel_IDNumber.IDNumber, tbl_Personnel_Information.LastName, " & _
            " tbl_Personnel_Information.FirstName, tbl_Personnel_Information.MiddleName"
    Case Else: Exit Sub
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

Private Sub txtSearch1_Change()
If Trim(txtSearch1.Text) = "" Then lstResult1.Clear: lstResult2.Clear: Exit Sub
lstResult1.Clear: lstResult2.Clear
's = "SELECT tbl_Personnel_IDNumber.ProfileKey, " & _
    " tbl_Personnel_Information.LastName + ',  ' + tbl_Personnel_Information.FirstName " & _
    " + '  ' + tbl_Personnel_Information.MiddleName AS EmployeeName " & _
    " FROM tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
    " WHERE (tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearch1.Text)) & "%') " & _
    " ORDER BY tbl_Personnel_IDNumber.IDNumber, tbl_Personnel_Information.LastName, " & _
    " tbl_Personnel_Information.FirstName, tbl_Personnel_Information.MiddleName"
s = "SELECT tbl_Personnel_Information.LastName, " & _
    " tbl_Personnel_Information.FirstName, " & _
    " tbl_Personnel_Information.MiddleName, " & _
    " tbl_Personnel_IDNumber.ProfileKey " & _
    " FROM tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
    " tbl_Personnel_Information ON tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK " & _
    " GROUP BY tbl_Personnel_Information.LastName, tbl_Personnel_Information.FirstName, tbl_Personnel_Information.MiddleName, " & _
    " tbl_Personnel_IDNumber.ProfileKey " & _
    " HAVING (tbl_Personnel_Information.LastName LIKE '" & FORMATSQL(Trim(txtSearch1.Text)) & "%') " & _
    " ORDER BY tbl_Personnel_Information.LastName, tbl_Personnel_Information.FirstName, tbl_Personnel_Information.MiddleName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResult1.AddItem rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName 'rs!EmployeeName
    lstResult1.ItemData(lstResult1.NewIndex) = rs!ProfileKey
    rs.MoveNext
Wend
rs.Close
If lstResult1.ListCount Then lstResult1.ListIndex = 0
End Sub

Private Sub txtSearch1_GotFocus()
HTEXT txtSearch1
End Sub

Private Sub txtSearch1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then lstResult1.SetFocus
End Sub
