VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOperationLockerRoom 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11370
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOperationLockerRoom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   2280
      ScaleHeight     =   2295
      ScaleWidth      =   6735
      TabIndex        =   1
      Top             =   1320
      Width           =   6735
      Begin VB.PictureBox picSetFocus 
         BackColor       =   &H00C6B8A4&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3720
         ScaleHeight     =   255
         ScaleWidth      =   735
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C6B8A4&
         Caption         =   "Fairway Towel"
         Height          =   1575
         Left            =   4080
         TabIndex        =   14
         Top             =   720
         Width           =   2655
         Begin VB.TextBox txtBalance 
            Height          =   315
            Left            =   1080
            MaxLength       =   100
            TabIndex        =   20
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtReturned 
            Height          =   315
            Left            =   1080
            MaxLength       =   100
            TabIndex        =   18
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtBorrowed 
            Height          =   315
            Left            =   1080
            MaxLength       =   100
            TabIndex        =   16
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Balance"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1080
            Width           =   2055
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Returned"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Borrowed"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.PictureBox picReturn 
         BackColor       =   &H00C6B8A4&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1800
         ScaleHeight     =   255
         ScaleWidth      =   1815
         TabIndex        =   12
         Top             =   1920
         Width           =   1815
         Begin VB.CheckBox chkReturn 
            BackColor       =   &H00C6B8A4&
            Caption         =   "Returned"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.TextBox txtLockerKeyNo 
         Height          =   315
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   10
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   6
         Top             =   0
         Width           =   1455
      End
      Begin VB.TextBox txtCtrlNo 
         Height          =   315
         Left            =   5040
         MaxLength       =   100
         TabIndex        =   5
         Top             =   0
         Width           =   1695
      End
      Begin VB.TextBox txtPassport 
         Height          =   315
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdSelectPassport 
         Caption         =   "..."
         Height          =   315
         Left            =   3390
         MouseIcon       =   "frmOperationLockerRoom.frx":08CA
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   360
         Width           =   300
      End
      Begin VB.TextBox txtPassportKey 
         Height          =   315
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Locker Key #"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   1590
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   30
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl #"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4200
         TabIndex        =   8
         Top             =   30
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Passport No"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   390
         Width           =   1095
      End
   End
   Begin RPVGCC.b8Container picAddLockerKey 
      Height          =   3375
      Left            =   3480
      TabIndex        =   28
      Top             =   360
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5953
      BackColor       =   15396057
      Begin VB.CommandButton cmdOKLocker 
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
         Picture         =   "frmOperationLockerRoom.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2745
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelLocker 
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
         Picture         =   "frmOperationLockerRoom.frx":1246
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2745
         Width           =   1560
      End
      Begin VB.TextBox txtSearchLocker 
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   480
         Width           =   4215
      End
      Begin VB.ListBox lstResultLocker 
         Height          =   1815
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   4215
      End
      Begin RPVGCC.b8TitleBar b8TitleBar5 
         Height          =   345
         Left            =   40
         TabIndex        =   33
         Top             =   40
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   609
         Caption         =   "Search Passport"
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
         Icon            =   "frmOperationLockerRoom.frx":19A2
         ShadowVisible   =   0   'False
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15600
      TabIndex        =   22
      Top             =   0
      Width           =   15600
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   23
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
         MouseIcon       =   "frmOperationLockerRoom.frx":1F3C
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   9900
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   24
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmOperationLockerRoom.frx":2256
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   2760
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
            Picture         =   "frmOperationLockerRoom.frx":2969
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationLockerRoom.frx":3643
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationLockerRoom.frx":431D
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationLockerRoom.frx":4FF7
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationLockerRoom.frx":5CD1
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationLockerRoom.frx":69AB
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationLockerRoom.frx":7685
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationLockerRoom.frx":835F
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationLockerRoom.frx":9039
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationLockerRoom.frx":9913
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationLockerRoom.frx":A5ED
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationLockerRoom.frx":B2C7
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationLockerRoom.frx":BFA1
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationLockerRoom.frx":CC7B
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationLockerRoom.frx":D955
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   4005
      Width           =   11370
      _ExtentX        =   20055
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
   Begin RPVGCC.b8Container picSearchPassport 
      Height          =   1815
      Left            =   4200
      TabIndex        =   25
      Top             =   2040
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3201
      BackColor       =   15396057
      Begin VB.ListBox lstPassportAdd 
         Height          =   1230
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox txtPassportAdd 
         Height          =   315
         Left            =   120
         MaxLength       =   100
         TabIndex        =   26
         Top             =   120
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmOperationLockerRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim iPassport, sCtrl, Arr

Private Sub BROWSER(Ctrl, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If Ctrl <> "" Then
            s = "SELECT TOP 1 tbl_Operation_LockerRoom.* " & _
                " FROM tbl_Operation_LockerRoom " & _
                " WHERE (CtrlNo = '" & Ctrl & "') " & _
                " ORDER BY CtrlNo"
        Else
            s = "SELECT TOP 1 tbl_Operation_LockerRoom.* " & _
                " FROM tbl_Operation_LockerRoom " & _
                " ORDER BY CtrlNo"
        End If
    Case "is_HOME"
        If picAddLockerKey.Visible = True Then Exit Sub
        If picSearchPassport.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Operation_LockerRoom.* " & _
            " FROM tbl_Operation_LockerRoom " & _
            " ORDER BY CtrlNo"
    Case "is_PAGEUP"
        If picAddLockerKey.Visible = True Then Exit Sub
        If picSearchPassport.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Operation_LockerRoom.* " & _
            " FROM tbl_Operation_LockerRoom " & _
            " WHERE (CtrlNo < '" & Ctrl & "') " & _
            " ORDER BY CtrlNo DESC"
    Case "is_PAGEDOWN"
        If picAddLockerKey.Visible = True Then Exit Sub
        If picSearchPassport.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Operation_LockerRoom.* " & _
            " FROM tbl_Operation_LockerRoom " & _
            " WHERE (CtrlNo > '" & Ctrl & "') " & _
            " ORDER BY CtrlNo"
    Case "is_END"
        If picAddLockerKey.Visible = True Then Exit Sub
        If picSearchPassport.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Operation_LockerRoom.* " & _
            " FROM tbl_Operation_LockerRoom " & _
            " ORDER BY CtrlNo DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtPassportKey.Text = rs!PassportKey
    txtPassport.Text = ""
    t = "SELECT tbl_Operation_Passport.* " & _
        " FROM tbl_Operation_Passport " & _
        " WHERE (PK = " & rs!PassportKey & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtPassport.Text = rt!PassportNo
    End If
    rt.Close
    txtDate.Text = Format(rs!dDate, "mm/dd/yyyy")
    txtCtrlNo.Text = rs!CtrlNo
    txtLockerKeyNo.Text = rs!LockerKeyNo
    txtBorrowed.Text = rs!FairwayTowelBorrow
    txtReturned.Text = rs!FairwayTowelReturn
    txtBalance.Text = CDbl(rs!FairwayTowelBorrow) - CDbl(rs!FairwayTowelReturn)
    chkReturn.Value = rs!LockerKeyReturn
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    imgPosted.Visible = IIf(rs!Posted = 1, True, False)
    SaveSetting App.EXEName, "OperationLockerRoom", "OperationLockerRoom", ""
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSearchPassport.Visible = True Then Exit Sub
If picAddLockerKey.Visible = True Then Exit Sub
If AccessRights("Locker Room", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
picAddLockerKey.ZOrder 0
txtSearchLocker.Text = ""
picMain.Enabled = False
picToolbar.Enabled = False
picAddLockerKey.Visible = True
txtSearchLocker.SetFocus
'CLEARTEXT
'LOCKTEXT False
'TOOLBARFUNC 2
'txtDate.Text = Format(FormatDateTime(Date, vbShortDate), "mm/dd/yyyy")
'TRANSACTIONTYPE = is_ADDING
End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSearchPassport.Visible = True Then Exit Sub
If picAddLockerKey.Visible = True Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("Locker Room", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_EDITTING
End Sub

Private Sub PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSearchPassport.Visible = True Then Exit Sub
If picAddLockerKey.Visible = True Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("Locker Room", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If MsgBox("ARE YOU SURE IN DELETING THIS TRANSACTION?                           ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
On Error GoTo PG:
ConnOmega.Execute "UPDATE tbl_Operation_Passport " & _
                  " SET LockerKeyNo = '' " & _
                  " WHERE (PK = " & RETURNTEXTVALUE(txtPassportKey) & ")"
ConnOmega.Execute "DELETE FROM tbl_Operation_LockerRoom WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "OperationLockerRoom", "OperationLockerRoom", ""), "is_PAGEDOWN"
If Trim(txtCtrlNo.Text) = "" Then BROWSER GetSetting(App.EXEName, "OperationLockerRoom", "OperationLockerRoom", ""), "is_HOME"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
If picSearchPassport.Visible = True Then Exit Sub
If picAddLockerKey.Visible = True Then Exit Sub
If IsDate(txtDate.Text) = False Then MsgBox "Please Supply a Valid Date!                      ", vbCritical, "Error...": txtDate.SetFocus: Exit Sub
If RETURNTEXTVALUE(txtPassportKey) <= 0 Then MsgBox "Please Supply Passport Number!              ", vbCritical, "Error...": txtPassport.SetFocus: Exit Sub
If Trim(txtLockerKeyNo.Text) = "" Then MsgBox "Please Supply Locker Number!                 ", vbCritical, "Error...": txtLockerKeyNo.SetFocus: Exit Sub
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    sCtrl = ""
    s = "SELECT TOP 1 CtrlNo " & _
        " FROM tbl_Operation_LockerRoom " & _
        " WHERE (Year(DDate) = " & Format(FormatDateTime(txtDate.Text, vbShortDate), "yyyy") & ") " & _
        " ORDER BY CtrlNo DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        sCtrl = Format(CDbl(rs!CtrlNo) + 1, "0000000#")
    Else
        sCtrl = Format(FormatDateTime(txtDate.Text, vbShortDate), "yyyy") & "0000"
    End If
    rs.Close
    Do
        s = "SELECT tbl_Operation_LockerRoom.* " & _
            " FROM tbl_Operation_LockerRoom " & _
            " WHERE (CtrlNo = '" & sCtrl & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If
        rs.Close
        sCtrl = Format(CDbl(sCtrl) + 1, "0000000#")
    Loop
    ConnOmega.Execute "INSERT INTO tbl_Operation_LockerRoom " & _
                      " (CtrlNo, DDate, PassportKey, LockerKeyNo, FairwayTowelBorrow, " & _
                      " FairwayTowelReturn, LastModified) " & _
                      " VALUES ('" & sCtrl & "', '" & FormatDateTime(txtDate.Text, vbShortDate) & "', " & _
                      " " & RETURNTEXTVALUE(txtPassportKey) & ", '" & Trim(txtLockerKeyNo.Text) & "', " & _
                      " " & RETURNTEXTVALUE(txtBorrowed) & ", " & RETURNTEXTVALUE(txtReturned) & ", " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
End If
If TRANSACTIONTYPE = is_EDITTING Then
    sCtrl = Trim(txtCtrlNo.Text)
    ConnOmega.Execute "UPDATE tbl_Operation_LockerRoom " & _
                      " SET DDate = '" & FormatDateTime(txtDate.Text, vbShortDate) & "', " & _
                      " PassportKey = " & RETURNTEXTVALUE(txtPassportKey) & ", " & _
                      " LockerKeyNo = '" & Trim(txtLockerKeyNo.Text) & "', " & _
                      " FairwayTowelBorrow = " & RETURNTEXTVALUE(txtBorrowed) & ", " & _
                      " FairwayTowelReturn = " & RETURNTEXTVALUE(txtReturned) & ", " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
End If

ConnOmega.Execute "UPDATE tbl_Operation_Passport " & _
                  " SET LockerKeyNo = '" & Trim(txtLockerKeyNo.Text) & "' " & _
                  " WHERE (PK = " & RETURNTEXTVALUE(txtPassportKey) & ")"

CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER sCtrl, "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSearchPassport.Visible = True Then Exit Sub
End Sub

Private Sub PRESS_F8()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSearchPassport.Visible = True Then Exit Sub
If picAddLockerKey.Visible = True Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("Locker Room", "Post") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
End Sub

Private Sub PRESS_F9()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSearchPassport.Visible = True Then Exit Sub
If picAddLockerKey.Visible = True Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    Unload Me
Else
    If picSearchPassport.Visible = True Then
        picToolbar.Enabled = True
        picMain.Enabled = True
        picSearchPassport.Visible = False
        Exit Sub
    End If
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "OperationLockerRoom", "OperationLockerRoom", ""), "is_LOAD"
    If Trim(txtCtrlNo.Text) = "" Then BROWSER GetSetting(App.EXEName, "OperationLockerRoom", "OperationLockerRoom", ""), "is_HOME"
End If
End Sub

Private Sub CLEARTEXT()
txtPassportKey.Text = ""
txtPassport.Text = ""
txtDate.Text = ""
txtCtrlNo.Text = ""
txtLockerKeyNo.Text = ""
txtBorrowed.Text = ""
txtReturned.Text = ""
txtBalance.Text = ""
chkReturn.Value = 0
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtPassport.Locked = True
txtDate.Locked = bln
txtCtrlNo.Locked = bln
txtLockerKeyNo.Locked = bln
txtBorrowed.Locked = bln
txtReturned.Locked = bln
txtBalance.Locked = True
picReturn.Enabled = IIf(bln = True, False, True)
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
    End Select
End With
End Sub

Private Sub cmdCancelLocker_Click()
picAddLockerKey.Visible = False
picMain.Enabled = True
picToolbar.Enabled = True
End Sub

Private Sub cmdOKLocker_Click()
If lstResultLocker.ListIndex = -1 Then Exit Sub
CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
Arr = Split(lstResultLocker.List(lstResultLocker.ListIndex), " - ", -1, 1)
txtPassportKey.Text = lstResultLocker.ItemData(lstResultLocker.ListIndex)
t = "SELECT tbl_Operation_Passport.* " & _
    " FROM tbl_Operation_Passport " & _
    " WHERE (PK = " & RETURNTEXTVALUE(txtPassportKey) & ")"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then
    txtDate.Text = Format(rt!DateAll, "mm/dd/yyyy")
Else
    txtDate.Text = Format(Date, "mm/dd/yyyy")
End If
rt.Close
txtPassport.Text = Arr(0)
cmdCancelLocker_Click
End Sub

Private Sub cmdSelectPassport_Click()
picSetFocus.SetFocus
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    picSearchPassport.ZOrder 0
    txtPassportAdd.Text = ""
    picMain.Enabled = False
    picToolbar.Enabled = False
    picSearchPassport.Visible = True
    txtPassportAdd.SetFocus
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert:   PRESS_INSERT
    Case vbKeyF2:       PRESS_F2
    Case vbKeyDelete:   PRESS_DELETE
    Case vbKeyF5:       PRESS_F5
    Case vbKeyF6:       PRESS_F6
    Case vbKeyF8:       PRESS_F8
    Case vbKeyF9:       PRESS_F9
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "OperationLockerRoom", "OperationLockerRoom", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "OperationLockerRoom", "OperationLockerRoom", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "OperationLockerRoom", "OperationLockerRoom", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "OperationLockerRoom", "OperationLockerRoom", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.Height - Me.Height) / 3
Me.Left = (MainForm.Width - Me.Width) / 5
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "OperationLockerRoom", "OperationLockerRoom", ""), "is_LOAD"
If Trim(txtCtrlNo.Text) = "" Then BROWSER GetSetting(App.EXEName, "OperationLockerRoom", "OperationLockerRoom", ""), "is_HOME"
End Sub


Private Sub lstPassportAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If lstPassportAdd.ListIndex = -1 Then Exit Sub
    txtPassportKey.Text = lstPassportAdd.ItemData(lstPassportAdd.ListIndex)
    txtPassport.Text = lstPassportAdd.List(lstPassportAdd.ListIndex)
    picSearchPassport.Visible = False
    picMain.Enabled = True
    picToolbar.Enabled = True
    txtPassport.SetFocus
End If
End Sub

Private Sub lstResultLocker_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKLocker_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Refresh":
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "OperationLockerRoom", "OperationLockerRoom", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "OperationLockerRoom", "OperationLockerRoom", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "OperationLockerRoom", "OperationLockerRoom", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "OperationLockerRoom", "OperationLockerRoom", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":   PRESS_F9
    Case "Post":    PRESS_F8
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtBorrowed_Change()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    txtBalance.Text = RETURNTEXTVALUE(txtBorrowed) - RETURNTEXTVALUE(txtReturned)
End If
End Sub

Private Sub txtBorrowed_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtPassportAdd_Change()
If Trim(txtPassportAdd.Text) = "" Then lstPassportAdd.Clear: Exit Sub
lstPassportAdd.Clear
s = "SELECT PK, PassportNo " & _
    " From dbo.tbl_Operation_Passport " & _
    " WHERE (PostedRegistration = 1) " & _
    " AND (PassportNo LIKE '" & Trim(txtPassportAdd.Text) & "%') " & _
    " ORDER BY PassportNo"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstPassportAdd.AddItem rs!PassportNo
    lstPassportAdd.ItemData(lstPassportAdd.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstPassportAdd.ListCount Then lstPassportAdd.ListIndex = 0
End Sub

Private Sub txtPassportAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstPassportAdd.SetFocus
End Sub

Private Sub txtReturned_Change()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    txtBalance.Text = RETURNTEXTVALUE(txtBorrowed) - RETURNTEXTVALUE(txtReturned)
End If
End Sub

Private Sub txtReturned_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtSearchLocker_Change()
If Trim(txtSearchLocker.Text) = "" Then lstResultLocker.Clear: Exit Sub
lstResultLocker.Clear
s = "SELECT PK, PassportNo, PlayerName " & _
    " FROM tbl_Operation_Passport " & _
    " WHERE (PassportNo LIKE '" & FORMATSQL(Trim(txtSearchLocker.Text)) & "%') " & _
    " AND (RegistrationAdded = 1) " & _
    " ORDER BY PassportNo"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultLocker.AddItem rs!PassportNo & " - " & rs!PlayerName
    lstResultLocker.ItemData(lstResultLocker.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResultLocker.ListCount Then lstResultLocker.ListIndex = 0
End Sub

Private Sub txtSearchLocker_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultLocker.SetFocus
End Sub

Private Sub txtSearchLocker_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub
