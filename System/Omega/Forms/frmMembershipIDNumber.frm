VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMembershipIDNumber 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMembershipIDNumber.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picAdd 
      Height          =   3255
      Left            =   1920
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5741
      BackColor       =   15396057
      Begin VB.CommandButton cmdOKAdd 
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
         Picture         =   "frmMembershipIDNumber.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2600
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelAdd 
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
         Picture         =   "frmMembershipIDNumber.frx":133C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2600
         Width           =   1560
      End
      Begin VB.TextBox txtSearchAdd 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   4215
      End
      Begin VB.ListBox lstResultAdd 
         Height          =   1620
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   4215
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   13
         Top             =   45
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   609
         Caption         =   "Search Member"
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
         Icon            =   "frmMembershipIDNumber.frx":1A98
         ShadowVisible   =   0   'False
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   40
         TabIndex        =   14
         Top             =   40
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   609
         Caption         =   "Search ID Number"
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
         Icon            =   "frmMembershipIDNumber.frx":2032
         ShadowVisible   =   0   'False
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   360
      ScaleHeight     =   2415
      ScaleWidth      =   8295
      TabIndex        =   1
      Top             =   1200
      Width           =   8295
      Begin VB.PictureBox picMember 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   5880
         ScaleHeight     =   2385
         ScaleWidth      =   2385
         TabIndex        =   24
         Top             =   0
         Width           =   2415
         Begin VB.Image imgMember 
            Height          =   2385
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2385
         End
         Begin VB.Image imgMemberLogo 
            Height          =   2385
            Left            =   0
            Picture         =   "frmMembershipIDNumber.frx":25CC
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2385
         End
      End
      Begin VB.TextBox txtCounter 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5040
         MaxLength       =   100
         TabIndex        =   22
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtIDNumber 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   6
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtMemberType 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1320
         Width           =   4695
      End
      Begin VB.TextBox txtMemberName 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   2
         Top             =   960
         Width           =   4695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Number"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Member Type"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   1350
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Member Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   990
         Width           =   1095
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   25
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   26
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
         MouseIcon       =   "frmMembershipIDNumber.frx":6F82
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
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   3870
      Width           =   9090
      _ExtentX        =   16034
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
      Height          =   3255
      Left            =   1920
      TabIndex        =   15
      Top             =   360
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5741
      BackColor       =   15396057
      Begin VB.TextBox txtSearchCnt 
         Height          =   315
         Left            =   120
         MaxLength       =   100
         TabIndex        =   23
         Top             =   2640
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ListBox lstResult 
         Height          =   1620
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   18
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
         Picture         =   "frmMembershipIDNumber.frx":729C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2600
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
         Picture         =   "frmMembershipIDNumber.frx":79F8
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2600
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar3 
         Height          =   345
         Left            =   45
         TabIndex        =   20
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
         Icon            =   "frmMembershipIDNumber.frx":806A
         ShadowVisible   =   0   'False
      End
      Begin RPVGCC.b8TitleBar b8TitleBar4 
         Height          =   345
         Left            =   40
         TabIndex        =   21
         Top             =   40
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   609
         Caption         =   "Search ID Number"
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
         Icon            =   "frmMembershipIDNumber.frx":8604
         ShadowVisible   =   0   'False
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10440
      Top             =   840
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
            Picture         =   "frmMembershipIDNumber.frx":8B9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipIDNumber.frx":9878
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipIDNumber.frx":A552
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipIDNumber.frx":B22C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipIDNumber.frx":BF06
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipIDNumber.frx":CBE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipIDNumber.frx":D8BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipIDNumber.frx":E594
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipIDNumber.frx":F26E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipIDNumber.frx":FB48
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipIDNumber.frx":10822
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipIDNumber.frx":114FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipIDNumber.frx":121D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipIDNumber.frx":12EB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipIDNumber.frx":13B8A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMembershipIDNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public isFind As Long
Public SearchAdd As Long

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim tmp As Long

Dim TmpIDNumber, FIDNumber, sFIDNumber, _
sFullName, Arr, TmpID01, TmpID02, TmpID03, _
i, j, IDNumberCnt, MemberKey, MemberType, MemberName, MemberClass


Private Sub BROWSER(sID, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If sID <> "" Then
            s = "SELECT TOP 1 tbl_Member_IDNumber.* " & _
                " FROM tbl_Member_IDNumber " & _
                " WHERE (IDNumber + ' - ' + Convert(varchar(50),IDCounter) = '" & sID & "') " & _
                " AND (ViewNot = 0) " & _
                " ORDER BY IDNumber + ' - ' + Convert(varchar(50),IDCounter)"
        Else
            s = "SELECT TOP 1 tbl_Member_IDNumber.* " & _
                " FROM tbl_Member_IDNumber " & _
                " WHERE (ViewNot = 0) " & _
                " ORDER BY IDNumber + ' - ' + Convert(varchar(50),IDCounter)"
        End If
    Case "is_HOME"
        If picAdd.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Member_IDNumber.* " & _
            " FROM tbl_Member_IDNumber " & _
            " WHERE (ViewNot = 0) " & _
            " ORDER BY IDNumber + ' - ' + Convert(varchar(50),IDCounter)"
    Case "is_PAGEUP"
        If picAdd.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Member_IDNumber.* " & _
            " FROM tbl_Member_IDNumber " & _
            " WHERE (IDNumber + ' - ' + Convert(varchar(50),IDCounter) < '" & sID & "') " & _
            " AND (ViewNot = 0) " & _
            " ORDER BY IDNumber + ' - ' + Convert(varchar(50),IDCounter) DESC"
    Case "is_PAGEDOWN"
        If picAdd.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Member_IDNumber.* " & _
            " FROM tbl_Member_IDNumber " & _
            " WHERE (IDNumber + ' - ' + Convert(varchar(50),IDCounter) > '" & sID & "') " & _
            " AND (ViewNot = 0) " & _
            " ORDER BY IDNumber + ' - ' + Convert(varchar(50),IDCounter) "
    Case "is_END"
        If picAdd.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Member_IDNumber.* " & _
            " FROM tbl_Member_IDNumber " & _
            " WHERE (ViewNot = 0) " & _
            " ORDER BY IDNumber + ' - ' + Convert(varchar(50),IDCounter) DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    MemberKey = rs!MemberKey
    MemberType = rs!MemberType
    txtIDNumber.Text = rs!IDNumber
    txtCounter.Text = rs!IDCounter
    txtMemberName.Text = rs!MemberName
    Select Case MemberType
        Case 1
            t = "SELECT tbl_Share_IDNumber.* " & _
                " FROM tbl_Share_IDNumber " & _
                " WHERE (IDNumber = '" & rs!IDNumber & "')"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                If rt!ShareType = 3 Then
                    txtMemberType.Text = ""
                    u = "SELECT Name " & _
                        " From tbl_Corporate_Account " & _
                        " WHERE (PK = " & rt!CorporateKey & ")"
                    If ru.State = adStateOpen Then ru.Close
                    ru.Open u, ConnOmega
                    If ru.RecordCount > 0 Then
                        txtMemberType.Text = "NOMINEE OF " & ru!Name
                    End If
                    ru.Close
                Else
                    txtMemberType.Text = "MAIN MEMBER"
                End If
            Else
                txtMemberType.Text = "ASSIGNEE"
            End If
            rt.Close
            
        Case 2
            t = "SELECT LastName, FirstName, MiddleName " & _
                " From tbl_Member_Information " & _
                " WHERE (PK = " & MemberKey & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                txtMemberType.Text = "SPOUSE OF " & rt!LastName & ",  " & rt!FirstName & "  " & rt!MiddleName
            End If
            rt.Close
        Case 3
            t = "SELECT LastName, FirstName, MiddleName " & _
                " From tbl_Member_Information " & _
                " WHERE (PK = " & MemberKey & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                txtMemberType.Text = "DEPENDENT OF " & rt!LastName & ",  " & rt!FirstName & "  " & rt!MiddleName
            End If
            rt.Close
        Case Else: txtMemberType.Text = ""
    End Select
    
    imgMember.Picture = LoadPicture("")
    If IsNull(rs!MemberPicture) = False Then
        imgMember.Picture = LoadPicture(SHOW_IMAGES(rs!PK, 0, "Member ID Number"))
    End If
    
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    SaveSetting App.EXEName, "MemberIDNumber", "MemberID", rs!IDNumber & " - " & rs!IDCounter
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If AccessRights("Membership ID Number", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
'Me.Height = 3765
picAdd.ZOrder 0
b8TitleBar1.Visible = True
b8TitleBar2.Visible = False
txtSearchAdd.Text = ""
SearchAdd = 1
picAdd.Visible = True
txtSearchAdd.SetFocus
End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If AccessRights("Membership ID Number", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If MemberType <> 1 Then MsgBox "Please edit only Main Member!                       ", vbCritical, "Error...": Exit Sub
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_EDITTING
End Sub

Private Sub PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If AccessRights("Membership ID Number", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If MemberType <> 1 Then MsgBox "Can't Delete Dependent ID!                      ", vbCritical, "Error...": Exit Sub
If MsgBox("CAUTION:" & vbCrLf & _
          "All Dependent ID Number will be Deleted!                 " & vbCrLf & vbCrLf & _
          "ARE YOU SURE IN DELETING THIS RECORD?                    ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub

's = "SELECT tbl_Member_Action.* " & _
'    " FROM tbl_Member_Action " & _
'    " WHERE (
'If rs.State = adStateOpen Then rs.Close
'rs.Open s, ConnOmega
'If rs.RecordCount > 0 Then
'
'End If
'rs.Close

On Error GoTo PG:
ConnOmega.Execute "UPDATE tbl_Share_IDNumber " & _
                  " SET MemberKey = Null " & _
                  " WHERE (IDNumber = '" & Trim(txtIDNumber.Text) & "')"
ConnOmega.Execute "DELETE FROM tbl_Member_IDNumber WHERE (MemberKey = " & MemberKey & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "MemberIDNumber", "MemberID", ""), "is_PAGEDOWN"
If Trim(txtIDNumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "MemberIDNumber", "MemberID", ""), "is_HOME"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If MemberKey = 0 Then MsgBox "Please Select Member!                 ", vbCritical, "Error...": Exit Sub
If Trim(txtIDNumber.Text) = "" Then MsgBox "Please Select ID Number!                  ", vbCritical, "Error...": Exit Sub
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    Arr = Split(Trim(txtIDNumber.Text), "-", -1, 1)
    TmpID01 = Arr(0): TmpID02 = Arr(1): TmpID03 = Arr(2)
    s = "SELECT LastName, FirstName, MiddleName, " & _
        " CivilStatus, SpouseLName, SpouseGName, " & _
        " SpouseMName, MemberPicture, SpousePicture " & _
        " From tbl_Member_Information " & _
        " WHERE (PK = " & MemberKey & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        FIDNumber = TmpID01 & "-" & TmpID02 & "-" & TmpID03
        sFIDNumber = FIDNumber
        sFullName = rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName
        
        IDNumberCnt = 1
        t = "SELECT TOP 1 tbl_Member_IDNumber.* " & _
            " FROM tbl_Member_IDNumber " & _
            " WHERE (IDNumber = '" & FIDNumber & "') " & _
            " ORDER BY IDCounter DESC "
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            IDNumberCnt = rt!IDCounter
        End If
        rt.Close
        
        Do
            t = "SELECT TOP 1 tbl_Member_IDNumber.* " & _
                " FROM tbl_Member_IDNumber " & _
                " WHERE (IDNumber = '" & FIDNumber & "') " & _
                " AND (IDCounter = " & IDNumberCnt & ") "
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount = 0 Then
                rt.Close
                Exit Do
            End If
            rt.Close
            IDNumberCnt = IDNumberCnt + 1
        Loop
        
        ConnOmega.Execute "INSERT INTO tbl_Member_IDNumber " & _
                          " (MemberKey, MemberName, IDNumber, IDCounter, MemberType, LastModified, MemberAssignor, MemberChildLine) " & _
                          " VALUES (" & MemberKey & ", '" & FORMATSQL(CStr(sFullName)) & "', " & _
                          " '" & FIDNumber & "', " & IDNumberCnt & ", 1, '" & CStr(Now) & " - " & gbl_CompleteName & "', 1, 0)"
        
        If IsNull(rs!MemberPicture) = False Then
            u = "SELECT PK " & _
                " FROM tbl_Member_IDNumber " & _
                " WHERE (IDNumber = '" & FIDNumber & "')"
            If ru.State = adStateOpen Then ru.Close
            ru.Open u, ConnOmega
            If ru.RecordCount > 0 Then
                SAVE_IMAGES ru!PK, 0, SHOW_IMAGES(MemberKey, 0, "Member"), "Member ID Number"
            End If
            ru.Close
        End If
        
        If rs!CivilStatus = 2 Then
            If Trim(rs!SpouseLName) <> "" And Trim(rs!SpouseGName) <> "" Then
                FIDNumber = TmpID01 & "-" & TmpID02 & "-" & Format(CDbl(TmpID03) + 1, "0#")
                sFullName = rs!SpouseLName & ",  " & rs!SpouseGName & "  " & rs!SpouseMName
                ConnOmega.Execute "INSERT INTO tbl_Member_IDNumber " & _
                                  " (MemberKey, MemberName, IDNumber, IDCounter, MemberType, LastModified, MemberChildLine) " & _
                                  " VALUES (" & MemberKey & ", '" & FORMATSQL(CStr(sFullName)) & "', " & _
                                  " '" & FIDNumber & "', " & IDNumberCnt & ", 2, '" & CStr(Now) & " - " & gbl_CompleteName & "', 0)"
                                
                If IsNull(rs!SpousePicture) = False Then
                    u = "SELECT PK " & _
                        " FROM tbl_Member_IDNumber " & _
                        " WHERE (IDNumber = '" & FIDNumber & "')"
                    If ru.State = adStateOpen Then ru.Close
                    ru.Open u, ConnOmega
                    If ru.RecordCount > 0 Then
                        SAVE_IMAGES ru!PK, 0, SHOW_IMAGES(MemberKey, 0, "Member Spouse"), "Member ID Number"
                    End If
                    ru.Close
                End If
            End If
        End If
        
        j = 0
        i = 2
        t = "SELECT ChildLName, ChildGName, ChildMName, " & _
            " ChildStatus, ChildBirthDate, ChildPicture " & _
            " From tbl_Member_Dependent " & _
            " Where (MemberKey = " & MemberKey & ") " & _
            " ORDER BY ChildBirthDate"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        While Not rt.EOF
            j = j + 1
            If rt!ChildStatus = 1 Then
                If Get_Age(FormatDateTime(rt!ChildBirthDate, vbShortDate), FormatDateTime(Date, vbShortDate)) <= 25 Then
                    i = i + 1
                    FIDNumber = TmpID01 & "-" & TmpID02 & "-" & Format(i, "0#")
                    sFullName = rt!ChildLName & ",  " & rt!ChildGName & "  " & rt!ChildMName
                    ConnOmega.Execute "INSERT INTO tbl_Member_IDNumber " & _
                                      " (MemberKey, MemberName, IDNumber, IDCounter, MemberType, MemberCStatus, MemberBDay, LastModified, MemberChildLine) " & _
                                      " VALUES (" & MemberKey & ", '" & FORMATSQL(CStr(sFullName)) & "', " & _
                                      " '" & FIDNumber & "', " & IDNumberCnt & ", 3, " & rt!ChildStatus & ", '" & FormatDateTime(rt!ChildBirthDate, vbShortDate) & "', " & _
                                      " '" & CStr(Now) & " - " & gbl_CompleteName & "', " & j & ")"
                    
                    If IsNull(rt!ChildPicture) = False Then
                        u = "SELECT PK " & _
                            " FROM tbl_Member_IDNumber " & _
                            " WHERE (IDNumber = '" & FIDNumber & "')"
                        If ru.State = adStateOpen Then ru.Close
                        ru.Open u, ConnOmega
                        If ru.RecordCount > 0 Then
                            SAVE_IMAGES ru!PK, j, SHOW_IMAGES(MemberKey, j, "Member Child"), "Member ID Number (Child)"
                        End If
                        ru.Close
                    End If
                    
                End If
            End If
            rt.MoveNext
        Wend
        rt.Close
    End If
    rs.Close
    'IDHolder
    s = "SELECT tbl_Share_IDNumber.* " & _
        " FROM tbl_Share_IDNumber " & _
        " WHERE (IDNumber = '" & sFIDNumber & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        ConnOmega.Execute "UPDATE tbl_Share_IDNumber " & _
                          " SET MemberKey = " & MemberKey & " " & _
                          " WHERE (IDNumber = '" & sFIDNumber & "')"
        If IIf(IsNull(rs!IDHolder) = True, 0, rs!IDHolder) = 0 Then
            ConnOmega.Execute "UPDATE tbl_Share_IDNumber " & _
                              " SET IDHolder = 1 " & _
                              " WHERE (IDNumber = '" & sFIDNumber & "')"
        End If
    End If
    rs.Close
End If
If TRANSACTIONTYPE = is_EDITTING Then
    
End If
CLEARTEXT
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER sFIDNumber & " - " & IDNumberCnt, "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picAdd.Visible = True Then Exit Sub
If picSearch.Visible = True Then Exit Sub
PopupMenu MainFormPopupF.mnuMemberIDFind, , Toolbar1.Buttons(15).Left, Toolbar1.Buttons(15).Top + Toolbar1.Buttons(15).Height
'Me.Height = 3765
'picSearch.ZOrder 0
'txtSearch.Text = ""
'picSearch.Visible = True
'txtSearch.SetFocus
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picAdd.Visible = True Then cmdCancelAdd_Click: Exit Sub
    If picSearch.Visible = True Then cmdCancel_Click: Exit Sub
    Unload Me
Else
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "MemberIDNumber", "MemberID", ""), "is_LOAD"
    If Trim(txtIDNumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "MemberIDNumber", "MemberID", ""), "is_HOME"
End If
End Sub

Private Sub CLEARTEXT()
MemberKey = 0
MemberType = 0
txtIDNumber.Text = ""
txtCounter.Text = "0000"
txtMemberName.Text = ""
txtMemberType.Text = ""
imgMember.Picture = LoadPicture("")
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtIDNumber.Locked = bln
txtMemberName.Locked = bln
txtMemberType.Locked = bln
txtCounter.Locked = True
End Sub

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
cmdCancelAdd_Click
End Sub

Private Sub b8TitleBar2_CLoseClick()
cmdCancelAdd_Click
End Sub

Private Sub b8TitleBar3_CLoseClick()
cmdCancel_Click
End Sub

Private Sub cmdCancel_Click()
picToolbar.Enabled = True
picMain.Enabled = True
'Me.Height = 2895
picSearch.Visible = False
End Sub

Private Sub cmdCancelAdd_Click()
picToolbar.Enabled = True
picMain.Enabled = True
'Me.Height = 2895
picAdd.Visible = False
End Sub

Private Sub cmdOK_Click()
If lstResult.ListIndex = -1 Then Exit Sub
Arr = Split(lstResult.List(lstResult.ListIndex), " - ", -1, 1)
'MsgBox CStr(Arr(0))
BROWSER CStr(Arr(0)) & " - " & txtSearchCnt.Text, "is_LOAD"
cmdCancel_Click
End Sub

Private Sub cmdOKAdd_Click()
If lstResultAdd.ListIndex = -1 Then Exit Sub
Select Case SearchAdd
    Case 1
        TmpIDNumber = lstResultAdd.List(lstResultAdd.ListIndex)
        MemberClass = 0
        s = "SELECT ShareType " & _
            " FROM tbl_Share_IDNumber " & _
            " WHERE (IDNumber = '" & CStr(TmpIDNumber) & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            MemberClass = rs!ShareType
        End If
        rs.Close
        b8TitleBar1.Visible = False
        b8TitleBar2.Visible = True
        txtSearchAdd.Text = ""
        txtSearchAdd.SetFocus
        SearchAdd = 2
    Case 2
        CLEARTEXT
        TOOLBARFUNC 2
        TRANSACTIONTYPE = is_ADDING
        txtIDNumber.Text = TmpIDNumber
        MemberKey = lstResultAdd.ItemData(lstResultAdd.ListIndex)
        txtMemberName.Text = lstResultAdd.List(lstResultAdd.ListIndex)
        cmdCancelAdd_Click
        txtMemberType.SetFocus
End Select
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
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "MemberIDNumber", "MemberID", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "MemberIDNumber", "MemberID", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "MemberIDNumber", "MemberID", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "MemberIDNumber", "MemberID", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
'Me.Height = 2895
Me.Top = (MainForm.Height - Me.Height) / 3
Me.Left = (MainForm.Width - Me.Width) / 5
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER GetSetting(App.EXEName, "MemberIDNumber", "MemberID", ""), "is_LOAD"
If Trim(txtIDNumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "MemberIDNumber", "MemberID", ""), "is_HOME"

tmp = SetWindowLong(txtSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearchAdd.hwnd, GWL_STYLE, GetWindowLong(txtSearchAdd.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtIDNumber.hwnd, GWL_STYLE, GetWindowLong(txtIDNumber.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtMemberName.hwnd, GWL_STYLE, GetWindowLong(txtMemberName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtMemberType.hwnd, GWL_STYLE, GetWindowLong(txtMemberType.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picAdd.Visible = True Then Cancel = -1
If picSearch.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstResult_Click()
If lstResult.ListIndex = -1 Then txtSearchCnt.Text = "": Exit Sub
txtSearchCnt.Text = ""
t = "SELECT IDCounter " & _
    " From tbl_Member_IDNumber " & _
    " WHERE (PK = " & lstResult.ItemData(lstResult.ListIndex) & ") "
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount > 0 Then
    txtSearchCnt.Text = rt!IDCounter
End If
rt.Close
End Sub

Private Sub lstResult_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub lstResultAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAdd_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "MemberIDNumber", "MemberID", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "MemberIDNumber", "MemberID", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "MemberIDNumber", "MemberID", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "MemberIDNumber", "MemberID", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstResult.Clear: Exit Sub
lstResult.Clear
Select Case isFind
    Case 1
        s = "SELECT PK, IDNumber, MemberName " & _
            " From tbl_Member_IDNumber " & _
            " WHERE (MemberName LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
            " ORDER BY MemberName, IDNumber DESC"
    Case 2
        s = "SELECT PK, IDNumber, MemberName " & _
            " From tbl_Member_IDNumber " & _
            " WHERE (IDNumber LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
            " ORDER BY MemberName, IDNumber DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResult.AddItem rs!IDNumber & " - " & rs!MemberName
    lstResult.ItemData(lstResult.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResult.ListCount Then lstResult.ListIndex = 0
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResult.SetFocus
End Sub

Private Sub txtSearchAdd_Change()
If Trim(txtSearchAdd.Text) = "" Then lstResultAdd.Clear: Exit Sub
lstResultAdd.Clear
Select Case SearchAdd
    Case 1
        's = "SELECT PK, IDNumber as TmpName " & _
            " From tbl_Share_IDNumber " & _
            " WHERE ((MemberKey IS NULL) AND (ShareType = 1) AND (IDNumber LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%')) " & _
            " OR ((CorporateKey IS NOT NULL) AND (ShareType = 3) AND (IDNumber LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%')) " & _
            " ORDER BY IDNumber"
        s = "SELECT PK, IDNumber  as TmpName " & _
            " From tbl_Share_IDNumber " & _
            " WHERE ((MemberKey IS NULL) AND (CorporateKey IS NULL) AND (ShareType <> 3) AND (IDNumber LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%')) " & _
            " OR ((MemberKey IS NULL) AND (CorporateKey IS NOT NULL) AND (ShareType = 3) AND (IDNumber LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%')) " & _
            " ORDER BY IDNumber"
    Case 2
        s = "SELECT PK, LastName + ',  ' + FirstName + '  ' + MiddleName AS TmpName " & _
            " From tbl_Member_Information " & _
            " WHERE (LastName LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%') " & _
            " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultAdd.AddItem rs!TmpName
    lstResultAdd.ItemData(lstResultAdd.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
If lstResultAdd.ListCount Then lstResultAdd.ListIndex = 0
End Sub

Private Sub txtSearchAdd_GotFocus()
HTEXT txtSearchAdd
End Sub

Private Sub txtSearchAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResultAdd.SetFocus
End Sub
