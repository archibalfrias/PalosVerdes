VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMembershipCorporate 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   9075
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMembershipCorporate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   9075
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   20
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   21
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
         MouseIcon       =   "frmMembershipCorporate.frx":0CCA
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
      Height          =   1815
      Left            =   840
      ScaleHeight     =   1815
      ScaleWidth      =   7335
      TabIndex        =   1
      Top             =   1200
      Width           =   7335
      Begin VB.TextBox txtContact 
         Height          =   315
         Left            =   960
         MaxLength       =   100
         TabIndex        =   18
         Top             =   720
         Width           =   6375
      End
      Begin VB.CommandButton cmdID2 
         Caption         =   "..."
         Height          =   315
         Left            =   2280
         MouseIcon       =   "frmMembershipCorporate.frx":0FE4
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   1440
         Width           =   315
      End
      Begin VB.CommandButton cmdID1 
         Caption         =   "..."
         Height          =   315
         Left            =   2280
         MouseIcon       =   "frmMembershipCorporate.frx":12EE
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1080
         Width           =   315
      End
      Begin VB.TextBox txtIDNumber2 
         Height          =   315
         Left            =   960
         MaxLength       =   100
         TabIndex        =   8
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtIDNumber1 
         Height          =   315
         Left            =   960
         MaxLength       =   100
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   960
         MaxLength       =   100
         TabIndex        =   3
         Top             =   0
         Width           =   6375
      End
      Begin VB.TextBox txtAddress 
         Height          =   315
         Left            =   960
         MaxLength       =   100
         TabIndex        =   2
         Top             =   360
         Width           =   6375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact #"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   750
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Assign ID 2"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   1470
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Assign ID 1"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   1110
         Width           =   1095
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   390
         Width           =   1095
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   3195
      Width           =   9075
      _ExtentX        =   16007
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
   Begin RPVGCC.b8Container picAdd 
      Height          =   2955
      Left            =   2280
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5212
      BackColor       =   15396057
      Begin VB.ListBox lstResultAdd 
         Height          =   1425
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtSearchAdd 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   4215
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
         Picture         =   "frmMembershipCorporate.frx":15F8
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2355
         Width           =   1560
      End
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
         Picture         =   "frmMembershipCorporate.frx":1D54
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2355
         Width           =   1560
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   40
         TabIndex        =   17
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
         Icon            =   "frmMembershipCorporate.frx":23C6
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
            Picture         =   "frmMembershipCorporate.frx":2960
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipCorporate.frx":363A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipCorporate.frx":4314
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipCorporate.frx":4FEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipCorporate.frx":5CC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipCorporate.frx":69A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipCorporate.frx":767C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipCorporate.frx":8356
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipCorporate.frx":9030
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipCorporate.frx":990A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipCorporate.frx":A5E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipCorporate.frx":B2BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipCorporate.frx":BF98
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipCorporate.frx":CC72
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMembershipCorporate.frx":D94C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMembershipCorporate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim iIDSearch   As Long
Dim tmp         As Long

Dim tmpID1, tmpID2, CorporateKey, sCName

Private Sub BROWSER(sName, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If sName <> "" Then
            s = "SELECT TOP 1 tbl_Corporate_Account.* " & _
                " From tbl_Corporate_Account " & _
                " WHERE (Name = '" & FORMATSQL(Trim(CStr(sName))) & "') " & _
                " ORDER BY Name"
        Else
            s = "SELECT TOP 1 tbl_Corporate_Account.* " & _
                " From tbl_Corporate_Account " & _
                " ORDER BY Name"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Corporate_Account.* " & _
            " From tbl_Corporate_Account " & _
            " ORDER BY Name"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Corporate_Account.* " & _
            " From tbl_Corporate_Account " & _
            " WHERE (Name < '" & FORMATSQL(Trim(CStr(sName))) & "') " & _
            " ORDER BY Name DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Corporate_Account.* " & _
            " From tbl_Corporate_Account " & _
            " WHERE (Name > '" & FORMATSQL(Trim(CStr(sName))) & "') " & _
            " ORDER BY Name "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Corporate_Account.* " & _
            " From tbl_Corporate_Account " & _
            " ORDER BY Name DESC"
    Case Else: Exit Sub
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtName.Text = rs!Name
    txtAddress.Text = rs!Address
    txtContact.Text = rs!Contact
    txtIDNumber1.Text = rs!ID1
    txtIDNumber2.Text = rs!ID2
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    SaveSetting App.EXEName, "CorporateName", "CorpName", rs!Name
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If AccessRights("Corporate Account", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
txtName.SetFocus
End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("Corporate Account", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_EDITTING
tmpID1 = txtIDNumber1.Text
tmpID2 = txtIDNumber2.Text
End Sub

Private Sub PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("Corporate Account", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If

End Sub

Private Sub PRESS_F5()
If Trim(txtName.Text) = "" Then MsgBox "Please Supply Corporate Name!                 ", vbCritical, "Error...": txtName.SetFocus: Exit Sub
If Trim(txtAddress.Text) = "" Then MsgBox "Please Supply Corporate Address!                   ", vbCritical, "Error..": txtAddress.SetFocus: Exit Sub

On Error GoTo PG:
sCName = Trim(txtName.Text)
If TRANSACTIONTYPE = is_ADDING Then
    If Trim(txtIDNumber1.Text) <> "" Then
        s = "SELECT CorporateKey " & _
            " From tbl_Share_IDNumber " & _
            " WHERE (ShareType = 3) " & _
            " AND (IDNumber = '" & FORMATSQL(Trim(txtIDNumber1.Text)) & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            If IsNull(rs!CorporateKey) = False Then MsgBox "This ID '" & Trim(txtIDNumber1.Text) & "' Already Assign!                   ", vbCritical, "Error...": Exit Sub
        End If
        rs.Close
    End If
    
    If Trim(txtIDNumber2.Text) <> "" Then
        s = "SELECT CorporateKey " & _
            " From tbl_Share_IDNumber " & _
            " WHERE (ShareType = 3) " & _
            " AND (IDNumber = '" & FORMATSQL(Trim(txtIDNumber2.Text)) & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            If IsNull(rs!CorporateKey) = False Then MsgBox "This ID '" & Trim(txtIDNumber2.Text) & "' Already Assign!                   ", vbCritical, "Error...": Exit Sub
        End If
        rs.Close
    End If
    
    ConnOmega.Execute "INSERT INTO tbl_Corporate_Account " & _
                      " (Name, Address, Contact, LastModified, ID1, ID2) " & _
                      " VALUES ('" & FORMATSQL(Trim(txtName.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtAddress.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtContact.Text)) & "', " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "', " & _
                      " '" & FORMATSQL(Trim(txtIDNumber1.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtIDNumber2.Text)) & "')"
                   
    CorporateKey = 0
    s = "SELECT PK " & _
        " FROM tbl_Corporate_Account " & _
        " WHERE (Name = '" & FORMATSQL(Trim(txtName.Text)) & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        CorporateKey = rs!PK
    End If
    rs.Close
    If CDbl(CorporateKey) <> 0 Then
        If Trim(txtIDNumber1.Text) <> "" Then
            ConnOmega.Execute "UPDATE tbl_Share_IDNumber " & _
                              " SET CorporateKey = " & CorporateKey & ", " & _
                              " IDHolder = 3 " & _
                              " WHERE (IDNumber = '" & FORMATSQL(Trim(txtIDNumber1.Text)) & "')"
        End If
        If Trim(txtIDNumber2.Text) <> "" Then
            ConnOmega.Execute "UPDATE tbl_Share_IDNumber " & _
                              " SET CorporateKey = " & CorporateKey & ", " & _
                              " IDHolder = 3 " & _
                              " WHERE (IDNumber = '" & FORMATSQL(Trim(txtIDNumber2.Text)) & "')"
        End If
    End If
End If
If TRANSACTIONTYPE = is_EDITTING Then
    If Trim(txtIDNumber1.Text) <> "" Then
        s = "SELECT CorporateKey " & _
            " From tbl_Share_IDNumber " & _
            " WHERE (ShareType = 3) " & _
            " AND (IDNumber = '" & FORMATSQL(Trim(txtIDNumber1.Text)) & "')" & _
            " AND (CorporateKey <> " & Statusbar1.Panels(1).Text & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            If IsNull(rs!CorporateKey) = False Then MsgBox "This ID '" & Trim(txtIDNumber1.Text) & "' Already Assign!                   ", vbCritical, "Error...": Exit Sub
        End If
        rs.Close
    End If
    If Trim(txtIDNumber2.Text) <> "" Then
        s = "SELECT CorporateKey " & _
            " From tbl_Share_IDNumber " & _
            " WHERE (ShareType = 3) " & _
            " AND (IDNumber = '" & FORMATSQL(Trim(txtIDNumber2.Text)) & "')" & _
            " AND (CorporateKey <> " & Statusbar1.Panels(1).Text & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            If IsNull(rs!CorporateKey) = False Then MsgBox "This ID '" & Trim(txtIDNumber2.Text) & "' Already Assign!                   ", vbCritical, "Error...": Exit Sub
        End If
        rs.Close
    End If
    ConnOmega.Execute "UPDATE tbl_Corporate_Account " & _
                      " SET Name = '" & FORMATSQL(Trim(txtName.Text)) & "', " & _
                      " Address = '" & FORMATSQL(Trim(txtAddress.Text)) & "', " & _
                      " Contact = '" & FORMATSQL(Trim(txtContact.Text)) & "', " & _
                      " ID1 = '" & Trim(txtIDNumber1.Text) & "', " & _
                      " ID2 = '" & Trim(txtIDNumber2.Text) & "', " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    
    If Trim(CStr(tmpID1)) <> Trim(txtIDNumber1.Text) Then
        ConnOmega.Execute "UPDATE tbl_Share_IDNumber " & _
                          " SET CorporateKey = Null " & _
                          " WHERE (IDNumber = '" & FORMATSQL(Trim(CStr(tmpID1))) & "')"
    End If
    If Trim(CStr(tmpID2)) <> Trim(txtIDNumber2.Text) Then
        ConnOmega.Execute "UPDATE tbl_Share_IDNumber " & _
                          " SET CorporateKey = Null " & _
                          " WHERE (IDNumber = '" & FORMATSQL(Trim(CStr(tmpID2))) & "')"
    End If
    
    If Trim(txtIDNumber1.Text) <> "" Then
         ConnOmega.Execute "UPDATE tbl_Share_IDNumber " & _
                           " SET CorporateKey = " & Statusbar1.Panels(1).Text & ", " & _
                           " IDHolder = 3 " & _
                           " WHERE (IDNumber = '" & FORMATSQL(Trim(txtIDNumber1.Text)) & "')"
    End If
    If Trim(txtIDNumber2.Text) <> "" Then
        ConnOmega.Execute "UPDATE tbl_Share_IDNumber " & _
                           " SET CorporateKey = " & Statusbar1.Panels(1).Text & ", " & _
                           " IDHolder = 3 " & _
                           " WHERE (IDNumber = '" & FORMATSQL(Trim(txtIDNumber2.Text)) & "')"
    End If
End If
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER sCName, "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub

End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    Unload Me
Else
    If picAdd.Visible = True Then cmdCancelAdd_Click: Exit Sub
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "CorporateName", "CorpName", ""), "is_LOAD"
    If Trim(txtName.Text) = "" Then BROWSER GetSetting(App.EXEName, "CorporateName", "CorpName", ""), "is_HOME"
End If
End Sub

Private Sub CLEARTEXT()
txtName.Text = ""
txtAddress.Text = ""
txtContact.Text = ""
txtIDNumber1.Text = ""
txtIDNumber2.Text = ""
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtName.Locked = bln
txtAddress.Locked = bln
txtContact.Locked = bln
txtIDNumber1.Locked = True
txtIDNumber2.Locked = True
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

Private Sub b8TitleBar2_CLoseClick()
cmdCancelAdd_Click
End Sub

Private Sub cmdCancelAdd_Click()
picMain.Enabled = True
picToolbar.Enabled = True
picAdd.Visible = False
End Sub

Private Sub cmdID1_Click()
txtIDNumber1.SetFocus
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
iIDSearch = 1
picAdd.ZOrder 0
picToolbar.Enabled = False
picMain.Enabled = False
txtSearchAdd.Text = ""
picAdd.Visible = True
txtSearchAdd.SetFocus
End Sub

Private Sub cmdID2_Click()
txtIDNumber2.SetFocus
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
iIDSearch = 2
picAdd.ZOrder 0
picToolbar.Enabled = False
picMain.Enabled = False
txtSearchAdd.Text = ""
picAdd.Visible = True
txtSearchAdd.SetFocus
End Sub

Private Sub cmdOKAdd_Click()
If lstResultAdd.ListIndex <= -1 Then Exit Sub
If iIDSearch = 1 Then
    txtIDNumber1.Text = lstResultAdd.List(lstResultAdd.ListIndex)
ElseIf iIDSearch = 2 Then
    txtIDNumber2.Text = lstResultAdd.List(lstResultAdd.ListIndex)
End If
cmdCancelAdd_Click
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
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "CorporateName", "CorpName", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "CorporateName", "CorpName", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "CorporateName", "CorpName", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "CorporateName", "CorpName", ""), "is_END"
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
BROWSER GetSetting(App.EXEName, "CorporateName", "CorpName", ""), "is_LOAD"
If Trim(txtName.Text) = "" Then BROWSER GetSetting(App.EXEName, "CorporateName", "CorpName", ""), "is_HOME"

tmp = SetWindowLong(txtName.hwnd, GWL_STYLE, GetWindowLong(txtName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtAddress.hwnd, GWL_STYLE, GetWindowLong(txtAddress.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtContact.hwnd, GWL_STYLE, GetWindowLong(txtContact.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtIDNumber1.hwnd, GWL_STYLE, GetWindowLong(txtIDNumber1.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtIDNumber2.hwnd, GWL_STYLE, GetWindowLong(txtIDNumber2.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearchAdd.hwnd, GWL_STYLE, GetWindowLong(txtSearchAdd.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picAdd.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstResultAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAdd_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "CorporateName", "CorpName", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "CorporateName", "CorpName", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "CorporateName", "CorpName", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "CorporateName", "CorpName", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtSearchAdd_Change()
If Trim(txtSearchAdd.Text) = "" Then lstResultAdd.Clear: Exit Sub
lstResultAdd.Clear
If iIDSearch = 1 Then
    s = "SELECT IDNumber " & _
        " From tbl_Share_IDNumber " & _
        " WHERE (ShareType = 3) " & _
        " AND (IDHolder = 0) " & _
        " AND (CorporateKey IS NULL) " & _
        " AND (IDNumber LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%') " & _
        " AND (IDNumber <> '" & Trim(txtIDNumber2.Text) & "') " & _
        " ORDER BY IDNumber"
ElseIf iIDSearch = 2 Then
    s = "SELECT IDNumber " & _
        " From tbl_Share_IDNumber " & _
        " WHERE (ShareType = 3) " & _
        " AND (IDHolder = 0) " & _
        " AND (CorporateKey IS NULL) " & _
        " AND (IDNumber LIKE '" & FORMATSQL(Trim(txtSearchAdd.Text)) & "%') " & _
        " AND (IDNumber <> '" & Trim(txtIDNumber1.Text) & "') " & _
        " ORDER BY IDNumber"
Else
    Exit Sub
End If
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResultAdd.AddItem rs!IDNumber
    rs.MoveNext
Wend
rs.Close
If lstResultAdd.ListCount Then lstResultAdd.ListIndex = 0
End Sub

Private Sub txtSearchAdd_GotFocus()
HTEXT txtSearchAdd
End Sub

Private Sub txtSearchAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then lstResultAdd.SetFocus
End Sub
