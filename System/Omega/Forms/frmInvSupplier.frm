VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInvSupplier 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5730
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
   Icon            =   "frmInvSupplier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   32
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   33
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
         MouseIcon       =   "frmInvSupplier.frx":08CA
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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   840
      TabIndex        =   31
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   21
      Top             =   5415
      Width           =   9075
      _ExtentX        =   16007
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
   Begin VB.PictureBox picMain 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   1200
      ScaleHeight     =   3975
      ScaleWidth      =   6660
      TabIndex        =   0
      Top             =   1200
      Width           =   6660
      Begin VB.ComboBox cmbAccountName 
         Height          =   315
         Left            =   2520
         TabIndex        =   30
         Text            =   "Combo1"
         Top             =   3600
         Width           =   4095
      End
      Begin VB.TextBox txtAccountCode 
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   29
         Top             =   3600
         Width           =   1030
      End
      Begin VB.TextBox txtAddress3 
         Height          =   315
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   10
         Top             =   1800
         Width           =   5175
      End
      Begin VB.TextBox txtAddress2 
         Height          =   315
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   9
         Top             =   1440
         Width           =   5175
      End
      Begin VB.TextBox txtContact 
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   8
         Top             =   3240
         Width           =   5175
      End
      Begin VB.TextBox txtEmail 
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   7
         Top             =   2880
         Width           =   5175
      End
      Begin VB.TextBox txtFaxNo 
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2520
         Width           =   5175
      End
      Begin VB.TextBox txtTelNo 
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2160
         Width           =   5175
      End
      Begin VB.TextBox txtAddress1 
         Height          =   315
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   4
         Top             =   1080
         Width           =   5175
      End
      Begin VB.ComboBox cmbSuppType 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtSuppName 
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   2
         Top             =   360
         Width           =   5175
      End
      Begin VB.TextBox txtSuppCode 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "POSTING GROUP"
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS 3"
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS 2"
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "CONTACT PERSON"
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "EMAIL ADD"
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "FAX NO"
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "TEL NO"
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS 1"
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER TYPE"
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER NAME"
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER CODE"
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1215
      End
   End
   Begin RPVGCC.b8Container picSearch 
      Height          =   4095
      Left            =   2520
      TabIndex        =   22
      Top             =   1080
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
         Picture         =   "frmInvSupplier.frx":0BE4
         Style           =   1  'Graphical
         TabIndex        =   26
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
         Picture         =   "frmInvSupplier.frx":1256
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3480
         Width           =   1560
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   4215
      End
      Begin VB.ListBox lstResult 
         Height          =   2595
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   4215
      End
      Begin RPVGCC.b8TitleBar b8TitleBar2 
         Height          =   345
         Left            =   45
         TabIndex        =   27
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
         Icon            =   "frmInvSupplier.frx":19B2
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
            Picture         =   "frmInvSupplier.frx":1F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvSupplier.frx":2C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvSupplier.frx":3900
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvSupplier.frx":45DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvSupplier.frx":52B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvSupplier.frx":5F8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvSupplier.frx":6C68
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvSupplier.frx":7942
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvSupplier.frx":861C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvSupplier.frx":8EF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvSupplier.frx":9BD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvSupplier.frx":A8AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvSupplier.frx":B584
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvSupplier.frx":C25E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvSupplier.frx":CF38
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInvSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2


Dim sCodeName           As String
Dim sCode               As String
Dim iSupplierType       As Long
Dim iSupplierTypeTmp    As Long
Dim tmp                 As Long

Dim Arr, i

Private Sub BROWSER(sCodName, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If sCodName <> "" Then
            s = "SELECT TOP 1 tbl_Inv_Supplier.* " & _
                " FROM tbl_Inv_Supplier " & _
                " WHERE (SupplierName + ' - ' + SupplierCode = '" & FORMATSQL(CStr(sCodName)) & "')" & _
                " ORDER BY SupplierName, SupplierCode"
        Else
            s = "SELECT TOP 1 tbl_Inv_Supplier.* " & _
                " FROM tbl_Inv_Supplier " & _
                " ORDER BY SupplierName + ' - ' + SupplierCode"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_Supplier.* " & _
            " FROM tbl_Inv_Supplier " & _
            " ORDER BY SupplierName + ' - ' + SupplierCode"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_Supplier.* " & _
            " FROM tbl_Inv_Supplier " & _
            " WHERE (SupplierName + ' - ' + SupplierCode < '" & FORMATSQL(CStr(sCodName)) & "')" & _
            " ORDER BY SupplierName + ' - ' + SupplierCode DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_Supplier.* " & _
            " FROM tbl_Inv_Supplier " & _
            " WHERE (SupplierName + ' - ' + SupplierCode > '" & FORMATSQL(CStr(sCodName)) & "')" & _
            " ORDER BY SupplierName + ' - ' + SupplierCode "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSearch.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_Supplier.* " & _
            " FROM tbl_Inv_Supplier " & _
            " ORDER BY SupplierName + ' - ' + SupplierCode DESC"
    Case Else: Exit Sub
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    iSupplierType = rs!Type
    txtSuppCode.Text = rs!SupplierCode
    txtSuppName.Text = rs!SupplierName
    txtAddress1.Text = rs!Address1
    txtAddress2.Text = rs!Address2
    txtAddress3.Text = rs!Address3
    txtTelNo.Text = rs!TelNo
    txtFaxNo.Text = rs!FaxNo
    txtEmail.Text = rs!Email
    txtContact.Text = rs!ContactPerson
    cmbSuppType.ListIndex = rs!Type
    If IsNull(rs!PostingGroup) = False Then
        'txtAccountCode.Text = rs!PostingGroup
        txtAccountCode.Text = ""
        cmbAccountName.Text = ""
    Else
        txtAccountCode.Text = ""
        cmbAccountName.Text = ""
        cmbAccountName.ListIndex = -1
    End If
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    SaveSetting App.EXEName, "SupplierCodeName", "SuppCodeName", rs!SupplierName & " - " & rs!SupplierCode
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If AccessRights("Inventory Supplier", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
'Me.Caption = "Supplier - New"
txtSuppName.SetFocus
End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If AccessRights("Inventory Supplier", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_EDITTING
'Me.Caption = "Supplier - Edit"
End Sub

Private Sub PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If AccessRights("Inventory Supplier", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If MsgBox("ARE YOU SURE IN DELETING THIS RECORD?                        ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_Inv_Supplier WHERE (PK = " & Statusbar1.Panels(1) & ")"
CLEARTEXT

Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
If picSearch.Visible = True Then Exit Sub
If Trim(txtSuppName.Text) = "" Then MsgBox "Please Supply Supplier Name!                                ", vbCritical, "Error...": txtSuppName.SetFocus: Exit Sub
If cmbSuppType.ItemData(cmbSuppType.ListIndex) <= 0 Then MsgBox "Please Select Supplier Type!                        ", vbCritical, "Error...": cmbSuppType.SetFocus: Exit Sub
'On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    sCode = Trim(txtSuppCode.Text)
    Do
        s = "SELECT tbl_Inv_Supplier.* " & _
            " FROM tbl_Inv_Supplier " & _
            " WHERE (SupplierCode = '" & sCode & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If 'gbl_CompleteName
        rs.Close
        sCode = Format(CDbl(sCode) + 1, "0000#")
    Loop
    sCodeName = FORMATSQL(Trim(txtSuppName.Text) & " - " & Trim(CStr(sCode)))
    ConnOmega.Execute "INSERT INTO tbl_Inv_Supplier " & _
                      " (SupplierCode, SupplierName, Type, Address1, Address2, Address3, " & _
                      " TelNo, FaxNo, Email,  ContactPerson, LastModified) " & _
                      " VALUES('" & sCode & "', '" & FORMATSQL(Trim(txtSuppName.Text)) & "', " & _
                      " " & cmbSuppType.ItemData(cmbSuppType.ListIndex) & ", " & _
                      " '" & FORMATSQL(Trim(txtAddress1.Text)) & "', '" & FORMATSQL(Trim(txtAddress2.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtAddress3.Text)) & "', '" & FORMATSQL(Trim(txtTelNo.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtFaxNo.Text)) & "', '" & FORMATSQL(Trim(txtEmail.Text)) & "', " & _
                      " '" & FORMATSQL(Trim(txtContact.Text)) & "', '" & CStr(Now) & " - " & gbl_CompleteName & "')"
End If
If TRANSACTIONTYPE = is_EDITTING Then
    sCode = Trim(txtSuppCode.Text)
    Do
        s = "SELECT tbl_Inv_Supplier.* " & _
            " FROM tbl_Inv_Supplier " & _
            " WHERE (SupplierCode = '" & sCode & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If
        rs.Close
        sCode = Format(CDbl(sCode) + 1, "0000#")
    Loop
    sCodeName = FORMATSQL(Trim(txtSuppName.Text) & " - " & Trim(CStr(sCode)))
    ConnOmega.Execute "UPDATE tbl_Inv_Supplier " & _
                      " SET SupplierCode = '" & sCode & "', " & _
                      " SupplierName = '" & FORMATSQL(Trim(txtSuppName.Text)) & "', " & _
                      " Type = " & cmbSuppType.ItemData(cmbSuppType.ListIndex) & ", " & _
                      " Address1 = '" & FORMATSQL(Trim(txtAddress1.Text)) & "', " & _
                      " Address2 = '" & FORMATSQL(Trim(txtAddress2.Text)) & "', " & _
                      " Address3 = '" & FORMATSQL(Trim(txtAddress3.Text)) & "', " & _
                      " TelNo = '" & FORMATSQL(Trim(txtTelNo.Text)) & "', " & _
                      " FaxNo = '" & FORMATSQL(Trim(txtFaxNo.Text)) & "', " & _
                      " Email = '" & FORMATSQL(Trim(txtEmail.Text)) & "', " & _
                      " ContactPerson = '" & FORMATSQL(Trim(txtContact.Text)) & "', " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
End If
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
'Me.Caption = "Supplier - Browse"
BROWSER sCodeName, "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSearch.Visible = True Then Exit Sub
picMain.Enabled = False
picToolbar.Enabled = False
picSearch.ZOrder 0
txtSearch.Text = ""
picSearch.Visible = True
txtSearch.SetFocus
End Sub

Private Sub PRESS_F9()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSearch.Visible = True Then Exit Sub
PopupMenu MainFormPopupF.mnuSupplierReport, , Toolbar1.Buttons(17).Left, Toolbar1.Buttons(17).Top + Toolbar1.Buttons(17).Height
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSearch.Visible = True Then cmdCancel_Click: Exit Sub
    Unload Me
Else
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    'Me.Caption = "Supplier - Browse"
    BROWSER GetSetting(App.EXEName, "SupplierCodeName", "SuppCodeName", ""), "is_LOAD"
    If Trim(txtSuppCode.Text) = "" Then BROWSER GetSetting(App.EXEName, "SupplierCodeName", "SuppCodeName", ""), "is_HOME"
End If
End Sub

Private Sub CLEARTEXT()
iSupplierType = 0
txtSuppCode.Text = ""
txtSuppName.Text = ""
txtAddress1.Text = ""
txtAddress2.Text = ""
txtAddress3.Text = ""
txtTelNo.Text = ""
txtFaxNo.Text = ""
txtEmail.Text = ""
txtContact.Text = ""
cmbSuppType.ListIndex = 0
txtAccountCode.Text = ""
cmbAccountName.Text = ""
cmbAccountName.ListIndex = -1
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtSuppCode.Locked = True
txtSuppName.Locked = bln
txtAddress1.Locked = bln
txtAddress2.Locked = bln
txtAddress3.Locked = bln
txtTelNo.Locked = bln
txtFaxNo.Locked = bln
txtEmail.Locked = bln
txtContact.Locked = bln
cmbSuppType.Locked = bln
txtAccountCode.Locked = bln
cmbAccountName.Locked = bln
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
cmdCancel_Click
End Sub

Private Sub cmbSuppType_Click()
If TRANSACTIONTYPE = is_REFRESH Then Exit Sub
If cmbSuppType.ItemData(cmbSuppType.ListIndex) = 0 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Then
    s = "SELECT TOP 1 tbl_Inv_Supplier.* " & _
        " FROM tbl_Inv_Supplier " & _
        " WHERE (Type = " & cmbSuppType.ItemData(cmbSuppType.ListIndex) & ") " & _
        " ORDER BY SupplierCode DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        txtSuppCode.Text = Format(CDbl(rs!SupplierCode) + 1, "0000#")
    Else
        txtSuppCode.Text = CStr(cmbSuppType.ItemData(cmbSuppType.ListIndex)) & "0001"
    End If
    rs.Close
End If
If TRANSACTIONTYPE = is_EDITTING Then
    If iSupplierType <> cmbSuppType.ItemData(cmbSuppType.ListIndex) Then
        s = "SELECT TOP 1 tbl_Inv_Supplier.* " & _
            " FROM tbl_Inv_Supplier " & _
            " WHERE (Type = " & cmbSuppType.ItemData(cmbSuppType.ListIndex) & ") " & _
            " ORDER BY SupplierCode DESC"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            txtSuppCode.Text = Format(CDbl(rs!SupplierCode) + 1, "0000#")
        Else
            txtSuppCode.Text = CStr(cmbSuppType.ItemData(cmbSuppType.ListIndex)) & "0001"
        End If
        rs.Close
    Else
        s = "SELECT TOP 1 tbl_Inv_Supplier.* " & _
            " FROM tbl_Inv_Supplier " & _
            " WHERE (PK = " & Statusbar1.Panels(1).Text & ") "
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            txtSuppCode.Text = rs!SupplierCode
        End If
        rs.Close
    End If
End If
End Sub

Private Sub cmdCancel_Click()
picMain.Enabled = True
picToolbar.Enabled = True
picSearch.Visible = False
End Sub

Private Sub cmdOK_Click()
If lstResult.ListIndex <= -1 Then Exit Sub
Arr = Split(lstResult.List(lstResult.ListIndex), " - ", -1, 1)
BROWSER CStr(Arr(1)) & " - " & CStr(Arr(0)), "is_LOAD"
cmdCancel_Click
End Sub

Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
i = 0: sCode = ""
s = "SELECT PK, SupplierCode, SupplierName, Type " & _
    " From dbo.tbl_Inv_Supplier " & _
    " WHERE  (Type = 3) " & _
    " ORDER BY PK"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    i = i + 1
    sCode = CStr(rs!Type) & Format(i, "000#")
    ConnOmega.Execute "UPDATE tbl_Inv_Supplier " & _
                      " SET SupplierCode = '" & sCode & "' " & _
                      " WHERE (PK = " & rs!PK & ")"
    rs.MoveNext
Wend
rs.Close
Screen.MousePointer = vbDefault
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
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "SupplierCodeName", "SuppCodeName", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "SupplierCodeName", "SuppCodeName", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "SupplierCodeName", "SuppCodeName", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "SupplierCodeName", "SuppCodeName", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
cmbSuppType.Clear
cmbSuppType.AddItem "-- Select --"
cmbSuppType.ItemData(cmbSuppType.NewIndex) = 0
s = "SELECT tbl_Inv_SupplierType.* " & _
    " FROM tbl_Inv_SupplierType " & _
    " ORDER BY PK"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbSuppType.AddItem rs!SupplierType
    cmbSuppType.ItemData(cmbSuppType.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
'Me.Caption = "Supplier - Browse"
BROWSER GetSetting(App.EXEName, "SupplierCodeName", "SuppCodeName", ""), "is_LOAD"
If Trim(txtSuppCode.Text) = "" Then BROWSER GetSetting(App.EXEName, "SupplierCodeName", "SuppCodeName", ""), "is_HOME"

tmp = SetWindowLong(txtSuppName.hwnd, GWL_STYLE, GetWindowLong(txtSuppName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtAddress1.hwnd, GWL_STYLE, GetWindowLong(txtAddress1.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtAddress2.hwnd, GWL_STYLE, GetWindowLong(txtAddress2.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtAddress3.hwnd, GWL_STYLE, GetWindowLong(txtAddress3.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtTelNo.hwnd, GWL_STYLE, GetWindowLong(txtTelNo.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtFaxNo.hwnd, GWL_STYLE, GetWindowLong(txtFaxNo.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtEmail.hwnd, GWL_STYLE, GetWindowLong(txtEmail.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtContact.hwnd, GWL_STYLE, GetWindowLong(txtContact.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtSearch.hwnd, GWL_STYLE, GetWindowLong(txtSearch.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picSearch.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub


Private Sub lstResult_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "SupplierCodeName", "SuppCodeName", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "SupplierCodeName", "SuppCodeName", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "SupplierCodeName", "SuppCodeName", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "SupplierCodeName", "SuppCodeName", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":   PRESS_F9
    Case "Close":   PRESS_ESCAPE
    Case Else: Exit Sub
End Select
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then lstResult.Clear: Exit Sub
lstResult.Clear
s = "SELECT tbl_Inv_Supplier.* " & _
    " FROM tbl_Inv_Supplier " & _
    " WHERE (SupplierName LIKE '" & FORMATSQL(Trim(txtSearch.Text)) & "%') " & _
    " ORDER BY SupplierCode + ' - ' + SupplierName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    lstResult.AddItem rs!SupplierCode & " - " & rs!SupplierName
    rs.MoveNext
Wend
rs.Close
If lstResult.ListCount Then lstResult.ListIndex = 0
End Sub

Private Sub txtSearch_GotFocus()
HTEXT txtSearch
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lstResult.SetFocus
If KeyCode = vbKeyDown Then lstResult.SetFocus
End Sub
