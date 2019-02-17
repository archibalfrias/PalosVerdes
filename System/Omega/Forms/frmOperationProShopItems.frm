VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOperationProShopItems 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOperationProShopItems.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
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
            Picture         =   "frmOperationProShopItems.frx":08CA
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShopItems.frx":15A4
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShopItems.frx":227E
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShopItems.frx":2F58
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShopItems.frx":3C32
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShopItems.frx":490C
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShopItems.frx":55E6
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShopItems.frx":62C0
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShopItems.frx":6F9A
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShopItems.frx":7874
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShopItems.frx":854E
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShopItems.frx":9228
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShopItems.frx":9F02
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShopItems.frx":ABDC
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperationProShopItems.frx":B8B6
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   10
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   11
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
               Caption         =   "Refresh"
               Key             =   "Refresh"
               ImageKey        =   "IMG12"
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
               ImageKey        =   "IMG13"
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         MousePointer    =   99
         MouseIcon       =   "frmOperationProShopItems.frx":C590
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   11460
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   12
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmOperationProShopItems.frx":C8AA
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
      Height          =   4095
      Left            =   480
      ScaleHeight     =   4095
      ScaleWidth      =   8535
      TabIndex        =   1
      Top             =   1320
      Width           =   8535
      Begin VB.Frame Frame1 
         BackColor       =   &H00C6B8A4&
         Caption         =   "S R P"
         Height          =   1575
         Left            =   1080
         TabIndex        =   21
         Top             =   2520
         Width           =   3495
         Begin MSComctlLib.ListView lstSRP 
            Height          =   1215
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   2143
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Date of Effectivity"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "SRP"
               Object.Width           =   2293
            EndProperty
         End
      End
      Begin VB.ComboBox cmbColor 
         Height          =   315
         Left            =   1080
         TabIndex        =   19
         Text            =   "Combo5"
         Top             =   2160
         Width           =   7455
      End
      Begin VB.ComboBox cmbSizes 
         Height          =   315
         Left            =   1080
         TabIndex        =   17
         Text            =   "Combo4"
         Top             =   1800
         Width           =   7455
      End
      Begin VB.ComboBox cmbModel 
         Height          =   315
         Left            =   1080
         TabIndex        =   15
         Text            =   "Combo3"
         Top             =   1440
         Width           =   7455
      End
      Begin VB.ComboBox cmbBrand 
         Height          =   315
         Left            =   1080
         TabIndex        =   13
         Text            =   "Combo2"
         Top             =   1080
         Width           =   7455
      End
      Begin VB.ComboBox cmbItemType 
         Height          =   315
         Left            =   1080
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   720
         Width           =   7455
      End
      Begin VB.TextBox txtSRP 
         Height          =   315
         Left            =   6600
         MaxLength       =   100
         TabIndex        =   6
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtCode 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   3
         Top             =   0
         Width           =   1455
      End
      Begin VB.TextBox txtDescription 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   2
         Top             =   360
         Width           =   7455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Sizes"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Type"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SRP"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5520
         TabIndex        =   7
         Top             =   2670
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   30
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
      Top             =   5640
      Width           =   9495
      _ExtentX        =   16748
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
Attribute VB_Name = "frmOperationProShopItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Dim tmp As Long

Dim iType, iBrand, iModel, iSizes, iColor, sDescription, x, i, iPK

Private Sub BROWSER(sDesc, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If sDesc <> "" Then
            s = "SELECT TOP 1 tbl_Operation_ProShop_Items.* " & _
                " FROM tbl_Operation_ProShop_Items " & _
                " WHERE (ItemDescription = '" & FORMATSQL(CStr(sDesc)) & "') " & _
                " ORDER BY ItemDescription"
        Else
            s = "SELECT TOP 1 tbl_Operation_ProShop_Items.* " & _
                " FROM tbl_Operation_ProShop_Items " & _
                " ORDER BY ItemDescription"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Operation_ProShop_Items.* " & _
            " FROM tbl_Operation_ProShop_Items " & _
            " ORDER BY ItemDescription"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Operation_ProShop_Items.* " & _
            " FROM tbl_Operation_ProShop_Items " & _
            " WHERE (ItemDescription < '" & FORMATSQL(CStr(sDesc)) & "') " & _
            " ORDER BY ItemDescription DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Operation_ProShop_Items.* " & _
            " FROM tbl_Operation_ProShop_Items " & _
            " WHERE (ItemDescription > '" & FORMATSQL(CStr(sDesc)) & "') " & _
            " ORDER BY ItemDescription"
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Operation_ProShop_Items.* " & _
            " FROM tbl_Operation_ProShop_Items " & _
            " ORDER BY ItemDescription DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtCode.Text = rs!Code
    txtDescription.Text = rs!ItemDescription
    iType = rs!TypeKey
    iBrand = rs!BrandKey
    iModel = rs!ModelKey
    iSizes = rs!SizeKey
    iColor = rs!ColorKey
    cmbItemType.Text = "": cmbItemType.ListIndex = -1
    t = "SELECT tbl_Operation_ProShop_Items_ItemType.* " & _
        " FROM tbl_Operation_ProShop_Items_ItemType " & _
        " WHERE (PK = " & rs!TypeKey & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        cmbItemType.Text = rt!ItemType
    End If
    rt.Close
    cmbBrand.Text = "": cmbBrand.ListIndex = -1
    t = "SELECT tbl_Operation_ProShop_Items_Brand.* " & _
        " FROM tbl_Operation_ProShop_Items_Brand " & _
        " WHERE (PK = " & rs!BrandKey & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        cmbBrand.Text = rt!ProItemsBrand
    End If
    rt.Close
    cmbModel.Text = "": cmbModel.ListIndex = -1
    t = "SELECT tbl_Operation_ProShop_Items_Model.* " & _
        " FROM tbl_Operation_ProShop_Items_Model " & _
        " WHERE (PK = " & rs!ModelKey & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        cmbModel.Text = rt!ProItemsModel
    End If
    rt.Close
    cmbSizes.Text = "": cmbSizes.ListIndex = -1
    t = "SELECT tbl_Operation_ProShop_Items_Sizes.* " & _
        " FROM tbl_Operation_ProShop_Items_Sizes " & _
        " WHERE (PK = " & rs!SizeKey & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        cmbSizes.Text = rt!ProItemsSizes
    End If
    rt.Close
    cmbColor.Text = "": cmbColor.ListIndex = -1
    t = "SELECT tbl_Operation_ProShop_Items_Color.* " & _
        " FROM tbl_Operation_ProShop_Items_Color " & _
        " WHERE (PK = " & rs!ColorKey & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        cmbColor.Text = rt!ProItemsColor
    End If
    rt.Close
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    SaveSetting App.EXEName, "ProShopItems", "ProShopItems", rs!ItemDescription
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If AccessRights("Pro Shop Items", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
CLEARTEXT
LOCKTEXT False
TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("Pro Shop Items", "Edit") = False Then
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
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("Pro Shop Items", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
If MsgBox("ARE YOU SURE IN DELETING THIS RECORD?                    ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
On Error GoTo PG:
ConnOmega.Execute "DELETE FROM tbl_Operation_ProShop_Items WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
CLEARTEXT
BROWSER GetSetting(App.EXEName, "ProShopItems", "ProShopItems", ""), "is_PAGEDOWN"
If Trim(txtCode.Text) = "" Then BROWSER GetSetting(App.EXEName, "ProShopItems", "ProShopItems", ""), "is_HOME"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F5()
If Trim(txtCode.Text) = "" Then MsgBox "Please Supply Item Code!                  ", vbCritical, "Error...": txtCode.SetFocus: Exit Sub
If Trim(txtDescription.Text) = "" Then MsgBox "Please Supply Item Description!                ", vbCritical, "Error...": txtDescription.SetFocus: Exit Sub
If CDbl(iType) = -1 Then MsgBox "Please Select Item Type!                         ", vbCritical, "Error": cmbItemType.SetFocus: Exit Sub
If CDbl(iBrand) = -1 Then MsgBox "Please Select Brand!                        ", vbCritical, "Error...": cmbBrand.SetFocus: Exit Sub
If CDbl(iModel) = -1 Then MsgBox "Please Select Model!                        ", vbCritical, "Error...": cmbModel.SetFocus: Exit Sub
If CDbl(iSizes) = -1 Then MsgBox "Please Select Sizes!                        ", vbCritical, "Error...": cmbSizes.SetFocus: Exit Sub
If CDbl(iColor) = -1 Then MsgBox "Please Select Color!                            ", vbCritical, "Error...": cmbColor.SetFocus: Exit Sub
sDescription = Trim(txtDescription.Text)
iPK = 0
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    ConnOmega.Execute "INSERT INTO tbl_Operation_ProShop_Items " & _
                      " (Code, ItemDescription, TypeKey, BrandKey, ColorKey, ModelKey, SizeKey, LastModified) " & _
                      " VALUES ('" & FORMATSQL(Trim(txtCode.Text)) & "', '" & FORMATSQL(CStr(sDescription)) & "', " & _
                      " " & iType & ", " & iBrand & ", " & iColor & ", " & iModel & ", " & iSizes & ", " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
    
    s = "SELECT PK " & _
        " FROM tbl_Operation_ProShop_Items " & _
        " WHERE (ItemDescription  = '" & FORMATSQL(CStr(sDescription)) & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        iPK = rs!PK
    End If
    rs.Close
End If
If TRANSACTIONTYPE = is_EDITTING Then
    iPK = Statusbar1.Panels(1).Text
    ConnOmega.Execute "UPDATE tbl_Operation_ProShop_Items " & _
                      " SET Code = '" & FORMATSQL(Trim(txtCode.Text)) & "', " & _
                      " ItemDescription = '" & FORMATSQL(CStr(sDescription)) & "', " & _
                      " TypeKey = " & iType & ", BrandKey = " & iBrand & ", " & _
                      " ColorKey = " & iColor & ", ModelKey = " & iModel & ", " & _
                      " SizeKey = " & iSizes & ", LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
End If
If CDbl(iPK) > 0 Then
    
End If
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
BROWSER sDescription, "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    Unload Me
Else
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    BROWSER GetSetting(App.EXEName, "ProShopItems", "ProShopItems", ""), "is_LOAD"
    If Trim(txtCode.Text) = "" Then BROWSER GetSetting(App.EXEName, "ProShopItems", "ProShopItems", ""), "is_HOME"
End If
End Sub

Private Sub CLEARTEXT()
iType = -1
iBrand = -1
iModel = -1
iSizes = -1
iColor = -1
txtCode.Text = ""
txtDescription.Text = ""
txtSRP.Text = ""
cmbItemType.Text = ""
cmbItemType.ListIndex = -1
cmbBrand.Text = ""
cmbBrand.ListIndex = -1
cmbBrand.Text = ""
cmbModel.ListIndex = -1
cmbModel.Text = ""
cmbSizes.ListIndex = -1
cmbSizes.Text = ""
cmbColor.ListIndex = -1
cmbColor.Text = ""
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
CLEAR_DETAILS
End Sub

Private Sub CLEAR_DETAILS()
With lstSRP.ListItems
    .Clear
    Set x = .Add()
    x.Text = ""
    x.SubItems(1) = " "
    x.SubItems(2) = " "
End With
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtCode.Locked = bln
txtDescription.Locked = bln
txtSRP.Locked = bln
cmbItemType.Locked = bln
cmbBrand.Locked = bln
cmbModel.Locked = bln
cmbSizes.Locked = bln
cmbColor.Locked = bln
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

Private Sub cmbBrand_Click()
If cmbBrand.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    iBrand = cmbBrand.ItemData(cmbBrand.ListIndex)
End If
End Sub

Private Sub cmbColor_Click()
If cmbColor.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    iColor = cmbColor.ItemData(cmbColor.ListIndex)
End If
End Sub

Private Sub cmbItemType_Click()
If cmbItemType.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    iType = cmbItemType.ItemData(cmbItemType.ListIndex)
End If
End Sub

Private Sub cmbModel_Click()
If cmbModel.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    iModel = cmbModel.ItemData(cmbModel.ListIndex)
End If
End Sub

Private Sub cmbSizes_Click()
If cmbSizes.ListIndex = -1 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    iSizes = cmbSizes.ItemData(cmbSizes.ListIndex)
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert:   PRESS_INSERT
    Case vbKeyF2:       PRESS_F2
    Case vbKeyDelete:   PRESS_DELETE
    Case vbKeyF5:       PRESS_F5
    Case vbKeyF6:       PRESS_F6
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "ProShopItems", "ProShopItems", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "ProShopItems", "ProShopItems", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "ProShopItems", "ProShopItems", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "ProShopItems", "ProShopItems", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.Height - Me.Height) / 3
Me.Left = (MainForm.Width - Me.Width) / 5
POPULATE_COMBO "PK", "ItemType", "tbl_Operation_ProShop_Items_ItemType", "ItemType", cmbItemType
POPULATE_COMBO "PK", "ProItemsBrand", "tbl_Operation_ProShop_Items_Brand", "ProItemsBrand", cmbBrand
POPULATE_COMBO "PK", "ProItemsModel", "tbl_Operation_ProShop_Items_Model", "ProItemsModel", cmbModel
POPULATE_COMBO "PK", "ProItemsSizes", "tbl_Operation_ProShop_Items_Sizes", "ProItemsSizes", cmbSizes
POPULATE_COMBO "PK", "ProItemsColor", "tbl_Operation_ProShop_Items_Color", "ProItemsColor", cmbColor
'iRow = 0
'iFocus = 0
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH

BROWSER GetSetting(App.EXEName, "ProShopItems", "ProShopItems", ""), "is_LOAD"
If Trim(txtDescription.Text) = "" Then BROWSER GetSetting(App.EXEName, "ProShopItems", "ProShopItems", ""), "is_HOME"

tmp = SetWindowLong(txtDescription.hwnd, GWL_STYLE, GetWindowLong(txtDescription.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtCode.hwnd, GWL_STYLE, GetWindowLong(txtCode.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "ProShopItems", "ProShopItems", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "ProShopItems", "ProShopItems", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "ProShopItems", "ProShopItems", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "ProShopItems", "ProShopItems", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":
    Case "Refresh":
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtSRP_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub
