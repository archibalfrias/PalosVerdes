VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOperationProShopItemsModel 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picAddEdit 
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2143
      BackColor       =   15396057
      Begin VB.TextBox txtDestination1 
         Height          =   315
         Left            =   2640
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtPK 
         Height          =   315
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtDestination 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   6015
      End
      Begin RPVGCC.b8TitleBar b8TitleBar5 
         Height          =   345
         Left            =   45
         TabIndex        =   4
         Top             =   45
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   609
         Caption         =   "Supply PO Number"
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
         Icon            =   "frmOperationProShopItemsModel.frx":0000
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   6015
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Top             =   5700
      Width           =   7425
      _ExtentX        =   13097
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
      Height          =   5415
      Left            =   120
      ScaleHeight     =   5415
      ScaleWidth      =   7095
      TabIndex        =   7
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton cmdClose 
         Caption         =   "CLOSE"
         Height          =   495
         Left            =   6240
         MouseIcon       =   "frmOperationProShopItemsModel.frx":059A
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "DELETE"
         Height          =   495
         Left            =   6240
         MouseIcon       =   "frmOperationProShopItemsModel.frx":08A4
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "EDIT"
         Height          =   495
         Left            =   6240
         MouseIcon       =   "frmOperationProShopItemsModel.frx":0BAE
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "ADD"
         Height          =   495
         Left            =   6240
         MouseIcon       =   "frmOperationProShopItemsModel.frx":0EB8
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   0
         Width           =   855
      End
      Begin MSComctlLib.ListView lstRecords 
         Height          =   5415
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   9551
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "PK"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "MODEL"
            Object.Width           =   10213
         EndProperty
      End
   End
End
Attribute VB_Name = "frmOperationProShopItemsModel"
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

Dim iRow, x, iFocus, Arr, i

Private Sub LOAD_RECORDS()
With lstRecords.ListItems
    .Clear
    s = "SELECT tbl_Operation_ProShop_Items_Model.* " & _
        " FROM tbl_Operation_ProShop_Items_Model " & _
        " WHERE (Visible = 1) " & _
        " ORDER BY ProItemsModel "
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        Set x = .Add()
        x.Text = rs!PK
        x.SubItems(1) = rs!ProItemsModel
        rs.MoveNext
    Wend
    rs.Close
End With
End Sub


Private Sub b8TitleBar5_CLoseClick()
With lstRecords.ListItems
    If TRANSACTIONTYPE = is_ADDING Then
        If .Count > 1 Then
            .Remove .Count
        Else
            .Item(1).Text = "0"
            .Item(1).SubItems(1) = " "
        End If
        iRow = .Count
    End If
    If TRANSACTIONTYPE = is_EDITTING Then
        .Item(iRow).SubItems(1) = txtDestination1.Text
    End If
    TRANSACTIONTYPE = is_REFRESH
    picAddEdit.Visible = False
    picMain.Enabled = True
    lstRecords.SetFocus
End With
End Sub

Private Sub cmdAdd_Click()
If picAddEdit.Visible = True Then Exit Sub
If AccessRights("Pro Shop Items (Model)", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
With lstRecords.ListItems
    If .Count = 0 Then
        Set x = .Add()
        x.Text = "0"
        x.SubItems(1) = " "
    Else
        If CDbl(.Item(iRow).Text) <> 0 Then
            Set x = .Add()
            x.Text = "0"
            x.SubItems(1) = " "
        End If
    End If
    iRow = .Count
End With
lstRecords.ListItems(iRow).EnsureVisible
lstRecords.ListItems(iRow).Selected = True
txtDestination.Text = ""
picAddEdit.ZOrder 0
b8TitleBar5.Caption = "Add"
TRANSACTIONTYPE = is_ADDING
picAddEdit.Visible = True
txtDestination.SetFocus
End Sub

Private Sub cmdClose_Click()
If picAddEdit.Visible = True Then Exit Sub
Unload Me
End Sub

Private Sub cmdDelete_Click()
If picAddEdit.Visible = True Then Exit Sub
If AccessRights("Pro Shop Items (Model)", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
With lstRecords.ListItems
    If MsgBox("ARE YOU SURE IN DELETING THIS RECORD?                    ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    On Error GoTo PG:
    ConnOmega.Execute "DELETE FROM tbl_Operation_ProShop_Items_Model WHERE (PK = " & .Item(iRow).Text & ")"
    .Remove iRow
    If CDbl(iRow) > .Count Then
        iRow = .Count
    End If
    lstRecords.ListItems(iRow).EnsureVisible
    lstRecords.ListItems(iRow).Selected = True
End With
lstRecords.SetFocus
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub cmdEdit_Click()
If picAddEdit.Visible = True Then Exit Sub
If AccessRights("Pro Shop Items (Model)", "Edit") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
With lstRecords.ListItems
    txtPK.Text = .Item(iRow).Text
    txtDestination.Text = .Item(iRow).SubItems(1)
    txtDestination1.Text = .Item(iRow).SubItems(1)
End With
lstRecords.ListItems(iRow).EnsureVisible
lstRecords.ListItems(iRow).Selected = True
picAddEdit.ZOrder 0
b8TitleBar5.Caption = "Edit"
TRANSACTIONTYPE = is_EDITTING
picAddEdit.Visible = True
txtDestination.SetFocus
End Sub

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert:   cmdAdd_Click
    Case vbKeyF2:       cmdEdit_Click
    Case vbKeyDelete:   cmdDelete_Click
    Case vbKeyEscape:   If picAddEdit.Visible = True Then b8TitleBar5_CLoseClick Else cmdClose_Click
End Select
End Sub

Private Sub Form_Load()
'Save_Loaded_Form Me
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Top = (MainForm.Height - Me.Height) / 6
Me.Left = (MainForm.Width - Me.Width) / 3
Me.Caption = gbl_Form_Caption
iFocus = 0
iRow = 0
LOAD_RECORDS
TRANSACTIONTYPE = is_REFRESH
lstRecords.TabIndex = 0
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
'tmp = SetWindowLong(txtCode.hWnd, GWL_STYLE, GetWindowLong(txtCode.hWnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtDestination.hwnd, GWL_STYLE, GetWindowLong(txtDestination.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picAddEdit.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
'Delete_Loaded_Form Me
End Sub

Private Sub lstRecords_GotFocus()
If lstRecords.ListItems.Count = 0 Then Exit Sub
iRow = lstRecords.SelectedItem.Index
iFocus = 1
u = "SELECT tbl_Operation_ProShop_Items_Model.* " & _
    " FROM tbl_Operation_ProShop_Items_Model " & _
    " WHERE (PK = " & IIf(IsNumeric(lstRecords.ListItems.Item(iRow).Text) = False, 0, lstRecords.ListItems.Item(iRow).Text) & ")"
If ru.State = adStateOpen Then ru.Close
ru.Open u, ConnOmega
If ru.RecordCount > 0 Then
    Statusbar1.Panels(1).Text = ru!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(ru!LastModified), "", ru!LastModified)
End If
ru.Close
End Sub

Private Sub lstRecords_ItemClick(ByVal Item As MSComctlLib.ListItem)
If lstRecords.ListItems.Count = 0 Then Exit Sub
iRow = lstRecords.SelectedItem.Index
u = "SELECT tbl_Operation_ProShop_Items_Model.* " & _
    " FROM tbl_Operation_ProShop_Items_Model " & _
    " WHERE (PK = " & IIf(IsNumeric(lstRecords.ListItems.Item(iRow).Text) = False, 0, lstRecords.ListItems.Item(iRow).Text) & ")"
If ru.State = adStateOpen Then ru.Close
ru.Open u, ConnOmega
If ru.RecordCount > 0 Then
    Statusbar1.Panels(1).Text = ru!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(ru!LastModified), "", ru!LastModified)
End If
ru.Close
End Sub

Private Sub lstRecords_LostFocus()
iFocus = 0
End Sub

Private Sub txtDestination_Change()
With lstRecords.ListItems
    .Item(iRow).SubItems(1) = Trim(txtDestination.Text)
End With
End Sub

Private Sub txtDestination_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    On Error GoTo PG:
    If TRANSACTIONTYPE = is_ADDING Then
        ConnOmega.Execute "INSERT INTO tbl_Operation_ProShop_Items_Model " & _
                         " (ProItemsModel, LastModified) " & _
                         " VALUES ('" & FORMATSQL(Trim(txtDestination.Text)) & "', " & _
                         " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
        t = "SELECT PK " & _
            " FROM tbl_Operation_ProShop_Items_Model " & _
            " WHERE (ProItemsModel = '" & FORMATSQL(Trim(txtDestination.Text)) & "')"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            With lstRecords.ListItems
                .Item(iRow).Text = rt!PK
            End With
        End If
        rt.Close
    End If
    If TRANSACTIONTYPE = is_EDITTING Then
        ConnOmega.Execute "UPDATE tbl_Operation_ProShop_Items_Model " & _
                         " SET ProItemsModel = '" & FORMATSQL(Trim(txtDestination.Text)) & "', " & _
                         " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                         " WHERE (PK = " & txtPK.Text & ")"
    End If
    TRANSACTIONTYPE = is_REFRESH
    picAddEdit.Visible = False
    picMain.Enabled = True
    lstRecords.SetFocus
End If
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub









