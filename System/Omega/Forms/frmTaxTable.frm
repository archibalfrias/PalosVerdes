VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTaxTable 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTaxTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container Picture1 
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox txtEmployee 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5520
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtEmployer 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4080
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtTo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2640
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtFrom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtBracket 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PERCENT"
         Height          =   255
         Left            =   5520
         TabIndex        =   15
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CONSTANT"
         Height          =   255
         Left            =   4080
         TabIndex        =   14
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TO"
         Height          =   255
         Left            =   2640
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FROM"
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BRACKET"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox picBody 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   120
      ScaleHeight     =   3135
      ScaleWidth      =   7455
      TabIndex        =   3
      Top             =   840
      Width           =   7455
      Begin MSComctlLib.ListView lstPHICDetail 
         Height          =   3075
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   5424
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "BRACKET"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "FROM"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "TO"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "CONSTANT"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "PERCENT"
            Object.Width           =   2822
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8280
      Top             =   600
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
            Picture         =   "frmTaxTable.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaxTable.frx":09CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaxTable.frx":0B50
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaxTable.frx":0E6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaxTable.frx":1223
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaxTable.frx":1675
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaxTable.frx":1AC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaxTable.frx":1E7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaxTable.frx":1F91
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaxTable.frx":24D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaxTable.frx":262D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaxTable.frx":2B6F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   770
      Left            =   0
      ScaleHeight     =   765
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
            NumButtons      =   16
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
               Caption         =   "Close"
               Key             =   "Close"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   15000
         Y1              =   690
         Y2              =   690
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
         Y1              =   750
         Y2              =   750
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   4050
      Width           =   7530
      _ExtentX        =   13282
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
End
Attribute VB_Name = "frmTaxTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ROW As Long

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2
Const is_FINDING = 3

Dim x

Private Function LOAD_TABLE()
'Screen.MousePointer = vbHourglass
s = "SELECT tbl_Personnel_TaxTable.Bracket, " & _
    " tbl_Personnel_TaxTable.PFrom, " & _
    " tbl_Personnel_TaxTable.PTo, " & _
    " tbl_Personnel_TaxTable.PConstant, " & _
    " tbl_Personnel_TaxTable.PPercent, " & _
    " tbl_Personnel_TaxTable.PK " & _
    " From tbl_Personnel_TaxTable " & _
    " ORDER BY tbl_Personnel_TaxTable.Bracket"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
With lstPHICDetail.ListItems
    .Clear
    While Not rs.EOF
        Set x = .Add()
        x.Text = rs!PK
        x.SubItems(1) = rs!Bracket
        x.SubItems(2) = Format(rs!PFrom, "##,##0.00")
        x.SubItems(3) = Format(rs!PTo, "##,##0.00")
        x.SubItems(4) = Format(rs!PConstant, "##,##0.00")
        x.SubItems(5) = Format(rs!PPercent, "##,##0.00")
        rs.MoveNext
    Wend
End With
rs.Close
'Screen.MousePointer = vbDefault
End Function

Private Function PRESS_INSERT()

If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
'If AccessRights("Personnel Tax Table", "Add") = False Then
'    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
'           "ACCESS DENIED!                                      ", vbCritical, "Alert"
'    Exit Function
'End If
TRANSACTIONTYPE = is_ADDING
TOOLBARBUTTON False
Picture1.Visible = True
With lstPHICDetail.ListItems
    Set x = .Add()
    x.Text = ""
    x.SubItems(1) = .Count
    x.SubItems(2) = " "
    x.SubItems(3) = " "
    x.SubItems(4) = " "
    x.SubItems(5) = " "
    ROW = .Count
End With
txtBracket.Text = ROW
txtFrom.Text = "0.00"
txtTo.Text = "0.00"
txtEmployer.Text = "0.00'"
txtEmployee.Text = "0.00"
txtFrom.SetFocus
lstPHICDetail.Enabled = False
'Me.Caption = "TAX TABLE - NEW"
End Function

Private Function PRESS_F2()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
'If AccessRights("Personnel Tax Table", "Edit") = False Then
'    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
'           "ACCESS DENIED!                                      ", vbCritical, "Alert"
'    Exit Function
'End If
If ROW = 0 Then Exit Function
TRANSACTIONTYPE = is_EDITTING
TOOLBARBUTTON False
With lstPHICDetail.ListItems
    txtBracket.Text = .Item(ROW).SubItems(1)
    txtFrom.Text = .Item(ROW).SubItems(2)
    txtTo.Text = .Item(ROW).SubItems(3)
    txtEmployer.Text = .Item(ROW).SubItems(4)
    txtEmployee.Text = .Item(ROW).SubItems(5)
End With
Picture1.Visible = True
txtFrom.SetFocus
picBody.Enabled = False
'Me.Caption = "TAX TABLE - EDIT"
End Function

Private Function DELETE_TAXTABLE(intPK)

s = "DELETE FROM Tax_Table" & _
    " WHERE (PK = " & intPK & ")"
ConnOmega.Execute s, , -1
End Function

Private Function PRESS_DELETE()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Function
'If AccessRights("Personnel Tax Table", "Delete") = False Then
'    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
'           "ACCESS DENIED!                                      ", vbCritical, "Alert"
'    Exit Function
'End If
If ROW = 0 Then Exit Function
If MsgBox("ARE YOU SURE TO DELETE THIS RECORD?          ", vbCritical + vbYesNo + vbDefaultButton2, "CONFIRMATION") = vbNo Then Exit Function
DELETE_TAXTABLE lstPHICDetail.ListItems.Item(ROW).Text
LOAD_TABLE
lstPHICDetail.SetFocus
End Function

Private Function INSERT_PHIC(intBracket, dblFrom, _
dblTo, dblEmployer, dblEmployee)

s = "INSERT INTO Tax_Table" & _
    " (Bracket, PFrom, PTo, PConstant, " & _
    " PPercent) " & _
    " VALUES(" & intBracket & ", " & CDbl(dblFrom) & ", " & _
    " " & CDbl(dblTo) & ", " & CDbl(dblEmployer) & ", " & _
    " " & CDbl(dblEmployee) & ")"
ConnOmega.Execute s, , -1
End Function

Private Function UPDATE_PHIC(intPK, intBracket, _
dblFrom, dblTo, dblEmployer, dblEmployee)

s = "UPDATE Tax_Table " & _
    " SET Bracket = " & intBracket & ", " & _
    " PFrom = " & CDbl(dblFrom) & ", " & _
    " PTo = " & CDbl(dblTo) & ", " & _
    " PConstant = " & CDbl(dblEmployer) & ", " & _
    " PPercent = " & CDbl(dblEmployee) & " " & _
    " WHERE(PK = " & intPK & ")"
ConnOmega.Execute s, , -1
End Function

Private Function PRESS_F5()
If TRANSACTIONTYPE = is_ADDING Then
    On Error GoTo PG:
    INSERT_PHIC Trim(txtBracket.Text), Trim(txtFrom.Text), _
        Trim(txtTo.Text), Trim(txtEmployer.Text), _
        Trim(txtEmployee.Text)
    LOAD_TABLE
    picBody.Enabled = True
    TRANSACTIONTYPE = is_REFRESH
    TOOLBARBUTTON True
    Picture1.Visible = False
    lstPHICDetail.SetFocus
    lstPHICDetail.ListItems(ROW).EnsureVisible
    lstPHICDetail.ListItems(ROW).Selected = True
    'Me.Height = 4515
    'Me.Caption = "TAX TABLE - BROWSE"
ElseIf TRANSACTIONTYPE = is_EDITTING Then
    On Error GoTo PG:
    UPDATE_PHIC lstPHICDetail.ListItems.Item(ROW).Text, _
        Trim(txtBracket.Text), Trim(txtFrom.Text), _
        Trim(txtTo.Text), Trim(txtEmployer.Text), _
        Trim(txtEmployee.Text)
    LOAD_TABLE
    picBody.Enabled = True
    TRANSACTIONTYPE = is_REFRESH
    TOOLBARBUTTON True
    Picture1.Visible = False
    lstPHICDetail.SetFocus
    lstPHICDetail.ListItems(ROW).EnsureVisible
    lstPHICDetail.ListItems(ROW).Selected = True
    'Me.Height = 4515
    'Me.Caption = "TAX TABLE - BROWSE"
End If
Exit Function
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Function
End Function

Private Function PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    Unload Me
Else
    If TRANSACTIONTYPE = is_ADDING Then
        ROW = ROW - 1
    End If
    TRANSACTIONTYPE = is_REFRESH
    LOAD_TABLE
    Picture1.Visible = False
    picBody.Enabled = True
    lstPHICDetail.SetFocus
    lstPHICDetail.ListItems(ROW).EnsureVisible
    lstPHICDetail.ListItems(ROW).Selected = True
    TOOLBARBUTTON True
    'Me.Height = 4755
    'Me.Caption = "TAX TABLE - BROWSE"
End If
End Function

Private Sub TOOLBARBUTTON(blnTag As Boolean)
'Set Toolbar1.ImageList = ImageList1
With Toolbar1
    If blnTag Then
        .Buttons(1).Image = 1
        .Buttons(3).Image = 2
        .Buttons(5).Image = 3
        .Buttons(11).Image = 6
        .Buttons(13).Image = 7
        .Buttons(15).Image = 10
'        .Buttons(17).Image = 9
'        .Buttons(19).Image = 10
        .Buttons(1).Enabled = True
        .Buttons(3).Enabled = True
        .Buttons(5).Enabled = True
        .Buttons(7).Image = 4
        .Buttons(7).Caption = "First"
        .Buttons(9).Image = 5
        .Buttons(9).Caption = "Back"
        .Buttons(7).Enabled = True
        .Buttons(9).Enabled = True
        .Buttons(11).Enabled = True
        .Buttons(13).Enabled = True
        .Buttons(15).Enabled = True
'        .Buttons(17).Enabled = True
'        .Buttons(19).Enabled = True
        .Buttons(1).ToolTipText = "NEW (Ins)"
        .Buttons(3).ToolTipText = "EDIT (F2)"
        .Buttons(5).ToolTipText = "DELETE (Del)"
        .Buttons(7).ToolTipText = "FIRST (Home)"
        .Buttons(9).ToolTipText = "BACK (Up)"
        .Buttons(11).ToolTipText = "NEXT (Down)"
        .Buttons(13).ToolTipText = "LAST (End)"
        .Buttons(15).ToolTipText = "CLOSE (Esc)"
'        .Buttons(17).ToolTipText = "PRINT (F9)"
'        .Buttons(19).ToolTipText = "CLOSE (Esc)"
    Else
        .Buttons(1).Image = 1
        .Buttons(3).Image = 2
        .Buttons(5).Image = 3
        .Buttons(11).Image = 6
        .Buttons(13).Image = 7
        .Buttons(15).Image = 10
'        .Buttons(17).Image = 9
'        .Buttons(19).Image = 10
        .Buttons(1).Enabled = False
        .Buttons(3).Enabled = False
        .Buttons(5).Enabled = False
        .Buttons(7).Image = 11
        .Buttons(7).Caption = "Save"
        .Buttons(9).Image = 12
        .Buttons(9).Caption = "Undo"
        .Buttons(7).Enabled = True
        .Buttons(9).Enabled = True
        .Buttons(11).Enabled = False
        .Buttons(13).Enabled = False
        .Buttons(15).Enabled = False
'        .Buttons(17).Enabled = False
'        .Buttons(19).Enabled = False
        .Buttons(1).ToolTipText = ""
        .Buttons(3).ToolTipText = ""
        .Buttons(5).ToolTipText = ""
        .Buttons(7).ToolTipText = "SAVE (F5)"
        .Buttons(9).ToolTipText = "UNDO (Esc)"
        .Buttons(11).ToolTipText = ""
        .Buttons(13).ToolTipText = ""
        .Buttons(15).ToolTipText = ""
'        .Buttons(17).ToolTipText = ""
'        .Buttons(19).ToolTipText = ""
    End If
End With
End Sub

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert: PRESS_INSERT
    Case vbKeyF2:     PRESS_F2
    Case vbKeyDelete: PRESS_DELETE
    Case vbKeyEscape: PRESS_ESCAPE
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
LOAD_TABLE
TRANSACTIONTYPE = is_REFRESH
TOOLBARBUTTON True
'Me.Caption = "TAX TABLE - BROWSE"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If TRANSACTIONTYPE <> is_REFRESH Then
    Cancel = -1
End If
End Sub

Private Sub lstPHICDetail_GotFocus()
ROW = lstPHICDetail.SelectedItem.Index
End Sub

Private Sub lstPHICDetail_ItemClick(ByVal Item As MSComctlLib.ListItem)
ROW = lstPHICDetail.SelectedItem.Index
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":      PRESS_INSERT
    Case "Edit":     PRESS_F2
    Case "Delete":   PRESS_DELETE
    Case "First"
        If Toolbar1.Buttons(7).Caption = "Save" Then
            PRESS_F5
        Else
            lstPHICDetail.ListItems(1).EnsureVisible
            lstPHICDetail.ListItems(1).Selected = True
            ROW = 1
        End If
    Case "Back"
        If Toolbar1.Buttons(9).Caption = "Undo" Then
            PRESS_ESCAPE
        Else
            If ROW > 1 Then
                lstPHICDetail.ListItems(ROW - 1).EnsureVisible
                lstPHICDetail.ListItems(ROW - 1).Selected = True
                ROW = ROW - 1
            End If
        End If
    Case "Next"
        If lstPHICDetail.ListItems.Count > ROW Then
            lstPHICDetail.ListItems(ROW + 1).EnsureVisible
            lstPHICDetail.ListItems(ROW + 1).Selected = True
            ROW = ROW + 1
        End If
    Case "Last"
        lstPHICDetail.ListItems(lstPHICDetail.ListItems.Count).EnsureVisible
        lstPHICDetail.ListItems(lstPHICDetail.ListItems.Count).Selected = True
        ROW = lstPHICDetail.ListItems.Count
    Case "Find":
    Case "Close":  PRESS_ESCAPE
End Select
End Sub

Private Sub txtEmployee_GotFocus()
HTEXT txtEmployee
If TRANSACTIONTYPE = is_ADDING Then
    Toolbar1.Buttons(7).Enabled = False
End If
End Sub

Private Sub txtEmployee_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    PRESS_F5
ElseIf KeyCode = vbKeyBack Then
    If Len(txtEmployee.Text) = 0 Then
        txtEmployer.SetFocus
    End If
End If
End Sub

Private Sub txtEmployee_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtEmployee_LostFocus()
'If TRANSACTIONTYPE = is_ADDING Or _
'TRANSACTIONTYPE = is_EDITTING Then
'    If Trim(txtEmployee.Text) <> "" Then
'        txtEmployee.Text = Format(txtEmployee.Text, "##,##0.00")
'    Else
'        txtEmployee.Text = "0.00"
'    End If
'    With lstPHICDetail.ListItems
'        .Item(ROW).SubItems(5) = txtEmployee.Text
'    End With
'End If
End Sub

Private Sub txtEmployer_GotFocus()
HTEXT txtEmployer
If TRANSACTIONTYPE = is_ADDING Then
    Toolbar1.Buttons(7).Enabled = False
End If
End Sub

Private Sub txtEmployer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtEmployee.SetFocus
ElseIf KeyCode = vbKeyBack Then
    If Len(txtEmployer.Text) = 0 Then
        txtTo.SetFocus
    End If
End If
End Sub

Private Sub txtEmployer_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtEmployer_LostFocus()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If Trim(txtEmployer.Text) <> "" Then
        txtEmployer.Text = Format(txtEmployer.Text, "##,##0.00")
    Else
        txtEmployer.Text = "0.00"
    End If
    With lstPHICDetail.ListItems
        .Item(ROW).SubItems(4) = txtEmployer.Text
    End With
End If
End Sub

Private Sub txtFrom_GotFocus()
HTEXT txtFrom
If TRANSACTIONTYPE = is_ADDING Then
    Toolbar1.Buttons(7).Enabled = False
End If
End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtTo.SetFocus
End If
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtFrom_LostFocus()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If Trim(txtFrom.Text) <> "" Then
        txtFrom.Text = Format(txtFrom.Text, "##,##0.00")
    Else
        txtFrom.Text = "0.00"
    End If
    With lstPHICDetail.ListItems
        .Item(ROW).SubItems(2) = txtFrom.Text
    End With
End If
End Sub

Private Sub txtTo_GotFocus()
HTEXT txtTo
If TRANSACTIONTYPE = is_ADDING Then
    Toolbar1.Buttons(7).Enabled = False
End If
End Sub

Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtEmployer.SetFocus
ElseIf KeyCode = vbKeyBack Then
    If Len(txtTo.Text) = 0 Then
        txtFrom.SetFocus
    End If
End If
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtTo_LostFocus()
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    If Trim(txtTo.Text) <> "" Then
        txtTo.Text = Format(txtTo.Text, "##,##0.00")
    Else
        txtTo.Text = "0.00"
    End If
    With lstPHICDetail.ListItems
        .Item(ROW).SubItems(3) = txtTo.Text
    End With
End If
End Sub







