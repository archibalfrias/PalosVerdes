VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInvStockTransfer 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvStockTransfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15600
      TabIndex        =   34
      Top             =   0
      Width           =   15600
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   35
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
               Caption         =   " Post   "
               Key             =   "Post"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               Key             =   "Refresh"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         MousePointer    =   99
         MouseIcon       =   "frmInvStockTransfer.frx":08CA
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   9900
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   36
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmInvStockTransfer.frx":0BE4
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
   Begin RPVGCC.b8Container picSLine 
      Height          =   855
      Left            =   1080
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox txtUnit1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4560
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtUnit 
         Height          =   315
         Left            =   6120
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtItemCode1 
         Height          =   285
         Left            =   4080
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtItemDescription1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtQty1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtItemKey1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtItemKey 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7440
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtItemDescription 
         Height          =   315
         Left            =   1440
         TabIndex        =   18
         Top             =   360
         Width           =   4575
      End
      Begin VB.TextBox txtItemCode 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "UNIT"
         Height          =   255
         Left            =   6120
         TabIndex        =   28
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "QTY"
         Height          =   255
         Left            =   7440
         TabIndex        =   21
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM DESCRIPTION"
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM CODE"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Top             =   5580
      Width           =   11415
      _ExtentX        =   20135
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
      Height          =   4215
      Left            =   720
      ScaleHeight     =   4215
      ScaleWidth      =   9855
      TabIndex        =   7
      Top             =   1200
      Width           =   9855
      Begin VB.TextBox txtPostedDate 
         Height          =   315
         Left            =   8520
         TabIndex        =   31
         Top             =   0
         Width           =   1215
      End
      Begin VB.TextBox txtPostedTime 
         Height          =   315
         Left            =   8520
         TabIndex        =   30
         Top             =   360
         Width           =   1215
      End
      Begin MSComctlLib.ListView lstDetail 
         Height          =   2655
         Left            =   0
         TabIndex        =   5
         Top             =   1200
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   4683
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "#"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ItemKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ITEM CODE"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ITEM DESCRIPTION"
            Object.Width           =   9085
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "UNIT"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "QTY"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.ComboBox cmbDestination 
         Height          =   315
         Left            =   3960
         TabIndex        =   4
         Text            =   "cmbDestination"
         Top             =   360
         Width           =   2895
      End
      Begin VB.ComboBox cmbSource 
         Height          =   315
         Left            =   3960
         TabIndex        =   3
         Text            =   "cmbSource"
         Top             =   0
         Width           =   2895
      End
      Begin VB.TextBox txtSTDate 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtSTNumber 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   0
         Width           =   1215
      End
      Begin VB.TextBox txtRemarks 
         Height          =   315
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   2
         Top             =   720
         Width           =   8775
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "POSTED DATE"
         Height          =   255
         Left            =   7320
         TabIndex        =   33
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "POSTED TIME"
         Height          =   255
         Left            =   7320
         TabIndex        =   32
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblTotalQty 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8280
         TabIndex        =   14
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL QTY >>"
         Height          =   255
         Left            =   6960
         TabIndex        =   13
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "DESTINATION"
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "SOURCE"
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ST DATE"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ST #"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "REMARKS"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11280
      Top             =   600
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
            Picture         =   "frmInvStockTransfer.frx":12F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockTransfer.frx":1FD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockTransfer.frx":2CAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockTransfer.frx":3985
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockTransfer.frx":465F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockTransfer.frx":5339
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockTransfer.frx":6013
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockTransfer.frx":6CED
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockTransfer.frx":79C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockTransfer.frx":82A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockTransfer.frx":8F7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockTransfer.frx":9C55
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockTransfer.frx":A92F
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockTransfer.frx":B609
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvStockTransfer.frx":C2E3
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInvStockTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2
Const is_FINDING = 3

Dim TRANS_DETAIL As Long
Const is_DET_REFRESH = 0
Const is_DET_ADDING = 1
Const is_DET_EDITTING = 2

Dim iRow             As Long
Dim isFocus         As Long

Dim iSource         As Long
Dim iDestination    As Long
Dim tmp             As Long

Dim x, sCtrl, iPK, i, iLine, dQty, dAvailableQty

Private Sub BROWSER(Ctrl, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If Ctrl <> "" Then
            s = "SELECT TOP 1 tbl_Inv_StockTransfer.* " & _
                " FROM tbl_Inv_StockTransfer " & _
                " WHERE (STNumber = '" & Ctrl & "')" & _
                " ORDER BY STNumber"
        Else
            s = "SELECT TOP 1 tbl_Inv_StockTransfer.* " & _
                " FROM tbl_Inv_StockTransfer " & _
                " ORDER BY STNumber"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_StockTransfer.* " & _
            " FROM tbl_Inv_StockTransfer " & _
            " ORDER BY STNumber"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_StockTransfer.* " & _
            " FROM tbl_Inv_StockTransfer " & _
            " WHERE (STNumber < '" & Ctrl & "')" & _
            " ORDER BY STNumber DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_StockTransfer.* " & _
            " FROM tbl_Inv_StockTransfer " & _
            " WHERE (STNumber > '" & Ctrl & "')" & _
            " ORDER BY STNumber "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        If picSLine.Visible = True Then Exit Sub
        s = "SELECT TOP 1 tbl_Inv_StockTransfer.* " & _
            " FROM tbl_Inv_StockTransfer " & _
            " ORDER BY STNumber DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtSTNumber.Text = rs!STNumber
    txtSTDate.Text = Format(rs!STDate, "mm/dd/yyyy")
    txtRemarks.Text = rs!Remarks
    iSource = rs!Source
    iDestination = rs!Destination
    cmbSource.Text = ""
    t = "SELECT LocName " & _
        " FROM tbl_Inv_Location " & _
        " WHERE (PK = " & iSource & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        cmbSource.Text = rt!LocName
    End If
    rt.Close
    cmbDestination.Text = ""
    t = "SELECT LocName " & _
        " FROM tbl_Inv_Location " & _
        " WHERE (PK = " & iDestination & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        cmbDestination.Text = rt!LocName
    End If
    rt.Close
    txtPostedDate.Text = ""
    txtPostedTime.Text = ""
    If IsNull(rs!PostedDateTime) = False Then
        txtPostedDate.Text = Format(rs!PostedDateTime, "mm/dd/yyyy")
        txtPostedTime.Text = Format(rs!PostedDateTime, "hh:mm:ss AM/PM")
    End If
    
    lblTotalQty.Caption = "0.00"
    imgPosted.Visible = IIf(rs!Posted = 1, True, False)
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    CLEARDETAIL
    t = "SELECT tbl_Inv_StockTransfer_Detail.ItemKey, " & _
        " tbl_Inv_Items.ItemCode, " & _
        " tbl_Inv_Items.ItemDesc, " & _
        " tbl_Inv_Items.Unit, " & _
        " tbl_Inv_StockTransfer_Detail.Qty " & _
        " FROM tbl_Inv_StockTransfer_Detail LEFT OUTER JOIN " & _
        " tbl_Inv_Items ON tbl_Inv_StockTransfer_Detail.ItemKey = tbl_Inv_Items.PK " & _
        " WHERE (tbl_Inv_StockTransfer_Detail.STKey = " & rs!PK & ") " & _
        " ORDER BY tbl_Inv_StockTransfer_Detail.Line"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        With lstDetail.ListItems
            .Clear
            iLine = 0
            dQty = 0
            While Not rt.EOF
                iLine = iLine + 1
                dQty = dQty + CDbl(rt!Qty)
                Set x = .Add()
                x.Text = ""
                x.SubItems(1) = Format(iLine, "0#")
                x.SubItems(2) = rt!ItemKey
                x.SubItems(3) = rt!ItemCode
                x.SubItems(4) = rt!ItemDesc
                x.SubItems(5) = rt!Unit
                x.SubItems(6) = Format(rt!Qty, "#,##0.00")
                rt.MoveNext
            Wend
        End With
    End If
    rt.Close
    lblTotalQty.Caption = Format(dQty, "#,##0.00")
    SaveSetting App.EXEName, "STCtrlNo", "STCtrl", rs!STNumber
End If
rs.Close
End Sub

Private Sub PRESS_INSERT()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSLine.Visible = True Then Exit Sub
    If AccessRights("Stock Transfer", "Add") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    CLEARTEXT
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_ADDING
    txtSTDate.Text = Format(Date, "mm/dd/yyyy")
    'Me.Caption = "STOCK TRANSFER - NEW"
    txtSTDate.SetFocus
Else
    If picSLine.Visible = True Then Exit Sub
    If isFocus = 0 Then Exit Sub
    If AccessRights("Stock Transfer", "Add") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
    With lstDetail.ListItems
        If CDbl(.Item(.Count).SubItems(2)) = 0 Then
            .Item(.Count).SubItems(1) = Format(.Count, "0#")
            .Item(.Count).SubItems(2) = "0"
            .Item(.Count).SubItems(3) = " "
            .Item(.Count).SubItems(4) = " "
            .Item(.Count).SubItems(5) = " "
            .Item(.Count).SubItems(6) = " "
            iRow = .Count
            txtItemCode.Text = ""
            txtItemDescription.Text = ""
            txtQty.Text = ""
            picToolbar.Enabled = False
            picMain.Enabled = False
            picSLine.ZOrder 0
            picSLine.Visible = True
            TRANS_DETAIL = is_DET_ADDING
            TOOLBARFUNC 3
            txtItemCode.SetFocus
        Else
            Set x = .Add()
            x.Text = ""
            x.SubItems(1) = Format(.Count, "0#")
            x.SubItems(2) = "0"
            x.SubItems(3) = " "
            x.SubItems(4) = " "
            x.SubItems(5) = " "
            x.SubItems(6) = " "
            iRow = .Count
            lstDetail.ListItems(iRow).EnsureVisible
            lstDetail.ListItems(iRow).Selected = True
            txtItemCode.Text = ""
            txtItemDescription.Text = ""
            txtQty.Text = ""
            picToolbar.Enabled = False
            picMain.Enabled = False
            picSLine.ZOrder 0
            picSLine.Visible = True
            TRANS_DETAIL = is_DET_ADDING
            TOOLBARFUNC 3
            txtItemCode.SetFocus
        End If
    End With
End If
End Sub

Private Sub PRESS_F2()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSLine.Visible = True Then Exit Sub
    If Statusbar1.Panels(1).Text = "" Then Exit Sub
    If AccessRights("Stock Transfer", "Edit") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_EDITTING
    'Me.Caption = "STOCK TRANSFER - EDIT"
Else
    If picSLine.Visible = True Then Exit Sub
    If isFocus = 0 Then Exit Sub
    If AccessRights("Stock Transfer", "Edit") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
    With lstDetail.ListItems
        txtItemKey.Text = .Item(iRow).SubItems(2)
        txtItemCode.Text = .Item(iRow).SubItems(3)
        txtItemDescription.Text = .Item(iRow).SubItems(4)
        txtUnit.Text = .Item(iRow).SubItems(5)
        txtQty.Text = .Item(iRow).SubItems(6)
        
        txtItemKey1.Text = .Item(iRow).SubItems(2)
        txtItemCode1.Text = .Item(iRow).SubItems(3)
        txtItemDescription1.Text = .Item(iRow).SubItems(4)
        txtUnit1.Text = .Item(iRow).SubItems(5)
        txtQty1.Text = .Item(iRow).SubItems(6)
        
        picToolbar.Enabled = False
        picMain.Enabled = False
        picSLine.ZOrder 0
        picSLine.Visible = True
        TRANS_DETAIL = is_DET_EDITTING
        TOOLBARFUNC 3
        txtItemCode.SetFocus
    End With
End If
End Sub

Private Sub PRESS_DELETE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSLine.Visible = True Then Exit Sub
    If Statusbar1.Panels(1).Text = "" Then Exit Sub
    If AccessRights("Stock Transfer", "Delete") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
    If MsgBox("ARE YOU SURE IN DELETING THIS TRANSACTION?                       ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    On Error GoTo PG:
    ConnOmega.Execute "DELETE FROM tbl_Inv_StockTransfer WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
    CLEARTEXT
    BROWSER GetSetting(App.EXEName, "STCtrlNo", "STCtrl", ""), "is_PAGEDOWN"
    If Trim(txtSTNumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "STCtrlNo", "STCtrl", ""), "is_HOME"
Else
    If picSLine.Visible = True Then Exit Sub
    If isFocus = 0 Then Exit Sub
    If AccessRights("Stock Transfer", "Delete") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub
    With lstDetail.ListItems
        If .Count = 1 Then
            .Item(.Count).SubItems(1) = " "
            .Item(.Count).SubItems(2) = "0"
            .Item(.Count).SubItems(3) = " "
            .Item(.Count).SubItems(4) = " "
            .Item(.Count).SubItems(5) = " "
            .Item(.Count).SubItems(6) = " "
            iRow = 1
        ElseIf .Count > 1 Then
            .Remove iRow
            If CDbl(iRow) > CDbl(.Count) Then
                iRow = .Count
            End If
        End If
        lstDetail.ListItems(iRow).EnsureVisible
        lstDetail.ListItems(iRow).Selected = True
    End With
End If

Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub

End Sub

Private Sub PRESS_F5()
If IsDate(txtSTDate.Text) = False Then MsgBox "Please Supply a Valid Date!                        ", vbCritical, "Error...": txtSTDate.SetFocus: Exit Sub
If iSource = 0 Then MsgBox "Please Select Source!                           ", vbCritical, "Error...": cmbSource.SetFocus: Exit Sub
If iDestination = 0 Then MsgBox "Please Select Destination!                     ", vbCritical, "Error...": cmbDestination.SetFocus: Exit Sub
txtSTDate.Text = Format(FormatDateTime(txtSTDate.Text, vbShortDate), "mm/dd/yyyy")
On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    sCtrl = ""
    s = "SELECT TOP 1 tbl_Inv_StockTransfer.* " & _
        " FROM tbl_Inv_StockTransfer " & _
        " WHERE (Year(STDate) = " & Format(txtSTDate.Text, "yyyy") & ") " & _
        " ORDER BY STNumber DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        sCtrl = CDbl(rs!STNumber) + 1
    Else
        sCtrl = Format(txtSTDate.Text, "yyyy") & "0000"
    End If
    rs.Close
    Do
        s = "SELECT tbl_Inv_StockTransfer.* " & _
            " FROM tbl_Inv_StockTransfer " & _
            " WHERE (STNumber = '" & sCtrl & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If
        rs.Close
        sCtrl = CDbl(sCtrl) + 1
    Loop
    
    ConnOmega.Execute "INSERT INTO tbl_Inv_StockTransfer " & _
                      " (STNumber, STDate, Remarks, Source, Destination, LastModified) " & _
                      " VALUES ('" & sCtrl & "', '" & FormatDateTime(txtSTDate.Text, vbShortDate) & "', " & _
                      " '" & FORMATSQL(Trim(txtRemarks.Text)) & "', " & cmbSource.ItemData(cmbSource.ListIndex) & ", " & _
                      " " & cmbDestination.ItemData(cmbDestination.ListIndex) & ", '" & CStr(Now) & " - " & gbl_CompleteName & "')"
                      
    iPK = 0
    s = "SELECT tbl_Inv_StockTransfer.* " & _
        " FROM tbl_Inv_StockTransfer " & _
        " WHERE (STNumber = '" & sCtrl & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        iPK = rs!PK
    End If
    rs.Close
    
    If CDbl(iPK) > 0 Then
        iLine = 0
        With lstDetail.ListItems
            For i = 1 To .Count
                If CDbl(.Item(i).SubItems(2)) > 0 Then
                    iLine = iLine + 1
                    ConnOmega.Execute "INSERT INTO tbl_Inv_StockTransfer_Detail " & _
                                      " (STKey, Line, ItemKey, Qty) " & _
                                      " VALUES (" & iPK & ", " & iLine & ", " & _
                                      " " & .Item(i).SubItems(2) & ", " & _
                                      " " & CDbl(.Item(i).SubItems(6)) & ")"
                End If
            Next i
        End With
    End If
End If
If TRANSACTIONTYPE = is_EDITTING Then
    sCtrl = Trim(txtSTNumber.Text)
    iPK = Statusbar1.Panels(1).Text
    
    ConnOmega.Execute "UPDATE tbl_Inv_StockTransfer " & _
                      " SET STDate = '" & FormatDateTime(txtSTDate.Text, vbShortDate) & "', " & _
                      " Remarks = '" & FORMATSQL(Trim(txtRemarks.Text)) & "', " & _
                      " Source = " & cmbSource.ItemData(cmbSource.ListIndex) & ", " & _
                      " Destination = " & cmbDestination.ItemData(cmbDestination.ListIndex) & ", " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & iPK & ")"
    
    ConnOmega.Execute "DELETE FROM tbl_Inv_StockTransfer_Detail WHERE (STKey = " & iPK & ")"
    iLine = 0
    With lstDetail.ListItems
        For i = 1 To .Count
            If CDbl(.Item(i).SubItems(2)) > 0 Then
                iLine = iLine + 1
                ConnOmega.Execute "INSERT INTO tbl_Inv_StockTransfer_Detail " & _
                                  " (STKey, Line, ItemKey, Qty) " & _
                                  " VALUES (" & iPK & ", " & iLine & ", " & _
                                  " " & .Item(i).SubItems(2) & ", " & _
                                  " " & CDbl(.Item(i).SubItems(6)) & ")"
            End If
        Next i
    End With
End If
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANS_DETAIL = is_DET_REFRESH
'Me.Caption = "STOCK TRANSFER - BROWSE"
txtSTNumber.SetFocus
BROWSER sCtrl, "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSLine.Visible = True Then Exit Sub
CLEARTEXT
TOOLBARFUNC 3
TRANSACTIONTYPE = is_FINDING
txtSTNumber.Locked = False
'Me.Caption = "STOCK TRANSFER - FIND"
txtSTNumber.SetFocus
End Sub

Private Sub PRESS_F8()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSLine.Visible = True Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If AccessRights("Stock Transfer", "Post") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
BROWSER GetSetting(App.EXEName, "STCtrlNo", "STCtrl", ""), "is_LOAD"
If imgPosted.Visible = True Then MsgBox "Already Posted!                     ", vbCritical, "Error...": Exit Sub

'   Checking
With lstDetail.ListItems
    For i = 1 To .Count
        If CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2))) <> 0 Then
            dQty = CDbl(IIf(IsNumeric(.Item(i).SubItems(6)) = False, 0, .Item(i).SubItems(6)))
            If CDbl(dQty) > 0 Then
                s = "SELECT tbl_Inv_Items_Available_Location.* " & _
                    " FROM tbl_Inv_Items_Available_Location " & _
                    " WHERE (ItemKey = " & .Item(i).SubItems(2) & ") " & _
                    " AND (LocationKey = " & iSource & ")"
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                If rs.RecordCount > 0 Then
                    If CDbl(rs!Quantity) < CDbl(dQty) Then
                        MsgBox "Not Enough Quantity to Transfer!                        ", vbCritical, "Error..."
                        lstDetail.ListItems(i).EnsureVisible
                        lstDetail.ListItems(i).Selected = True
                        lstDetail.SetFocus
                        If rs.State = adStateOpen Then rs.Close
                        Exit Sub
                    End If
                Else
                    MsgBox "Not Enough Quantity to Transfer!                        ", vbCritical, "Error..."
                    lstDetail.ListItems(i).EnsureVisible
                    lstDetail.ListItems(i).Selected = True
                    lstDetail.SetFocus
                    If rs.State = adStateOpen Then rs.Close
                    Exit Sub
                End If
                rs.Close
            End If
        End If
    Next i
End With


'   Posting

With lstDetail.ListItems
    For i = 1 To .Count
        If CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2))) <> 0 Then
            dQty = CDbl(IIf(IsNumeric(.Item(i).SubItems(6)) = False, 0, .Item(i).SubItems(6)))
            If CDbl(dQty) > 0 Then
                s = "SELECT tbl_Inv_Items_Transaction.*, " & _
                    " QuantityIn - QuantityUsed as AvailableQty " & _
                    " FROM tbl_Inv_Items_Transaction " & _
                    " WHERE (ItemKey = " & .Item(i).SubItems(2) & ") " & _
                    " AND (InOut = 'I') AND (Cleared = 0) " & _
                    " AND (Location = " & iSource & ") " & _
                    " ORDER BY PK"
                If rs.State = adStateOpen Then rs.Close
                rs.Open s, ConnOmega
                If rs.RecordCount > 0 Then
                    While Not rs.EOF
                        dAvailableQty = rs!AvailableQty
                        If CDbl(dAvailableQty) >= CDbl(dQty) Then
                            '  Out
                            ConnOmega.Execute "INSERT INTO tbl_Inv_Items_Transaction " & _
                                              " (ItemKey, Cleared, InOut, DocType, DocNumber, DocDate, " & _
                                              " Location, ReferenceKey, QuantityOut, QuantityUsed, Cost, " & _
                                              " PurcDisc, NetCost, NetVAT) " & _
                                              " VALUES (" & .Item(i).SubItems(2) & ", 1, 'O', 3, " & _
                                              " '" & FORMATSQL(Trim(txtSTNumber.Text)) & "', " & _
                                              " '" & FormatDateTime(txtSTDate.Text, vbShortDate) & "', " & _
                                              " " & iSource & ", " & rs!PK & ", " & CDbl(dQty) & ", " & _
                                              " " & CDbl(dQty) & ", " & CDbl(rs!Cost) & ", " & _
                                              " '" & IIf(IsNull(rs!PurcDisc), "", rs!PurcDisc) & "', " & _
                                              " " & CDbl(rs!NetCost) & ", " & CDbl(rs!NetVAT) & ")"
                            '  In
                            ConnOmega.Execute "INSERT INTO tbl_Inv_Items_Transaction " & _
                                              " (ItemKey, Cleared, InOut, DocType, DocNumber, DocDate, " & _
                                              " Location, ReferenceKey, QuantityIn, Cost, PurcDisc, NetCost, NetVAT) " & _
                                              " VALUES (" & .Item(i).SubItems(2) & ", 0, 'I', 3, " & _
                                              " '" & FORMATSQL(Trim(txtSTNumber.Text)) & "', " & _
                                              " '" & FormatDateTime(txtSTDate.Text, vbShortDate) & "', " & _
                                              " " & iDestination & ", " & rs!PK & ", " & CDbl(dQty) & ", " & _
                                              " " & CDbl(rs!Cost) & ", '" & IIf(IsNull(rs!PurcDisc), "", rs!PurcDisc) & "', " & _
                                              " " & CDbl(rs!NetCost) & ", " & CDbl(rs!NetVAT) & ")"
                            '  Update Reference
                            ConnOmega.Execute "UPDATE tbl_Inv_Items_Transaction " & _
                                              " SET QuantityUsed = QuantityUsed + " & CDbl(dQty) & " " & _
                                              " WHERE (PK = " & rs!PK & ")"
                            GoTo POST_NEXT:
                        Else
                            '  Out
                            ConnOmega.Execute "INSERT INTO tbl_Inv_Items_Transaction " & _
                                              " (ItemKey, Cleared, InOut, DocType, DocNumber, DocDate, " & _
                                              " Location, ReferenceKey, QuantityOut, QuantityUsed, Cost, " & _
                                              " PurcDisc, NetCost, NetVAT) " & _
                                              " VALUES (" & .Item(i).SubItems(2) & ", 1, 'O', 3, " & _
                                              " '" & FORMATSQL(Trim(txtSTNumber.Text)) & "', " & _
                                              " '" & FormatDateTime(txtSTDate.Text, vbShortDate) & "', " & _
                                              " " & iSource & ", " & rs!PK & ", " & CDbl(dAvailableQty) & ", " & _
                                              " " & CDbl(dAvailableQty) & ", " & CDbl(rs!Cost) & ", " & _
                                              " '" & IIf(IsNull(rs!PurcDisc), "", rs!PurcDisc) & "', " & _
                                              " " & CDbl(rs!NetCost) & ", " & CDbl(rs!NetVAT) & ")"
                            '  In
                            ConnOmega.Execute "INSERT INTO tbl_Inv_Items_Transaction " & _
                                              " (ItemKey, Cleared, InOut, DocType, DocNumber, DocDate, " & _
                                              " Location, ReferenceKey, QuantityIn, Cost, PurcDisc, NetCost, NetVAT) " & _
                                              " VALUES (" & .Item(i).SubItems(2) & ", 0, 'I', 3, " & _
                                              " '" & FORMATSQL(Trim(txtSTNumber.Text)) & "', " & _
                                              " '" & FormatDateTime(txtSTDate.Text, vbShortDate) & "', " & _
                                              " " & iDestination & ", " & rs!PK & ", " & CDbl(dAvailableQty) & ", " & _
                                              " " & CDbl(rs!Cost) & ", '" & IIf(IsNull(rs!PurcDisc), "", rs!PurcDisc) & "', " & _
                                              " " & CDbl(rs!NetCost) & ", " & CDbl(rs!NetVAT) & ")"
                            '  Update Reference
                            ConnOmega.Execute "UPDATE tbl_Inv_Items_Transaction " & _
                                              " SET QuantityUsed = QuantityUsed + " & CDbl(dAvailableQty) & " " & _
                                              " WHERE (PK = " & rs!PK & ")"
                            
                            dQty = dQty - CDbl(dAvailableQty)
                            If CDbl(dQty) = 0 Then GoTo POST_NEXT:
                        End If
                        rs.MoveNext
                    Wend
                End If
                rs.Close
            End If
        End If
POST_NEXT:
    Next i
End With

ConnOmega.Execute "UPDATE tbl_Inv_StockTransfer " & _
                  " SET Posted = 1, " & _
                  " PostedDateTime = '" & Now & "', " & _
                  " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                  " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"

BROWSER GetSetting(App.EXEName, "STCtrlNo", "STCtrl", ""), "is_LOAD"

End Sub

Private Sub PRESS_F9()
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picSLine.Visible = True Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub

End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    If picSLine.Visible = True Then Exit Sub
    Unload Me
Else
    If picSLine.Visible = True Then
        If TRANS_DETAIL = is_DET_ADDING Then
            With lstDetail.ListItems
                If .Count = 1 Then
                    .Item(.Count).SubItems(1) = " "
                    .Item(.Count).SubItems(2) = "0"
                    .Item(.Count).SubItems(3) = " "
                    .Item(.Count).SubItems(4) = " "
                    .Item(.Count).SubItems(5) = " "
                    .Item(.Count).SubItems(6) = " "
                ElseIf .Count > 1 Then
                    .Remove .Count
                End If
            End With
        End If
        If TRANS_DETAIL = is_DET_EDITTING Then
            With lstDetail.ListItems
                .Item(iRow).SubItems(2) = txtItemKey1.Text
                .Item(iRow).SubItems(3) = txtItemCode1.Text
                .Item(iRow).SubItems(4) = txtItemDescription1.Text
                .Item(iRow).SubItems(5) = txtUnit1.Text
                .Item(iRow).SubItems(6) = txtQty1.Text
            End With
        End If
        picSLine.Visible = False
        picToolbar.Enabled = True
        picMain.Enabled = True
        lstDetail.SetFocus
        Exit Sub
    End If
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    TRANS_DETAIL = is_DET_REFRESH
    txtSTNumber.SetFocus
    'Me.Caption = "STOCK TRANSFER - BROWSE"
    BROWSER GetSetting(App.EXEName, "STCtrlNo", "STCtrl", ""), "is_LOAD"
    If Trim(txtSTNumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "STCtrlNo", "STCtrl", ""), "is_HOME"
End If

End Sub

Private Sub CLEARTEXT()
iSource = 0
iDestination = 0
txtSTNumber.Text = ""
txtSTDate.Text = ""
txtRemarks.Text = ""
cmbSource.Text = ""
cmbSource.ListIndex = -1
cmbDestination.Text = ""
cmbDestination.ListIndex = -1
txtPostedDate.Text = ""
txtPostedTime.Text = ""
lblTotalQty.Caption = "0.00"
imgPosted.Visible = False
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
CLEARDETAIL
End Sub

Private Sub CLEARDETAIL()
lstDetail.ListItems.Clear
Set x = lstDetail.ListItems.Add()
x.Text = ""
x.SubItems(1) = " "
x.SubItems(2) = "0"
x.SubItems(3) = " "
x.SubItems(4) = " "
x.SubItems(5) = " "
x.SubItems(6) = " "
End Sub

Private Sub LOCKTEXT(bln As Boolean)
txtSTNumber.Locked = True
txtPostedDate.Locked = True
txtPostedTime.Locked = True
txtSTDate.Locked = bln
txtRemarks.Locked = bln
cmbSource.Locked = bln
cmbDestination.Locked = bln
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

Private Sub cmbDestination_Click()
If cmbDestination.ListIndex = -1 Then Exit Sub
iDestination = cmbDestination.ItemData(cmbDestination.ListIndex)
End Sub

Private Sub cmbSource_Click()
If cmbSource.ListIndex = -1 Then Exit Sub
iSource = cmbSource.ItemData(cmbSource.ListIndex)
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
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "STCtrlNo", "STCtrl", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "STCtrlNo", "STCtrl", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "STCtrlNo", "STCtrl", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "STCtrlNo", "STCtrl", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
cmbSource.Clear
s = "SELECT tbl_Inv_Location.* " & _
    " FROM tbl_Inv_Location " & _
    " ORDER BY LocName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbSource.AddItem rs!LocName
    cmbSource.ItemData(cmbSource.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
cmbDestination.Clear
s = "SELECT tbl_Inv_Location.* " & _
    " FROM tbl_Inv_Location " & _
    " ORDER BY LocName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    cmbDestination.AddItem rs!LocName
    cmbDestination.ItemData(cmbDestination.NewIndex) = rs!PK
    rs.MoveNext
Wend
rs.Close
isFocus = 0
iRow = 0
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANS_DETAIL = is_DET_REFRESH
'Me.Caption = "STOCK TRANSFER - BROWSE"
BROWSER GetSetting(App.EXEName, "STCtrlNo", "STCtrl", ""), "is_LOAD"
If Trim(txtSTNumber.Text) = "" Then BROWSER GetSetting(App.EXEName, "STCtrlNo", "STCtrl", ""), "is_HOME"

tmp = SetWindowLong(txtRemarks.hwnd, GWL_STYLE, GetWindowLong(txtRemarks.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picSLine.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstDetail_GotFocus()
iRow = lstDetail.SelectedItem.Index
TRANS_DETAIL = is_DET_REFRESH
isFocus = 1
If TRANSACTIONTYPE = is_REFRESH Then
    If Statusbar1.Panels(1).Text = "" Then Exit Sub
    TRANSACTIONTYPE = is_EDITTING
    'Me.Caption = "STOCK TRANSFER - EDIT"
    BROWSER GetSetting(App.EXEName, "STCtrlNo", "STCtrl", ""), "is_LOAD"
    If imgPosted.Visible = True Then TOOLBARFUNC 3: Exit Sub
End If
With lstDetail.ListItems
    If .Count = 1 Then
        If CDbl(.Item(iRow).SubItems(2)) > 0 Then
            TOOLBARFUNC 5
        Else
            TOOLBARFUNC 4
        End If
    ElseIf .Count > 1 Then
        TOOLBARFUNC 5
    End If
End With
End Sub

Private Sub lstDetail_ItemClick(ByVal Item As MSComctlLib.ListItem)
iRow = lstDetail.SelectedItem.Index
End Sub

Private Sub lstDetail_LostFocus()
isFocus = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "STCtrlNo", "STCtrl", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "STCtrlNo", "STCtrl", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "STCtrlNo", "STCtrl", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "STCtrlNo", "STCtrl", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Print":   PRESS_F9
    Case "Post":    PRESS_F8
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtItemCode_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(3) = txtItemCode.Text
    End With
End If
End Sub

Private Sub txtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If Trim(txtItemCode.Text) = "" Then MsgBox "Please Supply ItemCode!                   ", vbCritical, "Error...": txtItemCode.SetFocus: Exit Sub
    t = "SELECT PK, ItemCode, ItemDesc, Unit " & _
        " FROM tbl_Inv_Items " & _
        " WHERE (ItemCode = '" & Trim(txtItemCode.Text) & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        txtItemKey.Text = rt!PK
        txtItemCode.Text = rt!ItemCode
        txtItemDescription.Text = rt!ItemDesc
        txtUnit.Text = rt!Unit
    Else
        MsgBox "'" & txtItemCode.Text & "' Not Found!                  ", vbCritical, "Error..."
        txtItemCode.SetFocus
        rt.Close
        Exit Sub
    End If
    rt.Close
    txtQty.SetFocus
End If
End Sub

Private Sub txtItemCode_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtItemDescription_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(4) = txtItemDescription.Text
    End With
End If
End Sub

Private Sub txtItemKey_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(2) = RETURNTEXTVALUE(txtItemKey)
    End With
End If
End Sub

Private Sub txtQty_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(6) = Format(RETURNTEXTVALUE(txtQty), "#,##0.00")
    End With
End If
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    picSLine.Visible = False
    picToolbar.Enabled = True
    picMain.Enabled = True
    lstDetail.SetFocus
End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtSTNumber_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If TRANSACTIONTYPE <> is_FINDING Then Exit Sub
    sCtrl = Trim(txtSTNumber.Text)
    s = "SELECT tbl_Inv_StockTransfer.* " & _
        " FROM tbl_Inv_StockTransfer " & _
        " WHERE (STNumber = '" & sCtrl & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount = 0 Then
        MsgBox "'" & Trim(txtSTNumber.Text) & "' not Found!                     ", vbCritical, "Error..."
        rs.Close
        Exit Sub
    End If
    rs.Close
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    TRANS_DETAIL = is_DET_REFRESH
    'Me.Caption = "STOCK TRANSFER - BROWSE"
    BROWSER sCtrl, "is_LOAD"
End If
End Sub

Private Sub txtUnit_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetail.ListItems
        .Item(iRow).SubItems(5) = txtUnit.Text
    End With
End If
End Sub
