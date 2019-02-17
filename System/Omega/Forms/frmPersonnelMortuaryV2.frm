VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonnelMortuaryV2 
   Appearance      =   0  'Flat
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11475
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
   ScaleHeight     =   4905
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picSLPerPeriod 
      Height          =   855
      Left            =   6000
      TabIndex        =   20
      Top             =   1320
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox txtPerPayroll1 
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtPerPayroll 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1680
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtPayrollDate 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1680
         TabIndex        =   25
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Period"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   2640
      ScaleHeight     =   2415
      ScaleWidth      =   6135
      TabIndex        =   0
      Top             =   1440
      Width           =   6135
      Begin VB.TextBox txtCtrl 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtYear 
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   0
         Width           =   975
      End
      Begin VB.ComboBox cmbMonth 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   0
         Width           =   1335
      End
      Begin VB.TextBox txtRemarks 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   4575
      End
      Begin MSComctlLib.ListView lstDetails 
         Height          =   1575
         Left            =   0
         TabIndex        =   4
         Top             =   840
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   2778
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
            Text            =   "PostLevelKey"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Position Level"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView lstPerPayroll 
         Height          =   1575
         Left            =   3480
         TabIndex        =   18
         Top             =   840
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2778
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
            Text            =   "Payroll Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Month / Year"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
   End
   Begin RPVGCC.b8Container picSLLines 
      Height          =   855
      Left            =   2400
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1508
      BackColor       =   8438015
      Begin VB.TextBox txtPositionLevel 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2640
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtAmount1 
         Height          =   315
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Position Level"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2640
         TabIndex        =   14
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   7
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   8
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
         MouseIcon       =   "frmPersonnelMortuaryV2.frx":0000
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   9900
            ScaleHeight     =   495
            ScaleWidth      =   2055
            TabIndex        =   9
            Top             =   120
            Width           =   2055
            Begin VB.Image imgPosted 
               Height          =   345
               Left            =   0
               Picture         =   "frmPersonnelMortuaryV2.frx":031A
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
      Left            =   10440
      Top             =   2400
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
            Picture         =   "frmPersonnelMortuaryV2.frx":0A2D
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelMortuaryV2.frx":1707
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelMortuaryV2.frx":23E1
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelMortuaryV2.frx":30BB
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelMortuaryV2.frx":3D95
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelMortuaryV2.frx":4A6F
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelMortuaryV2.frx":5749
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelMortuaryV2.frx":6423
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelMortuaryV2.frx":70FD
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelMortuaryV2.frx":79D7
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelMortuaryV2.frx":86B1
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelMortuaryV2.frx":938B
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelMortuaryV2.frx":A065
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelMortuaryV2.frx":AD3F
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonnelMortuaryV2.frx":BA19
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   10
      Top             =   4605
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   2469
            MinWidth        =   2469
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
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
   Begin MSComctlLib.ListView lstMortuary 
      Height          =   1575
      Left            =   120
      TabIndex        =   19
      Top             =   2880
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2778
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "PostLevelKey"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Position Level"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Amount"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "PayrollDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmPersonnelMortuaryV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmp As Long

Public TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1
Const is_EDITTING = 2

Private TRANS_DETAIL As Long
Const is_DET_REFRESH = 0
Const is_DET_ADDING = 1
Const is_DET_EDITTING = 2

Dim isFocus, iRow       As Long
Dim isFocusDed, iRowDed       As Long

Dim x, y, i, j, sCtrl, iPK
Dim iPeriod, dDate, iLineCnt

Private Sub PRESS_INSERT()
If picSLPerPeriod.Visible = True Then Exit Sub
If picSLLines.Visible = True Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If AccessRights("Mortuary", "Add") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
CLEARTEXT
LOCKTEXT False
'TOOLBARFUNC 2
TRANSACTIONTYPE = is_ADDING
cmbMonth.ListIndex = Month(Date) - 1
txtYear.Text = Format(Date, "yyyy")
s = "SELECT tbl_Personnel_Position_Level.* " & _
    " FROM tbl_Personnel_Position_Level " & _
    " ORDER BY PK"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
lstDetails.ListItems.Clear
While Not rs.EOF
    Set x = lstDetails.ListItems.Add()
    x.Text = rs!PK
    x.SubItems(1) = rs!LevelName
    x.SubItems(2) = " "
    For i = 1 To 2
        Set y = lstMortuary.ListItems.Add()
        y.Text = rs!PK
        y.SubItems(1) = rs!LevelName
        y.SubItems(2) = " "
        If i = 1 Then
            y.SubItems(3) = Format(DateSerial(RETURNTEXTVALUE(txtYear), cmbMonth.ListIndex + 1, 15), "mm/dd/yyyy")
        Else
            y.SubItems(3) = Format(DateSerial(RETURNTEXTVALUE(txtYear), (cmbMonth.ListIndex + 1) + 1, 0), "mm/dd/yyyy")
        End If
        y.SubItems(4) = " "
    Next i
    rs.MoveNext
Wend
rs.Close
iRow = 1
lstDetails_Click
TOOLBARFUNC 2
cmbMonth.SetFocus
End Sub

Private Sub PRESS_F2()
If picSLPerPeriod.Visible = True Then Exit Sub
If picSLLines.Visible = True Then Exit Sub
If TRANSACTIONTYPE = is_REFRESH Then
    If Statusbar1.Panels(1).Text = "" Then Exit Sub
    If imgPosted.Visible = True Then MsgBox "Already Posted!                 ", vbCritical, "Error...": Exit Sub
    If AccessRights("Mortuary", "Edit") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    LOCKTEXT False
    TOOLBARFUNC 2
    TRANSACTIONTYPE = is_EDITTING
Else
    If isFocus = 1 Then
        With lstDetails.ListItems
            txtPositionLevel.Text = .Item(iRow).SubItems(1)
            txtAmount.Text = .Item(iRow).SubItems(2)
            txtAmount1.Text = .Item(iRow).SubItems(2)
        End With
        picSLLines.ZOrder 0
        picMain.Enabled = False
        picToolbar.Enabled = False
        picSLLines.Visible = True
        TRANS_DETAIL = is_DET_EDITTING
        txtAmount.SetFocus
        Exit Sub
    End If
    If isFocusDed = 1 Then
        With lstPerPayroll.ListItems
            txtPayrollDate.Text = .Item(iRowDed).SubItems(1)
            txtPerPayroll.Text = .Item(iRowDed).SubItems(2)
            txtPerPayroll1.Text = .Item(iRowDed).SubItems(2)
        End With
        picSLPerPeriod.ZOrder 0
        picMain.Enabled = False
        picToolbar.Enabled = False
        picSLPerPeriod.Visible = True
        TRANS_DETAIL = is_DET_EDITTING
        txtPerPayroll.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub PRESS_DELETE()
If picSLPerPeriod.Visible = True Then Exit Sub
If picSLLines.Visible = True Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If AccessRights("Mortuary", "Delete") = False Then
    MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
           "ACCESS DENIED!                                      ", vbCritical, "Alert"
    Exit Sub
End If
End Sub

Private Sub PRESS_F5()
If picSLPerPeriod.Visible = True Then Exit Sub
If picSLLines.Visible = True Then Exit Sub
If cmbMonth.ListIndex = -1 Then MsgBox "Please select month!                 ", vbCritical, "Error...": cmbMonth.SetFocus: Exit Sub
If RETURNTEXTVALUE(txtYear) <= 0 Then MsgBox "Invalid year!                  ", vbCritical, "Error...": txtYear.SetFocus: Exit Sub
With lstDetails.ListItems
    For i = 1 To .Count
        If CDbl(IIf(IsNumeric(.Item(i).SubItems(2)) = False, 0, .Item(i).SubItems(2))) <= 0 Then MsgBox "Invalid Amount!                ", vbCritical, "Error...": Exit Sub
        'If CDbl(IIf(IsNumeric(.Item(i).SubItems(4)) = False, 0, .Item(i).SubItems(4))) <= 0 Then MsgBox "Invalid Amount!                ", vbCritical, "Error...": Exit Sub
    Next i
End With

With lstMortuary.ListItems
    For i = 1 To .Count
        'If CDbl(IIf(IsNumeric(.Item(i).SubItems(3)) = False, 0, .Item(i).SubItems(3))) <= 0 Then MsgBox "Invalid Amount!                ", vbCritical, "Error...": Exit Sub
        If CDbl(IIf(IsNumeric(.Item(i).SubItems(4)) = False, 0, .Item(i).SubItems(4))) <= 0 Then MsgBox "Invalid Amount!                ", vbCritical, "Error...": Exit Sub
    Next i
End With

On Error GoTo PG:
If TRANSACTIONTYPE = is_ADDING Then
    sCtrl = ""
    s = "SELECT TOP (1) Ctrl " & _
        " FROM tbl_Personnel_Mortuary " & _
        " WHERE (Mor_Year = " & RETURNTEXTVALUE(txtYear) & ")" & _
        " ORDER BY Ctrl DESC"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        sCtrl = Format(CDbl(rs!Ctrl) + 1, "0000000#")
    Else
        sCtrl = Format(RETURNTEXTVALUE(txtYear), "000#") & "0000"
    End If
    rs.Close
    
    Do
        s = "SELECT tbl_Personnel_Mortuary.* " & _
            " FROM tbl_Personnel_Mortuary " & _
            " WHERE (Ctrl = '" & sCtrl & "')"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            rs.Close
            Exit Do
        End If
        rs.Close
        sCtrl = Format(CDbl(sCtrl) + 1, "0000000#")
    Loop
    
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Mortuary " & _
                      " (Ctrl, Mor_Month, Mor_Year, Remarks, LastModified) " & _
                      " VALUES ('" & sCtrl & "', " & cmbMonth.ListIndex + 1 & ", " & _
                      " " & RETURNTEXTVALUE(txtYear) & ", '" & FORMATSQL(Trim(txtRemarks.Text)) & "', " & _
                      " '" & CStr(Now) & " - " & gbl_CompleteName & "')"
    iPK = 0
    s = "SELECT tbl_Personnel_Mortuary.* " & _
        " FROM tbl_Personnel_Mortuary " & _
        " WHERE (Ctrl = '" & sCtrl & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        iPK = rs!PK
    End If
    rs.Close
    
End If
If TRANSACTIONTYPE = is_EDITTING Then
    sCtrl = Trim(txtCtrl.Text)
    iPK = Statusbar1.Panels(1).Text
    
    ConnOmega.Execute "UPDATE tbl_Personnel_Mortuary " & _
                      " SET Mor_Month = " & cmbMonth.ListIndex + 1 & ", " & _
                      " Mor_Year = " & RETURNTEXTVALUE(txtYear) & ", " & _
                      " Remarks = '" & FORMATSQL(Trim(txtRemarks.Text)) & "', " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & iPK & ")"
    
End If

If CDbl(iPK) <> 0 Then
    ConnOmega.Execute "DELETE FROM tbl_Personnel_Mortuary_Det WHERE (MasterKey = " & iPK & ")"
    With lstDetails.ListItems
        For i = 1 To .Count
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Mortuary_Det " & _
                              " (MasterKey, PositionLevelKey, Amount) " & _
                              " VALUES (" & iPK & ", " & .Item(i).Text & ",  " & _
                              " " & CDbl(.Item(i).SubItems(2)) & ")"
        Next i
    End With
    With lstMortuary.ListItems
        For j = 1 To .Count
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Mortuary_Det_Det " & _
                              " (MasterKey, PositionLevelKey, PayrollDate, Amount) " & _
                              " VALUES (" & iPK & ", " & .Item(j).Text & ", " & _
                              " '" & FormatDateTime(.Item(j).SubItems(3), vbShortDate) & "', " & _
                              " " & CDbl(.Item(j).SubItems(4)) & ")"
        Next j
    End With
End If

CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANS_DETAIL = is_DET_REFRESH
BROWSER sCtrl, "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F6()
If picSLPerPeriod.Visible = True Then Exit Sub
If picSLLines.Visible = True Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
End Sub

Private Sub PRESS_F8()
If picSLPerPeriod.Visible = True Then Exit Sub
If picSLLines.Visible = True Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If imgPosted.Visible = False Then
    If AccessRights("Mortuary", "Post") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If MsgBox("ARE YOU SURE IN POSTING THIS TRANSACTION?                   ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    ConnOmega.Execute "UPDATE tbl_Personnel_Mortuary " & _
                      " SET Posted = 1, " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
End If
If imgPosted.Visible = True Then
    If AccessRights("Mortuary", "UnPost") = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Sub
    End If
    If MsgBox("ARE YOU SURE IN UNPOSTING THIS TRANSACTION?                   ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
    ConnOmega.Execute "UPDATE tbl_Personnel_Mortuary " & _
                      " SET Posted = 0, " & _
                      " LastModified = '" & CStr(Now) & " - " & gbl_CompleteName & "' " & _
                      " WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
End If
CLEARTEXT
BROWSER GetSetting(App.EXEName, "PersonnelMortuary", "PersonnelMortuary", ""), "is_LOAD"
Exit Sub
PG:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub PRESS_F9()
If picSLPerPeriod.Visible = True Then Exit Sub
If picSLLines.Visible = True Then Exit Sub
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
End Sub

Private Sub PRESS_ESCAPE()
If TRANSACTIONTYPE = is_REFRESH Then
    Unload Me
Else
    If picSLLines.Visible = True Then
        lstDetails.ListItems.Item(iRow).SubItems(2) = txtAmount1.Text
        picSLLines.Visible = False
        picMain.Enabled = True
        picToolbar.Enabled = True
        lstDetails.SetFocus
        Exit Sub
    End If
    If picSLPerPeriod.Visible = True Then
        lstPerPayroll.ListItems.Item(iRow).SubItems(2) = txtPerPayroll1.Text
        picSLPerPeriod.Visible = False
        picMain.Enabled = True
        picToolbar.Enabled = True
        lstPerPayroll.SetFocus
        Exit Sub
    End If
    CLEARTEXT
    LOCKTEXT True
    TOOLBARFUNC 1
    TRANSACTIONTYPE = is_REFRESH
    TRANS_DETAIL = is_DET_REFRESH
    BROWSER GetSetting(App.EXEName, "PersonnelMortuary", "PersonnelMortuary", ""), "is_LOAD"
    If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelMortuary", "PersonnelMortuary", ""), "is_HOME"
End If
End Sub

Private Sub BROWSER(Ctrl, isAction As String)
Select Case isAction
    Case "is_LOAD"
        If Ctrl <> "" Then
            s = "SELECT TOP (1) tbl_Personnel_Mortuary.* " & _
                " FROM tbl_Personnel_Mortuary " & _
                " WHERE (Ctrl = '" & sCtrl & "') " & _
                " ORDER BY Ctrl"
        Else
            s = "SELECT TOP (1) tbl_Personnel_Mortuary.* " & _
                " FROM tbl_Personnel_Mortuary " & _
                " ORDER BY Ctrl"
        End If
    Case "is_HOME"
        If picSLPerPeriod.Visible = True Then Exit Sub
        If picSLLines.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) tbl_Personnel_Mortuary.* " & _
            " FROM tbl_Personnel_Mortuary " & _
            " ORDER BY Ctrl"
    Case "is_PAGEUP"
        If picSLPerPeriod.Visible = True Then Exit Sub
        If picSLLines.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) tbl_Personnel_Mortuary.* " & _
            " FROM tbl_Personnel_Mortuary " & _
            " WHERE (Ctrl < '" & sCtrl & "') " & _
            " ORDER BY Ctrl DESC"
    Case "is_PAGEDOWN"
        If picSLPerPeriod.Visible = True Then Exit Sub
        If picSLLines.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) tbl_Personnel_Mortuary.* " & _
            " FROM tbl_Personnel_Mortuary " & _
            " WHERE (Ctrl > '" & sCtrl & "') " & _
            " ORDER BY Ctrl"
    Case "is_END"
        If picSLPerPeriod.Visible = True Then Exit Sub
        If picSLLines.Visible = True Then Exit Sub
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP (1) tbl_Personnel_Mortuary.* " & _
            " FROM tbl_Personnel_Mortuary " & _
            " ORDER BY Ctrl DESC"
    Case "is_FIND"
        s = "SELECT TOP (1) tbl_Personnel_Mortuary.* " & _
            " FROM tbl_Personnel_Mortuary " & _
            " WHERE (PK = " & sCtrl & ") " & _
            " ORDER BY Ctrl DESC"
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    txtCtrl.Text = rs!Ctrl
    cmbMonth.ListIndex = rs!Mor_Month - 1
    txtYear.Text = rs!Mor_Year
    txtRemarks.Text = rs!Remarks
    imgPosted.Visible = False
    Statusbar1.Panels(1).Text = rs!PK
    Statusbar1.Panels(2).Text = IIf(IsNull(rs!LastModified), "", rs!LastModified)
    
    imgPosted.Visible = IIf(rs!Posted = 1, True, False)
    Toolbar1.Buttons(19).Caption = IIf(rs!Posted = 1, "UnPost", " Post ")
    Toolbar1.Buttons(19).Image = IIf(rs!Posted = 1, 11, 10)
    
    CLEAR_Details
    CLEAR_Details_Payroll
    lstMortuary.ListItems.Clear
    t = "SELECT dbo.tbl_Personnel_Mortuary_Det.PositionLevelKey, dbo.tbl_Personnel_Position_Level.LevelName, " & _
        " dbo.tbl_Personnel_Mortuary_Det.Amount " & _
        " FROM  dbo.tbl_Personnel_Mortuary_Det LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Position_Level ON dbo.tbl_Personnel_Mortuary_Det.PositionLevelKey = dbo.tbl_Personnel_Position_Level.PK " & _
        " Where (dbo.tbl_Personnel_Mortuary_Det.MasterKey = " & rs!PK & ") " & _
        " ORDER BY dbo.tbl_Personnel_Mortuary_Det.PositionLevelKey"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lstDetails.ListItems.Clear
        While Not rt.EOF
            Set x = lstDetails.ListItems.Add()
            x.Text = rt!PositionLevelKey
            x.SubItems(1) = rt!LevelName
            x.SubItems(2) = Format(rt!Amount, "#,##0.00")
            rt.MoveNext
        Wend
    End If
    
    t = "SELECT dbo.tbl_Personnel_Mortuary_Det.MasterKey, dbo.tbl_Personnel_Mortuary_Det.PositionLevelKey, " & _
        " dbo.tbl_Personnel_Position_Level.LevelName, dbo.tbl_Personnel_Mortuary_Det.Amount, " & _
        " dbo.tbl_Personnel_Mortuary_Det_Det.PayrollDate, dbo.tbl_Personnel_Mortuary_Det_Det.Amount AS perPayroll " & _
        " FROM  dbo.tbl_Personnel_Mortuary_Det LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Position_Level ON dbo.tbl_Personnel_Mortuary_Det.PositionLevelKey = dbo.tbl_Personnel_Position_Level.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Mortuary_Det_Det ON dbo.tbl_Personnel_Mortuary_Det.MasterKey = dbo.tbl_Personnel_Mortuary_Det_Det.MasterKey AND dbo.tbl_Personnel_Mortuary_Det.PositionLevelKey = dbo.tbl_Personnel_Mortuary_Det_Det.PositionLevelKey " & _
        " WHERE (dbo.tbl_Personnel_Mortuary_Det.MasterKey = " & rs!PK & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    While Not rt.EOF
        Set y = lstMortuary.ListItems.Add()
        y.Text = rt!PositionLevelKey
        y.SubItems(1) = rt!LevelName
        y.SubItems(2) = Format(rt!Amount, "#,##0.00")
        y.SubItems(3) = Format(rt!PayrollDate, "mm/dd/yyyy")
        y.SubItems(4) = Format(rt!perPayroll, "#,##0.00")
        rt.MoveNext
    Wend
    rt.Close
    iRow = 1
    lstDetails_Click
    SaveSetting App.EXEName, "PersonnelMortuary", "PersonnelMortuary", rs!Ctrl
    
End If
rs.Close
End Sub


Private Sub CLEARTEXT()
txtCtrl.Text = ""
cmbMonth.ListIndex = -1
txtYear.Text = ""
txtRemarks.Text = ""
imgPosted.Visible = False
Statusbar1.Panels(1).Text = ""
Statusbar1.Panels(2).Text = ""
CLEAR_Details
CLEAR_Details_Payroll
lstMortuary.ListItems.Clear
End Sub

Private Sub CLEAR_Details()
With lstDetails.ListItems
    .Clear
    Set x = .Add()
    x.Text = "0"
    x.SubItems(1) = " "
    x.SubItems(2) = " "
End With
End Sub

Private Sub CLEAR_Details_Payroll()
With lstPerPayroll.ListItems
    .Clear
    Set x = .Add()
    x.Text = "0"
    x.SubItems(1) = " "
    x.SubItems(2) = " "
End With
End Sub

Private Sub LOCKTEXT(bln As Boolean)
cmbMonth.Locked = bln
txtYear.Locked = bln
txtRemarks.Locked = bln
txtPayrollDate.Locked = True
txtPositionLevel.Locked = True
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
            .Buttons(1).Enabled = False
            .Buttons(3).Enabled = True
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
            .Buttons(3).ToolTipText = "EDIT (F2)"
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



Private Sub cmbMonth_Click()
If cmbMonth.ListIndex = -1 Then Exit Sub
If RETURNTEXTVALUE(txtYear) <= 0 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    With lstDetails.ListItems
        For i = 1 To .Count
            iPeriod = 0
            For j = 1 To lstMortuary.ListItems.Count
                If CDbl(.Item(i).Text) = CDbl(lstMortuary.ListItems.Item(j).Text) Then
                    iPeriod = iPeriod + 1
                    If iPeriod = 1 Then
                        dDate = DateSerial(RETURNTEXTVALUE(txtYear), cmbMonth.ListIndex + 1, 15)
                    Else
                        dDate = DateSerial(RETURNTEXTVALUE(txtYear), (cmbMonth.ListIndex + 1) + 1, 0)
                    End If
                    lstMortuary.ListItems.Item(j).SubItems(3) = Format(dDate, "mm/dd/yyyy")
                End If
            Next j
        Next i
    End With
End If
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
    Case vbKeyF8:       PRESS_F8
    Case vbKeyF9:       PRESS_F9
    Case vbKeyEscape:   PRESS_ESCAPE
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "PersonnelMortuary", "PersonnelMortuary", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "PersonnelMortuary", "PersonnelMortuary", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "PersonnelMortuary", "PersonnelMortuary", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "PersonnelMortuary", "PersonnelMortuary", ""), "is_END"
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
Me.Top = (MainForm.ScaleHeight - Me.Height) / 3
Me.Left = (MainForm.ScaleWidth - Me.Width) / 3
With cmbMonth
    .Clear
    .AddItem "January"
    .AddItem "February"
    .AddItem "March"
    .AddItem "April"
    .AddItem "May"
    .AddItem "June"
    .AddItem "July"
    .AddItem "August"
    .AddItem "September"
    .AddItem "October"
    .AddItem "November"
    .AddItem "December"
End With
isFocus = 0
iRow = 0
isFocusDed = 0
iRowDed = 0
CLEARTEXT
LOCKTEXT True
TOOLBARFUNC 1
TRANSACTIONTYPE = is_REFRESH
TRANS_DETAIL = is_DET_REFRESH
BROWSER GetSetting(App.EXEName, "PersonnelMortuary", "PersonnelMortuary", ""), "is_LOAD"
If Trim(txtCtrl.Text) = "" Then BROWSER GetSetting(App.EXEName, "PersonnelMortuary", "PersonnelMortuary", ""), "is_HOME"
tmp = SetWindowLong(txtRemarks.hwnd, GWL_STYLE, GetWindowLong(txtRemarks.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If picSLLines.Visible = True Then Cancel = -1
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub lstDetails_Click()
isFocus = 1
iRow = lstDetails.SelectedItem.Index
CLEAR_Details_Payroll
If CDbl(lstDetails.ListItems.Item(iRow).Text) <> 0 Then
    iLineCnt = 0
    For i = 1 To lstMortuary.ListItems.Count
        If CDbl(lstDetails.ListItems.Item(iRow).Text) = CDbl(lstMortuary.ListItems.Item(i).Text) Then
            iLineCnt = iLineCnt + 1
        End If
    Next i
    
    If CDbl(iLineCnt) > 0 Then
        lstPerPayroll.ListItems.Clear
        For i = 1 To lstMortuary.ListItems.Count
            If CDbl(lstDetails.ListItems.Item(iRow).Text) = CDbl(lstMortuary.ListItems.Item(i).Text) Then
                Set x = lstPerPayroll.ListItems.Add()
                x.Text = lstMortuary.ListItems.Item(i).Text
                x.SubItems(1) = lstMortuary.ListItems.Item(i).SubItems(3)
                x.SubItems(2) = lstMortuary.ListItems.Item(i).SubItems(4)
            End If
        Next i
    End If
End If

TRANS_DETAIL = is_DET_REFRESH
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    TOOLBARFUNC 5
End If
End Sub

Private Sub lstDetails_GotFocus()
isFocus = 1
iRow = lstDetails.SelectedItem.Index
CLEAR_Details_Payroll
If CDbl(lstDetails.ListItems.Item(iRow).Text) <> 0 Then
    iLineCnt = 0
    For i = 1 To lstMortuary.ListItems.Count
        If CDbl(lstDetails.ListItems.Item(iRow).Text) = CDbl(lstMortuary.ListItems.Item(i).Text) Then
            iLineCnt = iLineCnt + 1
        End If
    Next i
    
    If CDbl(iLineCnt) > 0 Then
        lstPerPayroll.ListItems.Clear
        For i = 1 To lstMortuary.ListItems.Count
            If CDbl(lstDetails.ListItems.Item(iRow).Text) = CDbl(lstMortuary.ListItems.Item(i).Text) Then
                Set x = lstPerPayroll.ListItems.Add()
                x.Text = lstMortuary.ListItems.Item(i).Text
                x.SubItems(1) = lstMortuary.ListItems.Item(i).SubItems(3)
                x.SubItems(2) = lstMortuary.ListItems.Item(i).SubItems(4)
            End If
        Next i
    End If
End If
TRANS_DETAIL = is_DET_REFRESH
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    TOOLBARFUNC 5
End If
End Sub

Private Sub lstDetails_ItemClick(ByVal Item As MSComctlLib.ListItem)
iRow = lstDetails.SelectedItem.Index
CLEAR_Details_Payroll
If CDbl(lstDetails.ListItems.Item(iRow).Text) <> 0 Then
    iLineCnt = 0
    For i = 1 To lstMortuary.ListItems.Count
        If CDbl(lstDetails.ListItems.Item(iRow).Text) = CDbl(lstMortuary.ListItems.Item(i).Text) Then
            iLineCnt = iLineCnt + 1
        End If
    Next i
    
    If CDbl(iLineCnt) > 0 Then
        lstPerPayroll.ListItems.Clear
        For i = 1 To lstMortuary.ListItems.Count
            If CDbl(lstDetails.ListItems.Item(iRow).Text) = CDbl(lstMortuary.ListItems.Item(i).Text) Then
                Set x = lstPerPayroll.ListItems.Add()
                x.Text = lstMortuary.ListItems.Item(i).Text
                x.SubItems(1) = lstMortuary.ListItems.Item(i).SubItems(3)
                x.SubItems(2) = lstMortuary.ListItems.Item(i).SubItems(4)
            End If
        Next i
    End If
End If
End Sub

Private Sub lstDetails_LostFocus()
isFocus = 0
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    TOOLBARFUNC 2
End If
End Sub

Private Sub lstPerPayroll_Click()
isFocusDed = 1
iRowDed = lstPerPayroll.SelectedItem.Index
TRANS_DETAIL = is_DET_REFRESH
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    TOOLBARFUNC 5
End If
End Sub

Private Sub lstPerPayroll_GotFocus()
isFocusDed = 1
iRowDed = lstPerPayroll.SelectedItem.Index
TRANS_DETAIL = is_DET_REFRESH
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    TOOLBARFUNC 5
End If
End Sub

Private Sub lstPerPayroll_ItemClick(ByVal Item As MSComctlLib.ListItem)
iRowDed = lstPerPayroll.SelectedItem.Index
End Sub

Private Sub lstPerPayroll_LostFocus()
isFocusDed = 0
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    TOOLBARFUNC 2
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit":    PRESS_F2
    Case "Delete":  PRESS_DELETE
    Case "First":   If Toolbar1.Buttons(7).Caption = "Save" Then PRESS_F5 Else BROWSER GetSetting(App.EXEName, "PersonnelMortuary", "PersonnelMortuary", ""), "is_HOME"
    Case "Back":    If Toolbar1.Buttons(9).Caption = "Undo" Then PRESS_ESCAPE Else BROWSER GetSetting(App.EXEName, "PersonnelMortuary", "PersonnelMortuary", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "PersonnelMortuary", "PersonnelMortuary", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "PersonnelMortuary", "PersonnelMortuary", ""), "is_END"
    Case "Find":    PRESS_F6
    Case "Post":    PRESS_F8
    Case "Print":   PRESS_F9
    Case "Close":   PRESS_ESCAPE
End Select
End Sub

Private Sub txtAmount_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstDetails.ListItems
        .Item(iRow).SubItems(2) = Format(RETURNTEXTVALUE(txtAmount), "#,##0.00")
    End With
End If
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then txtPerPayroll.SetFocus
If KeyCode = vbKeyReturn Then
    picMain.Enabled = True
    picToolbar.Enabled = True
    picSLLines.Visible = False
    lstDetails.SetFocus
End If
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtPerPayroll_Change()
If TRANS_DETAIL = is_DET_ADDING Or _
TRANS_DETAIL = is_DET_EDITTING Then
    With lstPerPayroll.ListItems
        .Item(iRowDed).SubItems(2) = Format(RETURNTEXTVALUE(txtPerPayroll), "#,##0.00")
        
        For j = 1 To lstMortuary.ListItems.Count
            If CDbl(.Item(iRowDed).Text) = CDbl(lstMortuary.ListItems.Item(j).Text) And _
            DateValue(.Item(iRowDed).SubItems(1)) = DateValue(lstMortuary.ListItems.Item(j).SubItems(3)) Then
                lstMortuary.ListItems.Item(j).SubItems(4) = Format(RETURNTEXTVALUE(txtPerPayroll), "#,##0.00")
            End If
        Next j
        
    End With
End If
End Sub

Private Sub txtPerPayroll_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    picMain.Enabled = True
    picToolbar.Enabled = True
    picSLPerPeriod.Visible = False
    lstPerPayroll.SetFocus
End If
End Sub

Private Sub txtPerPayroll_KeyPress(KeyAscii As Integer)
KeyAscii = NUMBERKEYASCII(KeyAscii)
End Sub

Private Sub txtYear_Change()

'If cmbMonth.ListIndex = -1 Then Exit Sub
'If RETURNTEXTVALUE(txtYear) <= 0 Then Exit Sub
'If TRANSACTIONTYPE = is_ADDING Or _
'TRANSACTIONTYPE = is_EDITTING Then
'    With lstDetails.ListItems
'        For i = 1 To .Count
'            .Item(i).SubItems(2) = Format(DateSerial(RETURNTEXTVALUE(txtYear), cmbMonth.ListIndex + 1, 15), "mm/dd/yyyy")
'        Next i
'    End With
'End If

If cmbMonth.ListIndex = -1 Then Exit Sub
If RETURNTEXTVALUE(txtYear) <= 0 Then Exit Sub
If TRANSACTIONTYPE = is_ADDING Or _
TRANSACTIONTYPE = is_EDITTING Then
    With lstDetails.ListItems
        For i = 1 To .Count
            iPeriod = 0
            For j = 1 To lstMortuary.ListItems.Count
                If CDbl(.Item(i).Text) = CDbl(lstMortuary.ListItems.Item(j).Text) Then
                    iPeriod = iPeriod + 1
                    If iPeriod = 1 Then
                        dDate = DateSerial(RETURNTEXTVALUE(txtYear), cmbMonth.ListIndex + 1, 15)
                    Else
                        dDate = DateSerial(RETURNTEXTVALUE(txtYear), (cmbMonth.ListIndex + 1) + 1, 0)
                    End If
                    lstMortuary.ListItems.Item(j).SubItems(3) = Format(dDate, "mm/dd/yyyy")
                End If
            Next j
        Next i
    End With
End If
End Sub
