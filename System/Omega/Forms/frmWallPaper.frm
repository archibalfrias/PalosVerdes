VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWallPaper 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWallPaper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   Begin RPVGCC.b8Container picAddImages 
      Height          =   5055
      Left            =   600
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8916
      BackColor       =   15396057
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         Height          =   3915
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   3015
      End
      Begin VB.FileListBox File2 
         Appearance      =   0  'Flat
         Height          =   4320
         Left            =   3240
         TabIndex        =   12
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
         Left            =   7560
         Picture         =   "frmWallPaper.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1080
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
         Left            =   7560
         Picture         =   "frmWallPaper.frx":20DE
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   480
         Width           =   1560
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   3015
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   45
         TabIndex        =   14
         Top             =   45
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   609
         Caption         =   "b8TitleBar"
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
         Icon            =   "frmWallPaper.frx":2750
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   120
      ScaleHeight     =   7215
      ScaleWidth      =   10215
      TabIndex        =   4
      Top             =   1200
      Width           =   10215
      Begin RPVGCC.b8Container b8Container1 
         Height          =   7200
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   12700
         BackColor       =   16185592
         Begin VB.FileListBox File1 
            Height          =   1065
            Left            =   1200
            Pattern         =   "*.jpg"
            TabIndex        =   7
            Top             =   840
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   495
            Left            =   4680
            TabIndex        =   6
            Top             =   2040
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Image imgWallpaper 
            Appearance      =   0  'Flat
            Height          =   7080
            Left            =   60
            Stretch         =   -1  'True
            Top             =   60
            Width           =   10095
         End
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15000
      TabIndex        =   2
      Top             =   0
      Width           =   15000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   0
         TabIndex        =   3
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
         MouseIcon       =   "frmWallPaper.frx":2CEA
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
   Begin VB.TextBox txtWallPaper 
      Height          =   285
      Left            =   10560
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar Statusbar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   8520
      Width           =   10485
      _ExtentX        =   18494
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10560
      Top             =   1800
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
            Picture         =   "frmWallPaper.frx":3004
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWallPaper.frx":3CDE
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWallPaper.frx":49B8
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWallPaper.frx":5692
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWallPaper.frx":636C
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWallPaper.frx":7046
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWallPaper.frx":7D20
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWallPaper.frx":89FA
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWallPaper.frx":96D4
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWallPaper.frx":9FAE
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWallPaper.frx":AC88
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWallPaper.frx":B962
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWallPaper.frx":C63C
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWallPaper.frx":D316
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWallPaper.frx":DFF0
            Key             =   "IMG15"
         EndProperty
      EndProperty
   End
   Begin VB.Image ImageTmp 
      Height          =   135
      Left            =   10680
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmWallPaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRANSACTIONTYPE As Long
Const is_REFRESH = 0
Const is_ADDING = 1

Dim PrimaryKey As Double
Dim Filename, i, sPath, iPKTmp, iPK


Private Sub BROWSER(PriKey, is_Action As String)
Select Case is_Action
    Case "is_LOAD"
        If PriKey <> "" Then
            s = "SELECT TOP 1 tbl_Wallpaper.* " & _
                " FROM tbl_Wallpaper " & _
                " WHERE (PK = " & PriKey & ")" & _
                " ORDER BY PK"
        Else
            s = "SELECT TOP 1 tbl_Wallpaper.* " & _
                " FROM tbl_Wallpaper " & _
                " ORDER BY PK"
        End If
    Case "is_HOME"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Wallpaper.* " & _
            " FROM tbl_Wallpaper " & _
            " ORDER BY PK"
    Case "is_PAGEUP"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Wallpaper.* " & _
            " FROM tbl_Wallpaper " & _
            " WHERE (PK < " & PriKey & ")" & _
            " ORDER BY PK DESC"
    Case "is_PAGEDOWN"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Wallpaper.* " & _
            " FROM tbl_Wallpaper " & _
            " WHERE (PK > " & PriKey & ")" & _
            " ORDER BY PK "
    Case "is_END"
        If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
        s = "SELECT TOP 1 tbl_Wallpaper.* " & _
            " FROM tbl_Wallpaper " & _
            " ORDER BY PK DESC"
    Case Else: Exit Sub
End Select
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    PrimaryKey = rs!PK
    'imgWallpaper.Picture = LoadPicture(SHOW_WALLPAPER(PrimaryKey))
    imgWallpaper.Picture = LoadPicture(SHOW_IMAGES(rs!PK, 0, "Wallpaper"))
    Statusbar1.Panels(1).Text = rs!PK
    SaveSetting App.EXEName, "WallPaper", "WPaper", rs!PK
End If
rs.Close
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

Private Sub PRESS_INSERT()
'MsgBox "pass"
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
If picAddImages.Visible = True Then Exit Sub
picToolbar.Enabled = False
picAddImages.ZOrder 0
picAddImages.Visible = True
Dir1.Path = App.Path
File2.Path = Dir1.Path ' App.Path & "\Wallpaper"
'File1.Pattern = "*.JPG;*.JPEG;*.JPE;*.BMP;*.RLE;*.DIB;*.GIF;*.PNG;*.TIF;*.TIFF"
File2.Pattern = "*.JPG;*.JPEG"
Dir1.SetFocus
'MainForm.CommonDialog1.CancelError = True
'On Error GoTo ErrorHandler
'MainForm.CommonDialog1.Filter = "Image Files|*.JPG;*.JPEG;*.JPE;*.BMP;*.RLE;*.DIB;*.GIF;*.PNG;*.TIF;*.TIFF"
'MainForm.CommonDialog1.ShowOpen
'Filename = Trim(MainForm.CommonDialog1.Filename)
'txtWallPaper.Text = Filename
'imgWallpaper.Picture = LoadPicture(Filename)
'TRANSACTIONTYPE = is_ADDING
'Me.Caption = "Wallpaper - New"
'imgWallpaper.ToolTipText = "Press [F5] to save, Press [Esc] to Cancel"
'Exit Function
'ErrorHandler:
'Exit Function
End Sub

Private Sub PRESS_DELETE()
If Statusbar1.Panels(1).Text = "" Then Exit Sub
If TRANSACTIONTYPE <> is_REFRESH Then Exit Sub
'If PrimaryKey = 0 Then Exit Sub
If MsgBox("ARE YOU SURE IN DELETING THIS IMAGE?                    ", vbCritical + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
ConnOmega.Execute "DELETE FROM tbl_Wallpaper WHERE (PK = " & Statusbar1.Panels(1).Text & ")"
PrimaryKey = 0
BROWSER GetSetting(App.EXEName, "WallPaper", "WPaper", ""), "is_PAGEDOWN"
If PrimaryKey = 0 Then BROWSER GetSetting(App.EXEName, "WallPaper", "WPaper", ""), "is_HOME"
End Sub

Private Function PRESS_F5()
'If TRANSACTIONTYPE = is_ADDING Then
'    SAVE_IMAGES 0, 0, Trim(txtWallPaper.Text), "Wallpaper"
'    TRANSACTIONTYPE = is_REFRESH
'    Me.Caption = "Wallpaper - Browse"
'    imgWallpaper.ToolTipText = "Press [Ins] to add, Press [Del] to delete, Press [Home], [Page Up], [Page Down], [End] for Browsing, Press [Esc] to Close"
''    ImageTmp.Picture
'    BROWSER GetSetting(App.EXEName, "WallPaper", "WPaper", ""), "is_END"
'End If
End Function

Private Sub PRESS_ESCAPE()
If picAddImages.Visible = True Then cmdCancel_Click: Exit Sub
Unload Me
'If TRANSACTIONTYPE = is_REFRESH Then
'    Unload Me
'Else
'    TRANSACTIONTYPE = is_REFRESH
'    Me.Caption = "Wallpaper - Browse"
'    imgWallpaper.ToolTipText = "Press [Ins] to add, Press [Del] to delete, Press [Home], [Page Up], [Page Down], [End] for Browsing, Press [Esc] to Close"
'    BROWSER GetSetting(App.EXEName, "WallPaper", "WPaper", ""), "is_LOAD"
'    If PrimaryKey = 0 Then BROWSER GetSetting(App.EXEName, "WallPaper", "WPaper", ""), "is_HOME"
'End If
End Sub

Private Sub b8TitleBar1_CLoseClick()
cmdCancel_Click
End Sub

Private Sub cmdCancel_Click()
picAddImages.Visible = False
picToolbar.Enabled = True
End Sub

Private Sub cmdOK_Click()
If File2.ListCount = 0 Then MsgBox "No Image file/s!                    ", vbCritical, "Error...": Exit Sub
Screen.MousePointer = vbHourglass
File2.SetFocus
For i = 1 To File2.ListCount
    File2.ListIndex = i - 1
    On Error GoTo PG:
    sPath = File2.Path & "\" & File2.List(i - 1)
    iPKTmp = 1
    Do
        t = "SELECT tbl_Wallpaper.* " & _
            " FROM tbl_Wallpaper " & _
            " WHERE (FileName = '" & File2.List(i - 1) & "')"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount = 0 Then
            s = "SELECT tbl_Wallpaper.* " & _
                " FROM tbl_Wallpaper " & _
                " WHERE (PK = " & iPKTmp & ")"
            If rs.State = adStateOpen Then rs.Close
            rs.Open s, ConnOmega
            If rs.RecordCount = 0 Then
                Exit Do: rs.Close
            End If
            rs.Close
            iPKTmp = iPKTmp + 1
        End If
        rt.Close
    Loop

    ConnOmega.Execute "INSERT INTO tbl_Wallpaper " & _
                     " (PK, FileName) " & _
                     " VALUES (" & iPKTmp & ", '" & File2.List(i - 1) & "')"
    iPK = iPKTmp
    If CDbl(iPK) <> 0 Then
        SAVE_IMAGES iPK, 0, CStr(sPath), "Wallpaper"
    End If
Next i
Screen.MousePointer = vbDefault
MsgBox "Successfully Added!                     ", vbInformation, "Success"
cmdCancel_Click

Exit Sub
PG:
Screen.MousePointer = vbDefault
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub Dir1_Change()
File2.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
'txtWallPaper.Text = File1.FileName ' & File1.List(File1.ListIndex)
'On Error Resume Next
'imgWallpaper.Picture = LoadPicture(File1.FileName)
'SAVE_WALLPAPER File1.FileName
End Sub

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert:   PRESS_INSERT
    Case vbKeyDelete:   PRESS_DELETE
    Case vbKeyF5:       PRESS_F5
    Case vbKeyF8:       frmBackground.imgSlide.Picture = LoadPicture(SHOW_IMAGES(PrimaryKey, 0, "Wallpaper")): frmBackground.txtTimer.Text = "0"
    Case vbKeyHome:     BROWSER GetSetting(App.EXEName, "WallPaper", "WPaper", ""), "is_HOME"
    Case vbKeyPageUp:   BROWSER GetSetting(App.EXEName, "WallPaper", "WPaper", ""), "is_PAGEUP"
    Case vbKeyPageDown: BROWSER GetSetting(App.EXEName, "WallPaper", "WPaper", ""), "is_PAGEDOWN"
    Case vbKeyEnd:      BROWSER GetSetting(App.EXEName, "WallPaper", "WPaper", ""), "is_END"
    Case vbKeyEscape:   PRESS_ESCAPE
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
'Me.Height = 7620
'Me.Width = 10335
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
Me.Caption = "Wallpaper - Browse"
TRANSACTIONTYPE = is_REFRESH
TOOLBARFUNC 1
PrimaryKey = 0
BROWSER GetSetting(App.EXEName, "WallPaper", "WPaper", ""), "is_LOAD"
If PrimaryKey = 0 Then BROWSER GetSetting(App.EXEName, "WallPaper", "WPaper", ""), "is_HOME"
imgWallpaper.ToolTipText = "Press [Ins] to add, Press [Del] to delete, Press [Home], [Page Up], [Page Down], [End] for Browsing, Press [Esc] to Close"
'File1.FileName = App.Path & "\Wallpaper\"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If TRANSACTIONTYPE <> is_REFRESH Then Cancel = -1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Add":     PRESS_INSERT
    Case "Edit"
    Case "Delete":  PRESS_DELETE
    Case "First":   BROWSER GetSetting(App.EXEName, "WallPaper", "WPaper", ""), "is_HOME"
    Case "Back":    BROWSER GetSetting(App.EXEName, "WallPaper", "WPaper", ""), "is_PAGEUP"
    Case "Next":    BROWSER GetSetting(App.EXEName, "WallPaper", "WPaper", ""), "is_PAGEDOWN"
    Case "Last":    BROWSER GetSetting(App.EXEName, "WallPaper", "WPaper", ""), "is_END"
    Case "Find":
    Case "Refresh":
    Case "Close":   PRESS_ESCAPE
End Select
End Sub
