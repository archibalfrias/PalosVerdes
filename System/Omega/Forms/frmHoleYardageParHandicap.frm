VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHoleYardageParHandicap 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHoleYardageParHandicap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picHoleYardageParHandicap 
      BackColor       =   &H00C6B8A4&
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   120
      ScaleHeight     =   4695
      ScaleWidth      =   7815
      TabIndex        =   1
      Top             =   120
      Width           =   7815
      Begin MSFlexGridLib.MSFlexGrid FGrid 
         Height          =   4665
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   8229
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   13023396
         ForeColorFixed  =   255
         BackColorSel    =   8388608
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         FocusRect       =   0
      End
   End
End
Attribute VB_Name = "frmHoleYardageParHandicap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim HEADER1$

Dim i

Private Function CUSTOM_GRID()

With FGrid
    .Clear
    HEADER1$ = HEADER1$ & "|" & _
               "Hole #" & "|" & _
               "Par" & "|" & _
               "Handicap Index" & "|" & _
               "Gold" & "|" & _
               "Blue" & "|" & _
               "White" & "|" & _
               "Red"
    .FormatString = HEADER1$
    .ColWidth(1) = 1000     'Hole #
    '.ColWidth(2) = 0        'PK
    .ColWidth(2) = 1000     'Par
    .ColWidth(3) = 1300     'Handicap Index
    .ColWidth(4) = 1000     'Gold
    .ColWidth(5) = 1000     'Blue
    .ColWidth(6) = 1000     'White
    .ColWidth(7) = 1000     'Red
    .ColAlignment(1) = 3 'flexAlignRightCenter
    .ColAlignment(2) = 3 'flexAlignRightCenter
    .ColAlignment(3) = 3 'flexAlignRightCenter
    .ColAlignment(4) = flexAlignRightCenter
    .ColAlignment(5) = flexAlignRightCenter
    .ColAlignment(6) = flexAlignRightCenter
    .ColAlignment(7) = flexAlignRightCenter
'    .ColAlignment(8) = flexAlignRightCenter
    .Rows = 2
End With
End Function

Private Function LOAD_HOLE_INFO()
s = "SELECT tbl_Scoring_Yardage_Par_HandicapIndex.* " & _
    " FROM tbl_Scoring_Yardage_Par_HandicapIndex " & _
    " ORDER BY Hole"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    i = 0
    With FGrid
        While Not rs.EOF
            i = i + 1
            .Rows = i + 1
            .TextMatrix(i, 1) = rs!Hole
            '.TextMatrix(i, 2) = rs!PK
            .TextMatrix(i, 2) = rs!Par
            .TextMatrix(i, 3) = rs!HandicapIndex
            .TextMatrix(i, 4) = rs!Gold
            .TextMatrix(i, 5) = rs!Blue
            .TextMatrix(i, 6) = rs!White
            .TextMatrix(i, 7) = rs!Red
            rs.MoveNext
        Wend
    End With
End If
rs.Close
End Function

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Top = (MainForm.ScaleHeight - Me.Height) / 4
Me.Left = (MainForm.ScaleWidth - Me.Width) / 2
Me.Icon = MainForm.ImageListMother.ListImages(MainForm.trView.Nodes(iTreeViewIndex).Image).Picture
Me.Caption = gbl_Form_Caption
CUSTOM_GRID
LOAD_HOLE_INFO
End Sub
