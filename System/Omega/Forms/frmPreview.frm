VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPreview 
   BackColor       =   &H00C6B8A4&
   Caption         =   "Preview"
   ClientHeight    =   6840
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11610
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C6B8A4&
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   -70
      Width           =   11460
      Begin VB.Timer Timer_Items_Section 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2520
         Top             =   120
      End
      Begin VB.Timer Timer_ItemsTransaction 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2040
         Top             =   120
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00C6B8A4&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   255
         Left            =   6720
         ScaleHeight     =   255
         ScaleWidth      =   2415
         TabIndex        =   0
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdOpen 
         BackColor       =   &H00D4D4D4&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   600
         Picture         =   "frmPreview.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   150
         Width           =   400
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00D4D4D4&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   120
         Picture         =   "frmPreview.frx":2D2C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   150
         Width           =   400
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00D4D4D4&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1080
         Picture         =   "frmPreview.frx":31E2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   150
         Width           =   400
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00D4D4D4&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1560
         Picture         =   "frmPreview.frx":3714
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   150
         Width           =   400
      End
      Begin VB.Label lblPrinter 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3600
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   6375
      End
   End
   Begin MSComctlLib.ListView lstReport 
      Height          =   5835
      Left            =   0
      TabIndex        =   7
      Top             =   550
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   10292
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LastPage     As Long
Public PrintType    As Long
Dim strPath         As String

Public iItemKey, sItemCode, sItemDesc, sSectCode
Public sSectName, sClassCode, sClassName

Dim iLoading As Long


Dim i, cnt, Page, iNo, StrFile, x, Pages, Array1, dblWidth

Private Sub cmdClose_Click()
Picture4.SetFocus
Unload Me
End Sub

Private Sub cmdOpen_Click()
Picture4.SetFocus
MainForm.CommonDialog1.DialogTitle = "OPEN FILE"
MainForm.CommonDialog1.Filename = ""
MainForm.CommonDialog1.Filter = "Text Files (*.txt)|*.txt"
MainForm.CommonDialog1.FilterIndex = 1
MainForm.CommonDialog1.ShowOpen
strPath = MainForm.CommonDialog1.Filename
If strPath <> "" Then
    With lstReport.ListItems
        .Clear
        i = 0
        Pages = 1
        Open strPath For Input As #1
            Do Until EOF(1)
                Line Input #1, StrFile
                i = i + 1
                If i = 1 Then
                    If Trim(StrFile) <> "" Then
                        Array1 = Split(StrFile, "|", -1, 1)
                        If UBound(Array1) = 1 Then
                            dblWidth = Array1(0)
                            PrintType = CLng(Array1(1))
                        ElseIf UBound(Array1) = 0 Then
                            dblWidth = Array1(0)
                            PrintType = 0
                        End If
                    Else
                        dblWidth = 11000.13
                    End If
                    lstReport.ColumnHeaders(1).Width = CDbl(dblWidth)
                    If Trim(StrFile) = "" Then
                        Set x = .Add()
                        x.Text = StrFile
                    End If
                Else
                    If Trim(StrFile) = Chr(12) Then
                        Pages = Pages + 1
                    End If
                    Set x = .Add()
                    x.Text = StrFile
                End If
            Loop
        Close #1
        LastPage = Pages
    End With
End If
End Sub

Private Sub cmdPrint_Click()
Picture4.SetFocus
If lstReport.ListItems.Count = 0 Then Exit Sub
frmPrinter.PrintType = PrintType
frmPrinter.PRINT_TRANSACTION = 0
frmPrinter.txtCopies.Text = "1"
frmPrinter.txtPgFrom.Text = "1"
frmPrinter.txtPgTo.Text = LastPage
frmPrinter.Show 1
End Sub

Private Sub cmdSave_Click()
Picture4.SetFocus
If lstReport.ListItems.Count = 0 Then Exit Sub
MainForm.CommonDialog1.DialogTitle = "SAVE FILE"
MainForm.CommonDialog1.Filename = ""
MainForm.CommonDialog1.Filter = "Text Files (*.txt)|*.txt|Excel Files (*.xls)|*.xls|Lutos 123 (*.wk1)|*.wk1"
MainForm.CommonDialog1.ShowSave
strPath = MainForm.CommonDialog1.Filename
If strPath <> "" Then
    With lstReport.ListItems
        If .Count > 0 Then
            Open strPath For Output As #1
            Print #1, lstReport.ColumnHeaders(1).Width & "|" & PrintType
            For i = 1 To .Count
                Print #1, .Item(i).Text
            Next i
            Close #1
        End If
    End With
    
    If MainForm.CommonDialog1.FilterIndex = 1 Then
        If MsgBox("Would you like to open the file just saved?              ", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm") = vbYes Then
            Shell "notepad.exe " & strPath, vbMaximizedFocus
        End If
    End If
    
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    If iLoading = 0 Then Unload Me
End If
End Sub

Private Sub Form_Load()
KeyPreview = True
iLoading = 0
End Sub

Private Sub Form_Resize()
On Error Resume Next
Frame1.Top = -80
Frame1.Left = 0
Frame1.Width = Me.ScaleWidth
lstReport.Top = 540
lstReport.Height = Me.ScaleHeight - (MainForm.Statusbar1.Height + 230)
lstReport.Width = Me.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
If iLoading = 1 Then Cancel = -1
End Sub

Private Sub Timer_Items_Section_Timer()
Timer_Items_Section.Enabled = False
iLoading = 1
lstReport.ColumnHeaders(1).Width = 11000
With lstReport.ListItems
    iNo = 0: cnt = 0: Page = 1: sClassCode = ""
    .Clear
    .Add , , ""
    .Add , , ""
    .Add , , Space(2) & "Item's List" & Space(85 - Len("Page " & Page)) & _
             "Page " & Page
    .Add , , Space(2) & "Section : " & sSectCode & " - " & sSectName
    .Add , , Space(2) & "================================================================================================"
    .Add , , Space(2) & " No" & _
             Space(2) & "ItemCode" & _
             Space(2) & "Description" & Space(50 - Len("Description")) & _
             Space(2) & "Unit" & Space(15 - Len("Unit")) & _
             Space(2) & Space(12 - Len("Cost")) & "Cost"
    .Add , , Space(2) & "================================================================================================"
    s = "SELECT tbl_Inv_Class.ClassCode, tbl_Inv_Class.ClassName, " & _
        " tbl_Inv_Items.ItemCode, tbl_Inv_Items.ItemDesc, tbl_Inv_Items.Unit, " & _
        " tbl_Inv_Items.Cost " & _
        " FROM tbl_Inv_Items LEFT OUTER JOIN " & _
        " tbl_Inv_Class ON tbl_Inv_Items.ClassKey = tbl_Inv_Class.PK LEFT OUTER JOIN " & _
        " tbl_Inv_Section ON tbl_Inv_Class.SectKey = tbl_Inv_Section.PK " & _
        " WHERE (tbl_Inv_Section.SectCode = '" & sSectCode & "') " & _
        " ORDER BY tbl_Inv_Class.ClassCode, tbl_Inv_Items.ItemCode"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        If Trim(CStr(sClassCode)) <> "" Then
            If Trim(CStr(sClassCode)) <> rs!ClassCode Then
                sClassCode = rs!ClassCode
                .Add , , Space(7) & rs!ClassCode & " - " & rs!ClassName
                cnt = cnt + 1
                iNo = 0
                If cnt = 50 Then
                    cnt = 0: Page = Page + 1
                    .Add , , Chr(12)
                    .Add , , ""
                    .Add , , ""
                    .Add , , Space(2) & "Item's List" & Space(85 - Len("Page " & Page)) & _
                             "Page " & Page
                    .Add , , Space(2) & "Section : " & sSectCode & " - " & sSectName
                    .Add , , Space(2) & "================================================================================================"
                    .Add , , Space(2) & " No" & _
                             Space(2) & "ItemCode" & _
                             Space(2) & "Description" & Space(50 - Len("Description")) & _
                             Space(2) & "Unit" & Space(15 - Len("Unit")) & _
                             Space(2) & Space(12 - Len("Cost")) & "Cost"
                    .Add , , Space(2) & "================================================================================================"
                End If
            End If
        Else
            sClassCode = rs!ClassCode
            .Add , , Space(7) & rs!ClassCode & " - " & rs!ClassName
            cnt = cnt + 1
            iNo = 0
            If cnt = 50 Then
                cnt = 0: Page = Page + 1
                .Add , , Chr(12)
                .Add , , ""
                .Add , , ""
                .Add , , Space(2) & "Item's List" & Space(85 - Len("Page " & Page)) & _
                         "Page " & Page
                .Add , , Space(2) & "Section : " & sSectCode & " - " & sSectName
                .Add , , Space(2) & "================================================================================================"
                .Add , , Space(2) & " No" & _
                         Space(2) & "ItemCode" & _
                         Space(2) & "Description" & Space(50 - Len("Description")) & _
                         Space(2) & "Unit" & Space(15 - Len("Unit")) & _
                         Space(2) & Space(12 - Len("Cost")) & "Cost"
                .Add , , Space(2) & "================================================================================================"
            End If
        End If
        iNo = iNo + 1
        .Add , , Space(5 - Len(iNo)) & iNo & "." & _
                 Space(1) & Trim(rs!ItemCode) & Space(8 - Len(Trim(rs!ItemCode))) & _
                 Space(2) & Trim(rs!ItemDesc) & Space(50 - Len(Trim(rs!ItemDesc))) & _
                 Space(2) & Trim(rs!Unit) & Space(15 - Len(Trim(rs!Unit))) & _
                 Space(2) & Space(12 - Len(Format(rs!Cost, "#,##0.00"))) & Format(rs!Cost, "#,##0.00")
        cnt = cnt + 1
        If cnt = 50 Then
            cnt = 0: Page = Page + 1
            .Add , , Chr(12)
            .Add , , ""
            .Add , , ""
            .Add , , Space(2) & "Item's List" & Space(85 - Len("Page " & Page)) & _
                     "Page " & Page
            .Add , , Space(2) & "Section : " & sSectCode & " - " & sSectName
            .Add , , Space(2) & "================================================================================================"
            .Add , , Space(2) & " No" & _
                     Space(2) & "ItemCode" & _
                     Space(2) & "Description" & Space(50 - Len("Description")) & _
                     Space(2) & "Unit" & Space(15 - Len("Unit")) & _
                     Space(2) & Space(12 - Len("Cost")) & "Cost"
            .Add , , Space(2) & "================================================================================================"
        End If
        rs.MoveNext
    Wend
    rs.Close
End With
iLoading = 0
End Sub

Private Sub Timer_ItemsTransaction_Timer()
Timer_ItemsTransaction.Enabled = False
iLoading = 1
lstReport.ColumnHeaders(1).Width = 21000
With lstReport.ListItems
    .Clear
    .Add , , ""
    .Add , , ""
    .Add , , Space(2) & "Item's Transaction"
    .Add , , Space(2) & sItemCode & " - " & sItemDesc
    .Add , , ""
    .Add , , Space(2) & "==============================================================================================================================="
    .Add , , Space(2) & "   Date   " & _
             Space(2) & "Document Number     " & _
             Space(2) & "Document Name       " & _
             Space(2) & "Location            " & _
             Space(2) & "       Quantity" & _
             Space(2) & "     Total Cost" & _
             Space(2) & "  Total Netcost"
    .Add , , Space(2) & "==============================================================================================================================="
    s = "SELECT tbl_Inv_Items_Transaction.DocDate, " & _
        " tbl_Inv_Items_Transaction.DocNumber, " & _
        " tbl_Inv_DocumentType.DocumentName, " & _
        " tbl_Inv_Location.LocName, " & _
        " tbl_Inv_Items_Transaction.Quantity, " & _
        " tbl_Inv_Items_Transaction.TotalCost, " & _
        " tbl_Inv_Items_Transaction.TotalNetCost " & _
        " FROM tbl_Inv_Items_Transaction LEFT OUTER JOIN " & _
        " tbl_Inv_Location ON tbl_Inv_Items_Transaction.Location = tbl_Inv_Location.PK LEFT OUTER JOIN " & _
        " tbl_Inv_DocumentType ON tbl_Inv_Items_Transaction.DocType = tbl_Inv_DocumentType.PK LEFT OUTER JOIN " & _
        " tbl_Inv_Items ON tbl_Inv_Items_Transaction.ItemKey = tbl_Inv_Items.PK " & _
        " Where (tbl_Inv_Items_Transaction.ItemKey = " & iItemKey & ") " & _
        " ORDER BY tbl_Inv_Items_Transaction.DocDate"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        .Add , , Space(2) & Format(rs!DocDate, "mm/dd/yyyy") & _
                 Space(2) & Mid(rs!DocNumber, 1, 20) & Space(20 - Len(Mid(rs!DocNumber, 1, 20))) & _
                 Space(2) & Mid(rs!DocumentName, 1, 20) & Space(20 - Len(Mid(rs!DocumentName, 1, 20))) & _
                 Space(2) & Mid(rs!LocName, 1, 20) & Space(20 - Len(Mid(rs!LocName, 1, 20))) & _
                 Space(2) & Space(15 - Len(Format(rs!Quantity, "#,##0.00"))) & Format(rs!Quantity, "#,##0.00") & _
                 Space(2) & Space(15 - Len(Format(rs!TotalCost, "#,##0.00"))) & Format(rs!TotalCost, "#,##0.00") & _
                 Space(2) & Space(15 - Len(Format(rs!TotalNetCost, "#,##0.00"))) & Format(rs!TotalNetCost, "#,##0.00")
        rs.MoveNext
    Wend
    rs.Close
End With
iLoading = 0
End Sub
