VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MainForm.frx":0CCA
   ScaleHeight     =   3375
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerExportData 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4080
      Top             =   3000
   End
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3000
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer TimerImportData 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   3000
   End
   Begin VB.CommandButton cmdExportData 
      Caption         =   "Export Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      MouseIcon       =   "MainForm.frx":40CF0
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton cmdImportData 
      Caption         =   "Import Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      MouseIcon       =   "MainForm.frx":40FFA
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton cmdScoring 
      Caption         =   "Scoring"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      MouseIcon       =   "MainForm.frx":41304
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.PictureBox picFocus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12360
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdTournamentInfo 
      Caption         =   "Tournament Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      MouseIcon       =   "MainForm.frx":4160E
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim s As String
Dim rs As New ADODB.Recordset
Dim t As String
Dim rt As New ADODB.Recordset
Dim u As String
Dim ru As New ADODB.Recordset

Dim sPath

Dim WorkbookName As String
Dim iWorkSheet As Integer
Dim RowCnt, ColCnt, strRange, i, l, k, strValue, iReset, strAmount, iPK, Arr, Arr1, iLocationKey, StrFile, sLine, _
sFileNameMaster, sFileNameDetail

Private Sub cmdExportData_Click()
CommonDialog1.DialogTitle = "OPEN FILE"
CommonDialog1.FileName = ""
CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx|Text File|*.txt"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowSave
sPath = CommonDialog1.FileName
txtPath.Text = sPath
If Trim(txtPath.Text) = "" Then Exit Sub

WorkbookName = txtPath.Text

TimerExportData.Enabled = True

Exit Sub
ErrorHandler:
Exit Sub

End Sub

Private Sub cmdImportData_Click()
CommonDialog1.DialogTitle = "OPEN FILE"
CommonDialog1.FileName = ""
CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx|Text File|*.txt"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowOpen
sPath = CommonDialog1.FileName
txtPath.Text = sPath
If Trim(txtPath.Text) = "" Then Exit Sub
TimerImportData.Enabled = True
End Sub

Private Sub cmdScoring_Click()
frmScoreCard.Show 1
End Sub

Private Sub cmdTournamentInfo_Click()
picFocus.SetFocus
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Caption = "Remote Scoring"
DataOpen ConnRS
picFocus.TabIndex = 0
End Sub

Private Sub TimerExportData_Timer()
TimerExportData.Enabled = False

Arr = Split(Trim(Trim(txtPath.Text)), "\", -1, 1)
Arr1 = Split(CStr(Arr(UBound(Arr))), ".", -1, 1)

Screen.MousePointer = vbHourglass
'On Error GoTo PG:
If Arr1(UBound(Arr1)) = "txt" Then    'Text File
    sFileNameMaster = Trim(txtPath.Text)
    'sFileNameDetail = Replace(Trim(txtPath.Text), CStr(Arr(UBound(Arr))), Arr1(0) & "_det.txt")
    Open sFileNameMaster For Output As #1
        'Score Card
        s = "SELECT PK, TournamentKey, LocationKey, PlayerKey, DDate " & _
            " FROM tbl_Scoring_ScoreCard"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnRS
        While Not rs.EOF
            sLine = "ScoreCard["
            For i = 0 To rs.Fields.Count - 1
                sLine = sLine & rs.Fields(i).Value & "|"
            Next i
            Print #1, Mid(CStr(sLine), 1, Len(sLine) - 1)
            
            t = "SELECT ScoreCardKey, Hole, Par, Handicap, Score, Gross, Net " & _
                " FROM tbl_Scoring_ScoreCard_Detail " & _
                " WHERE (ScoreCardKey = " & rs!PK & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnRS
            While Not rt.EOF
                sLine = "ScoreCardDetail["
                For i = 0 To rt.Fields.Count - 1
                    sLine = sLine & rt.Fields(i).Value & "|"
                Next i
                Print #1, Mid(CStr(sLine), 1, Len(sLine) - 1)
                rt.MoveNext
            Wend
            rt.Close
            
            
            rs.MoveNext
        Wend
        rs.Close
    Close #1
    
'    Open sFileNameDetail For Output As #1
'        'Score Card Detail
'        s = "SELECT ScoreCardKey, Hole, Par, Handicap, Score, Gross, Net " & _
'            " FROM tbl_Scoring_ScoreCard_Detail"
'        If rs.State = adStateOpen Then rs.Close
'        rs.Open s, ConnRS
'        While Not rs.EOF
'            sLine = ""
'            For i = 0 To rs.Fields.Count - 1
'                sLine = sLine & rs.Fields(i).Value & "|"
'            Next i
'            Print #1, Mid(CStr(sLine), 1, Len(sLine) - 1)
'            rs.MoveNext
'        Wend
'        rs.Close
'    Close #1
    
    If MsgBox("Would you like to open the file just saved?              ", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm") = vbYes Then
        Shell "Notepad.exe " & sFileNameMaster, vbMaximizedFocus
    End If
    
    Screen.MousePointer = vbDefault
Else
    ColCnt = 0: RowCnt = 0
    'cnt = 0
    Set xlsApp = CreateObject("Excel.Application")
    With xlsApp
        .Visible = False
        .Workbooks.Add
        .DisplayAlerts = False
        iWorkSheet = 0
        .Workbooks(1).Sheets(2).Delete
        '.Workbooks(1).Sheets(2).Delete
        
        s = "SELECT PK, TournamentKey, LocationKey, PlayerKey, DDate " & _
            " FROM tbl_Scoring_ScoreCard"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnRS
        If rs.RecordCount > 0 Then
            iWorkSheet = iWorkSheet + 1
            .Workbooks(1).Sheets(iWorkSheet).Activate
            .Workbooks(1).Sheets(iWorkSheet).Name = "ScoreCard"
            RowCnt = 0: ColCnt = 0
            RowCnt = RowCnt + 1
            For i = 0 To rs.Fields.Count - 1
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = rs.Fields(i).Name
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = True
            Next i
            While Not rs.EOF
                RowCnt = RowCnt + 1
                ColCnt = 0
                For i = 0 To rs.Fields.Count - 1
                    ColCnt = ColCnt + 1
                    strRange = EXCEL_RANGE(ColCnt, RowCnt)
                    .Range(strRange).Value = rs.Fields(i).Value
                    .Range(strRange).Font.Name = "Tahoma"
                    .Range(strRange).Font.Size = 8
                    .Range(strRange).Font.Bold = False
                Next i
                rs.MoveNext
            Wend
            'rs.Close
        End If
        rs.Close
        
        iWorkSheet = iWorkSheet + 1
        .Workbooks(1).Sheets(iWorkSheet).Activate
        .Workbooks(1).Sheets(iWorkSheet).Name = "ScoreCardDetails"
        RowCnt = 0: ColCnt = 0
        s = "SELECT ScoreCardKey, Hole, Par, Handicap, Score, Gross, Net " & _
            " FROM tbl_Scoring_ScoreCard_Detail"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnRS
        RowCnt = RowCnt + 1
        For i = 0 To rs.Fields.Count - 1
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            .Range(strRange).Value = rs.Fields(i).Name
            .Range(strRange).Font.Name = "Tahoma"
            .Range(strRange).Font.Size = 8
            .Range(strRange).Font.Bold = True
        Next i
        While Not rs.EOF
            RowCnt = RowCnt + 1
            ColCnt = 0
            For i = 0 To rs.Fields.Count - 1
                ColCnt = ColCnt + 1
                strRange = EXCEL_RANGE(ColCnt, RowCnt)
                .Range(strRange).Value = rs.Fields(i).Value
                .Range(strRange).Font.Name = "Tahoma"
                .Range(strRange).Font.Size = 8
                .Range(strRange).Font.Bold = False
            Next i
            rs.MoveNext
        Wend
        rs.Close
        
        
         If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
        .ActiveWorkbook.SaveAs FileName:=WorkbookName
        .Visible = True
        Set xlsApp = Nothing
    End With
    Screen.MousePointer = vbDefault
End If
Exit Sub
PG:
Screen.MousePointer = vbDefault
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
Exit Sub
End Sub

Private Sub TimerImportData_Timer()
TimerImportData.Enabled = False

Arr = Split(Trim(Trim(txtPath.Text)), "\", -1, 1)
Arr1 = Split(CStr(Arr(UBound(Arr))), ".", -1, 1)

Screen.MousePointer = vbHourglass
If Arr1(UBound(Arr1)) = "txt" Then    'Text File
    ConnRS.Execute "DELETE FROM tbl_Scoring_Location_Details"
    ConnRS.Execute "DELETE FROM tbl_Scoring_TournamentInfo_Class"
    ConnRS.Execute "DELETE FROM tbl_Scoring_TournamentInfo_Index"
    ConnRS.Execute "DELETE FROM tbl_Scoring_TournamentInfo_Location"
    Open CStr(txtPath.Text) For Input As #1
        Do Until EOF(1)
            Line Input #1, StrFile
            Arr = Split(StrFile, "[", -1, 1)
            Select Case Arr(0)
                Case "LOCATION"
                    Arr1 = Split(Arr(1), "|", -1, 1)
                    t = "SELECT tbl_Scoring_Location.* " & _
                        " FROM tbl_Scoring_Location " & _
                        " WHERE (PK = " & Arr1(0) & ")"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnRS
                    If rt.RecordCount = 0 Then
                        ConnRS.Execute "INSERT INTO tbl_Scoring_Location " & _
                                       " (PK, ScoringLocation) " & _
                                       " VALUES (" & Arr1(0) & ", '" & FORMATSQL(CStr(Arr1(1))) & "')"
                    Else
                        ConnRS.Execute "UPDATE tbl_Scoring_Location " & _
                                       " SET ScoringLocation = '" & FORMATSQL(CStr(Arr1(1))) & "' " & _
                                       " WHERE (PK = " & Arr1(0) & ")"
                    End If
                    rt.Close
                Case "LOCATION_DETAILS"
                    Arr1 = Split(Arr(1), "|", -1, 1)
                    ConnRS.Execute "INSERT INTO tbl_Scoring_Location_Details " & _
                                   " (MasterKey, Line, Description, H1, H2, H3, H4, H5, H6, H7, H8, H9, H10, H11, H12, H13, H14, H15, H16, H17, H18) " & _
                                   " VALUES (" & Arr1(0) & ", " & Arr1(1) & ", '" & FORMATSQL(CStr(Arr1(2))) & "', " & Arr1(3) & ", " & Arr1(4) & ", " & Arr1(5) & ", " & _
                                   " " & Arr1(6) & ", " & Arr1(7) & ", " & Arr1(8) & ", " & Arr1(9) & ", " & Arr1(10) & ", " & Arr1(11) & ", " & Arr1(12) & ", " & _
                                   " " & Arr1(13) & ", " & Arr1(14) & ", " & Arr1(15) & ", " & Arr1(16) & ", " & Arr1(17) & ", " & Arr1(18) & ", " & Arr1(19) & ", " & _
                                   " " & Arr1(20) & ")"
                Case "SCORING_SYSTEM"
                    Arr1 = Split(Arr(1), "|", -1, 1)
                    t = "SELECT tbl_Scoring_System.* " & _
                        " FROM tbl_Scoring_System " & _
                        " WHERE (PK = " & Arr1(0) & ")"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnRS
                    If rt.RecordCount = 0 Then
                        ConnRS.Execute "INSERT INTO tbl_Scoring_System " & _
                                       " (PK, ScoringSystem) " & _
                                       " VALUES (" & Arr1(0) & ", '" & FORMATSQL(CStr(Arr1(1))) & "')"
                    Else
                        ConnRS.Execute "UPDATE tbl_Scoring_System " & _
                                       " SET ScoringSystem = '" & FORMATSQL(CStr(Arr1(1))) & "' " & _
                                       " WHERE (PK = " & Arr1(0) & ")"
                    End If
                    rt.Close
                Case "TOURNAMENT_INFO"
                    Arr1 = Split(Arr(1), "|", -1, 1)
                    t = "SELECT tbl_Scoring_TournamentInfo.* " & _
                        " FROM tbl_Scoring_TournamentInfo " & _
                        " WHERE (PK = " & Arr1(0) & ")"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnRS
                    If rt.RecordCount = 0 Then
                        ConnRS.Execute "INSERT INTO tbl_Scoring_TournamentInfo " & _
                                       " (PK, TournamentName, TournamentStart, TournamentEnd, TournamentDays, NoofPlays, NoofPlayerPerTeam, Remarks, Locked, Activated, Scoring, TeamPlay, " & _
                                       " PlayerToCount, HandicapDivisor, PointsToCountTeam, TeamAverage, IndividualPlay, AllowTeamPerPlayer, PointsToCountIndi, ParGrossPoints) " & _
                                       " VALUES (" & Arr1(0) & ", '" & FORMATSQL(CStr(Arr1(1))) & "', '" & FormatDateTime(Arr1(2), vbShortDate) & "','" & FormatDateTime(Arr1(3), vbShortDate) & "', " & _
                                       " " & Arr1(4) & ", " & Arr1(5) & ", " & Arr1(6) & ", '" & FORMATSQL(CStr(Arr1(7))) & "', " & Arr1(8) & ", " & Arr1(9) & ", " & Arr1(11) & ", " & _
                                       " " & Arr1(12) & ", " & Arr1(13) & ", " & Arr1(14) & ", " & Arr1(15) & ", " & Arr1(16) & ", " & Arr1(17) & ", " & _
                                       " " & Arr1(18) & ", " & Arr1(19) & ", " & Arr1(20) & ")"
                    Else
                        ConnRS.Execute "UPDATE tbl_Scoring_TournamentInfo " & _
                                       " SET TournamentName = '" & FORMATSQL(CStr(Arr1(1))) & "', TournamentStart = '" & FormatDateTime(Arr1(2), vbShortDate) & "', " & _
                                       " TournamentEnd = '" & FormatDateTime(Arr1(3), vbShortDate) & "', TournamentDays = " & Arr1(4) & ", NoofPlays = " & Arr1(5) & ", " & _
                                       " NoofPlayerPerTeam = " & Arr1(6) & ", Remarks = '" & FORMATSQL(CStr(Arr1(7))) & "', Locked = " & Arr1(8) & ", Activated = " & Arr1(9) & ", " & _
                                       " Scoring = " & Arr1(11) & ", TeamPlay = " & Arr1(12) & ", PlayerToCount = " & Arr1(13) & ", HandicapDivisor = " & Arr1(14) & ", " & _
                                       " PointsToCountTeam = " & Arr1(15) & ", TeamAverage = " & Arr1(16) & ", IndividualPlay = " & Arr1(17) & ", " & _
                                       " AllowTeamPerPlayer = " & Arr1(18) & ", PointsToCountIndi = " & Arr1(19) & ", ParGrossPoints = " & Arr1(20) & " " & _
                                       " WHERE (PK = " & Arr1(0) & ")"
                    End If
                    rt.Close
                    t = "SELECT tbl_Scoring_TournamentInfo.* " & _
                        " FROM tbl_Scoring_TournamentInfo " & _
                        " WHERE (PK <> " & Arr1(0) & ")"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnRS
                    While Not rt.EOF
                        ConnRS.Execute "UPDATE tbl_Scoring_TournamentInfo " & _
                                       " SET Activated = 0 " & _
                                       " WHERE (PK = " & Arr1(0) & ")"
                        rt.MoveNext
                    Wend
                    rt.Close
                Case "TOURNAMENT_INFO_CLASS"
                    Arr1 = Split(Arr(1), "|", -1, 1)
                    ConnRS.Execute "INSERT INTO tbl_Scoring_TournamentInfo_Class " & _
                                   " (TournamentKey, Class, HFrom, HTo) " & _
                                   " VALUES (" & Arr1(0) & ", '" & Arr1(1) & "', " & Arr1(2) & ", " & Arr1(3) & ")"
                Case "TOURNAMENT_INFO_INDEX"
                    Arr1 = Split(Arr(1), "|", -1, 1)
                    ConnRS.Execute "INSERT INTO tbl_Scoring_TournamentInfo_Index " & _
                                   " (TournamentKey, Class, HFrom, HTo) " & _
                                   " VALUES (" & Arr1(0) & ", '" & Arr1(1) & "', " & Arr1(2) & ", " & Arr1(3) & ")"
                Case "TOURNAMENT_INFO_LOCATION"
                    Arr1 = Split(Arr(1), "|", -1, 1)
                    ConnRS.Execute "INSERT INTO tbl_Scoring_TournamentInfo_Location " & _
                                   " (MasterKey, LocationKey) " & _
                                   " VALUES (" & Arr1(0) & ", " & Arr1(1) & ")"
                Case "PLAYER_NAME"
                    Arr1 = Split(Arr(1), "|", -1, 1)
                    t = "SELECT tbl_Scoring_PlayerName.* " & _
                        " FROM tbl_Scoring_PlayerName " & _
                        " WHERE (PK = " & Arr1(0) & ")"
                    If rt.State = adStateOpen Then rt.Close
                    rt.Open t, ConnRS
                    If rt.RecordCount = 0 Then
                        ConnRS.Execute "INSERT INTO tbl_Scoring_PlayerName " & _
                                       " (PK, TournamentKey, LastName, FirstName, MiddleName, Gender, HandiCap, Class, AllowedTeam, iIndex) " & _
                                       " VALUES (" & Arr1(0) & ", " & Arr1(1) & ", '" & FORMATSQL(CStr(Arr1(2))) & "', '" & FORMATSQL(CStr(Arr1(3))) & "', " & _
                                       " '" & FORMATSQL(CStr(Arr1(4))) & "', '" & FORMATSQL(CStr(Arr1(5))) & "', " & Arr1(6) & ", '" & FORMATSQL(CStr(Arr1(7))) & "', " & _
                                       " " & Arr1(8) & ", " & Arr1(9) & ")"
                    Else
                        ConnRS.Execute "UPDATE tbl_Scoring_PlayerName " & _
                                       " SET TournamentKey = " & Arr1(1) & ", LastName = '" & FORMATSQL(CStr(Arr1(2))) & "', " & _
                                       " FirstName = '" & FORMATSQL(CStr(Arr1(3))) & "', MiddleName = '" & FORMATSQL(CStr(Arr1(4))) & "', " & _
                                       " Gender = '" & FORMATSQL(CStr(Arr1(5))) & "', HandiCap = " & Arr1(6) & ", " & _
                                       " Class = '" & FORMATSQL(CStr(Arr1(7))) & "', AllowedTeam = " & Arr1(8) & ", " & _
                                       " iIndex = " & Arr1(9) & " " & _
                                       " WHERE (PK = " & Arr1(0) & ")"
                    End If
                    rt.Close
            End Select
        Loop
    Close #1
    Screen.MousePointer = vbDefault
Else
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    'Microsoft.ACE.OLEDB.12.0;
    'cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cn.Provider = "Microsoft.ACE.OLEDB.12.0;"
    cn.ConnectionString = "Data Source=" & Trim(txtPath.Text) & ";" & _
                          "Extended Properties=Excel 8.0;"
    cn.CursorLocation = adUseClient
    If cn.State = adStateOpen Then cn.Close
    cn.Open
    
    'Location
    Set rs = New ADODB.Recordset
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT * FROM [Location$] ", cn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        t = "SELECT tbl_Scoring_Location.* " & _
            " FROM tbl_Scoring_Location " & _
            " WHERE (PK = " & rs!PK & ")"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnRS
        If rt.RecordCount = 0 Then
            ConnRS.Execute "INSERT INTO tbl_Scoring_Location " & _
                           " (PK, ScoringLocation) " & _
                           " VALUES (" & rs!PK & ", '" & FORMATSQL(rs!ScoringLocation) & "')"
        Else
            ConnRS.Execute "UPDATE tbl_Scoring_Location " & _
                           " SET ScoringLocation = '" & FORMATSQL(rs!ScoringLocation) & "' " & _
                           " WHERE (PK = " & rs!PK & ")"
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    'Location Details
    Set rs = New ADODB.Recordset
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT * FROM [Location_Details$] ", cn, adOpenDynamic, adLockOptimistic
    ConnRS.Execute "DELETE FROM tbl_Scoring_Location_Details"
    While Not rs.EOF
        ConnRS.Execute "INSERT INTO tbl_Scoring_Location_Details " & _
                       " (MasterKey, Line, Description, H1, H2, H3, H4, H5, H6, H7, H8, H9, H10, H11, H12, H13, H14, H15, H16, H17, H18) " & _
                       " VALUES (" & rs!MasterKey & ", " & rs!Line & ", '" & FORMATSQL(rs!Description) & "', " & rs!H1 & ", " & rs!H2 & ", " & rs!H3 & ", " & _
                       " " & rs!H4 & ", " & rs!H5 & ", " & rs!H6 & ", " & rs!H7 & ", " & rs!H8 & ", " & rs!H9 & ", " & rs!H10 & ", " & rs!H11 & ", " & rs!H12 & ", " & _
                       " " & rs!H13 & ", " & rs!H14 & ", " & rs!H15 & ", " & rs!H16 & ", " & rs!H17 & ", " & rs!H18 & ")"
        rs.MoveNext
    Wend
    rs.Close
    
    'Scoring System
    Set rs = New ADODB.Recordset
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT * FROM [Scoring_System$] ", cn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        t = "SELECT tbl_Scoring_System.* " & _
            " FROM tbl_Scoring_System " & _
            " WHERE (PK = " & rs!PK & ")"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnRS
        If rt.RecordCount = 0 Then
            ConnRS.Execute "INSERT INTO tbl_Scoring_System " & _
                           " (PK, ScoringSystem) " & _
                           " VALUES (" & rs!PK & ", '" & FORMATSQL(rs!ScoringSystem) & "')"
        Else
            ConnRS.Execute "UPDATE tbl_Scoring_System " & _
                           " SET ScoringSystem = '" & FORMATSQL(rs!ScoringSystem) & "' " & _
                           " WHERE (PK = " & rs!PK & ")"
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    'Tournament Info
    Set rs = New ADODB.Recordset
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT * FROM [TournamentInfo$] ", cn, adOpenDynamic, adLockOptimistic
    If rs.RecordCount > 0 Then
    'While Not rs.EOF
        t = "SELECT tbl_Scoring_TournamentInfo.* " & _
            " FROM tbl_Scoring_TournamentInfo " & _
            " WHERE (PK = " & rs!PK & ")"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnRS
        If rt.RecordCount = 0 Then
            ConnRS.Execute "INSERT INTO tbl_Scoring_TournamentInfo " & _
                           " (PK, TournamentName, TournamentStart, TournamentEnd, TournamentDays, NoofPlays, NoofPlayerPerTeam, Remarks, Locked, Activated, Scoring, TeamPlay, " & _
                           " PlayerToCount, HandicapDivisor, PointsToCountTeam, TeamAverage, IndividualPlay, AllowTeamPerPlayer, PointsToCountIndi, ParGrossPoints) " & _
                           " VALUES (" & rs!PK & ", '" & FORMATSQL(rs!TournamentName) & "', '" & FormatDateTime(rs!TournamentStart, vbShortDate) & "','" & FormatDateTime(rs!TournamentEnd, vbShortDate) & "', " & _
                           " " & rs!TournamentDays & ", " & rs!NoofPlays & ", " & rs!NoofPlayerPerTeam & ", '" & FORMATSQL(IIf(IsNull(rs!Remarks), "", rs!Remarks)) & "', " & rs!Locked & ", " & rs!Activated & ", " & rs!Scoring & ", " & _
                           " " & rs!TeamPlay & ", " & rs!PlayerToCount & ", " & rs!HandicapDivisor & ", " & rs!PointsToCountTeam & ", " & rs!TeamAverage & ", " & rs!IndividualPlay & ", " & _
                           " " & rs!AllowTeamPerPlayer & ", " & rs!PointsToCountIndi & ", " & rs!ParGrossPoints & ")"
        Else
            ConnRS.Execute "UPDATE tbl_Scoring_TournamentInfo " & _
                           " SET TournamentName = '" & FORMATSQL(rs!TournamentName) & "', TournamentStart = '" & FormatDateTime(rs!TournamentStart, vbShortDate) & "', " & _
                           " TournamentEnd = '" & FormatDateTime(rs!TournamentEnd, vbShortDate) & "', TournamentDays = " & rs!TournamentDays & ", NoofPlays = " & rs!NoofPlays & ", " & _
                           " NoofPlayerPerTeam = " & rs!NoofPlayerPerTeam & ", Remarks = '" & FORMATSQL(IIf(IsNull(rs!Remarks), "", rs!Remarks)) & "', Locked = " & rs!Locked & ", Activated = " & rs!Activated & ", " & _
                           " Scoring = " & rs!Scoring & ", TeamPlay = " & rs!TeamPlay & ", PlayerToCount = " & rs!PlayerToCount & ", HandicapDivisor = " & rs!HandicapDivisor & ", " & _
                           " PointsToCountTeam = " & rs!PointsToCountTeam & ", TeamAverage = " & rs!TeamAverage & ", IndividualPlay = " & rs!IndividualPlay & ", " & _
                           " AllowTeamPerPlayer = " & rs!AllowTeamPerPlayer & ", PointsToCountIndi = " & rs!PointsToCountIndi & ", ParGrossPoints = " & rs!ParGrossPoints & " " & _
                           " WHERE (PK = " & rs!PK & ")"
        End If
        rt.Close
        t = "SELECT tbl_Scoring_TournamentInfo.* " & _
            " FROM tbl_Scoring_TournamentInfo " & _
            " WHERE (PK <> " & rs!PK & ")"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnRS
        While Not rt.EOF
            ConnRS.Execute "UPDATE tbl_Scoring_TournamentInfo " & _
                           " SET Activated = 0 " & _
                           " WHERE (PK = " & rt!PK & ")"
            rt.MoveNext
        Wend
        rt.Close
    '    rs.MoveNext
    'Wend
    End If
    rs.Close
    
    'Tournament Class
    Set rs = New ADODB.Recordset
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT * FROM [TournamentInfo_Class$] ", cn, adOpenDynamic, adLockOptimistic
    ConnRS.Execute "DELETE FROM tbl_Scoring_TournamentInfo_Class"
    While Not rs.EOF
        ConnRS.Execute "INSERT INTO tbl_Scoring_TournamentInfo_Class " & _
                       " (TournamentKey, Class, HFrom, HTo) " & _
                       " VALUES (" & rs!TournamentKey & ", '" & rs!Class & "', " & rs!HFrom & ", " & rs!HTo & ")"
        rs.MoveNext
    Wend
    rs.Close
    
    'Tournament Index
    Set rs = New ADODB.Recordset
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT * FROM [TournamentInfo_Index$] ", cn, adOpenDynamic, adLockOptimistic
    ConnRS.Execute "DELETE FROM tbl_Scoring_TournamentInfo_Index"
    While Not rs.EOF
        ConnRS.Execute "INSERT INTO tbl_Scoring_TournamentInfo_Index " & _
                       " (TournamentKey, Class, HFrom, HTo) " & _
                       " VALUES (" & rs!TournamentKey & ", '" & rs!Class & "', " & rs!HFrom & ", " & rs!HTo & ")"
        rs.MoveNext
    Wend
    rs.Close
    
    'Tournament Location
    Set rs = New ADODB.Recordset
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT * FROM [TournamentInfo_Location$] ", cn, adOpenDynamic, adLockOptimistic
    ConnRS.Execute "DELETE FROM tbl_Scoring_TournamentInfo_Location"
    While Not rs.EOF
        ConnRS.Execute "INSERT INTO tbl_Scoring_TournamentInfo_Location " & _
                       " (MasterKey, LocationKey) " & _
                       " VALUES (" & rs!MasterKey & ", " & rs!LocationKey & ")"
        rs.MoveNext
    Wend
    rs.Close
    
    'Player Name
    Set rs = New ADODB.Recordset
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT * FROM [PlayerName$] ", cn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        t = "SELECT tbl_Scoring_PlayerName.* " & _
            " FROM tbl_Scoring_PlayerName " & _
            " WHERE (PK = " & rs!PK & ")"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnRS
        If rt.RecordCount = 0 Then
            ConnRS.Execute "INSERT INTO tbl_Scoring_PlayerName " & _
                           " (PK, TournamentKey, LastName, FirstName, MiddleName, Gender, HandiCap, Class, AllowedTeam, iIndex) " & _
                           " VALUES (" & rs!PK & ", " & rs!TournamentKey & ", '" & FORMATSQL(rs!LastName) & "', '" & FORMATSQL(rs!FirstName) & "', " & _
                           " '" & FORMATSQL(rs!MiddleName) & "', '" & FORMATSQL(rs!Gender) & "', " & rs!HandiCap & ", '" & FORMATSQL(rs!Class) & "', " & _
                           " " & rs!AllowedTeam & ", " & rs!iIndex & ")"
        Else
            ConnRS.Execute "UPDATE tbl_Scoring_PlayerName " & _
                           " SET TournamentKey = " & rs!TournamentKey & ", LastName = '" & FORMATSQL(rs!LastName) & "', " & _
                           " FirstName = '" & FORMATSQL(rs!FirstName) & "', MiddleName = '" & FORMATSQL(rs!MiddleName) & "', " & _
                           " Gender = '" & FORMATSQL(rs!Gender) & "', HandiCap = " & rs!HandiCap & ", " & _
                           " Class = '" & FORMATSQL(rs!Class) & "', AllowedTeam = " & rs!AllowedTeam & ", " & _
                           " iIndex = " & rs!iIndex & " " & _
                           " WHERE (PK = " & rs!PK & ")"
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    
    If cn.State = adStateOpen Then cn.Close

    Screen.MousePointer = vbDefault
End If

MsgBox "Successfully Imported!                      ", vbInformation, "Import"

End Sub
