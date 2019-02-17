VERSION 5.00
Begin VB.Form frmProgressBar 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3735
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
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerProfile 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   720
   End
   Begin RPVGCC.b8Container picAlphalist 
      Height          =   1815
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3201
      BackColor       =   15396057
      Begin VB.CommandButton cmdOKAlphalist 
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
         Left            =   240
         Picture         =   "frmProgressBar.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1080
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancelAlphalist 
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
         Left            =   1920
         Picture         =   "frmProgressBar.frx":0672
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1080
         Width           =   1560
      End
      Begin VB.TextBox txtAsof 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin RPVGCC.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   45
         TabIndex        =   6
         Top             =   45
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   609
         Caption         =   "Alpha List"
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
         Icon            =   "frmProgressBar.frx":0DCE
         ShadowVisible   =   0   'False
      End
      Begin VB.Label Label89 
         BackStyle       =   0  'Transparent
         Caption         =   "As of"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
   End
   Begin RPVGCC.b8Container picProgressReport 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1720
      BackColor       =   15266266
      Begin VB.Timer TimerActive 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   5640
         Top             =   720
      End
      Begin VB.Timer TimerInactive 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   5160
         Top             =   720
      End
      Begin VB.Timer TimerHeadCount 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   4680
         Top             =   720
      End
      Begin VB.Timer TimerAlphaActive 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   4200
         Top             =   720
      End
      Begin VB.Timer TimerHistory 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   3720
         Top             =   720
      End
      Begin VB.PictureBox picProgress 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5955
         TabIndex        =   1
         Top             =   120
         Width           =   6015
      End
   End
End
Attribute VB_Name = "frmProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FileName_xls As String
Dim WorkbookName As String
Dim iWorkSheet As Integer

Public iEmployee

Dim i, l, TableName, DetailTableName, Columns, ColumnsDet, Clustered, j, k, _
RowCnt, ColCnt, strRange, Arr, Arr1, iTotRecord, iPK, sPosition, sDateHired, _
sIDNumber, sDepartment, sLevel, sDateMarriage, sSpouseBDay, sFatherBDay, _
sMotherBDay, sChildBDay, sBrotherSisterBDay

Private Sub b8TitleBar1_CLoseClick()
cmdCancelAlphalist_Click
End Sub

Private Sub cmdCancelAlphalist_Click()
Unload Me
End Sub

Private Sub cmdOKAlphalist_Click()
If IsDate(txtAsof.Text) = False Then MsgBox "Please Supply a Valid Date!                  ", vbCritical, "Error...": txtAsof.SetFocus: Exit Sub

MainForm.CommonDialog1.CancelError = True
On Error GoTo ErrorHandler
MainForm.CommonDialog1.DialogTitle = "Save"
MainForm.CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
MainForm.CommonDialog1.ShowSave
FileName_xls = Trim(MainForm.CommonDialog1.Filename)



txtAsof.Text = Format(FormatDateTime(txtAsof.Text, vbShortDate), "mm/dd/yyyy")
picAlphalist.Visible = False
'picToolbar.Enabled = False
'picMain.Enabled = False

Me.Width = 6260
Me.Height = 970

picProgress.BackColor = &HFFFFFF
picProgressReport.ZOrder 0
picProgressReport.Visible = True
DoEvents
TimerAlphaActive.Enabled = True

Exit Sub
ErrorHandler:
Exit Sub
End Sub

Private Sub TimerActive_Timer()
TimerActive.Enabled = False
i = 0
picProgress.BackColor = &HFFFFFF
ConnOmega.Execute "DELETE FROM tbl_Personnel_Active_Inactive_Report WHERE (LogInName = '" & gbl_UserName & "')"
's = "sp_Personnel_Active_Inactive_Report (1,'" & FormatDateTime(Date, vbShortDate) & "')"
s = "SELECT dbo.tbl_Personnel_IDNumber.PK, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName + ',  ' + dbo.tbl_Personnel_Information.FirstName + '  ' + dbo.tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
    " ISNULL((SELECT TOP (1) PK From dbo.tbl_Personnel_ActionNew WHERE (EmpPK = dbo.tbl_Personnel_IDNumber.PK) AND (EffectivityDate <= '" & FormatDateTime(Date, vbShortDate) & "') ORDER BY EffectivityDate DESC), 0) AS ActionMemo " & _
    " FROM  dbo.tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
    " WHERE ((SELECT TOP (1) tbl_Personnel_EmploymentStatus_1.Active FROM  dbo.tbl_Personnel_ActionNew AS tbl_Personnel_ActionNew_1 LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_EmploymentStatus AS tbl_Personnel_EmploymentStatus_1 ON tbl_Personnel_ActionNew_1.EmpStatusKey = tbl_Personnel_EmploymentStatus_1.PK " & _
    " WHERE (tbl_Personnel_ActionNew_1.EmpPK = dbo.tbl_Personnel_IDNumber.PK) AND (tbl_Personnel_ActionNew_1.EffectivityDate <= '" & FormatDateTime(Date, vbShortDate) & "') ORDER BY tbl_Personnel_ActionNew_1.EffectivityDate DESC) = 1)"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount = 0 Then Exit Sub
DoEvents
While Not rs.EOF
    DoEvents
    i = i + 1
    
    't = "SELECT tbl_Personnel_Action.Division, tbl_Personnel_Action.Dept, tbl_Personnel_Department.DepartmentName, " & _
        " tbl_Personnel_Action.EmpStatus, tbl_Personnel_EmploymentStatus.StatusName, tbl_Personnel_Action.Positions, " & _
        " tbl_Personnel_Position.PositionName , tbl_Personnel_Action.EffectivityDate, tbl_Personnel_Action.Remarks " & _
        " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
        " tbl_Personnel_Position ON tbl_Personnel_Action.Positions = tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_Department ON tbl_Personnel_Action.Dept = tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_EmploymentStatus ON tbl_Personnel_Action.EmpStatus = tbl_Personnel_EmploymentStatus.PK " & _
        " WHERE (tbl_Personnel_Action.PK = " & rs!ActionMemo & ")"
    t = "SELECT dbo.tbl_Personnel_ActionNew.DivisionKey AS Division, dbo.tbl_Personnel_Division.Description AS DivisionName, " & _
        " dbo.tbl_Personnel_ActionNew.DeptKey AS Dept, dbo.tbl_Personnel_Department.DepartmentName, dbo.tbl_Personnel_ActionNew.EmpStatusKey AS EmpStatus, " & _
        " dbo.tbl_Personnel_EmploymentStatus.StatusName, dbo.tbl_Personnel_ActionNew.PositionsKey AS Positions, " & _
        " dbo.tbl_Personnel_Position.PositionName, dbo.tbl_Personnel_ActionNew.EffectivityDate, dbo.tbl_Personnel_ActionNew.Remarks " & _
        " FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_ActionNew.DivisionKey = dbo.tbl_Personnel_Division.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_EmploymentStatus ON dbo.tbl_Personnel_ActionNew.EmpStatusKey = dbo.tbl_Personnel_EmploymentStatus.PK " & _
        " WHERE (dbo.tbl_Personnel_ActionNew.PK = " & rs!ActionMemo & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Active_Inactive_Report " & _
                          " (LogInName, Division, DivisionName, Department, DepartmentName, StatusKey, StatusName, " & _
                          " PositionKey, PositionName, EmpKey, IDNumber, EmployeeName) " & _
                          " VALUES ('" & gbl_UserName & "', " & rt!Division & ", '" & FORMATSQL(rt!DivisionName) & "', " & _
                          " " & rt!Dept & ", '" & FORMATSQL(rt!DepartmentName) & "', " & rt!EmpStatus & ", " & _
                          " '" & FORMATSQL(rt!StatusName) & "', " & rt!Positions & ", '" & FORMATSQL(rt!PositionName) & "', " & _
                          " " & rs!PK & ", '" & FORMATSQL(rs!IDNumber) & "', '" & FORMATSQL(rs!EmployeeName) & "')"
    End If
    rt.Close
    
    UpdateProgress picProgress, i / rs.RecordCount
    rs.MoveNext
Wend
rs.Close

s = "SELECT tbl_Personnel_Active_Inactive_Report.* " & _
    " FROM tbl_Personnel_Active_Inactive_Report " & _
    " WHERE (LogInName = '" & gbl_UserName & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
rs.Requery
rs.Close

Unload Me

frmCrystalReportViewer.PRINT_ACTIVE_EMPLOYEE gbl_CompanyName, gbl_UserName
If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show

End Sub

Private Sub TimerAlphaActive_Timer()
TimerAlphaActive.Enabled = False

WorkbookName = CStr(FileName_xls)

DoEvents
'picToolbar.Enabled = False
'picMain.Enabled = False
'picProgressReport.ZOrder 0
'picProgressReport.Visible = True

Screen.MousePointer = vbHourglass

iWorkSheet = 1
Set xlsApp = CreateObject("Excel.Application")
xlsApp.Visible = False
xlsApp.Workbooks.Add
xlsApp.DisplayAlerts = False
xlsApp.Workbooks(1).Sheets(2).Delete
xlsApp.Workbooks(1).Sheets(2).Delete
xlsApp.Workbooks(1).Sheets(iWorkSheet).Activate
xlsApp.Workbooks(1).Sheets(iWorkSheet).Name = "Alphalist"

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = gbl_CompanyName
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 12
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = gbl_CompanyAddress1
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 9
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True

RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = gbl_CompanyAddress2
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 9
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True


RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = "As of " & Format(txtAsof.Text, "mmmm dd, yyyy")
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 9
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True


RowCnt = RowCnt + 1
ColCnt = 0
ColCnt = ColCnt + 1
strRange = EXCEL_RANGE(ColCnt, RowCnt)
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 9
xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False


i = 0
picProgress.BackColor = &HFFFFFF
s = "sp_Personnel_Alphalist(1, '" & FormatDateTime(txtAsof.Text, vbShortDate) & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    RowCnt = RowCnt + 1
    ColCnt = 0
    ColCnt = ColCnt + 1
    strRange = EXCEL_RANGE(ColCnt, RowCnt)
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = ""
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
    xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
    
    For k = 1 To rs.Fields.Count
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs.Fields(k - 1).Name
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Arial"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = True
    Next k
    
    While Not rs.EOF
        DoEvents
        i = i + 1
        RowCnt = RowCnt + 1
        ColCnt = 0
        ColCnt = ColCnt + 1
        strRange = EXCEL_RANGE(ColCnt, RowCnt)
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = i
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
        xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
        For k = 1 To rs.Fields.Count
            ColCnt = ColCnt + 1
            strRange = EXCEL_RANGE(ColCnt, RowCnt)
            If IsNumeric(rs.Fields(k - 1).Value) Then
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "@"
            End If
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Value = rs.Fields(k - 1).Value
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Name = "Tahoma"
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Size = 10
            xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).Font.Bold = False
            If IsDate(rs.Fields(k - 1).Value) Then
                xlsApp.ActiveWorkbook.Sheets(iWorkSheet).Range(strRange).NumberFormat = "mm/dd/yyyy"
            End If
            
        Next k
        
        UpdateProgress picProgress, i / rs.RecordCount
        
        rs.MoveNext
    Wend
End If
rs.Close

SAVING:
On Error GoTo err_saving:
If InStr(WorkbookName, ".") = 0 Then WorkbookName = WorkbookName & ".xls"
xlsApp.ActiveWorkbook.SaveAs Filename:=WorkbookName

xlsApp.Visible = True

Screen.MousePointer = vbDefault

Unload Me

'picProgress.BackColor = &HFFFFFF
'picProgressReport.Visible = False
'picToolbar.Enabled = True
'picMain.Enabled = True

Exit Sub
err_saving:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Please Check if File Currently Open!              ", vbCritical, "Error..."
GoTo SAVING:

Exit Sub
ErrorHandler:
Screen.MousePointer = vbDefault
Exit Sub
End Sub

Private Sub TimerHeadCount_Timer()
TimerHeadCount.Enabled = False
i = 0
picProgress.BackColor = &HFFFFFF
ConnOmega.Execute "DELETE FROM tbl_Personnel_HeadCount WHERE (LogInName = '" & gbl_UserName & "')"
's = "sp_Personnel_Active_Inactive_Report (1,'" & FormatDateTime(Date, vbShortDate) & "')"
s = "SELECT dbo.tbl_Personnel_IDNumber.PK, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName + ',  ' + dbo.tbl_Personnel_Information.FirstName + '  ' + dbo.tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
    " ISNULL((SELECT TOP (1) PK From dbo.tbl_Personnel_ActionNew WHERE (EmpPK = dbo.tbl_Personnel_IDNumber.PK) AND (EffectivityDate <= '" & FormatDateTime(Date, vbShortDate) & "') ORDER BY EffectivityDate DESC), 0) AS ActionMemo " & _
    " FROM  dbo.tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
    " WHERE ((SELECT TOP (1) tbl_Personnel_EmploymentStatus_1.Active FROM  dbo.tbl_Personnel_ActionNew AS tbl_Personnel_ActionNew_1 LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_EmploymentStatus AS tbl_Personnel_EmploymentStatus_1 ON tbl_Personnel_ActionNew_1.EmpStatusKey = tbl_Personnel_EmploymentStatus_1.PK " & _
    " WHERE (tbl_Personnel_ActionNew_1.EmpPK = dbo.tbl_Personnel_IDNumber.PK) AND (tbl_Personnel_ActionNew_1.EffectivityDate <= '" & FormatDateTime(Date, vbShortDate) & "') ORDER BY tbl_Personnel_ActionNew_1.EffectivityDate DESC) = 1)"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount = 0 Then Exit Sub

DoEvents

TableName = "tmp_" & gbl_UserName & "_Personnel_HeadCount"

Columns = ""
Columns = Columns & "|DivisionKey:int:NOT NULL:DEFAULT(0)"
Columns = Columns & "|DepartmentKey:int:NOT NULL:DEFAULT(0)"
Columns = Columns & "|StatusKey:int:NOT NULL:DEFAULT(0)"
Columns = Columns & "|EmployeeKey:int:NOT NULL:DEFAULT(0)"
CreateTable gbl_Database, TableName, Columns

While Not rs.EOF
    DoEvents
    i = i + 1
    't = "SELECT tbl_Personnel_Action.Division, tbl_Personnel_Action.Dept, tbl_Personnel_Department.DepartmentName, " & _
        " tbl_Personnel_Action.EmpStatus, tbl_Personnel_EmploymentStatus.StatusName, tbl_Personnel_Action.Positions, " & _
        " tbl_Personnel_Position.PositionName , tbl_Personnel_Action.EffectivityDate, tbl_Personnel_Action.Remarks " & _
        " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
        " tbl_Personnel_Position ON tbl_Personnel_Action.Positions = tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_Department ON tbl_Personnel_Action.Dept = tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_EmploymentStatus ON tbl_Personnel_Action.EmpStatus = tbl_Personnel_EmploymentStatus.PK " & _
        " WHERE (tbl_Personnel_Action.PK = " & rs!ActionMemo & ")"
    t = "SELECT dbo.tbl_Personnel_ActionNew.DivisionKey AS Division, dbo.tbl_Personnel_Division.Description AS DivisionName, " & _
        " dbo.tbl_Personnel_ActionNew.DeptKey AS Dept, dbo.tbl_Personnel_Department.DepartmentName, dbo.tbl_Personnel_ActionNew.EmpStatusKey AS EmpStatus, " & _
        " dbo.tbl_Personnel_EmploymentStatus.StatusName, dbo.tbl_Personnel_ActionNew.PositionsKey AS Positions, " & _
        " dbo.tbl_Personnel_Position.PositionName, dbo.tbl_Personnel_ActionNew.EffectivityDate, dbo.tbl_Personnel_ActionNew.Remarks " & _
        " FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_ActionNew.DivisionKey = dbo.tbl_Personnel_Division.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_EmploymentStatus ON dbo.tbl_Personnel_ActionNew.EmpStatusKey = dbo.tbl_Personnel_EmploymentStatus.PK " & _
        " WHERE (dbo.tbl_Personnel_ActionNew.PK = " & rs!ActionMemo & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        ConnOmega.Execute "INSERT INTO " & TableName & " " & _
                          " (DivisionKey, DepartmentKey, StatusKey, EmployeeKey) " & _
                          " VALUES (" & rt!Division & ", " & rt!Dept & ", " & _
                          " " & rt!EmpStatus & ", " & rs!PK & ")"
    End If
    rt.Close
    
    UpdateProgress picProgress, i / rs.RecordCount
    rs.MoveNext
Wend
rs.Close

s = "SELECT " & TableName & ".DivisionKey, " & TableName & ".DepartmentKey, " & _
    " tbl_Personnel_Department.DepartmentName, " & TableName & ".StatusKey, " & _
    " tbl_Personnel_EmploymentStatus.StatusName, COUNT(" & TableName & ".EmployeeKey) AS HeadCount " & _
    " FROM " & TableName & " LEFT OUTER JOIN " & _
    " tbl_Personnel_EmploymentStatus ON " & _
    " " & TableName & ".StatusKey = tbl_Personnel_EmploymentStatus.PK LEFT OUTER JOIN " & _
    " tbl_Personnel_Department ON " & TableName & ".DepartmentKey = tbl_Personnel_Department.PK " & _
    " GROUP BY " & TableName & ".DivisionKey, " & TableName & ".DepartmentKey, " & _
    " " & TableName & ".StatusKey, tbl_Personnel_Department.DepartmentName, " & _
    " tbl_Personnel_EmploymentStatus.StatusName " & _
    " ORDER BY " & TableName & ".DivisionKey, " & TableName & ".DepartmentKey, " & _
    " tbl_Personnel_EmploymentStatus.StatusName "
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
While Not rs.EOF
    ConnOmega.Execute "INSERT INTO tbl_Personnel_HeadCount " & _
                      " (LogInName, DivKey, DivName, DeptKey, DeptName, StatusKey, StatusName, EmpCount) " & _
                      " VALUES ('" & gbl_UserName & "', " & rs!DivisionKey & ", " & _
                      " '" & IIf(rs!DivisionKey = 1, "CLUB HOUSE", "MAINTENANCE") & "', " & _
                      " " & rs!DepartmentKey & ", '" & FORMATSQL(rs!DepartmentName) & "', " & _
                      " " & rs!StatusKey & ", '" & FORMATSQL(rs!StatusName) & "', " & _
                      " " & rs!HeadCount & ")"
    rs.MoveNext
Wend
rs.Close


s = "SELECT tbl_Personnel_HeadCount.* " & _
    " FROM tbl_Personnel_HeadCount " & _
    " WHERE (LogInName = '" & gbl_UserName & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
rs.Requery
rs.Close

Unload Me

frmCrystalReportViewer.PRINT_EMPLOYEE_HEADCOUNT gbl_CompanyName, gbl_UserName
If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show

End Sub

Private Sub TimerInactive_Timer()
TimerInactive.Enabled = False
i = 0
picProgress.BackColor = &HFFFFFF
ConnOmega.Execute "DELETE FROM tbl_Personnel_Active_Inactive_Report WHERE (LogInName = '" & gbl_UserName & "')"
's = "sp_Personnel_Active_Inactive_Report (2,'" & FormatDateTime(Date, vbShortDate) & "')"
s = "SELECT dbo.tbl_Personnel_IDNumber.PK, dbo.tbl_Personnel_IDNumber.IDNumber, dbo.tbl_Personnel_Information.LastName + ',  ' + dbo.tbl_Personnel_Information.FirstName + '  ' + dbo.tbl_Personnel_Information.MiddleName AS EmployeeName, " & _
    " ISNULL((SELECT TOP (1) PK From dbo.tbl_Personnel_ActionNew WHERE (EmpPK = dbo.tbl_Personnel_IDNumber.PK) AND (EffectivityDate <= '" & FormatDateTime(Date, vbShortDate) & "') ORDER BY EffectivityDate DESC), 0) AS ActionMemo " & _
    " FROM  dbo.tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_Information ON dbo.tbl_Personnel_IDNumber.ProfileKey = dbo.tbl_Personnel_Information.PK " & _
    " WHERE ((SELECT TOP (1) tbl_Personnel_EmploymentStatus_1.Active FROM  dbo.tbl_Personnel_ActionNew AS tbl_Personnel_ActionNew_1 LEFT OUTER JOIN " & _
    " dbo.tbl_Personnel_EmploymentStatus AS tbl_Personnel_EmploymentStatus_1 ON tbl_Personnel_ActionNew_1.EmpStatusKey = tbl_Personnel_EmploymentStatus_1.PK " & _
    " WHERE (tbl_Personnel_ActionNew_1.EmpPK = dbo.tbl_Personnel_IDNumber.PK) AND (tbl_Personnel_ActionNew_1.EffectivityDate <= '" & FormatDateTime(Date, vbShortDate) & "') ORDER BY tbl_Personnel_ActionNew_1.EffectivityDate DESC) = 2)"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount = 0 Then Exit Sub
DoEvents

While Not rs.EOF
    DoEvents
    i = i + 1
    
    't = "SELECT tbl_Personnel_Action.Division, tbl_Personnel_Action.Dept, tbl_Personnel_Department.DepartmentName, " & _
        " tbl_Personnel_Action.EmpStatus, tbl_Personnel_EmploymentStatus.StatusName, tbl_Personnel_Action.Positions, " & _
        " tbl_Personnel_Position.PositionName , tbl_Personnel_Action.EffectivityDate, tbl_Personnel_Action.Remarks " & _
        " FROM tbl_Personnel_Action LEFT OUTER JOIN " & _
        " tbl_Personnel_Position ON tbl_Personnel_Action.Positions = tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_Department ON tbl_Personnel_Action.Dept = tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
        " tbl_Personnel_EmploymentStatus ON tbl_Personnel_Action.EmpStatus = tbl_Personnel_EmploymentStatus.PK " & _
        " WHERE (tbl_Personnel_Action.PK = " & rs!ActionMemo & ")"
    t = "SELECT dbo.tbl_Personnel_ActionNew.DivisionKey AS Division, dbo.tbl_Personnel_Division.Description AS DivisionName, " & _
        " dbo.tbl_Personnel_ActionNew.DeptKey AS Dept, dbo.tbl_Personnel_Department.DepartmentName, dbo.tbl_Personnel_ActionNew.EmpStatusKey AS EmpStatus, " & _
        " dbo.tbl_Personnel_EmploymentStatus.StatusName, dbo.tbl_Personnel_ActionNew.PositionsKey AS Positions, " & _
        " dbo.tbl_Personnel_Position.PositionName, dbo.tbl_Personnel_ActionNew.EffectivityDate, dbo.tbl_Personnel_ActionNew.Remarks " & _
        " FROM  dbo.tbl_Personnel_ActionNew LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Position ON dbo.tbl_Personnel_ActionNew.PositionsKey = dbo.tbl_Personnel_Position.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Department ON dbo.tbl_Personnel_ActionNew.DeptKey = dbo.tbl_Personnel_Department.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_Division ON dbo.tbl_Personnel_ActionNew.DivisionKey = dbo.tbl_Personnel_Division.PK LEFT OUTER JOIN " & _
        " dbo.tbl_Personnel_EmploymentStatus ON dbo.tbl_Personnel_ActionNew.EmpStatusKey = dbo.tbl_Personnel_EmploymentStatus.PK " & _
        " WHERE (dbo.tbl_Personnel_ActionNew.PK = " & rs!ActionMemo & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Active_Inactive_Report " & _
                          " (LogInName, Division, DivisionName, Department, DepartmentName, StatusKey, StatusName, " & _
                          " PositionKey, PositionName, EmpKey, IDNumber, EmployeeName, EffecDate, Reason) " & _
                          " VALUES ('" & gbl_UserName & "', " & rt!Division & ", '" & FORMATSQL(rt!DivisionName) & "', " & _
                          " " & rt!Dept & ", '" & FORMATSQL(rt!DepartmentName) & "', " & rt!EmpStatus & ", " & _
                          " '" & FORMATSQL(rt!StatusName) & "', " & rt!Positions & ", '" & FORMATSQL(rt!PositionName) & "', " & _
                          " " & rs!PK & ", '" & FORMATSQL(rs!IDNumber) & "', '" & FORMATSQL(rs!EmployeeName) & "', " & _
                          " '" & FormatDateTime(rt!EffectivityDate, vbShortDate) & "', '" & FORMATSQL(rt!Remarks) & "')"
    End If
    rt.Close
    
    UpdateProgress picProgress, i / rs.RecordCount
    rs.MoveNext
Wend
rs.Close

'picProgressReport.Visible = False
'picToolbar.Enabled = True
'picMain.Enabled = True

s = "SELECT tbl_Personnel_Active_Inactive_Report.* " & _
    " FROM tbl_Personnel_Active_Inactive_Report " & _
    " WHERE (LogInName = '" & gbl_UserName & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
rs.Requery
rs.Close

Unload Me

frmCrystalReportViewer.PRINT_INACTIVE_EMPLOYEE gbl_CompanyName, gbl_UserName
If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show

End Sub



Private Sub TimerProfile_Timer()
TimerProfile.Enabled = False

picProgress.BackColor = &HFFFFFF

CREATE_PROFILE_DATASHEET_TABLE "tbl_Personnel_DataSheet"

ConnOmega.Execute "DELETE FROM tbl_Personnel_DataSheet WHERE (LogInName = '" & gbl_UserName & "')"

iTotRecord = 1: i = 0
s = "SELECT tbl_Personnel_Information.* " & _
    " FROM tbl_Personnel_Information " & _
    " WHERE (PK = " & iEmployee & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    t = "SELECT tbl_Personnel_BrotherSister.* " & _
        " FROM tbl_Personnel_BrotherSister " & _
        " WHERE (ProfileKey = " & rs!PK & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    iTotRecord = iTotRecord + rt.RecordCount
    rt.Close
    
    t = "SELECT tbl_Personnel_Children.* " & _
        " FROM tbl_Personnel_Children " & _
        " WHERE (ProfileKey = " & rs!PK & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    iTotRecord = iTotRecord + rt.RecordCount
    rt.Close
    
    t = "SELECT tbl_Personnel_Education.* " & _
        " FROM tbl_Personnel_Education " & _
        " WHERE (ProfileKey = " & rs!PK & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    iTotRecord = iTotRecord + rt.RecordCount
    rt.Close
    
    t = "SELECT tbl_Personnel_Employment.* " & _
        " FROM tbl_Personnel_Employment " & _
        " WHERE (ProfileKey = " & rs!PK & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    iTotRecord = iTotRecord + rt.RecordCount
    rt.Close
    
    
    
    sPosition = ""
    t = "SELECT TOP 1 tbl_Personnel_Position.PositionName " & _
        " FROM tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
        " tbl_Personnel_Action ON tbl_Personnel_IDNumber.PK = tbl_Personnel_Action.EmpPK LEFT OUTER JOIN " & _
        " tbl_Personnel_Position ON tbl_Personnel_Action.Positions = tbl_Personnel_Position.PK " & _
        " WHERE (tbl_Personnel_IDNumber.ProfileKey = " & rs!PK & ") " & _
        " AND (tbl_Personnel_Action.EffectivityDate <= '" & FormatDateTime(Date, vbShortDate) & "') " & _
        " ORDER BY tbl_Personnel_Action.EffectivityDate DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        sPosition = UCase(rt!PositionName)
    End If
    rt.Close
    
    sDateHired = "": sIDNumber = ""
    t = "SELECT TOP 1 DateHired, IDNumber " & _
        " From tbl_Personnel_IDNumber " & _
        " Where (ProfileKey = " & rs!PK & ") " & _
        " ORDER BY DateHired DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        sDateHired = Format(rt!DateHired, "mm/dd/yyyy")
        sIDNumber = rt!IDNumber
    End If
    rt.Close
    
    sDepartment = ""
    t = "SELECT TOP 1 tbl_GL_Department.DeptName " & _
        " FROM tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
        " tbl_Personnel_Action ON tbl_Personnel_IDNumber.PK = tbl_Personnel_Action.EmpPK LEFT OUTER JOIN " & _
        " tbl_GL_Department ON tbl_Personnel_Action.Dept = tbl_GL_Department.PK " & _
        " WHERE (tbl_Personnel_IDNumber.ProfileKey = " & rs!PK & ") " & _
        " AND (tbl_Personnel_Action.EffectivityDate <= '" & FormatDateTime(Date, vbShortDate) & "') " & _
        " ORDER BY tbl_Personnel_Action.EffectivityDate DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        sDepartment = UCase(rt!DeptName)
    End If
    rt.Close
    
    sLevel = ""
    t = "SELECT TOP 1 tbl_Personnel_Position.PositionLevel " & _
        " FROM tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
        " tbl_Personnel_Action ON tbl_Personnel_IDNumber.PK = tbl_Personnel_Action.EmpPK LEFT OUTER JOIN " & _
        " tbl_Personnel_Position ON tbl_Personnel_Action.Positions = tbl_Personnel_Position.PK " & _
        " WHERE (tbl_Personnel_IDNumber.ProfileKey = " & rs!PK & ") " & _
        " AND (tbl_Personnel_Action.EffectivityDate <= '" & FormatDateTime(Date, vbShortDate) & "') " & _
        " ORDER BY tbl_Personnel_Action.EffectivityDate DESC"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        sLevel = IIf(rt!PositionLevel = 1, "RANK IN FILE", IIf(rt!PositionLevel = 2, "SUPERVISORY", ""))
    End If
    rt.Close
    
    If IsNull(rs!DateMarriage) = True Then
        sDateMarriage = ""
    Else
        If DateValue(rs!DateMarriage) = DateValue("01/01/1900") Then
            sDateMarriage = ""
        Else
            sDateMarriage = Format(rs!DateMarriage, "mm/dd/yyyy")
        End If
    End If
    
    If IsNull(rs!SpouseBDay) = True Then
        sSpouseBDay = ""
    Else
        If DateValue(rs!SpouseBDay) = DateValue("01/01/1900") Then
            sSpouseBDay = ""
        Else
            sSpouseBDay = Format(rs!SpouseBDay, "mm/dd/yyyy")
        End If
    End If
    
    If IsNull(rs!FatherBDay) = True Then
        sFatherBDay = ""
    Else
        If DateValue(rs!FatherBDay) = DateValue("01/01/1900") Then
            sFatherBDay = ""
        Else
            sFatherBDay = Format(rs!FatherBDay, "mm/dd/yyyy")
        End If
    End If
    
    If IsNull(rs!MotherBDay) = True Then
        sMotherBDay = ""
    Else
        If DateValue(rs!MotherBDay) = DateValue("01/01/1900") Then
            sMotherBDay = ""
        Else
            sMotherBDay = Format(rs!MotherBDay, "mm/dd/yyyy")
        End If
    End If
    
    i = i + 1
    ConnOmega.Execute "INSERT INTO tbl_Personnel_DataSheet " & _
                      " (LogInName, Positions, DateHired, Department, Levels, Name, PresentAddress, OwnedHouse, Rent, BirthDate, Age, BirthPlace, Religion, LivingParents, CivilStatus, " & _
                      " DateMarriage, Height, Weight, Nationality, SSSNumber, TIN, DriversLicense, PHICNumber, PagIbigNumber, IDNumber, SpouseName, SpouseBirthDate, " & _
                      " SpouseOccupation, SpouseAddress, FatherName, FatherBirthDate, FatherOccupation, FatherAddress, MotherName, MotherBirthDate, MotherOccupation, " & _
                      " MotherAddress, Skills, OrgClub, RelatedName, RelatedContact, RelatedAddress, RelativeName, RelativeContact, RelativeAddress, " & _
                      " EmergencyName , EmergencyAddress, EmergencyRelation, EmergencyContact) " & _
                      " VALUES ('" & gbl_UserName & "', '" & FORMATSQL(CStr(sPosition)) & "', '" & FORMATSQL(CStr(sDateHired)) & "', '" & FORMATSQL(CStr(sDepartment)) & "', " & _
                      " '" & FORMATSQL(CStr(sLevel)) & "', '" & FORMATSQL(rs!LastName & ",  " & rs!FirstName & "  " & rs!MiddleName) & "', '" & FORMATSQL(rs!PresentAddress) & "', " & _
                      " '" & IIf(rs!OwnedHouse = 1, "NO", IIf(rs!OwnedHouse = 2, "YES", "")) & "', '" & IIf(rs!Rented = 1, "NO", IIf(rs!Rented = 2, "YES", "")) & "', " & _
                      " '" & Format(rs!BirthDate, "mm/dd/yyyy") & "','" & CStr(Get_Age(FormatDateTime(rs!BirthDate, vbShortDate), FormatDateTime(Date, vbShortDate))) & "', " & _
                      " '" & FORMATSQL(rs!BirthPlace) & "', '" & FORMATSQL(rs!Religion) & "', '" & IIf(rs!LivingWParents = 1, "NO", IIf(rs!LivingWParents = 2, "YES", "")) & "', " & _
                      " '" & IIf(rs!CivilStatus = 1, "SINGLE", IIf(rs!CivilStatus = 2, "MARRIED", IIf(rs!CivilStatus = 3, "WIDOWED", IIf(rs!CivilStatus = 4, "WIDOWER", "")))) & "', " & _
                      " '" & sDateMarriage & "', '" & FORMATSQL(rs!Height) & "', '" & rs!Weight & "', '" & FORMATSQL(rs!Nationality) & "', '" & FORMATSQL(rs!SSSNumber) & "', " & _
                      " '" & FORMATSQL(rs!TIN) & "','" & FORMATSQL(rs!DriverLicense) & "', '" & FORMATSQL(rs!PHICNumber) & "', '" & FORMATSQL(rs!HDMFNumber) & "', " & _
                      " '" & FORMATSQL(CStr(sIDNumber)) & "', '" & FORMATSQL(rs!SpouseName) & "', '" & FORMATSQL(CStr(sSpouseBDay)) & "', '" & FORMATSQL(rs!SpouseOccupation) & "', " & _
                      " '" & FORMATSQL(rs!SpouseAddress) & "', '" & FORMATSQL(rs!FatherName) & "', '" & sFatherBDay & "', '" & FORMATSQL(rs!FatherOccupation) & "', " & _
                      " '" & FORMATSQL(rs!FatherAddress) & "', '" & FORMATSQL(rs!MotherName) & "', '" & sMotherBDay & "', '" & FORMATSQL(rs!MotherOccupation) & "', " & _
                      " '" & FORMATSQL(rs!MotherAddress) & "', '" & FORMATSQL(rs!Skills) & "', '" & FORMATSQL(rs!OrganizationClubs) & "', '" & FORMATSQL(rs!RefName) & "', " & _
                      " '" & FORMATSQL(rs!RefContact) & "', '" & FORMATSQL(rs!RefAddress) & "', '" & FORMATSQL(rs!RefCompName) & "', '" & FORMATSQL(rs!RefCompContact) & "', " & _
                      " '" & FORMATSQL(rs!RefCompAddress) & "', '" & FORMATSQL(rs!EmergencyName) & "', '" & FORMATSQL(rs!EmergencyAddress) & "','" & FORMATSQL(rs!EmergencyRelation) & "', " & _
                      " '" & FORMATSQL(rs!EmergencyContact) & "')"
    UpdateProgress picProgress, i / iTotRecord
    
    iPK = 0
    t = "SELECT tbl_Personnel_DataSheet.* " & _
        " FROM tbl_Personnel_DataSheet " & _
        " WHERE (LogInName = '" & gbl_UserName & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        iPK = rt!PK
    End If
    rt.Close
    
    If CDbl(iPK) <> 0 Then
        l = 0
        t = "SELECT tbl_Personnel_BrotherSister.* " & _
            " FROM tbl_Personnel_BrotherSister " & _
            " WHERE (ProfileKey = " & rs!PK & ") " & _
            " ORDER BY Line"
        If rt.State = adStateOpen Then rs.Close
        rt.Open t, ConnOmega
        While Not rt.EOF
            i = i + 1
            l = l + 1
            If IsNull(rt!BrotherSisterBDay) = True Then
                sBrotherSisterBDay = ""
            Else
                If DateValue(rt!BrotherSisterBDay) = DateValue("01/01/1900") Then
                    sBrotherSisterBDay = ""
                Else
                    sBrotherSisterBDay = Format(rt!BrotherSisterBDay, "mm/dd/yyyy")
                End If
            End If
            
            ConnOmega.Execute "INSERT INTO tbl_Personnel_DataSheet_Sibling " & _
                              " (MasterKey, Line, FullName, BirthDate, Occupation, Address) " & _
                              " VALUES (" & iPK & ", " & l & ", '" & FORMATSQL(rt!BrotherSisterName) & "', " & _
                              " '" & sBrotherSisterBDay & "', '" & FORMATSQL(rt!BrotherSisterOccupation) & "', " & _
                              " '" & FORMATSQL(rt!BrotherSisterAddress) & "')"
            UpdateProgress picProgress, i / iTotRecord
            rt.MoveNext
        Wend
        rt.Close
        
        ConnOmega.Execute "UPDATE tbl_Personnel_DataSheet SET Sibling = " & CDbl(l) & " WHERE (PK = " & iPK & ")"
        
        l = 0
        t = "SELECT tbl_Personnel_Children.* " & _
            " FROM tbl_Personnel_Children " & _
            " WHERE (ProfileKey = " & rs!PK & ") " & _
            " ORDER BY Line"
        If rt.State = adStateOpen Then rs.Close
        rt.Open t, ConnOmega
        While Not rt.EOF
            i = i + 1
            l = l + 1
            If IsNull(rt!ChildBDay) = True Then
                sChildBDay = ""
            Else
                If DateValue(rt!ChildBDay) = DateValue("01/01/1900") Then
                    sChildBDay = ""
                Else
                    sChildBDay = Format(rt!ChildBDay, "mm/dd/yyyy")
                End If
            End If
            
            ConnOmega.Execute "INSERT INTO tbl_Personnel_DataSheet_Children " & _
                              " (MasterKey, Line, FullName, BirthDate, Occupation, Address) " & _
                              " VALUES (" & iPK & ", " & l & ", '" & FORMATSQL(rt!ChildName) & "', " & _
                              " '" & sChildBDay & "', '" & FORMATSQL(rt!ChildOccupation) & "', " & _
                              " '" & FORMATSQL(rt!ChildAddress) & "')"
            UpdateProgress picProgress, i / iTotRecord
            rt.MoveNext
        Wend
        rt.Close
        
        ConnOmega.Execute "UPDATE tbl_Personnel_DataSheet SET Children = " & CDbl(l) & " WHERE (PK = " & iPK & ")"
        
        l = 0
        t = "SELECT tbl_Personnel_Education.* " & _
            " FROM tbl_Personnel_Education " & _
            " WHERE (ProfileKey = " & rs!PK & ") " & _
            " ORDER BY Line"
        If rt.State = adStateOpen Then rs.Close
        rt.Open t, ConnOmega
        While Not rt.EOF
            i = i + 1
            l = l + 1
            ConnOmega.Execute "INSERT INTO tbl_Personnel_DataSheet_Education " & _
                              " (MasterKey, Line, SchoolName, InclusiveDate, Course, Address) " & _
                              " VALUES (" & iPK & ", " & l & ", '" & FORMATSQL(rt!SchoolName) & "', " & _
                              " '" & FORMATSQL(rt!InclusiveDate) & "', '" & FORMATSQL(rt!Course) & "', " & _
                              " '" & FORMATSQL(rt!Address) & "')"
            UpdateProgress picProgress, i / iTotRecord
            rt.MoveNext
        Wend
        rt.Close
        
        ConnOmega.Execute "UPDATE tbl_Personnel_DataSheet SET Education = " & CDbl(l) & " WHERE (PK = " & iPK & ")"
        
        l = 0
        t = "SELECT tbl_Personnel_Employment.* " & _
            " FROM tbl_Personnel_Employment " & _
            " WHERE (ProfileKey = " & rs!PK & ") " & _
            " ORDER BY Line"
        If rt.State = adStateOpen Then rs.Close
        rt.Open t, ConnOmega
        While Not rt.EOF
            i = i + 1
            l = l + 1
            ConnOmega.Execute "INSERT INTO tbl_Personnel_DataSheet_Employment " & _
                              " (MasterKey, Line, Company, Positions, Salary, IncDates, Address) " & _
                              " VALUES (" & iPK & ", " & l & ", '" & FORMATSQL(rt!Company) & "', " & _
                              " '" & FORMATSQL(rt!Positions) & "', '" & FORMATSQL(rt!Salary) & "', " & _
                              " '" & FORMATSQL(rt!InclusiveDate) & "', '" & FORMATSQL(rt!Address) & "')"
            UpdateProgress picProgress, i / iTotRecord
            rt.MoveNext
        Wend
        rt.Close
        
        ConnOmega.Execute "UPDATE tbl_Personnel_DataSheet SET Employment = " & CDbl(l) & " WHERE (PK = " & iPK & ")"
        
        SAVE_IMAGES iPK, 0, SHOW_IMAGES(rs!PK, 0, "Employee Profile"), "Employee DataSheet"
        
    End If
    
End If
rs.Close

s = "SELECT tbl_Personnel_DataSheet.* " & _
    " FROM tbl_Personnel_DataSheet " & _
    " WHERE (LogInName = '" & gbl_UserName & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
rs.Requery
rs.Close

Unload Me

frmCrystalReportViewer.PRINT_PERSONNAL_DATASHEET gbl_UserName
If IsLoaded(frmCrystalReportViewer) Then frmCrystalReportViewer.ZOrder 0 Else frmCrystalReportViewer.Show

End Sub


Private Sub txtAsof_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOKAlphalist_Click
End Sub
