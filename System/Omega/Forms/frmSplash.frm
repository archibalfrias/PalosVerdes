VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
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
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   4035
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerOpen 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   3480
      Top             =   3360
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   6360
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "The maximum extend possible under law."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   2310
      Width           =   2655
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "may result in severe civil and penalties and will be prosecute"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   2130
      Width           =   3855
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "International treaties Unauthorized reproduction of program"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   1965
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Warning : This program is protected by copyright law and "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "RANCHO PALOS VERDES GOLF AND COUNTRY CLUB"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label lblLoading 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   2760
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   3135
      Left            =   0
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim x As Long

Dim sPath
Dim FadeValue As Byte


Private Sub Form_Load()
x = 0
Me.Height = 3135
Me.Width = 6015
TimerOpen.Enabled = True
'FadeValue = 254
End Sub

Private Sub TimerOpen_Timer()
TimerOpen.Enabled = False
x = x + 1
Select Case x

    Case 1
        DoEvents
        
        lblLoading.Caption = "Connecting to [" & UCase(gbl_Server) & "] ....."
        
        TimerOpen.Enabled = True
    Case 2
        DoEvents
        'MsgBox EncryptDecryptLogIn(CStr(sPasswordL))
        lblLoading.Caption = "Opening Database ....."
        DELETE_DNS_SQL_ODBC CStr(gbl_DatabaseL), CStr(gbl_ServerL), CStr(gbl_DatabaseL), CStr(sLogInL), EncryptDecryptLogIn(CStr(sPasswordL))
        DataOpen ConnOmega
        
        SaveSetting App.EXEName, "MainServerL", "MServerL", CStr(gbl_Server)
        SaveSetting App.EXEName, "MainDatabaseL", "MDatabaseL", CStr(gbl_Database)
        SaveSetting App.EXEName, "MainLogInL", "MLogInL", CStr(sLogIn)
        SaveSetting App.EXEName, "MainPasswordL", "MPasswordL", CStr(sPassword)
        
        SaveSetting App.EXEName, "ConnectionAttempt", "ConnectAttempt", "0"
        
        TimerOpen.Enabled = True
    Case 3
        DoEvents
        lblLoading.Caption = "Configuring Database ....."
        ConnOmega.Execute "set quoted_identifier off" & _
                          " set implicit_transactions off" & _
                          " set cursor_close_on_commit off" & _
                          " set ansi_warnings off " & _
                          " set ansi_padding off" & _
                          " set ansi_nulls off" & _
                          " set concat_null_yields_null off " & _
                          " set language us_english" & _
                          " set dateformat mdy" & _
                          " set datefirst 7"
        TimerOpen.Enabled = True
    Case 4
        DoEvents
        lblLoading.Caption = "Getting Server Info ....."
        ConnOmega.Execute "exec sp_server_info 18"
        TimerOpen.Enabled = True
    Case 5
        DoEvents
        lblLoading.Caption = "Use " & UCase(gbl_Database) & " ....."
        ConnOmega.Execute "use [" & gbl_Database & "]"
        'If Trim(gbl_ServerL) <> "" Then
        'DELETE_DNS_SQL_ODBC CStr(gbl_DatabaseL), CStr(gbl_ServerL), CStr(gbl_DatabaseL), "sa", ""
        'End If
        If Not checkWantedSQLDSN(gbl_Database) Then
            CREATE_SQL_DNS
        End If
        TimerOpen.Enabled = True
    Case 6
        DoEvents
        lblLoading.Caption = "Set TextSize ....."
        ConnOmega.Execute "SET TEXTSIZE 32768"
        TimerOpen.Enabled = True
    Case 7
        DoEvents
        lblLoading.Caption = "Select system user  ....."
        ConnOmega.Execute "select name from sysusers where uid = user_id()"
        TimerOpen.Enabled = True
    Case 8
        DoEvents
        lblLoading.Caption = "Set Lock Type  ....."
        ConnOmega.Execute "SET LOCK_TIMEOUT 200"
        TimerOpen.Enabled = True
    Case 9
        DoEvents
        TimerOpen.Enabled = True
    Case 10
        DoEvents
        lblLoading.Caption = "Getting Date From Server ....."
        s = "SELECT tbl_ApplicationDateTime.* " & _
            " FROM tbl_ApplicationDateTime " & _
            " WHERE (PK = 1)"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount = 0 Then
            ConnOmega.Execute "INSERT INTO tbl_ApplicationDateTime " & _
                              " (PK, ApplicationName) " & _
                              " VALUES (1, '" & App.EXEName & "')"
        Else
            ConnOmega.Execute "UPDATE tbl_ApplicationDateTime " & _
                              " SET ApplicationName = '" & App.EXEName & "' " & _
                              " WHERE (PK = 1)"
        End If
        rs.Close
        TimerOpen.Enabled = True
    Case 11
        DoEvents
        lblLoading.Caption = "Changing Client Date/Time ....."
        s = "SELECT tbl_ApplicationDateTime.*" & _
            " FROM tbl_ApplicationDateTime " & _
            " WHERE (PK = 1)"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            Date = Format(rs!ApplicationTime, "mm/dd/yyyy")
            Time = Format(rs!ApplicationTime, "hh:mm:ss AM/PM")
        End If
        rs.Close
        TimerOpen.Enabled = True
    Case 12
        DoEvents
        lblLoading.Caption = "Getting Company Information ....."
        s = "SELECT PK, CompanyName, Address1, " & _
            " Address2, TelNo, FaxNo, SSSNo, PHICNo, TIN " & _
            " From tbl_Company " & _
            " WHERE (PK = 1)"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            gbl_CompanyName = rs!CompanyName
            gbl_CompanyAddress1 = rs!Address1
            gbl_CompanyAddress2 = rs!Address2
            gbl_CompanyTelNo = rs!TelNo
            gbl_CompanyFaxNo = rs!FaxNo
            gbl_CompanySSSNo = rs!SSSNo
            gbl_CompanyPHICNo = rs!PHICNo
            gbl_CompanyTIN = rs!TIN
        End If
        rs.Close
        
        gbl_MinTakeHomePay = 0
'        s = "SELECT TOP (1) tbl_System_Settings.* " & _
'            " FROM tbl_System_Settings " & _
'            " WHERE (EffectDate <= '" & FormatDateTime(Date, vbShortDate) & "') " & _
'            " ORDER BY EffectDate DESC"
'        If rs.State = adStateOpen Then rs.Close
'        rs.Open s, ConnOmega
'        If rs.RecordCount > 0 Then
'            gbl_MinTakeHomePay = rs!MinTakeHomePay
'        End If
'        rs.Close
        
        TournamentKey = 0
        WithTeamPlay = 0
        WithIndividualPlay = 0
        TournamentName = ""
        TournamentRange = ""
        TeamPlayer2Cnt = 0
        AllowedTeam = 0
        NoofPlayerPerTeam = 0
        HandicapDivisor = 0
        DaysPlayerToPlay = 0
        ScoringType = 0
        TopHandicap = 0
        PointsToCnt = 0
        PointsToCntIndi = 0
        TopIndex = 0
        ParGrossPoints = 0
        LocationCnt = 0
        TeamAverage = 0
        TeamDivisorOrder = -1
        
        s = "SELECT tbl_Scoring_TournamentInfo.* " & _
            " FROM tbl_Scoring_TournamentInfo " & _
            " WHERE (Activated = 1)"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        If rs.RecordCount > 0 Then
            TournamentKey = rs!PK
            WithTeamPlay = rs!TeamPlay
            WithIndividualPlay = rs!IndividualPlay
            TournamentName = rs!TournamentName
            TournamentRange = Format(rs!TournamentStart, "mm/dd/yyyy") & " - " & Format(rs!TournamentEnd, "mm/dd/yyyy")
            TeamPlayer2Cnt = rs!PlayerToCount
            AllowedTeam = rs!AllowTeamPerPlayer
            NoofPlayerPerTeam = rs!NoofPlayerPerTeam
            HandicapDivisor = rs!HandicapDivisor
            DaysPlayerToPlay = rs!NoofPlays
            ScoringType = rs!Scoring
            PointsToCnt = rs!PointsToCountTeam
            PointsToCntIndi = rs!PointsToCountIndi
            TeamAverage = rs!TeamAverage
            ParGrossPoints = rs!ParGrossPoints
            TeamDivisorOrder = rs!TeamDivisorOrder
            
            t = "SELECT TOP 1 HTo " & _
                " From tbl_Scoring_TournamentInfo_Class " & _
                " Where (TournamentKey = " & TournamentKey & ") " & _
                " ORDER BY HTo DESC"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                TopHandicap = CDbl(rt!HTo)
            End If
            rt.Close
            
            t = "SELECT TOP 1 HTo " & _
                " From tbl_Scoring_TournamentInfo_Index " & _
                " Where (TournamentKey = " & TournamentKey & ") " & _
                " ORDER BY HTo DESC"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                TopIndex = CDbl(rt!HTo)
            End If
            rt.Close
            
            t = "SELECT COUNT(*) AS LocCnt " & _
                " From dbo.tbl_Scoring_TournamentInfo_Location " & _
                " WHERE (MasterKey = " & TournamentKey & ")"
            If rt.State = adStateOpen Then rt.Close
            rt.Open t, ConnOmega
            If rt.RecordCount > 0 Then
                LocationCnt = CDbl(IIf(IsNull(rt!LocCnt), 0, rt!LocCnt))
            End If
            rt.Close
            
            
            
        End If
        rs.Close
        
        gbl_VAT = 1.12

        TimerOpen.Enabled = True
        
    Case 13
        DoEvents
        lblLoading.Caption = "Configuring background ....."
        DoEvents
        s = "SELECT tbl_Wallpaper.* " & _
            " FROM tbl_Wallpaper"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, ConnOmega
        While Not rs.EOF
            DoEvents
            sPath = App.Path & "\Tmp\Back\" & rs!PK & ".jpg"
            If Dir(sPath) = "" Then
                Image1.Picture = LoadPicture(SHOW_IMAGES(rs!PK, 0, "Background"))
            End If
            rs.MoveNext
        Wend
        rs.Close
        TimerOpen.Enabled = True
    Case Else
        
        DoEvents
        lblLoading.Caption = ""
        Unload frmSplash
        MainForm.Show
        frmBackground.Quotes
        frmBackground.picQuotes.Visible = True
        frmBackground.picFreeMem.Visible = True
'        frmBackground.picDayTime.Visible = True
        MainForm.Timer_CheckIdle.Enabled = True
        
End Select
End Sub
