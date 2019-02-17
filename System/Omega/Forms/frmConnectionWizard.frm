VERSION 5.00
Begin VB.Form frmConnectionWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Connection Wizard"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   ControlBox      =   0   'False
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
   ScaleHeight     =   4860
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic01 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   2280
      ScaleHeight     =   3975
      ScaleWidth      =   4695
      TabIndex        =   3
      Top             =   0
      Width           =   4695
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "To continue, click Next"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "* Create ODBC Connection"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "* Connect to Server"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "This wizard helps to you :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to New Connection Wizard"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   4335
      End
   End
   Begin RPVGCC.b8Line b8Line1 
      Height          =   60
      Left            =   0
      TabIndex        =   28
      Top             =   3960
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   106
   End
   Begin VB.PictureBox picFocus 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   240
      ScaleHeight     =   255
      ScaleWidth      =   375
      TabIndex        =   9
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   5520
      MouseIcon       =   "frmConnectionWizard.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Height          =   495
      Left            =   4200
      MouseIcon       =   "frmConnectionWizard.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmConnectionWizard.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
   Begin VB.PictureBox pic02 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   2280
      ScaleHeight     =   3975
      ScaleWidth      =   4695
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox txtLogIn 
         Height          =   315
         Left            =   1680
         TabIndex        =   25
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   24
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox txtDatabase 
         Height          =   315
         Left            =   1680
         TabIndex        =   17
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox txtServerName 
         Height          =   315
         Left            =   1680
         TabIndex        =   16
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   255
         Left            =   600
         TabIndex        =   27
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Log In"
         Height          =   255
         Left            =   600
         TabIndex        =   26
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to New Connection Wizard"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Fill in Server Name :"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Server"
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Database"
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "To continue, click Next"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   3480
         Width           =   1935
      End
   End
   Begin VB.PictureBox pic03 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   2280
      ScaleHeight     =   3975
      ScaleWidth      =   4695
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
      Begin VB.PictureBox picProgress1 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         ScaleHeight     =   195
         ScaleWidth      =   3795
         TabIndex        =   29
         Top             =   3160
         Width           =   3855
      End
      Begin VB.Timer TimerConnection 
         Enabled         =   0   'False
         Interval        =   400
         Left            =   3960
         Top             =   3360
      End
      Begin VB.PictureBox picProgress 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         ScaleHeight     =   195
         ScaleWidth      =   3795
         TabIndex        =   22
         Top             =   2880
         Width           =   3855
      End
      Begin VB.Label lblProgress 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   2640
         Width           =   3855
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait . . . . ."
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   3480
         Width           =   4095
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Checking Connection :"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to New Connection Wizard"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Image Image2 
      Height          =   3975
      Left            =   0
      Picture         =   "frmConnectionWizard.frx":091E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   3975
      Left            =   0
      Picture         =   "frmConnectionWizard.frx":E966
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "frmConnectionWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmp     As Long
Dim TotCnt  As Double
Dim x, i    As Double

Dim sPath

Private Sub cmdBack_Click()
picFocus.SetFocus
If pic02.Visible = True Then
    pic02.Visible = False
    pic01.ZOrder 0
    pic01.Visible = True
End If
End Sub

Private Sub cmdCancel_Click()
picFocus.SetFocus
End
End Sub

Private Sub cmdNext_Click()
picFocus.SetFocus
If pic01.Visible = True Then
    pic01.Visible = False
    pic02.ZOrder 0
    txtServerName.Text = "SERVER"
    txtDatabase.Text = "OMEGA_FINAL"
    txtLogIn.Text = "Arch|e"
    'txtPassword.Text = "Ãáóôéììïî±¹·¹"
    txtPassword.Text = "íáîáçåíåîô"
    pic02.Visible = True
    txtServerName.SetFocus
    cmdBack.Enabled = True
    Exit Sub
End If
If pic02.Visible = True Then
    If Trim(txtServerName.Text) = "" Then MsgBox "Please Supply Server Name!              ", vbCritical, "Error...": txtServerName.SetFocus: HTEXT txtServerName: Exit Sub
    pic02.Visible = False
    pic03.ZOrder 0
    pic03.Visible = True
    cmdBack.Enabled = False
    cmdNext.Enabled = False
    cmdCancel.Enabled = False
    
    If PingServer(Trim(txtServerName.Text)) = False Then
        MsgBox "Server Offline!                    ", vbCritical, "Error..."
        End
    End If
    
    SaveSetting App.EXEName, "MainServer", "MServer", Trim(txtServerName.Text)
    SaveSetting App.EXEName, "MainDatabase", "MDatabase", Trim(txtDatabase.Text)
    SaveSetting App.EXEName, "MainLogIn", "MLogIn", Trim(txtLogIn.Text)
    SaveSetting App.EXEName, "MainPassword", "MPassword", Trim(txtPassword.Text)
    
    gbl_Server = GetSetting(App.EXEName, "MainServer", "MServer", "")
    gbl_Database = GetSetting(App.EXEName, "MainDatabase", "MDatabase", "")
    sLogIn = GetSetting(App.EXEName, "MainLogIn", "MLogIn", "")
    sPassword = GetSetting(App.EXEName, "MainPassword", "MPassword", "")
    
    TimerConnection.Enabled = True
End If
End Sub

Private Sub Form_Load()
Dim sPw
TotCnt = 12
x = 0
pic01.ZOrder 0
pic01.Visible = True
pic02.Visible = False
pic03.Visible = False
cmdBack.Enabled = False
picProgress1.Visible = False
picProgress1.BackColor = &HFFFFFF
'txtDatabase.Text = "Omega"

'MsgBox EncryptDecrypt("Ãáóôéììïî±¹·¹")

tmp = SetWindowLong(txtServerName.hwnd, GWL_STYLE, GetWindowLong(txtServerName.hwnd, GWL_STYLE) Or ES_UPPERCASE)
tmp = SetWindowLong(txtDatabase.hwnd, GWL_STYLE, GetWindowLong(txtDatabase.hwnd, GWL_STYLE) Or ES_UPPERCASE)
End Sub

Private Sub TimerConnection_Timer()
TimerConnection.Enabled = False
x = x + 1
If x = 1 Then
    DoEvents
    lblProgress.Caption = "Opening Database . . . "
    DELETE_DNS_SQL_ODBC CStr(gbl_DatabaseL), CStr(gbl_ServerL), CStr(gbl_DatabaseL), CStr(sLogInL), EncryptDecryptLogIn(CStr(sPasswordL))
    DataOpen ConnOmega
    
    SaveSetting App.EXEName, "MainServerL", "MServerL", CStr(gbl_Server)
    SaveSetting App.EXEName, "MainDatabaseL", "MDatabaseL", CStr(gbl_Database)
    SaveSetting App.EXEName, "MainLogInL", "MLogInL", CStr(sLogIn)
    SaveSetting App.EXEName, "MainPasswordL", "MPasswordL", CStr(sPassword)
    
    SaveSetting App.EXEName, "ConnectionAttempt", "ConnectAttempt", "0"
    
    DoEvents
    TimerConnection.Interval = 700
    UpdateProgress_No_Percent picProgress, x / TotCnt
    TimerConnection.Enabled = True
ElseIf x = 2 Then
    DoEvents
    lblProgress.Caption = "Configuring Database ....."
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
    DoEvents
    TimerConnection.Interval = 400
    UpdateProgress_No_Percent picProgress, x / TotCnt
    TimerConnection.Enabled = True
ElseIf x = 3 Then
    DoEvents
    lblProgress.Caption = "Getting Server Info ....."
    ConnOmega.Execute "exec sp_server_info 18"
    UpdateProgress_No_Percent picProgress, x / TotCnt
    TimerConnection.Enabled = True
ElseIf x = 4 Then
    DoEvents
    lblProgress.Caption = "Use " & gbl_Database & " ....."
    ConnOmega.Execute "use [" & gbl_Database & "]"
    UpdateProgress_No_Percent picProgress, x / TotCnt
    TimerConnection.Enabled = True
ElseIf x = 5 Then
    DoEvents
    lblProgress.Caption = "Set TextSize ....."
    ConnOmega.Execute "SET TEXTSIZE 32768"
    UpdateProgress_No_Percent picProgress, x / TotCnt
    TimerConnection.Enabled = True
ElseIf x = 6 Then
    DoEvents
    lblProgress.Caption = "Select system user  ....."
    ConnOmega.Execute "select name from sysusers where uid = user_id()"
    UpdateProgress_No_Percent picProgress, x / TotCnt
    TimerConnection.Enabled = True
ElseIf x = 7 Then
    DoEvents
    lblProgress.Caption = "Set Lock Type  ....."
    ConnOmega.Execute "SET LOCK_TIMEOUT 200"
    UpdateProgress_No_Percent picProgress, x / TotCnt
    TimerConnection.Enabled = True
ElseIf x = 8 Then
    DoEvents
    lblProgress.Caption = "Getting Date From Server ....."
    'ConnOmega.Execute "UPDATE tbl_ApplicationDateTime SET ApplicationName = '" & App.EXEName & "'"
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
    UpdateProgress_No_Percent picProgress, x / TotCnt
    TimerConnection.Enabled = True
ElseIf x = 9 Then
    DoEvents
    lblProgress.Caption = "Changing Client Date/Time ....."
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
    UpdateProgress_No_Percent picProgress, x / TotCnt
    TimerConnection.Enabled = True
ElseIf x = 10 Then
    DoEvents
    lblProgress.Caption = "Creating ODBC Connection ....."
    If Trim(gbl_ServerL) <> "" Then
        DELETE_DNS_SQL_ODBC CStr(gbl_DatabaseL), CStr(gbl_ServerL), CStr(gbl_DatabaseL), sLogIn, EncryptDecryptLogIn(sPassword)
    End If
    If Not checkWantedSQLDSN(gbl_Database) Then
        CREATE_SQL_DNS
    End If
    UpdateProgress_No_Percent picProgress, x / TotCnt
    TimerConnection.Enabled = True
ElseIf x = 11 Then
    DoEvents
    lblProgress.Caption = "Configuring background ....."
    picProgress1.Visible = True
    i = 0
    DoEvents
    s = "SELECT tbl_Wallpaper.* " & _
        " FROM tbl_Wallpaper"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    While Not rs.EOF
        DoEvents
        i = i + 1
        sPath = App.Path & "\Tmp\Back\" & rs!PK & ".jpg"
        If Dir(sPath) = "" Then
            Image1.Picture = LoadPicture(SHOW_IMAGES(rs!PK, 0, "Background"))
        End If
        UpdateProgress_No_Percent picProgress1, i / rs.RecordCount
        rs.MoveNext
    Wend
    rs.Close
    TimerConnection.Enabled = True
ElseIf x = 12 Then
    DoEvents
    picProgress1.Visible = False
    lblProgress.Caption = "Done!"
    UpdateProgress_No_Percent picProgress, x / TotCnt
    TimerConnection.Interval = 300
    TimerConnection.Enabled = True
ElseIf x = 13 Then
    
    gbl_MinTakeHomePay = 0
'    s = "SELECT TOP (1) tbl_System_Settings.* " & _
'        " FROM tbl_System_Settings " & _
'        " WHERE (EffectDate <= '" & FormatDateTime(Date, vbShortDate) & "') " & _
'        " ORDER BY EffectDate DESC"
'    If rs.State = adStateOpen Then rs.Close
'    rs.Open s, ConnOmega
'    If rs.RecordCount > 0 Then
'        gbl_MinTakeHomePay = rs!MinTakeHomePay
'    End If
'    rs.Close
    
    s = "SELECT tbl_Company.* " & _
        " FROM tbl_Company"
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
        TeamAverage = 0
        TopIndex = 0
        ParGrossPoints = 0
        LocationCnt = 0
        TeamDivisorOrder = -1
        
        u = "SELECT tbl_Scoring_TournamentInfo.* " & _
            " FROM tbl_Scoring_TournamentInfo " & _
            " WHERE (Activated = 1)"
        If ru.State = adStateOpen Then ru.Close
        ru.Open u, ConnOmega
        If ru.RecordCount > 0 Then
            TournamentKey = ru!PK
            WithTeamPlay = ru!TeamPlay
            WithIndividualPlay = ru!IndividualPlay
            TournamentName = ru!TournamentName
            TournamentRange = Format(ru!TournamentStart, "mm/dd/yyyy") & " - " & Format(ru!TournamentEnd, "mm/dd/yyyy")
            TeamPlayer2Cnt = ru!PlayerToCount
            AllowedTeam = ru!AllowTeamPerPlayer
            NoofPlayerPerTeam = ru!NoofPlayerPerTeam
            HandicapDivisor = ru!HandicapDivisor
            DaysPlayerToPlay = ru!NoofPlays
            ScoringType = ru!Scoring
            PointsToCnt = ru!PointsToCountTeam
            PointsToCntIndi = ru!PointsToCountIndi
            TeamAverage = ru!TeamAverage
            ParGrossPoints = ru!ParGrossPoints
            TeamDivisorOrder = ru!TeamDivisorOrder
            
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
        ru.Close
        
        gbl_VAT = 1.12
        
        Unload Me
        MainForm.Show
        frmBackground.Quotes
        frmBackground.picQuotes.Visible = True
        frmBackground.picFreeMem.Visible = True
        'frmBackground.picDayTime.Visible = True
        MainForm.Timer_CheckIdle.Enabled = True
        
    Else
        
        gbl_CompanyName = rs!CompanyName
        gbl_CompanyAddress1 = rs!Address1
        gbl_CompanyAddress2 = rs!Address2
        gbl_CompanyTelNo = rs!TelNo
        gbl_CompanyFaxNo = rs!FaxNo
        gbl_CompanySSSNo = rs!SSSNo
        gbl_CompanyPHICNo = rs!PHICNo
        gbl_CompanyTIN = rs!TIN
        
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
        TopHandicap = 0
        PointsToCnt = 0
        PointsToCntIndi = 0
        TeamAverage = 0
        ParGrossPoints = 0
        LocationCnt = 0
        TeamDivisorOrder = -1
        
        Unload Me
        frmCompanyModal.Show 1
    End If
    If rs.State = adStateOpen Then rs.Close
    
End If
End Sub

Private Sub txtServerName_GotFocus()
HTEXT txtServerName
End Sub
