VERSION 5.00
Begin VB.Form frmBackground 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Background"
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9990
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
   MDIChild        =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimerGL 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8640
      Top             =   3120
   End
   Begin VB.CommandButton Command3 
      Caption         =   "migrate GL"
      Height          =   375
      Left            =   3960
      TabIndex        =   28
      Top             =   3720
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "migrate"
      Height          =   375
      Left            =   3960
      TabIndex        =   27
      Top             =   4920
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Timer TimerBDayList 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   4560
   End
   Begin VB.Timer TimerBDayMarquee 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1320
      Top             =   4560
   End
   Begin VB.PictureBox picBirthDays 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   120
      ScaleHeight     =   405
      ScaleWidth      =   5295
      TabIndex        =   24
      Top             =   5040
      Visible         =   0   'False
      Width           =   5295
      Begin VB.PictureBox picBirthDaysInside 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   30
         ScaleHeight     =   375
         ScaleWidth      =   5055
         TabIndex        =   25
         Top             =   30
         Width           =   5055
         Begin VB.Label lblListBirthDay 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label2"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   75
            Width           =   465
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "migrate"
      Height          =   375
      Left            =   3960
      TabIndex        =   23
      Top             =   2520
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3960
      ScaleHeight     =   555
      ScaleWidth      =   4395
      TabIndex        =   22
      Top             =   3000
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.TextBox txtImage 
      Height          =   315
      Left            =   120
      TabIndex        =   21
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtTimer 
      Height          =   315
      Left            =   120
      TabIndex        =   20
      Text            =   "0"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer TimerRandomizeCount 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   2880
   End
   Begin VB.Timer TimerRandomize 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2160
      Top             =   2880
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   120
      Pattern         =   "*.jpg"
      TabIndex        =   19
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox picFreeMem 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   800
      Left            =   5880
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   12
      Top             =   5640
      Visible         =   0   'False
      Width           =   2355
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   30
         ScaleHeight     =   735
         ScaleWidth      =   2295
         TabIndex        =   13
         Top             =   30
         Width           =   2295
         Begin VB.Timer TimerMemoryStatus 
            Interval        =   1000
            Left            =   1920
            Top             =   480
         End
         Begin VB.Label lblVirtual 
            BackStyle       =   0  'Transparent
            Caption         =   "0 MB"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   1200
            TabIndex        =   18
            Top             =   495
            Width           =   1095
         End
         Begin VB.Label lblPhysical 
            BackStyle       =   0  'Transparent
            Caption         =   "0 MB"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   1200
            TabIndex        =   17
            Top             =   300
            Width           =   1095
         End
         Begin VB.Line Line3 
            BorderColor     =   &H0000FF00&
            X1              =   1080
            X2              =   1080
            Y1              =   255
            Y2              =   985
         End
         Begin VB.Label lblVirCap 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Virtual"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   495
            Width           =   975
         End
         Begin VB.Label lblPhyCap 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Physical"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   0
            TabIndex        =   15
            Top             =   300
            Width           =   975
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0000FF00&
            X1              =   0
            X2              =   3735
            Y1              =   250
            Y2              =   250
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Available Free Memory"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   30
            Width           =   2055
         End
      End
   End
   Begin VB.PictureBox picQuotes 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   800
      Left            =   120
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   8
      Top             =   5880
      Visible         =   0   'False
      Width           =   3675
      Begin VB.PictureBox picQuotesInside 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   735
         Left            =   30
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   241
         TabIndex        =   9
         Top             =   30
         Width           =   3615
         Begin VB.Timer TimerQuotes 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   1920
            Top             =   240
         End
         Begin VB.TextBox txtQuotesCounter 
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Text            =   "0"
            Top             =   600
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtQuotes 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H0000FF00&
            Height          =   615
            Left            =   75
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   75
            Width           =   3480
         End
      End
   End
   Begin VB.PictureBox picDayTime 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   1400
      Left            =   6120
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3675
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   30
         ScaleHeight     =   1335
         ScaleWidth      =   3615
         TabIndex        =   1
         Top             =   30
         Width           =   3615
         Begin VB.Timer TimerSeparator 
            Enabled         =   0   'False
            Interval        =   500
            Left            =   480
            Top             =   0
         End
         Begin VB.Timer TimerDateTime 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   0
            Top             =   0
         End
         Begin VB.Label lblSeparator 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   495
            Left            =   1410
            TabIndex        =   7
            Top             =   820
            Width           =   255
         End
         Begin VB.Label lblAMPM 
            BackStyle       =   0  'Transparent
            Caption         =   "PM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   495
            Left            =   2050
            TabIndex        =   6
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lblMinute 
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   495
            Left            =   1600
            TabIndex        =   5
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   495
            Left            =   600
            TabIndex        =   4
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lblDate 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "September 24, 2009"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   495
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   3375
         End
         Begin VB.Label lblWeekDay 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Thursday"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   495
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   3375
         End
      End
   End
   Begin VB.Image imgSlide 
      Height          =   2655
      Left            =   0
      Picture         =   "frmBackground.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cnnS As String
Dim cnn As New ADODB.Connection
Dim cn As New ADODB.Connection
Dim RecordCount As Double

Dim myvalue As Integer

Dim sPath, Array1, sFileName, _
gbl_DatabaseO, gbl_ServerO, ProfileKey, i, strPath1, _
strTime, Array2, CurrDate, TommDate, strCelebrant

Public Function Quotes()
aa = "SELECT tbl_Greeting.* " & _
    " FROM tbl_Greeting"
If raa.State = adStateOpen Then raa.Close
raa.Open aa, ConnOmega
RecordCount = raa.RecordCount
If raa.State = adStateOpen Then raa.Close

aa:
Randomize
Array1 = Split(Format((RecordCount * Rnd + 1), "##0.00"), ".", -1, 1)
myvalue = Array1(0)
aa = "SELECT PK, GGreeting, GAuthor" & _
    " From tbl_Greeting" & _
    " WHERE  (PK = " & myvalue & ")"
If raa.State = adStateOpen Then raa.Close
raa.Open aa, ConnOmega
If raa.RecordCount > 0 Then
    txtQuotes.Text = "     " & raa!GGreeting & " (" & raa!GAuthor & ")"
Else
    GoTo AB:
End If
If raa.State = adStateOpen Then raa.Close
Exit Function

AB:
Randomize
Array1 = Split(Format((RecordCount * Rnd + 1), "##0.00"), ".", -1, 1)
myvalue = Array1(0)
aa = "SELECT PK, GGreeting, GAuthor" & _
    " From tbl_Greeting" & _
    " WHERE  (PK = " & myvalue & ")"
If raa.State = adStateOpen Then raa.Close
raa.Open aa, ConnOmega
If raa.RecordCount > 0 Then
    txtQuotes.Text = "     " & raa!GGreeting & " (" & raa!GAuthor & ")"
Else
    GoTo aa:
End If
If raa.State = adStateOpen Then raa.Close
End Function

Private Function GetFileIndex(iTotCnt) As Long
Randomize
Array1 = Split(Format((iTotCnt * Rnd + 1), "##0.00"), ".", -1, 1)
myvalue = Array1(0)
GetFileIndex = myvalue
End Function


Private Sub Command1_Click()

gbl_DatabaseO = "Omega": gbl_ServerO = "SERVER"
Set cnn = New ADODB.Connection
cnnS = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" + gbl_DatabaseO + ";Data Source=" + gbl_ServerO
cnn.CursorLocation = adUseClient
cnn.Mode = adModeReadWrite
cnn.IsolationLevel = adXactIsolated
cnn.Open cnnS
i = 0
s = "SELECT LName AS LastName, FName AS FirstName, MName AS MiddleName, MAX(BDate) AS BirthDate, MAX(BPlace) AS BirthPlace, MAX(ContactPerson) " & _
    " AS EmergencyName, MAX(Address) AS PresentAddress, MAX(Sex) AS Gender, MAX(Status) AS CivilStatus, MAX(DateMarried) AS DateMarriage, " & _
    " MAX(SpouseName) AS SpouseName, MAX(CelNo) AS ContactNumber, MAX(SSS) AS SSSNumber, MAX(PHIC) AS PHICNumber, MAX(PagIbig) " & _
    " AS HDMFNumber, MAX(TIN) AS TIN, MAX(License) AS DriverLicense, MAX(TaxStatus) AS TaxStatus, MAX(NoChildren) AS NoDependent " & _
    " From tbl_PersonnelProfile " & _
    " GROUP BY LName, FName, MName"
If rs.State = adStateOpen Then rs.Close
rs.Open s, cnn
While Not rs.EOF
    DoEvents
    i = i + 1
    '== Profile
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Information " & _
                      " (LastName, FirstName, MiddleName, BirthDate, BirthPlace, EmergencyName, PresentAddress, Gender, CivilStatus, DateMarriage, SpouseName, " & _
                      " ContactNumber, SSSNumber, PHICNumber, HDMFNumber, TIN, DriverLicense, TaxStatus, NoDependent) " & _
                      " VALUES ('" & rs!LastName & "', '" & rs!FirstName & "', '" & rs!MiddleName & "', '" & rs!BirthDate & "', '" & FORMATSQL(rs!BirthPlace) & "', " & _
                      " '" & rs!EmergencyName & "', '" & FORMATSQL(rs!PresentAddress) & "', " & rs!Gender & ", " & rs!CivilStatus & ", '" & rs!DateMarriage & "', " & _
                      " '" & rs!SpouseName & "', '" & rs!ContactNumber & "', '" & rs!SSSNumber & "', '" & rs!PHICNumber & "', '" & rs!HDMFNumber & "', '" & rs!TIN & "', " & _
                      " '" & rs!DriverLicense & "', " & rs!TaxStatus & ", " & rs!NoDependent & ")"
    t = "SELECT tbl_Personnel_Information.* " & _
        " FROM tbl_Personnel_Information " & _
        " WHERE (LastName = '" & rs!LastName & "')" & _
        " AND (FirstName = '" & rs!FirstName & "') " & _
        " AND (MiddleName = '" & rs!MiddleName & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        ProfileKey = rt!PK
    End If
    rt.Close
    '== IDNumber
    t = "SELECT PK, IDNumber, DHired AS DateHired, LastModified AS LastModified " & _
        " From tbl_PersonnelProfile " & _
        " WHERE (LName = '" & rs!LastName & "') " & _
        " AND (FName = '" & rs!FirstName & "') " & _
        " AND (MName = '" & rs!MiddleName & "')"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, cnn
    While Not rt.EOF
    
        ConnOmega.Execute "INSERT INTO tbl_Personnel_IDNumber " & _
                          " (PK, ProfileKey, IDNumber, DateHired, LastModified) " & _
                          " VALUES (" & rt!PK & ", " & ProfileKey & ", '" & rt!IDNumber & "', " & _
                          " '" & rt!DateHired & "', '" & IIf(IsNull(rt!LastModified), "", rt!LastModified) & "')"
        '== Action Memo
        u = "SELECT tbl_PersonnelAction.* " & _
            " FROM tbl_PersonnelAction " & _
            " WHERE (EmpPK = " & rt!PK & ")"
        If ru.State = adStateOpen Then ru.Close
        ru.Open u, cnn
        While Not ru.EOF
            
            v = "SELECT tbl_Personnel_Position.* " & _
                " FROM tbl_Personnel_Position " & _
                " WHERE (PK = " & ru!Positions & ")"
            If rv.State = adStateOpen Then rv.Close
            rv.Open v, ConnOmega
            If rv.RecordCount = 0 Then
                w = "SELECT tbl_PersonnelPosition.* " & _
                    " FROM tbl_PersonnelPosition " & _
                    " WHERE (PK = " & ru!Positions & ")"
                If rw.State = adStateOpen Then rw.Close
                rw.Open w, cnn
                If rw.RecordCount > 0 Then
                    ConnOmega.Execute "INSERT INTO tbl_Personnel_Position " & _
                                      " (PK, PositionCode, PositionName, PositionLevel, LastModified) " & _
                                      " VALUES (" & rw!PK & ", '" & rw!PositionCode & "', '" & rw!PositionName & "', " & _
                                      " " & rw!PositionLevel & ", '" & IIf(IsNull(rw!LastModified), "", rw!LastModified) & "') "
                End If
                rw.Close
            End If
            rv.Close
            
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Action " & _
                              " (PK, CntrlNo, EmpPK, Division, Dept, EmpStatus, TaxStatus, Positions, CompensationRate, Is_SSS, SSS, Is_PHIC, PHIC, Is_PAGIBIG, PAGIBIG, Is_TIN, " & _
                              " TIN, EffectivityDate, Basic, RatePerHour, Allowance, RatePerHourAllow, Remarks, LastModified, Locked) " & _
                              " VALUES (" & ru!PK & ", '" & ru!CntrlNo & "', " & ru!EmpPK & ", " & ru!Division & ", " & ru!Dept & ", " & ru!EmpStatus & ", " & ru!TaxStatus & ", " & ru!Positions & ", " & _
                              " " & ru!CompensationRate & ", " & ru!Is_SSS & ", '" & ru!SSS & "', " & ru!Is_PHIC & ", '" & ru!PHIC & "', " & ru!Is_PAGIBIG & ", '" & ru!PAGIBIG & "', " & ru!Is_TIN & ", " & _
                              " '" & ru!TIN & "', '" & ru!EffectivityDate & "', " & ru!Basic & ", " & ru!RatePerHour & ", " & ru!Allowance & ", " & ru!RatePerHourAllow & ", '" & ru!Remarks & "', " & _
                              " '" & ru!LastModified & "', " & ru!Locked & ")"
            ru.MoveNext
        Wend
        ru.Close
        
        '== Loans
        u = "SELECT tbl_PersonnelLoans.* " & _
            " FROM tbl_PersonnelLoans " & _
            " WHERE (EmpPK = " & rt!PK & ")"
        If ru.State = adStateOpen Then ru.Close
        ru.Open u, cnn
        While Not ru.EOF
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Loans " & _
                              " (PK, EmpPK, LoanType, DateGranted, LoanAmount, InterestType, Interest, " & _
                              " TotalAmount, Amortization, NoMonths, DateFrom, DateTo, ZeroOut, TotalPaid, " & _
                              " Balance, LastModified) " & _
                              " VALUES (" & ru!PK & ", " & ru!EmpPK & ", " & ru!LoanType & ", '" & ru!DateGranted & "', " & _
                              " " & ru!LoanAmount & ", " & ru!InterestType & ", " & ru!Interest & ", " & ru!TotalAmount & ", " & _
                              " " & ru!Amortization & ", " & ru!NoMonths & ", '" & ru!DateFrom & "', '" & ru!DateTo & "', " & _
                              " " & ru!ZeroOut & ", " & ru!TotalPaid & ", " & ru!Balance & ", '" & ru!LastModified & "')"
            ru.MoveNext
        Wend
        ru.Close
        
        '== Compensation
        u = "SELECT tbl_PersonnelPayroll.* " & _
            " FROM tbl_PersonnelPayroll " & _
            " WHERE (EmpPK = " & rt!PK & ")"
        If ru.State = adStateOpen Then ru.Close
        ru.Open u, cnn
        While Not ru.EOF
            ConnOmega.Execute "INSERT INTO tbl_Personnel_Compensation " & _
                              " (PK, EmpPK, Division, Dept, Status, Positions, Period, Basic, RatePerHour, ActionMemo, NoHours, SH_Hours, LH_Hours, SL_Hours, " & _
                              " Adjustment, Reg_OT_Hours, RD_OT_Hours, SH_OT_Hours, LH_OT_Hours, Amount_Earned, SH_Amount, LH_Amount, SL_Amount, Reg_OT_Amount, " & _
                              " RD_OT_Amount, SH_OT_Amount, LH_OT_Amount, TotalEarning, Mortuary, AR_Others, Advances, Shortages, Uniforms, Others, Is_Have_Loan, " & _
                              " SSSLoan_No, SSSLoan, SSSBalance, PagIbigLoan_No, PagIbigLoan, PagIbigBalance, Is_Have_Cont, SSS, SSS_Employer, SSS_EC, PHIC, " & _
                              " PHIC_Employer, PagIbig, PagIbig_Employer, WithHeld, TotalDeduction, NetEarning, Locked, LastModified) " & _
                              " VALUES (" & ru!PK & ", " & ru!EmpPK & ", " & ru!Division & ", " & ru!Dept & ", " & ru!Status & ", " & ru!Positions & ", " & ru!Period & ", " & _
                              " " & ru!Basic & ", " & ru!RatePerHour & ", " & ru!ActionMemo & ", " & ru!NoHours & ", " & ru!SH_Hours & ", " & ru!LH_Hours & ", " & ru!SL_Hours & ", " & _
                              " " & ru!Adjustment & ", " & ru!Reg_OT_Hours & ", " & ru!RD_OT_Hours & ", " & ru!SH_OT_Hours & ", " & ru!LH_OT_Hours & ", " & ru!Amount_Earned & ", " & _
                              " " & ru!SH_Amount & ", " & ru!LH_Amount & ", " & ru!SL_Amount & ", " & ru!Reg_OT_Amount & ", " & ru!RD_OT_Amount & ", " & ru!SH_OT_Amount & ", " & _
                              " " & ru!LH_OT_Amount & ", " & ru!TotalEarning & ", " & ru!Mortuary & ", " & ru!AR_Others & ", " & ru!Advances & ", " & ru!Shortages & ", " & ru!Uniforms & ", " & _
                              " " & ru!Others & ", " & ru!Is_Have_Loan & ", " & ru!SSSLoan_No & ", " & ru!SSSLoan & ", " & ru!SSSBalance & ", " & ru!PagIbigLoan_No & ", " & ru!PagIbigLoan & ", " & _
                              " " & ru!PagIbigBalance & ", " & ru!Is_Have_Cont & ", " & ru!SSS & ", " & ru!SSS_Employer & ", " & ru!SSS_EC & ", " & ru!PHIC & ", " & ru!PHIC_Employer & ", " & _
                              " " & ru!PAGIBIG & ", " & ru!PagIbig_Employer & ", " & ru!WithHeld & ", " & ru!TotalDeduction & ", " & ru!NetEarning & ", " & ru!Locked & ", '" & ru!LastModified & "')"
            ru.MoveNext
        Wend
        ru.Close
        
        rt.MoveNext
    Wend
    rt.Close
    
    UpdateProgress Picture1, i / rs.RecordCount
    
    rs.MoveNext
Wend
rs.Close
If cnn.State = adStateOpen Then cnn.Close
End Sub

Private Sub Command2_Click()

gbl_DatabaseO = "Omega_Final": gbl_ServerO = "Programmer"
Set cnn = New ADODB.Connection
cnnS = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" + gbl_DatabaseO + ";Data Source=" + gbl_ServerO
cnn.CursorLocation = adUseClient
cnn.Mode = adModeReadWrite
cnn.IsolationLevel = adXactIsolated
cnn.Open cnnS
i = 0
s = "SELECT tbl_Personnel_Information.* " & _
    " FROM tbl_Personnel_Information"
If rs.State = adStateOpen Then rs.Close
rs.Open s, cnn
While Not rs.EOF
    DoEvents
    i = i + 1
    ConnOmega.Execute "INSERT INTO tbl_Personnel_Information " & _
                      " (PK, LastName, FirstName, MiddleName, PresentAddress, OwnedHouse, Rented, Gender, LivingWParents, CivilStatus, BirthDate, BirthPlace, Religion, " & _
                      " Height, Weight, Nationality, ContactNumber, SSSNumber, PHICNumber, HDMFNumber, TIN, DriverLicense, SpouseName, SpouseOccupation, " & _
                      " SpouseAddress, FatherName, FatherOccupation, FatherAddress, MotherName, MotherOccupation, MotherAddress, Skills, OrganizationClubs, RefName, " & _
                      " RefContact, RefAddress, RefCompName, RefCompContact, RefCompAddress, EmergencyName, EmergencyRelation, EmergencyAddress, EmergencyContact, " & _
                      " LastModified, NoDependent, TaxStatus) " & _
                      " VALUES (" & rs!PK & ",'" & FORMATSQL(rs!LastName) & "', '" & FORMATSQL(rs!FirstName) & "', '" & FORMATSQL(rs!MiddleName) & "', '" & FORMATSQL(rs!PresentAddress) & "', " & rs!OwnedHouse & ", " & rs!Rented & ", " & _
                      " " & rs!Gender & ", " & rs!LivingWParents & ", " & rs!CivilStatus & ", '" & rs!BirthDate & "', '" & FORMATSQL(rs!BirthPlace) & "', '" & FORMATSQL(rs!Religion) & "', " & rs!Height & ", " & _
                      " " & rs!Weight & ", '" & FORMATSQL(rs!Nationality) & "', '" & FORMATSQL(rs!ContactNumber) & "', '" & rs!SSSNumber & "', '" & rs!PHICNumber & "', '" & rs!HDMFNumber & "', '" & rs!TIN & "', " & _
                      " '" & rs!DriverLicense & "', '" & FORMATSQL(rs!SpouseName) & "', '" & FORMATSQL(rs!SpouseOccupation) & "', '" & FORMATSQL(rs!SpouseAddress) & "', '" & FORMATSQL(rs!FatherName) & "', '" & FORMATSQL(rs!FatherOccupation) & "', " & _
                      " '" & FORMATSQL(rs!FatherAddress) & "', '" & FORMATSQL(rs!MotherName) & "', '" & FORMATSQL(rs!MotherOccupation) & "', '" & FORMATSQL(rs!MotherAddress) & "', '" & FORMATSQL(rs!Skills) & "', '" & FORMATSQL(rs!OrganizationClubs) & "', " & _
                      " '" & FORMATSQL(rs!RefName) & "', '" & FORMATSQL(rs!RefContact) & "', '" & FORMATSQL(rs!RefAddress) & "', '" & FORMATSQL(rs!RefCompName) & "', '" & FORMATSQL(rs!RefCompContact) & "', '" & FORMATSQL(rs!RefCompAddress) & "', " & _
                      " '" & FORMATSQL(rs!EmergencyName) & "', '" & FORMATSQL(rs!EmergencyRelation) & "', '" & FORMATSQL(rs!EmergencyAddress) & "', '" & FORMATSQL(rs!EmergencyContact) & "', '" & FORMATSQL(IIf(IsNull(rs!LastModified), "", rs!LastModified)) & "', " & rs!NoDependent & ", " & rs!TaxStatus & ")"
    
    If IsNull(rs!DateMarriage) = False Then
        ConnOmega.Execute "UPDATE tbl_Personnel_Information SET DateMarriage = '" & rs!DateMarriage & "' WHERE (PK = " & rs!PK & ")"
    End If
    If IsNull(rs!SpouseBDay) = False Then
        ConnOmega.Execute "UPDATE tbl_Personnel_Information SET SpouseBDay = '" & rs!SpouseBDay & "' WHERE (PK = " & rs!PK & ")"
    End If
    If IsNull(rs!FatherBDay) = False Then
        ConnOmega.Execute "UPDATE tbl_Personnel_Information SET FatherBDay = '" & rs!FatherBDay & "' WHERE (PK = " & rs!PK & ")"
    End If
    If IsNull(rs!MotherBDay) = False Then
        ConnOmega.Execute "UPDATE tbl_Personnel_Information SET MotherBDay = '" & rs!MotherBDay & "' WHERE (PK = " & rs!PK & ")"
    End If
    
    t = "SELECT tbl_Personnel_BrotherSister.* " & _
        " FROM tbl_Personnel_BrotherSister " & _
        " WHERE (ProfileKey = " & rs!PK & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, cnn
    While Not rt.EOF
        ConnOmega.Execute "INSERT INTO tbl_Personnel_BrotherSister " & _
                          " (ProfileKey, Line, BrotherSisterName, BrotherSisterOccupation, BrotherSisterAddress) " & _
                          " VALUES (" & rt!PK & ", " & rt!line & ", '" & FORMATSQL(rt!BrotherSisterName) & "', " & _
                          " '" & FORMATSQL(rt!BrotherSisterOccupation) & "', '" & FORMATSQL(rt!BrotherSisterAddress) & "')"
        rt.MoveNext
    Wend
    rt.Close
    
    t = "SELECT tbl_Personnel_Children.* " & _
        " FROM tbl_Personnel_Children " & _
        " WHERE (ProfileKey = " & rs!PK & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, cnn
    While Not rt.EOF
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Children " & _
                          " (ProfileKey, Line, ChildName, ChildOccupation, ChildAddress) " & _
                          " VALUES (" & rt!ProfileKey & ", " & rt!line & ", '" & FORMATSQL(rt!ChildName) & "', " & _
                          " '" & FORMATSQL(rt!ChildOccupation) & "', '" & FORMATSQL(rt!ChildAddress) & "')"
        rt.MoveNext
    Wend
    rt.Close
    
    t = "SELECT tbl_Personnel_Training.* " & _
        " FROM tbl_Personnel_Training " & _
        " WHERE (ProfileKey = " & rs!PK & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, cnn
    While Not rt.EOF
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Training " & _
                          " (ProfileKey, Line, Title, InclusiveDate, Venue) " & _
                          " VALUES (" & rt!ProfileKey & ", " & rt!line & ", '" & FORMATSQL(rt!Title) & "', " & _
                          " '" & FORMATSQL(rt!InclusiveDate) & "', '" & FORMATSQL(rt!Venue) & "')"
        rt.MoveNext
    Wend
    rt.Close
    
    t = "SELECT tbl_Personnel_Education.* " & _
        " FROM tbl_Personnel_Education " & _
        " WHERE (ProfileKey = " & rs!PK & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, cnn
    While Not rt.EOF
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Education " & _
                          " (ProfileKey, Line, SchoolName, InclusiveDate, Course, Address) " & _
                          " VALUES (" & rt!ProfileKey & ", " & rt!line & ", '" & FORMATSQL(rt!SchoolName) & "', " & _
                          " '" & FORMATSQL(rt!InclusiveDate) & "', " & _
                          " '" & FORMATSQL(rt!Course) & "', " & _
                          " '" & FORMATSQL(rt!Address) & "')"
        rt.MoveNext
    Wend
    rt.Close
    
    t = "SELECT tbl_Personnel_Employment.* " & _
        " FROM tbl_Personnel_Employment " & _
        " WHERE (ProfileKey = " & rs!PK & ")"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, cnn
    While Not rt.EOF
        ConnOmega.Execute "INSERT INTO tbl_Personnel_Employment " & _
                          " (ProfileKey, Line, Company, Positions, Salary, InclusiveDate, Address) " & _
                          " VALUES (" & rt!ProfileKey & ", " & rt!line & ", '" & FORMATSQL(rt!Company) & "', " & _
                          " '" & FORMATSQL(rt!Positions) & "', " & CDbl(rt!Salary) & ", " & _
                          " '" & FORMATSQL(rt!InclusiveDate) & "', '" & FORMATSQL(rt!Address) & "')"
        rt.MoveNext
    Wend
    rt.Close
    
    UpdateProgress Picture1, i / rs.RecordCount
    rs.MoveNext
Wend
rs.Close

i = 0
s = "SELECT tbl_Personnel_IDNumber.* " & _
    " FROM tbl_Personnel_IDNumber "
If rs.State = adStateOpen Then rs.Close
rs.Open s, cnn
While Not rs.EOF
    DoEvents
    i = i + 1
    ConnOmega.Execute "INSERT INTO tbl_Personnel_IDNumber " & _
                      " (PK, ProfileKey, IDNumber, DateHired, AccountNumber, LastModified) " & _
                      " VALUES (" & rs!PK & ", " & rs!ProfileKey & ", '" & rs!IDNumber & "', " & _
                      " '" & FormatDateTime(rs!DateHired, vbShortDate) & "', " & _
                      " '" & rs!AccountNumber & "', '" & rs!LastModified & "')"
    UpdateProgress Picture1, i / rs.RecordCount
    rs.MoveNext
Wend
rs.Close

If cnn.State = adStateOpen Then cnn.Close
End Sub

Private Sub Command3_Click()
With MainForm
    
    .CommonDialog1.DialogTitle = "OPEN FILE"
    .CommonDialog1.Filename = ""
    .CommonDialog1.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx"
    .CommonDialog1.FilterIndex = 1
    .CommonDialog1.ShowOpen
    strPath1 = .CommonDialog1.Filename
    If Trim(strPath1) = "" Then Exit Sub
    sPath = strPath1
    TimerGL.Enabled = True
End With
End Sub



Private Sub File1_Click()
txtImage.Text = App.Path & "\Wallpaper\" & File1.List(File1.ListIndex)
End Sub

Private Sub Form_Activate()
MainForm.txtActiveForm.Text = Me.Name
Me.ZOrder 1
End Sub

Private Sub Form_Load()
'Select Case Weekday(Now, vbMonday)
'    Case 1
'        lblWeekDay.ForeColor = &HFF00&
'        lblWeekDay.Caption = "Monday"
'    Case 2
'        lblWeekDay.ForeColor = &HFF00&
'        lblWeekDay.Caption = "Tuesday"
'    Case 3
'        lblWeekDay.ForeColor = &HFF00&
'        lblWeekDay.Caption = "Wednesday"
'    Case 4
'        lblWeekDay.ForeColor = &HFF00&
'        lblWeekDay.Caption = "Thursday"
'    Case 5
'        lblWeekDay.ForeColor = &HFF00&
'        lblWeekDay.Caption = "Friday"
'    Case 6
'        lblWeekDay.ForeColor = &HFF00&
'        lblWeekDay.Caption = "Saturday"
'    Case 7
'        lblWeekDay.ForeColor = &HFF&
'        lblWeekDay.Caption = "Sunday"
'End Select
'lblDate.Caption = Format(Now, "mmmm dd, yyyy")
'strTime = Format(Time, "hh:mm:ss AM/PM")
'Array1 = Split(strTime, ":", -1, 1)
'lblHour.Caption = Array1(0)
'lblMinute.Caption = Array1(1)
'Array2 = Split(Array1(2), " ", -1, 1)
'lblAMPM.Caption = Array2(1)

'For i = 1 To 3
'    Load Image1(i)
''    MsgBox Image1(i).Left
'Next i
'
'For i = 1 To 3
'    'Text1(i).Left = Text1(i - 1).Left + Text1(i - 1).Width
'    'Text1(i).Top = 120
'    Image1(i).ZOrder 0
'    Image1(i).Move Image1(i - 1).Left + Image1(i - 1).Width, 120
'    Image1(i).Visible = True
'Next i

DoEvents
Call GlobalMemoryStatus(MEM_STAT)
lblPhysical.Caption = Format((MEM_STAT.dwAvailPhys / 1024) / 1024, "#,##0.0") & " MB"
lblVirtual.Caption = Format((MEM_STAT.dwAvailVirtual / 1024) / 1024, "#,##0.0") & " MB"

'On Error Resume Next
'File1.FileName = App.Path & "\Wallpaper\"
'If File1.ListCount > 0 Then
'    TimerRandomize_Timer
'    TimerRandomizeCount.Enabled = True
'End If

TimerBDayList.Enabled = True

sPath = App.Path & "\Tmp\Back"
File1.Pattern = "*.JPG;*.JPEG"
File1.Path = sPath
TimerRandomizeCount.Enabled = True
If File1.ListCount = 0 Then TimerRandomizeCount.Enabled = False: Exit Sub

sFileName = File1.Path & "\" & File1.List(GetFileIndex(File1.ListCount))
On Error Resume Next
imgSlide.Picture = LoadPicture(sFileName)

's = "SELECT TOP 1 tbl_Wallpaper.* " & _
'    " FROM tbl_Wallpaper"
'If rs.State = adStateOpen Then rs.Close
'rs.Open s, ConnOmega
'If rs.RecordCount > 0 Then
'    TimerRandomize_Timer
'    TimerRandomizeCount.Enabled = True
'End If
'If rs.State = adStateOpen Then rs.Close



End Sub

Private Sub Form_Resize()
On Error Resume Next
With imgSlide
    .Top = 0
    .Left = 0
    .Height = Me.ScaleHeight
    .Width = Me.ScaleWidth
End With

With picDayTime
    .Top = 0
    .Left = Me.ScaleWidth - (.Width + 5)
End With

picBirthDays.Left = 0
picBirthDays.Top = ScaleHeight - (picFreeMem.Height + picBirthDays.Height)
picBirthDays.Width = ScaleWidth
picBirthDaysInside.Width = picBirthDays.Width - 60


With picFreeMem
    .Top = Me.ScaleHeight - (.Height + 0)
    .Left = Me.ScaleWidth - (.Width + 5)
End With

With picQuotes
    .Top = Me.ScaleHeight - (.Height + 0)
    .Left = 0
    .Width = Me.ScaleWidth - picFreeMem.Width
    
    
    Me.ScaleMode = 3
    
    picQuotesInside.Width = .Width - 2
    txtQuotes.Width = picQuotesInside.Width - 10
    
    Me.ScaleMode = 1
    
    '.Visible = True
    
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
'MsgBox Check_Open_Forms
If Trim(Check_Open_Forms) <> "" Then MsgBox "Please Close All Opened Forms!            ", vbInformation, "Close Forms": frmBackground.ZOrder 1: Cancel = -1: Exit Sub
End Sub

Private Sub TimerBDayList_Timer()
TimerBDayList.Enabled = False
lblListBirthDay.Caption = ""
CurrDate = FormatDateTime(Date, vbShortDate)
TommDate = FormatDateTime(DateSerial(Year(CurrDate), Month(CurrDate), Day(CurrDate) + 1), vbShortDate)
strCelebrant = ""
aa = "SELECT LastName + ',  ' + FirstName + '  ' + MiddleName AS BirthdayCelebrant " & _
    " From tbl_Personnel_Information " & _
    " WHERE (MONTH(BirthDate) = " & Month(CurrDate) & ") " & _
    " AND (DAY(BirthDate) = " & Day(CurrDate) & ") " & _
    " AND (ISNULL ((SELECT TOP 1 tbl_Personnel_EmploymentStatus.Active " & _
    " FROM tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
    " tbl_Personnel_Action ON tbl_Personnel_IDNumber.PK = tbl_Personnel_Action.EmpPK LEFT OUTER JOIN " & _
    " tbl_Personnel_EmploymentStatus ON tbl_Personnel_Action.EmpStatus = tbl_Personnel_EmploymentStatus.PK " & _
    " WHERE (tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK) " & _
    " AND (tbl_Personnel_Action.EffectivityDate <= '" & FormatDateTime(CurrDate, vbShortDate) & "') " & _
    " ORDER BY tbl_Personnel_Action.EffectivityDate DESC), 0) = 1) " & _
    " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
If raa.State = adStateOpen Then raa.Close
raa.Open aa, ConnOmega
If raa.RecordCount > 0 Then
    lblListBirthDay.Caption = lblListBirthDay.Caption & "Employee/s Celebrating birthday today " & Format(CurrDate, "mmmm dd, yyyy") & ":"
    strCelebrant = ""
    While Not raa.EOF
        strCelebrant = strCelebrant & _
                       "; " & raa!BirthdayCelebrant
        raa.MoveNext
    Wend
    lblListBirthDay.Caption = lblListBirthDay.Caption & _
                              " " & Mid(strCelebrant, 3, Len(strCelebrant))
                              
    strCelebrant = ""
    t = "SELECT LastName + ',  ' + FirstName + '  ' + MiddleName AS BirthdayCelebrant " & _
        " From tbl_Personnel_Information " & _
        " WHERE (MONTH(BirthDate) = " & Month(TommDate) & ") " & _
        " AND (DAY(BirthDate) = " & Day(TommDate) & ") " & _
        " AND (ISNULL ((SELECT TOP 1 tbl_Personnel_EmploymentStatus.Active " & _
        " FROM tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
        " tbl_Personnel_Action ON tbl_Personnel_IDNumber.PK = tbl_Personnel_Action.EmpPK LEFT OUTER JOIN " & _
        " tbl_Personnel_EmploymentStatus ON tbl_Personnel_Action.EmpStatus = tbl_Personnel_EmploymentStatus.PK " & _
        " WHERE (tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK) " & _
        " AND (tbl_Personnel_Action.EffectivityDate <= '" & FormatDateTime(TommDate, vbShortDate) & "') " & _
        " ORDER BY tbl_Personnel_Action.EffectivityDate DESC), 0) = 1) " & _
        " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lblListBirthDay.Caption = lblListBirthDay.Caption & _
                                  Space(15) & "Employee/s Celebrating birthday tomorrow " & Format(TommDate, "mmmm dd, yyyy") & ":"
        While Not rt.EOF
            strCelebrant = strCelebrant & _
                           "; " & rt!BirthdayCelebrant
            rt.MoveNext
        Wend
        
        lblListBirthDay.Caption = lblListBirthDay.Caption & _
                                  " " & Mid(strCelebrant, 3, Len(strCelebrant))
    End If
    rt.Close
    
    picBirthDays.Visible = True
    TimerBDayMarquee.Enabled = True
    
Else
    
    strCelebrant = ""
    t = "SELECT LastName + ',  ' + FirstName + '  ' + MiddleName AS BirthdayCelebrant " & _
        " From tbl_Personnel_Information " & _
        " WHERE (MONTH(BirthDate) = " & Month(TommDate) & ") " & _
        " AND (DAY(BirthDate) = " & Day(TommDate) & ") " & _
        " AND (ISNULL ((SELECT TOP 1 tbl_Personnel_EmploymentStatus.Active " & _
        " FROM tbl_Personnel_IDNumber LEFT OUTER JOIN " & _
        " tbl_Personnel_Action ON tbl_Personnel_IDNumber.PK = tbl_Personnel_Action.EmpPK LEFT OUTER JOIN " & _
        " tbl_Personnel_EmploymentStatus ON tbl_Personnel_Action.EmpStatus = tbl_Personnel_EmploymentStatus.PK " & _
        " WHERE (tbl_Personnel_IDNumber.ProfileKey = tbl_Personnel_Information.PK) " & _
        " AND (tbl_Personnel_Action.EffectivityDate <= '" & FormatDateTime(TommDate, vbShortDate) & "') " & _
        " ORDER BY tbl_Personnel_Action.EffectivityDate DESC), 0) = 1) " & _
        " ORDER BY LastName + ',  ' + FirstName + '  ' + MiddleName"
    If rt.State = adStateOpen Then rt.Close
    rt.Open t, ConnOmega
    If rt.RecordCount > 0 Then
        lblListBirthDay.Caption = lblListBirthDay.Caption & _
                                  Space(10) & "Celebrating birthday tomorrow " & Format(TommDate, "mmmm dd, yyyy") & ":"
        While Not rt.EOF
            strCelebrant = strCelebrant & _
                           "; " & rt!BirthdayCelebrant
            rt.MoveNext
        Wend
        
        lblListBirthDay.Caption = lblListBirthDay.Caption & _
                                  " " & Mid(strCelebrant, 3, Len(strCelebrant))
        
        picBirthDays.Visible = True
        TimerBDayMarquee.Enabled = True
        
    End If
    rt.Close
    
End If
raa.Close
End Sub

Private Sub TimerBDayMarquee_Timer()
TimerBDayMarquee.Enabled = False
Static x As Double
If x = 0 Then
    lblListBirthDay.Left = picBirthDays.Width
End If
x = x + 10
lblListBirthDay.Left = picBirthDays.Width - x
If x >= picBirthDays.Width + lblListBirthDay.Width Then
    x = 0
End If
TimerBDayMarquee.Enabled = True
End Sub

Private Sub TimerDateTime_Timer()
Select Case Weekday(Now, vbMonday)
    Case 1
        lblWeekDay.ForeColor = &HFF00&
        lblWeekDay.Caption = "Monday"
    Case 2
        lblWeekDay.ForeColor = &HFF00&
        lblWeekDay.Caption = "Tuesday"
    Case 3
        lblWeekDay.ForeColor = &HFF00&
        lblWeekDay.Caption = "Wednesday"
    Case 4
        lblWeekDay.ForeColor = &HFF00&
        lblWeekDay.Caption = "Thursday"
    Case 5
        lblWeekDay.ForeColor = &HFF00&
        lblWeekDay.Caption = "Friday"
    Case 6
        lblWeekDay.ForeColor = &HFF00&
        lblWeekDay.Caption = "Saturday"
    Case 7
        lblWeekDay.ForeColor = &HFF&
        lblWeekDay.Caption = "Sunday"
End Select
lblDate.Caption = Format(Now, "mmmm dd, yyyy")
strTime = Format(Time, "hh:mm:ss AM/PM")
Array1 = Split(strTime, ":", -1, 1)
lblHour.Caption = Array1(0)
lblMinute.Caption = Array1(1)
Array2 = Split(Array1(2), " ", -1, 1)
lblAMPM.Caption = Array2(1)
End Sub

Private Sub TimerGL_Timer()
TimerGL.Enabled = False
Screen.MousePointer = vbHourglass
Set cn = New ADODB.Connection
cn.Provider = "Microsoft.Jet.OLEDB.4.0"
cn.ConnectionString = _
    "Data Source= " & sPath & ";" & _
    "Extended Properties=Excel 8.0;"
cn.CursorLocation = adUseClient
If cn.State = adStateOpen Then cn.Close
cn.Open

i = 0
s = "SELECT * FROM [GLAccount$]"
If rs.State = adStateOpen Then rs.Close
rs.Open s, cn, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    DoEvents
    i = i + 1
    If IsNull(rs![Code]) = False Then
        t = "SELECT tbl_GL_Accounts.* " & _
            " FROM tbl_GL_Accounts " & _
            " WHERE (AccountCode = '" & rs![Code] & "')"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount = 0 Then
            ConnOmega.Execute "INSERT INTO tbl_GL_Accounts " & _
                              " (AccountCode, AccountName, Dept) " & _
                              " VALUES ('" & rs![Code] & "', " & _
                              " '" & UCase(FORMATSQL(rs![Name])) & "', " & _
                              " " & rs![Dept] & ")"
'        Else
'            ConnOmega.Execute "UPDATE tbl_GL_Accounts " & _
'                              " SET CurrentAmount = 0 " & _
'                              " WHERE (AccountCode = '" & rs![Code] & "')"
        End If
        rt.Close
    End If
    UpdateProgress Picture1, i / rs.RecordCount
    rs.MoveNext
Wend
rs.Close
If cn.State = adStateOpen Then cn.Close
Screen.MousePointer = vbDefault
End Sub

Private Sub TimerMemoryStatus_Timer()
DoEvents
Call GlobalMemoryStatus(MEM_STAT)
lblPhysical.Caption = Format((MEM_STAT.dwAvailPhys / 1024) / 1024, "#,##0.0") & " MB"
lblVirtual.Caption = Format((MEM_STAT.dwAvailVirtual / 1024) / 1024, "#,##0.0") & " MB"
End Sub

Private Sub TimerQuotes_Timer()
txtQuotesCounter.Text = RETURNTEXTVALUE(txtQuotesCounter) + 1
End Sub

Private Sub TimerRandomize_Timer()
TimerRandomize.Enabled = False
TimerRandomizeCount.Enabled = False
'Dim myvalue As Integer
'Dim strImage
'Randomize
'myvalue = Int(File1.ListCount * Rnd)
'File1.ListIndex = myvalue
'On Error Resume Next
'imgSlide.Picture = LoadPicture(txtImage.Text)
'txtTimer.Text = "0"


txtTimer.Text = "0"

RANDOM:
s = "SELECT tbl_Wallpaper.* " & _
    " FROM tbl_Wallpaper"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
Randomize
myvalue = Int(rs.RecordCount * Rnd)
If rs.State = adStateOpen Then rs.Close

s = "SELECT tbl_Wallpaper.* " & _
    " FROM tbl_Wallpaper " & _
    " WHERE (PK = " & myvalue & ")"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount = 0 Then
    GoTo RANDOM:
Else
    On Error GoTo RANDOM:
    imgSlide.Picture = LoadPicture(SHOW_IMAGES(rs!PK, 0, "Wallpaper"))
End If
If rs.State = adStateOpen Then rs.Close

TimerRandomizeCount.Enabled = True
End Sub

Private Sub TimerRandomizeCount_Timer()
If gbl_Slides_Background = 1 Then
    txtTimer.Text = RETURNTEXTVALUE(txtTimer) + 1
End If
End Sub

Private Sub TimerSeparator_Timer()
If lblSeparator.Visible = True Then
    lblSeparator.Visible = False
Else
    lblSeparator.Visible = True
End If
End Sub

Private Sub txtQuotesCounter_Change()
If RETURNTEXTVALUE(txtQuotesCounter) >= CDbl(gbl_Quotes_Time) Then
    TimerQuotes.Enabled = False
    txtQuotesCounter.Text = 0
    Quotes
    TimerQuotes.Enabled = True
End If
End Sub

Private Sub txtTimer_Change()
'If RETURNTEXTVALUE(txtTimer) >= CDbl(gbl_Slides_Time) Then
'    TimerRandomizeCount.Enabled = False
'    TimerRandomize.Enabled = True
'End If
If RETURNTEXTVALUE(txtTimer) >= CDbl(gbl_Slides_Time) Then
    TimerRandomizeCount.Enabled = False
    txtTimer.Text = 0
    TimerRandomizeCount.Enabled = True
    'sPath = App.Path & "\Tmp\Background"
    'sFileName = sPath & "\" & GetRecordIndex & ".jpg"
    'On Error GoTo PG:
    'imgSlide.Picture = LoadPicture(sFileName)
    If File1.ListCount = 0 Then Exit Sub
    On Error Resume Next
    imgSlide.Picture = LoadPicture(File1.Path & "\" & File1.List(GetFileIndex(File1.ListCount)))
End If
End Sub
