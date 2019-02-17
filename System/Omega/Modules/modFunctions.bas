Attribute VB_Name = "modFunctions"
Option Explicit

Public Sub LOAD_HIDE_MENU(bln As Boolean)
With MainForm
    .mnuUpdateMenu.Visible = bln
    .mnuUpdateGovtTables.Visible = bln
End With
End Sub

Public Function RoundOffIndex(dIndex As Double) As Double
Dim iIndexTmp, Arr, iRound, iRound1st, iRound2nd
iIndexTmp = Format(dIndex, "##0.00")
Arr = Split(iIndexTmp, ".", -1, 1)
iRound = Arr(1)
iRound1st = Mid(iRound, 1, 1)
iRound2nd = Mid(iRound, 2, 1)
If CDbl(iRound2nd) >= 0 And _
CDbl(iRound2nd) <= 5 Then
    RoundOffIndex = Arr(0) + CDbl("." & CStr(CDbl(iRound1st) + 0))
ElseIf CDbl(iRound2nd) >= 6 And _
CDbl(iRound2nd) <= 9 Then
    'MsgBox CDbl(iRound1st) + 1
    If CDbl(iRound1st) + 1 = 10 Then
        RoundOffIndex = Arr(0) + 1
    Else
        RoundOffIndex = Arr(0) + CDbl("." & CStr(CDbl(iRound1st) + 1))
    End If
End If
End Function

Public Function EncryptDecrypt(strWord As String) As String
Dim charval, i, j, str2Encrypt, ENCRYPTED
charval = 0
str2Encrypt = ""
ENCRYPTED = ""
For i = 1 To Len(strWord)
    str2Encrypt = Trim(Mid(strWord, i, 1))
  
    ' get char value
    For j = 0 To 255
        If str2Encrypt = Chr(j) Then
            charval = j
        End If
    Next j
  
    'encrypt now
    If charval < 128 Then
        ENCRYPTED = ENCRYPTED + Chr(charval + 128)
    Else
        ENCRYPTED = ENCRYPTED + Chr(charval - 128)
    End If
Next i
EncryptDecrypt = ENCRYPTED
End Function

Public Function EncryptDecryptLogIn(strWord As String) As String
Dim charval, i, j, str2Encrypt, ENCRYPTED
charval = 0
str2Encrypt = ""
ENCRYPTED = ""
For i = 1 To Len(strWord)
    str2Encrypt = Trim(Mid(strWord, i, 1))
  
    ' get char value
    For j = 0 To 255
        If str2Encrypt = Chr(j) Then
            charval = j
        End If
    Next j
  
    'encrypt now
    If charval < 128 Then
        ENCRYPTED = ENCRYPTED + Chr(charval) ' + 128)
    Else
        ENCRYPTED = ENCRYPTED + Chr(charval - 128)
    End If
Next i
EncryptDecryptLogIn = ENCRYPTED
End Function


'=======================================================
'   ENCRIPTING AND DECRIPTING TEXT
'=======================================================
Public Function Decript(strName) As String
Dim i, n, m, ostr
For i = 1 To Len(strName)
    n = Asc(Mid(strName, i, 1))
    m = IIf(i Mod 2 = 1, -1, 1)
    n = n - m * (i Mod 8)
    ostr = ostr & Chr(n)
Next
    Decript = ostr
End Function

Public Function Encript(strName) As String
Dim i, j, k, estr
For i = 1 To Len(strName)
    j = Asc(Mid(strName, i, 1))
    k = IIf(i Mod 2 = 1, -1, 1)
    j = j + k * (i Mod 8)
    estr = estr & Chr(j)
Next
    Encript = estr
End Function
'======================================================

Public Function LOAD_FORM(strModule, strModuleAction, frmForm As Form, frmModal As Long)
'LogInWithOutLoading = 0
gbl_MODULE = strModule
gbl_MODULE_Action = strModuleAction
Set gbl_FORM = frmForm
gbl_FORM_Modal = frmModal

If Trim(MainForm.Statusbar1.Panels(3).Text) = "" Then
    LogInWithOutLoading = 1
    aLogIn.Show 1
Else
    
    If AccessRights(gbl_MODULE, gbl_MODULE_Action) = False Then
        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
               "ACCESS DENIED!                                      ", vbCritical, "Alert"
        Exit Function
    End If
    
    If frmModal = 1 Then
        gbl_FORM.Show 1
    Else
        If IsLoaded(gbl_FORM) Then
            gbl_FORM.ZOrder 0
        Else
            gbl_FORM.Show
        End If
    End If
End If
End Function

'Public Function LOAD_FORM_X(strModule, strModuleAction, frmForm As Object, frmModal As Long)
''LogInWithOutLoading = 0
'gbl_MODULE = strModule
'gbl_MODULE_Action = strModuleAction
'Set gbl_FORMx = frmForm
'gbl_FORM_Modal = frmModal
'
'If Trim(MainForm.StatusBar1.Panels(3).Text) = "" Then
'    LogInWithOutLoading = 1
'    aLogIn.Show 1
'Else
'
'    If AccessRights(gbl_MODULE, gbl_MODULE_Action) = False Then
'        MsgBox "INSUFICIENT RIGHTS TO PERFORM THIS OPERATION.       " & vbCrLf & _
'               "ACCESS DENIED!                                      ", vbCritical, "Alert"
'        Exit Function
'    End If
'
'    If frmModal = 1 Then
'        gbl_FORMx.Show 1
'    Else
'        If IsLoadedx(gbl_FORMx) Then
'            gbl_FORMx.ZOrder 0
'        Else
'            gbl_FORMx.Show
'        End If
'    End If
'End If
'End Function

Public Function IsLoaded(ByVal frm As Form) As Boolean
    On Error Resume Next
    Dim f As Form
    For Each f In Forms
        On Error GoTo PG:
        If f.Name = frm.Name Then
            IsLoaded = True
            Exit Function
        End If
    Next
    IsLoaded = False
Exit Function
PG:
Exit Function
End Function


'Public Function IsLoadedx(ByVal frm As Object) As Boolean
'    On Error Resume Next
'    Dim f As Form
'    For Each f In Forms
'        On Error GoTo PG:
'        If f.Name = frm.Name Then
'            IsLoadedx = True
'            Exit Function
'        End If
'    Next
'    IsLoadedx = False
'Exit Function
'PG:
'Exit Function
'End Function



Public Function COMPUTE_AGE(dtmBDate As Date) As Long
Dim dtmBDate1 As Date
dtmBDate1 = DateSerial(Year(Now), Month(dtmBDate), Day(dtmBDate))
If DateValue(dtmBDate1) >= DateValue(Date) Then
    COMPUTE_AGE = DateDiff("yyyy", dtmBDate, Date) - 1
Else
    COMPUTE_AGE = DateDiff("yyyy", dtmBDate, Date)
End If
End Function

'   REMOVING Char(13) in Text
Public Function FORMATENTER(StrFieldVal As String) As String
FORMATENTER = Replace(Replace(Replace(StrFieldVal, Chr(13), Chr(8)), Chr(8), ""), Chr(10), "")
End Function

'   Formating Ñ character
Public Function FORMATINYI(StrFieldVal As String) As String
FORMATINYI = Replace(StrFieldVal, Chr(209), Chr(165))
End Function

'   Removing Single Quate
Public Function FORMATSQL(StrFieldVal As String) As String
FORMATSQL = Replace(StrFieldVal, "'", "''")
End Function

Public Function HTEXT(objText As Object)
objText.SelStart = 0
objText.SelLength = Len(objText.Text)
End Function

Public Function NUMBERKEYASCII(intKeyascii As Integer) As Integer
    If intKeyascii >= 1 And intKeyascii <= 7 Then
        NUMBERKEYASCII = 0
    ElseIf intKeyascii >= 9 And intKeyascii <= 12 Then
        NUMBERKEYASCII = 0
    ElseIf intKeyascii >= 14 And intKeyascii <= 44 Then
        NUMBERKEYASCII = 0
    ElseIf intKeyascii = 47 Then
        NUMBERKEYASCII = 0
    ElseIf intKeyascii >= 58 And intKeyascii <= 126 Then
        NUMBERKEYASCII = 0
    Else
        NUMBERKEYASCII = intKeyascii
    End If
End Function

Public Function NUMBERKEYASCII_NEG(intKeyascii As Integer) As Integer
    If intKeyascii >= 1 And intKeyascii <= 7 Then
        NUMBERKEYASCII_NEG = 0
    ElseIf intKeyascii >= 9 And intKeyascii <= 12 Then
        NUMBERKEYASCII_NEG = 0
    ElseIf intKeyascii >= 14 And intKeyascii <= 44 Then
        NUMBERKEYASCII_NEG = 0
    ElseIf intKeyascii = 47 Then
        NUMBERKEYASCII_NEG = 0
    ElseIf intKeyascii >= 58 And intKeyascii <= 126 Then
        NUMBERKEYASCII_NEG = 0
    Else
        NUMBERKEYASCII_NEG = intKeyascii
    End If
End Function

Public Function RETURNTEXTVALUE(objTxt As Object) As Double
Dim varValue
If objTxt.Text <> "" And IsNumeric(objTxt.Text) = True Then
   varValue = CDbl(objTxt.Text)
Else
   varValue = 0
End If
RETURNTEXTVALUE = varValue
End Function

Public Function RETURNLABELVALUE(objLbl As Object) As Double
Dim varValue
If objLbl.Caption <> "" And IsNumeric(objLbl.Caption) = True Then
   varValue = CDbl(objLbl.Caption)
Else
   varValue = 0
End If
RETURNLABELVALUE = varValue
End Function

Public Function UpdateProgress(pic As PictureBox, ByVal sngPercent As Single, Optional colBackground As Long = 16777215, Optional colForeground As Long = 8388608) As Boolean
On Local Error GoTo Err: 'if we encounter errors don't bother continuing
    Dim strPercent As String 'define all variables
    Dim intWidth As Integer
    Dim intHeight As Integer
    Dim intX As Integer
    Dim intY As Integer
    Dim intPercent As Integer
    pic.AutoRedraw = True 'make sure autoredraw is on
    pic.ForeColor = colForeground 'set colors
    pic.BackColor = colBackground
    intPercent = Int(100 * sngPercent) 'format into percentage
    strPercent = Str(intPercent) & "%" 'put a percentage sign after
    intWidth = pic.TextWidth(strPercent) 'use width and height of text to
    intHeight = pic.TextHeight(strPercent)
    intX = pic.Width / 2 - intWidth / 2 'calculate where to center the percentage
    intY = pic.Height / 2 - intHeight / 2
    pic.DrawMode = 13 'copy pen
    pic.Line (intX, intY)-Step(intWidth, intHeight), pic.BackColor, BF
    pic.CurrentX = intX 'position writing place
    pic.CurrentY = intY
    pic.Print strPercent 'print text
    pic.DrawMode = 10 'not xor
    If sngPercent > 0 Then 'if percentage is greater that 0 then
        pic.Line (0, 0)-(pic.Width * sngPercent, pic.Height), pic.ForeColor, BF 'paint a rectangle of the appropriate width
    Else 'otherwise
        pic.Line (0, 0)-(pic.Width, pic.Height), pic.BackColor, BF 'clear it
    End If
    pic.Refresh 'refresh it
UpdateProgress = True 'if no errors encountered return true
Exit Function
Err:
UpdateProgress = False 'we encountered errors - return false
End Function

Public Function UpdateProgress_No_Percent(pic As PictureBox, ByVal sngPercent As Single, Optional colBackground As Long = 16777215, Optional colForeground As Long = 8388608) As Boolean
On Local Error GoTo Err: 'if we encounter errors don't bother continuing
    Dim strPercent As String 'define all variables
    Dim intWidth As Integer
    Dim intHeight As Integer
    Dim intX As Integer
    Dim intY As Integer
    Dim intPercent As Integer
    pic.AutoRedraw = True 'make sure autoredraw is on
    pic.ForeColor = colForeground 'set colors
    pic.BackColor = colBackground
    intPercent = CDbl(100 * CDbl(sngPercent)) 'format into percentage
    strPercent = Str(intPercent) & "%" 'put a percentage sign after
    intWidth = pic.TextWidth(strPercent) 'use width and height of text to
    intHeight = pic.TextHeight(strPercent)
    intX = pic.Width / 2 - intWidth / 2 'calculate where to center the percentage
    intY = pic.Height / 2 - intHeight / 2
    pic.DrawMode = 13 'copy pen
    pic.Line (intX, intY)-Step(intWidth, intHeight), pic.BackColor, BF
    pic.CurrentX = intX 'position writing place
    pic.CurrentY = intY
'    pic.Print strPercent 'print text
    pic.DrawMode = 10 'not xor
    If sngPercent > 0 Then 'if percentage is greater that 0 then
        pic.Line (0, 0)-(pic.Width * sngPercent, pic.Height), pic.ForeColor, BF 'paint a rectangle of the appropriate width
    Else 'otherwise
        pic.Line (0, 0)-(pic.Width, pic.Height), pic.BackColor, BF 'clear it
    End If
    pic.Refresh 'refresh it
UpdateProgress_No_Percent = True 'if no errors encountered return true
Exit Function
Err:
UpdateProgress_No_Percent = False 'we encountered errors - return false
End Function

Public Function UpdateProgress_Caption(strCaption As String, pic As PictureBox, ByVal sngPercent As Single, Optional colBackground As Long = 16777215, Optional colForeground As Long = 8388608) As Boolean
On Local Error GoTo Err: 'if we encounter errors don't bother continuing
    Dim strPercent As String 'define all variables
    Dim intWidth As Integer
    Dim intHeight As Integer
    Dim intX As Integer
    Dim intY As Integer
    Dim intPercent As Integer
    pic.AutoRedraw = True 'make sure autoredraw is on
    pic.ForeColor = colForeground 'set colors
    pic.BackColor = colBackground
    intPercent = Int(100 * sngPercent) 'format into percentage
    strPercent = strCaption 'Str(intPercent) & "%" 'put a percentage sign after
    intWidth = pic.TextWidth(strPercent) 'use width and height of text to
    intHeight = pic.TextHeight(strPercent)
    intX = pic.Width / 2 - intWidth / 2 'calculate where to center the percentage
    intY = pic.Height / 2 - intHeight / 2
    pic.DrawMode = 13 'copy pen
    pic.Line (intX, intY)-Step(intWidth, intHeight), pic.BackColor, BF
    pic.CurrentX = intX 'position writing place
    pic.CurrentY = intY
    pic.Print strPercent 'print text
    pic.DrawMode = 10 'not xor
    If sngPercent > 0 Then 'if percentage is greater that 0 then
        pic.Line (0, 0)-(pic.Width * sngPercent, pic.Height), pic.ForeColor, BF 'paint a rectangle of the appropriate width
    Else 'otherwise
        pic.Line (0, 0)-(pic.Width, pic.Height), pic.BackColor, BF 'clear it
    End If
    pic.Refresh 'refresh it
UpdateProgress_Caption = True 'if no errors encountered return true
Exit Function
Err:
UpdateProgress_Caption = False 'we encountered errors - return false
End Function

Public Function ShowProgressInStatusBar(ByVal bShowProgressBar As Boolean)

    Dim tRC As RECT
    
'    If bShowProgressBar Then
'
' Get the size of the Panel (2) Rectangle from the status bar
' remember that Indexes in the API are always 0 based (well,
' nearly always) - therefore Panel(2) = Panel(1) to the api
'
'
        SendMessageAny MainForm.Statusbar1.hwnd, SB_GETRECT, 5, tRC
'
' and convert it to twips....
'
        With tRC
            .Top = (.Top * Screen.TwipsPerPixelY)
            .Left = (.Left * Screen.TwipsPerPixelX)
            .Bottom = (.Bottom * Screen.TwipsPerPixelY) - .Top
            .Right = (.Right * Screen.TwipsPerPixelX) - .Left
        End With
'
' Now Reparent the ProgressBar to the statusbar
'
        With MainForm.picProgressBar
            SetParent .hwnd, MainForm.Statusbar1.hwnd
            .Move tRC.Left + 10, tRC.Top + 10, tRC.Right - 40, tRC.Bottom - 40
            .Visible = True
'            .Value = 0
        End With
        
'    Else
'
' Reparent the progress bar back to the form and hide it
'
'        SetParent ProgressBar1.hwnd, Me.hwnd
'        ProgressBar1.Visible = False
'    End If
    
End Function

Public Function Check_Open_Forms() As String
Dim strNames, strForms As String
Dim Form As Form
strNames = ""
For Each Form In Forms
    strNames = strNames & ";" & Form.Name
Next Form
strForms = Replace(Replace(Replace(strNames, ";frmBackground", ""), ";MainFormPopupF", ""), ";MainForm", "")
Check_Open_Forms = Mid(strForms, 2, Len(strForms))
End Function


Public Function Get_Gross_Points(dblPar, dblScore) As Double
Dim dblPts
Select Case CDbl(dblPar)
    Case 3
        If CDbl(dblScore) = CDbl(dblPar) Then
            Get_Gross_Points = ParGrossPoints
        ElseIf CDbl(dblScore) < CDbl(dblPar) Then
            Select Case CDbl(dblScore)
                Case 2: Get_Gross_Points = ParGrossPoints + 1
                Case 1: Get_Gross_Points = ParGrossPoints + 2 '3
            End Select
        ElseIf CDbl(dblScore) > CDbl(dblPar) Then
            dblPts = (CDbl(dblScore) - CDbl(dblPar))
            Get_Gross_Points = ParGrossPoints - CDbl(dblPts)
        End If
    Case 4
        If CDbl(dblScore) = CDbl(dblPar) Then
            Get_Gross_Points = ParGrossPoints
        ElseIf CDbl(dblScore) < CDbl(dblPar) Then
            Select Case CDbl(dblScore)
                Case 3: Get_Gross_Points = ParGrossPoints + 1
                Case 2: Get_Gross_Points = ParGrossPoints + 2
                Case 1: Get_Gross_Points = ParGrossPoints + 3
            End Select
        ElseIf CDbl(dblScore) > CDbl(dblPar) Then
            dblPts = (CDbl(dblScore) - CDbl(dblPar))
            Get_Gross_Points = ParGrossPoints - CDbl(dblPts)
        End If
    Case 5
        If CDbl(dblScore) = CDbl(dblPar) Then
            Get_Gross_Points = ParGrossPoints '2
        ElseIf CDbl(dblScore) < CDbl(dblPar) Then
            Select Case CDbl(dblScore)
                Case 4: Get_Gross_Points = ParGrossPoints + 1
                Case 3: Get_Gross_Points = ParGrossPoints + 2
                Case 2: Get_Gross_Points = ParGrossPoints + 3 '4 '4
                Case 1: Get_Gross_Points = ParGrossPoints + 4 '3
            End Select
        ElseIf CDbl(dblScore) > CDbl(dblPar) Then
            dblPts = (CDbl(dblScore) - CDbl(dblPar))
            Get_Gross_Points = ParGrossPoints - CDbl(dblPts)
        End If
End Select
End Function

Public Function Get_Net_Points(dblHandicap, dblHandicapIndex, dblGrossPts) As Double
Dim dblNet, dblH, dblReturn, dblHPost, ArrH, dblHTmp
If CDbl(dblHandicap) <= 18 Then
    If CDbl(dblHandicap) >= CDbl(dblHandicapIndex) Then
        dblNet = CDbl(dblGrossPts) + 1
    Else
        dblNet = CDbl(dblGrossPts)
    End If
Else
    dblH = Format(CDbl(dblHandicap) / 18, "#0.00")
    ArrH = Split(CStr(dblH), ".", -1, 1)
    dblReturn = CDbl(ArrH(0))
    dblHTmp = 18 * CDbl(dblReturn)
    dblH = CDbl(dblHandicap) - CDbl(dblHTmp)
    If CDbl(dblH) <= 18 Then
        If CDbl(dblH) >= CDbl(dblHandicapIndex) Then
            dblNet = CDbl(dblGrossPts) + CDbl(dblReturn) + 1
        Else
            dblNet = CDbl(dblGrossPts) + CDbl(dblReturn)
        End If
    Else
        dblNet = CDbl(dblGrossPts) + CDbl(dblReturn)
    End If
End If
Get_Net_Points = IIf(CDbl(dblNet) <= 0, 0, CDbl(dblNet))
End Function

Public Function Get_Age(dtmBDay As Date, dtmDate As Date) As Double
If Month(FormatDateTime(dtmDate, vbShortDate)) >= Month(FormatDateTime(dtmBDay, vbShortDate)) Then
    If Day(FormatDateTime(dtmDate, vbShortDate)) >= Day(FormatDateTime(dtmBDay, vbShortDate)) Then
        Get_Age = DateDiff("yyyy", FormatDateTime(dtmBDay, vbShortDate), FormatDateTime(dtmDate, vbShortDate))
    Else
        Get_Age = DateDiff("yyyy", FormatDateTime(dtmBDay, vbShortDate), FormatDateTime(dtmDate, vbShortDate)) - 1
    End If
Else
    Get_Age = DateDiff("yyyy", FormatDateTime(dtmBDay, vbShortDate), FormatDateTime(dtmDate, vbShortDate)) - 1
End If
End Function

Public Function GET_HANDICAP_BEST_BALL() As Double

End Function

Public Function Center_Object(strmsg As String, dblWidth As Double) As Double
Dim LenMsg, LenMsgHalf As Double
Dim Array1
LenMsg = Len(strmsg)
Array1 = Split(CStr(CDbl(LenMsg) / 2), ".", -1, 1)
LenMsgHalf = CDbl(Array1(0))
Center_Object = CDbl(dblWidth) - CDbl(LenMsgHalf)
End Function

Public Function Stuff(sStr, cPos As Byte, cStuff As String, nNo As Byte) As String

    Dim sString As String, x As Byte
    sString = ""
    For x = 1 To nNo
        sString = sString & cStuff
    Next x
    If cPos = 1 Then
        sString = sString & sStr
    End If
    
    If cPos = 2 Then
        sString = sStr & sString
    End If
    
    Stuff = sString
    
End Function

Public Function AMT2WORDS(nInAmount As Double) As String

    Dim sInWords As String, sNum As String, nCent As Double, sThree As String
    Dim sNum1 As String, nCtr As Integer, sWord As String, lcont As Boolean
    
    Dim aTens(9) As String, aOnes(9) As String, aCValue(9) As String
    Dim nLen As Integer, x As Integer, nSingle As Integer
    
    
    aOnes(1) = "One"
    aOnes(2) = "Two"
    aOnes(3) = "Three"
    aOnes(4) = "Four"
    aOnes(5) = "Five"
    aOnes(6) = "Six"
    aOnes(7) = "Seven"
    aOnes(8) = "Eight"
    aOnes(9) = "Nine"

    aTens(1) = "Ten"
    aTens(2) = "Twenty"
    aTens(3) = "Thirty"
    aTens(4) = "Fourty"
    aTens(5) = "Fifty"
    aTens(6) = "Sixty"
    aTens(7) = "Seventy"
    aTens(8) = "Eigthy"
    aTens(9) = "Ninety"

    aCValue(1) = "Eleven"
    aCValue(2) = "Twelve"
    aCValue(3) = "Thirteen"
    aCValue(4) = "Fourteen"
    aCValue(5) = "Fifteen"
    aCValue(6) = "Sixteen"
    aCValue(7) = "Seventeen"
    aCValue(8) = "Eighteen"
    aCValue(9) = "Nineteen"
    
    nInAmount = Abs(nInAmount)
    sNum = Trim(Str(Int(nInAmount)))
    nCent = 0
    If Val(sNum) > 0 Then
        nCent = nInAmount - Val(sNum)
    Else
        nCent = nInAmount
    End If

    nCent = nCent * 100
    nLen = Len(sNum)
    If nLen < 12 Then
        sNum1 = Stuff(sNum, 1, "0", 12 - Len(sNum))
    Else
        sNum1 = sNum
    End If
    sInWords = ""
    
    nCtr = 1
    Do While True
        sThree = Mid(sNum1, nCtr, 3)
        sWord = ""
        For x = 1 To 3
            nSingle = Val(Mid(sThree, x, 1))
            lcont = True
            If nSingle > 0 Then
                If x = 1 Then
                    sWord = sWord + aOnes(nSingle) + " Hundred"
                End If
                If x = 2 Then
                    If nSingle = 1 And Val(Mid(sThree, 3, 1)) > 0 Then
                        sWord = sWord + " " + aCValue(Val(Mid(sThree, 3, 1)))
                        lcont = False
                    Else
                        If nSingle > 0 Then
                            sWord = sWord + " " + aTens(nSingle)
                        End If
                    End If
                End If
            
                If Not lcont Then
                    Exit For
                End If
                If x = 3 Then
                    sWord = sWord + " " + aOnes(nSingle)
                End If
            End If
        Next x
    
        sInWords = sInWords + " " + sWord
        If nCtr = 1 And Len(Trim(sInWords)) > 1 Then
            sInWords = sInWords + " Billion"
        End If
    
        If nCtr = 4 And Len(Trim(sInWords)) > 1 Then
            sInWords = sInWords + " Million"
        End If
    
        If nCtr = 7 And Len(Trim(sInWords)) > 1 Then
            sInWords = sInWords & " Thousand"
        End If
    
        nCtr = nCtr + 3
        If nCtr > 13 Then
            Exit Do
        End If
    
    Loop
    
    'I use Peso coz its our currency name in the Philippines
    'Just change it whatever currency word you have...
    
    If Val(sNum) > 1 Then
        sInWords = sInWords & "Pesos"
    End If
    
    If Val(sNum) = 1 Then
        sInWords = sInWords + "Peso"
    End If
    
    nCent = Format(nCent, "0.00")
    
    If nCent > 0 And Val(sNum) > 1 Then
        sInWords = sInWords + " & " + Trim(Str(nCent)) + "/100"
    End If

    If nCent > 0 And Val(sNum) = 0 Then
        sInWords = sInWords + " & " + Trim(Str(nCent)) + "/100"
    End If
    
    sInWords = sInWords + " Only"
    AMT2WORDS = Trim(sInWords)
    
End Function
