Attribute VB_Name = "modFunction"
Option Explicit


Public Function EXCEL_RANGE(iCol, iRow) As String
If CDbl(iCol) > 104 Then
    EXCEL_RANGE = Chr$(64 + 4) & Chr$(64 + (CDbl(iCol) - 104)) & CStr(iRow)
ElseIf CDbl(iCol) > 78 Then
    EXCEL_RANGE = Chr$(64 + 3) & Chr$(64 + (CDbl(iCol) - 78)) & CStr(iRow)
ElseIf CDbl(iCol) > 52 Then
    EXCEL_RANGE = Chr$(64 + 2) & Chr$(64 + (CDbl(iCol) - 52)) & CStr(iRow)
ElseIf CDbl(iCol) > 26 Then
    EXCEL_RANGE = Chr$(64 + 1) & Chr$(64 + (CDbl(iCol) - 26)) & CStr(iRow)
Else
    EXCEL_RANGE = Chr$(64 + iCol) & CStr(iRow)
End If
End Function


Public Function Get_Gross_Points(dblPar, dblScore) As Double
Dim dblPts
Select Case CDbl(dblPar)
    Case 3
        If CDbl(dblScore) = CDbl(dblPar) Then
            Get_Gross_Points = dParGrossPoints '2
        ElseIf CDbl(dblScore) < CDbl(dblPar) Then
            Select Case CDbl(dblScore)
                Case 2: Get_Gross_Points = dParGrossPoints + 1 '3
                Case 1: Get_Gross_Points = dParGrossPoints + 3 '5
            End Select
        ElseIf CDbl(dblScore) > CDbl(dblPar) Then
            dblPts = (CDbl(dblScore) - CDbl(dblPar))
            Get_Gross_Points = dParGrossPoints - CDbl(dblPts) '2 - CDbl(dblPts)
        End If
    Case 4
        If CDbl(dblScore) = CDbl(dblPar) Then
            Get_Gross_Points = dParGrossPoints '2
        ElseIf CDbl(dblScore) < CDbl(dblPar) Then
            Select Case CDbl(dblScore)
                Case 3: Get_Gross_Points = dParGrossPoints + 1 '3
                Case 2: Get_Gross_Points = dParGrossPoints + 2 '4
                Case 1: Get_Gross_Points = dParGrossPoints + 3 '5
            End Select
        ElseIf CDbl(dblScore) > CDbl(dblPar) Then
            dblPts = (CDbl(dblScore) - CDbl(dblPar))
            Get_Gross_Points = dParGrossPoints - CDbl(dblPts) '2 - CDbl(dblPts)
        End If
    Case 5
        If CDbl(dblScore) = CDbl(dblPar) Then
            Get_Gross_Points = dParGrossPoints '2
        ElseIf CDbl(dblScore) < CDbl(dblPar) Then
            Select Case CDbl(dblScore)
                Case 4: Get_Gross_Points = dParGrossPoints + 1
                Case 3: Get_Gross_Points = dParGrossPoints + 2
                Case 2: Get_Gross_Points = dParGrossPoints + 4
                Case 1: Get_Gross_Points = dParGrossPoints + 3
            End Select
        ElseIf CDbl(dblScore) > CDbl(dblPar) Then
            dblPts = (CDbl(dblScore) - CDbl(dblPar))
            Get_Gross_Points = dParGrossPoints - CDbl(dblPts) '2 - CDbl(dblPts)
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

Public Function HTEXT(objText As Object)
objText.SelStart = 0
objText.SelLength = Len(objText.Text)
End Function

Public Sub LOAD_CARD_LOCATION(iLocation, Grid As MSFlexGrid)
Dim GRow, HEADER1$, a, b, i, j, Tot, Tot1, Tot2
Dim s As String
Dim rs As New ADODB.Recordset
GRow = 0: a = -1: HEADER1$ = ""
s = "SELECT Description, H1, H2, H3, H4, H5, H6, H7, H8, H9, H10, H11, H12, H13, H14, H15, H16, H17, H18 " & _
    " From tbl_Scoring_Location_Details " & _
    " Where (MasterKey = " & iLocation & ") " & _
    " ORDER BY Line"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnRS
While Not rs.EOF
    a = a + 1
    With Grid
        If a = 0 Then
            .Clear
            i = 0: j = 0
            For i = 0 To rs.Fields.Count - 1
                If i = 9 Then
                    HEADER1$ = HEADER1$ & "|" & rs.Fields(i).Value
                    HEADER1$ = HEADER1$ & "|" & "OUT"
                Else
                    HEADER1$ = HEADER1$ & "|" & rs.Fields(i).Value
                End If
            Next i
            HEADER1$ = HEADER1$ & "|" & "IN" & "|" & "TOT"
            .FormatString = HEADER1$
            For i = 1 To .Cols - 1
                If i = 1 Then
                    .ColWidth(i) = 2000
                    .ColAlignment(i) = 1
                ElseIf i = 11 Or _
                i = 21 Or i = 22 Then
                    .ColWidth(i) = 550
                    .ColAlignment(i) = flexAlignRightCenter
                Else
                    .ColWidth(i) = 450
                    .ColAlignment(i) = flexAlignRightCenter
                End If
            Next i
        
        Else
            If CDbl(a) > 1 Then
                .Rows = .Rows + 1
            End If
            b = 0: Tot1 = 0: Tot2 = 0: Tot = 0
            For i = 0 To rs.Fields.Count - 1
                b = b + 1
                If i >= 0 And i <= 9 Then
                    If i = 0 Then
                        .TextMatrix(a, b) = rs.Fields(i).Value
                    ElseIf i = 9 Then
                        Tot1 = Tot1 + CDbl(rs.Fields(i).Value)
                        .TextMatrix(a, b) = rs.Fields(i).Value
                        b = b + 1
                        If a = 2 Then
                            .TextMatrix(a, b) = ""
                        Else
                            .TextMatrix(a, b) = Tot1
                        End If
                    Else
                        Tot1 = Tot1 + CDbl(rs.Fields(i).Value)
                        .TextMatrix(a, b) = rs.Fields(i).Value
                    End If
                Else
                    .TextMatrix(a, b) = rs.Fields(i).Value
                    Tot2 = Tot2 + CDbl(rs.Fields(i).Value)
                End If
            Next i
            Tot = Tot1 + Tot2
            b = b + 1
            If a = 2 Then
                .TextMatrix(a, b) = ""
            Else
                .TextMatrix(a, b) = Tot2
            End If
            b = b + 1
            If a = 2 Then
                .TextMatrix(a, b) = ""
            Else
                .TextMatrix(a, b) = Tot
            End If
        End If
    End With
    rs.MoveNext
Wend
rs.Close
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
