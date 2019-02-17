Attribute VB_Name = "modScoring"
Option Explicit

Dim GRow As Long
Dim HEADER1$, a, b, i, j, Tot1, Tot2, Tot

Dim s As String
Dim rs As New ADODB.Recordset
Dim t As String
Dim rt As New ADODB.Recordset
Dim u As String
Dim ru As New ADODB.Recordset
Dim v As String
Dim rv As New ADODB.Recordset

Public Sub LOAD_CARD(dEffectDate, Grid As MSFlexGrid)
v = "SELECT TOP 1 tbl_Scoring_Yardage_Par_HandicapIndex_Master.* " & _
    " FROM tbl_Scoring_Yardage_Par_HandicapIndex_Master " & _
    " WHERE (EffectDate <= '" & FormatDateTime(dEffectDate, vbShortDate) & "') " & _
    " ORDER BY EffectDate DESC"
If rv.State = adStateOpen Then rv.Close
rv.Open v, ConnOmega
If rv.RecordCount > 0 Then
    GRow = 0
    u = "SELECT Hole, Par, HandicapIndex, " & _
        " Gold, Blue, White, Red " & _
        " From tbl_Scoring_Yardage_Par_HandicapIndex " & _
        " WHERE (MasterKey = " & rv!PK & ") " & _
        " ORDER BY Hole"
    If ru.State = adStateOpen Then ru.Close
    ru.Open u, ConnOmega
    If ru.RecordCount > 0 Then
        With Grid
            .Clear
            i = 0: j = 0
            HEADER1$ = HEADER1$ & "|" & "HOLE"
            ru.MoveFirst
            While Not ru.EOF
                If i = 9 Then
                    i = 0
                    HEADER1$ = HEADER1$ & "|" & "OUT"
                    HEADER1$ = HEADER1$ & "|" & CStr(ru!Hole)
                Else
                    HEADER1$ = HEADER1$ & "|" & CStr(ru!Hole)
                End If
                i = i + 1
                ru.MoveNext
            Wend
            HEADER1$ = HEADER1$ & "|" & "IN" & "|" & "TOT" ' & "|" & "TOT"
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
        End With
    End If
    ru.Close
    
    'Par
    i = 1
    GRow = GRow + 1
    a = 0: Tot1 = 0: Tot2 = 0: Tot = 0
    u = "SELECT Hole, Par, HandicapIndex, " & _
        " Gold, Blue, White, Red " & _
        " From tbl_Scoring_Yardage_Par_HandicapIndex " & _
        " WHERE (MasterKey = " & rv!PK & ") " & _
        " ORDER BY Hole"
    If ru.State = adStateOpen Then ru.Close
    ru.Open u, ConnOmega
    If ru.RecordCount > 0 Then
        With Grid
            ru.MoveFirst
            .TextMatrix(GRow, i) = "PAR"
            While Not ru.EOF
                i = i + 1
                If a = 9 Then
                    .TextMatrix(GRow, i) = Tot
                    i = i + 1
                    .TextMatrix(GRow, i) = ru!Par
                    Tot1 = Tot
                    a = 0
                    Tot = 0
                Else
                    .TextMatrix(GRow, i) = ru!Par
                End If
                Tot = Tot + CDbl(ru!Par)
                a = a + 1
                ru.MoveNext
            Wend
            Tot2 = CDbl(Tot) + CDbl(Tot1)
            i = i + 1
            .TextMatrix(GRow, i) = Tot
            i = i + 1
            .TextMatrix(GRow, i) = Tot2 'Tot1
            'i = i + 1
            '.TextMatrix(GRow, i) = Tot2
        End With
    End If
    ru.Close
    
    'Handicap
    i = 1
    GRow = GRow + 1
    a = 0: Tot1 = 0: Tot2 = 0: Tot = 0
    u = "SELECT Hole, Par, HandicapIndex, " & _
        " Gold, Blue, White, Red " & _
        " From tbl_Scoring_Yardage_Par_HandicapIndex " & _
        " WHERE (MasterKey = " & rv!PK & ") " & _
        " ORDER BY Hole"
    If ru.State = adStateOpen Then ru.Close
    ru.Open u, ConnOmega
    If ru.RecordCount > 0 Then
        With Grid
            ru.MoveFirst
            .Rows = .Rows + 1
            .TextMatrix(GRow, i) = "HANDICAP"
            While Not ru.EOF
                i = i + 1
                If a = 9 Then
                    .TextMatrix(GRow, i) = ""
                    i = i + 1
                    .TextMatrix(GRow, i) = ru!HandicapIndex
                    Tot1 = Tot
                    a = 0
                    Tot = 0
                Else
                    .TextMatrix(GRow, i) = ru!HandicapIndex
                End If
                Tot = Tot + CDbl(ru!Par)
                a = a + 1
                ru.MoveNext
            Wend
            Tot2 = CDbl(Tot) + CDbl(Tot1)
            i = i + 1
            .TextMatrix(GRow, i) = ""
            i = i + 1
            .TextMatrix(GRow, i) = ""
            'i = i + 1
            '.TextMatrix(GRow, i) = ""
        End With
    End If
    ru.Close
    
    'Gold
    i = 1
    GRow = GRow + 1
    a = 0: Tot1 = 0: Tot2 = 0: Tot = 0
    u = "SELECT Hole, Par, HandicapIndex, " & _
        " Gold, Blue, White, Red " & _
        " From tbl_Scoring_Yardage_Par_HandicapIndex " & _
        " WHERE (MasterKey = " & rv!PK & ") " & _
        " ORDER BY Hole"
    If ru.State = adStateOpen Then ru.Close
    ru.Open u, ConnOmega
    If ru.RecordCount > 0 Then
        With Grid
            ru.MoveFirst
            .Rows = .Rows + 1
            .TextMatrix(GRow, i) = "GOLD"
            While Not ru.EOF
                i = i + 1
                If a = 9 Then
                    .TextMatrix(GRow, i) = Tot
                    i = i + 1
                    .TextMatrix(GRow, i) = ru!Gold
                    Tot1 = Tot
                    a = 0
                    Tot = 0
                Else
                    .TextMatrix(GRow, i) = ru!Gold
                End If
                Tot = Tot + CDbl(ru!Gold)
                a = a + 1
                ru.MoveNext
            Wend
            Tot2 = CDbl(Tot) + CDbl(Tot1)
            i = i + 1
            .TextMatrix(GRow, i) = Tot
            i = i + 1
            .TextMatrix(GRow, i) = Tot2 'Tot1
            'i = i + 1
            '.TextMatrix(GRow, i) = Tot2
        End With
    End If
    ru.Close
    
    'Blue
    i = 1
    GRow = GRow + 1
    a = 0: Tot1 = 0: Tot2 = 0: Tot = 0
    u = "SELECT Hole, Par, HandicapIndex, " & _
        " Gold, Blue, White, Red " & _
        " From tbl_Scoring_Yardage_Par_HandicapIndex " & _
        " WHERE (MasterKey = " & rv!PK & ") " & _
        " ORDER BY Hole"
    If ru.State = adStateOpen Then ru.Close
    ru.Open u, ConnOmega
    If ru.RecordCount > 0 Then
        With Grid
            ru.MoveFirst
            .Rows = .Rows + 1
            .TextMatrix(GRow, i) = "BLUE"
            While Not ru.EOF
                i = i + 1
                If a = 9 Then
                    .TextMatrix(GRow, i) = Tot
                    i = i + 1
                    .TextMatrix(GRow, i) = ru!Blue
                    Tot1 = Tot
                    a = 0
                    Tot = 0
                Else
                    .TextMatrix(GRow, i) = ru!Blue
                End If
                Tot = Tot + CDbl(ru!Blue)
                a = a + 1
                ru.MoveNext
            Wend
            Tot2 = CDbl(Tot) + CDbl(Tot1)
            i = i + 1
            .TextMatrix(GRow, i) = Tot
            i = i + 1
            .TextMatrix(GRow, i) = Tot2 'Tot1
            'i = i + 1
            '.TextMatrix(GRow, i) = Tot2
        End With
    End If
    ru.Close
    
    'White
    i = 1
    GRow = GRow + 1
    a = 0: Tot1 = 0: Tot2 = 0: Tot = 0
    u = "SELECT Hole, Par, HandicapIndex, " & _
        " Gold, Blue, White, Red " & _
        " From tbl_Scoring_Yardage_Par_HandicapIndex " & _
        " WHERE (MasterKey = " & rv!PK & ") " & _
        " ORDER BY Hole"
    If ru.State = adStateOpen Then ru.Close
    ru.Open u, ConnOmega
    If ru.RecordCount > 0 Then
        With Grid
            ru.MoveFirst
            .Rows = .Rows + 1
            .TextMatrix(GRow, i) = "WHITE"
            While Not ru.EOF
                i = i + 1
                If a = 9 Then
                    .TextMatrix(GRow, i) = Tot
                    i = i + 1
                    .TextMatrix(GRow, i) = ru!White
                    Tot1 = Tot
                    a = 0
                    Tot = 0
                Else
                    .TextMatrix(GRow, i) = ru!White
                End If
                Tot = Tot + CDbl(ru!White)
                a = a + 1
                ru.MoveNext
            Wend
            Tot2 = CDbl(Tot) + CDbl(Tot1)
            i = i + 1
            .TextMatrix(GRow, i) = Tot
            i = i + 1
            .TextMatrix(GRow, i) = Tot2 'Tot1
            'i = i + 1
            '.TextMatrix(GRow, i) = Tot2
        End With
    End If
    ru.Close
    
    'Red
    i = 1
    GRow = GRow + 1
    a = 0: Tot1 = 0: Tot2 = 0: Tot = 0
    u = "SELECT Hole, Par, HandicapIndex, " & _
        " Gold, Blue, White, Red " & _
        " From tbl_Scoring_Yardage_Par_HandicapIndex " & _
        " WHERE (MasterKey = " & rv!PK & ") " & _
        " ORDER BY Hole"
    If ru.State = adStateOpen Then ru.Close
    ru.Open u, ConnOmega
    If ru.RecordCount > 0 Then
        With Grid
            ru.MoveFirst
            .Rows = .Rows + 1
            .TextMatrix(GRow, i) = "RED"
            While Not ru.EOF
                i = i + 1
                If a = 9 Then
                    .TextMatrix(GRow, i) = Tot
                    i = i + 1
                    .TextMatrix(GRow, i) = ru!Red
                    Tot1 = Tot
                    a = 0
                    Tot = 0
                Else
                    .TextMatrix(GRow, i) = ru!Red
                End If
                Tot = Tot + CDbl(ru!Red)
                a = a + 1
                ru.MoveNext
            Wend
            Tot2 = CDbl(Tot) + CDbl(Tot1)
            i = i + 1
            .TextMatrix(GRow, i) = Tot
            i = i + 1
            .TextMatrix(GRow, i) = Tot2 'Tot1
            'i = i + 1
            '.TextMatrix(GRow, i) = Tot2
        End With
    End If
    ru.Close
End If
rv.Close
End Sub

Public Sub LOAD_CARD_LOCATION(iLocation, dEffectDate, Grid As MSFlexGrid)
GRow = 0: a = -1: HEADER1$ = ""
v = "SELECT TOP 1 PK " & _
    " From dbo.tbl_Scoring_Location_Master " & _
    " WHERE (MasterKey = " & iLocation & ") " & _
    " AND (EffectDate <= '" & FormatDateTime(dEffectDate, vbShortDate) & "') " & _
    " ORDER BY EffectDate DESC"
If rv.State = adStateOpen Then rv.Close
rv.Open v, ConnOmega
If rv.RecordCount > 0 Then
    u = "SELECT Description, H1, H2, H3, H4, H5, H6, H7, H8, H9, H10, H11, H12, H13, H14, H15, H16, H17, H18 " & _
        " From dbo.tbl_Scoring_Location_Details " & _
        " Where (MasterKey = " & rv!PK & ") " & _
        " ORDER BY Line"
    If ru.State = adStateOpen Then ru.Close
    ru.Open u, ConnOmega
    While Not ru.EOF
        a = a + 1
        With Grid
            If a = 0 Then
                .Clear
                i = 0: j = 0
                For i = 0 To ru.Fields.Count - 1
                    If i = 9 Then
                        HEADER1$ = HEADER1$ & "|" & ru.Fields(i).Value
                        HEADER1$ = HEADER1$ & "|" & "OUT"
                    Else
                        HEADER1$ = HEADER1$ & "|" & ru.Fields(i).Value
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
                For i = 0 To ru.Fields.Count - 1
                    b = b + 1
                    If i >= 0 And i <= 9 Then
                        If i = 0 Then
                            .TextMatrix(a, b) = ru.Fields(i).Value
                        ElseIf i = 9 Then
                            Tot1 = Tot1 + CDbl(ru.Fields(i).Value)
                            .TextMatrix(a, b) = ru.Fields(i).Value
                            b = b + 1
                            If a = 2 Then
                                .TextMatrix(a, b) = ""
                            Else
                                .TextMatrix(a, b) = Tot1
                            End If
                        Else
                            Tot1 = Tot1 + CDbl(ru.Fields(i).Value)
                            .TextMatrix(a, b) = ru.Fields(i).Value
                        End If
                    Else
                        .TextMatrix(a, b) = ru.Fields(i).Value
                        Tot2 = Tot2 + CDbl(ru.Fields(i).Value)
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
        ru.MoveNext
    Wend
    ru.Close
End If
rv.Close
End Sub
