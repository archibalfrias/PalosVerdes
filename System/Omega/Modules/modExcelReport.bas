Attribute VB_Name = "modExcelReport"
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


