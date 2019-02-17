VERSION 5.00
Begin VB.Form frmPrinter 
   BackColor       =   &H00C6B8A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrinter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00C6B8A4&
      Caption         =   " Page Range "
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   2775
      Begin VB.PictureBox picPageRange 
         BackColor       =   &H00C6B8A4&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   2535
         TabIndex        =   12
         Top             =   360
         Width           =   2535
         Begin VB.OptionButton optAll 
            BackColor       =   &H00C6B8A4&
            Caption         =   "All"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optPages 
            BackColor       =   &H00C6B8A4&
            Caption         =   "Pages"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtPgFrom 
            Height          =   315
            Left            =   960
            TabIndex        =   14
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtPgTo 
            Height          =   315
            Left            =   1920
            TabIndex        =   13
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1680
            TabIndex        =   17
            Top             =   360
            Width           =   255
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C6B8A4&
      Caption         =   "PRINTER"
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5175
      Begin VB.ComboBox cmbPrinter 
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   360
         Width           =   4815
      End
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   510
      Left            =   2760
      Picture         =   "frmPrinter.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   1560
   End
   Begin VB.CommandButton cmdOK 
      Height          =   510
      Left            =   1080
      Picture         =   "frmPrinter.frx":2EFE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2400
      Width           =   1560
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C6B8A4&
      Caption         =   "Copies"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3000
      TabIndex        =   0
      Top             =   1080
      Width           =   2295
      Begin VB.TextBox txtCopies 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Copies"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C6B8A4&
      Caption         =   "Character per Inch"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Top             =   1680
      Width           =   2295
      Begin VB.PictureBox picCPI 
         BackColor       =   &H00C6B8A4&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   2055
         TabIndex        =   4
         Top             =   240
         Width           =   2055
         Begin VB.OptionButton opt12cpi 
            BackColor       =   &H00C6B8A4&
            Caption         =   "12 cpi"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1080
            TabIndex        =   6
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton opt10cpi 
            BackColor       =   &H00C6B8A4&
            Caption         =   "10 cpi"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   0
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "frmPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public PRINT_TRANSACTION As Long
Const is_GENERAL = 0
Const is_PO = 1
Const is_RR = 2
Const is_CHARGE_INVOICE = 3
Const is_CV = 4
Const is_CHECK = 5

Public LastPage     As Long
Public PrintType    As Long

Dim lhPrinter       As Long
Dim lReturn         As Long
Dim lpcWritten      As Long
Dim lDoc            As Long
Dim sWrittenData    As String
Dim MyDocInfo       As DOCINFO

Dim iPrinter As Integer
Dim pDeviceName As String

Dim sSupplierInfo, Array1, iRRLine, sItem


Private cSetPrinter As New clsSetDefaultPrinter

Dim i, iDefaultPrinter, pstrt, strt, stp, LPage, dblDiscount, strDiscount


Private Function Print_Report_All()
lReturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)
If lReturn = 0 Then
    MsgBox "The Printer Name you typed wasn't recognized."
Exit Function
End If

MyDocInfo.pDocName = ""
MyDocInfo.pOutputFile = vbNullString
MyDocInfo.pDatatype = vbNullString
lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
Call StartPagePrinter(lhPrinter)
        
If PrintType = 1 Then
    
    ' 12 cpi
    sWrittenData = (Chr(27) & Chr(40) & Chr(7)) & Chr(27) & Chr(15) & (Chr(27) & Chr(77))
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
        Len(sWrittenData), lpcWritten)
        
    For i = 1 To frmPreview.lstReport.ListItems.Count
        sWrittenData = frmPreview.lstReport.ListItems.Item(i).Text & vbCrLf
        lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
            Len(sWrittenData), lpcWritten)
    Next i
    
Else

    ' 10 cpi
    sWrittenData = (Chr(27) & Chr(40) & Chr(7)) & Chr(27) & Chr(15) & (Chr(27) & Chr(80))
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
        Len(sWrittenData), lpcWritten)
        
    For i = 1 To frmPreview.lstReport.ListItems.Count
        sWrittenData = frmPreview.lstReport.ListItems.Item(i).Text & vbCrLf
        lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
            Len(sWrittenData), lpcWritten)
    Next i
    
End If


'sWrittenData = Chr(27) & Chr(15) & vbFormFeed
sWrittenData = Chr(12) ' & Chr(15) & vbFormFeed
lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
    Len(sWrittenData), lpcWritten)
    
lReturn = EndPagePrinter(lhPrinter)
lReturn = EndDocPrinter(lhPrinter)
lReturn = ClosePrinter(lhPrinter)

End Function

Private Function Print_Report_Per_Page(PageF As Integer, PageT As Integer)



lReturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)
If lReturn = 0 Then
    MsgBox "The Printer Name you typed wasn't recognized."
Exit Function
End If

MyDocInfo.pDocName = ""
MyDocInfo.pOutputFile = vbNullString
MyDocInfo.pDatatype = vbNullString
lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
Call StartPagePrinter(lhPrinter)


   ' Get Starting Page
    
    pstrt = 1
    
    If PageF = 1 Then
    
        strt = 1
        
    Else
        
        For i = 1 To frmPreview.lstReport.ListItems.Count
        
           If frmPreview.lstReport.ListItems.Item(i).Text = Chr(12) Then
             pstrt = pstrt + 1
           End If
           
           If CDbl(pstrt) = CDbl(PageF) Then
               strt = i + 1
               Exit For
           End If
        Next i
    
    End If
    
    ' Get Ending Page
    pstrt = 1
    
    If PageT = 1 Then
    
        For i = 1 To frmPreview.lstReport.ListItems.Count
           If frmPreview.lstReport.ListItems.Item(i).Text = Chr(12) Then
             pstrt = pstrt + 1
           End If
           
           If CDbl(pstrt) = CDbl(PageT) Then
               stp = i
               Exit For
           End If
        Next i
      
      
    Else
    
        For i = 1 To frmPreview.lstReport.ListItems.Count
           If frmPreview.lstReport.ListItems.Item(i).Text = Chr(12) Then
             pstrt = pstrt + 1
           End If
           
           If CDbl(pstrt) = CDbl(PageT) Then
               stp = i + 1
               Exit For
           End If
        Next i
        
    End If
        
    
    'Get Last Page
    
    LPage = 0
    For i = stp To frmPreview.lstReport.ListItems.Count
       If frmPreview.lstReport.ListItems.Item(i).Text = Chr(12) Then
          Exit For
       Else
       LPage = LPage + 1
       End If
    Next i
    
'sWrittenData = (Chr(27) & Chr(40) & Chr(7))
'    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
'        Len(sWrittenData), lpcWritten)

'sWrittenData = (Chr(27) & Chr(40) & Chr(7)) & (Chr(27) & Chr(69)) & "" & (Chr(27) & Chr(70)) & vbCrLf

'sWrittenData = (Chr(27) & Chr(40) & Chr(7)) & (Chr(27) & Chr(112) & 1) & (Chr(27) & Chr(119) & 1) & (Chr(27) & Chr(69)) & "" & (Chr(27) & Chr(112) & 0) & (Chr(27) & Chr(119) & 0) & (Chr(27) & Chr(70)) & vbCrLf
'lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
'        Len(sWrittenData), lpcWritten)

If PrintType = 1 Then

    '   12 cpi
    sWrittenData = (Chr(27) & Chr(40) & Chr(7)) & Chr(27) & Chr(15) & (Chr(27) & Chr(77))
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
        Len(sWrittenData), lpcWritten)
        
    For i = strt To (stp + LPage) - 1
        sWrittenData = frmPreview.lstReport.ListItems.Item(i).Text & vbCrLf
        lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
            Len(sWrittenData), lpcWritten)
    Next i
    
Else
    
    '   10 cpi
    sWrittenData = (Chr(27) & Chr(40) & Chr(7)) & Chr(27) & Chr(15) & (Chr(27) & Chr(80))
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
        Len(sWrittenData), lpcWritten)
        
    For i = strt To (stp + LPage) - 1
        sWrittenData = frmPreview.lstReport.ListItems.Item(i).Text & vbCrLf
        lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
            Len(sWrittenData), lpcWritten)
    Next i
    
End If


'sWrittenData = Chr(27) & Chr(15) & vbFormFeed
sWrittenData = Chr(12) ' & Chr(15) & vbFormFeed
lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
    Len(sWrittenData), lpcWritten)
    
lReturn = EndPagePrinter(lhPrinter)
lReturn = EndDocPrinter(lhPrinter)
lReturn = ClosePrinter(lhPrinter)
'lstReport.SetFocus

End Function

Private Sub PRINT_RR()
With frmInvRR
    lReturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)
    If lReturn = 0 Then
        MsgBox "The Printer Name you typed wasn't recognized."
    Exit Sub
    End If
    
    sSupplierInfo = ""
    s = "SELECT SupplierName, " & _
        " Address1 + ' ' + Address2 + ' ' + Address3 as Address, " & _
        " TelNo + ', FAX # ' + FaxNo as ContactNo " & _
        " FROM tbl_Inv_Supplier " & _
        " WHERE (PK = " & .iSupplier & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        sSupplierInfo = rs!SupplierName & "|" & _
                        rs!Address & "|" & _
                        rs!ContactNo
    End If
    If rs.State = adStateOpen Then rs.Close
    
    Array1 = Split(CStr(sSupplierInfo), "|", -1, 1)
    
    MyDocInfo.pDocName = App.EXEName & "_RR"
    MyDocInfo.pOutputFile = vbNullString
    MyDocInfo.pDatatype = vbNullString
    lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
    Call StartPagePrinter(lhPrinter)
    
'====================
    For i = 1 To 5
        sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(15)) & "" & (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(15)) & vbCrLf
        lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
            Len(sWrittenData), lpcWritten)
    Next i
    
    sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(69)) & (Chr(27) & Chr(15)) & Space(18) & Mid(FORMATINYI(CStr(Array1(0))), 1, 78) & Space(78 - Len(Mid(FORMATINYI(CStr(Array1(0))), 1, 78))) & Format(.txtRRDate.Text, "mm") & Space(4) & Format(.txtRRDate.Text, "dd") & Space(4) & Format(.txtRRDate.Text, "yyyy") & (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(70)) & (Chr(27) & Chr(15)) & vbCrLf
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
        Len(sWrittenData), lpcWritten)
                
    sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(69)) & (Chr(27) & Chr(15)) & Space(18) & Mid(FORMATINYI(CStr(Array1(1))), 1, 88) & Space(88 - Len(Mid(FORMATINYI(CStr(Array1(1))), 1, 88))) & (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(70)) & (Chr(27) & Chr(15)) & vbCrLf
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
        Len(sWrittenData), lpcWritten)
    
    sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(69)) & (Chr(27) & Chr(15)) & Space(106) & .txtPONumber.Text & (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(70)) & (Chr(27) & Chr(15)) & vbCrLf
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
        Len(sWrittenData), lpcWritten)
    
    sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(69)) & (Chr(27) & Chr(15)) & Space(18) & Mid(FORMATINYI(Trim(.txtInvNumber.Text)), 1, 88) & Space(88 - Len(Mid(FORMATINYI(Trim(.txtInvNumber.Text)), 1, 88))) & Trim(.txtRefNo.Text) & (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(70)) & (Chr(27) & Chr(15)) & vbCrLf
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
        Len(sWrittenData), lpcWritten)
    
    For i = 1 To 3
        sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(15)) & "" & (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(15)) & vbCrLf
        lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
            Len(sWrittenData), lpcWritten)
    Next i
'================== Detail
    iRRLine = 0
    For i = 1 To .lstDetail.ListItems.Count
        With .lstDetail.ListItems
            
            'sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(69)) & (Chr(27) & Chr(15)) & _
                Space(15 - Len(.Item(i).SubItems(4))) & .Item(i).SubItems(4) & _
                Space(8) & Mid(.Item(i).SubItems(5), 1, 5) & _
                Space(2 - Len(Mid(.Item(i).SubItems(5), 1, 5))) & "[" & .Item(i).SubItems(6) & "] " & Mid(.Item(i).SubItems(7), 1, 45) & _
                Space(62 - Len("[" & .Item(i).SubItems(6) & "] " & Mid(.Item(i).SubItems(7), 1, 45))) & _
                Space(13 - Len(Format(.Item(i).SubItems(9), "#,##0.00"))) & Format(.Item(i).SubItems(9), "#,##0.00") & _
                Space(13 - Len(Format(.Item(i).SubItems(10), "#,##0.00"))) & Format(.Item(i).SubItems(10), "#,##0.00") & _
                Space(15) & .Item(i).SubItems(16) & _
                (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(70)) & (Chr(27) & Chr(15)) & vbCrLf
                'lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
                    Len(sWrittenData), lpcWritten)
            
            If CDbl(IIf(IsNumeric(.Item(i).SubItems(4)) = False, 0, .Item(i).SubItems(4))) > 0 Then
                iRRLine = iRRLine + 1
                sItem = Mid("[" & .Item(i).SubItems(6) & "] " & .Item(i).SubItems(7), 1, 42)
                sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(69)) & (Chr(27) & Chr(15)) & _
                    Space(15 - Len(.Item(i).SubItems(4))) & .Item(i).SubItems(4) & _
                    Space(4) & Mid(.Item(i).SubItems(5), 1, 5) & Space(5 - Len(Mid(.Item(i).SubItems(5), 1, 5))) & _
                    Space(4) & sItem & Space(42 - Len(sItem)) & _
                    Space(12 - Len(Format(.Item(i).SubItems(9), "#,##0.00"))) & Format(.Item(i).SubItems(9), "#,##0.00") & _
                    Space(12 - Len(Format(.Item(i).SubItems(10), "#,##0.00"))) & Format(.Item(i).SubItems(10), "#,##0.00") & _
                    Space(23) & Mid(.Item(i).SubItems(16), 1, 15) & _
                    (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(70)) & (Chr(27) & Chr(15)) & vbCrLf
                lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
                    Len(sWrittenData), lpcWritten)
            End If
            
        End With
    Next i
    
    iRRLine = iRRLine + 1
    sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(15)) & Space(66) & "================================" & (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(15)) & vbCrLf
        lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
            Len(sWrittenData), lpcWritten)
            
    iRRLine = iRRLine + 1
    sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(69)) & (Chr(27) & Chr(15)) & Space(28) & "TOTAL >>" & Space(58 - Len(.lblTotalNetCost.Caption)) & .lblTotalNetCost.Caption & (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(70)) & (Chr(27) & Chr(15)) & vbCrLf
        lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
            Len(sWrittenData), lpcWritten)
    
    iRRLine = iRRLine + 1
    sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(69)) & (Chr(27) & Chr(15)) & Space(2) & Space(Center_Object("--------------- Nothing Follows ---------------", 66)) & "--------------- Nothing Follows ---------------" & (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(70)) & (Chr(27) & Chr(15)) & vbCrLf
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
        Len(sWrittenData), lpcWritten)
    
    'Line Space
    For i = 1 To (15 - CDbl(iRRLine))
        sWrittenData = (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(15)) & "" & vbCrLf
        lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
            Len(sWrittenData), lpcWritten)
    Next i
    
    For i = 1 To 3
        sWrittenData = (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(15)) & "" & vbCrLf
        lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
            Len(sWrittenData), lpcWritten)
    Next i
    
    sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(69)) & (Chr(27) & Chr(15)) & Space(15) & .txtStockClerk.Text & Space(30 - Len(.txtStockClerk.Text)) & _
                   .txtPurchaser.Text & Space(30 - Len(.txtPurchaser.Text)) & .txtDeptHead.Text & (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(70)) & (Chr(27) & Chr(15)) & vbCrLf
        lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
            Len(sWrittenData), lpcWritten)
            
    sWrittenData = Chr(27) & Chr(15) & vbFormFeed
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
        Len(sWrittenData), lpcWritten)
    lReturn = EndPagePrinter(lhPrinter)
    lReturn = EndDocPrinter(lhPrinter)
    lReturn = ClosePrinter(lhPrinter)
    
End With
Unload Me
End Sub

Private Sub PRINT_PO()
With frmInvPO
    lReturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)
    If lReturn = 0 Then
        MsgBox "The Printer Name you typed wasn't recognized."
    Exit Sub
    End If
    
    sSupplierInfo = ""
    s = "SELECT SupplierName, " & _
        " Address1 + ' ' + Address2 + ' ' + Address3 as Address, " & _
        " TelNo + ', FAX # ' + FaxNo as ContactNo " & _
        " FROM tbl_Inv_Supplier " & _
        " WHERE (PK = " & .txtSuppKey.Text & ")"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        sSupplierInfo = rs!SupplierName & "|" & _
                        rs!Address & "|" & _
                        rs!ContactNo
    End If
    If rs.State = adStateOpen Then rs.Close
    
    Array1 = Split(CStr(sSupplierInfo), "|", -1, 1)
    
    MyDocInfo.pDocName = App.EXEName & "_PO"
    MyDocInfo.pOutputFile = vbNullString
    MyDocInfo.pDatatype = vbNullString
    lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
    Call StartPagePrinter(lhPrinter)
    
    
'====================
    For i = 1 To 7
        sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(15)) & "" & (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(15)) & vbCrLf
        lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
            Len(sWrittenData), lpcWritten)
    Next i
    
    sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(69)) & (Chr(27) & Chr(15)) & Space(18) & FORMATINYI(CStr(Array1(0))) & (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(70)) & (Chr(27) & Chr(15)) & vbCrLf
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
        Len(sWrittenData), lpcWritten)
            
    sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(69)) & (Chr(27) & Chr(15)) & Space(18) & Mid(FORMATINYI(CStr(Array1(1))), 1, 100) & Space(100 - Len(Mid(FORMATINYI(CStr(Array1(1))), 1, 100))) & .txtPODate.Text & (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(70)) & (Chr(27) & Chr(15)) & vbCrLf
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
        Len(sWrittenData), lpcWritten)
            
    sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(69)) & (Chr(27) & Chr(15)) & Space(118) & .txtRefNo.Text & (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(70)) & (Chr(27) & Chr(15)) & vbCrLf
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
        Len(sWrittenData), lpcWritten)
            
    sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(69)) & (Chr(27) & Chr(15)) & Space(18) & FORMATINYI(CStr(Array1(2))) & Space(100 - Len(Mid(FORMATINYI(CStr(Array1(2))), 1, 100))) & .txtTerms.Text & (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(70)) & (Chr(27) & Chr(15)) & vbCrLf
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
        Len(sWrittenData), lpcWritten)
    
    For i = 1 To 2
        sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(15)) & "" & (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(15)) & vbCrLf
        lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
            Len(sWrittenData), lpcWritten)
    Next i
    
    '(Chr(27) & Chr(69)) &
    For i = 1 To .lstDetail.ListItems.Count
        With .lstDetail.ListItems
            'sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(69)) & (Chr(27) & Chr(15)) & _
                Space(10 - Len(.Item(i).SubItems(2))) & .Item(i).SubItems(2) & _
                Space(8) & .Item(i).SubItems(3) & _
                Space(13 - Len(.Item(i).SubItems(3))) & "[" & .Item(i).SubItems(4) & "] " & Mid(.Item(i).SubItems(5), 1, 45) & _
                Space(62 - Len("[" & .Item(i).SubItems(4) & "] " & Mid(.Item(i).SubItems(5), 1, 45))) & _
                Space(16 - Len(Format(.Item(i).SubItems(7), "#,##0.00"))) & Format(.Item(i).SubItems(7), "#,##0.00") & _
                Space(23 - Len(Format(.Item(i).SubItems(8), "#,##0.00"))) & Format(.Item(i).SubItems(8), "#,##0.00") & _
                (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(70)) & (Chr(27) & Chr(15)) & vbCrLf
                'lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
                    Len(sWrittenData), lpcWritten)
            
            sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(69)) & (Chr(27) & Chr(15)) & _
                Space(10 - Len(.Item(i).SubItems(4))) & .Item(i).SubItems(4) & _
                Space(8) & .Item(i).SubItems(5) & _
                Space(13 - Len(.Item(i).SubItems(5))) & "[" & .Item(i).SubItems(6) & "] " & Mid(.Item(i).SubItems(7), 1, 45) & _
                Space(62 - Len("[" & .Item(i).SubItems(6) & "] " & Mid(.Item(i).SubItems(7), 1, 45))) & _
                Space(16 - Len(Format(.Item(i).SubItems(9), "#,##0.00"))) & Format(.Item(i).SubItems(9), "#,##0.00") & _
                Space(23 - Len(Format(.Item(i).SubItems(10), "#,##0.00"))) & Format(.Item(i).SubItems(10), "#,##0.00") & _
                (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(70)) & (Chr(27) & Chr(15)) & vbCrLf
                lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
                    Len(sWrittenData), lpcWritten)
        
        End With
    Next i
    
    sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(69)) & (Chr(27) & Chr(15)) & Space(2) & Space(Center_Object("--------------- Nothing Follows ---------------", 66)) & "--------------- Nothing Follows ---------------" & (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(70)) & (Chr(27) & Chr(15)) & vbCrLf
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
        Len(sWrittenData), lpcWritten)
    
    'Line Space
    For i = 1 To (13 - .lstDetail.ListItems.Count)
        sWrittenData = (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(15)) & "" & vbCrLf
        lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
            Len(sWrittenData), lpcWritten)
    Next i
    
    dblDiscount = (CDbl(.lblTotalCost.Caption) - CDbl(.lblTotalNetCost.Caption)) * -1
    strDiscount = IIf(CDbl(dblDiscount) = 0, "", Format(dblDiscount, "#,##0.00"))
    
    'Total Amount
    sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(69)) & (Chr(27) & Chr(15)) & _
                   Space(94) & Space(15 - Len(strDiscount)) & strDiscount & Space(23 - Len(.lblTotalCost.Caption)) & .lblTotalNetCost.Caption & _
                   (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(70)) & (Chr(27) & Chr(15)) & vbCrLf
        lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
            Len(sWrittenData), lpcWritten)
            
    
    sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(15)) & "" & vbCrLf
        lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
            Len(sWrittenData), lpcWritten)
    
    'Remarks/Purposes
    sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(69)) & (Chr(27) & Chr(15)) & Space(31) & FORMATINYI(.txtRemarks.Text) & (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(70)) & (Chr(27) & Chr(15)) & vbCrLf
        lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
            Len(sWrittenData), lpcWritten)
    
    For i = 1 To 2
        sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(15)) & "" & vbCrLf
        lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
            Len(sWrittenData), lpcWritten)
    Next i
    
    sWrittenData = (Chr(27) & Chr(107) & 0) & (Chr(27) & Chr(69)) & (Chr(27) & Chr(15)) & Space(2) & _
        Space(Center_Object(FORMATINYI(.txtRequested.Text), 20)) & FORMATINYI(.txtRequested.Text) & Space(Center_Object(FORMATINYI(.txtRequested.Text), 20)) & Space(5) & _
        Space(Center_Object(FORMATINYI(.txtChecked.Text), 20)) & FORMATINYI(.txtChecked.Text) & Space(Center_Object(FORMATINYI(.txtChecked.Text), 20)) & Space(5) & _
        Space(Center_Object(FORMATINYI(.txtApproved.Text), 20)) & FORMATINYI(.txtApproved.Text) & Space(Center_Object(FORMATINYI(.txtApproved.Text), 20)) & _
        (Chr(27) & Chr(107) & 2) & (Chr(27) & Chr(70)) & (Chr(27) & Chr(15)) & vbCrLf
        lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
            Len(sWrittenData), lpcWritten)
    
    sWrittenData = Chr(27) & Chr(15) & vbFormFeed
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
        Len(sWrittenData), lpcWritten)
    lReturn = EndPagePrinter(lhPrinter)
    lReturn = EndDocPrinter(lhPrinter)
    lReturn = ClosePrinter(lhPrinter)
    
    ConnOmega.Execute "UPDATE tbl_Inv_PO SET Printed = 1 WHERE (PK = " & .Statusbar1.Panels(1).Text & ")"
    
    Unload Me
End With
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If cmbPrinter.ListIndex = -1 Then Exit Sub
pDeviceName = cmbPrinter.List(cmbPrinter.ListIndex)
If cSetPrinter.SetPrinterAsDefault(pDeviceName) = False Then MsgBox pDeviceName & " has failed to be set as the default printer.                  ", vbCritical, "Error...": Exit Sub
SaveSetting App.EXEName, "DefaultPrinter", "DftPntr", cmbPrinter.ListIndex
'Exit Sub
Select Case PRINT_TRANSACTION
    Case is_GENERAL:
        
        If optAll.Value = True Then
            For i = 1 To RETURNTEXTVALUE(txtCopies)
                Print_Report_All
            Next i
        End If
        
        If optPages.Value = True Then
            For i = 1 To RETURNTEXTVALUE(txtCopies)
                Print_Report_Per_Page txtPgFrom.Text, txtPgTo.Text
            Next i
        End If
        
        Unload Me
        
    Case is_PO:             PRINT_PO
    Case is_RR:             PRINT_RR
    Case is_CHARGE_INVOICE:
    Case is_CV
        Dim ReportCVPre As New rptCheckVoucher_PrePrinted
        ReportCVPre.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
        ReportCVPre.DiscardSavedData
        ReportCVPre.ParameterFields(1).ClearCurrentValueAndRange
        ReportCVPre.ParameterFields(1).AddCurrentValue CStr(gbl_UserName)
        ReportCVPre.PrintOut False
        Unload Me
    Case is_CHECK
        Dim ReportCheck As New rptCheckVoucher_Check
        ReportCheck.Database.LogOnServer "Pdsodbc.dll", gbl_Database, gbl_Database, sLogIn, sPassword
        ReportCheck.DiscardSavedData
        ReportCheck.ParameterFields(1).ClearCurrentValueAndRange
        ReportCheck.ParameterFields(1).AddCurrentValue CStr(gbl_UserName)
        ReportCheck.PrintOut False
        Unload Me
    Case Else: Exit Sub
End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Height = 3480
Me.Width = 5535
cmbPrinter.Clear
For iPrinter = 0 To Printers.Count - 1
    cmbPrinter.AddItem Printers(iPrinter).DeviceName
    'cmbPrinter.ListIndex = 0
Next iPrinter
iDefaultPrinter = GetSetting(App.EXEName, "DefaultPrinter", "DftPntr", 0)
If cmbPrinter.ListCount = 0 Then
    iDefaultPrinter = -1
Else
    If cmbPrinter.ListCount < CDbl(iDefaultPrinter) Then
        iDefaultPrinter = 0
    End If
End If
cmbPrinter.ListIndex = CInt(iDefaultPrinter)

optAll.Value = True
txtPgFrom.Text = "1"
txtPgTo.Text = LastPage
txtCopies.Text = "1"
opt10cpi.Value = True
End Sub
