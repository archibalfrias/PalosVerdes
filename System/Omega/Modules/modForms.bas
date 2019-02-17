Attribute VB_Name = "modForms"
' *********************************************
' * Code by Robert Wright - <rob@xenonic.com> *
' *********************************************

' *********************************************
' *        -> This code is FREEWARE <-        *
' * You are free to use this any of this VB   *
' * Project (including the images) in your    *
' * own programs.                             *
' *                                           *
' * All I ask is that you vote for me on PSC! *
' *********************************************

' Used to set the shape of the form
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
' Used to create the rounded rectangle region
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
' Used to make the form draggable
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
' Also used to make the form draggable
Public Declare Function ReleaseCapture Lib "user32" () As Long
' Used to make the window always on top
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
' Various constants used by the above functions
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Sub DoTransparency(TheForm As Form)
' TheForm:  The form you want to be rounded rectangle shape
    
    Dim TempRegions(6) As Long
    Dim FormWidthInPixels As Long
    Dim FormHeightInPixels As Long
    Dim a
    
' Convert the form's height and width from twips to pixels
    FormWidthInPixels = TheForm.Width / Screen.TwipsPerPixelX
    FormHeightInPixels = TheForm.Height / Screen.TwipsPerPixelY
    
' Make a rounded rectangle shaped region with the dimentions of the form
    a = CreateRoundRectRgn(0, 0, FormWidthInPixels, FormHeightInPixels, 24, 24)
    
' Set this region as the shape for "TheForm"
    a = SetWindowRgn(TheForm.hWnd, a, True)
End Sub

Public Sub DoDrag(TheForm As Form)
' TheForm:  The form you want to start dragging
    
    ReleaseCapture
    SendMessage TheForm.hWnd, &HA1, 2, 0&
End Sub

Public Sub MakeWindowInstantMessaging(TheForm As Form)
' TheForm:  The form you want to make graphical
    TheForm.BackColor = RGB(207, 207, 207)
    Dim i
'    For i = 0 To 67
'        TheForm!Shape1(i).BackColor = RGB(207, 207, 207)
'        TheForm!Shape1(i).Visible = False
'    Next i
'    For i = 0 To 12
''        TheForm!Label1(i).ForeColor = RGB(207, 207, 207)
'        TheForm!Label1(i).ForeColor = &HFF00&
'    Next i
'    For Each ShapeName In Shapes
''        ShapeNameTemp = "Shape" & CStr(i)
''        Set ShapeName = TheForm!ShapeNameTemp
'        TheForm!ShapeName.BackColor = RGB(207, 207, 207)
'    Next
    
    TheForm.Caption = TheForm!lblTitle.Caption
    TheForm!lblTitle.Left = 16
    TheForm!lblTitle.Top = 7
    
    With TheForm!imgTitleLeft
        .Top = 0
        .Left = 0
    End With
    
    With TheForm!imgTitleRight
        .Top = 0
        .Left = (TheForm.Width / Screen.TwipsPerPixelX) - 19
    End With
    
    With TheForm!imgTitleMain
        .Top = 0
        .Left = 19
        .Width = (TheForm.Width / Screen.TwipsPerPixelX) - 19
    End With
    
    With TheForm!imgWindowLeft
        .Top = 30
        .Left = 0
        .Height = (TheForm.Height / Screen.TwipsPerPixelY) - 60
    End With
    
    With TheForm!imgWindowBottomLeft
        .Top = (TheForm.Height / Screen.TwipsPerPixelY) - 30
        .Left = 0
    End With
    
    With TheForm!imgWindowBottom
        .Top = (TheForm.Height / Screen.TwipsPerPixelY) - 30
        .Left = 19
        .Width = (TheForm.Width / Screen.TwipsPerPixelX) - 38
    End With
    
    With TheForm!imgWindowBottomRight
        .Top = (TheForm.Height / Screen.TwipsPerPixelY) - 30
        .Left = (TheForm.Width / Screen.TwipsPerPixelX) - 19
    End With
    
    With TheForm!imgWindowRight
        .Top = 30
        .Left = (TheForm.Width / Screen.TwipsPerPixelX) - 19
        .Height = (TheForm.Height / Screen.TwipsPerPixelY) - 60 '38
    End With
    
    With TheForm!imgTitleClose
        .Top = 8
        .Left = (TheForm.Width / Screen.TwipsPerPixelX) - 22
    End With
    
    With TheForm!imgTitleMinimize
        .Top = 8
        .Left = (TheForm.Width / Screen.TwipsPerPixelX) - 39
    End With
    
    With TheForm!imgTitleHelp
        .Top = 8
        .Left = (TheForm.Width / Screen.TwipsPerPixelX) - 56
    End With
    
    DoTransparency TheForm
    
End Sub



