Attribute VB_Name = "modTransparent"
Option Explicit


Public Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const LWA_ALPHA = &H2


Public Function TransparentEffect(hwnd As Long, Rate As Byte)
    Dim WinInfo As Long
    WinInfo = GetWindowLong(hwnd, GWL_EXSTYLE)
    WinInfo = WinInfo Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, WinInfo
    SetLayeredWindowAttributes hwnd, 0, Rate, LWA_ALPHA
End Function

Public Function GET_OS_VERSION() As Long
Dim verinfo As OSVERSIONINFO
Dim ret As Integer
verinfo.dwOSVersionInfoSize = Len(verinfo)
ret = GetVersionEx(verinfo)
GET_OS_VERSION = verinfo.dwPlatformId
End Function
