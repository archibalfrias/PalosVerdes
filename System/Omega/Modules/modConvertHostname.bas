Attribute VB_Name = "modConvertHostname"
Option Explicit

Private Const IP_SUCCESS As Long = 0
Private Const MAX_WSADescription As Long = 256
Private Const MAX_WSASYSStatus As Long = 128
Private Const WS_VERSION_REQD As Long = &H101
Private Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD As Long = 1
Private Const SOCKET_ERROR As Long = -1
Private Const ERROR_SUCCESS As Long = 0

Private Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To MAX_WSADescription) As Byte
   szSystemStatus(0 To MAX_WSASYSStatus) As Byte
   wMaxSockets As Long
   wMaxUDPDG As Long
   dwVendorInfo As Long
End Type

Private Declare Function gethostbyname Lib "wsock32.dll" _
  (ByVal hostname As String) As Long
 
Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (xDest As Any, _
   xSource As Any, _
   ByVal nbytes As Long)

Private Declare Function lstrlenA Lib "kernel32" _
  (lpString As Any) As Long

Private Declare Function WSAStartup Lib "wsock32.dll" _
   (ByVal wVersionRequired As Long, _
    lpWSADATA As WSADATA) As Long
   
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long

Private Declare Function inet_ntoa Lib "wsock32.dll" _
  (ByVal addr As Long) As Long

Private Declare Function lstrcpyA Lib "kernel32" _
  (ByVal RetVal As String, _
   ByVal Ptr As Long) As Long
                       
Private Declare Function gethostname Lib "wsock32.dll" _
   (ByVal szHost As String, _
    ByVal dwHostLen As Long) As Long

Public Declare Function GetRTTAndHopCount _
    Lib "iphlpapi.dll" _
   (ByVal lDestIPAddr As Long, _
    ByRef lHopCount As Long, _
    ByVal lMaxHops As Long, _
    ByRef lRTT As Long) As Long
        
Public Declare Function inet_addr _
    Lib "wsock32.dll" _
   (ByVal cp As String) As Long



Public Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA
   Dim success As Long
  
   SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
   
End Function


Public Sub SocketsCleanup()
  
   If WSACleanup() <> 0 Then
       MsgBox "Windows Sockets error occurred in Cleanup.", vbExclamation
   End If
   
End Sub
  

Public Function GetMachineName() As String

   Dim sHostName As String * 256
  
   If gethostname(sHostName, 256) = ERROR_SUCCESS Then
      GetMachineName = Trim$(sHostName)
   End If
  
End Function


Public Function GetIPFromHostName(ByVal sHostName As String) As String

  'converts a host name to an IP address

   Dim nbytes As Long
   Dim ptrHosent As Long  'address of HOSENT structure
   Dim ptrName As Long    'address of name pointer
   Dim ptrAddress As Long 'address of address pointer
   Dim ptrIPAddress As Long
   Dim ptrIPAddress2 As Long

   ptrHosent = gethostbyname(sHostName & vbNullChar)

   If ptrHosent <> 0 Then

     'assign pointer addresses and offset

     'Null-terminated list of addresses for the host.
     'The Address is offset 12 bytes from the start of
     'the HOSENT structure. Note: Here we are retrieving
     'only the first address returned. To return more than
     'one, define sAddress as a string array and loop through
     'the 4-byte ptrIPAddress members returned. The last
     'item is a terminating null. All addresses are returned
     'in network byte order.
      ptrAddress = ptrHosent + 12
     
     'get the IP address
      CopyMemory ptrAddress, ByVal ptrAddress, 4
      CopyMemory ptrIPAddress, ByVal ptrAddress, 4
      CopyMemory ptrIPAddress2, ByVal ptrIPAddress, 4

      GetIPFromHostName = GetInetStrFromPtr(ptrIPAddress2)

   End If
  
End Function


Private Function GetStrFromPtrA(ByVal lpszA As Long) As String

   GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
  
End Function


Private Function GetInetStrFromPtr(Address As Long) As String
 
   GetInetStrFromPtr = GetStrFromPtrA(inet_ntoa(Address))

End Function

Public Function PingServer(prmIPaddrH As String) As Boolean
Dim IPaddr As Long, HopsCount As Long, RTT As Long
Dim sIpAddress As String
Dim MaxHops As Long

If SocketsInitialize() Then
    sIpAddress = GetIPFromHostName(prmIPaddrH)
    SocketsCleanup
Else
    sIpAddress = ""
End If

    Const success = 1
    MaxHops = 20               ' should be enough ...
    'sIpAddress = GetIPFromHostName(prmIPaddrH)
    IPaddr = inet_addr(sIpAddress)
    PingServer = (GetRTTAndHopCount(IPaddr, HopsCount, MaxHops, RTT) = success)
End Function
