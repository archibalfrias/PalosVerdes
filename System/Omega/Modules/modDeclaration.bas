Attribute VB_Name = "modDeclaration"

Public gbl_Item_Module      As String

Public TournamentKey        As Double
Public WithTeamPlay         As Integer
Public TeamPlayer2Cnt       As Integer
Public NoofPlayerPerTeam    As Integer
Public AllowedTeam          As Integer
Public WithIndividualPlay   As Integer
Public HandicapDivisor      As Integer
Public TeamDivisorOrder     As Integer
Public DaysPlayerToPlay     As Integer
Public TournamentName       As String
Public TournamentRange      As String
Public ScoringType          As Long
Public PointsToCnt          As Long
Public TopHandicap          As Double
Public PointsToCntIndi      As Long
Public TeamAverage          As Long
Public TopIndex             As Double
Public ParGrossPoints    As Double
Public LocationCnt          As Double
Public LocationKey          As Long

Public gbl_CompanyName      As String
Public gbl_CompanyAddress1  As String
Public gbl_CompanyAddress2  As String
Public gbl_CompanyTelNo     As String
Public gbl_CompanyFaxNo     As String
Public gbl_CompanySSSNo     As String
Public gbl_CompanyPHICNo    As String
Public gbl_CompanyTIN       As String

Public iAdmin               As Long

Public gbl_Server           As String
Public gbl_Database         As String
Public sLogIn               As String
Public sPassword            As String

Public gbl_ServerL          As String
Public gbl_DatabaseL        As String
Public sLogInL              As String
Public sPasswordL           As String

Public ConnOmega            As New ADODB.Connection

Public gbl_FORM             As Form
Public gbl_FORMx            As Object
Public gbl_FORM_Modal       As Long
Public gbl_Form_Caption     As String

Public SystemIdleTime       As Double
Public blnIsIdle            As Boolean

Public LogInWithOutLoading  As Long

Public gbl_LockWhenIdle     As Long
Public gbl_Idle_Time        As Double
Public gbl_Slides_Background As Long
Public gbl_Slides_Time      As Double
Public gbl_Quotes_Time      As Double

Public gbl_VAT              As Double
Public gbl_MinTakeHomePay   As Double
Public gbl_TotalEarning     As Double

Public xlsApp               As Object

Public PassStartWizard      As Long

Public iMsgLoaded           As Long

Public iTreeViewIndex       As Long

Public gbl_MpnthlyDivisor   As Double

Public a                    As String
Public ra                   As New ADODB.Recordset
Public AA                   As String
Public raa                  As New ADODB.Recordset
Public b                    As String
Public rb                   As New ADODB.Recordset
Public C                    As String
Public rc                   As New ADODB.Recordset
Public D                    As String
Public rd                   As New ADODB.Recordset
Public s                    As String
Public rs                   As New ADODB.Recordset
Public t                    As String
Public rt                   As New ADODB.Recordset
Public u                    As String
Public ru                   As New ADODB.Recordset
Public v                    As String
Public rv                   As New ADODB.Recordset
Public w                    As String
Public rw                   As New ADODB.Recordset
Public z                    As String
Public rz                   As New ADODB.Recordset

Public C_Application        As New CRAXDDRT.Application
Public C_Report             As New CRAXDDRT.Report

'------- Sleep PC
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'------- PRINTING
Public Type DOCINFO
    pDocName As String
    pOutputFile As String
    pDatatype As String
End Type

Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Public Declare Function EndDocPrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Public Declare Function EndPagePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Public Declare Function OpenPrinter Lib "winspool.drv" Alias _
   "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
    ByVal pDefault As Long) As Long
Public Declare Function StartDocPrinter Lib "winspool.drv" Alias _
   "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, _
   pDocInfo As DOCINFO) As Long
Public Declare Function StartPagePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Public Declare Function WritePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, _
   pcWritten As Long) As Long
   
'------- Get Printer Information
Public Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, ByRef pPrinter As Any, ByVal cbBuf As Long, ByRef pcbNeeded As Long) As Long
Public Declare Function IsBadStringPtrByLong Lib "kernel32" Alias "IsBadStringPtrA" (ByVal lpsz As Long, ByVal ucchMax As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Type PRINTER_INFO_2
   pServerName As String
   pPrinterName As String
   pShareName As String
   pPortName As String
   pDriverName As String
   pComment As String
   pLocation As String
   pDevMode As Long
   pSepFile As String
   pPrintProcessor As String
   pDatatype As String
   pParameters As String
   pSecurityDescriptor As Long
   Attributes As Long
   Priority As Long
   DefaultPriority As Long
   StartTime As Long
   UntilTime As Long
   Status As Long
   JobsCount As Long
   AveragePPM As Long
End Type

Private Const PRINTER_ENUM_CONNECTIONS = &H4
Private Const PRINTER_ENUM_LOCAL = &H2
Private Const PRINTER_CONTROL_RESUME = 2
Private Const PRINTER_CONTROL_SET_STATUS = 4
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const PRINTER_ATTRIBUTE_WORK_OFFLINE = &H400

Private Const ERROR_ACCESS_DENIED = 5           ' Access is denied.
Private Const ERROR_BAD_NETPATH = 53            ' The network path was not found
Private Const ERROR_UNEXP_NET_ERR = 59          ' An unexpected network error occurred
Private Const ERROR_INSUFFICIENT_BUFFER = 122   ' The data area passed to a system call is too small
Private Const ERROR_NETWORK_UNREACHABLE = 1231  ' The remote network is not reachable by the transport
Private Const RPC_S_SERVER_UNAVAILABLE = 1722   ' The RPC server is unavailable
Private Const RPC_S_CALL_FAILED = 1726          ' The remote procedure call failed

Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Type PRINTER_DEFAULTS
    pDatatype As String
    pDevMode As DEVMODE
    DesiredAccess As Long
End Type

Public PrinterDriverUse As String
Public pPrinterNames() As String
Public pShareNames() As String
Public pNrPrinters As Integer

'' DOS PRINTING
'Public Type DOCINFO
'    pDocName As String
'    pOutputFile As String
'    pDatatype As String
'End Type
'
'Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal _
'   hPrinter As Long) As Long
'Public Declare Function EndDocPrinter Lib "winspool.drv" (ByVal _
'   hPrinter As Long) As Long
'Public Declare Function EndPagePrinter Lib "winspool.drv" (ByVal _
'   hPrinter As Long) As Long
'Public Declare Function OpenPrinter Lib "winspool.drv" Alias _
'   "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
'    ByVal pDefault As Long) As Long
'Public Declare Function StartDocPrinter Lib "winspool.drv" Alias _
'   "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, _
'   pDocInfo As DOCINFO) As Long
'Public Declare Function StartPagePrinter Lib "winspool.drv" (ByVal _
'   hPrinter As Long) As Long
'Public Declare Function WritePrinter Lib "winspool.drv" (ByVal _
'   hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, _
'   pcWritten As Long) As Long
''=================================================


'------- PUT OBJECT IN STATUSBAR
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, lparam As Any) As Long

'Private Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type


Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const WM_USER As Long = &H400
Public Const SB_GETRECT As Long = (WM_USER + 10)

'------- UPPER CASE
Public Const ES_UPPERCASE = &H8&
Public Const GWL_STYLE = (-16)
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'------- GET OS VERSION
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'------- MEMMORY STATUS
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
                
Private Type MEMORYSTATUS 'Type variable for memory info
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Public MEM_STAT As MEMORYSTATUS


'       Resize TreeView
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" _
                                   (ByVal hwnd As Long, ByVal wMsg As Long, _
                                    ByVal wParam As Long, lparam As Any) As Long

Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Sub ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI)
Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
'================================================

