Attribute VB_Name = "modCreateODBC"
'==================================
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const conSwNormal = 1
'===================================
Private Const KEY_QUERY_VALUE = &H1
Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const REG_DWORD = 4
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long ' Note that If you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hwndParent As Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long

Public Function DELETE_DNS_SQL_ODBC(sDSNName As String, _
sServerName As String, sDatabaseName As String, sUserID As String, _
sPassword As String)
Dim sAttributes As String
Dim iRetVal As Integer

sAttributes = "DSN=" & sDSNName & Chr$(0) & _
              "Description=SQL Server on server " & sServerName & Chr$(0) & _
              "SERVER=" & sServerName & Chr$(0) & _
              "Database=" & sDatabaseName & Chr$(0) & Chr$(0)
              
iRetVal = SQLConfigDataSource(0, 6, "SQL Server", sAttributes)

End Function

Public Function CREATE_SQL_DNS()
Dim szDriverName As String
Dim szWantedDSN As String

szDriverName = String(255, Chr(32))
szWantedDSN = gbl_Database

If Not checkSQLDriver(szDriverName) Then
    MsgBox "You must Install SQL ODBC Drivers before use this program.", vbOK + vbCritical
End If

If Not (checkWantedSQLDSN(szWantedDSN)) Then

    If szDriverName = "" Then
        MsgBox "Can't find SQL ODBC driver.", vbOK + vbCritical
    Else

        If Not createSQLDSN(szDriverName, szWantedDSN) Then
            MsgBox "Can't create database ODBC.", vbOK + vbCritical
        End If
    End If
End If

End Function

Public Function CREATE_SQL_DNS_Acctg()
Dim szDriverName As String
Dim szWantedDSN As String

szDriverName = String(255, Chr(32))
szWantedDSN = gbl_DATABASE_Acctg

If Not checkSQLDriver(szDriverName) Then
    MsgBox "You must Install SQL ODBC Drivers before use this program.", vbOK + vbCritical
End If

If Not (checkWantedSQLDSN(szWantedDSN)) Then

    If szDriverName = "" Then
        MsgBox "Can't find SQL ODBC driver.", vbOK + vbCritical
    Else

        If Not createSQLDSN_Acctg(szDriverName, szWantedDSN) Then
            MsgBox "Can't create database ODBC.", vbOK + vbCritical
        End If
    End If
End If

End Function


Public Function isSZKeyExist(szKeyPath As String, _
    szKeyName As String, _
    ByRef szKeyValue As String) As Boolean
    Dim bRes As Boolean
    Dim lRes As Long
    Dim hKey As Long
    lRes = RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
    szKeyPath, _
    0&, _
    KEY_QUERY_VALUE, _
    hKey)


    If lRes <> ERROR_SUCCESS Then
        isSZKeyExist = False
        Exit Function
    End If
    lRes = RegQueryValueEx(hKey, _
    szKeyName, _
    0&, _
    REG_SZ, _
    ByVal szKeyValue, _
    Len(szKeyValue))
    RegCloseKey (hKey)


    If lRes <> ERROR_SUCCESS Then
        isSZKeyExist = False
        Exit Function
    End If
    isSZKeyExist = True
End Function


Public Function checkSQLDriver(ByRef szDriverName As String) As Boolean
    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String
    Dim bRes As Boolean
    
    
    bRes = False
    
    szKeyPath = "SOFTWARE\ODBC\ODBCINST.INI\SQL Server"
    szKeyName = "Driver"
    szKeyValue = String(255, Chr(32))

    If isSZKeyExist(szKeyPath, szKeyName, szKeyValue) Then
        szDriverName = szKeyValue
        bRes = True
    Else
        bRes = False
    End If
    
    checkSQLDriver = bRes
End Function


Public Function checkWantedSQLDSN(szWantedDSN As String) As Boolean
    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String
    Dim bRes As Boolean
    
    szKeyPath = "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources"
    szKeyName = szWantedDSN
    szKeyValue = String(255, Chr(32))
    


    If isSZKeyExist(szKeyPath, szKeyName, szKeyValue) Then
        bRes = True
    Else
        bRes = False
    End If
    
    checkWantedSQLDSN = bRes
    
End Function


Public Function createSQLDSN(szDriverName As String, _
    szWantedDSN As String) As Boolean
    
    Dim hKey As Long
    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String
    Dim lKeyValue As Long
    Dim lRes As Long
    Dim lSize As Long
    Dim szEmpty As String
    
    szEmpty = Chr(0)
    
    
    lSize = 4
    lRes = RegCreateKey(HKEY_LOCAL_MACHINE, _
    "SOFTWARE\ODBC\ODBC.INI\" & _
    szWantedDSN, _
    hKey)
    

    If lRes <> ERROR_SUCCESS Then
        createSQLDSN = False
        Exit Function
    End If
    
    szKeyValue = szDriverName
    lRes = RegSetValueExString(hKey, "Driver", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
    
    szKeyValue = "us_english"
    lRes = RegSetValueExString(hKey, "Language", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
    
    szKeyValue = gbl_Server
    lRes = RegSetValueExString(hKey, "Server", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
    
    szKeyValue = gbl_Database
    lRes = RegSetValueExString(hKey, "Database", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
    
    szKeyValue = "sa" 'sLogIn
    lRes = RegSetValueExString(hKey, "LastUser", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
        
    
    
'    szKeyValue = sPassword '"sa"
'    lRes = RegSetValueExString(hKey, "LastPassword", 0&, REG_SZ, _
'    szKeyValue, Len(szKeyValue))
    
    lRes = RegCloseKey(hKey)
        
    If lRes <> ERROR_SUCCESS Then
        createSQLDSN = False
        Exit Function
    End If
    
    lRes = RegCreateKey(HKEY_LOCAL_MACHINE, _
    "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", _
    hKey)
    
    If lRes <> ERROR_SUCCESS Then
        createSQLDSN = False
        Exit Function
    End If
    
    szKeyValue = "SQL Server"
    lRes = RegSetValueExString(hKey, szWantedDSN, 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
    
    lRes = RegCloseKey(hKey)
    createSQLDSN = True
    
End Function

Public Function createSQLDSN_Acctg(szDriverName As String, _
    szWantedDSN As String) As Boolean
    
    Dim hKey As Long
    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String
    Dim lKeyValue As Long
    Dim lRes As Long
    Dim lSize As Long
    Dim szEmpty As String
    
    szEmpty = Chr(0)
    
    
    lSize = 4
    lRes = RegCreateKey(HKEY_LOCAL_MACHINE, _
    "SOFTWARE\ODBC\ODBC.INI\" & _
    szWantedDSN, _
    hKey)
    

    If lRes <> ERROR_SUCCESS Then
        createSQLDSN_Acctg = False
        Exit Function
    End If
    
    szKeyValue = gbl_DATABASE_Acctg
    
    lRes = RegSetValueExString(hKey, "Database", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
    
    szKeyValue = szDriverName
    lRes = RegSetValueExString(hKey, "Driver", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
    
    szKeyValue = "us_english"
    lRes = RegSetValueExString(hKey, "Language", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
    
    szKeyValue = "sa"
    lRes = RegSetValueExString(hKey, "LastUser", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
    
    szKeyValue = gbl_SERVER_Acctg
    lRes = RegSetValueExString(hKey, "Server", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
    
    lRes = RegCloseKey(hKey)
        
    If lRes <> ERROR_SUCCESS Then
        createSQLDSN_Acctg = False
        Exit Function
    End If
    
    lRes = RegCreateKey(HKEY_LOCAL_MACHINE, _
    "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", _
    hKey)
    
    If lRes <> ERROR_SUCCESS Then
        createSQLDSN_Acctg = False
        Exit Function
    End If
    
    szKeyValue = "SQL Server"
    lRes = RegSetValueExString(hKey, szWantedDSN, 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
    
    lRes = RegCloseKey(hKey)
    createSQLDSN_Acctg = True
    
End Function

'=== ACCESS DSN ====

Public Function CREATE_ACCESS_DNS()
Dim szDriverName As String
Dim szWantedDSN As String

szDriverName = String(255, Chr(32))
szWantedDSN = "POS_Data"

If Not checkAccessDriver(szDriverName) Then
    MsgBox "You must Install Access ODBC Drivers before use this program.", vbOK + vbCritical
End If

If Not (checkWantedAccessDSN(szWantedDSN)) Then

    If szDriverName = "" Then
        MsgBox "Can't find SQL ODBC driver.", vbOK + vbCritical
    Else

        If Not createAccessDSN(szDriverName, szWantedDSN) Then
            MsgBox "Can't find Access ODBC driver.", vbOK + vbCritical
        End If
    End If
End If

End Function

Public Function checkAccessDriver(ByRef szDriverName As String) As Boolean
    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String
    Dim bRes As Boolean
    
    
    bRes = False
    
    szKeyPath = "SOFTWARE\ODBC\ODBCINST.INI\Microsoft Access Driver (*.mdb)"
    szKeyName = "Driver"
    szKeyValue = String(255, Chr(32))
    


    If isSZKeyExist(szKeyPath, szKeyName, szKeyValue) Then
        szDriverName = szKeyValue
        bRes = True
    Else
        bRes = False
    End If
    
    checkAccessDriver = bRes
End Function


Public Function checkWantedAccessDSN(szWantedDSN As String) As Boolean
    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String
    Dim bRes As Boolean
    
    szKeyPath = "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources"
    szKeyName = szWantedDSN
    szKeyValue = String(255, Chr(32))
    


    If isSZKeyExist(szKeyPath, szKeyName, szKeyValue) Then
        bRes = True
    Else
        bRes = False
    End If
    
    checkWantedAccessDSN = bRes
    
End Function

Public Function createAccessDSN(szDriverName As String, _
    szWantedDSN As String) As Boolean
    
    Dim hKey As Long
    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String
    Dim lKeyValue As Long
    Dim lRes As Long
    Dim lSize As Long
    Dim szEmpty As String
    
    szEmpty = Chr(0)
    
    
    lSize = 4
    lRes = RegCreateKey(HKEY_LOCAL_MACHINE, _
    "SOFTWARE\ODBC\ODBC.INI\" & _
    szWantedDSN, _
    hKey)
    


    If lRes <> ERROR_SUCCESS Then
        createAccessDSN = False
        Exit Function
    End If
    
    lRes = RegSetValueExString(hKey, "UID", 0&, REG_SZ, _
    szEmpty, Len(szEmpty))
    
    szKeyValue = App.Path & "\Database\POS_Data.mdb"
    lRes = RegSetValueExString(hKey, "DBQ", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
    szKeyValue = szDriverName
    lRes = RegSetValueExString(hKey, "Driver", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
    szKeyValue = "MS Access;"
    lRes = RegSetValueExString(hKey, "FIL", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
    lKeyValue = 25
    lRes = RegSetValueExLong(hKey, "DriverId", 0&, REG_DWORD, _
    lKeyValue, 4)
    
    lKeyValue = 0
    lRes = RegSetValueExLong(hKey, "SafeTransactions", 0&, REG_DWORD, _
    lKeyValue, 4)
    
    lRes = RegCloseKey(hKey)
    szKeyPath = "SOFTWARE\ODBC\ODBC.INI\" & szWantedDSN & "\Engines\Jet"
    
    lRes = RegCreateKey(HKEY_LOCAL_MACHINE, _
    szKeyPath, _
    hKey)
    


    If lRes <> ERROR_SUCCESS Then
        createAccessDSN = False
        Exit Function
    End If
    lRes = RegSetValueExString(hKey, "ImplicitCommitSync", 0&, REG_SZ, _
    szEmpty, Len(szEmpty))
    szKeyValue = "Yes"
    lRes = RegSetValueExString(hKey, "UserCommitSync", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
    lKeyValue = 2048
    lRes = RegSetValueExLong(hKey, "MaxBufferSize", 0&, REG_DWORD, _
    lKeyValue, 4)
    
    lKeyValue = 5
    lRes = RegSetValueExLong(hKey, "PageTimeout", 0&, REG_DWORD, _
    lKeyValue, 4)
    
    lKeyValue = 3
    lRes = RegSetValueExLong(hKey, "Threads", 0&, REG_DWORD, _
    lKeyValue, 4)
    
    lRes = RegCloseKey(hKey)
    lRes = RegCreateKey(HKEY_LOCAL_MACHINE, _
    "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", _
    hKey)
    


    If lRes <> ERROR_SUCCESS Then
        createAccessDSN = False
        Exit Function
    End If
    
    szKeyValue = "Microsoft Access Driver (*.mdb)"
    lRes = RegSetValueExString(hKey, szWantedDSN, 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
    
    lRes = RegCloseKey(hKey)
    createAccessDSN = True
    
End Function




