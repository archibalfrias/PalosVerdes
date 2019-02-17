Attribute VB_Name = "modConnection"

Public Sub Main()

'SaveSetting App.EXEName, "MainServer", "MServer", ""
'SaveSetting App.EXEName, "MainDatabase", "MDatabase", ""
'MsgBox CDbl(GetSetting(App.EXEName, "ConnectionAttempt", "ConnectAttempt", "0"))

If CDbl(GetSetting(App.EXEName, "ConnectionAttempt", "ConnectAttempt", "0")) >= 3 Then
    SaveSetting App.EXEName, "MainServer", "MServer", ""
    SaveSetting App.EXEName, "MainDatabase", "MDatabase", ""
    SaveSetting App.EXEName, "MainLogIn", "MLogIn", ""
    SaveSetting App.EXEName, "MainPassword", "MPassword", ""
    SaveSetting App.EXEName, "ConnectionAttempt", "ConnectAttempt", "0"
End If

'SaveSetting App.EXEName, "MainServer", "MServer", ""

'SaveSetting App.EXEName, "MainServer", "MServer", "Archie_lt"
'SaveSetting App.EXEName, "MainDatabase", "MDatabase", "Omega_Final"
'SaveSetting App.EXEName, "MainLogIn", "MLogIn", "sa"
'SaveSetting App.EXEName, "MainPassword", "MPassword", "sa123"

'SaveSetting App.EXEName, "MainServer", "MServer", "Archie"
'SaveSetting App.EXEName, "MainDatabase", "MDatabase", "Omega_Final"
'SaveSetting App.EXEName, "MainLogIn", "MLogIn", "sa"
'SaveSetting App.EXEName, "MainPassword", "MPassword", "sa123"

'SaveSetting App.EXEName, "MainServer", "MServer", "Server"
'SaveSetting App.EXEName, "MainDatabase", "MDatabase", "Omega_Final"
'SaveSetting App.EXEName, "MainLogIn", "MLogIn", "Archie"
'SaveSetting App.EXEName, "MainPassword", "MPassword", "√·ÛÙÈÏÏÔÓ±π∑π"

gbl_Server = GetSetting(App.EXEName, "MainServer", "MServer", "")
gbl_Database = GetSetting(App.EXEName, "MainDatabase", "MDatabase", "")
sLogIn = GetSetting(App.EXEName, "MainLogIn", "MLogIn", "")
sPassword = EncryptDecryptLogIn(GetSetting(App.EXEName, "MainPassword", "MPassword", ""))

gbl_ServerL = GetSetting(App.EXEName, "MainServerL", "MServerL", "")
gbl_DatabaseL = GetSetting(App.EXEName, "MainDatabaseL", "MDatabaseL", "")
sLogInL = GetSetting(App.EXEName, "MainLogInL", "MLogInL", "")
sPasswordL = EncryptDecryptLogIn(GetSetting(App.EXEName, "MainPasswordL", "MPasswordL", ""))

PassStartWizard = 0

gbl_Form_Caption = ""
sDefaultPW = "123456"
gbl_UserName = ""
gbl_Password = ""
gbl_CompleteName = ""
gbl_LockWhenIdle = 0
gbl_Idle_Time = 0
gbl_Slides_Background = 1
gbl_Slides_Time = 720
gbl_Quotes_Time = 360
gbl_Item_Module = ""

gbl_MpnthlyDivisor = 13.08333

On Error Resume Next
MkDir App.Path & "\Tmp"
MkDir App.Path & "\Tmp\Back"

If Trim(gbl_Server) = "" Then PassStartWizard = 1: frmConnectionWizard.Show 1: Exit Sub
If Trim(gbl_Database) = "" Then PassStartWizard = 1: frmConnectionWizard.Show 1: Exit Sub
If Trim(sLogIn) = "" Then PassStartWizard = 1: frmConnectionWizard.Show 1: Exit Sub
If Trim(sPassword) = "" Then PassStartWizard = 1: frmConnectionWizard.Show 1: Exit Sub

If PingServer(gbl_Server) = False Then
    MsgBox "Server is Offline!                    ", vbCritical, "Error..."
    SaveSetting App.EXEName, "MainServer", "MServer", ""
    SaveSetting App.EXEName, "MainDatabase", "MDatabase", ""
    SaveSetting App.EXEName, "MainLogIn", "MLogIn", ""
    SaveSetting App.EXEName, "MainPassword", "MPassword", ""
    End
End If

frmSplash.Show 1

End Sub

Public Function DataOpen(oConn As Connection) As Boolean
On Error GoTo open_EH:
    oConn.CursorLocation = adUseClient
    oConn.CommandTimeout = 0
    'oConn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" + gbl_Database + ";Data Source=" + gbl_Server
    oConn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" + sLogIn + ";Password=" + EncryptDecryptLogIn(sPassword) + ";Initial Catalog=" + gbl_Database + ";Data Source=" + gbl_Server
    oConn.Open oConn.ConnectionString
    oConn.Mode = adModeReadWrite
    oConn.IsolationLevel = adXactIsolated
    DataOpen = True
    
'    SaveSetting App.EXEName, "MainServerL", "MServerL", CStr(gbl_Server)
'    SaveSetting App.EXEName, "MainDatabaseL", "MDatabaseL", CStr(gbl_Database)
'    SaveSetting App.EXEName, "MainLogInL", "MLogInL", CStr(sLogIn)
'    SaveSetting App.EXEName, "MainPasswordL", "MPasswordL", CStr(sPassword)
'
'    SaveSetting App.EXEName, "ConnectionAttempt", "ConnectAttempt", "0"
    
Exit Function
open_EH:
    If Err.Number = 3705 Then Exit Function
    Call ErrorHandler(ConnOmega)
    DataOpen = False
    Exit Function
End Function

Public Function ErrorHandler(oConn As Connection)
Dim oErr As Error
Dim strmsg As String
For Each oErr In oConn.Errors
    strmsg = strmsg & _
             "Error #: " & _
             oErr.Number & vbCrLf
    strmsg = strmsg & _
             "Description: " & _
             oErr.Description & vbCrLf
    strmsg = strmsg & _
             "Source: " & _
             oErr.Source & vbCrLf
    strmsg = strmsg & _
             "SQL State: " & _
             oErr.SQLState & vbCrLf
    strmsg = strmsg & _
             "Native Error: " & _
             oErr.NativeError & vbCrLf
    strmsg = strmsg & vbCrLf
Next

MsgBox strmsg, vbCritical, "Error connection!"

'SaveSetting App.EXEName, "MainServer", "MServer", ""
'SaveSetting App.EXEName, "MainDatabase", "MDatabase", ""

DELETE_DNS_SQL_ODBC CStr(gbl_DatabaseL), CStr(gbl_ServerL), CStr(gbl_DatabaseL), CStr(sLogInL), CStr(sPasswordL)

SaveSetting App.EXEName, "ConnectionAttempt", "ConnectAttempt", CDbl(GetSetting(App.EXEName, "ConnectionAttempt", "ConnectAttempt", "0")) + 1

End
End Function

