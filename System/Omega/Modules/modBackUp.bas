Attribute VB_Name = "modBackUp"
Option Explicit


Public Sub Create_Backup(sPath As String, sFileName As String)
Dim conn As ADODB.Connection
Dim strSQL As String
Dim strBackPath
strBackPath = sPath & sFileName '"D:\Backup\" & Format(Date, "mmddyy") & "F"
On Error GoTo errHandler
Set conn = New ADODB.Connection
conn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & EncryptDecryptLogIn(sPassword) & ";Persist Security Info=True;User ID=" & sLogIn & ";Initial Catalog= " & gbl_Database & ";Data Source=" & gbl_Server
conn.Open
conn.BeginTrans
conn.CommandTimeout = 0
DoEvents
strSQL = "BACKUP DATABASE [" & gbl_Database & "] TO DISK = '" & strBackPath & "' WITH INIT"
conn.Execute strSQL
conn.CommitTrans
conn.Close
Set conn = Nothing
MsgBox "D O N E !                   ", vbInformation, "Success"
Exit Sub
errHandler:
If Err.Number <> 32755 Then
    MsgBox "Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, "CreateBackup"
End If
End Sub


