Attribute VB_Name = "modConnection"


'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=F:\MY FILES\PROGRAMS\RPVGCC\System\090113\Remote Scoring\Database\Data.accdb;Persist Security Info=False


Public ConnRS As New ADODB.Connection


Public Function DataOpen(oConn As Connection) As Boolean
On Error GoTo open_EH
    oConn.CursorLocation = adUseClient
    oConn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & App.Path & "\Data.mdb" & ";Persist Security Info=False"
    'oConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Data.mdb" & ";Persist Security Info=False"
    'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\My_Projects\ICAMS_POS\POS VPLUS Current\OffLine\Exe File\Database\POS_Data.mdb;Persist Security Info=False
    oConn.Open oConn.ConnectionString
    DataOpen = True
Exit Function
open_EH:
    If Err.Number = 3705 Then Exit Function
    DataOpen = False
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error..."
    End
End Function

