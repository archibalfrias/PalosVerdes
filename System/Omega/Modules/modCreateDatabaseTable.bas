Attribute VB_Name = "modCreateDatabaseTable"

Option Explicit

Dim i, j, Arr, Arr1, ClusterName, RelationshipName
Dim s As String
Dim rs As New ADODB.Recordset
Dim t As String
Dim rt As New ADODB.Recordset

Dim DetailTableName, Columns, ColumnsDet, Clustered

Public Sub CreateNewTable(Database, TableName, Columns)
s = "SELECT * " & _
    " From sysobjects " & _
    " WHERE (name = N'" & TableName & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount = 0 Then
    t = "CREATE TABLE " & Database & ".." & TableName & " ("
    Arr = Split(CStr(Columns), "|", -1, 1)
    For i = 1 To UBound(Arr)
        Arr1 = Split(CStr(Arr(i)), ":", -1, 1)
        For j = 0 To UBound(Arr1)
            t = t & " " & Arr1(j)
        Next j
        t = t & ","
    Next i
    t = Mid(t, 1, Len(t) - 1) & ")"
    ConnOmega.Execute t
End If
rs.Close
End Sub

Public Sub CreatePrimaryKey(Database, TableName, PrimaryKeyField, _
Optional Clustered As Long = 1, Optional MultiplePK As Long = 0)
If MultiplePK = 0 Then
    t = "ALTER TABLE " & Database & ".." & TableName & " " & _
        " WITH NOCHECK ADD " & _
        " CONSTRAINT [PK_" & TableName & "] PRIMARY KEY " & IIf(Clustered = 0, "NONCLUSTERED", "CLUSTERED") & " " & _
        " ([" & PrimaryKeyField & "])"
Else
    t = "ALTER TABLE " & Database & ".." & TableName & " " & _
        " WITH NOCHECK ADD " & _
        " CONSTRAINT [PK_" & TableName & "] PRIMARY KEY " & IIf(Clustered = 0, "NONCLUSTERED", "CLUSTERED") & " ("
    Arr = Split(CStr(PrimaryKeyField), "|", -1, 1)
    For i = 0 To UBound(Arr)
        t = t & "[" & Arr(i) & "],"
    Next i
    t = Mid(t, 1, Len(t) - 1) & ")"
End If
ConnOmega.Execute t
End Sub

Public Sub CreateUniqueCluster(TableName, ClusteredName, ClusteredField, Optional MultiplePK As Long = 0)
If MultiplePK = 0 Then
    t = "CREATE UNIQUE CLUSTERED INDEX " & ClusteredName & " ON " & TableName & " " & _
        " ([" & ClusteredField & "])"
Else
    t = "CREATE UNIQUE CLUSTERED INDEX " & ClusteredName & " ON " & TableName & " ("
    Arr = Split(CStr(ClusteredField), "|", -1, 1)
    For i = 0 To UBound(Arr)
        t = t & "[" & Arr(i) & "],"
    Next i
    t = Mid(t, 1, Len(t) - 1) & ")"
End If
ConnOmega.Execute t
End Sub

Public Sub CreateCluster(TableName, ClusteredName, ClusteredField, Optional MultiplePK As Long = 0)
If MultiplePK = 0 Then
    t = "CREATE CLUSTERED INDEX " & ClusteredName & " ON " & TableName & " " & _
        " ([" & ClusteredField & "])"
Else
    t = "CREATE CLUSTERED INDEX " & ClusteredName & " ON " & TableName & " ("
    Arr = Split(CStr(ClusteredField), "|", -1, 1)
    For i = 0 To UBound(Arr)
        t = t & "[" & Arr(i) & "],"
    Next i
    t = Mid(t, 1, Len(t) - 1) & ")"
End If
ConnOmega.Execute t
End Sub

Public Sub CreateRelationship(Database, MasterTable, MasterKey, DetailTable, DetailKey, _
Optional CascadeOnDelete As Long = 0)
If CascadeOnDelete = 0 Then
    t = "ALTER TABLE " & Database & ".." & DetailTable & " ADD CONSTRAINT" & _
        " [FK_" & DetailTable & "_" & MasterTable & "] FOREIGN KEY" & _
        " ([" & DetailKey & "]) REFERENCES " & Database & ".." & MasterTable & _
        " ([" & MasterKey & "])"
Else
    t = "ALTER TABLE " & Database & ".." & DetailTable & " ADD CONSTRAINT" & _
        " [FK_" & DetailTable & "_" & MasterTable & "] FOREIGN KEY" & _
        " ([" & DetailKey & "]) REFERENCES " & Database & ".." & MasterTable & _
        " ([" & MasterKey & "]) ON DELETE CASCADE"
End If
ConnOmega.Execute t
End Sub

Public Function CreateTable(Database, TableName, Columns, _
Optional Clustered As String = "", Optional wDetail As Long = 0, _
Optional DetailTable As String = "", Optional DetailColums As String = "")

s = "SELECT * " & _
    " From sysobjects " & _
    " WHERE (name = N'" & TableName & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount > 0 Then
    If CLng(wDetail) = 1 Then
        RelationshipName = "FK_" & DetailTable & "_" & TableName & ""
        t = "SELECT * " & _
            " From sysobjects " & _
            " WHERE (name = N'" & RelationshipName & "')"
        If rt.State = adStateOpen Then rt.Close
        rt.Open t, ConnOmega
        If rt.RecordCount > 0 Then
            ConnOmega.Execute "ALTER TABLE " & Database & ".." & DetailTable & " " & _
                              " DROP CONSTRAINT " & RelationshipName
        End If
        rt.Close
    End If
    ConnOmega.Execute "DROP TABLE " & Database & ".." & TableName & ""
End If
rs.Close

s = "CREATE TABLE " & Database & ".." & TableName & " " & _
    " (PK int IDENTITY(1,1) NOT NULL, "
Arr = Split(CStr(Columns), "|", -1, 1)
For i = 1 To UBound(Arr)
    Arr1 = Split(CStr(Arr(i)), ":", -1, 1)
    For j = 0 To UBound(Arr1)
        s = s & " " & Arr1(j)
    Next j
    s = s & ","
Next i
s = Mid(s, 1, Len(s) - 1) & ")"
ConnOmega.Execute s

ConnOmega.Execute "ALTER TABLE " & Database & ".." & TableName & " " & _
                  " WITH NOCHECK ADD " & _
                  " CONSTRAINT [PK_" & TableName & "] PRIMARY KEY NONCLUSTERED " & _
                  " ([PK])"
                  
If Trim(CStr(Clustered)) <> "" Then
    ClusterName = TableName
    Arr = Split(Clustered, "|", -1, 1)
    For i = 1 To UBound(Arr)
        ClusterName = ClusterName & "_" & Arr(i)
    Next i
    s = "CREATE CLUSTERED INDEX " & ClusterName & " ON " & TableName & "("
    Arr = Split(Clustered, "|", -1, 1)
    For i = 1 To UBound(Arr)
        s = s & " " & Arr(i) & ","
    Next i
    s = Mid(s, 1, Len(s) - 1) & ")"
    ConnOmega.Execute s
End If

'== with Detailed
If CLng(wDetail) = 1 Then
    s = "SELECT * " & _
        " From sysobjects " & _
        " WHERE (name = N'" & DetailTable & "')"
    If rs.State = adStateOpen Then rs.Close
    rs.Open s, ConnOmega
    If rs.RecordCount > 0 Then
        ConnOmega.Execute "DROP TABLE " & Database & ".." & DetailTable & ""
    End If
    rs.Close
    
    s = "CREATE TABLE " & Database & ".." & DetailTable & " " & _
        " (MasterKey int NOT NULL, Line int NOT NULL, "
    Arr = Split(CStr(DetailColums), "|", -1, 1)
    For i = 1 To UBound(Arr)
        Arr1 = Split(CStr(Arr(i)), ":", -1, 1)
        For j = 0 To UBound(Arr1)
            s = s & " " & Arr1(j)
        Next j
        s = s & ","
    Next i
    s = Mid(s, 1, Len(s) - 1) & ")"
    ConnOmega.Execute s
    
    ConnOmega.Execute "ALTER TABLE " & Database & ".." & DetailTable & " " & _
                      " WITH NOCHECK ADD " & _
                      " CONSTRAINT [PK_" & DetailTable & "] PRIMARY KEY CLUSTERED " & _
                      " ([MasterKey], [Line])"
    
    ConnOmega.Execute "ALTER TABLE " & Database & ".." & DetailTable & " " & _
                      " ADD" & _
                      " CONSTRAINT [FK_" & DetailTable & "_" & TableName & "] FOREIGN KEY" & _
                      " ([MasterKey])" & _
                      " REFERENCES " & Database & ".." & TableName & _
                      " ([PK])"
    
End If
End Function

