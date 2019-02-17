Attribute VB_Name = "modCreateDropTable"
Option Explicit

Dim s As String
Dim rs As New ADODB.Recordset
Dim t As String
Dim rt As New ADODB.Recordset

Dim DetailTableName, Columns, ColumnsDet, i, iDay, Arr, sDetTmp, sComputedField



Public Sub CREATE_MODIFIED_STABLE_FORD(MasterTable)
Columns = ""
Columns = Columns & "|TeamKey:int:NOT NULL:DEFAULT(0)"
Columns = Columns & "|TeamName:varchar:(50):NOT NULL:DEFAULT('')"
Columns = Columns & "|AveHandicap:float:NOT NULL:DEFAULT(0)"
Columns = Columns & "|TeamIndex:float:NOT NULL:DEFAULT(0)"
Columns = Columns & "|TeamTotal:float:NOT NULL:DEFAULT(0)"
Columns = Columns & "|LastPlayer:float:NOT NULL:DEFAULT(0)"
Columns = Columns & "|Back9:float:NOT NULL:DEFAULT(0)"
Columns = Columns & "|Front9:float:NOT NULL:DEFAULT(0)"

DetailTableName = MasterTable & "_Detail"
ColumnsDet = ""
ColumnsDet = ColumnsDet & "|PlayerName:varchar:(100):NOT NULL:DEFAULT('')"
Arr = Split(TournamentRange, " - ", -1, 1)
iDay = 0: sDetTmp = "": sComputedField = ""
For i = 0 To DateDiff("d", Arr(0), Arr(1), vbMonday)
    ColumnsDet = ColumnsDet & "|Day" & i + 1 & ":float:NOT NULL:DEFAULT(0)"
    iDay = iDay + 1
    sDetTmp = sDetTmp & "|Day" & CStr(i + 1)
Next i

Arr = Split(sDetTmp, "|", -1, 1)
For i = 1 To UBound(Arr)
    sComputedField = sComputedField & Arr(i) & "+"
Next i
sComputedField = "(" & Mid(sComputedField, 1, Len(sComputedField) - 1) & ")"

ColumnsDet = ColumnsDet & "|Total: AS " & sComputedField & ""

CreateTable gbl_Database, MasterTable, Columns, "", 1, CStr(DetailTableName), CStr(ColumnsDet)
End Sub


Public Sub CREATE_CV_CHECK(TableName)
s = "SELECT * " & _
    " FROM sysobjects " & _
    " WHERE (name = N'" & CStr(TableName) & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount = 0 Then
    Columns = ""
    Columns = Columns & "|PK:int:IDENTITY(1,1):NOT NULL"
    Columns = Columns & "|LogInName:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|CheckDate:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|CheckAmt:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|PayTo:varchar:(100):NOT NULL:DEFAULT('')"
    Columns = Columns & "|Pesos:varchar:(255):NOT NULL:DEFAULT('')"
    
    CreateNewTable gbl_Database, TableName, Columns
    CreatePrimaryKey gbl_Database, TableName, "PK", 0
    CreateUniqueCluster TableName, TableName & "_LogInName", "LogInName"
End If
If rs.State = adStateOpen Then rs.Close
End Sub

Public Sub CREATE_CV_TABLES(TableName)
s = "SELECT * " & _
    " FROM sysobjects " & _
    " WHERE (name = N'" & CStr(TableName) & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount = 0 Then
    Columns = ""
    Columns = Columns & "|PK:int:IDENTITY(1,1):NOT NULL"
    Columns = Columns & "|LogInName:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|CVNumber:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|CVDate:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|PayTo:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|Pesos:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|ORNumber:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|TotAmt:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|ChkNumber:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|Approved:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|Received:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|Prepared:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|Checked:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|Entered:varchar:(50):NOT NULL:DEFAULT('')"
    
    
    CreateNewTable gbl_Database, TableName, Columns
    CreatePrimaryKey gbl_Database, TableName, "PK", 0
    CreateUniqueCluster TableName, TableName & "_LogInName", "LogInName"

End If
If rs.State = adStateOpen Then rs.Close

DetailTableName = TableName & "_Explanation"

t = "SELECT * " & _
    " FROM sysobjects " & _
    " WHERE (name = N'" & CStr(DetailTableName) & "')"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount = 0 Then
    ColumnsDet = ""
    ColumnsDet = ColumnsDet & "|MasterKey:int:NOT NULL"
    ColumnsDet = ColumnsDet & "|Line:int:NOT NULL"
    ColumnsDet = ColumnsDet & "|Description:varchar:(100):NOT NULL:DEFAULT('')"
    ColumnsDet = ColumnsDet & "|Amount:varchar:(50):NOT NULL:DEFAULT('')"
    
    CreateNewTable gbl_Database, DetailTableName, ColumnsDet
    CreatePrimaryKey gbl_Database, DetailTableName, "MasterKey|Line", 1, 1
    CreateRelationship gbl_Database, TableName, "PK", DetailTableName, "MasterKey", 1
    
End If
rt.Close

DetailTableName = TableName & "_AD"

t = "SELECT * " & _
    " FROM sysobjects " & _
    " WHERE (name = N'" & CStr(DetailTableName) & "')"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount = 0 Then
    ColumnsDet = ""
    ColumnsDet = ColumnsDet & "|MasterKey:int:NOT NULL"
    ColumnsDet = ColumnsDet & "|Line:int:NOT NULL"
    ColumnsDet = ColumnsDet & "|AccountCode:varchar:(50):NOT NULL:DEFAULT('')"
    ColumnsDet = ColumnsDet & "|AccountName:varchar:(100):NOT NULL:DEFAULT('')"
    ColumnsDet = ColumnsDet & "|Debit:varchar:(50):NOT NULL:DEFAULT('')"
    ColumnsDet = ColumnsDet & "|Credit:varchar:(50):NOT NULL:DEFAULT('')"
    
    CreateNewTable gbl_Database, DetailTableName, ColumnsDet
    CreatePrimaryKey gbl_Database, DetailTableName, "MasterKey|Line", 1, 1
    CreateRelationship gbl_Database, TableName, "PK", DetailTableName, "MasterKey", 1
    
End If
rt.Close
End Sub


Public Sub CREATE_PROFILE_DATASHEET_TABLE(TableName)
s = "SELECT * " & _
    " FROM sysobjects " & _
    " WHERE (name = N'" & CStr(TableName) & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount = 0 Then
    
    Columns = ""
    Columns = Columns & "|PK:int:IDENTITY(1,1):NOT NULL"
    Columns = Columns & "|LogInName:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|Positions:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|DateHired:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|Department:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|Levels:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|Name:varchar:(255):NOT NULL:DEFAULT('')"
    Columns = Columns & "|PresentAddress:varchar:(255):NOT NULL:DEFAULT('')"
    Columns = Columns & "|OwnedHouse:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|Rent:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|BirthDate:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|Age:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|BirthPlace:varchar:(255):NOT NULL:DEFAULT('')"
    Columns = Columns & "|Religion:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|LivingParents:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|CivilStatus:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|DateMarriage:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|Height:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|Weight:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|Nationality:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|SSSNumber:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|TIN:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|DriversLicense:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|PHICNumber:varchar:(255):NOT NULL:DEFAULT('')"
    Columns = Columns & "|PagIbigNumber:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|IDNumber:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|SpouseName:varchar:(100):NOT NULL:DEFAULT('')"
    Columns = Columns & "|SpouseBirthDate:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|SpouseOccupation:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|SpouseAddress:varchar:(255):NOT NULL:DEFAULT('')"
    Columns = Columns & "|FatherName:varchar:(100):NOT NULL:DEFAULT('')"
    Columns = Columns & "|FatherBirthDate:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|FatherOccupation:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|FatherAddress:varchar:(255):NOT NULL:DEFAULT('')"
    Columns = Columns & "|MotherName:varchar:(100):NOT NULL:DEFAULT('')"
    Columns = Columns & "|MotherBirthDate:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|MotherOccupation:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|MotherAddress:varchar:(255):NOT NULL:DEFAULT('')"
    Columns = Columns & "|Skills:varchar:(255):NOT NULL:DEFAULT('')"
    Columns = Columns & "|OrgClub:varchar:(255):NOT NULL:DEFAULT('')"
    Columns = Columns & "|RelatedName:varchar:(255):NOT NULL:DEFAULT('')"
    Columns = Columns & "|RelatedContact:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|RelatedAddress:varchar:(255):NOT NULL:DEFAULT('')"
    Columns = Columns & "|RelativeName:varchar:(255):NOT NULL:DEFAULT('')"
    Columns = Columns & "|RelativeContact:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|RelativeAddress:varchar:(255):NOT NULL:DEFAULT('')"
    Columns = Columns & "|EmergencyName:varchar:(255):NOT NULL:DEFAULT('')"
    Columns = Columns & "|EmergencyAddress:varchar:(255):NOT NULL:DEFAULT('')"
    Columns = Columns & "|EmergencyRelation:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|EmergencyContact:varchar:(50):NOT NULL:DEFAULT('')"
    Columns = Columns & "|Sibling:float:NOT NULL:DEFAULT(0)"
    Columns = Columns & "|Children:float:NOT NULL:DEFAULT(0)"
    Columns = Columns & "|Education:float:NOT NULL:DEFAULT(0)"
    Columns = Columns & "|Employment:float:NOT NULL:DEFAULT(0)"
    Columns = Columns & "|Picture:image:NULL"
        
    CreateNewTable gbl_Database, TableName, Columns
    CreatePrimaryKey gbl_Database, TableName, "PK", 0
    CreateUniqueCluster TableName, TableName & "_LogInName", "LogInName"
                      
End If
If rs.State = adStateOpen Then rs.Close

DetailTableName = TableName & "_Children"

t = "SELECT * " & _
    " FROM sysobjects " & _
    " WHERE (name = N'" & CStr(DetailTableName) & "')"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount = 0 Then
    ColumnsDet = ""
    ColumnsDet = ColumnsDet & "|MasterKey:int:NOT NULL"
    ColumnsDet = ColumnsDet & "|Line:int:NOT NULL"
    ColumnsDet = ColumnsDet & "|FullName:varchar:(100):NOT NULL:DEFAULT('')"
    ColumnsDet = ColumnsDet & "|BirthDate:varchar:(50):NOT NULL:DEFAULT('')"
    ColumnsDet = ColumnsDet & "|Occupation:varchar:(50):NOT NULL:DEFAULT('')"
    ColumnsDet = ColumnsDet & "|Address:varchar:(255):NOT NULL:DEFAULT('')"
    
    CreateNewTable gbl_Database, DetailTableName, ColumnsDet
    CreatePrimaryKey gbl_Database, DetailTableName, "MasterKey|Line", 1, 1
    CreateRelationship gbl_Database, TableName, "PK", DetailTableName, "MasterKey", 1
    
End If
rt.Close

DetailTableName = TableName & "_Sibling"

t = "SELECT * " & _
    " FROM sysobjects " & _
    " WHERE (name = N'" & CStr(DetailTableName) & "')"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount = 0 Then
    ColumnsDet = ""
    ColumnsDet = ColumnsDet & "|MasterKey:int:NOT NULL"
    ColumnsDet = ColumnsDet & "|Line:int:NOT NULL"
    ColumnsDet = ColumnsDet & "|FullName:varchar:(100):NOT NULL:DEFAULT('')"
    ColumnsDet = ColumnsDet & "|BirthDate:varchar:(50):NOT NULL:DEFAULT('')"
    ColumnsDet = ColumnsDet & "|Occupation:varchar:(50):NOT NULL:DEFAULT('')"
    ColumnsDet = ColumnsDet & "|Address:varchar:(255):NOT NULL:DEFAULT('')"
    
    CreateNewTable gbl_Database, DetailTableName, ColumnsDet
    CreatePrimaryKey gbl_Database, DetailTableName, "MasterKey|Line", 1, 1
    CreateRelationship gbl_Database, TableName, "PK", DetailTableName, "MasterKey", 1
    
End If
rt.Close


DetailTableName = TableName & "_Employment"

t = "SELECT * " & _
    " FROM sysobjects " & _
    " WHERE (name = N'" & CStr(DetailTableName) & "')"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount = 0 Then
    ColumnsDet = ""
    ColumnsDet = ColumnsDet & "|MasterKey:int:NOT NULL"
    ColumnsDet = ColumnsDet & "|Line:int:NOT NULL"
    ColumnsDet = ColumnsDet & "|Company:varchar:(100):NOT NULL:DEFAULT('')"
    ColumnsDet = ColumnsDet & "|Positions:varchar:(50):NOT NULL:DEFAULT('')"
    ColumnsDet = ColumnsDet & "|Salary:varchar:(50):NOT NULL:DEFAULT('')"
    ColumnsDet = ColumnsDet & "|IncDates:varchar:(50):NOT NULL:DEFAULT('')"
    ColumnsDet = ColumnsDet & "|Address:varchar:(255):NOT NULL:DEFAULT('')"
    
    CreateNewTable gbl_Database, DetailTableName, ColumnsDet
    CreatePrimaryKey gbl_Database, DetailTableName, "MasterKey|Line", 1, 1
    CreateRelationship gbl_Database, TableName, "PK", DetailTableName, "MasterKey", 1
    
End If
rt.Close

DetailTableName = TableName & "_Education"

t = "SELECT * " & _
    " FROM sysobjects " & _
    " WHERE (name = N'" & CStr(DetailTableName) & "')"
If rt.State = adStateOpen Then rt.Close
rt.Open t, ConnOmega
If rt.RecordCount = 0 Then
    ColumnsDet = ""
    ColumnsDet = ColumnsDet & "|MasterKey:int:NOT NULL"
    ColumnsDet = ColumnsDet & "|Line:int:NOT NULL"
    ColumnsDet = ColumnsDet & "|SchoolName:varchar:(100):NOT NULL:DEFAULT('')"
    ColumnsDet = ColumnsDet & "|InclusiveDate:varchar:(50):NOT NULL:DEFAULT('')"
    ColumnsDet = ColumnsDet & "|Course:varchar:(50):NOT NULL:DEFAULT('')"
    ColumnsDet = ColumnsDet & "|Address:varchar:(255):NOT NULL:DEFAULT('')"
    
    CreateNewTable gbl_Database, DetailTableName, ColumnsDet
    CreatePrimaryKey gbl_Database, DetailTableName, "MasterKey|Line", 1, 1
    CreateRelationship gbl_Database, TableName, "PK", DetailTableName, "MasterKey", 1
    
End If
rt.Close

End Sub
