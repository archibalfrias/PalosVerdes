Attribute VB_Name = "modPictures"
Option Explicit

Dim Filename, strPath

Dim s           As String
Dim rs          As New ADODB.Recordset
Dim myStream    As New ADODB.Stream
Dim cnnS        As String
Dim cnn         As ADODB.Connection

Public Sub SAVE_IMAGES(iKey, iLine, strPath As String, isTable As String)
If Trim(strPath) = "" Then Exit Sub
Set cnn = New ADODB.Connection
cnnS = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" + sLogIn + ";Password=" + EncryptDecryptLogIn(sPassword) + ";Initial Catalog=" + gbl_Database + ";Data Source=" + gbl_Server
cnn.CursorLocation = adUseClient
cnn.Mode = adModeReadWrite
cnn.IsolationLevel = adXactIsolated
If cnn.State = adStateOpen Then cnn.Close
cnn.Open cnnS
myStream.Type = adTypeBinary
Select Case isTable
    Case "Company Logo"
        s = "SELECT tbl_Company.* " & _
            " FROM tbl_Company "
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, cnn, adOpenForwardOnly, adLockOptimistic
        If rs.RecordCount > 0 Then
            myStream.Open
            myStream.LoadFromFile strPath
            rs.Fields("Logo").Value = myStream.Read
            rs.Update
            myStream.Close
        End If
        rs.Close
    Case "Wallpaper"
        s = "SELECT tbl_Wallpaper.* " & _
            " FROM tbl_Wallpaper " & _
            " WHERE (PK = " & iKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, cnn, adOpenForwardOnly, adLockOptimistic
        If rs.RecordCount > 0 Then
            myStream.Open
            myStream.LoadFromFile strPath
            rs.Fields("WallPaper").Value = myStream.Read
            rs.Update
            myStream.Close
        End If
        rs.Close
    Case "Employee Profile"
        s = "SELECT tbl_Personnel_Information.* " & _
            " FROM tbl_Personnel_Information " & _
            " WHERE (PK = " & iKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, cnn, adOpenForwardOnly, adLockOptimistic
        If rs.RecordCount > 0 Then
            myStream.Open
            myStream.LoadFromFile strPath
            rs.Fields("Picture").Value = myStream.Read
            rs.Update
            myStream.Close
        End If
        rs.Close
    Case "Employee DataSheet"
        s = "SELECT tbl_Personnel_DataSheet.* " & _
            " FROM tbl_Personnel_DataSheet " & _
            " WHERE (PK = " & iKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, cnn, adOpenForwardOnly, adLockOptimistic
        If rs.RecordCount > 0 Then
            myStream.Open
            myStream.LoadFromFile strPath
            rs.Fields("Picture").Value = myStream.Read
            rs.Update
            myStream.Close
        End If
        rs.Close
    Case "Member"
        s = "SELECT tbl_Member_Information.* " & _
            " FROM tbl_Member_Information " & _
            " WHERE (PK = " & iKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, cnn, adOpenForwardOnly, adLockOptimistic
        If rs.RecordCount > 0 Then
            myStream.Open
            myStream.LoadFromFile strPath
            rs.Fields("MemberPicture").Value = myStream.Read
            rs.Update
            myStream.Close
        End If
        rs.Close
    Case "Member Spouse"
        s = "SELECT tbl_Member_Information.* " & _
            " FROM tbl_Member_Information " & _
            " WHERE (PK = " & iKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, cnn, adOpenForwardOnly, adLockOptimistic
        If rs.RecordCount > 0 Then
            myStream.Open
            myStream.LoadFromFile strPath
            rs.Fields("SpousePicture").Value = myStream.Read
            rs.Update
            myStream.Close
        End If
        rs.Close
    Case "Member Child"
        s = "SELECT tbl_Member_Dependent.* " & _
            " FROM tbl_Member_Dependent " & _
            " WHERE (MemberKey = " & iKey & ") " & _
            " AND (Line = " & iLine & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, cnn, adOpenForwardOnly, adLockOptimistic
        If rs.RecordCount > 0 Then
            myStream.Open
            myStream.LoadFromFile strPath
            rs.Fields("ChildPicture").Value = myStream.Read
            rs.Update
            myStream.Close
        End If
        rs.Close
    Case "Member ID Number"
        s = "SELECT tbl_Member_IDNumber.* " & _
            " FROM tbl_Member_IDNumber " & _
            " WHERE (PK = " & iKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, cnn, adOpenForwardOnly, adLockOptimistic
        If rs.RecordCount > 0 Then
            myStream.Open
            myStream.LoadFromFile strPath
            rs.Fields("MemberPicture").Value = myStream.Read
            rs.Update
            myStream.Close
        End If
        rs.Close
    Case "Member ID Number (Child)"
        s = "SELECT tbl_Member_IDNumber.* " & _
            " FROM tbl_Member_IDNumber " & _
            " WHERE (PK = " & iKey & ") " & _
            " AND (MemberType = 3) " & _
            " AND (MemberChildLine = " & iLine & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, cnn, adOpenForwardOnly, adLockOptimistic
        If rs.RecordCount > 0 Then
            myStream.Open
            myStream.LoadFromFile strPath
            rs.Fields("MemberPicture").Value = myStream.Read
            rs.Update
            myStream.Close
        End If
        rs.Close
    Case "Menu Management"
        s = "SELECT tbl_Menu.* " & _
            " FROM tbl_Menu " & _
            " WHERE (PK = " & iKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, cnn, adOpenForwardOnly, adLockOptimistic
        If rs.RecordCount > 0 Then
            myStream.Open
            myStream.LoadFromFile strPath
            rs.Fields("Picture").Value = myStream.Read
            rs.Update
            myStream.Close
        End If
        rs.Close
    Case "Caddy Information"
        s = "SELECT tbl_Caddy_Information.* " & _
            " FROM tbl_Caddy_Information " & _
            " WHERE (PK = " & iKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, cnn, adOpenForwardOnly, adLockOptimistic
        If rs.RecordCount > 0 Then
            myStream.Open
            myStream.LoadFromFile strPath
            rs.Fields("Picture").Value = myStream.Read
            rs.Update
            myStream.Close
        End If
        rs.Close
    
End Select
If cnn.State = adStateOpen Then cnn.Close
End Sub

Public Function SHOW_IMAGES(iKey, iLine, isTable As String) As String
Set cnn = New ADODB.Connection
cnnS = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" + sLogIn + ";Password=" + EncryptDecryptLogIn(sPassword) + ";Initial Catalog=" + gbl_Database + ";Data Source=" + gbl_Server
cnn.CursorLocation = adUseClient
cnn.Mode = adModeReadWrite
cnn.IsolationLevel = adXactIsolated
If cnn.State = adStateOpen Then cnn.Close
cnn.Open cnnS
Select Case isTable
    Case "Company Logo"
        strPath = App.Path & "\Tmp"
        On Error Resume Next
        MkDir strPath
        Filename = strPath & "\CompanyLogo"
        s = "SELECT tbl_Company.* " & _
            " FROM tbl_Company " & _
            " WHERE (PK = " & iKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, cnn, adOpenDynamic, adLockOptimistic
        If rs.RecordCount > 0 Then
            If IsNull(rs!Logo) = True Then
                SHOW_IMAGES = ""
            Else
                myStream.Type = adTypeBinary
                myStream.Open
                myStream.Write rs.Fields("Logo").Value
                myStream.SaveToFile Filename, adSaveCreateOverWrite
                myStream.Close
                SHOW_IMAGES = Filename
            End If
        End If
        rs.Close
    Case "Wallpaper"
        strPath = App.Path & "\Tmp"
        On Error Resume Next
        MkDir strPath
        Filename = strPath & "\Wallpaper"
        s = "SELECT tbl_Wallpaper.* " & _
            " FROM tbl_Wallpaper " & _
            " WHERE (PK = " & iKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, cnn, adOpenDynamic, adLockOptimistic
        If rs.RecordCount > 0 Then
            If IsNull(rs!Wallpaper) = True Then
                SHOW_IMAGES = ""
            Else
                myStream.Type = adTypeBinary
                myStream.Open
                myStream.Write rs.Fields("WallPaper").Value
                myStream.SaveToFile Filename, adSaveCreateOverWrite
                myStream.Close
                SHOW_IMAGES = Filename
            End If
        End If
        rs.Close
    Case "Background"
        strPath = App.Path & "\Tmp\Back"
        'On Error Resume Next
        'MkDir strPath
        Filename = strPath & "\" & iKey & ".jpg"
        s = "SELECT tbl_Wallpaper.* " & _
            " FROM tbl_Wallpaper " & _
            " WHERE (PK = " & iKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, cnn, adOpenDynamic, adLockOptimistic
        If rs.RecordCount > 0 Then
            If IsNull(rs!Wallpaper) = True Then
                SHOW_IMAGES = ""
            Else
                myStream.Type = adTypeBinary
                myStream.Open
                myStream.Write rs.Fields("WallPaper").Value
                myStream.SaveToFile Filename, adSaveCreateOverWrite
                myStream.Close
                SHOW_IMAGES = Filename
            End If
        End If
        rs.Close
    Case "Employee Profile"
        strPath = App.Path & "\Tmp"
        On Error Resume Next
        MkDir strPath
        Filename = strPath & "\Employee"
        s = "SELECT Picture " & _
            " FROM tbl_Personnel_Information " & _
            " WHERE (PK = " & iKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, cnn, adOpenDynamic, adLockOptimistic
        If rs.RecordCount > 0 Then
            If IsNull(rs!Picture) = True Then
                SHOW_IMAGES = ""
            Else
                myStream.Type = adTypeBinary
                myStream.Open
                myStream.Write rs.Fields("Picture").Value
                myStream.SaveToFile Filename, adSaveCreateOverWrite
                myStream.Close
                SHOW_IMAGES = Filename
            End If
        End If
        rs.Close
    Case "Member"
        strPath = App.Path & "\Tmp"
        On Error Resume Next
        MkDir strPath
        Filename = strPath & "\Member"
        s = "SELECT MemberPicture " & _
            " FROM tbl_Member_Information " & _
            " WHERE (PK = " & iKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, cnn, adOpenDynamic, adLockOptimistic
        If rs.RecordCount > 0 Then
            If IsNull(rs!MemberPicture) = True Then
                SHOW_IMAGES = ""
            Else
                myStream.Type = adTypeBinary
                myStream.Open
                myStream.Write rs.Fields("MemberPicture").Value
                myStream.SaveToFile Filename, adSaveCreateOverWrite
                myStream.Close
                SHOW_IMAGES = Filename
            End If
        End If
        rs.Close
    Case "Member Spouse"
        strPath = App.Path & "\Tmp"
        On Error Resume Next
        MkDir strPath
        Filename = strPath & "\Member_Spouse"
        s = "SELECT SpousePicture " & _
            " FROM tbl_Member_Information " & _
            " WHERE (PK = " & iKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, cnn, adOpenDynamic, adLockOptimistic
        If rs.RecordCount > 0 Then
            If IsNull(rs!SpousePicture) = True Then
                SHOW_IMAGES = ""
            Else
                myStream.Type = adTypeBinary
                myStream.Open
                myStream.Write rs.Fields("SpousePicture").Value
                myStream.SaveToFile Filename, adSaveCreateOverWrite
                myStream.Close
                SHOW_IMAGES = Filename
            End If
        End If
        rs.Close
    Case "Member Child"
        strPath = App.Path & "\Tmp"
        On Error Resume Next
        MkDir strPath
        Filename = strPath & "\Member_Child_" & iLine & ".jpg"
        s = "SELECT ChildPicture " & _
            " FROM tbl_Member_Dependent " & _
            " WHERE (MemberKey = " & iKey & ") " & _
            " AND (Line = " & iLine & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, cnn, adOpenDynamic, adLockOptimistic
        If rs.RecordCount > 0 Then
            If IsNull(rs!ChildPicture) = True Then
                SHOW_IMAGES = ""
            Else
                myStream.Type = adTypeBinary
                myStream.Open
                myStream.Write rs.Fields("ChildPicture").Value
                myStream.SaveToFile Filename, adSaveCreateOverWrite
                myStream.Close
                SHOW_IMAGES = Filename
            End If
        End If
        rs.Close
    Case "Member ID Number"
        strPath = App.Path & "\Tmp"
        On Error Resume Next
        MkDir strPath
        Filename = strPath & "\MemberID"
        s = "SELECT MemberPicture " & _
            " FROM tbl_Member_IDNumber " & _
            " WHERE (PK = " & iKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, cnn, adOpenDynamic, adLockOptimistic
        If rs.RecordCount > 0 Then
            If IsNull(rs!MemberPicture) = True Then
                SHOW_IMAGES = ""
            Else
                myStream.Type = adTypeBinary
                myStream.Open
                myStream.Write rs.Fields("MemberPicture").Value
                myStream.SaveToFile Filename, adSaveCreateOverWrite
                myStream.Close
                SHOW_IMAGES = Filename
            End If
        End If
        rs.Close
    Case "Menu Management"
        strPath = App.Path & "\Tmp"
        On Error Resume Next
        MkDir strPath
        Filename = strPath & "\Menu"
        s = "SELECT Picture " & _
            " FROM tbl_Menu " & _
            " WHERE (PK = " & iKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, cnn, adOpenDynamic, adLockOptimistic
        If rs.RecordCount > 0 Then
            If IsNull(rs!Picture) = True Then
                SHOW_IMAGES = ""
            Else
                myStream.Type = adTypeBinary
                myStream.Open
                myStream.Write rs.Fields("Picture").Value
                myStream.SaveToFile Filename, adSaveCreateOverWrite
                myStream.Close
                SHOW_IMAGES = Filename
            End If
        End If
        rs.Close
    Case "Caddy Information"
        strPath = App.Path & "\Tmp"
        On Error Resume Next
        MkDir strPath
        Filename = strPath & "\Caddy"
        s = "SELECT Picture " & _
            " FROM tbl_Caddy_Information " & _
            " WHERE (PK = " & iKey & ")"
        If rs.State = adStateOpen Then rs.Close
        rs.Open s, cnn, adOpenDynamic, adLockOptimistic
        If rs.RecordCount > 0 Then
            If IsNull(rs!Picture) = True Then
                SHOW_IMAGES = ""
            Else
                myStream.Type = adTypeBinary
                myStream.Open
                myStream.Write rs.Fields("Picture").Value
                myStream.SaveToFile Filename, adSaveCreateOverWrite
                myStream.Close
                SHOW_IMAGES = Filename
            End If
        End If
        rs.Close
End Select
If cnn.State = adStateOpen Then cnn.Close
End Function



