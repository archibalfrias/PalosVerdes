Attribute VB_Name = "modAccessRights"
Public gbl_UserName             As String
Public gbl_Password             As String
Public gbl_CompleteName         As String
Public UserLevel                As Long
Public sDefaultPW               As String

Public gbl_Last_User            As String
Public SystemSetting            As String

Public gbl_MODULE               As String
Public gbl_MODULE_Action        As String

Public Function AccessRights(strModule, strAction) As Boolean
Dim ArrAccess
Dim s As String
Dim rs As New ADODB.Recordset
s = "SELECT tbl_Users_Account.*" & _
    " From tbl_Users_Account " & _
    " WHERE (UserName = '" & gbl_UserName & "') " & _
    " AND (Password = '" & EncryptDecrypt(gbl_Password) & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open s, ConnOmega
If rs.RecordCount = 0 Then AccessRights = False: Exit Function
Select Case strModule
    Case "User's Account"
        ArrAccess = Split(rs!UserAccount, "/", -1, 1)
        Select Case strAction
            Case "Open":    AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":     AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":    AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":  AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Admin":   AccessRights = IIf(rs!Admin = 1, True, False)
            Case Else:      AccessRights = False
        End Select
    Case "Company Information"
        ArrAccess = Split(rs!CompanyInfo, "/", -1, 1)
        Select Case strAction
            Case "Open":    AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Edit":    AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case Else:      AccessRights = False
        End Select
    Case "Personnel Information"
        ArrAccess = Split(rs!PersonnelInfo, "/", -1, 1)
        Select Case strAction
            Case "Open":    AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":     AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":    AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":  AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:      AccessRights = False
        End Select
    Case "Personnel ID Number"
        ArrAccess = Split(rs!PersonnelID, "/", -1, 1)
        Select Case strAction
            Case "Open":    AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":     AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":    AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":  AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:      AccessRights = False
        End Select
    Case "Personnel Action Memo"
        ArrAccess = Split(rs!PersonnelAction, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Print":       AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case "Supervisory": AccessRights = IIf(ArrAccess(5) = 1, True, False)
            Case Else:      AccessRights = False
        End Select
    Case "Personnel Gov't Table"
        ArrAccess = Split(rs!PersonnelGovt, "/", -1, 1)
        Select Case strAction
            Case "SSS":             AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "PHIC":            AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "PAGIBIG":         AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "TAX":             AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "PERSONAL_EXEMP":  AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case Else:              AccessRights = False
        End Select
    Case "Personnel Department"
        ArrAccess = Split(rs!PersonnelDept, "/", -1, 1)
        Select Case strAction
            Case "Open":    AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":     AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":    AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":  AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:      AccessRights = False
        End Select
    Case "Personnel Position"
        ArrAccess = Split(rs!PersonnelPost, "/", -1, 1)
        Select Case strAction
            Case "Open":    AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":     AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":    AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":  AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:      AccessRights = False
        End Select
    Case "Personnel Employment Status"
        ArrAccess = Split(rs!PersonnelStatus, "/", -1, 1)
        Select Case strAction
            Case "Open":    AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":     AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":    AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":  AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:      AccessRights = False
        End Select
    Case "Personnel Overtime/Restday Rate"
        ArrAccess = Split(rs!PersonnelOvertimeRestDay, "/", -1, 1)
        Select Case strAction
            Case "Open":    AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Edit":    AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case Else:      AccessRights = False
        End Select
    Case "Personnel Generate Payroll Period"
    
    Case "Personnel Loans"
        ArrAccess = Split(rs!PersonnelLoan, "/", -1, 1)
        Select Case strAction
            Case "Open":    AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":     AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":    AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":  AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Post":    AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case "UnPost":  AccessRights = IIf(ArrAccess(6) = 1, True, False)
            Case Else:      AccessRights = False
        End Select
    Case "Personnel Deduction"
        ArrAccess = Split(rs!PersonnelDeduction, "/", -1, 1)
        Select Case strAction
            Case "Open":    AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":     AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":    AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":  AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Post":    AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case "UnPost":  AccessRights = IIf(ArrAccess(5) = 1, True, False)
            Case Else:      AccessRights = False
        End Select
    Case "Personnel Compensation"
        ArrAccess = Split(rs!PersonnelCompensation, "/", -1, 1)
        Select Case strAction
            Case "Open":            AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":             AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":            AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":          AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Supervisory":     AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case "Locked Payroll":  AccessRights = IIf(ArrAccess(5) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Service Charge Setup"
        ArrAccess = Split(rs!ServiceChargeSetup, "/", -1, 1)
        Select Case strAction
            Case "Open":    AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":     AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":    AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":  AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:      AccessRights = False
        End Select
    Case "Perfect Days (Daily)"
        ArrAccess = Split(rs!PersonnelSetUpPerfectDays, "/", -1, 1)
        Select Case strAction
            Case "Open":    AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":     AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":    AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":  AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:      AccessRights = False
        End Select
    Case "PagIbig Additional Contribution"
        ArrAccess = Split(rs!PersonnelPagIbigAddContri, "/", -1, 1)
        Select Case strAction
            Case "Open":    AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":     AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":    AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":  AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:      AccessRights = False
        End Select
    Case "Absent/Late/Undertime Employee"
        ArrAccess = Split(rs!AbsentUndertime, "/", -1, 1)
        Select Case strAction
            Case "Open":    AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":     AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":    AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":  AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Post":    AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case "UnPost":  AccessRights = IIf(ArrAccess(5) = 1, True, False)
            Case Else:      AccessRights = False
        End Select
    Case "Service Charge"
        ArrAccess = Split(rs!ServiceCharge, "/", -1, 1)
        Select Case strAction
            Case "Open":    AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":     AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":    AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":  AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Post":    AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case "UnPost":  AccessRights = IIf(ArrAccess(5) = 1, True, False)
            Case Else:      AccessRights = False
        End Select
    Case "Service Charge Summary"
        ArrAccess = Split(rs!ServiceChargeSumm, "/", -1, 1)
        Select Case strAction
            Case "Open":    AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":     AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":    AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":  AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Post":    AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case Else:      AccessRights = False
        End Select
    Case "Scoring Tournament Information"
        ArrAccess = Split(rs!ScoringTournamentInfo, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Scoring Player Information"
        ArrAccess = Split(rs!ScoringPlayerInfo, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Scoring Team Information"
        ArrAccess = Split(rs!ScoringTeamInfo, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Scoring Score Card"
        ArrAccess = Split(rs!ScoringScoreCard, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Inventory Section"
        ArrAccess = Split(rs!Sections, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Inventory Classification"
        ArrAccess = Split(rs!Classification, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Inventory Supplier"
        ArrAccess = Split(rs!Supplier, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Inventory Items"
        ArrAccess = Split(rs!ItemInfo, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Membership Information"
        ArrAccess = Split(rs!MemberInfo, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Membership ID Number"
        ArrAccess = Split(rs!MemberIDNumber, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Membership Action"
        ArrAccess = Split(rs!MemberAction, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Corporate Account"
        ArrAccess = Split(rs!CorporateAccount, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Golf Cart Information"
        ArrAccess = Split(rs!GolfCart, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Caddy Information"
        ArrAccess = Split(rs!CaddyInfo, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Share ID Number"
        ArrAccess = Split(rs!ShareID, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Mortuary"
        ArrAccess = Split(rs!Mortuary, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Post":        AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case "UnPost":      AccessRights = IIf(ArrAccess(5) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Allowance"
        ArrAccess = Split(rs!Allowance, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Generate":    AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Allow Backup"
        ArrAccess = Split(rs!AllowBackup, "/", -1, 1)
        Select Case strAction
            Case "Backup":      AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Purchase Order"
        ArrAccess = Split(rs!PurchaseOrder, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Print":       AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case "Post":        AccessRights = IIf(ArrAccess(5) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Menu Management"
        ArrAccess = Split(rs!MenuMngt, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Fixed Assets"
        ArrAccess = Split(rs!FixedAsset, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Receiving Report"
        ArrAccess = Split(rs!ReceivingReport, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Post Rcd":    AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case "Post Inv":    AccessRights = IIf(ArrAccess(5) = 1, True, False)
            Case "Print":       AccessRights = IIf(ArrAccess(6) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Purchase Invoice"
        ArrAccess = Split(rs!PurchaseInvoice, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Print":       AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case "Post Inv":    AccessRights = IIf(ArrAccess(5) = 1, True, False)
            Case "Post GL":     AccessRights = IIf(ArrAccess(6) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Stock Transfer"
        ArrAccess = Split(rs!StockTransfer, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Post":        AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Stock Adjustment"
        ArrAccess = Split(rs!StockAdjustment, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Post":        AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Stock Issuance"
        ArrAccess = Split(rs!StockIssuance, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Post":        AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case "Post Inv":    AccessRights = IIf(ArrAccess(5) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Check Voucher"
        ArrAccess = Split(rs!CheckVoucher, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Post":        AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Journal Voucher"
        ArrAccess = Split(rs!JournalVoucher, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Post":        AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Chart Of Accounts"
        ArrAccess = Split(rs!ChartOfAccounts, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "PostRange":   AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Bag Drop"
        ArrAccess = Split(rs!BagDrop, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Post":        AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Registration"
        ArrAccess = Split(rs!Registration, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Post":        AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Pro Shop"
        ArrAccess = Split(rs!ProShop, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Post":        AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Pro Shop Items"
        ArrAccess = Split(rs!ProShopItems, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Driving Range"
        ArrAccess = Split(rs!DrivingRange, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Post":        AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Locker Room"
        ArrAccess = Split(rs!LockerRoom, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Post":        AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Golf Cart Operation"
        ArrAccess = Split(rs!GolfCartOP, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Post":        AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "FnB Location"
        ArrAccess = Split(rs!FnBLocation, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Pro Shop Items (Brand)"
        ArrAccess = Split(rs!ProShopItemsBrand, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Pro Shop Items (Model)"
        ArrAccess = Split(rs!ProShopItemsModel, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Pro Shop Items (Sizes)"
        ArrAccess = Split(rs!ProShopItemsSizes, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Pro Shop Items (Color)"
        ArrAccess = Split(rs!ProShopItemsColor, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Pro Shop Items (Item Type)"
        ArrAccess = Split(rs!ProShopItemsItemType, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Personnel - Hours"
        ArrAccess = Split(rs!PersonnelHours, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Post":        AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case "UnPost":      AccessRights = IIf(ArrAccess(5) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Personnel - For Deduction"
        ArrAccess = Split(rs!PersonnelForDeduction, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Post":        AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case "UnPost":      AccessRights = IIf(ArrAccess(5) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case "Personnel - Manual Payment"
        ArrAccess = Split(rs!PersonnelManualPayment, "/", -1, 1)
        Select Case strAction
            Case "Open":        AccessRights = IIf(ArrAccess(0) = 1, True, False)
            Case "Add":         AccessRights = IIf(ArrAccess(1) = 1, True, False)
            Case "Edit":        AccessRights = IIf(ArrAccess(2) = 1, True, False)
            Case "Delete":      AccessRights = IIf(ArrAccess(3) = 1, True, False)
            Case "Post":        AccessRights = IIf(ArrAccess(4) = 1, True, False)
            Case "UnPost":      AccessRights = IIf(ArrAccess(5) = 1, True, False)
            Case Else:          AccessRights = False
        End Select
    Case Else
        AccessRights = True
End Select
rs.Close
End Function
