Attribute VB_Name = "UserReport Mod"
Option Compare Database
Option Explicit

Public Function UserReportCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
            frm("lblChargedAmount").Height = frm("lblChargedAmount").Height * 2
            frm("lblChargedAmount").Width = frm("lblChargedAmount").Width * 2 / 3
            frm("ChargedAmount").Width = frm("lblChargedAmount").Width
            
            Dim Right: Right = frm("ChargedAmount").Width + frm("ChargedAmount").Left
            frm("DiscountGiven").Left = Right
            frm("lblDiscountGiven").Left = Right
            
            Right = GetRight(frm("DiscountGiven"))
            frm("UserName").Left = Right
            frm("lblUserName").Left = Right
            
            CreateTotalControl frm, "StandardFee"
            CreateTotalControl frm, "ChargedAmount"
            CreateTotalControl frm, "DiscountGiven"
            CreateTotalLabel frm, "StandardFee"
            
            OffsetControlPositions frm, 50
            
            frm.Printer.Orientation = acPRORLandscape
            
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function

Public Function Open_frmUserReports(frm As Form)
    
    Dim fltrUser: fltrUser = frm("fltrUser")
    Dim fltrDateFrom: fltrDateFrom = frm("fltrDateFrom")
    Dim fltrDateTo: fltrDateTo = frm("fltrDateTo")
    Dim fltrMinistry: fltrMinistry = frm("fltrMinistry")
    Dim fltrTask: fltrTask = frm("fltrTask")
    
    CloseThisForm frm
    
    open_form "frmUserReports"
    
    Set frm = GetForm("frmUserReports")
    
    frm("fltrUser") = fltrUser
    frm("fltrDateFrom") = fltrDateFrom
    frm("fltrDateTo") = fltrDateTo
    frm("fltrMinistry") = fltrMinistry
    frm("fltrTask") = fltrTask
    
    frmUserReports_fltr_AfterUpdate frm
    
End Function

Public Function Open_frmAnalytics(frm As Form)
    
    Dim fltrUser: fltrUser = frm("fltrUser")
    Dim fltrDateFrom: fltrDateFrom = frm("fltrDateFrom")
    Dim fltrDateTo: fltrDateTo = frm("fltrDateTo")
    Dim fltrMinistry: fltrMinistry = frm("fltrMinistry")
    Dim fltrTask: fltrTask = frm("fltrTask")
    
    CloseThisForm frm
    
    open_form "frmAnalytics"
    
    Set frm = GetForm("frmAnalytics")
    
    frm("fltrUser") = fltrUser
    frm("fltrDateFrom") = fltrDateFrom
    frm("fltrDateTo") = fltrDateTo
    frm("fltrMinistry") = fltrMinistry
    frm("fltrTask") = fltrTask
    
    frmAnalytics_OnLoad frm
    
End Function

Public Function fltrMinistry_AfterUpdate(frm As Form)
    
    If frm.Name = "frmUserReports" Then
        frmUserReports_fltr_AfterUpdate frm, True
    Else
        frmAnalytics_fltr_AfterUpdate frm, True
    End If
    
End Function

Public Function frmUserReports_fltr_AfterUpdate(frm As Form, Optional Reset_fltrTask As Boolean = False)
    
    Set_fltr_RowSource frm, Reset_fltrTask
    Set_subform_RecordSource frm
    
End Function

Public Function Reset_fltrs(frm As Form)
    
    frm("fltrUser") = -2
    frm("fltrDateFrom") = Null
    frm("fltrDateTo") = Null
    frm("fltrMinistry") = -2
    frm("fltrTask") = -2
    
    If frm.Name = "frmUserReports" Then
        frmUserReports_fltr_AfterUpdate frm
    Else
        frmAnalytics_fltr_AfterUpdate frm
    End If
    
End Function

Public Function frmUserReports_cmdPrint_OnClick(frm As Form)
    
    Dim rs As Recordset: Set rs = ReturnRecordset("qryUserReports")
    If CountRecordset(rs) = 0 Then
        ShowError "There is no record to print."
        Exit Function
    End If
    
    DoCmd.OpenReport "rptUserReports", acViewPreview
    
End Function

Private Sub Set_fltr_RowSource(frm As Form, Optional Reset_fltrTask As Boolean = False)
    
    Set_fltrUser_RowSource frm
    
    Dim sqlStr: sqlStr = "SELECT ALL_Number as MinistryID, [All] as Ministry FROM tblAlls where All_Number = -2"
    sqlStr = "SELECT MinistryID,Ministry FROM tblMinistries ORDER BY Ministry UNION ALL " & sqlStr
    
    sqlStr = "SELECT * FROM (" & sqlStr & ") temp ORDER BY MinistryID"
    frm("fltrMinistry").RowSource = sqlStr
    
    Dim fltrMinistry: fltrMinistry = frm("fltrMinistry")
    
    sqlStr = "SELECT ALL_Number as MinistryTaskID, [All] as Task FROM tblAlls where All_Number = -2"
    Dim filterStr: filterStr = "MinistryTaskID > 0"
    If Not isFalse(fltrMinistry) Then
        If fltrMinistry > 0 Then filterStr = "MinistryID = " & fltrMinistry
    End If
    
    sqlStr = "SELECT MinistryTaskID,Task FROM tblMinistryTasks WHERE " & filterStr & " ORDER BY Task UNION ALL " & sqlStr
    sqlStr = "SELECT * FROM (" & sqlStr & ") temp ORDER BY MinistryTaskID"
    frm("fltrTask").RowSource = sqlStr
    
    If Reset_fltrTask Then
        frm("fltrTask") = -2
    End If
    
End Sub

Public Function frmUserReports_OnLoad(frm As Form)
    
    frmUserReports_AlignControlsBasedOnUser frm
    Set_subform_RecordSource frm
    Set_fltr_RowSource frm
    TranslateToArabic frm
    Set_fltr_AfterUpdate frm
    
End Function

Public Sub frmUserReports_AlignControlsBasedOnUser(frm As Form)
    
    If GetIsAdmin Then Exit Sub
    
    Dim DistanceBetweenControls: DistanceBetweenControls = InchToTwip(0.1)
    Dim Top: Top = frm("lblfltrUser").Top
    ''Hide these controls
    ''lblfltrUser,fltrUser,fltrTask,fltrMinistry,lblfltrMinistry,fltrTask,lblfltrTask,cmdOpenAnalytics
    Dim controlArr As New clsArray: controlArr.arr = "lblfltrUser,fltrUser,fltrTask,fltrMinistry,lblfltrMinistry," & _
        "fltrTask,lblfltrTask,cmdOpenAnalytics"
    
    Dim item
    For Each item In controlArr.arr
        frm(item).Visible = False
    Next item
    
    ''Adjust the width
    frm("lblDateFrom").Width = frm("lblDateFrom").Width * 2 / 3
    
    ''Move the fltrDateFrom,lblDateFrom,fltrDateTo,lblDateTo at the top position of fltrUser
    Set controlArr = New clsArray: controlArr.arr = "lblDateFrom,fltrDateFrom,lblDateTo,fltrDateTo"
    Dim i: i = 0
    For Each item In controlArr.arr
        frm(item).Top = Top
        If i <> 0 Then
            frm(item).Left = GetRight(frm(controlArr.arr(i - 1))) + DistanceBetweenControls
        End If
        i = i + 1
    Next item
    
    ''Move the cmdRefresh at the right of fltrDateTo + 0.25 Inches
    frm("cmdRefresh").Left = GetRight(frm("fltrDateTo")) + DistanceBetweenControls
    
    ''Move the cmdMainMenu to the left of  txtPrint - 0.25 Inches
    frm("cmdMainMenu").Left = frm("txtPrint").Left - frm("cmdMainMenu").Width - DistanceBetweenControls
    
    ''Move the line and subform at the top
    frm("Line57").Top = GetBottom(frm("lblDateFrom")) + InchToTwip(0.25)
    frm("subform").Top = GetBottom(frm("lblDateFrom")) + InchToTwip(0.5)
    
End Sub

Private Sub Set_fltr_AfterUpdate(frm As Form)
    
    'fltrUser,fltrDateFrom,fltrDateTo,fltrMinistry,fltrTask
    Dim fltrArr As New clsArray: fltrArr.arr = "fltrUser,fltrDateFrom,fltrDateTo,fltrMinistry,fltrTask"
    
    Dim item, items As New clsArray
    For Each item In fltrArr.arr
        frm(item).AfterUpdate = "=frmUserReports_fltr_AfterUpdate([Form])"
    Next item
    
    frm("fltrMinistry").AfterUpdate = "=fltrMinistry_AfterUpdate([Form])"
    
End Sub

Private Sub Set_fltrUser_RowSource(frm As Form)
    Dim sqlStr: sqlStr = "SELECT ALL_Number as UserID, [All] as Username FROM tblAlls where All_Number = -2"
    sqlStr = "SELECT UserID, Username FROM tblUsers ORDER BY Username UNION ALL " & sqlStr
    
    sqlStr = "SELECT * FROM (" & sqlStr & ") temp ORDER BY UserID"
    frm("fltrUser").RowSource = sqlStr
End Sub



Private Sub Set_subform_RecordSource(frm As Form)

    Dim fltrUser: fltrUser = frm("fltrUser")
    Dim fltrDateFrom: fltrDateFrom = frm("fltrDateFrom")
    Dim fltrDateTo: fltrDateTo = frm("fltrDateTo")
    Dim fltrMinistry: fltrMinistry = frm("fltrMinistry")
    Dim fltrTask: fltrTask = frm("fltrTask")
    
    Dim filterStr: filterStr = "TransactionID > 0"
    
    Dim fltrArr As New clsArray
    
    If Not isFalse(fltrUser) Then
        If fltrUser > 0 Then fltrArr.Add "CreatedBy = " & fltrUser
    End If
    
    If Not isFalse(fltrMinistry) Then
        If fltrMinistry > 0 Then fltrArr.Add "MinistryID = " & fltrMinistry
    End If
    
    If Not isFalse(fltrTask) Then
        If fltrTask > 0 Then fltrArr.Add "MinistryTaskID = " & fltrTask
    End If
    
    If Not isFalse(fltrDateFrom) And Not isFalse(fltrDateTo) Then
        fltrArr.Add "TransactionDate Between " & EscapeString(fltrDateFrom, "tblTransactions", "TransactionDate") & " AND " & _
            EscapeString(fltrDateTo, "tblTransactions", "TransactionDate")
    End If
    
    If fltrArr.count > 0 Then
        filterStr = fltrArr.JoinArr(" AND ")
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM qryTransactions WHERE " & filterStr & " ORDER BY TransactionDate ASC"
    SetQueryDefSQL "qryUserReports", sqlStr
    
    frm("subform").SourceObject = "Report.rptUserReports"
    TranslateToArabic frm
    
End Sub

Public Function GetReportFilterCaption() As String

    Dim frm As Form: Set frm = GetForm("frmUserReports")
    
    If frm Is Nothing Then Exit Function
    
    Dim fltrArr As New clsArray
    
    Dim fltrUser: fltrUser = frm("fltrUser")
    Dim fltrDateFrom: fltrDateFrom = frm("fltrDateFrom")
    Dim fltrDateTo: fltrDateTo = frm("fltrDateTo")
    Dim fltrMinistry: fltrMinistry = frm("fltrMinistry")
    Dim fltrTask: fltrTask = frm("fltrTask")
    
    If Not isFalse(fltrUser) Then
        fltrArr.Add "User: " & frm("fltrUser").Column(1)
    End If
    
    If Not isFalse(fltrMinistry) Then
        If fltrMinistry > 0 Then fltrArr.Add "Ministry: " & frm("fltrMinistry").Column(1)
    End If
    
    If Not isFalse(fltrTask) Then
        If fltrTask > 0 Then fltrArr.Add "Task: " & frm("fltrTask").Column(1)
    End If
    
    If Not isFalse(fltrDateFrom) And Not isFalse(fltrDateTo) Then
        fltrArr.Add "Date: " & fltrDateFrom & " to " & fltrDateTo
    End If
    
    If fltrArr.count > 0 Then
        GetReportFilterCaption = fltrArr.JoinArr("  ")
    End If
    
    
End Function

Public Sub rptUserReports_Create()
    
    Const HEADER = "By User Category"
    Const REPORT_NAME = "rptUserReports"
    Const RECORDSOURCE_NAME = ""
    
    Dim rpt As Report: Set rpt = CreateReport()
    SetCommonReportProperties rpt
    rpt.recordSource = RECORDSOURCE_NAME
    rpt.Caption = HEADER
    rpt.OnLoad = "=DefaultReportLoad([Report])"
    rpt.Printer.Orientation = acPRORLandscape
    
    Dim ctl As control
    Set ctl = CreateLabelControl(rpt, HEADER, "Header", "Heading1", acPageHeader)
    Set ctl = CreateTextboxControl(rpt, "GetReportFilterCaption()", "txtFilterCaption", , "Heading3", acPageHeader)
    
    ''AUFTRAG --> CustomerOrderReportOrders
    Set ctl = CreateLabelControl(rpt, "TRANSACTIONS", "TRANSACTIONS", "Heading3")
    Offset_ctlPositions rpt, ctl, , InchToTwip(0.25)
    Set ctl = CreateSubformControl(rpt, "srptUserReports", "subUserReports", , "")
    ctl.Left = 0
    ctl.Width = InchToTwip(11 - 0.5)
    
'    ''SUBLIEFERANTEN --> CustomerOrderReportSubSuppliers
'    Set ctl = CreateLabelControl(rpt, "SUBLIEFERANTEN", "SUBLIEFERANTEN", "Heading3")
'    Offset_ctlPositions rpt, ctl, , InchToTwip(0.1)
'    Set ctl = CreateSubformControl(rpt, "srptFinancialReportSubSuppliers", "subFinancialReportSubSuppliers", , "CustomerOrderID")
'
'    ''MATERIALIEN --> CustomerOrderReportMaterials
'    Set ctl = CreateLabelControl(rpt, "MATERIALIEN", "MATERIALIEN", "Heading3")
'    Offset_ctlPositions rpt, ctl, , InchToTwip(0.1)
'    Set ctl = CreateSubformControl(rpt, "srptFinancialReportOrderMaterials", "subFinancialReportOrderMaterials", , "CustomerOrderID")
'
'    Set ctl = CreateTextboxControl(rpt, "GetSubformValue([subFinancialReportSubSuppliers].[Report]![SumTotalCost])", _
'        "TotalSubSupplierCost", , , , True)
'    Set ctl = CreateTextboxControl(rpt, "GetSubformValue([subFinancialReportOrderMaterials].[Report]![SumTotalCost])", _
'        "TotalMaterialCost", , , , True)
'    Set ctl = CreateTextboxControl(rpt, "GetSubformValue([subFinancialReportOrders].[Report]![Revenue])", _
'        "Revenue", , , , True)
'    ''=GetSubformValue([subCustomerOrderReportSubSuppliers].[Form]![SumTotalCost])
'    ''=GetSubformValue([subCustomerOrderReportMaterials].[Form]![SumTotalCost])
'    ''=GetSubformValue([subCustomerOrderReportOrders].[Form]![Revenue])
'
'    ''Total Costs -> =[TotalSubSupplierCost]+[TotalMaterialCost]
'    Set ctl = CreateTextboxControl(rpt, "[TotalSubSupplierCost]+[TotalMaterialCost]", "TotalCost", "TOTAL COSTS")
'    ctl.Format = "Standard"
'    ''Performance -> =[Revenue]-[TotalCost]
'    Set ctl = CreateTextboxControl(rpt, "[Revenue]-[TotalCost]", "Performance", "PERFORMANCE")
'    ctl.Format = "Standard"
'    ''Margin -> =[Performance]/[Revenue]
'    Set ctl = CreateTextboxControl(rpt, "[Performance]/[Revenue]", "Margin", "MARGIN")
'    ctl.Format = "Percent"
'
'    Offset_ctlPositions rpt, rpt("lblTotalCost"), 50, InchToTwip(0.1)
'    Offset_ctlPositions rpt, rpt("TotalCost"), 50, InchToTwip(0.1)
'
'    RepositionControlsInRow rpt, "1,1,1", "lblTotalCost,lblPerformance,lblMargin", rpt("lblTotalCost").Left, rpt("lblTotalCost").Width * 0.9, _
'        InchToTwip(0.1), rpt("lblTotalCost").Top, True
'
'    RepositionControlsInRow rpt, "1,1,1", "TotalCost,Performance,Margin", rpt("TotalCost").Left, rpt("TotalCost").Width * 0.9, InchToTwip(0.1), _
'        rpt("TotalCost").Top, True
    
    CreateBannerControls rpt
    
    CleanUpReportProperties rpt
    
    RenameFormOrReport rpt.Name, REPORT_NAME
    
    GetFormOrReport REPORT_NAME, True, True
    
End Sub


