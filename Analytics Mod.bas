Attribute VB_Name = "Analytics Mod"
Option Compare Database
Option Explicit

Public Function AnalyticsCreate(frm As Object, FormTypeID)

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
            
            
            CreateTotalControl frm, "StandardFee"
            CreateTotalControl frm, "ChargedAmount"
            CreateTotalControl frm, "DiscountGiven"
            CreateTotalLabel frm, "StandardFee"
            
            OffsetControlPositions frm, 50
            
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function

Public Function frmAnalytics_fltrUser_AfterUpdate(frm As Form)
    
    SetChart_QueryDefs frm
    
End Function

Public Function frmAnalytics_fltrDateFrom_AfterUpdate(frm As Form)
    
    SetChart_QueryDefs frm
    
End Function

Public Function frmAnalytics_cmdPrint_OnClick(frm As Form)
    
    Dim rs As Recordset: Set rs = ReturnRecordset("qryAnalytics")
    If CountRecordset(rs) = 0 Then
        ShowError "There is no record to print."
        Exit Function
    End If
    
    DoCmd.OpenReport "rptAnalytics", acViewPreview
    
End Function

Public Function frmAnalytics_fltrDateTo_AfterUpdate(frm As Form)

    SetChart_QueryDefs frm
    
End Function

Public Function frmAnalytics_OnLoad(frm As Form)

    ''Set_subform_RecordSource frm
    DisplayDataLabel frm
    TranslateToArabic frm
    Set_fltr_AfterUpdate frm
    Set_fltr_RowSource frm
    SetChart_QueryDefs frm
    
End Function

Private Sub Set_fltr_AfterUpdate(frm As Form)
    
    'fltrUser,fltrDateFrom,fltrDateTo,fltrMinistry,fltrTask
    Dim fltrArr As New clsArray: fltrArr.arr = "fltrUser,fltrDateFrom,fltrDateTo,fltrMinistry,fltrTask"
    
    Dim item, items As New clsArray
    For Each item In fltrArr.arr
        frm(item).AfterUpdate = "=frmAnalytics_fltr_AfterUpdate([Form])"
    Next item
    
    frm("fltrMinistry").AfterUpdate = "=fltrMinistry_AfterUpdate([Form])"
    
End Sub

Public Function frmAnalytics_fltr_AfterUpdate(frm As Form, Optional Reset_fltrTask As Boolean = False)
    
    Set_fltr_RowSource frm, Reset_fltrTask
    SetChart_QueryDefs frm
    
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
    Debug.Print "Reset_fltrTask: " & Reset_fltrTask
    If Reset_fltrTask Then
        frm("fltrTask") = -2
    End If
    
End Sub

Private Sub SetChart_QueryDefs(frm As Object)
    
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
    
    Dim sqlStr: sqlStr = "SELECT UserName, Count(TransactionID) As CountOfTransactionID FROM qryTransactions WHERE " & filterStr & " GROUP BY UserName"
    SetQueryDefSQL "qryTransactionsPerUser", sqlStr
    
    frm("chtTransactionsPerUser").Requery
    
    sqlStr = "SELECT Ministry, Count(TransactionID) As CountOfTransactionID FROM qryTransactions WHERE " & filterStr & " GROUP BY Ministry"
    SetQueryDefSQL "qryTransactionsPerMinistry", sqlStr
    
    frm("chtTransactionsPerMinistry").Requery
    
    sqlStr = "SELECT Format(TransactionDate, ""MM/DD"") As Caption, Count(TransactionID) As CountOfTransactionID FROM qryTransactions WHERE " & filterStr & " GROUP BY TransactionDate " & _
        " ORDER BY TransactionDate"
        
    SetQueryDefSQL "qryTransactionsPerDate", sqlStr
    
    frm("chtTransactionsPerDate").Requery
    
End Sub

Private Sub DisplayDataLabel(frm As Form)

    Dim var As ChartSeries
    For Each var In frm("chtTransactionsPerUser").ChartSeriesCollection
        var.DisplayDataLabel = True
    Next
    
    For Each var In frm("chtTransactionsPerMinistry").ChartSeriesCollection
        var.DisplayDataLabel = True
    Next
    
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
    
    Dim filterStr: filterStr = "TransactionID = 0"
    
    Dim fltrArr As New clsArray
    
    If Not isFalse(fltrUser) Then
        If fltrUser > 0 Then fltrArr.Add "CreatedBy = " & fltrUser
    End If
    
    If Not isFalse(fltrDateFrom) Then
        fltrArr.Add "TransactionDate >= " & EscapeString(fltrDateFrom, "tblTransactions", "TransactionDate")
    End If
    
    If Not isFalse(fltrDateTo) Then
        fltrArr.Add "TransactionDate <= " & EscapeString(fltrDateTo, "tblTransactions", "TransactionDate")
    End If
    
    If fltrArr.count > 0 Then
        filterStr = fltrArr.JoinArr(" AND ")
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM qryTransactions WHERE " & filterStr & " ORDER BY TransactionDate ASC"
    SetQueryDefSQL "qryAnalytics", sqlStr
    
    frm("subform").SourceObject = "Report.rptAnalytics"
    
End Sub

Public Sub rptAnalytics_Create()
    
    Const HEADER = "Per User Report"
    Const REPORT_NAME = "rptAnalytics"
    Const RECORDSOURCE_NAME = ""
    
    Dim rpt As Report: Set rpt = CreateReport()
    SetCommonReportProperties rpt
    rpt.recordSource = RECORDSOURCE_NAME
    rpt.Caption = HEADER
    
    Dim ctl As control
    Set ctl = CreateLabelControl(rpt, HEADER, "Header", "Heading1", acPageHeader)
    
    ''AUFTRAG --> CustomerOrderReportOrders
    Set ctl = CreateLabelControl(rpt, "TRANSACTIONS", "TRANSACTIONS", "Heading3")
    Offset_ctlPositions rpt, ctl, , InchToTwip(0.25)
    Set ctl = CreateSubformControl(rpt, "srptAnalytics", "subAnalytics", , "")
    
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




