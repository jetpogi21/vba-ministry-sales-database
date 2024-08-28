Attribute VB_Name = "FinancialReport Mod"
Option Compare Database
Option Explicit

Public Function FinancialReportCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function

Public Function frmFinancialReports_cmdPrint_OnClick(frm As Form)
    
    Dim rs As Recordset: Set rs = ReturnRecordset("qryFinancialReports")
    If CountRecordset(rs) = 0 Then
        ShowError "There is no record to print."
        Exit Function
    End If
    
    DoCmd.OpenReport "rptFinancialReports", acViewPreview
    
End Function

Public Sub rptFinancialReports_Create()
    
    Const HEADER = "Financial Report"
    Const REPORT_NAME = "rptFinancialReports"
    Const RECORDSOURCE_NAME = "qryFinancialReports"
    
    Dim rpt As Report: Set rpt = CreateReport()
    SetCommonReportProperties rpt
    rpt.recordSource = RECORDSOURCE_NAME
    rpt.Caption = HEADER
    
    Dim ctl As control
    Set ctl = CreateLabelControl(rpt, HEADER, "Header", "Heading1", acPageHeader)
    Set ctl = CreateTextboxControl(rpt, """Komm.Nr. "" & CustomerOrderID2", "txtCustomerOrderID", , "Heading2", acPageHeader)
    
    ''AUFTRAG --> CustomerOrderReportOrders
    Set ctl = CreateLabelControl(rpt, "AUFTRAG", "AUFTRAG", "Heading3")
    Offset_ctlPositions rpt, ctl, , InchToTwip(0.25)
    Set ctl = CreateSubformControl(rpt, "rptFinancialReportOrders", "subFinancialReportOrders", , "CustomerOrderID")
    
    ''SUBLIEFERANTEN --> CustomerOrderReportSubSuppliers
    Set ctl = CreateLabelControl(rpt, "SUBLIEFERANTEN", "SUBLIEFERANTEN", "Heading3")
    Offset_ctlPositions rpt, ctl, , InchToTwip(0.1)
    Set ctl = CreateSubformControl(rpt, "srptFinancialReportSubSuppliers", "subFinancialReportSubSuppliers", , "CustomerOrderID")
    
    ''MATERIALIEN --> CustomerOrderReportMaterials
    Set ctl = CreateLabelControl(rpt, "MATERIALIEN", "MATERIALIEN", "Heading3")
    Offset_ctlPositions rpt, ctl, , InchToTwip(0.1)
    Set ctl = CreateSubformControl(rpt, "srptFinancialReportOrderMaterials", "subFinancialReportOrderMaterials", , "CustomerOrderID")
    
    Set ctl = CreateTextboxControl(rpt, "GetSubformValue([subFinancialReportSubSuppliers].[Report]![SumTotalCost])", _
        "TotalSubSupplierCost", , , , True)
    Set ctl = CreateTextboxControl(rpt, "GetSubformValue([subFinancialReportOrderMaterials].[Report]![SumTotalCost])", _
        "TotalMaterialCost", , , , True)
    Set ctl = CreateTextboxControl(rpt, "GetSubformValue([subFinancialReportOrders].[Report]![Revenue])", _
        "Revenue", , , , True)
    ''=GetSubformValue([subCustomerOrderReportSubSuppliers].[Form]![SumTotalCost])
    ''=GetSubformValue([subCustomerOrderReportMaterials].[Form]![SumTotalCost])
    ''=GetSubformValue([subCustomerOrderReportOrders].[Form]![Revenue])
    
    ''Total Costs -> =[TotalSubSupplierCost]+[TotalMaterialCost]
    Set ctl = CreateTextboxControl(rpt, "[TotalSubSupplierCost]+[TotalMaterialCost]", "TotalCost", "TOTAL COSTS")
    ctl.Format = "Standard"
    ''Performance -> =[Revenue]-[TotalCost]
    Set ctl = CreateTextboxControl(rpt, "[Revenue]-[TotalCost]", "Performance", "PERFORMANCE")
    ctl.Format = "Standard"
    ''Margin -> =[Performance]/[Revenue]
    Set ctl = CreateTextboxControl(rpt, "[Performance]/[Revenue]", "Margin", "MARGIN")
    ctl.Format = "Percent"
    
    Offset_ctlPositions rpt, rpt("lblTotalCost"), 50, InchToTwip(0.1)
    Offset_ctlPositions rpt, rpt("TotalCost"), 50, InchToTwip(0.1)
    
    RepositionControlsInRow rpt, "1,1,1", "lblTotalCost,lblPerformance,lblMargin", rpt("lblTotalCost").Left, rpt("lblTotalCost").Width * 0.9, _
        InchToTwip(0.1), rpt("lblTotalCost").Top, True
        
    RepositionControlsInRow rpt, "1,1,1", "TotalCost,Performance,Margin", rpt("TotalCost").Left, rpt("TotalCost").Width * 0.9, InchToTwip(0.1), _
        rpt("TotalCost").Top, True
    
    CreateBannerControls rpt
    
    CleanUpReportProperties rpt
    
    RenameFormOrReport rpt.Name, REPORT_NAME
    
    GetFormOrReport REPORT_NAME, True, True
    
End Sub
