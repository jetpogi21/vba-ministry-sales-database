Attribute VB_Name = "CustomerOrderReport Mod"
Option Compare Database
Option Explicit

Public Function CustomerOrderReportCreate(frm As Object, FormTypeID)

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

Public Function frmCustomerOrderReports_cmdPrint_OnClick(frm As Form)
    
    Dim rs As Recordset: Set rs = ReturnRecordset("qryCustomerOrderReports")
    If CountRecordset(rs) = 0 Then
        ShowError "There is no record to print."
        Exit Function
    End If
    
    DoCmd.OpenReport "rptCustomerOrderReports", acViewPreview
    
End Function

Public Function frmCustomerOrderReports_OnLoad(frm As Form)

    Set_subform_RecordSource frm
    
End Function

Public Sub rptCustomerOrderReports_Create()
    
    Const HEADER = "Customer Order Report"
    Const REPORT_NAME = "rptCustomerOrderReports"
    Const RECORDSOURCE_NAME = "qryCustomerOrderReports"
    
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
    Set ctl = CreateSubformControl(rpt, "rptCustomerOrderReportOrders", "subCustomerOrderReportOrders", , "CustomerOrderID")
    
    ''SUBLIEFERANTEN --> CustomerOrderReportSubSuppliers
    Set ctl = CreateLabelControl(rpt, "SUBLIEFERANTEN", "SUBLIEFERANTEN", "Heading3")
    Offset_ctlPositions rpt, ctl, , InchToTwip(0.1)
    Set ctl = CreateSubformControl(rpt, "srptCustomerOrderReportSubSuppliers", "subCustomerOrderReportSubSuppliers", , "CustomerOrderID")
    
    ''MATERIALIEN --> CustomerOrderReportMaterials
    Set ctl = CreateLabelControl(rpt, "MATERIALIEN", "MATERIALIEN", "Heading3")
    Offset_ctlPositions rpt, ctl, , InchToTwip(0.1)
    Set ctl = CreateSubformControl(rpt, "srptCustomerOrderReportMaterials", "subCustomerOrderReportMaterials", , "CustomerOrderID")
    
    ''AUSLIEFERUNG --> CustomerOrderReportDeliveries
    Set ctl = CreateLabelControl(rpt, "AUSLIEFERUNG", "AUSLIEFERUNG", "Heading3")
    Offset_ctlPositions rpt, ctl, , InchToTwip(0.1)
    Set ctl = CreateSubformControl(rpt, "srptCustomerOrderReportDeliveries", "subCustomerOrderReportDeliveries", , "CustomerOrderID")
    ''CustomerOrderReportNCs
    Set ctl = CreateSubformControl(rpt, "srptCustomerOrderReportNCs", "subCustomerOrderReportNCs", , "CustomerOrderID")
    ''CustomerOrderReportNCs
    Set ctl = CreateSubformControl(rpt, "srptCustomerOrderReportForceCloses", "subCustomerOrderReportForceCloses", , "CustomerOrderID")
    
    CreateBannerControls rpt
    
    CleanUpReportProperties rpt
    
    RenameFormOrReport rpt.Name, REPORT_NAME
    
    GetFormOrReport REPORT_NAME, True, True
    
End Sub

Public Sub frmCustomerOrderReports_Racalculate(Optional RunEvenWhenClosed As Boolean = False)
    
    Dim frmName: frmName = "frmCustomerOrderReports"
    Dim frm As Form
    If IsFormOpen(frmName) Then
        Set frm = Forms(frmName)
    Else
        If RunEvenWhenClosed Then
            DoCmd.OpenForm frmName
            Set frm = Forms(frmName)
        End If
    End If
    
    If frm Is Nothing Then
        Exit Sub
    End If
    
    frmCustomerOrderReports_fltrCommissionNumber_AfterUpdate frm
    
End Sub

Public Sub SetUp_frmCustomerOrderReports()
    
    Dim frmName: frmName = "frmCustomerOrderReports"
    
    If Not IsFormOpen(frmName) Then
        DoCmd.OpenForm frmName, acDesign
    End If
    
    Dim frm As Form: Set frm = Forms(frmName)
    
    CreateBannerControls frm
    
    Dim i As Integer: i = 16
    
    For i = 16 To 23
        Dim TextBoxName: TextBoxName = "Text" & i
        If DoesPropertyExists(frm, TextBoxName) Then
            DeleteControl frm.Name, TextBoxName
        End If
    Next i
    
End Sub


Public Function frmCustomerOrderReports_SetCustomerOrderID(frm As Form)

    Dim CustomerOrderID: CustomerOrderID = frm("fltrCommissionNumber")
    
    frm("CustomerOrderID") = CustomerOrderID
    
End Function

Public Function frmCustomerOrderReports_SetfltrCustomerShortName(frm As Form)

    Dim CustomerID: CustomerID = frm("fltrCommissionNumber").Column(2)
    
    frm("fltrCustomerShortName") = CustomerID
    
End Function

Public Function frmCustomerOrderReports_fltrCommissionNumber_AfterUpdate(frm As Form)
    
    frmCustomerOrderReports_SetCustomerOrderID frm
    frmCustomerOrderReports_SetfltrCustomerShortName frm
    frmCustomerOrderReports_fltrOrderDate_SetRowSource frm
    frmCustomerOrderReports_SetfltrOrderDate frm
    Set_subform_RecordSource frm
    ''frmCustomerOrderReports_fltrOrderDate_AfterUpdate frm
    
End Function

Private Sub Set_subform_RecordSource(frm As Form)

    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    Dim filterStr: filterStr = "CustomerOrderID = 0"
    If Not isFalse(CustomerOrderID) Then
        filterStr = "CustomerOrderID = " & CustomerOrderID
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM qryCustomerOrders WHERE " & filterStr & " ORDER BY CustomerOrderID2"
    SetQueryDefSQL "qryCustomerOrderReports", sqlStr
    
    frm("subform").SourceObject = "Report.rptCustomerOrderReports"
    
End Sub

Public Function frmCustomerOrderReports_SetfltrOrderDate(frm As Form)

    Dim CustomerOrderID: CustomerOrderID = frm("fltrCommissionNumber")
    
    frm("fltrOrderDate") = CustomerOrderID
    
End Function

Private Function frmCustomerOrderReports_fltrOrderDate_SetRowSource(frm As Form)

    Dim CustomerID: CustomerID = frm("fltrCustomerShortName")
    
    Dim filterStr: filterStr = "CustomerOrderID = 0"
    If Not isFalse(CustomerID) Then
        filterStr = "CustomerID = " & CustomerID
    End If
    
    Dim sqlStr: sqlStr = "SELECT CustomerOrderID, OrderDate FROM qryCustomerOrders WHERE " & filterStr & " ORDER BY OrderDate,CustomerOrderID"
    
    frm("fltrOrderDate").RowSource = sqlStr
    
End Function

Public Function frmCustomerOrderReports_fltrCustomerShortName_AfterUpdate(frm As Form)
    
    frmCustomerOrderReports_fltrOrderDate_SetRowSource frm
    frm("fltrOrderDate") = Null
    
End Function

Public Function frmCustomerOrderReports_fltrOrderDate_AfterUpdate(frm As Form)
    
    Dim CustomerOrderID: CustomerOrderID = frm("fltrOrderDate")
    frm("fltrCommissionNumber") = CustomerOrderID
    
    frmCustomerOrderReports_SetCustomerOrderID frm
    Set_subform_RecordSource frm
    
End Function
