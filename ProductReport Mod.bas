Attribute VB_Name = "ProductReport Mod"
Option Compare Database
Option Explicit

Public Function ProductReportCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
        
            CreateTotalControl frm, "QtyOnStock"
            CreateTotalLabel frm, "QtyOnStock"
            FormatFormAsReport frm, 8
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function

Public Function frmProductReports_cmdPrint_OnClick(frm As Form)

    Dim rs As Recordset: Set rs = ReturnRecordset("qryProductReports")
    Dim rs2 As Recordset: Set rs2 = ReturnRecordset("qryQuarantinedProductReports")
    If CountRecordset(rs) + CountRecordset(rs2) = 0 Then
        ShowError "There is no record to print."
        Exit Function
    End If
    
    DoCmd.OpenReport "rptProductReports", acViewPreview
    
End Function

Private Function Set_subform_RecordSource(ByVal frm As Form)
    
    Dim ProductID: ProductID = frm("fltrProduct")
    Dim fltrInStock: fltrInStock = frm("fltrInStock")

    Dim filterStr: filterStr = "ProductID = 0"
    
    Dim filterArr As New clsArray
    If Not isFalse(ProductID) Then
        filterArr.Add "ProductID = " & ProductID
        filterStr = "ProductID = " & ProductID
    End If
    
    If Not isFalse(fltrInStock) Then
        filterArr.Add "QtyOnStock > 0"
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM qryProducts WHERE " & filterStr & " ORDER BY ProductID"
    SetQueryDefSQL "qryProductReportProducts", sqlStr
    
    frm("subform").SourceObject = "Report.rptProductReports"
    
End Function

Public Sub rptProductReports_Create()
    
    Const HEADER = "Product Report"
    Const REPORT_NAME = "rptProductReports"
    Const RECORDSOURCE_NAME = "qryProductReportProducts"
    
    Dim rpt As Report: Set rpt = CreateReport()
    SetCommonReportProperties rpt
    rpt.recordSource = RECORDSOURCE_NAME
    rpt.Caption = HEADER
    
    Dim ctl As control
    Set ctl = CreateLabelControl(rpt, HEADER, "Header", "Heading1", acPageHeader)
     
    ''Product --> ProductReportProducts
    Set ctl = CreateLabelControl(rpt, "PRODUCT", "PRODUCT", "Heading3")
    Offset_ctlPositions rpt, ctl, , InchToTwip(0.25)
    Set ctl = CreateSubformControl(rpt, "srptProductReportProducts", "subProductReportProducts", , "ProductID")
    
    ''srptProductReports
    Set ctl = CreateSubformControl(rpt, "srptProductReports", "subProductReports", , "ProductID")
    Offset_ctlPositions rpt, ctl, , InchToTwip(0.25)
    
    ''QUARANTINE --> srptQuarantinedProductReports
    Set ctl = CreateLabelControl(rpt, "QUARANTINE", "QUARANTINE", "Heading3")
    Offset_ctlPositions rpt, ctl, , InchToTwip(0.25)
    Set ctl = CreateSubformControl(rpt, "srptQuarantinedProductReports", "subQuarantinedProductReports", , "ProductID")
    
    CreateBannerControls rpt
    
    CleanUpReportProperties rpt
    
    RenameFormOrReport rpt.Name, REPORT_NAME
    
    GetFormOrReport REPORT_NAME, True, True
    
End Sub

Private Sub Set_fltrProduct_RowSource(frm As Form)
    
    Dim fltrInStock: fltrInStock = frm("fltrInStock")
    Dim filterStr: filterStr = "ProductID > 0"
    If fltrInStock Then
        filterStr = "QtyOnStock > 0 AND NOT QtyOnStock IS NULL"
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM qryProducts WHERE " & filterStr & " ORDER BY SESEMSProductNumber"
    
    frm("fltrProduct").RowSource = sqlStr
    
    Dim fltrProduct: fltrProduct = frm("fltrProduct")
    If isFalse(fltrProduct) Then Exit Sub
    If Not isPresent("qryProducts", filterStr & " AND ProductID = " & fltrProduct) Then
        frm("fltrProduct") = Null
    End If
    
End Sub

Public Function frmProductReports_fltrInStock_AfterUpdate(frm As Form)
    
    Set_fltrProduct_RowSource frm
    
End Function


Public Function frmProductReports_fltrCustomer_AfterUpdate(frm As Form)
    
    fltrCustomerProdNumber_SetRowSource frm

End Function

Private Function fltrCustomerProdNumber_SetRowSource(frm As Form)
    
    Dim fltrCustomer: fltrCustomer = frm("fltrCustomer")
    Dim filterStr: filterStr = "ProductID = 0"
    
    If Not isFalse(fltrCustomer) Then
        filterStr = "CustomerID = " & fltrCustomer
    End If

    Dim sqlStr: sqlStr = "SELECT ProductID,CustomerProdNumber FROM qryProducts WHERE " & filterStr & " ORDER BY CustomerProdNumber"
    frm("fltrCustomerProdNumber").RowSource = sqlStr
    frm("fltrCustomerProdNumber").Requery
    frm("fltrCustomerProdNumber") = Null
    
End Function

Public Function frmProductReports_OnLoad(frm As Form)
    
    ResetOtherFilters frm, ""
    ''Filter_subProducts frm, Null
    SyncProductIDWith_tblProductReports Null
    Set_subform_RecordSource frm
    
End Function

Public Sub SetUp_frmProductReports()
    
    Dim frmName: frmName = "frmProductReports"
    
    If Not IsFormOpen(frmName) Then
        DoCmd.OpenForm frmName, acDesign
    End If
    
    Dim frm As Form: Set frm = Forms(frmName)
    
    CreateBannerControls frm
    
    frm("subCustomerOrderReportSubSuppliers").Form.recordSource = "SELECT * FROM tblProductReports WHERE NOT WarehousePlace IS NULL"
    frm("subQuarantinedProductReports").Form.recordSource = "SELECT * FROM tblProductReports WHERE WarehousePlace IS NULL"
    
End Sub

Private Function ResetOtherFilters(frm As Form, controlName)
    
    Dim ctl As control, ctlNames As New clsArray: ctlNames.arr = "fltrProduct,fltrCustomer,fltrCustomerProdNumber"
    For Each ctl In frm.Controls
        If ctl.Name Like "fltr*" And ctl.Name <> controlName Then
            ctl = Null
        End If
    Next ctl
    
    
End Function

Public Function frmProductReports_fltrProduct_AfterUpdate(frm As Form)

    Dim fltrProduct: fltrProduct = frm("fltrProduct")
    ''Filter_subProducts frm, fltrProduct
    ResetOtherFilters frm, "fltrProduct"
    SyncProductIDWith_tblProductReports fltrProduct
    Set_subform_RecordSource frm
    
End Function


Public Function frmProductReports_fltrCustomerProdNumber_AfterUpdate(frm As Form)

    Dim fltrCustomerProdNumber: fltrCustomerProdNumber = frm("fltrCustomerProdNumber")
    ''Filter_subProducts frm, fltrCustomerProdNumber
    ''ResetOtherFilters frm, "fltrCustomerProdNumber"
    
    frm("fltrProduct") = fltrCustomerProdNumber
    SyncProductIDWith_tblProductReports fltrCustomerProdNumber
    Set_subform_RecordSource frm
    
End Function

'Private Function Filter_subProducts(frm As Form, filterValue, Optional mode = "Product")
'
'    Dim filterStr: filterStr = "ProductID = 0"
'    If Not isFalse(filterValue) Then
'        If mode = "Product" Then
'            filterStr = "ProductID = " & filterValue
'        Else
'            filterStr = "CustomerID = " & filterValue
'        End If
'
'    End If
'
'    frm("subProducts").Form.Filter = filterStr
'    frm("subProducts").Form.FilterOn = True
'
'End Function

Private Function RequerySubforms(frm As Form)

    ''subCustomerOrderReportSubSuppliers, subQuarantinedProductReports
    frm("subCustomerOrderReportSubSuppliers").Form.Requery
    frm("subQuarantinedProductReports").Form.Requery
    
End Function

Public Function SyncProductIDWith_tblProductReports(ProductID)
    
    RunSQL "DELETE FROM tblProductReports"
    
    If isFalse(ProductID) Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM qryOrderAssignments WHERE ProductID = " & ProductID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim fields As New clsArray: fields.arr = "WarehousePlace,DeliveryDate,ProductSupplier,QtyOnStock"
    Dim fieldValues As New clsArray
    Do Until rs.EOF
        
        Dim ActualQuantity: ActualQuantity = rs.fields("ActualQuantity")
        Dim TransferredOutQty: TransferredOutQty = Coalesce(rs.fields("TransferredOutQty"), 0)
        Dim ScrapQty: ScrapQty = Coalesce(rs.fields("ScrapQty"), 0)
        Dim WHTQty: WHTQty = Coalesce(rs.fields("WHTQty"), 0)
        Dim DCQty: DCQty = Coalesce(rs.fields("DCQty"), 0)
        Dim SubDeliveryDate: SubDeliveryDate = rs.fields("SubDeliveryDate")
        Dim SupplierShortName: SupplierShortName = rs.fields("SupplierShortName")
        Dim WarehousePlace: WarehousePlace = rs.fields("WarehousePlace")
        Dim WHTWarehousePlace: WHTWarehousePlace = rs.fields("WHTWarehousePlace")
        Dim QualityControlStatus: QualityControlStatus = rs.fields("QualityControlStatus")
        
        If isFalse(QualityControlStatus) Then
            GoTo NextRecord:
        End If
        
        Dim RemainingQTY: RemainingQTY = 0
        If WHTQty = 0 Then
            RemainingQTY = ActualQuantity - TransferredOutQty - ScrapQty - DCQty
            If RemainingQTY > 0 Then
                Set fieldValues = New clsArray
                fieldValues.Add WarehousePlace
                fieldValues.Add SubDeliveryDate
                fieldValues.Add SupplierShortName
                fieldValues.Add RemainingQTY
                UpsertRecord "tblProductReports", fields, fieldValues
            End If
        Else
            ''There's a transfer here to the WarehousePlace
            RemainingQTY = ActualQuantity - WHTQty - ScrapQty - DCQty
            If RemainingQTY > 0 Then
                Set fieldValues = New clsArray
                fieldValues.Add WarehousePlace
                fieldValues.Add SubDeliveryDate
                fieldValues.Add SupplierShortName
                fieldValues.Add RemainingQTY
                UpsertRecord "tblProductReports", fields, fieldValues
            End If
            
            RemainingQTY = WHTQty - TransferredOutQty - ScrapQty - DCQty
            If RemainingQTY > 0 Then
                Set fieldValues = New clsArray
                fieldValues.Add WHTWarehousePlace
                fieldValues.Add SubDeliveryDate
                fieldValues.Add SupplierShortName
                fieldValues.Add RemainingQTY
                UpsertRecord "tblProductReports", fields, fieldValues
            End If
            
        End If
        
        
NextRecord:
        rs.MoveNext
    Loop
    
End Function

