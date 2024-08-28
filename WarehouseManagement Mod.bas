Attribute VB_Name = "WarehouseManagement Mod"
Option Compare Database
Option Explicit

Public Function WarehouseManagementCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            frm.OnCurrent = "=frmWarehouseManagement_OnCurrent([Form])"
            frm.OnLoad = "=frmWarehouseManagement_OnLoad([Form])"
            frm("fltrOrderStatus").AfterUpdate = "=frmWarehouseManagement_fltrOrderStatus_AfterUpdate([Form])"
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function frmWarehouseManagement_fltrSESEMSProductNr_AfterUpdate(frm As Form)
    
    Set_fltrCommissionNumber_RowSource frm
    frmWarehouseManagement_fltrCommissionNumber_AfterUpdate frm
    
End Function

Private Function Set_fltrCustomerOrderNumber_RowSource(frm As Form)
    
    Dim fltrShortName: fltrShortName = frm("fltrShortName")
    Dim fltrOrderStatus: fltrOrderStatus = frm("fltrOrderStatus")
    
    Dim filterArr As New clsArray
    
    If Not isFalse(fltrShortName) Then
        filterArr.Add "CustomerID = " & fltrShortName
    End If
    
    If fltrOrderStatus = "Open" Then
        filterArr.Add "OrderStatus = ""Open"""
    End If
    
    Dim filterStr: filterStr = "CustomerOrderID > 0"
    If filterArr.count > 0 Then
        filterStr = filterArr.JoinArr(" AND ")
    End If
    
    Dim sqlStr: sqlStr = "SELECT CustomerOrderMainID,CustomerOrderNumber FROM qryCustomerOrders WHERE " & filterStr & _
        " GROUP BY CustomerOrderMainID,CustomerOrderNumber ORDER BY CustomerOrderNumber"
    frm("fltrCustomerOrderNumber").RowSource = sqlStr
    
End Function

Private Function Set_fltrCustomerProdNumber_RowSource(frm As Form)
    
    Dim fltrShortName: fltrShortName = frm("fltrShortName")
    Dim fltrOrderStatus: fltrOrderStatus = frm("fltrOrderStatus")
    
    Dim filterArr As New clsArray
    
    If Not isFalse(fltrShortName) Then
        filterArr.Add "CustomerID = " & fltrShortName
    End If
    
    If fltrOrderStatus = "Open" Then
        filterArr.Add "OrderStatus = ""Open"""
    End If
    
    Dim filterStr: filterStr = "CustomerOrderID > 0"
    If filterArr.count > 0 Then
        filterStr = filterArr.JoinArr(" AND ")
    End If
    
    Dim sqlStr: sqlStr = "SELECT CustomerOrderID,CustomerProdNumber FROM qryCustomerOrders WHERE " & filterStr & _
        " ORDER BY CustomerProdNumber"
    frm("fltrCustomerProdNumber").RowSource = sqlStr
    
End Function

Private Function Set_fltrCommissionNumber_RowSource(frm As Form)
    
    Dim fltrSESEMSProductNr: fltrSESEMSProductNr = frm("fltrSESEMSProductNr")
    Dim fltrOrderStatus: fltrOrderStatus = frm("fltrOrderStatus")
    Dim fltrCustomerOrderNumber: fltrCustomerOrderNumber = frm("fltrCustomerOrderNumber")
    Dim fltrShortName: fltrShortName = frm("fltrShortName")
    
    Dim filterArr As New clsArray
    
    If Not isFalse(fltrSESEMSProductNr) Then
        filterArr.Add "ProductID = " & fltrSESEMSProductNr
    End If
    
    If Not isFalse(fltrShortName) Then
        filterArr.Add "CustomerID = " & fltrShortName
    End If
    
    If Not isFalse(fltrCustomerOrderNumber) Then
        filterArr.Add "CustomerOrderMainID = " & fltrCustomerOrderNumber
    End If
    
    If fltrOrderStatus = "Open" Then
        filterArr.Add "OrderStatus = ""Open"""
    End If
    
    Dim filterStr: filterStr = "CustomerOrderID > 0"
    If filterArr.count > 0 Then
        filterStr = filterArr.JoinArr(" AND ")
    End If
    
    Dim sqlStr: sqlStr = "SELECT CustomerOrderID,CustomerOrderID2 FROM qryCustomerOrders WHERE " & filterStr & _
        " ORDER BY CustomerOrderID2"
    frm("fltrCommissionNumber").RowSource = sqlStr
    
    Dim fltrCommissionNumber: fltrCommissionNumber = frm("fltrCommissionNumber")
    
    If Not isFalse(fltrCommissionNumber) Then
        filterArr.Add "CustomerOrderID = " & fltrCommissionNumber
        If Not isPresent("qryCustomerOrders", filterArr.JoinArr(" AND ")) Then
            frm("fltrCommissionNumber") = Null
        End If
    End If
    
End Function

Public Sub frmWarehouseManagement_Fix_controls()

    DoCmd.OpenForm "frmWarehouseManagement", acDesign
    
    Dim frm As Form: Set frm = Forms("frmWarehouseManagement")
    
    Dim baseSubForm As control: Set baseSubForm = frm("subTempDeliveryToCustomers")
    
    frm("subTempWarehouseTransfers").Top = baseSubForm.Top
    frm("subTempWarehouseTransfers").Left = baseSubForm.Left
    frm("subTempWarehouseTransfers").Width = baseSubForm.Width
    
End Sub

Public Function GetWarehouseManagementStatus(frm As Form) As String
    
    If frm.NewRecord Then Exit Function
    
    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    
    Dim StillToBeDelivered: StillToBeDelivered = frm("ScrappableQty")
    
    ''tblOrderAssignments without QualityControlStatus
    If Not isFalse(CustomerOrderID) Then
    
        Dim FirstOrderAssignmentID: FirstOrderAssignmentID = ELookup("tblOrderAssignments", "CustomerOrderID = " & _
        CustomerOrderID, "OrderAssignmentID", "OrderAssignmentOrder ASC, OrderAssignmentID ASC")
        
        Dim andClause As New clsArray
        andClause.Add "(Not isFalse(OutToSubsupplierDate) AND IsFalse(QualityControlStatus))"
        If Not isFalse(FirstOrderAssignmentID) Then
            andClause.Add "(isFalse(OutToSubsupplierDate) AND OrderAssignmentID = " & FirstOrderAssignmentID & "  AND IsFalse(QualityControlStatus))"
        End If
        
        
        Dim filterStr: filterStr = "CustomerOrderID = " & _
            CustomerOrderID & " AND (" & andClause.JoinArr(" OR ") & ")"

        Dim UnprocessedOrderAssignment: UnprocessedOrderAssignment = ELookup("tblOrderAssignments", filterStr, "ActualQuantity")
        StillToBeDelivered = StillToBeDelivered + Coalesce(UnprocessedOrderAssignment, 0)
        
    End If
    
    Dim OrderStatus
    
    If StillToBeDelivered > 0 Then
        GetWarehouseManagementStatus = "STILL OPEN " & StillToBeDelivered & " PIECES"
        OrderStatus = "Open"
    Else
        GetWarehouseManagementStatus = "CLOSED"
        OrderStatus = "Closed"
    End If
    
'    If OrderStatus = "Closed" Then
'        ''Make sure there's no pending
'        Dim sqlStr: sqlStr = "SELECT * FROM qryOrderAssignments WHERE CustomerOrderID = " & CustomerOrderID & " AND QualityControlStatus IS NULL"
'        Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
'        If Not rs.EOF Then
'            Dim ActualQuantity: ActualQuantity = rs.fields("ActualQuantity")
'            GetWarehouseManagementStatus = "STILL OPEN " & StillToBeDelivered & " PIECES"
'            OrderStatus = "Open"
'        End If
'    End If
    
    If Not isFalse(CustomerOrderID) Then
    
        Dim isForceClosed: isForceClosed = isPresent("tblCustomerOrders", "CustomerOrderID = " & CustomerOrderID & " AND FCCheck")
        If isForceClosed Then
            GetWarehouseManagementStatus = "FORCE CLOSED: " & StillToBeDelivered & " PIECES"
            OrderStatus = "Closed"
        End If
        
        RunSQL "UPDATE tblCustomerOrders SET OrderStatus = " & Esc(OrderStatus) & " WHERE CustomerOrderID = " & CustomerOrderID & " AND NOT FCCheck"
        
        Dim CustomerOrderMainID: CustomerOrderMainID = ELookup("tblCustomerOrders", "CustomerOrderID = " & CustomerOrderID, "CustomerOrderMainID")
        
        Dim AllClosed: AllClosed = Not isPresent("tblCustomerOrders", "CustomerOrderMainID = " & CustomerOrderMainID & _
            " AND OrderStatus = ""OPEN""")
        
        If AllClosed And OrderStatus = "Closed" Then
            RunSQL "UPDATE tblCustomerOrderMains SET OrderMainStatus = ""Closed"" WHERE CustomerOrderMainID = " & CustomerOrderMainID
        Else
            RunSQL "UPDATE tblCustomerOrderMains SET OrderMainStatus = ""Open"" WHERE CustomerOrderMainID = " & CustomerOrderMainID
        End If
    
    End If
    
    Dim ProductID: ProductID = frm("ProductID")
    Update_tblProducts_QtyOnStock ProductID
    frmCustomerOrders_Racalculate
    frmCustomerOrderReports_Racalculate
    frmOrderAssignmentsWithMaterialMain_Recalculate
    
'    Dim ManualOrderStatus: ManualOrderStatus = frm("ManualOrderStatus")
'    If isFalse(ManualOrderStatus) Then
'        frm("OrderStatus") = OrderStatus
'    End If
    
End Function

Public Function frmWarehouseManagement_SyncOtherTabs(TabIndex)

    If IsFormOpen("frmWarehouseManagement") Then
        Dim frm As Form:  Set frm = Forms("frmWarehouseManagement")
        Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
        
        If isFalse(CustomerOrderID) Then Exit Function
        
        Dim SubformName, SubformNames As New clsArray
        SubformNames.Add "subTempWarehouseTransferToSubSuppliers"
        SubformNames.Add "subTempWarehouseTransfers"
        SubformNames.Add "subTempWarehouseScraps"
        SubformNames.Add "subTempDeliveryToCustomers"
        
        Dim callbacks As New clsArray
        callbacks.Add "Sync_TempWarehouseTransferToSubSupplier"
        callbacks.Add "Sync_tblTempWarehouseTransfers_2"
        callbacks.Add "Sync_TempWarehouseScraps"
        callbacks.Add "Sync_tblTempDeliveryToCustomers_2"
        
        Dim i As Integer: i = 0
        For Each SubformName In SubformNames.arr
            If TabIndex <> i Then
                Run callbacks.arr(i), frm, CustomerOrderID
            End If
            i = i + 1
        Next SubformName
    End If
    
End Function

Public Function SyncAllWarehouseManagementSubTables(frm As Form, CustomerOrderID)
    
    Sync_tblTempDeliveryToCustomers_2 frm, CustomerOrderID
    Sync_tblTempWarehouseTransfers_2 frm, CustomerOrderID
    ''Sync_tblReleaseFromQuarantines frm, CustomerOrderID
    Sync_TempWarehouseTransferToSubSupplier frm, CustomerOrderID
    Sync_TempWarehouseScraps frm, CustomerOrderID
    
End Function

Public Function Sync_tblTempWarehouseTransfers_2(frm As Form, CustomerOrderID)
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblOrderAssignments WHERE CustomerOrderID = " & CustomerOrderID & _
        " AND Not isFalse(QualityControlStatus) ORDER " & _
        " BY OrderAssignmentOrder ASC, OrderAssignmentID ASC"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim fields As New clsArray: fields.arr = "OrderAssignmentID,PCSToTransfer,WarehousePlaceID,AvailablePCS," & _
        "TargetWarehousePlaceID,IsChecked,DescriptionOfRelease"
   
    Dim fieldValues As New clsArray
    
    RunSQL "DELETE FROM tblTempWarehouseTransfers"
    
    Do Until rs.EOF
        Dim OrderAssignmentID: OrderAssignmentID = rs.fields("OrderAssignmentID")
        Dim ActualQuantity: ActualQuantity = rs.fields("ActualQuantity")
        Dim TransferredOutDate: TransferredOutDate = rs.fields("TransferredOutDate")
        Dim WarehousePlaceID: WarehousePlaceID = rs.fields("WarehousePlaceID")
        Dim WHTWarehousePlaceID: WHTWarehousePlaceID = rs.fields("WHTWarehousePlaceID")
        Dim WHTQty: WHTQty = Coalesce(rs.fields("WHTQty"), 0)
        Dim WHTDescriptionOfRelease: WHTDescriptionOfRelease = rs.fields("WHTDescriptionOfRelease")
        Dim WHTConfirmation: WHTConfirmation = rs.fields("WHTConfirmation")
        Dim DCConfirmation: DCConfirmation = rs.fields("DCConfirmation")
        Dim ScrapConfirmation: ScrapConfirmation = rs.fields("ScrapConfirmation")
        Dim ScrapQty: ScrapQty = Coalesce(rs.fields("ScrapQty"), 0)
        Dim DCQty: DCQty = Coalesce(rs.fields("DCQty"), 0)
        Dim TransferredOutQty: TransferredOutQty = Coalesce(rs.fields("TransferredOutQty"), 0)
        
'        If ScrapConfirmation Then
'            GoTo NextRecord
'        End If
        
'        If DCConfirmation Then
'            GoTo NextRecord
'        End If
        
'        If Not isFalse(TransferredOutDate) And Not WHTConfirmation Then
'            GoTo NextRecord
'        End If
    
        Set fieldValues = New clsArray
        fieldValues.Add OrderAssignmentID
        fieldValues.Add IIf(WHTConfirmation, WHTQty, ActualQuantity - DCQty - TransferredOutQty)
        fieldValues.Add WarehousePlaceID
        fieldValues.Add ActualQuantity - DCQty - TransferredOutQty + WHTQty
        fieldValues.Add WHTWarehousePlaceID
        fieldValues.Add WHTConfirmation
        fieldValues.Add WHTDescriptionOfRelease
        
        UpsertRecord "tblTempWarehouseTransfers", fields, fieldValues
NextRecord:
        rs.MoveNext
    Loop
    
    frm("subTempWarehouseTransfers").Form.Requery
    
End Function

Public Function Sync_TempWarehouseScraps(frm As Form, CustomerOrderID)
    
    ''get tll transferrable products
    Dim sqlStr: sqlStr = "SELECT * FROM tblOrderAssignments WHERE CustomerOrderID = " & CustomerOrderID & _
        " AND Not isFalse(QualityControlStatus) ORDER " & _
        " BY OrderAssignmentOrder ASC, OrderAssignmentID ASC"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim fields As New clsArray: fields.arr = "OrderAssignmentID,QtyToScrap,WarehousePlaceID,Reason,IsChecked"
    ''TO DO: Adjustment of QTYToScrap and Reason and isChecked field based on if there is a recorded scrap
    ''Adjust QTYToScrap - transferred to subsupplier
     
    Dim fieldValues As New clsArray
    
    RunSQL "DELETE FROM tblTempWarehouseScraps"
    
    Do Until rs.EOF
        Dim OrderAssignmentID: OrderAssignmentID = rs.fields("OrderAssignmentID")
        Dim ActualQuantity: ActualQuantity = rs.fields("ActualQuantity")
        Dim TransferredOutDate: TransferredOutDate = rs.fields("TransferredOutDate")
        Dim TransferredOutQty: TransferredOutQty = Coalesce(rs.fields("TransferredOutQty"), 0)
        Dim WarehousePlaceID: WarehousePlaceID = rs.fields("WarehousePlaceID")
        Dim ReasonForScrap: ReasonForScrap = rs.fields("ReasonForScrap")
        Dim ScrapQty: ScrapQty = Coalesce(rs.fields("ScrapQty"), 0)
        Dim ScrapConfirmation: ScrapConfirmation = rs.fields("ScrapConfirmation")
        Dim DCQty: DCQty = Coalesce(rs.fields("DCQty"), 0)
        
        ActualQuantity = ActualQuantity - TransferredOutQty - DCQty
    
        Set fieldValues = New clsArray
        fieldValues.Add OrderAssignmentID
        fieldValues.Add ActualQuantity
        fieldValues.Add WarehousePlaceID
        fieldValues.Add ReasonForScrap
        fieldValues.Add ScrapConfirmation
        
        UpsertRecord "tblTempWarehouseScraps", fields, fieldValues
NextRecord:
        rs.MoveNext
    Loop
    
    frm("subTempWarehouseScraps").Form.Requery
    
End Function

Public Function Sync_TempWarehouseTransferToSubSupplier(frm As Form, CustomerOrderID)
    
'    Dim LastOrderAssignmentID: LastOrderAssignmentID = ELookup("tblOrderAssignments", "CustomerOrderID = " & _
'        CustomerOrderID, "OrderAssignmentID", "OrderAssignmentOrder DESC, OrderAssignmentID DESC")
    
    Dim filterStr: filterStr = "CustomerOrderID = " & CustomerOrderID
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblOrderAssignments WHERE " & filterStr & " ORDER BY OrderAssignmentOrder ASC, OrderAssignmentID ASC"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim fieldNames As New clsArray: fieldNames.arr = "OrderAssignmentID,WarehousePlaceID,TransferredOutQty"
    
    Dim fieldValues As New clsArray
     
    ''Missing TransferredOutToID,TransferredOutDate,TransferredOutQty
    RunSQL "DELETE FROM tblTempWarehouseTransferToSubSuppliers"
    Do Until rs.EOF
        
        Dim DCQty: DCQty = rs.fields("DCQty")
        Dim WHTQty: WHTQty = rs.fields("WHTQty")
        Dim QualityControlStatus: QualityControlStatus = rs.fields("QualityControlStatus")
        If isFalse(QualityControlStatus) Then
            GoTo NextRecord
        End If
        
        Dim ScrapConfirmation: ScrapConfirmation = rs.fields("ScrapConfirmation")
        If ScrapConfirmation Then
            Dim ActualQuantity: ActualQuantity = rs.fields("ActualQuantity")
            Dim ScrapQty: ScrapQty = rs.fields("ScrapQty")
            If ActualQuantity = ScrapQty Then
                GoTo NextRecord
            End If
        End If
        
        fieldNames.arr = "OrderAssignmentID,WarehousePlaceID"
        Set fieldValues = New clsArray
        Dim OrderAssignmentID: OrderAssignmentID = rs.fields("OrderAssignmentID")
        
        fieldValues.Add OrderAssignmentID
        fieldValues.Add rs.fields("WarehousePlaceID")
        
        Dim rs2 As Recordset: Set rs2 = GetNextOrderAssignment(OrderAssignmentID)
        If Not rs2.EOF Then
            Dim OutToSubsupplierDate: OutToSubsupplierDate = rs2.fields("OutToSubsupplierDate")
            If Not isFalse(OutToSubsupplierDate) Then
                fieldNames.Add "TransferredOutToID"
                fieldNames.Add "TransferredOutDate"
                fieldNames.Add "TransferredOutQty"
                fieldValues.Add rs2.fields("OrderAssignmentID")
                fieldValues.Add OutToSubsupplierDate
                fieldValues.Add rs2.fields("ActualQuantity")
            Else
                
                fieldNames.Add "TransferredOutQty"
                fieldValues.Add Coalesce(WHTQty, rs.fields("ActualQuantity"))
                
            End If
        Else
            fieldNames.Add "TransferredOutQty"
            fieldNames.Add "TransferredOutDate"
            fieldValues.Add Coalesce(DCQty, WHTQty, rs.fields("ActualQuantity"))
            fieldValues.Add Coalesce(GetDeliveryDate(frm))
        End If
        
        UpsertRecord "tblTempWarehouseTransferToSubSuppliers", fieldNames, fieldValues
NextRecord:
        rs.MoveNext
    Loop
    
    frm("subTempWarehouseTransferToSubSuppliers").Form.Requery
    
End Function

Public Function contOrderAssignments_Recalculate(CustomerOrderID)
    
    If Not IsFormOpen("frmWarehouseManagement") Then
        DoCmd.OpenForm "frmWarehouseManagement", , , , , acHidden
    End If
    
    Dim frm As Form: Set frm = Forms("frmWarehouseManagement")
    frm("fltrCommissionNumber") = CustomerOrderID
    
    frmWarehouseManagement_fltrCommissionNumber_AfterUpdate frm
    
End Function

Public Function frmWarehouseManagement_RunWorkflow(frm As Form, CustomerOrderID)
    
    frm("CustomerOrderID") = CustomerOrderID
    
    CustomerOrderID = IIf(isFalse(CustomerOrderID), 0, CustomerOrderID)
    Set_subCustomerOrder_Recordsource frm, CustomerOrderID
    SyncAllWarehouseManagementSubTables frm, CustomerOrderID
    
End Function

Public Function frmWarehouseManagement_fltrCommissionNumber_AfterUpdate(frm As Form)

    Dim CustomerOrderID: CustomerOrderID = frm("fltrCommissionNumber")
    frmWarehouseManagement_RunWorkflow frm, CustomerOrderID
    
End Function

Public Function Is_tblOrderAssignmentAdditionAllowed(frm As Form)
    
    Const ModelName = "DeliveryToCustomers"
    Const tempTable = "tblTemp" & ModelName
    Const TableName = "tbl" & ModelName
    Const SubformName = "subTemp" & ModelName
    
    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    ''Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    If isFalse(CustomerOrderID) Then Exit Function
    
    Dim QuarantinedQty: QuarantinedQty = ESum2("qryOrderAssignments", "QualityControlStatus = ""Stop"" AND CustomerOrderID = " & CustomerOrderID & _
        " AND OutToSubsupplierDate IS NULL", "ActualQuantity")
    
    Dim ReleasedFromQuarantine: ReleasedFromQuarantine = ELookup("tblReleaseFromQuarantines", "CustomerOrderID = " & CustomerOrderID, "PCSToTransfer")
    If ReleasedFromQuarantine = "" Then ReleasedFromQuarantine = 0

    Dim RemainingQuarantinedQty: RemainingQuarantinedQty = QuarantinedQty - ReleasedFromQuarantine
    
    Is_tblOrderAssignmentAdditionAllowed = RemainingQuarantinedQty = True
    
End Function

Public Function Set_tblOrderAssignments_QualityControlCaption(frm As Form)
    
    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    If isFalse(CustomerOrderID) Then
        Exit Function
    End If
    
    Dim QualityControlCaptions As New clsArray, QTY
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM qryTempDeliveryToCustomers")
    
    Do Until rs.EOF
        Dim WarehousePlace: WarehousePlace = rs.fields("WarehousePlace")
        Dim AvailablePCs: AvailablePCs = rs.fields("AvailablePCS")
        Dim PCSToDeliver: PCSToDeliver = rs.fields("PCSToDeliver")
        Dim IsChecked: IsChecked = rs.fields("IsChecked")
        
        QTY = AvailablePCs
        If IsChecked Then QTY = QTY - PCSToDeliver
        
        If QTY <> 0 Then QualityControlCaptions.Add WarehousePlace & ": " & Format$(QTY, "Standard")
        
        rs.MoveNext
    Loop
    
    Dim QuarantinedQty: QuarantinedQty = frm("QuarantinedQty")
    Dim txtReleasedFromQuarantine: txtReleasedFromQuarantine = frm("txtReleasedFromQuarantine")
    
    Dim RemainingQTY: RemainingQTY = QuarantinedQty - txtReleasedFromQuarantine
    If RemainingQTY <> 0 Then QualityControlCaptions.Add "Quarantine: " & Format$(RemainingQTY, "Standard")
    
    ''Find the last tblOrderAssignments (OrderAssignmentID) Where CustomerOrderID AND OutToSubsupplierDate IS NULL
    Dim OrderAssignmentID: OrderAssignmentID = ELookup("tblOrderAssignments", "CustomerOrderID = " & CustomerOrderID & _
        " AND Not IsFalse(QualityControlStatus)", "OrderAssignmentID", "OrderAssignmentOrder DESC,OrderAssignmentID DESC")
        
    ''Remove all the captions first
    RunSQL "UPDATE tblOrderAssignments SET QualityControlCaption = Null WHERE CustomerOrderID = " & CustomerOrderID
    
    If Not isFalse(OrderAssignmentID) Then
        Dim rs2 As Recordset: Set rs2 = GetNextOrderAssignment(OrderAssignmentID)
        If rs2.EOF Then
            GoTo DoTheUpdate:
        End If
        ''Check if the QualityControlStatus IS NOT false and OutToSubsupplierDate IS also NOT NULL
        ''If Not isFalse(rs2.fields("QualityControlStatus")) And Not isFalse(rs2.fields("OutToSubsupplierDate")) Then
        If Not isFalse(rs2.fields("OutToSubsupplierDate")) Then
            GoTo EndTheFunction:
        End If
    End If
    
DoTheUpdate:
    If Not isFalse(OrderAssignmentID) And QualityControlCaptions.count > 0 Then
        RunSQL "UPDATE tblOrderAssignments SET QualityControlCaption = " & Esc(QualityControlCaptions.JoinArr(" | ")) & _
            " WHERE OrderAssignmentID = " & OrderAssignmentID
    End If
    
EndTheFunction:
    If IsFormOpen("frmSubSupplierManagementMain") Then
        Forms("frmSubSupplierManagementMain")("subform").Form("subOrderAssignments").Form.Requery
    End If
    
End Function

Public Function GetNextOrderAssignment(OrderAssignmentID) As Recordset
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblOrderAssignments WHERE OrderAssignmentID = " & OrderAssignmentID)
    
    Dim CustomerOrderID: CustomerOrderID = rs.fields("CustomerOrderID")
    Dim OrderAssignmentOrder: OrderAssignmentOrder = EscapeString(rs.fields("OrderAssignmentOrder"), "tblOrderAssignments", "OrderAssignmentOrder")
    
    Set rs = ReturnRecordset("SELECT * FROM tblOrderAssignments WHERE CustomerOrderID = " & CustomerOrderID & " AND ((OrderAssignmentOrder = " & _
        OrderAssignmentOrder & " AND OrderAssignmentID > " & OrderAssignmentID & ") OR OrderAssignmentOrder > " & OrderAssignmentOrder & ") ORDER BY " & _
        "OrderAssignmentOrder, OrderAssignmentID")
        
    Set GetNextOrderAssignment = rs
    
End Function

Public Function GetPreviousOrderAssignment(OrderAssignmentID) As Recordset
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblOrderAssignments WHERE OrderAssignmentID = " & OrderAssignmentID)
    
    If rs.EOF Then Exit Function
    
    Dim CustomerOrderID: CustomerOrderID = rs.fields("CustomerOrderID")
    Dim OrderAssignmentOrder: OrderAssignmentOrder = EscapeString(rs.fields("OrderAssignmentOrder"), "tblOrderAssignments", "OrderAssignmentOrder")
    
    Set rs = ReturnRecordset("SELECT * FROM tblOrderAssignments WHERE CustomerOrderID = " & CustomerOrderID & " AND ((OrderAssignmentOrder = " & _
        OrderAssignmentOrder & " AND OrderAssignmentID < " & OrderAssignmentID & ") OR OrderAssignmentOrder < " & OrderAssignmentOrder & ") ORDER BY " & _
        "OrderAssignmentOrder DESC , OrderAssignmentID DESC")
        
    Set GetPreviousOrderAssignment = rs
End Function

Public Function frmWarehouseManagement_fltrCustomerDueDate_AfterUpdate(frm As Form)
    
    Dim CustomerOrderID: CustomerOrderID = frm("fltrCustomerDueDate")
    Set_subCustomerOrder_Recordsource frm, CustomerOrderID
    SyncAllWarehouseManagementSubTables frm, CustomerOrderID

End Function

Public Function frmWarehouseManagement_fltrCustomerOrderNumber_AfterUpdate(frm As Form)
    
    Set_fltrCommissionNumber_RowSource frm
    frmWarehouseManagement_fltrCommissionNumber_AfterUpdate frm
    
End Function

Public Function frmWarehouseManagement_fltrCustomerProdNumber_AfterUpdate(frm As Form)
    
    Dim CustomerOrderID: CustomerOrderID = frm("fltrCustomerProdNumber")
    frmWarehouseManagement_RunWorkflow frm, CustomerOrderID
    
End Function

Public Function frmWarehouseManagement_fltrOrderStatus_AfterUpdate(frm As Form)
    
    Set_fltrCommissionNumber_RowSource frm
    Set_fltrCustomerOrderNumber_RowSource frm
    Set_fltrCustomerProdNumber_RowSource frm
    
End Function

Private Sub FindCustomerOrderBasedOnFiltersOtherThanCommissionNumber(frm As Form)

    ''fltrOrderStatus,fltrShortName,fltrCustomerOrderNumber,fltrCustomerDueDate
    Dim sqlStr: sqlStr = "SELECT * FROM qryCustomerOrders ORDER BY CustomerOrderID"
    
    Dim fltrOrderStatus: fltrOrderStatus = frm("fltrOrderStatus")
    Dim fltrShortName: fltrShortName = frm("fltrShortName")
    Dim fltrCustomerOrderNumber: fltrCustomerOrderNumber = frm("fltrCustomerOrderNumber")
    Dim fltrCustomerDueDate: fltrCustomerDueDate = frm("fltrCustomerDueDate")
    
    Dim filterArr As New clsArray
    
    If fltrOrderStatus = "Open" Then
        filterArr.Add "OrderStatus = ""Open"""
    ElseIf fltrOrderStatus = "All" Then
        filterArr.Add "CustomerOrderID > 0"
    End If
    
    If Not isFalse(fltrShortName) Then
        filterArr.Add "CustomerID = " & fltrShortName
    End If
    
    If Not isFalse(fltrCustomerOrderNumber) Then
        filterArr.Add "CustomerOrderNumber = " & Esc(fltrCustomerOrderNumber)
    End If
    
    If Not isFalse(fltrCustomerDueDate) Then
        filterArr.Add "LastAgreedDueDate = #" & fltrCustomerDueDate & "#"
    End If
    
    If filterArr.count > 0 Then
        frm("fltrCommissionNumber") = Null
        Set frm = frm("subCustomerOrder").Form
        frm.recordSource = sqlStr
        FindFirst frm, filterArr.JoinArr(" AND ")
    Else
        Set_subCustomerOrder_Recordsource frm, 0
    End If
    
End Sub

Public Function frmWarehouseManagement_fltrShortName_AfterUpdate(frm As Form)
    
    Set_fltrCustomerOrderNumber_RowSource frm
    Set_fltrCustomerProdNumber_RowSource frm
    Set_fltrCommissionNumber_RowSource frm
        
    frmWarehouseManagement_fltrCommissionNumber_AfterUpdate frm
    
End Function

Public Function frmWarehouseManagement_OnCurrent(frm As Form)
    
    ''Sync_tblTempDeliveryToCustomers frm
    ''Sync_tblTempWarehouseTransfers frm
    ''Sync_tblReleaseFromQuarantines frm
    ''fltrCommissionNumber
    ''fltrCustomerOrderNumber
    ''fltrCustomerDueDate

End Function

'Public Sub Sync_tblReleaseFromQuarantines(frm As Form, CustomerOrderID)
'
'    If isFalse(CustomerOrderID) Then Exit Sub
'
'    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblReleaseFromQuarantines WHERE CustomerOrderID = " & CustomerOrderID)
'
'    If Not rs.EOF Then
'
'      Dim PCSToTransfer: PCSToTransfer = rs.fields("PCSToTransfer")
'      Dim WarehousePlaceID: WarehousePlaceID = rs.fields("WarehousePlaceID")
'      Dim DescriptionOfRelease: DescriptionOfRelease = rs.fields("DescriptionOfRelease")
'
'      frm("txtPCSToTransfer") = PCSToTransfer
'      frm("txtWarehousePlaceID") = WarehousePlaceID
'      frm("txtDescriptionOfRelease") = DescriptionOfRelease
'    Else
'        frm("txtPCSToTransfer") = Null
'        frm("txtWarehousePlaceID") = Null
'        frm("txtDescriptionOfRelease") = Null
'    End If
'
'End Sub

Public Sub Sync_tblTempWarehouseTransfers(frm As Form, CustomerOrderID)

    Const ModelName = "WarehouseTransfers"
    
    Const tempTable = "tblTemp" & ModelName
    Const TableName = "tbl" & ModelName
    Const SubformName = "subTemp" & ModelName
    
    ''Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    If isFalse(CustomerOrderID) Then Exit Sub
    
    RunSQL "DELETE FROM " & tempTable
    ''Build the tblTempWarehouseTransfers table based on the CustomerOrderID - First Part (Non-Quarantined only)
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    sqlStr = "SELECT CustomerOrderID,ActualQuantity AS AvailablePCS, WarehousePlaceID FROM qryOrderAssignments" & _
        " WHERE QualityControlStatus = ""OK"" AND NOT WarehousePlaceID IS NULL AND " & _
        " OutToSubsupplierDate IS NULL AND CustomerOrderID = " & CustomerOrderID
    
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = tempTable
          .fields = "WarehousePlaceID,AvailablePCS"
          .insertSQL = sqlStr
          .InsertFilterField = "WarehousePlaceID,AvailablePCS"
          rowsAffected = .Run
    End With
    
    Dim fieldItem, fieldsArr As New clsArray, fieldValuesArr As New clsArray, setStatementArr As New clsArray
    fieldsArr.arr = "PCSToTransfer,TargetWarehousePlaceID,IsChecked"
    
    ''Released from Quarantine
    Set rs = ReturnRecordset("SELECT * FROM tblReleaseFromQuarantines WHERE CustomerOrderID = " & CustomerOrderID)
    Dim PCSToTransfer, WarehousePlaceID
    If Not rs.EOF Then
        WarehousePlaceID = rs.fields("WarehousePlaceID")
        PCSToTransfer = rs.fields("PCSToTransfer")
        Dim rs2 As Recordset: Set rs2 = ReturnRecordset("SELECT * FROM " & tempTable & " WHERE WarehousePlaceID = " & WarehousePlaceID & " AND CustomerOrderID = " & CustomerOrderID)
        If rs2.EOF Then
            ''Insert here
            Set fieldValuesArr = New clsArray
            fieldValuesArr.Add CustomerOrderID
            fieldValuesArr.Add WarehousePlaceID
            fieldValuesArr.Add PCSToTransfer
            
            RunSQL "INSERT INTO " & tempTable & " (CustomerOrderID,WarehousePlaceID,AvailablePCS) VALUES (" & fieldValuesArr.JoinArr(",") & ")"
        Else
            ''Update here
            RunSQL "UPDATE " & tempTable & " SET AvailablePCS = [AvailablePCS] + " & PCSToTransfer & " WHERE WarehousePlaceID = " & WarehousePlaceID & " AND CustomerOrderID = " & CustomerOrderID
        End If
    End If
    
    Set rs = ReturnRecordset("SELECT * FROM " & TableName & "  WHERE CustomerOrderID = " & CustomerOrderID)
    Dim i As Integer
    
    Do Until rs.EOF
    
        WarehousePlaceID = rs.fields("WarehousePlaceID")
        Set setStatementArr = New clsArray
    
        For Each fieldItem In fieldsArr.arr
            setStatementArr.Add fieldItem & " = " & EscapeString(rs.fields(fieldItem), "tblTempWarehouseTransfers", fieldItem)
        Next fieldItem
        
        RunSQL "UPDATE " & tempTable & " SET " & setStatementArr.JoinArr(",") & " WHERE WarehousePlaceID = " & _
            WarehousePlaceID
        rs.MoveNext
    Loop
    
    
    
    frm(SubformName).Form.Requery
    
End Sub

Public Function frmWarehouseManagementCustomerOrder_OnCurrent(frm As Form)

    Dim ctl As control
    For Each ctl In frm.Controls
        If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Then
            ctl.Enabled = False
            ctl.Locked = True
        End If
    Next ctl
    
End Function

Private Sub Set_subCustomerOrder_Recordsource(ByVal frm As Form, CustomerOrderID)
    
    Dim sqlStr: sqlStr = "SELECT * FROM qryCustomerOrders"
    
    If CustomerOrderID = 0 Or isFalse(CustomerOrderID) Then
        sqlStr = sqlStr & " WHERE CustomerOrderID = 0"
    End If
    
    sqlStr = sqlStr & " ORDER BY CustomerOrderID"
    
    Set frm = frm("subCustomerOrder").Form
    frm.recordSource = sqlStr
    If CustomerOrderID <> 0 And Not isFalse(CustomerOrderID) Then
        FindFirst frm, "CustomerOrderID = " & CustomerOrderID
    End If
    
End Sub

Public Function frmWarehouseManagement_OnLoad(frm As Form)
    
    DefaultFormLoad frm, "CustomerOrderID"
    
    ''Delete all tempTable data
    RunSQL "DELETE FROM tblTempDeliveryToCustomers"
    RunSQL "DELETE FROM tblTempWarehouseTransfers"
    RunSQL "DELETE FROM tblTempWarehouseScraps"
    RunSQL "DELETE FROM tblTempWarehouseTransferToSubSuppliers"
    
    
    frm("subTempDeliveryToCustomers").Form.Requery
    frm("subTempWarehouseTransfers").Form.Requery
    
    frm("subTempWarehouseScraps").Form.Requery
    frm("subTempWarehouseTransferToSubSuppliers").Form.Requery
    
    Set_subCustomerOrder_Recordsource frm, 0
    DevOnly_Visibility frm
    
End Function

Private Sub DevOnly_Visibility(frm As Form)

    Dim ctl As control
    For Each ctl In frm.Controls
        If ctl.Tag Like "*devOnly*" Then
            If Environ$("computername") = "DESKTOP-3G3V8GO" Then
                ctl.Visible = True
            Else
                ctl.Visible = False
            End If
        End If
    Next ctl
    
End Sub

Public Function GetDeliverableQTY(frm As Form) As Double
    
    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    Dim sqlStr: sqlStr = "SELECT * FROM tblOrderAssignments WHERE OrderAssignmentID = " & OrderAssignmentID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If rs.EOF Then Exit Function
    
    Dim ActualQuantity: ActualQuantity = rs.fields("ActualQuantity")
    Dim QualityControlStatus: QualityControlStatus = rs.fields("QualityControlStatus")
    Dim WHTQty: WHTQty = Coalesce(rs.fields("WHTQty"), 0)
    Dim ScrapConfirmation: ScrapConfirmation = rs.fields("ScrapConfirmation")
    
    If (QualityControlStatus = "OK" And Not ScrapConfirmation) Or (QualityControlStatus = "STOP" And WHTQty > 0) Then
        GetDeliverableQTY = IIf(WHTQty > 0, WHTQty, ActualQuantity)
    End If
    
End Function

Public Function Sync_tblTempDeliveryToCustomers_2(frm As Form, CustomerOrderID)
        
    Dim sqlStr: sqlStr = "SELECT * FROM tblOrderAssignments WHERE CustomerOrderID = " & CustomerOrderID & " ORDER " & _
        " BY OrderAssignmentOrder DESC, OrderAssignmentID DESC"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim fields As New clsArray: fields.arr = "OrderAssignmentID,PCSToDeliver,AvailablePCS," & _
        "DeliveryDate,DeliveryNote,IsChecked"
   
    Dim fieldValues As New clsArray
    
    RunSQL "DELETE FROM tblTempDeliveryToCustomers"
    
    If Not rs.EOF Then
        Dim OrderAssignmentID: OrderAssignmentID = rs.fields("OrderAssignmentID")
        Dim ActualQuantity: ActualQuantity = rs.fields("ActualQuantity")
        Dim TransferredOutDate: TransferredOutDate = rs.fields("TransferredOutDate")
        Dim DCQty: DCQty = rs.fields("DCQty")
        Dim DCDeliveryDate: DCDeliveryDate = rs.fields("DCDeliveryDate")
        Dim DCDeliveryNote: DCDeliveryNote = rs.fields("DCDeliveryNote")
        Dim DCConfirmation: DCConfirmation = rs.fields("DCConfirmation")
        Dim ScrapConfirmation: ScrapConfirmation = rs.fields("ScrapConfirmation")
        Dim ScrapQty: ScrapQty = rs.fields("ScrapQty")
        Dim WHTQty: WHTQty = rs.fields("WHTQty")
        
        Dim QualityControlStatus: QualityControlStatus = rs.fields("QualityControlStatus")
        
        If (QualityControlStatus = "OK" And Not ScrapConfirmation) Or (QualityControlStatus = "STOP" And WHTQty > 0) Then
        
            Set fieldValues = New clsArray
            fieldValues.Add OrderAssignmentID
            fieldValues.Add Coalesce(DCQty, 0)
            fieldValues.Add Coalesce(WHTQty, ActualQuantity)
            fieldValues.Add Coalesce(DCDeliveryDate, Date)
            fieldValues.Add DCDeliveryNote
            fieldValues.Add DCConfirmation
            
            UpsertRecord "tblTempDeliveryToCustomers", fields, fieldValues
        
        End If
        
        
    End If
    
    frm("subTempDeliveryToCustomers").Form.Requery
    
End Function

Public Sub Sync_tblTempDeliveryToCustomers(frm As Form, CustomerOrderID)
    
    Const ModelName = "DeliveryToCustomers"
    Const tempTable = "tblTemp" & ModelName
    Const TableName = "tbl" & ModelName
    Const SubformName = "subTemp" & ModelName
    
    ''Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    If isFalse(CustomerOrderID) Then Exit Sub
    
    RunSQL "DELETE FROM " & tempTable
    ''Build the tblTempDeliveryToCustomers table based on the CustomerOrderID - First Part (Non-Quarantined only)
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    sqlStr = "SELECT QualityControlStatus,CustomerOrderID,ActualQuantity AS AvailablePCS, WarehousePlaceID FROM qryOrderAssignments" & _
        " WHERE CustomerOrderID = " & CustomerOrderID & " ORDER BY OrderAssignmentOrder DESC, OrderAssignmentID DESC"
    
    Set rs = ReturnRecordset(sqlStr)
    
    Dim QualityControlStatus: QualityControlStatus = rs.fields("QualityControlStatus")
    If QualityControlStatus = "OK" Then
        Set sqlObj = New clsSQL
        With sqlObj
              .SQLType = "INSERT"
              .Source = tempTable
              .fields = "CustomerOrderID,WarehousePlaceID,AvailablePCS"
              .insertSQL = sqlStr
              .InsertFilterField = "CustomerOrderID,WarehousePlaceID,AvailablePCS"
              rowsAffected = .Run
        End With
    End If
    
    Dim fieldItem, fieldsArr As New clsArray, fieldValuesArr As New clsArray, setStatementArr As New clsArray
    ''Warehouse transfers
    Set rs = ReturnRecordset("SELECT * FROM tblWarehouseTransfers WHERE CustomerOrderID = " & CustomerOrderID & " AND IsChecked AND " & _
        "NOT PCSToTransfer IS NULL AND NOT TargetWarehousePlaceID IS NULL")
    Dim WarehousePlaceID, PCSToTransfer, TargetWarehousePlaceID
    Do Until rs.EOF
        TargetWarehousePlaceID = rs.fields("TargetWarehousePlaceID")
        WarehousePlaceID = rs.fields("WarehousePlaceID")
        PCSToTransfer = rs.fields("PCSToTransfer")
        ''Destination Logic - see first if it's present from tempTable (the TargetWarehousePlaceID as WarehousePlaceID)
        ''if it is present update qty else insert
        Dim rs2 As Recordset: Set rs2 = ReturnRecordset("SELECT * FROM " & tempTable & " WHERE WarehousePlaceID = " & TargetWarehousePlaceID & " AND CustomerOrderID = " & CustomerOrderID)
        If rs2.EOF Then
            ''Insert here
            Set fieldValuesArr = New clsArray
            fieldValuesArr.Add CustomerOrderID
            fieldValuesArr.Add TargetWarehousePlaceID
            fieldValuesArr.Add PCSToTransfer
            RunSQL "INSERT INTO " & tempTable & " (CustomerOrderID,WarehousePlaceID,AvailablePCS) VALUES (" & fieldValuesArr.JoinArr(",") & ")"
        End If
        
        ''RunSQL "UPDATE " & tempTable & " SET AvailablePCS = [AvailablePCS] - " & PCSToTransfer & " WHERE WarehousePlaceID = " & WarehousePlaceID & " AND CustomerOrderID = " & CustomerOrderID
        
        rs.MoveNext
    Loop
    
    ''Released from Quarantine
    Set rs = ReturnRecordset("SELECT * FROM tblReleaseFromQuarantines WHERE CustomerOrderID = " & CustomerOrderID)
    If Not rs.EOF Then
        WarehousePlaceID = rs.fields("WarehousePlaceID")
        PCSToTransfer = rs.fields("PCSToTransfer")
        Set rs2 = ReturnRecordset("SELECT * FROM " & tempTable & " WHERE WarehousePlaceID = " & WarehousePlaceID & " AND CustomerOrderID = " & CustomerOrderID)
        If rs2.EOF Then
            ''Insert here
            Set fieldValuesArr = New clsArray
            fieldValuesArr.Add CustomerOrderID
            fieldValuesArr.Add WarehousePlaceID
            fieldValuesArr.Add PCSToTransfer
            RunSQL "INSERT INTO " & tempTable & " (CustomerOrderID,WarehousePlaceID,AvailablePCS) VALUES (" & fieldValuesArr.JoinArr(",") & ")"
        Else
            ''Update here
            RunSQL "UPDATE " & tempTable & " SET AvailablePCS = [AvailablePCS] + " & PCSToTransfer & " WHERE WarehousePlaceID = " & WarehousePlaceID & " AND CustomerOrderID = " & CustomerOrderID
        End If
    End If
    
    fieldsArr.arr = "PCSToDeliver,DeliveryDate,DeliveryNote,IsChecked"
    Set rs = ReturnRecordset("SELECT * FROM " & TableName & " WHERE CustomerOrderID = " & CustomerOrderID)
    Do Until rs.EOF
    
        WarehousePlaceID = rs.fields("WarehousePlaceID")
        CustomerOrderID = rs.fields("CustomerOrderID")
        
        Set fieldValuesArr = New clsArray
        fieldValuesArr.Add EscapeString(rs.fields("PCSToDeliver"), "tblTempDeliveryToCustomers", "PCSToDeliver")
        fieldValuesArr.Add EscapeString(rs.fields("DeliveryDate"), "tblTempDeliveryToCustomers", "DeliveryDate")
        fieldValuesArr.Add EscapeString(rs.fields("DeliveryNote"), "tblTempDeliveryToCustomers", "DeliveryNote")
        fieldValuesArr.Add EscapeString(rs.fields("IsChecked"), "tblTempDeliveryToCustomers", "IsChecked")
        
        Set setStatementArr = New clsArray
        Dim i As Integer: i = 0
        
        For Each fieldItem In fieldsArr.arr
            setStatementArr.Add fieldItem & " = " & fieldValuesArr.arr(i)
            i = i + 1
        Next fieldItem
        
        RunSQL "UPDATE " & tempTable & " SET " & setStatementArr.JoinArr(",") & " WHERE WarehousePlaceID = " & _
            WarehousePlaceID & " AND CustomerOrderID = " & CustomerOrderID
        rs.MoveNext
    Loop
    
    ''Run the updates from warehouse transfers
    Set rs = ReturnRecordset("SELECT * FROM tblWarehouseTransfers WHERE CustomerOrderID = " & CustomerOrderID & " AND IsChecked AND " & _
        "NOT PCSToTransfer IS NULL AND NOT TargetWarehousePlaceID IS NULL")
    Do Until rs.EOF
        TargetWarehousePlaceID = rs.fields("TargetWarehousePlaceID")
        WarehousePlaceID = rs.fields("WarehousePlaceID")
        PCSToTransfer = rs.fields("PCSToTransfer")
        ''Destination Logic - see first if it's present from tempTable (the TargetWarehousePlaceID as WarehousePlaceID)
        ''if it is present update qty else insert
'        Set rs2 = ReturnRecordset("SELECT * FROM " & tempTable & " WHERE WarehousePlaceID = " & TargetWarehousePlaceID & " AND CustomerOrderID = " & CustomerOrderID)
'        If Not rs2.EOF Then
'            RunSQL "UPDATE " & tempTable & " SET AvailablePCS = [AvailablePCS] + " & PCSToTransfer & " WHERE WarehousePlaceID = " & TargetWarehousePlaceID & " AND CustomerOrderID = " & CustomerOrderID
'        End If
        
        RunSQL "UPDATE " & tempTable & " SET AvailablePCS = [AvailablePCS] - " & PCSToTransfer & " WHERE WarehousePlaceID = " & WarehousePlaceID & " AND CustomerOrderID = " & CustomerOrderID
        
        rs.MoveNext
    Loop
    
    
    frm(SubformName).Form.Requery
    
    Dim QuarantinedQty: QuarantinedQty = ESum2("qryOrderAssignments", "QualityControlStatus = ""Stop"" AND CustomerOrderID = " & CustomerOrderID & _
        " AND SentBackDate IS NULL", "ActualQuantity")
    frm("QuarantinedQty") = QuarantinedQty
    
    Dim ReleasedFromQuarantine: ReleasedFromQuarantine = ELookup("tblReleaseFromQuarantines", "CustomerOrderID = " & CustomerOrderID, "PCSToTransfer")
    If ReleasedFromQuarantine = "" Then ReleasedFromQuarantine = 0
    frm("txtReleasedFromQuarantine") = ReleasedFromQuarantine
    
    Set_tblOrderAssignments_QualityControlCaption frm
    
End Sub
