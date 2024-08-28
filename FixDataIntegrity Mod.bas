Attribute VB_Name = "FixDataIntegrity Mod"
Option Compare Database
Option Explicit

Public Function FixDataIntegrityCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function SetUp_SubSupplierManagement_TestData(Optional TestNumber = 1)
    
    RunSQL "DELETE FROM tblOrderAssignments"
    Dim sqlStr: sqlStr = "SELECT * FROM tblOrderAssignments_test_" & TestNumber & " ORDER BY OrderAssignmentID"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim fld As field
    Dim fieldValuesArr As New clsArray
    Dim field, fieldsArr As New clsArray
    For Each fld In rs.fields
        If fld.Name <> "OrderAssignmentID" Then
            fieldsArr.Add fld.Name
        End If
    Next fld
    
    Do Until rs.EOF
        Set fieldValuesArr = New clsArray
        For Each field In fieldsArr.arr
            If field <> "OrderAssignmentID" Then
                fieldValuesArr.Add rs.fields(field)
            End If
        Next field
        UpsertRecord "tblOrderAssignments", fieldsArr, fieldValuesArr, "OrderAssignmentID = 0"
        rs.MoveNext
    Loop
    
End Function


Private Function tblOrderAssignments_ResetQualityControl(rs As Recordset)
    
    Dim field, fieldsToReset As New clsArray: fieldsToReset.arr = "WarehousePlaceID,DescriptionOfFailure,QualityControlStatus," & _
        "OutToSubsupplierDate,QualityControlCaption,SentOutDate,SentBackDate"
    
    rs.Edit
    For Each field In fieldsToReset.arr
        rs.fields(field) = Null
    Next field
    rs.Update
    
End Function

Public Function SetUp_TestData()
    
    RunSQL "DELETE FROM tblOrderAssignments WHERE OrderAssignmentID > 0"
    
End Function

Public Function PurgeData()

    RunSQL "DELETE FROM tblCustomerOrderMains WHERE CustomerOrderMainID > 0"
    RunSQL "DELETE FROM tblWarehouseTransfers WHERE WarehouseTransferID > 0"
    RunSQL "DELETE FROM tblDeliveryToCustomers WHERE DeliveryToCustomerID > 0"
    RunSQL "DELETE FROM tblReleaseFromQuarantines WHERE ReleaseFromQuarantineID > 0"
    
End Function


Public Function RunAllDataFixes()

    Fix_tblCustomerOrders_LastAgreedDueDate
    Fix_tblOrderAssignments
    Update_tblOrderAssignment_QualityControlStatus
    DeleteCustomerOrdersWithoutCustomerOrderMain
    DeleteOrderAssignmentsWithoutCustomerOrderID
    Fix_tblCustomerOrders_Qty
    Fix_tblMaterialDeliveries_ControlNumber
    Fix_tblCustomerOrderMains
    Fix_tblCustomerOrders_UpdateProducts
    Fix_tblSubSupplierServiceItems
    Fix_tblMaterialSupplierMaterialItems
    ''CheckQuarantineWarehousePlace
    
End Function

Public Sub Fix_tblSubSupplierServiceItems()
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSubSupplierServiceItems ORDER BY SubSupplierServiceItemID"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        Dim SubSupplierID: SubSupplierID = rs.fields("SubSupplierID"): If ExitIfTrue(isFalse(SubSupplierID), "SubSupplierID is empty..") Then Exit Sub
        Dim ServiceID: ServiceID = GetRandomID("tblSubSupplierServices", "ServiceID", "SubSupplierID = " & SubSupplierID)
        
        If Not IsNull(ServiceID) Then
            rs.Edit
            rs.fields("ServiceID") = ServiceID
            rs.Update
        Else
            rs.Delete
        End If
        rs.MoveNext
    Loop
    
End Sub

Public Sub Fix_tblMaterialSupplierMaterialItems()
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblMaterialSupplierMaterialItems ORDER BY MaterialSupplierMaterialItemID"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        Dim MaterialSupplierID: MaterialSupplierID = rs.fields("MaterialSupplierID"): If ExitIfTrue(isFalse(MaterialSupplierID), "MaterialSupplierID is empty..") Then Exit Sub
        Dim MaterialID: MaterialID = GetRandomID("tblMaterialSupplierMaterials", "MaterialID", "MaterialSupplierID = " & MaterialSupplierID)
        If Not IsNull(MaterialID) Then
            rs.Edit
            rs.fields("MaterialID") = MaterialID
            rs.Update
        Else
            rs.Delete
        End If
        rs.MoveNext
    Loop
    
End Sub

Public Sub Fix_tblCustomerOrders_UpdateProducts()

    RunSQL "UPDATE tblCustomerOrders SET ProductID1 = ProductID, ProductID2 = ProductID, ProductID3 = ProductID"
    
End Sub

Public Sub Fix_tblCustomerOrderMains()

    RunSQL "UPDATE tblCustomerOrderMains SET IsSaved = -1"
    
End Sub

Public Sub Fix_tblMaterialDeliveries_ControlNumber()
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblMaterialDeliveries ORDER BY DeliveryDate, MaterialDeliveryID"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''ControlNumber = Emax("tblMaterialDeliveries", "Year(DeliveryDate) = " & Year(DeliveryDate), "ControlNumber") + 1
    Dim ControlNumber: ControlNumber = 1
    Do Until rs.EOF
        rs.Edit
        rs.fields("ControlNumber") = ControlNumber
        rs.Update
        ControlNumber = ControlNumber + 1
        rs.MoveNext
    Loop
    
End Sub

Public Sub DeleteOrderAssignmentsWithoutCustomerOrderID()
    
    RunSQL "DELETE tblOrderAssignments.* FROM tblOrderAssignments LEFT JOIN tblCustomerOrders ON tblOrderAssignments.CustomerOrderID = " & _
        " tblCustomerOrders.CustomerOrderID WHERE tblCustomerOrders.CustomerOrderID IS NULL"
    
End Sub

Public Sub DeleteCustomerOrdersWithoutCustomerOrderMain()

    RunSQL "DELETE FROM tblCustomerOrders WHERE CustomerOrderMainID IS NULL"
    
End Sub

Public Sub Fix_tblCustomerOrders_LastAgreedDueDate()
    
    ''LastAgreedDueDate should be the latest Timestamp's CustomerDueDate from tblOrderDueDates for each CustomerOrderID or the tblCustomerOrders'
    ''CustomerDueDate if the previous condition is not applicable or its own CustomerDueDate isn't null
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblCustomerOrders ORDER BY CustomerOrderID")
    
    Do Until rs.EOF
        Dim CustomerOrderID: CustomerOrderID = rs.fields("CustomerOrderID"): If ExitIfTrue(isFalse(CustomerOrderID), "CustomerOrderID is empty..") Then Exit Sub
        Dim CustomerDueDate: CustomerDueDate = rs.fields("CustomerDueDate")
        Dim LastAgreedDueDate: LastAgreedDueDate = ELookup("tblOrderDueDates", "CustomerOrderID = " & CustomerOrderID, "CustomerDueDate", "OrderDueDateID DESC")
        
        If Not isFalse(LastAgreedDueDate) Then
            rs.Edit
            rs("LastAgreedDueDate") = CDate(LastAgreedDueDate)
            rs.Update
            GoTo NextRecord
        End If
        
        If Not isFalse(CustomerDueDate) Then
            rs.Edit
            rs("LastAgreedDueDate") = CustomerDueDate
            rs.Update
            GoTo NextRecord
        End If
NextRecord:
        rs.MoveNext
    Loop
    
End Sub

Public Sub Update_tblOrderAssignment_QualityControlStatus()
    
    Dim choiceArr As New clsArray: choiceArr.arr = "OK,Stop"
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblOrderAssignments ORDER BY OrderAssignmentID")
    Dim FieldToUpdate, fieldValue
    Do Until rs.EOF
    
        Dim randomIndex As Integer: randomIndex = GetRandomFromRange(0, 1)
        Dim QualityControlStatus: QualityControlStatus = choiceArr.items(randomIndex)
        
        If QualityControlStatus = "OK" Then
            fieldValue = GetRandomID("tblWarehousePlaces", "WarehousePlaceID")
            FieldToUpdate = "WarehousePlaceID"
        Else
            fieldValue = GetRandomID("tblMockData", "LongText")
            FieldToUpdate = "DescriptionOfFailure"
        End If
        
        rs.Edit
        rs(FieldToUpdate) = fieldValue
        rs(IIf(FieldToUpdate = "WarehousePlaceID", "DescriptionOfFailure", "WarehousePlaceID")) = Null
        rs("QualityControlStatus") = QualityControlStatus
        rs.Update
        
NextRecord:
        rs.MoveNext
    Loop
End Sub

Public Sub Fix_tblOrderAssignments()

    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM qryOrderAssignments ORDER BY OrderAssignmentID")
    Dim CustomerOrderID
    Do Until rs.EOF
        CustomerOrderID = rs.fields("CustomerOrderID")
        Dim MaterialDeliveryID: MaterialDeliveryID = rs.fields("MaterialDeliveryID")
        
        If Not isFalse(MaterialDeliveryID) Then
            If isPresent("tblMaterialDeliveryCustomerOrders", "CustomerOrderID = " & CustomerOrderID) Then
                MaterialDeliveryID = GetRandomID("tblMaterialDeliveryCustomerOrders", "MaterialDeliveryID", "CustomerOrderID = " & CustomerOrderID)
            End If
        End If
        Dim MaterialQuantity: MaterialQuantity = IIf(isFalse(MaterialDeliveryID), Null, GetRandomFromRange(0, 100))
        rs.Edit
        rs.fields("MaterialDeliveryID") = MaterialDeliveryID
        rs.fields("MaterialQuantity") = MaterialQuantity
        rs.fields("ActualCost") = rs.fields("Cost")
        rs.fields("ActualQuantity") = rs.fields("Qty")
        rs.Update
        
        rs.MoveNext
    Loop
    
    ''First check first if the CustomerOrderID is present in
    ''tblMaterialDeliveryCustomerOrders then Validate if the MaterialDeliveryID if present, a part of tblMaterialDeliveryCustomerOrders of the same CustomerOrderID.
    
    RunSQL "UPDATE tblOrderAssignments set OrderAssignmentOrder = 1 where OrderAssignmentOrder is null"
    ''There can never be a next order assignment when the QualityControlStatus is STOP.
    tblOrderAssignments_Delete_After_Stop_Status
    
    ''Set the OutToSubsupplierDate if the record is not the first record within the CustomerOrder and it's QualityControlStatus is not STOP
    Set rs = ReturnRecordset("SELECT * FROM tblOrderAssignments WHERE QualityControlStatus = ""OK"" ORDER BY CustomerOrderID,OrderAssignmentOrder,OrderAssignmentID")
    Do Until rs.EOF
        CustomerOrderID = rs.fields("CustomerOrderID"): If ExitIfTrue(isFalse(CustomerOrderID), "CustomerOrderID is empty..") Then Exit Sub
        Dim OrderAssignmentOrder: OrderAssignmentOrder = EscapeString(rs.fields("OrderAssignmentOrder"), "tblOrderAssignments", "OrderAssignmentOrder"): If ExitIfTrue(isFalse(OrderAssignmentOrder), "OrderAssignmentOrder is empty..") Then Exit Sub
        Dim OrderAssignmentID: OrderAssignmentID = rs.fields("OrderAssignmentID"): If ExitIfTrue(isFalse(OrderAssignmentID), "OrderAssignmentID is empty..") Then Exit Sub
        
        If isPresent("tblOrderAssignments", "((OrderAssignmentOrder = " & OrderAssignmentOrder & " AND OrderAssignmentID > " & OrderAssignmentID & ") " & _
         " OR OrderAssignmentOrder > " & OrderAssignmentOrder & ") " & _
         " AND CustomerOrderID = " & CustomerOrderID) Then
            rs.Edit
            rs.fields("OutToSubsupplierDate") = CDate(GetRandomFromRange(CLng(#1/1/2024#), CLng(#12/31/2024#)))
            rs.Update
        End If
        rs.MoveNext
    Loop
    
    ''Update the QualityControlCaption PLACE [pieces]" or "Quarantine [pieces]
    Set rs = ReturnRecordset("SELECT CustomerOrderID FROM qryOrderAssignments GROUP BY CustomerOrderID ORDER BY CustomerOrderID")
    Do Until rs.EOF
        CustomerOrderID = rs.fields("CustomerOrderID"): If ExitIfTrue(isFalse(CustomerOrderID), "CustomerOrderID is empty..") Then Exit Sub
        Dim rs2 As Recordset: Set rs2 = ReturnRecordset("SELECT * FROM qryOrderAssignments WHERE CustomerOrderID = " & CustomerOrderID & _
            " ORDER BY OrderAssignmentOrder DESC,OrderAssignmentID DESC")
        Dim i As Integer: i = 0
        Do Until rs2.EOF
            Dim WarehousePlace: WarehousePlace = rs2.fields("WarehousePlace")
            Dim QualityControlStatus: QualityControlStatus = rs2.fields("QualityControlStatus")
            Dim ActualQuantity: ActualQuantity = rs2.fields("ActualQuantity")
            Dim QualityControlCaption: QualityControlCaption = ""
            rs2.Edit
            If i = 0 Then
                Select Case QualityControlStatus
                    Case "OK":
                        QualityControlCaption = WarehousePlace & ": " & Format$(ActualQuantity, "Standard")
                    Case "STOP":
                        QualityControlCaption = "Quarantine: " & Format$(ActualQuantity, "Standard")
                End Select
                rs2.fields("QualityControlCaption") = QualityControlCaption
            Else
                rs2.fields("QualityControlCaption") = Null
            End If
            rs2.Update
            i = i + 1
            rs2.MoveNext
        Loop
        rs.MoveNext
    Loop
    
End Sub

Public Sub tblOrderAssignments_Delete_After_Stop_Status()

    Dim rs As Recordset
Reloop:
    Set rs = ReturnRecordset("SELECT * FROM tblOrderAssignments WHERE QualityControlStatus = ""STOP"" ORDER BY CustomerOrderID,OrderAssignmentOrder, OrderAssignmentID")
On Error GoTo Err_Handler:
    Do Until rs.EOF
        Dim CustomerOrderID: CustomerOrderID = rs.fields("CustomerOrderID"): If ExitIfTrue(isFalse(CustomerOrderID), "CustomerOrderID is empty..") Then Exit Sub
        Dim OrderAssignmentOrder: OrderAssignmentOrder = rs.fields("OrderAssignmentOrder"): If ExitIfTrue(isFalse(OrderAssignmentOrder), "OrderAssignmentOrder is empty..") Then Exit Sub
        
        OrderAssignmentOrder = EscapeString(OrderAssignmentOrder, "tblOrderAssignments", "OrderAssignmentOrder")
        
        Dim OrderAssignmentID: OrderAssignmentID = rs.fields("OrderAssignmentID"): If ExitIfTrue(isFalse(OrderAssignmentID), "OrderAssignmentID is empty..") Then Exit Sub
        
        RunSQL "DELETE FROM tblOrderAssignments " & _
         "WHERE ((OrderAssignmentOrder = " & OrderAssignmentOrder & " AND OrderAssignmentID > " & OrderAssignmentID & ") " & _
         " OR OrderAssignmentOrder > " & OrderAssignmentOrder & ") " & _
         " AND CustomerOrderID = " & CustomerOrderID
        rs.MoveNext
    Loop
    
    Exit Sub
    
Err_Handler:
    
    If Err.Number = 3167 Then
        rs.Close
        GoTo Reloop
    Else
        MsgBox Err.Description
    End If
    Exit Sub
    
End Sub

Public Sub Fix_tblCustomerOrders_Qty()

    RunSQL "UPDATE tblCustomerOrders SET QTY = CLng(QTY)"
    Fix_tblOrderAssignments
    
End Sub

'Public Sub CheckQuarantineWarehousePlace()
'
'    If Not isPresent("tblWarehousePlaces", "WarehousePlace = ""Q""") Then
'        RunSQL "INSERT INTO tblWarehousePlaces(WarehousePlace,WPDescription) VALUES (""Q"",""Quarantine"")"
'    End If
'
'End Sub
