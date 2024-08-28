Attribute VB_Name = "DeliveryToCustomer2 Mod"
Option Compare Database
Option Explicit

Public Function DeliveryToCustomer2Create(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
            frm.AfterUpdate = "=dshtDeliveryToCustomer2s_AfterUpdate([Form])"
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function

Public Function dshtDeliveryToCustomer2s_AfterUpdate(frm As Object)
    
    Set frm = GetForm("frmDeliveryToCustomerMain")
    
    If Not frm Is Nothing Then
        frmDeliveryToCustomerMain_cmdSaveClose_OnClick frm, False
    End If
    
End Function

Public Function Sync_tblDeliveryToCustomer2s(frm As Object)
    
    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    Dim sqlStr: sqlStr = "SELECT * FROM tblDeliveryToCustomers WHERE OrderAssignmentID = " & OrderAssignmentID & _
        " ORDER BY OrderAssignmentID"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    RunSQL "DELETE FROM tblDeliveryToCustomer2s"
    
    Dim fields As New clsArray: fields.arr = "PCSToDeliver,DeliveryDate,DeliveryNote"
    Dim fieldValues As New clsArray
    Do Until rs.EOF
        Dim PCSToDeliver: PCSToDeliver = rs.fields("PCSToDeliver")
        Dim DeliveryDate: DeliveryDate = rs.fields("DeliveryDate")
        Dim DeliveryNote: DeliveryNote = rs.fields("DeliveryNote")

        Set fieldValues = New clsArray
        fieldValues.Add PCSToDeliver
        fieldValues.Add DeliveryDate
        fieldValues.Add DeliveryNote
        UpsertRecord "tblDeliveryToCustomer2s", fields, fieldValues
        rs.MoveNext
    Loop
    
End Function

Public Function DeliveryToCustomer2Validation(frm As Object) As Boolean
    
    Dim DeliveryToCustomer2ID: DeliveryToCustomer2ID = frm("DeliveryToCustomer2ID")
    Dim PCSToDeliver: PCSToDeliver = frm("PCSToDeliver")
    Dim DeliveryDate: DeliveryDate = frm("DeliveryDate")
    Dim DeliveryNote: DeliveryNote = frm("DeliveryNote")
    Dim AvailablePCs: AvailablePCs = frm.parent("AvailablePCs")
    
    Dim OrderAssignmentID: OrderAssignmentID = frm.parent("OrderAssignmentID")
    If Not isFalse(OrderAssignmentID) Then
        Dim CustomerOrderID: CustomerOrderID = ELookup("tblOrderAssignments", "OrderAssignmentID = " & OrderAssignmentID, "CustomerOrderID")
        Dim isClosed: isClosed = isPresent("tblCustomerOrders", "CustomerOrderID = " & CustomerOrderID & " AND OrderStatus = ""Closed""")
        If isClosed Then
            ShowError "Komm. Nr.  is already closed. Delivery changes not possible."
            Exit Function
        End If
    End If
    
    Dim SumPCSToDeliver: SumPCSToDeliver = ESum2("tblDeliveryToCustomer2s", "DeliveryToCustomer2ID <> " & DeliveryToCustomer2ID, "PCSToDeliver")
    
    If AvailablePCs < (SumPCSToDeliver + PCSToDeliver) Then
        ShowError "Total Auslieferung Stk. will exceed the Lagerbestand Stk."
        frm("PCSToDeliver").SetFocus
        Exit Function
    End If

    DeliveryToCustomer2Validation = True
    
End Function
