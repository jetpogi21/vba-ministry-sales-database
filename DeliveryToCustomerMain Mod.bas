Attribute VB_Name = "DeliveryToCustomerMain Mod"
Option Compare Database
Option Explicit

Public Function DeliveryToCustomerMainCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            frm("AvailablePCs").ControlSource = "=GetDeliverableQTY([Form])"
            frm.Caption = "Delivery To Customers"
            frm("subDeliveryToCustomers").LinkMasterFields = ""
            frm("subDeliveryToCustomers").LinkChildFields = ""
            frm("subDeliveryToCustomers").SourceObject = "dshtDeliveryToCustomer2s"
            frm.OnCurrent = "=frmDeliveryToCustomerMain_OnCurrent([Form])"
            frm("cmdSaveClose").OnClick = "=frmDeliveryToCustomerMain_cmdSaveClose_OnClick([Form], True)"
            
            Dim fieldArr As New clsArray: fieldArr.arr = "TransferredFrom,WarehousePlace,AvailablePCs"
            
            Dim field
            For Each field In fieldArr.arr
                frm(field).Enabled = False
            Next field
            
            Dim ctl As control
            Set ctl = CreateTextboxControl(frm, "GetSubformValue([subDeliveryToCustomers]![SumPCSToDeliver])", "SumPCSToDeliver", , , , True)
            
            frm("cmdCancel").OnClick = "=CancelEdit([subDeliveryToCustomers]![Form], True)"
        
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function

Public Function frmDeliveryToCustomerMain_cmdSaveClose_OnClick(frm As Object, Optional closeForm As Boolean = False)
    
    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    Dim AvailablePCs: AvailablePCs = frm("AvailablePCs")
    Dim SumPCSToDeliver: SumPCSToDeliver = frm("SumPCSToDeliver")
    
    ''Validate first
    If SumPCSToDeliver > AvailablePCs Then
        ShowError "Total Auslieferung Stk. will exceed the Lagerbestand Stk."
        Exit Function
    End If
    
    Sync_tblDeliveryToCustomers frm
    Sync_tblDeliveryToCustomers_to_tblOrderAssignments frm
    
    Set frm = GetForm("frmWarehouseManagement")
    If Not frm Is Nothing Then
        frmWarehouseManagement_fltrCommissionNumber_AfterUpdate frm
    End If
    
    frmCustomerOrders_Racalculate
    
    If closeForm Then
        DoCmd.Close acForm, "frmDeliveryToCustomerMain", acSaveNo
    End If
    
End Function

Public Function frmDeliveryToCustomerMain_OnCurrent(frm As Object)
    
    Sync_tblDeliveryToCustomer2s frm
    
    frm("subDeliveryToCustomers").Form.Requery
    frm("subDeliveryToCustomers").SetFocus
    frm("subDeliveryToCustomers").Form("PCSToDeliver").SetFocus
    
End Function

Public Function Sync_tblDeliveryToCustomers(frm As Object)
    
    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    Dim sqlStr: sqlStr = "SELECT * FROM tblDeliveryToCustomer2s"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    RunSQL "DELETE FROM tblDeliveryToCustomers WHERE OrderAssignmentID = " & OrderAssignmentID
    
    Dim fields As New clsArray: fields.arr = "OrderAssignmentID,PCSToDeliver,DeliveryDate,DeliveryNote"
    Dim fieldValues As New clsArray
    Do Until rs.EOF
        Dim PCSToDeliver: PCSToDeliver = rs.fields("PCSToDeliver")
        Dim DeliveryDate: DeliveryDate = rs.fields("DeliveryDate")
        Dim DeliveryNote: DeliveryNote = rs.fields("DeliveryNote")

        Set fieldValues = New clsArray
        fieldValues.Add OrderAssignmentID
        fieldValues.Add PCSToDeliver
        fieldValues.Add DeliveryDate
        fieldValues.Add DeliveryNote
        UpsertRecord "tblDeliveryToCustomers", fields, fieldValues
        rs.MoveNext
    Loop
    
End Function

Public Function Sync_tblDeliveryToCustomers_to_tblOrderAssignments(frm As Object)
    
    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    
    Dim PCSToDeliver: PCSToDeliver = ESum2("tblDeliveryToCustomers", "OrderAssignmentID = " & OrderAssignmentID, "PCSToDeliver")
    Dim DCConfirmation: DCConfirmation = True
    If PCSToDeliver = 0 Then
        PCSToDeliver = Null
        DCConfirmation = False
    End If

    Dim fields As New clsArray: fields.arr = "DCQty,DCConfirmation"
    Dim fieldValues As New clsArray
    Set fieldValues = New clsArray
    fieldValues.Add PCSToDeliver
    fieldValues.Add DCConfirmation
    
    UpsertRecord "tblOrderAssignments", fields, fieldValues, "OrderAssignmentID = " & OrderAssignmentID
        
End Function

