Attribute VB_Name = "TempDeliveryToCustomer Mod"
Option Compare Database
Option Explicit

Public Function TempDeliveryToCustomerCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
            frm("IsChecked").Properties("DatasheetCaption") = " "
            frm.AllowAdditions = False
            frm.AllowDeletions = False
            frm("WarehousePlace").Enabled = False
            frm("AvailablePCS").Enabled = False
            frm.AfterUpdate = "=dshtTempDeliveryToCustomers_AfterUpdate([Form])"
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8:
            frm.AllowAdditions = False
            frm.AllowDeletions = False
            frm("WarehousePlace").Enabled = False
            frm("AvailablePCS").Enabled = False
            frm("PCSToDeliver").Enabled = False
            ''frm.AfterUpdate = "=dshtTempDeliveryToCustomers_AfterUpdate([Form])"
            
            Dim standardHeight: standardHeight = GetStandardControlHeight(frm)
            Dim ControlSource: ControlSource = "=" & Esc("Auslieferungen")
            CreateContinuousFormButton frm, standardHeight, ControlSource, "lblManageDeliveries", "cmdManageDeliveries"
            frm("cmdManageDeliveries").OnClick = "=Open_frmDeliveryToCustomerMain([Form])"
            
    End Select

End Function

Public Function Open_frmDeliveryToCustomerMain(frm As Form)

    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    If isFalse(OrderAssignmentID) Then Exit Function
    
    DoCmd.OpenForm "frmDeliveryToCustomerMain", , , "OrderAssignmentID = " & OrderAssignmentID
    
End Function

Public Function dshtTempDeliveryToCustomers_AfterUpdate(frm As Form)
    
    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    Dim ProductID: ProductID = frm("ProductID")
    Dim IsChecked: IsChecked = frm("isChecked")
    Dim PCSToDeliver: PCSToDeliver = frm("PCSToDeliver")
    Dim DeliveryDate: DeliveryDate = frm("DeliveryDate")
    Dim DeliveryNote: DeliveryNote = frm("DeliveryNote")
    
    
    If IsChecked Then
        Dim fields As New clsArray: fields.arr = "DCQty,DCDeliveryDate,DCDeliveryNote,DCConfirmation"
        Dim fieldValues As New clsArray
        Set fieldValues = New clsArray
        fieldValues.Add PCSToDeliver
        fieldValues.Add DeliveryDate
        fieldValues.Add DeliveryNote
        fieldValues.Add IsChecked
        UpsertRecord "tblOrderAssignments", fields, fieldValues, "OrderAssignmentID = " & OrderAssignmentID
    End If
    
    Update_tblProducts_QtyOnStock ProductID
    frmWarehouseManagement_SyncOtherTabs 3
    
End Function

Public Function TempDeliveryToCustomerValidation(frm As Form) As Boolean
    
    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    If isPresent("tblOrderAssignments", "OrderAssignmentID = " & OrderAssignmentID & " AND DCConfirmation") Then
        MsgBox "This product has already been delivered. Changes aren't allowed.", vbCritical + vbOKOnly
        TempDeliveryToCustomerValidation = False
        frm.Undo
        Exit Function
    End If
    
    Dim AvailablePCs: AvailablePCs = frm("AvailablePCS")
    Dim PCSToDeliver: PCSToDeliver = frm("PCSToDeliver")
    Dim DeliveryDate: DeliveryDate = frm("DeliveryDate")
    Dim DeliveryNote: DeliveryNote = frm("DeliveryNote")
    Dim IsChecked: IsChecked = frm("IsChecked")
    
    If CDbl(AvailablePCs) < PCSToDeliver Then
        MsgBox "PCS To Deliver should not exceed the Available PCs."
        Exit Function
    End If

    TempDeliveryToCustomerValidation = True
    
End Function
