Attribute VB_Name = "TempWarehouseTransfer Mod"
Option Compare Database
Option Explicit

Public Function TempWarehouseTransferCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
            frm("IsChecked").Properties("DatasheetCaption") = " "
            frm.AllowAdditions = False
            frm.AllowDeletions = False
            frm("WarehousePlace").Enabled = False
            frm("DescriptionOfRelease").Properties("DatasheetCaption") = "Begründung (when quarantined)"
            frm("AvailablePCS").Enabled = False
            frm.AfterUpdate = "=dshtTempWarehouseTransfers_AfterUpdate([Form])"
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function TempWarehouseTransferValidation(frm As Form) As Boolean
    
    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    If isPresent("tblOrderAssignments", "WHTConfirmation AND OrderAssignmentID = " & OrderAssignmentID) Then
        MsgBox "This product has been transferred already. Changes not allowed."
        frm.Undo
        Exit Function
    End If
    
    Dim IsChecked: IsChecked = frm("IsChecked")
    Dim TargetWarehousePlaceID: TargetWarehousePlaceID = frm("TargetWarehousePlaceID")
    Dim PCSToTransfer: PCSToTransfer = frm("PCSToTransfer")
    Dim AvailablePCs: AvailablePCs = frm("AvailablePCS")
    Dim WarehousePlaceID: WarehousePlaceID = frm("WarehousePlaceID")
    Dim DescriptionOfRelease: DescriptionOfRelease = frm("DescriptionOfRelease")
    Dim DCConfirmation: DCConfirmation = ELookup("tblOrderAssignments", "OrderAssignmentID = " & OrderAssignmentID, "DCConfirmation")
    
    If IsChecked Then
        
        If CBool(DCConfirmation) Then
            MsgBox "Product can't be transferred since it's already been delivered to the customer."
            Exit Function
        End If
        
        If AvailablePCs < PCSToTransfer Then
            MsgBox "PCS To Transfer should not exceed the Available PCs."
            Exit Function
        End If
        
        If isFalse(WarehousePlaceID) And isFalse(DescriptionOfRelease) Then
            MsgBox "Description of release is a required field when releasing from quarantine."
            frm("DescriptionOfRelease").SetFocus
            Exit Function
        End If
        
        If Not isFalse(WarehousePlaceID) And Not isFalse(DescriptionOfRelease) Then
            MsgBox "Description of release should be empty when not on quarantined."
            frm("DescriptionOfRelease").SetFocus
            Exit Function
        End If
    
    End If
    
    TempWarehouseTransferValidation = True
    
End Function

Public Function dshtTempWarehouseTransfers_AfterUpdate(frm As Form)

    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    
    Dim IsChecked: IsChecked = frm("isChecked")
    Dim PCSToTransfer: PCSToTransfer = frm("PCSToTransfer")
    Dim TargetWarehousePlaceID: TargetWarehousePlaceID = frm("TargetWarehousePlaceID")
    Dim DescriptionOfRelease: DescriptionOfRelease = frm("DescriptionOfRelease")
    
    
    If IsChecked Then
        Dim fields As New clsArray: fields.arr = "WHTWarehousePlaceID,WHTQty,WHTDescriptionOfRelease,WHTConfirmation"
        Dim fieldValues As New clsArray
        Set fieldValues = New clsArray
        fieldValues.Add TargetWarehousePlaceID
        fieldValues.Add PCSToTransfer
        fieldValues.Add DescriptionOfRelease
        fieldValues.Add IsChecked
        UpsertRecord "tblOrderAssignments", fields, fieldValues, "OrderAssignmentID = " & OrderAssignmentID
    End If
    
    frmWarehouseManagement_SyncOtherTabs 1
    
End Function

'Public Function dshtTempWarehouseTransfers_AfterUpdate(frm As Form)
'
'    ''It's either insert or update depending on wether a record is present on
'    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
'    Dim WarehousePlaceID: WarehousePlaceID = frm("WarehousePlaceID")
'    Dim PCSToTransfer: PCSToTransfer = frm("PCSToTransfer")
'    Dim TargetWarehousePlaceID: TargetWarehousePlaceID = frm("TargetWarehousePlaceID")
'    Dim IsChecked: IsChecked = frm("IsChecked")
'
'    Dim fieldArr As New clsArray, fieldValuearr As New clsArray
'    fieldArr.arr = "WarehousePlaceID,CustomerOrderID,PCSToTransfer,TargetWarehousePlaceID,IsChecked"
'
'    fieldValuearr.Add WarehousePlaceID
'    fieldValuearr.Add CustomerOrderID
'    fieldValuearr.Add PCSToTransfer
'    fieldValuearr.Add TargetWarehousePlaceID
'    fieldValuearr.Add IsChecked
'
'    UpsertRecord "tblWarehouseTransfers", fieldArr, fieldValuearr, "CustomerOrderID = " & CustomerOrderID & " AND WarehousePlaceID = " & WarehousePlaceID
'
'End Function
