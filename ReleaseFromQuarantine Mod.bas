Attribute VB_Name = "ReleaseFromQuarantine Mod"
Option Compare Database
Option Explicit

Public Function ReleaseFromQuarantineCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function Save_tblReleaseFromQuarantines(frm As Form)
    
    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    
    If isFalse(CustomerOrderID) Then Exit Function
    
    Dim PCSToTransfer: PCSToTransfer = frm("txtPCSToTransfer")
    Dim WarehousePlaceID: WarehousePlaceID = frm("txtWarehousePlaceID")
    Dim DescriptionOfRelease: DescriptionOfRelease = frm("txtDescriptionOfRelease")
    
    If isFalse(PCSToTransfer) Then
        MsgBox "PCS to Transfer is a required field."
        frm("txtPCSToTransfer").SetFocus
        Exit Function
    End If
    
    If isFalse(WarehousePlaceID) Then
        MsgBox "Warehouse Place is a required field."
        frm("txtWarehousePlaceID").SetFocus
        Exit Function
    End If
    
    If isFalse(DescriptionOfRelease) Then
        MsgBox "Description of Release is a required field."
        frm("txtDescriptionOfRelease").SetFocus
        Exit Function
    End If
    
    ''Validate first
    Dim QuarantinedQty: QuarantinedQty = frm("QuarantinedQty")
    If PCSToTransfer > QuarantinedQty Then
        MsgBox "Stk. Fur Transfer will exceed the Lagerbestand Stk.", vbCritical
        Exit Function
    End If
    
    Dim fieldsArr As New clsArray: fieldsArr.arr = "CustomerOrderID,PCSToTransfer,WarehousePlaceID,DescriptionOfRelease"
    Dim fieldValuesArr As New clsArray
    fieldValuesArr.Add CustomerOrderID
    fieldValuesArr.Add PCSToTransfer
    fieldValuesArr.Add WarehousePlaceID
    fieldValuesArr.Add DescriptionOfRelease
    
    UpsertRecord "tblReleaseFromQuarantines", fieldsArr, fieldValuesArr, "CustomerOrderID = " & CustomerOrderID
    
    Sync_tblTempDeliveryToCustomers frm, CustomerOrderID
    Sync_tblTempWarehouseTransfers frm, CustomerOrderID
    
    
    MsgBox "Successfully released from quarantine."
    
End Function
