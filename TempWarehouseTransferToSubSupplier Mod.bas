Attribute VB_Name = "TempWarehouseTransferToSubSupplier Mod"
Option Compare Database
Option Explicit

Public Function TempWarehouseTransferToSubSupplierCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
            Dim standardHeight: standardHeight = GetStandardControlHeight(frm)
            Dim ControlSource As String
            ControlSource = "=Get_lbl_cmdTransferToNext_ControlSource([Form])"

            CreateContinuousFormButton frm, standardHeight, ControlSource, "lbl_cmdTransferToNext", "cmdTransferToNext"
            frm("cmdTransferToNext").OnClick = "=TransferOutToNextSubSupplierOrSendBack([Form])"
            
            frm.AllowAdditions = False
            frm.AllowDeletions = False
            
            Dim ctl As control: Set ctl = frm("lbl_cmdTransferToNext")
            Dim cond As FormatCondition
            ctl.FormatConditions.Delete
            
            Set cond = ctl.FormatConditions.Add(acFieldValue, acEqual, Esc("Delivered"))
            cond.BackColor = vbRed
            cond.ForeColor = vbWhite
            
            Set cond = ctl.FormatConditions.Add(acFieldValue, acEqual, Esc("Transferred"))
            cond.BackColor = vbRed
            cond.ForeColor = vbWhite
        
            frm("TransferredFrom").Enabled = False
            frm("WarehousePlace").Enabled = False
            frm("TransferredTo").Enabled = False
            frm("ActualQuantity").Enabled = False
            frm("ActualQuantity").Format = "Standard"
            ''frm("TransferredOutDate").Enabled = False
            ''frm("ScrappedQty").Enabled = False
            
            Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , "TransferredOutToID", 0, 0, 0, 0)
            ctl.Visible = False
            ctl.Name = "TransferredOutToID"
            
            frm.AfterUpdate = "=contTempWarehouseTransferToSubSuppliers_AfterUpdate([Form])"
    End Select

End Function

Public Function Get_lbl_cmdTransferToNext_ControlSource(frm As Form) As String
    
    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    If isPresent("tblOrderAssignments", "DCConfirmation AND OrderAssignmentID = " & OrderAssignmentID) Then
        Dim TransferredOutQty: TransferredOutQty = frm("TransferredOutQty")
        Dim ActualQuantity: ActualQuantity = frm("ActualQuantity")
        If TransferredOutQty = ActualQuantity Then
            Get_lbl_cmdTransferToNext_ControlSource = "Delivered"
        Else
            Get_lbl_cmdTransferToNext_ControlSource = "Auslieferungen"
        End If
        Exit Function
    End If
    
    If isPresent("tblOrderAssignments", "NOT TransferredOutDate IS NULL AND OrderAssignmentID = " & OrderAssignmentID) Then
        Get_lbl_cmdTransferToNext_ControlSource = "Transferred"
        Exit Function
    End If
    
    Dim WarehousePlace: WarehousePlace = frm("WarehousePlace")
    If WarehousePlace = "Quarantined" Then
        Get_lbl_cmdTransferToNext_ControlSource = "Send back"
    Else
        Get_lbl_cmdTransferToNext_ControlSource = "Transfer to next"
    End If
    
End Function

Public Function TransferOutToNextSubSupplierOrSendBack(frm As Form)
    
    Dim lbl_cmdTransferToNext: lbl_cmdTransferToNext = frm("lbl_cmdTransferToNext")
    
    If lbl_cmdTransferToNext = "Auslieferungen" Then
        Open_frmDeliveryToCustomerMain frm
        Exit Function
    End If
    
    If lbl_cmdTransferToNext = "Delivered" Or lbl_cmdTransferToNext = "Transferred" Then Exit Function
    
    Dim OriginalFrm As Form:  Set OriginalFrm = frm
    Dim WarehousePlaceID: WarehousePlaceID = frm("WarehousePlaceID")
    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    Dim TransferredOutQty: TransferredOutQty = frm("TransferredOutQty")
    Dim WarehousePlace: WarehousePlace = frm("WarehousePlace")
    Dim ProductID: ProductID = frm("ProductID")
    
    If WarehousePlace = "Quarantined" Then
    
        If Not IsFormOpen("frmSubSupplierManagementMain") Then
            DoCmd.OpenForm "frmSubSupplierManagementMain", , , , , acHidden
        End If
        
        Set frm = Forms("frmSubSupplierManagementMain")
        ''Get the CustomerOrderID referencing the OrderAssignmentID
        frm("fltrCommissionNumber") = CustomerOrderID
        
        Set frm = frm("subform").Form("subOrderAssignments").Form
        
        FindFirst frm, "OrderAssignmentID = " & OrderAssignmentID
        
        SendBackToPreviousSubSupplier frm, TransferredOutQty
        ''Debug.Print frm("subform").Form("subOrderAssignments").Form("OrderAssignmentOrder")
        
        ''FindFirst OriginalFrm, "OrderAssignmentID = " & OrderAssignmentID
        TransferOutToNextSubSupplier OriginalFrm
        
    Else
        TransferOutToNextSubSupplier frm
    End If
    
    Update_tblProducts_QtyOnStock ProductID
    frmWarehouseManagement_SyncOtherTabs 0
    
End Function

Public Function TransferOutToNextSubSupplier(ByVal frm As Form)
    
    Dim TempWarehouseTransferToSubSupplierID: TempWarehouseTransferToSubSupplierID = frm("TempWarehouseTransferToSubSupplierID")
    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    Dim OriginalOrderAssignmentID: OriginalOrderAssignmentID = OrderAssignmentID
    Dim TransferredOutToID: TransferredOutToID = frm("TransferredOutToID")
    Dim TransferredOutQty: TransferredOutQty = frm("TransferredOutQty")
    Dim TransferredOutDate: TransferredOutDate = frm("TransferredOutDate")
    
    If isFalse(TransferredOutQty) Then
        MsgBox "Please provide a valid QTY to transfer."
        Exit Function
    End If
    
    If Not isFalse(TransferredOutToID) Then
        MsgBox "This product has been transferred out already."
        Exit Function
    End If

    Dim rs As Recordset: Set rs = GetNextOrderAssignment(OrderAssignmentID)
    If rs.EOF Then
        MsgBox "There's no subsupplier to transfer to."
        Exit Function
    End If
    
    OrderAssignmentID = rs.fields("OrderAssignmentID")
On Error GoTo ErrHandler:
    frm.Requery
    FindFirst frm, "TempWarehouseTransferToSubSupplierID = " & TempWarehouseTransferToSubSupplierID
    frm("TransferredOutToID") = OrderAssignmentID
    frm("TransferredOutDate") = Date
    DoCmd.RunCommand acCmdSaveRecord
    
'    Dim fields As New clsArray: fields.arr = "TransferredOutToID,TransferredOutDate"
'    Dim fieldValues As New clsArray
'    fieldValues.Add OrderAssignmentID
'    fieldValues.Add Date
'
'    UpsertRecord "tblTempWarehouseTransferToSubSuppliers", fields, fieldValues, "TempWarehouseTransferToSubSupplierID = " & _
'        TempWarehouseTransferToSubSupplierID
    
    Dim fields As New clsArray:
    fields.arr = "TransferredOutToID,TransferredOutQty,TransferredOutDate"
    Dim fieldValues As New clsArray
    Set fieldValues = New clsArray
    fieldValues.Add OrderAssignmentID
    fieldValues.Add TransferredOutQty
    fieldValues.Add Date
    
    UpsertRecord "tblOrderAssignments", fields, fieldValues, "OrderAssignmentID = " & OriginalOrderAssignmentID
    
    fields.arr = "OutToSubsupplierDate,ActualQuantity"
    
    Set fieldValues = New clsArray
    fieldValues.Add Date
    fieldValues.Add TransferredOutQty
    
    UpsertRecord "tblOrderAssignments", fields, fieldValues, "OrderAssignmentID = " & OrderAssignmentID
    
    If IsFormOpen("frmSubSupplierManagementMain") Then
        Forms("frmSubSupplierManagementMain")("subform").Form("subOrderAssignments").Form.Requery
    End If
    
    MsgBox "Products successfully transferred to the next subsupplier."
    
    Exit Function
    
ErrHandler:
    If Err.Number = 2001 Then
        Exit Function
    End If
    
End Function

Public Function TempWarehouseTransferToSubSupplierValidation(frm As Form) As Boolean
    
    Dim TransferredOutToID: TransferredOutToID = frm("TransferredOutToID")
    Dim TransferredOutQty: TransferredOutQty = frm("TransferredOutQty")
    Dim ActualQuantity: ActualQuantity = frm("ActualQuantity")
    Dim TransferredOutDate: TransferredOutDate = frm("TransferredOutDate")
    Dim OldTransferredOutDate: OldTransferredOutDate = frm("TransferredOutDate").oldValue
    Dim lbl_cmdTransferToNext: lbl_cmdTransferToNext = frm("lbl_cmdTransferToNext")
    
    If IsNull(OldTransferredOutDate) And Not IsNull(TransferredOutDate) And isFalse(TransferredOutToID) Then
        MsgBox "Changing Transferred out date from blank to filled is not allowed manually."
        DoCmd.CancelEvent
        frm.Undo
        Exit Function
    End If
    
    If lbl_cmdTransferToNext = "Delivered" Then
        ShowError "Changing dates manually for delivered products is not allowed."
        frm.Undo
        Exit Function
    End If
    
    If TransferredOutQty > CDbl(ActualQuantity) Then
        MsgBox "Product to transfer shouldn't exceed the actual quantity received.", vbOKOnly
        frm("TransferredOutQty").SetFocus
        Exit Function
    End If
    
    TempWarehouseTransferToSubSupplierValidation = True
    
End Function

Public Function contTempWarehouseTransferToSubSuppliers_AfterUpdate(frm As Form)
    
    
    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    Dim TransferredOutDate: TransferredOutDate = frm("TransferredOutDate")
    
    If Not isFalse(TransferredOutDate) Then
        Dim fields As New clsArray: fields.arr = "TransferredOutDate"
        Dim fieldValues As New clsArray
        Set fieldValues = New clsArray
        fieldValues.Add TransferredOutDate
        UpsertRecord "tblOrderAssignments", fields, fieldValues, "OrderAssignmentID = " & OrderAssignmentID
        ''Update the next order assignment if there's any
        Dim rs As Recordset: Set rs = GetNextOrderAssignment(OrderAssignmentID)
        
        If Not rs.EOF Then
            OrderAssignmentID = rs.fields("OrderAssignmentID")
            Set fields = New clsArray: fields.arr = "OutToSubsupplierDate"
            Set fieldValues = New clsArray
            fieldValues.Add TransferredOutDate
            UpsertRecord "tblOrderAssignments", fields, fieldValues, "OrderAssignmentID = " & OrderAssignmentID
        End If
    End If
    
    frmWarehouseManagement_SyncOtherTabs 0
    
End Function

