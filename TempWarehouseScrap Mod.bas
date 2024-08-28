Attribute VB_Name = "TempWarehouseScrap Mod"
Option Compare Database
Option Explicit

Public Function TempWarehouseScrapCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
            frm.AllowAdditions = False
            frm.AllowDeletions = False
            
            frm("TransferredFrom").Enabled = False
            frm("WarehousePlace").Enabled = False
            frm("QtyToScrap").Enabled = False
            
            frm.AfterUpdate = "=dshtTempWarehouseScraps_AfterUpdate([Form])"
            frm("WarehousePlace").ControlSource = "=GetScrapWarehousePlaces([OrderAssignmentID])"
            
            Dim ctl As control: Set ctl = CreateControl(frm.Name, acTextBox, acFooter, , , 0, 0, 0, 0)
            ctl.Name = "SumScrappableQty"
            ctl.ControlSource = "=CdblNz(Sum(IIf([isChecked],0,[QtyToScrap])))"
            
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function

Public Function GetScrapWarehousePlaces(OrderAssignmentID) As String
    
    If isFalse(OrderAssignmentID) Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM qryOrderAssignments WHERE OrderAssignmentID = " & OrderAssignmentID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If rs.EOF Then Exit Function
    
    Dim WarehousePlace: WarehousePlace = Coalesce(rs.fields("WarehousePlace"), "Quarantined")
    Dim WHTQty: WHTQty = Coalesce(rs.fields("WHTQty"), 0)
    Dim ActualQuantity: ActualQuantity = rs.fields("ActualQuantity")
    Dim WHTWarehousePlace: WHTWarehousePlace = rs.fields("WHTWarehousePlace")
    Dim TransferredOutQty: TransferredOutQty = Coalesce(rs.fields("TransferredOutQty"), 0)
    Dim DCQty: DCQty = Coalesce(rs.fields("DCQty"), 0)
    
    Dim Warehouses As New clsArray
    
    ''Dual warehouse
    If WHTQty <> 0 Then
        ''warehouses.Add WHTWarehousePlace & ": " & WHTQty
        Dim RemainingQTY: RemainingQTY = ActualQuantity - WHTQty
    
        If RemainingQTY <> 0 Then
            Warehouses.Add WarehousePlace & ": " & RemainingQTY
'        Else
'            Warehouses.Add WHTWarehousePlace & ": " & WHTQty
        End If
       
        Dim QtyAfterTransfer: QtyAfterTransfer = 0
        If TransferredOutQty <> 0 Then
            QtyAfterTransfer = WHTQty - TransferredOutQty
            If QtyAfterTransfer <> 0 Then
                Warehouses.Add WHTWarehousePlace & ": " & QtyAfterTransfer
            End If
        Else
            RemainingQTY = WHTQty - DCQty
            If RemainingQTY > 0 Then
                Warehouses.Add WHTWarehousePlace & ": " & RemainingQTY
            End If
        End If
        
    Else
        ActualQuantity = ActualQuantity - TransferredOutQty - DCQty
        Warehouses.Add WarehousePlace & ": " & ActualQuantity
    End If
    
    If Warehouses.count = 0 Then
        If isFalse(WarehousePlace) Then
            Warehouses.Add "Quarantined: " & ActualQuantity
        Else
            Warehouses.Add WarehousePlace & ": " & ActualQuantity
        End If
    End If
     
    GetScrapWarehousePlaces = Warehouses.JoinArr(" | ")
    
End Function

Public Function TempWarehouseScrapValidation(frm As Form) As Boolean

    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    If isPresent("tblOrderAssignments", "OrderAssignmentID = " & OrderAssignmentID & " AND ScrapConfirmation") Then
        MsgBox "This product has already been scrapped. Changes aren't allowed.", vbCritical + vbOKOnly
        TempWarehouseScrapValidation = False
        frm.Undo
        Exit Function
    End If
    
    
    Dim IsChecked: IsChecked = frm("isChecked")
    Dim Reason: Reason = frm("Reason")
    Dim QtyToScrap: QtyToScrap = frm("QtyToScrap")
    
    If IsChecked Then
    
        If QtyToScrap = 0 Then
            MsgBox "You can't scrap 0 qty.", vbCritical + vbOKOnly
            TempWarehouseScrapValidation = False
            Exit Function
        End If
        
        If isFalse(Reason) Then
            MsgBox "Please enter a reason for this scrapping.", vbCritical + vbOKOnly
            TempWarehouseScrapValidation = False
            frm("isChecked") = False
            frm("Reason").SetFocus
            Exit Function
        End If
    End If
    
    TempWarehouseScrapValidation = True
    
    
End Function

Public Function dshtTempWarehouseScraps_AfterUpdate(frm As Form)
    
    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    Dim ProductID: ProductID = frm("ProductID")
    Dim IsChecked: IsChecked = frm("isChecked")
    Dim Reason: Reason = frm("Reason")
    Dim QtyToScrap: QtyToScrap = frm("QtyToScrap")
    
    ''ReasonForScrap,ScrapQty,ScrapConfirmation
    If IsChecked Then
        Dim fields As New clsArray: fields.arr = "ReasonForScrap,ScrapQty,ScrapConfirmation"
        Dim fieldValues As New clsArray
        Set fieldValues = New clsArray
        fieldValues.Add Reason
        fieldValues.Add QtyToScrap
        fieldValues.Add True
        UpsertRecord "tblOrderAssignments", fields, fieldValues, "OrderAssignmentID = " & OrderAssignmentID
    End If
    
    Update_tblProducts_QtyOnStock ProductID
    frmWarehouseManagement_SyncOtherTabs 2
    
End Function
