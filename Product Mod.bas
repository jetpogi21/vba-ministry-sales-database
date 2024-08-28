Attribute VB_Name = "Product Mod"
Option Compare Database
Option Explicit

Public Function ProductCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Sub Update_tblProducts_QtyOnStock(ProductID)
    
    If isFalse(ProductID) Then Exit Sub
    
    Dim fields As New clsArray: fields.arr = "QtyOnStock"
    Dim fieldValues As New clsArray
    Set fieldValues = New clsArray
    fieldValues.Add GetProductQtyOnStock(ProductID)
    UpsertRecord "tblProducts", fields, fieldValues, "ProductID = " & ProductID
    
End Sub

Public Function GetProductQtyOnStock(ProductID) As Double
    
    If isFalse(ProductID) Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM qryOrderAssignments WHERE ProductID = " & ProductID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim fields As New clsArray: fields.arr = "WarehousePlace,DeliveryDate,ProductSupplier,QtyOnStock"
    Dim fieldValues As New clsArray
    
    Dim RunningBalance: RunningBalance = 0
    
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
                RunningBalance = RunningBalance + RemainingQTY
            End If
        Else
            ''There's a transfer here to the WarehousePlace
            RemainingQTY = ActualQuantity - WHTQty - ScrapQty - DCQty
            If RemainingQTY > 0 Then
                RunningBalance = RunningBalance + RemainingQTY
            End If
            
            RemainingQTY = WHTQty - TransferredOutQty - ScrapQty - DCQty
            If RemainingQTY > 0 Then
                RunningBalance = RunningBalance + RemainingQTY
            End If
            
        End If
          
NextRecord:
        rs.MoveNext
    Loop
    
    GetProductQtyOnStock = RunningBalance
    
End Function

Public Function GetProductNumber(ProductID) As String
    If isFalse(ProductID) Then Exit Function
    
    ''Format as 'S0000' so ProductID of 4 is 'S0004'
    GetProductNumber = "S" & Format(ProductID, "0000")

End Function
