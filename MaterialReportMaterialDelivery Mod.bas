Attribute VB_Name = "MaterialReportMaterialDelivery Mod"
Option Compare Database
Option Explicit

Public Function MaterialReportMaterialDeliveryCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
            FormatFormAsReport frm, 8
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function

Public Sub Create_qryMaterialReportMaterialDeliveries()
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblOrderAssignmentMaterialDeliveries"
          .fields = "MaterialDeliveryID,SUM(QTY) As DeliveredQTY"
          .GroupBy = "MaterialDeliveryID"
          sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "qryMaterialDeliveries"
          .fields = "qryMaterialDeliveries.*,Quantity - CdblNz(temp!DeliveredQTY) AS QuantityOnStock, " & _
            "tblMaterialSuppliers.ShortName As MaterialSupplierShortName"
          .joins.Add GenerateJoinObj(sqlStr, "MaterialDeliveryID", "temp", , "LEFT")
          .joins.Add GenerateJoinObj("tblMaterialSuppliers", "MaterialSupplierID")
          .OrderBy = "qryMaterialDeliveries.MaterialDeliveryID"
          sqlStr = .sql
    End With
    
    Dim qDef As QueryDef: Set qDef = CurrentDb.QueryDefs("qryMaterialReportMaterialDeliveries")
    qDef.sql = sqlStr
    
End Sub
