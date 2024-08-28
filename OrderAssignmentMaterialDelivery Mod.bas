Attribute VB_Name = "OrderAssignmentMaterialDelivery Mod"
Option Compare Database
Option Explicit

Public Function OrderAssignmentMaterialDeliveryCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function

Public Function GetOrderAssignmentMaterialDeliveries(OrderAssignmentID) As String

    If isFalse(OrderAssignmentID) Then Exit Function
    GetOrderAssignmentMaterialDeliveries = Elookups("qryOrderAssignmentMaterialDeliveries", "OrderAssignmentID = " & _
        OrderAssignmentID, "MaterialControlNumber")
        
    GetOrderAssignmentMaterialDeliveries = TruncateText(GetOrderAssignmentMaterialDeliveries, 16)
        
End Function
