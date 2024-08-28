Attribute VB_Name = "MaterialDeliveryMultiSelector Mod"
Option Compare Database
Option Explicit

Public Function MaterialDeliveryMultiSelectorCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
            frm.RecordSelectors = False
            frm.NavigationButtons = False
            frm("MaterialControlNumber").Enabled = False
            OffsetControlPositions frm, 50
            frm.AllowAdditions = False
            
            frm("AvailableQTY").ControlSource = "=GetMaterialAvailableQTY([Form])"
            frm("AvailableQTY").Enabled = False
            frm("AvailableQTY").Format = "Standard"
            
            Dim ctl As control
            Set ctl = CreateControl(frm.Name, acTextBox, acFooter, "", , 0, 0, 0)
            SetControlProperties ctl
            ctl.Name = "SumQTY"
            ctl.ControlSource = "=CdblNz(Sum([QTY]))"
            
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function

Public Function GetMaterialAvailableQTY(frm As Form) As Double
    
    If IsFormOpen("mainMaterialDeliveryMultiSelector") Then
        Dim MaterialID: MaterialID = Forms("mainMaterialDeliveryMultiSelector")("MaterialID")
        Dim OrderAssignmentID: OrderAssignmentID = Forms("mainMaterialDeliveryMultiSelector")("OrderAssignmentID")
        Dim MaterialDeliveryID: MaterialDeliveryID = frm("MaterialDeliveryID")
        Dim DeliveredQTY: DeliveredQTY = frm("DeliveredQTY")
        
        Dim ReleasedQTY: ReleasedQTY = ESum2("qryOrderAssignmentMaterialDeliveries", "MaterialDeliveryID = " & MaterialDeliveryID & _
            " AND Not OrderAssignmentID = " & OrderAssignmentID, "QTY")
             
        GetMaterialAvailableQTY = DeliveredQTY - ReleasedQTY
        ''Debug.Print MaterialID, OrderAssignmentID, MaterialDeliveryID
    End If
    
End Function

Public Function MaterialDeliveryMultiSelectorValidation(frm As Form) As Boolean
    
    Dim AvailableQTY: AvailableQTY = frm("AvailableQTY")
    Dim QTY: QTY = frm("QTY")
    
    If QTY > AvailableQTY Then
        MsgBox "Qty should not exceed what's available", vbOKOnly + vbCritical
        frm.Undo
        DoCmd.CancelEvent
        Exit Function
    End If
    
    MaterialDeliveryMultiSelectorValidation = True
End Function
