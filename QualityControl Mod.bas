Attribute VB_Name = "QualityControl Mod"
Option Compare Database
Option Explicit

Public Function QualityControlCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            frm("pgQualityControlItems").Caption = "Order Assignments"
            frm("Label0").Caption = "Search Customer Order"
            frm.OnCurrent = "=frmQualityControl_OnCurrent([Form])"
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function frmOrderAssignmentsWithMaterialMain_OnLoad(frm As Form)

    frm("subform").Form.Filter = "CustomerOrderID = 0"
    frm("subform").Form.FilterOn = True
    
    ''frmOrderAssignmentsWithMaterialMain_fltrSupplierShortName_RowSource frm
    
    SetNavigationData frm, False, , "subsupplier"
    
End Function


