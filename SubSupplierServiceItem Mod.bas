Attribute VB_Name = "SubSupplierServiceItem Mod"
Option Compare Database
Option Explicit

Public Function SubSupplierServiceItemCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
            CreateContinuousFormDeleteButton frm, 742
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function
