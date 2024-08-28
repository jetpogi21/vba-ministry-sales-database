Attribute VB_Name = "QualityControlCustomerOrder Mod"
Option Compare Database
Option Explicit

Public Function QualityControlCustomerOrderCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function frmQualityControlCustomerOrder_OnCurrent(frm As Form)

    frm("WhiteSquare").Visible = frm.NewRecord
    
End Function
