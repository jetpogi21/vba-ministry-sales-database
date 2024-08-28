Attribute VB_Name = "OpenCustomerOrderMain Mod"
Option Compare Database
Option Explicit

Public Function OpenCustomerOrderMainCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
            frm("subform").Width = contFrm.Width
            frm("subform")("CustomerOrderMainID").TextAlign = 2
            frm.Width = contFrm.Width
    End Select

End Function
