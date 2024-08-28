Attribute VB_Name = "QuarantinedProductReport Mod"
Option Compare Database
Option Explicit

Public Function QuarantinedProductReportCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
            CreateTotalControl frm, "QtyOnStock"
            CreateTotalLabel frm, "QtyOnStock"
            FormatFormAsReport frm, 8
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function
