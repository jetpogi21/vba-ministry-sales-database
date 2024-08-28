Attribute VB_Name = "qryCustomerOrderReportForceCloses Mod"
Option Compare Database
Option Explicit

Public Function qryCustomerOrderReportForceClosesCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
            FormatFormAsReport frm, FormTypeID
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function
