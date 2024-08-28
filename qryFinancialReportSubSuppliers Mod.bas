Attribute VB_Name = "qryFinancialReportSubSuppliers Mod"
Option Compare Database
Option Explicit

Public Function qryFinancialReportSubSuppliersCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
            
            Dim ctl As control
            
            CreateTotalControl frm, "ActualCost"
            CreateTotalControl frm, "DeliveryCost"
            CreateTotalControl frm, "TotalCost"
            
            CreateTotalLabel frm, "ActualCost"
            
            FormatFormAsReport frm, 8
        
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function
