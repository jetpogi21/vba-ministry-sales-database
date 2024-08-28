Attribute VB_Name = "qryCustomerOrderReportSubSuppliers Mod"
Option Compare Database
Option Explicit

Public Function qryCustomerOrderReportSubSuppliersCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
          
            Dim ctl As control: Set ctl = frm("WasInQuarantine")
            ctl.Enabled = False
            ''ctl.Locked = True
            ctl.FontName = "Wingdings"
            ctl.Format = "ü;\û"
            ctl.fontSize = 12
            ctl.TextAlign = 2
            
            FormatFormAsReport frm, FormTypeID
            
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function


