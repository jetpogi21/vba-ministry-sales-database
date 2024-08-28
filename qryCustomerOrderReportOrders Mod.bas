Attribute VB_Name = "qryCustomerOrderReportOrders Mod"
Option Compare Database
Option Explicit

Public Function qryCustomerOrderReportOrdersCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        
            FormatFormAsReport frm, FormTypeID, "CustomerFullName"
'            frm.AllowAdditions = False
'            frm.AllowEdits = False
'            frm.AllowDeletions = False
'
'            Dim ctl As control: Set ctl = frm("CustomerFullName")
'            OffsetControlPositions frm, (ctl.Left * -1) + 25, (frm("lblCustomerFullName").Top * -1) + 25
            
            
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function
