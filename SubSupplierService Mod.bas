Attribute VB_Name = "SubSupplierService Mod"
Option Compare Database
Option Explicit

Public Function SubSupplierServiceCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
    End Select

End Function

Public Function frmSubSupplierServiceSelector_OnCurrent(frm As Form)

    Dim SubSupplierID: SubSupplierID = frm("SubSupplierID")
    Dim filterStr: filterStr = "SubSupplierServiceID = 0"
    
    If Not isFalse(SubSupplierID) Then
       filterStr = "SubSupplierID = " & SubSupplierID
        
    End If
    
    Dim sqlStr: sqlStr = "SELECT ServiceID,Service FROM qrySubSupplierServices WHERE " & filterStr & " ORDER BY Service"
    frm("ServiceID").RowSource = sqlStr
    
End Function

Public Function frmSubSupplierServiceSelector_cmdConfirm_OnClick(frm As Form)

    Dim ServiceID: ServiceID = frm("ServiceID")
    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    
    If isFalse(OrderAssignmentID) Then Exit Function
    If isFalse(ServiceID) Then
        MsgBox "Select a valid service.", vbCritical + vbOKOnly
        Exit Function
    End If
    
    If IsFormOpen("frmSubSupplierManagementMain") Then
        Forms("frmSubSupplierManagementMain")("subform").Form("subOrderAssignments").Form("ServiceID") = ServiceID
    Else
        RunSQL "UPDATE tblOrderAssignments SET ServiceID = " & ServiceID & " WHERE OrderAssignmentID = " & OrderAssignmentID
    End If
    
    DoCmd.Close acForm, frm.Name, acSaveNo
    
End Function
