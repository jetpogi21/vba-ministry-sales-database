Attribute VB_Name = "OrderDueDate Mod"
Option Compare Database
Option Explicit

Public Function OrderDueDateCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
            frm("CustomerDueDate").Tag = "DontAutoWidth"
            frm("CustomerDueDate").AfterUpdate = "=dshtOrderDueDates_CustomerDueDate_AfterUpdate([Form])"
            frm.AfterUpdate = "=dshtOrderDueDates_AfterUpdate([Form])"
            frm("CustomerOrderDueDate").ColumnHidden = True
            frm("CustomerOrderDueDate").Tag = "alwaysHideOnDatasheet"
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function dshtOrderDueDates_CustomerDueDate_AfterUpdate(frm As Form)
    
'    Dim CustomerDueDate: CustomerDueDate = frm("CustomerDueDate")
'    frm("CustomerOrderDueDate") = CustomerDueDate
    
End Function

Public Function GetLastAgreedDueDate(CustomerOrderID)
    
    Dim labelCaption: labelCaption = "Liefertermin 1" ''Liefertermin Neu
    If isFalse(CustomerOrderID) Then GoTo Finally
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblOrderDueDates WHERE CustomerOrderID = " & CustomerOrderID & " ORDER BY [Timestamp] DESC")
    Dim CustomerDueDate
    
    If rs.EOF Then
        GetLastAgreedDueDate = ELookup("tblCustomerOrders", "CustomerOrderID = " & CustomerOrderID, "CustomerDueDate")
    Else
        labelCaption = "Liefertermin Neu"
        GetLastAgreedDueDate = rs.fields("CustomerDueDate")
    End If

'    If Not frm Is Nothing Then
'        If Not isFalse(GetLastAgreedDueDate) Then frm("subform")("LastAgreedDueDate") = EscapeString(GetLastAgreedDueDate, "tblCustomerOrders", "LastAgreedDueDate")
'    Else
'        If Not isFalse(GetLastAgreedDueDate) Then
'            RunSQL "UPDATE tblCustomerOrders SET LastAgreedDueDate = " & EscapeString(GetLastAgreedDueDate, "tblCustomerOrders", "LastAgreedDueDate") & " WHERE " & _
'                "CustomerOrderID = " & CustomerOrderID
'        End If
'    End If

Finally:
    Dim frm As Form: Set frm = GetForm("frmSubSupplierManagementMain")
    If Not frm Is Nothing Then
        frm("subform")("lblLastAgreedDueDate").Caption = labelCaption
    End If
    
End Function

Public Function dshtOrderDueDates_AfterUpdate(frm As Form)
    
    Dim CustomerOrderID: CustomerOrderID = frm.parent("CustomerOrderID")
    If isFalse(CustomerOrderID) Then Exit Function
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblOrderDueDates WHERE CustomerOrderID = " & CustomerOrderID & " ORDER BY [Timestamp] DESC")
    Dim CustomerDueDate
    If rs.EOF Then
        If IsFormOpen("frmCustomerOrders") Then
            Set frm = Forms("frmCustomerOrders")("subform").Form
            frm("LastAgreedDueDate") = frm("CustomerDueDate")
        End If
        Exit Function
    End If
    
    CustomerDueDate = rs.fields("CustomerDueDate")
    
    If IsFormOpen("frmCustomerOrders") Then
        Set frm = Forms("frmCustomerOrders")("subform").Form
        frm("LastAgreedDueDate") = CustomerDueDate
        ''frm("CustomerDuedate") = CustomerDueDate
        Forms("frmCustomerOrders")("subform1").Form("txtLastAgreedDueDate").Requery
    End If
    
    If IsFormOpen("frmSubSupplierManagementMain") Then
        Set frm = Forms("frmSubSupplierManagementMain")("subform").Form
        frm.Requery
    End If
    
End Function

Public Function dshtOrderDueDates_remove_focus(frm As Form)

    frm.parent("OrderDate").SetFocus

End Function
