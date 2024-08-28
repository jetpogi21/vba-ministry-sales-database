Attribute VB_Name = "SubSupplierManagement Mod"
Option Compare Database
Option Explicit

Public Function SubSupplierManagementCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function frmSubSupplierManagementMain_fltrCustomer_AfterUpdate(frm As Form)

    Dim CustomerID: CustomerID = frm("fltrCustomer")
    
    Dim sqlStr: sqlStr = "SELECT CustomerOrderID,CustomerOrderID2 FROM qryCustomerOrders"
    If Not isFalse(CustomerID) Then
        sqlStr = sqlStr & " WHERE CustomerID = " & CustomerID
    End If
    
    sqlStr = sqlStr & " ORDER BY CustomerOrderID2"
    frm("fltrCommissionNumber").RowSource = sqlStr
    
End Function

Public Function GetSubSupplierManagementDueDateWarning(LastAgreedDueDate, MaxSubDueDate) As String
    
    Dim frm As Form: Set frm = Forms("frmSubSupplierManagementMain")("subform").Form
    If isFalse(LastAgreedDueDate) Then
        frm("Text38").Visible = False
    End If
    
    If isFalse(MaxSubDueDate) Then
        frm("Text38").Visible = False
    End If

On Error GoTo ErrHandler:
    If SQLDate(MaxSubDueDate) > SQLDate(LastAgreedDueDate) Then
        GetSubSupplierManagementDueDateWarning = "Einer von Sublieferterminen überschreitet den Liefertermin zum Kunden."
        frm("Text38").Visible = True
    Else
        frm("Text38").Visible = False
    End If

ErrHandler:
    If Err.Number = 57097 Then
        Exit Function
    End If

End Function

Public Function frmSubSupplierManagementMain_OnCurrent(frm As Form)
        
    SetFocusOnForm frm, ""
    ''SetNavigationData frm, False
    
End Function

Public Function frmSubSupplierManagementMain_OnLoad(frm As Form)
    
    DefaultFormLoad frm, "CustomerOrderID"
    DoCmd.GoToRecord acDataForm, frm.Name, acNewRec
    
End Function

Public Function frmSubSupplierManagementMain_txtRecordNumber_AfterUpdate(frm As Form)

    Dim txtRecordNumber: txtRecordNumber = frm("txtRecordNumber")
    
    If isFalse(txtRecordNumber) Then Exit Function
    
    GoToSpecificRecord frm, txtRecordNumber, False
    
End Function

Public Function frmSubSupplierManagement_OnCurrent(frm As Form)

    SetFocusOnForm frm, ""
    
    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    Dim OrderStatus: OrderStatus = frm("OrderStatus")
    
    If isFalse(CustomerOrderID) Then
        frm("subOrderAssignments").Form.AllowAdditions = False
    Else
        ''Check if there is an orderassignment for this. if it not then allowadditions
        If OrderStatus = "Closed" Then
            frm("subOrderAssignments").Form.AllowAdditions = False
        Else
            frm("subOrderAssignments").Form.AllowAdditions = Not isPresent("tblOrderAssignments", "CustomerOrderID = " & CustomerOrderID)
        End If
    End If
    
    ''Set_contOrderAssignments_MaterialDeliveryID_RowSource frm
    contOrderAssignments_ActualQuantity_Set_Default frm("subOrderAssignments").Form
    
End Function

'Public Function Set_contOrderAssignments_MaterialDeliveryID_RowSource(frm As Form)
'
'    Dim sqlStr: sqlStr = "SELECT MaterialDeliveryID,MaterialControlNumber FROM qryMaterialDeliveries  ORDER BY MaterialControlNumber"
'    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
'    If Not isFalse(CustomerOrderID) Then
'        ''Check first if there's a record from the tblMaterialDeliveryCustomerOrders
'        If isPresent("tblMaterialDeliveryCustomerOrders", "CustomerOrderID = " & CustomerOrderID) Then
'            sqlStr = "SELECT MaterialDeliveryID,MaterialControlNumber FROM qryMaterialDeliveryCustomerOrders WHERE CustomerOrderID = " & CustomerOrderID & _
'                " GROUP BY MaterialDeliveryID,MaterialControlNumber ORDER BY MaterialControlNumber,MaterialDeliveryID"
'        End If
'    End If
'
'    frm("subOrderAssignments").Form("MaterialDeliveryID").RowSource = sqlStr
'
'End Function

Public Function contOrderAssignments_MaterialDeliveryID_Requery(frm As Form)
    
    frm("MaterialDeliveryID").Requery
    
End Function
