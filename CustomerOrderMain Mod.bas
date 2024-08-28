Attribute VB_Name = "CustomerOrderMain Mod"
Option Compare Database
Option Explicit

Public Function CustomerOrderMainCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            frm.OnCurrent = "=frmCustomerOrders_OnCurrent([Form])"
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Private Function Set_subform1_ProductIDs_RowSource(frm As Form)
    
    Dim CustomerID: CustomerID = frm("CustomerID")
    Dim filterStr: filterStr = "CustomerID = 0"
    If Not isFalse(CustomerID) Then
        filterStr = "CustomerID = " & CustomerID
    End If
    
    Dim field, fieldArr As New clsArray: fieldArr.arr = "CustomerProdNumber,ProductDescription"
    
    Dim i As Integer: i = 2
    
    For Each field In fieldArr.arr
        Dim sqlStr: sqlStr = "SELECT ProductID,[field] FROM qryProducts WHERE " & filterStr & " ORDER BY [field]"
        sqlStr = Replace(sqlStr, "[field]", field)
        frm("subform1").Form("ProductID" & i).RowSource = sqlStr
        i = i + 1
    Next field
    
    frm("subform1").Form.Requery
    
    
End Function

Public Function dshtCustomerOrders_DeliveryDueDate_AfterUpdate(frm As Form)
    
    If areDataValid2(frm, "CustomerOrderMain") Then
        DoCmd.RunCommand acCmdSaveRecord
        SetCustomerDueDate_DefaultValue frm
        
    End If
    
End Function

Public Function SetCustomerDueDate_DefaultValue(frm As Form)
    
    Dim DeliveryDueDate: DeliveryDueDate = frm("DeliveryDueDate")
    
    If Not isFalse(DeliveryDueDate) Then
        frm("subform1").Form("CustomerDueDate").DefaultValue = "=#" & SQLDate(frm("DeliveryDueDate")) & "#"
    Else
        frm("subform1").Form("CustomerDueDate").DefaultValue = "Null"
    End If
    
    frm("subform1").Form("CustomerDueDate").Requery
    
End Function

Public Function frmCustomerOrders_OnCurrent(frm As Form)
        
    SetFocusOnForm frm, "CommissionNumber"
    Dim IsSaved: IsSaved = frm("isSaved")
    
    If IsSaved Then
        frm("subform1").Form("CustomerDueDate").Enabled = False
    End If
    
    SetCustomerDueDate_DefaultValue frm
    Set_subform1_ProductIDs_RowSource frm
    
End Function

Public Function frmCustomerOrders_CustomerID_AfterUpdate(frm As Form)
    
    Set_subform1_ProductIDs_RowSource frm
    
End Function

Public Function frmCustomerOrders_OnLoad(frm As Form)
        
    DefaultFormLoad frm, "fltrCommissionNumber"
    DoCmd.GoToRecord acDataForm, frm.Name, acNewRec
    
End Function





