Attribute VB_Name = "CustomerOrderForceClose Mod"
Option Compare Database
Option Explicit

Public Function CustomerOrderForceCloseCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        
            Dim ctl As control: Set ctl = frm("FCDate")
            ctl.Width = ctl.Width / 2
            frm("lblFCDate").Width = frm("lblFCDate").Width / 2
            
            frm("FCCheck").Left = ctl.Left + ctl.Width + InchToTwip(0.25)
            frm("lblFCCheck").Left = frm("FCCheck").Left + frm("FCCheck").Width
            frm("FCCheck").AfterUpdate = "=frmCustomerOrderForceCloses_frmCustomerOrderForceCloses_AfterUpdate([Form])"
            
            frm.AfterUpdate = "=frmCustomerOrderForceCloses_AfterUpdate([Form])"
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function

Public Function frmCustomerOrderForceCloses_AfterUpdate(frm As Form)
    tblCustomerOrders_Update_OrderStatus frm
End Function

Public Function frmCustomerOrderForceCloses_frmCustomerOrderForceCloses_AfterUpdate(frm As Form)
    
    If areDataValid2(frm, "CustomerOrderForceClose") Then
        DoCmd.RunCommand acCmdSaveRecord
    End If
    
End Function

Private Function tblCustomerOrders_Update_OrderStatus(frm As Form)
    
    Dim fields As New clsArray: fields.arr = "OrderStatus"
    Dim fieldValues As New clsArray
    
    Dim FCCheck: FCCheck = frm("FCCheck")
    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    
    If FCCheck Then
        fieldValues.Add "Closed"
        UpsertRecord "tblCustomerOrders", fields, fieldValues, "CustomerOrderID = " & CustomerOrderID
    Else
        ''Calculate normally...
        Set frm = GetForm("frmCustomerOrders")
        If Not frm Is Nothing Then
            Set frm = frm("subform1").Form
            Open_frmWarehouseManagement frm, True
        End If
        
    End If
    
End Function
