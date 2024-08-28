Attribute VB_Name = "DeliveryToCustomer Mod"
Option Compare Database
Option Explicit

Public Function DeliveryToCustomerCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
            ''frm("IsChecked").Properties("DatasheetCaption") = " "
            frm.AllowDeletions = False
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function

Public Function DeliveryToCustomerValidation(frm As Form) As Boolean
    
    Dim PCSToDeliver: PCSToDeliver = frm("PCSToDeliver")
    
    If PCSToDeliver <= 0 Then
        ShowError ("Auslieferung Stk. should be greater than 0.")
        frm("PCSToDeliver").SetFocus
        Exit Function
    End If

    DeliveryToCustomerValidation = True
    
End Function
