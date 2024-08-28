Attribute VB_Name = "Transaction Mod"
Option Compare Database
Option Explicit

Public Function TransactionCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            
            frm("TransactionDate").Width = frm("TransactionDate").Width / 2
            frm("StandardFee").Width = frm("StandardFee").Width / 2
            frm("StandardFee").Locked = True
            frm("StandardFee").Enabled = False
            frm("ChargedAmount").Width = frm("ChargedAmount").Width / 2
            
            frm.OnCurrent = "=frmTransactions_OnCurrent([Form])"
            frm("MinistryID").AfterUpdate = "=frmTransactions_MinistryID_AfterUpdate([Form])"
            
            frm.OnLoad = "=frmTransactions_OnLoad([Form])"
            
        Case 5: ''Datasheet Form
            frm.AllowAdditions = False
            frm.AllowEdits = False
            frm.AllowDeletions = False
            
            frm("Timestamp").ColumnHidden = True
            frm("Timestamp").Tag = "alwaysHideOnDatasheet"
            
            SetDatasheetCaption2 frm("CreatedBy"), "User"
            
        Case 6: ''Main Form
        
            Create_mainForm_CloseButton frm
            
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function

Public Function frmTransactions_OnLoad(frm As Form)

    DefaultFormLoad frm, "TransactionID"
    If GetIsAdmin Then Exit Function
    
    DoCmd.GoToRecord acDataForm, frm.Name, acNewRec
    
End Function

Public Function frmTransactions_MinistryID_AfterUpdate(frm As Form)
    
    Set_MinistryTaskID_RowSource frm
    
End Function


Public Function frmTransactions_OnCurrent(frm As Form)
    
    SetFocusOnForm frm, "TransactionDate"
    Set_MinistryTaskID_RowSource frm
    
End Function

Private Function Set_MinistryTaskID_RowSource(frm As Form)
    
    Dim MinistryID: MinistryID = frm("MinistryID")
    
    Dim filterStr: filterStr = "MinistryTaskID = 0"
    
    If Not isFalse(MinistryID) Then
        filterStr = "MinistryID = " & MinistryID
    End If
    
    Dim sqlStr: sqlStr = "SELECT MinistryTaskID,Task FROM tblMinistryTasks WHERE " & filterStr
    
    frm("MinistryTaskID").RowSource = sqlStr

End Function
