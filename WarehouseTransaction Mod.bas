Attribute VB_Name = "WarehouseTransaction Mod"
Option Compare Database
Option Explicit

Public Function WarehouseTransactionCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
    End Select

End Function

Public Function tblWarehouseTransactions_InsertInitial(OrderAssignmentID)
    
    If isFalse(OrderAssignmentID) Then Exit Function
    ''There can only be one initial for an OrderAssignment
    Dim fieldArr As New clsArray: fieldArr.arr = "OrderAssignmentID,TransactionType"
    Dim fieldValueArr As New clsArray
    fieldValueArr.Add OrderAssignmentID
    fieldValueArr.Add "Initial"
    UpsertRecord "tblWarehouseTransactions", fieldArr, fieldValueArr, "OrderAssignmentID = " & OrderAssignmentID & " AND TransactionType = ""Initial"""

End Function


