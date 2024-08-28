Attribute VB_Name = "DataFix Mod"
Option Compare Database
Option Explicit

Public Function DataFixCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function

Public Sub FixTransactions()
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblTransactions"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim toBeDeleted As New clsArray
     
    Do Until rs.EOF
        rs.Edit
        rs.fields("CreatedBy") = GetRandomID("tblUsers", "UserID")
        
        Dim MinistryTaskID: MinistryTaskID = GetRandomID("tblMinistryTasks", "MinistryTaskID", "MinistryID = " & rs.fields("MinistryID"))
        Dim TransactionID: TransactionID = rs.fields("TransactionID")
        
        If isFalse(MinistryTaskID) Then
            toBeDeleted.Add TransactionID
        Else
            rs.fields("MinistryTaskID") = MinistryTaskID
        End If
        
        rs.Update
        rs.MoveNext
    Loop
    
    
    Dim item
    For Each item In toBeDeleted.arr
        RunSQL "Delete from tblTransactions where TransactionID = " & item
    Next item
    
End Sub
