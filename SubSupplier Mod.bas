Attribute VB_Name = "SubSupplier Mod"
Option Compare Database
Option Explicit

Public Function SubSupplierCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            frm.OnCurrent = "=frmSubSuppliers_OnCurrent([Form])"
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function frmSubSuppliers_OnCurrent(frm As Form)

    SetFocusOnForm frm, "ShortName"
    ''InsertTo_tblSubSupplierServices frm
    ''Set_subform1_RecordSource frm
    Dim SubSupplierID: SubSupplierID = frm("SubSupplierID")
    Dim filterStr: filterStr = "SubSupplierID = 0"
    
    If Not isFalse(SubSupplierID) Then
        filterStr = "SubSupplierID = " & SubSupplierID
    End If
    
    Dim sqlStr: sqlStr = "SELECT ServiceID,Service FROM qrySubSupplierServices WHERE " & filterStr
    frm("subform1").Form("ServiceID").RowSource = sqlStr
    frm("subform1").Form.Requery
    
End Function

Private Sub InsertTo_tblSubSupplierServices(frm As Form)
    
    RunSQL "DELETE FROM tblSubSupplierServices"
    
    Dim SubSupplierID: SubSupplierID = frm("SubSupplierID")
    If isFalse(SubSupplierID) Then Exit Sub

    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblOrderAssignments"
          .AddFilter "SubSupplierID = " & SubSupplierID
          .fields = "ServiceID"
          .GroupBy = "ServiceID"
          sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblSubSupplierServices"
          .fields = "ServiceID"
          .insertSQL = sqlStr
          .InsertFilterField = "ServiceID"
          rowsAffected = .Run
    End With
    
    frm("subform").Form.Requery
    
End Sub

Private Sub Set_subform1_RecordSource(frm As Form)
    
    Dim SubSupplierID: SubSupplierID = frm("SubSupplierID")
    
    Dim filters, filterArr As New clsArray
    
    If Not isFalse(SubSupplierID) Then
        filterArr.Add "SubSupplierID = " & SubSupplierID
    End If
    
    Dim ServiceID, ServiceIDs As New clsArray: ServiceIDs.arr = Elookups("tblSubSupplierServices", "IsChecked", "ServiceID")
    
    If ServiceIDs.count > 0 Then
        filterArr.Add "ServiceID In(" & ServiceIDs.JoinArr(",") & ")"
    End If
    
    Dim filterStr: filterStr = IIf(filterArr.count > 0 And ServiceIDs.count > 0, filterArr.JoinArr(" AND "), "SubSupplierID = 0")
    
    Dim sqlStr: sqlStr = "Select * from qryOrderAssignments WHERE " & filterStr
    
    frm("subform1").Form.recordSource = sqlStr
    frm("subform1").Form.Requery
    
End Sub

Public Function contSubSupplierMaterials_IsChecked_AfterUpdate(frm As Form)
    
    Dim SubSupplierServiceID: SubSupplierServiceID = frm("SubSupplierServiceID")
    Dim IsChecked: IsChecked = frm("IsChecked")
    DoCmd.RunCommand acCmdSaveRecord
    Set_subform1_RecordSource frm.parent.Form
    
End Function

