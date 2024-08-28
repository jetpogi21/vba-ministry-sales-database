Attribute VB_Name = "MaterialSupplier Mod"
Option Compare Database
Option Explicit

Public Function MaterialSupplierCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            frm.OnCurrent = "=frmMaterialSuppliers_OnCurrent([Form])"
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function frmMaterialSuppliers_OnCurrent(frm As Form)
    
    SetFocusOnForm frm, "ShortName"
    
    Dim MaterialSupplierID: MaterialSupplierID = frm("MaterialSupplierID")
    Dim filterStr: filterStr = "MaterialSupplierID = 0"
    
    If Not isFalse(MaterialSupplierID) Then
        filterStr = "MaterialSupplierID = " & MaterialSupplierID
    End If
    
    Dim sqlStr: sqlStr = "SELECT MaterialID,MaterialName FROM qryMaterialSupplierMaterials WHERE " & filterStr
    frm("subform1").Form("MaterialID").RowSource = sqlStr
    frm("subform1").Form.Requery
    ''InsertTo_tblMaterialSupplierMaterials frm
    ''Set_subform1_RecordSource frm
    
End Function


Private Sub InsertTo_tblMaterialSupplierMaterials(frm As Form)
    
    RunSQL "DELETE FROM tblMaterialSupplierMaterials"
    
    Dim MaterialSupplierID: MaterialSupplierID = frm("MaterialSupplierID")
    If isFalse(MaterialSupplierID) Then Exit Sub
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblMaterialDeliveries"
          .AddFilter "MaterialSupplierID = " & MaterialSupplierID
          .fields = "MaterialID"
          .GroupBy = "MaterialID"
          sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblMaterialSupplierMaterials"
          .fields = "MaterialID"
          .insertSQL = sqlStr
          .InsertFilterField = "MaterialID"
          rowsAffected = .Run
    End With
    
    frm("subform").Form.Requery
    
End Sub

Private Sub Set_subform1_RecordSource(frm As Form)
    
    Dim MaterialSupplierID: MaterialSupplierID = frm("MaterialSupplierID")
    
    Dim filters, filterArr As New clsArray
    
    If Not isFalse(MaterialSupplierID) Then
        filterArr.Add "MaterialSupplierID = " & MaterialSupplierID
    End If
    
    Dim MaterialID, MaterialIDs As New clsArray: MaterialIDs.arr = Elookups("tblMaterialSupplierMaterials", "IsChecked", "MaterialID")
    
    If MaterialIDs.count > 0 Then
        filterArr.Add "MaterialID In(" & MaterialIDs.JoinArr(",") & ")"
    End If
    
    Dim filterStr: filterStr = IIf(filterArr.count > 0 And MaterialIDs.count > 0, filterArr.JoinArr(" AND "), "MaterialSupplierID = 0")
    
    Dim sqlStr: sqlStr = "Select * from qryMaterialDeliveries WHERE " & filterStr
    
    frm("subform1").Form.recordSource = sqlStr
    frm("subform1").Form.Requery
    
End Sub

Public Function contMaterialSupplierMaterials_IsChecked_AfterUpdate(frm As Form)
    
    Dim MaterialSupplierMaterialID: MaterialSupplierMaterialID = frm("MaterialSupplierMaterialID")
    Dim IsChecked: IsChecked = frm("IsChecked")
    DoCmd.RunCommand acCmdSaveRecord
    Set_subform1_RecordSource frm.parent.Form
    
End Function
