Attribute VB_Name = "MaterialDelivery Mod"
Option Compare Database
Option Explicit

Public Function MaterialDeliveryCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            frm.OnCurrent = "=frmMaterialDeliveries_OnCurrent([Form])"
            frm("MaterialDeliveryDate") = "=MaterialDeliveryDate_AfterUpdate([Form])"
        Case 5: ''Datasheet Form
            frm.AllowAdditions = False
            frm.AllowEdits = False
            frm.AllowDeletions = False
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function MaterialDeliveryValidation(frm As Form) As Boolean
    
    Dim MaterialDeliveryID: MaterialDeliveryID = frm("MaterialDeliveryID")
    If isFalse(MaterialDeliveryID) Then
        MaterialDeliveryValidation = True
        Exit Function
    End If
    Dim Quantity: Quantity = frm("Quantity")
    Dim ReleasedQTY: ReleasedQTY = ESum2("tblOrderAssignmentMaterialDeliveries", "MaterialDeliveryID = " & MaterialDeliveryID, "QTY")
    
    If Quantity < ReleasedQTY Then
        MsgBox "Quantity will be less than the total delivery to sub-suppliers.", vbCritical + vbOKOnly
        MaterialDeliveryValidation = False
        frm.Undo
        DoCmd.CancelEvent
        Exit Function
    End If
    
    MaterialDeliveryValidation = True
End Function

Public Function Open_mainMaterialDeliveryMultiSelector_OnCurrent(frm As Form)

    Dim MaterialID: MaterialID = frm("MaterialID")
    Dim filterStr: filterStr = "MaterialDeliveryID = 0"
    
    If Not isFalse(MaterialID) Then
       filterStr = "MaterialID = " & MaterialID
        
    End If
    
    Dim sqlStr: sqlStr = "SELECT MaterialDeliveryID,MaterialControlNumber FROM qryMaterialDeliveries WHERE " & filterStr & " ORDER BY MaterialControlNumber"
    frm("MaterialDeliveryID").RowSource = sqlStr
    
End Function

''Rename the Open_mainMaterialDeliveryMultiSelector to mainMaterialDeliveryMultiSelector
Public Function Open_mainMaterialDeliveryMultiSelector(frm As Form)
    
    Dim MaterialID: MaterialID = frm("MaterialID")
    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    
    If isFalse(MaterialID) Then
        MsgBox "Please select a valid material first.", vbOKOnly
        Exit Function
    End If
    
    DoCmd.OpenForm "mainMaterialDeliveryMultiSelector", , , "MaterialID = " & MaterialID
    
    ''TODO:Sync the tblOrderAssignmentMaterialDeliveries with tblMaterialDeliveryMultiSelectors via the
    ''OrderAssignmentID
    Sync_tblMaterialDeliveryMultiSelectors_tblOrderAssignmentMaterialDeliveries MaterialID, OrderAssignmentID
    
    If Not isFalse(OrderAssignmentID) Then
        Forms("mainMaterialDeliveryMultiSelector")("OrderAssignmentID") = OrderAssignmentID
    End If
    
    RequeryForm "mainMaterialDeliveryMultiSelector", "subform"
    
End Function

Private Sub Sync_tblMaterialDeliveryMultiSelectors_tblOrderAssignmentMaterialDeliveries(MaterialID, OrderAssignmentID)
    
    RunSQL "DELETE FROM tblMaterialDeliveryMultiSelectors"
    
    Dim sqlStr: sqlStr = "SELECT * FROM qryMaterialDeliveries " & _
        " WHERE MaterialID = " & MaterialID & " ORDER BY MaterialControlNumber"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim fields As New clsArray: fields.arr = "MaterialDeliveryID,QTY"
    Dim fieldValues As New clsArray
    Do Until rs.EOF
        
        Set fieldValues = New clsArray
        
        Dim MaterialDeliveryID: MaterialDeliveryID = rs.fields("MaterialDeliveryID")
        
        Dim QTY: QTY = ELookup("tblOrderAssignmentMaterialDeliveries", "OrderAssignmentID = " & OrderAssignmentID & _
            " AND MaterialDeliveryID = " & MaterialDeliveryID, "QTY")
        
        fieldValues.Add MaterialDeliveryID
        fieldValues.Add IIf(isFalse(QTY), 0, QTY)
        UpsertRecord "tblMaterialDeliveryMultiSelectors", fields, fieldValues
        
        rs.MoveNext
    Loop
    
End Sub

Public Function Open_mainMaterialDeliveryMultiSelector_cmdConfirm_OnClick(frm As Form)

    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    
    If isFalse(OrderAssignmentID) Then Exit Function
    
    
    Dim fields As New clsArray: fields.arr = "OrderAssignmentID,MaterialDeliveryID,QTY"
    Dim fieldValues As New clsArray
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblMaterialDeliveryMultiSelectors WHERE QTY > 0"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim TotalQTY: TotalQTY = 0
    
    Do Until rs.EOF
        Dim MaterialDeliveryID: MaterialDeliveryID = rs.fields("MaterialDeliveryID")
        Dim QTY: QTY = rs.fields("QTY")
        
        Set fieldValues = New clsArray
        fieldValues.Add OrderAssignmentID
        fieldValues.Add MaterialDeliveryID
        fieldValues.Add QTY
        
        UpsertRecord "tblOrderAssignmentMaterialDeliveries", fields, fieldValues, "OrderAssignmentID = " & OrderAssignmentID & _
            " AND MaterialDeliveryID = " & MaterialDeliveryID
            
        TotalQTY = TotalQTY + QTY
        rs.MoveNext
    Loop
    
    If IsFormOpen("frmSubSupplierManagementMain") Then
        Forms("frmSubSupplierManagementMain")("subform").Form("subOrderAssignments").Form("MaterialQuantity") = TotalQTY
        Forms("frmSubSupplierManagementMain")("subform").Form("subOrderAssignments").Form("lblcmdMaterialDeliveryID").Requery
    End If
    
    RequeryForm "frmCustomerOrderReports"
    
    DoCmd.Close acForm, frm.Name, acSaveNo
    
End Function

Public Function MaterialDeliveryDate_AfterUpdate(frm As Form)

    Dim ControlNumber: ControlNumber = frm("ControlNumber")
    
    If Not isFalse(ControlNumber) Then
        Exit Function
    End If
    
    Dim DeliveryDate: DeliveryDate = frm("DeliveryDate")
    If isFalse(DeliveryDate) Then
        frm("Controlnumber") = Null
        Exit Function
    End If
    
    ControlNumber = Emax("tblMaterialDeliveries", "Year(DeliveryDate) = " & Year(DeliveryDate), "ControlNumber") + 1
    frm("ControlNumber") = ControlNumber
    
End Function

Public Function frmMaterialDeliveries_OnCurrent(frm As Form)
    
    SetFocusOnForm frm, "cboMaterialName"
    If frm.NewRecord Then
        frm("Controlnumber") = Null
        frm("cboMaterialName") = Null
        frm("cboMaterialQuality") = Null
    Else
        frm("cboMaterialName") = frm("MaterialName")
        frm("cboMaterialQuality") = frm("MaterialQuality")
    End If
    
    frmMaterialDeliveries_set_subform_Visibility frm
    SetMaterialSupplierID_RowSource frm
    
End Function

Public Function frmMaterialDeliveries_set_subform_Visibility(frm As Form)
    
    Dim MaterialID: MaterialID = frm("MaterialID")
    Dim isVisible: isVisible = Not isFalse(MaterialID)
    
    
    
    frm("subform").Visible = isVisible
    frm("Label49").Visible = isVisible ''KommNr. Label
    
End Function

Public Function GetMaterialControlNumber(DeliveryDate, ControlNumber) As String

    If isFalse(DeliveryDate) Then Exit Function
    If isFalse(ControlNumber) Then Exit Function
    
    GetMaterialControlNumber = Format(DeliveryDate, "YYYY") & "M" & Format(ControlNumber, "000")
    
End Function

Public Function MaterialSearchAfterUpdate(frm As Form, Optional AsMaterialName As Boolean = True)
    
    Dim fieldName: fieldName = IIf(AsMaterialName, "MaterialName", "MaterialQuality")
    Dim oppositeFieldName: oppositeFieldName = IIf(fieldName = "MaterialName", "MaterialQuality", "MaterialName")
    Dim sqlStr: sqlStr = "SELECT MaterialID,MaterialName,MaterialQuality,MaterialDescription FROM tblMaterials"
    Dim fieldValue: fieldValue = frm("cbo" & fieldName)
    
    If isFalse(fieldValue) Then
        frm("MaterialID") = Null
        frmMaterialDeliveries_set_subform_Visibility frm
        Exit Function
    End If
    
    sqlStr = sqlStr & " WHERE " & fieldName & " = " & Esc(fieldValue)
    
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If rs.EOF Then
        frm("MaterialID") = Null
        MsgBox "Bitte zuerst Material anlegen > Menu / Material", vbCritical + vbOKOnly
        frm("cbo" & fieldName) = Null
        frm("cbo" & fieldName).SetFocus
        frmMaterialDeliveries_set_subform_Visibility frm
        Exit Function
    End If
    
    frm("MaterialID") = rs.fields("MaterialID")
    frm("cbo" & oppositeFieldName) = rs.fields(oppositeFieldName)
    frmMaterialDeliveries_set_subform_Visibility frm
    
    SetMaterialSupplierID_RowSource frm
    
    
End Function

Private Function SetMaterialSupplierID_RowSource(frm As Form)

    Dim MaterialID: MaterialID = frm("MaterialID")
    If isFalse(MaterialID) Then MaterialID = 0
    
    Dim sqlStr: sqlStr = "SELECT MaterialSupplierID,ShortName FROM qryMaterialSupplierMaterials WHERE MaterialID = " & MaterialID
    frm("MaterialSupplierID").RowSource = sqlStr
    
End Function
