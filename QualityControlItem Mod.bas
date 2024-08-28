Attribute VB_Name = "QualityControlItem Mod"
Option Compare Database
Option Explicit

Public Function QualityControlItemCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            ''frm.OnCurrent = "=frmOrderAssignmentsWithMaterial_OnCurrent([Form])"
            Dim ctrl, controlsArr As New clsArray: controlsArr.arr = "SupplierShortName,Service,SubDueDate"
            For Each ctrl In controlsArr.arr
                frm(ctrl).Enabled = False
                frm(ctrl).Locked = True
            Next ctrl
            
            frm.AfterUpdate = "=frmQualityControlItems_AfterUpdate([Form])"
            
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function frmOrderAssignmentsWithMaterialMain_Recalculate()
    
    Dim frm As Form: Set frm = GetForm("frmOrderAssignmentsWithMaterialMain")
    
    If Not frm Is Nothing Then frmOrderAssignmentsWithMaterialMain_fltrOrderStatus_AfterUpdate frm
    
End Function

Public Function frmOrderAssignmentsWithMaterialMain_fltrOrderStatus_AfterUpdate(frm As Form)
    
    Set_fltrCommissionNumber_RowSource frm
    frmOrderAssignmentsWithMaterial_fltrCommissionNumber_AfterUpdate frm
    
End Function

Private Function Set_fltrCommissionNumber_RowSource(frm As Form)
    
    Dim fltrOrderStatus: fltrOrderStatus = frm("fltrOrderStatus")
    Dim fltrStr: fltrStr = "CustomerOrderID > 0"
    
    If fltrOrderStatus = "Open" Then
        fltrStr = "OrderStatus = ""Open"""
    End If
    
    Dim CustomerOrderID: CustomerOrderID = frm("fltrCommissionNumber")
    frm("fltrCommissionNumber").RowSource = "SELECT CustomerOrderID,CustomerOrderID2 FROM qryCustomerOrders WHERE " & fltrStr & " ORDER BY CustomerOrderID2"
    
    If fltrOrderStatus = "Open" And Not isFalse(CustomerOrderID) Then
        fltrStr = fltrStr & " AND CustomerOrderID = " & CustomerOrderID
    End If
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT CustomerOrderID FROM qryCustomerOrders WHERE " & fltrStr & " ORDER BY CustomerOrderID2")
    If rs.EOF Then
        frm("fltrCommissionNumber") = Null
    Else
        frm("fltrCommissionNumber") = CustomerOrderID
    End If
    
    frm("fltrCommissionNumber").Requery
    
End Function

Public Function Open_mainOpenCustomerOrderSelector(frm As Form)
    
    Dim TargetForm: TargetForm = frm.Name
    Dim RecordIDName: RecordIDName = "CustomerOrderID"
    Dim DropdownName: DropdownName = "fltrCommissionNumber"
    Dim AfterUpdateCallback: AfterUpdateCallback = "frmOrderAssignmentsWithMaterial_fltrCommissionNumber_AfterUpdate"
    
    DoCmd.OpenForm "mainOpenCustomerOrderSelector"
    Set frm = Forms("mainOpenCustomerOrderSelector")
    
    frm("subform").Form("cmdSelect").OnClick = "=SelectRecordFromSelector([Form]," & Esc(TargetForm) & _
        "," & Esc(RecordIDName) & "," & Esc(DropdownName) & "," & Esc(AfterUpdateCallback) & " )"
    
End Function

Public Function frmQualityControlItems_AfterUpdate(frm As Form)
    
    If IsFormOpen("frmSubSupplierManagementMain") Then
        Set frm = Forms("frmSubSupplierManagementMain")
        frm("subform").Requery
    End If
    
End Function

Private Function ToggleControlVisibility(frm As Form)

    Dim QualityControlStatus: QualityControlStatus = frm("QualityControlStatus")
    Dim ctl As control
    
    Select Case QualityControlStatus
        Case "OK":
            For Each ctl In frm.Controls
                If ctl.Tag Like "*OK*" Then
                    ctl.Visible = True
                ElseIf ctl.Tag Like "*Stop*" Then
                    ctl.Visible = False
                End If
            Next ctl
        Case "Stop":
            For Each ctl In frm.Controls
                If ctl.Tag Like "*Stop*" Then
                    ctl.Visible = True
                ElseIf ctl.Tag Like "*OK*" Then
                    ctl.Visible = False
                End If
            Next ctl
        Case Else:
            For Each ctl In frm.Controls
                If ctl.Tag Like "*Stop*" Or ctl.Tag Like "*OK*" Then
                    ctl.Visible = False
                End If
            Next ctl
        
    End Select
    
End Function

Public Function SetQualityControlStatus(frm As Form, QualityControlStatus)
        
    If frm.NewRecord Then Exit Function
    
    If frm("QualityControlStatus") = QualityControlStatus Then
        Exit Function
    End If
    
    If frm("QualityControlStatus") = "OK" Then
        MsgBox "Status can't be changed.", vbCritical + vbOKOnly
        Exit Function
    End If
    
    If frm("QualityControlStatus") = "Stop" Then
        MsgBox "Status can't be changed.", vbCritical + vbOKOnly
        Exit Function
    End If
    
    frm("QualityControlStatus") = QualityControlStatus
    SetQualityButtonFormat frm
    ToggleControlVisibility frm
    
    If QualityControlStatus = "OK" Then
        ''frm("ActualQuantity") = frm("Qty")
        frm("DescriptionOfFailure") = Null
    ElseIf QualityControlStatus = "Stop" Then
        frm("WarehousePlaceID") = Null
    End If
    
    ''frmQualityControlItems_SetQualityControlCaption frm
    
End Function


Public Function ReleasePreviousProducts(OrderAssignmentID)
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblOrderAssignments WHERE OrderAssignmentID = " & OrderAssignmentID)
    Dim CustomerOrderID: CustomerOrderID = rs.fields("CustomerOrderID")
    Dim OrderAssignmentOrder: OrderAssignmentOrder = EscapeString(rs.fields("OrderAssignmentOrder"), "tblOrderAssignments", "OrderAssignmentOrder")
    
    ''Update the SentOutDate of the previous items
    RunSQL "UPDATE tblOrderAssignments SET SentOutDate = " & EscapeString(Date, "tblOrderAssignments", "SentOutDate") & " WHERE " & _
        "CustomerOrderID = " & CustomerOrderID & " AND ((OrderAssignmentOrder = " & OrderAssignmentOrder & " AND OrderAssignmentID < " & OrderAssignmentID & _
        ") OR OrderAssignmentOrder < " & OrderAssignmentOrder & ") AND SentOutDate IS NULL"
        
End Function

Public Function frmQualityControlItems_txtWarehousePlaceID_AfterUpdate(frm As Form)

    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    Dim ProductID: ProductID = frm("ProductID")
    
    If areDataValid(frm, "OrderAssignment") Then
        'frmQualityControlItems_SetQualityControlCaption frm
        DoCmd.RunCommand acCmdSaveRecord
        'ReleasePreviousProducts OrderAssignmentID
        contOrderAssignments_Recalculate CustomerOrderID
        tblWarehouseTransactions_InsertInitial OrderAssignmentID
        Update_tblProducts_QtyOnStock ProductID
        RequeryForm "frmSubSupplierManagementMain", "subform"
        frmCustomerOrders_Racalculate
    End If
    
End Function


Private Sub SetButtonStyles(frm As Form, ctlName, IsActive As Boolean)
    
    Dim ctl As control
    Set ctl = frm(ctlName)
    
    If IsActive Then
        ctl.Properties("BorderStyle") = 0
        ctl.Properties("BorderWidth") = 0
        ctl.Properties("BackColor") = 10738157
        ctl.Properties("HoverColor") = 12642530
        ctl.Properties("PressedColor") = 12642530
        ctl.Properties("HoverForeColor") = 2500134
        ctl.Properties("PressedForeColor") = 2500134
        ctl.Properties("ForeColor") = 2500134
    Else
        ctl.Properties("BorderStyle") = 1
        ctl.Properties("BorderWidth") = 0
        ctl.Properties("BackColor") = 8298702
        ctl.Properties("HoverColor") = 12642530
        ctl.Properties("PressedColor") = 12642530
        ctl.Properties("HoverForeColor") = 2500134
        ctl.Properties("PressedForeColor") = 2500134
        ctl.Properties("ForeColor") = 2500134
    End If
    
    
End Sub

Public Function SetQualityButtonFormat(frm As Form)

    Dim QualityControlStatus: QualityControlStatus = frm("QualityControlStatus")
    
    Select Case QualityControlStatus
        Case "OK":
            SetButtonStyles frm, "cmdOK", True
            SetButtonStyles frm, "cmdStop", False
        Case "Stop":
            SetButtonStyles frm, "cmdOK", False
            SetButtonStyles frm, "cmdStop", True
        Case Else:
            SetButtonStyles frm, "cmdOK", False
            SetButtonStyles frm, "cmdStop", False
        
    End Select
    
End Function

Public Function frmQualityControlItems_OnCurrent(frm As Form)
    
    ''SetFocusOnForm frm, "fltrCommissionNumber"
    Dim editable, editableArr As New clsArray
    editableArr.arr = "SubDeliveryDate,DeliveryNote,ActualCost,ActualQuantity,DeliveryCost"
    
    For Each editable In editableArr.arr
        frm(editable).Enabled = Not frm.NewRecord
        frm(editable).Locked = frm.NewRecord
    Next editable
    
    SetQualityButtonFormat frm
    ToggleControlVisibility frm
    
    Dim QualityControlStatus: QualityControlStatus = frm("QualityControlStatus")
    Dim IsOK: IsOK = QualityControlStatus = "OK"
    If isFalse(QualityControlStatus) Then
        IsOK = False
    End If
    
    frm("txtWarehousePlaceID").Enabled = Not IsOK
    frm("txtWarehousePlaceID").Locked = IsOK
    
    frm("cmdOK").Visible = Not frm.NewRecord
    frm("cmdStop").Visible = Not frm.NewRecord

End Function

'Public Function frmQualityControlItems_SetQualityControlCaption(frm As Form)
'
'    Dim WarehousePlace: WarehousePlace = frm("WarehousePlace")
'    Dim QualityControlStatus: QualityControlStatus = frm("QualityControlStatus")
'    Dim ActualQuantity: ActualQuantity = frm("ActualQuantity")
'    Dim OutToSubsupplierDate: OutToSubsupplierDate = frm("OutToSubsupplierDate")
'
'    Dim QualityControlCaption: QualityControlCaption = Null
'
'    If Not IsNull(OutToSubsupplierDate) Then
'        frm("QualityControlCaption") = QualityControlCaption
'        Exit Function
'    End If
'
'    Select Case QualityControlStatus
'        Case "OK":
'            QualityControlCaption = WarehousePlace & ": " & Format$(ActualQuantity, "Standard")
'        Case "STOP":
'            QualityControlCaption = "Quarantine: " & Format$(ActualQuantity, "Standard")
'    End Select
'
'    frm("QualityControlCaption") = QualityControlCaption
'
'End Function

Public Function frmOrderAssignmentsWithMaterial_fltrCommissionNumber_AfterUpdate(frm As Form)

    Dim CustomerOrderID: CustomerOrderID = frm("fltrCommissionNumber")
    
    ''frm("fltrSupplierShortName") = Null
    
    ''Set_fltrSupplierShortName_RowSource frm
    
    If Not isFalse(CustomerOrderID) Then
        frm("subform").Form.Filter = "CustomerOrderID = " & CustomerOrderID
        frm("subform").Form.FilterOn = True
    Else
        frm("subform").Form.Filter = "CustomerOrderID = 0"
        frm("subform").Form.FilterOn = True
    End If
    
    SetNavigationData frm, False, , "subsupplier"
    
End Function

'Public Function frmOrderAssignmentsWithMaterialMain_fltrSupplierShortName_RowSource(frm As Form)
'
'    Set_fltrSupplierShortName_RowSource frm
'
'End Function

'Private Sub Set_fltrSupplierShortName_RowSource(frm As Form)
'
'    Dim CustomerOrderID: CustomerOrderID = frm("fltrCommissionNumber")
'
'    If isFalse(CustomerOrderID) Then
'        CustomerOrderID = 0
'    End If
'
'    Dim sqlStr: sqlStr = "SELECT OrderAssignmentID, SupplierShortName FROM qryOrderAssignments WHERE CustomerOrderID = " & _
'        CustomerOrderID & " ORDER BY OrderAssignmentOrder,OrderAssignmentID"
'
'    frm("fltrSupplierShortName").RowSource = sqlStr
'
'End Sub

Private Sub SubformFilterOff(frm As Form)

    frm("subOrderAssignments").Form.FilterOn = False
    
End Sub

Public Function frmOrderAssignmentsWithMaterial_fltrSupplierShortName_AfterUpdate(frm As Form)
    
    Dim OrderAssignmentID: OrderAssignmentID = frm("fltrSupplierShortName")
    
    If Not isFalse(OrderAssignmentID) Then
        FindFirst frm("subform").Form, "OrderAssignmentID = " & OrderAssignmentID
        SetNavigationData frm, False, , "subsupplier"
    End If
    
End Function
'Public Function frmQualityControlCustomerOrder_OnCurrent(frm As Form)
'
'    SetFocusOnForm frm, ""
'    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
'    If isFalse(CustomerOrderID) Then Exit Function
'
'    Set frm = Forms("frmOrderAssignmentsWithMaterialMain")
'
'    frm("txtCustomerOrderID") = CustomerOrderID
'    Dim sqlStr: sqlStr = "SELECT OrderAssignmentID,SupplierShortName FROM qryOrderAssignmentsWithMaterial WHERE CustomerOrderID = " & CustomerOrderID & _
'        " ORDER BY SupplierShortName"
'    frm("fltrSupplierShortName").RowSource = sqlStr
'
'End Function


