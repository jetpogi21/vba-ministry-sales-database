Attribute VB_Name = "OrderAssignment Mod"
Option Compare Database
Option Explicit

Public Function OrderAssignmentCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
            frm("MaterialDeliveryID").RowSource = "SELECT MaterialDeliveryID,MaterialControlNumber FROM qryMaterialDeliveries  ORDER BY MaterialControlNumber"
            frm("Cost").AfterUpdate = "=dshtOrderAssignments_Cost_AfterUpdate([Form])"
            frm("MaterialQuantity").AfterUpdate = "=dshtOrderAssignments_MaterialQuantity_AfterUpdate([Form])"
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
            frm("lblOrderAssignmentOrder").Caption = ""
            frm("lblQualityControlCaption").Caption = ""
            ''frm("MaterialDeliveryID").RowSource = "SELECT MaterialDeliveryID,MaterialControlNumber FROM qryMaterialDeliveries  ORDER BY MaterialControlNumber"
            frm("Cost").AfterUpdate = "=dshtOrderAssignments_Cost_AfterUpdate([Form])"
            ''frm("MaterialQuantity").AfterUpdate = "=dshtOrderAssignments_MaterialQuantity_AfterUpdate([Form])"
            frm("MaterialQuantity").Enabled = False
            frm("lblSymbol").Caption = ""
            
            frm("Symbol").Locked = True
            frm("Symbol").TabStop = False
            ''frm("OutToSubsupplierDate").Locked = True
            frm("QualityControlCaption").Locked = True
            frm("QualityControlCaption").ControlSource = "=GetQualityControlCaption([OrderAssignmentID])"
            
            Dim ctl As control: Set ctl = frm("Symbol")
            ctl.ControlSource = "=""l"""
            ctl.FontName = "Wingdings"
            ctl.fontSize = 11
            ctl.TextAlign = 2
            
            Dim cond As FormatCondition

            Set cond = ctl.FormatConditions.Add(acExpression, acEqual, "[QualityControlStatus] = ""OK""")
            cond.ForeColor = vbGreen
            
            Set cond = ctl.FormatConditions.Add(acExpression, acEqual, "[QualityControlStatus] = ""STOP""")
            cond.ForeColor = vbRed
            
            Set cond = ctl.FormatConditions.Add(acExpression, acEqual, "NOT [OutToSubsupplierDate] IS NULL")
            cond.ForeColor = vbYellow
            
            Set ctl = frm("OutToSubsupplierDate")
            
            Set cond = ctl.FormatConditions.Add(acExpression, acEqual, "[OrderAssignmentOrder] <> 1")
            cond.BackColor = RGB(240, 240, 240)
            cond.ForeColor = RGB(100, 100, 100)
            
            frm.AllowAdditions = False
            frm.AllowDeletions = False
            
            Set ctl = CreateControl(frm.Name, acTextBox, acFooter, , , 0, 0, 0, 0)
            ctl.ControlSource = "=Max([SubDueDate])"
            ctl.Visible = False
            ctl.Name = "MaxSubDueDate"
            
            Dim sqlStr: sqlStr = "SELECT * from qryOrderAssignments ORDER BY OrderAssignmentOrder,OrderAssignmentID"
            frm.recordSource = sqlStr
            
            frm.AfterUpdate = "=contOrderAssignments_AfterUpdate([Form])"
            
            Dim maxX: maxX = GetMaxX(frm)
            Dim standardHeight: standardHeight = frm("Symbol").Height
            Dim ControlSource
            
'            ControlSource = "=iif([QualityControlStatus] = ""STOP"" AND SentBackDate IS NULL," & Esc("Send Back") & "," & Esc("") & ")"
'            CreateContinuousFormButton frm, standardHeight, ControlSource, "lblReturnToPrevious", "cmdButton"
'            frm("cmdButton").OnClick = "=SendBackToPreviousSubSupplier([Form])"
            
            CreateContinuousFormButton frm, standardHeight, "=Get_cmdQC_ControlSource(CustomerOrderID, OrderAssignmentID, OutToSubsupplierDate)", "lblQC", "cmdQC"
            frm("cmdQC").OnClick = "=Open_frmOrderAssignmentsWithMaterialMain([Form])"
            
            CreateContinuousFormButton frm, standardHeight, "=Get_cmdWH_ControlSource([Form])", "lblWH", "cmdWH"
            frm("cmdWH").OnClick = "=Open_frmWarehouseManagement([Form])"
            
            ''Create the ServiceID button to select service id
            CreateContinuousFormButton frm, standardHeight, ControlSource, "lblcmdServiceID", "cmdServiceID"
            CopyProperties frm, "lblcmdServiceID", "TextControlInTab"
            frm("cmdServiceID").OnClick = "=Open_frmSubSupplierServiceSelector([Form])"
            frm("lblcmdServiceID").ControlSource = "=[Service]"
            
            Set ctl = CreateControl(frm.Name, acLabel, acHeader, , , 0, 0)
            ctl.Name = "lblServiceID"
            ctl.Caption = "Dienstleistung"
            CopyProperties frm, "lblServiceID", "LabelControl"
            ctl.InSelection = True
            ctl.TextAlign = 2
            
            frm("Service").Visible = False
            frm("Service").Left = 0
            frm("Service").Top = 0
            frm("lblService").Visible = False
            frm("lblService").Left = 0
            frm("lblService").Top = 0
            
            ''Create the MaterialDeliveryID button to select MaterialDeliveryID
            CreateContinuousFormButton frm, standardHeight, ControlSource, "lblcmdMaterialDeliveryID", "cmdMaterialDeliveryID"
            CopyProperties frm, "lblcmdMaterialDeliveryID", "TextControlInTab"
            frm("cmdMaterialDeliveryID").OnClick = "=Open_mainMaterialDeliveryMultiSelector([Form])"
            frm("lblcmdMaterialDeliveryID").ControlSource = "=GetOrderAssignmentMaterialDeliveries([OrderAssignmentID])"
            
            Set ctl = CreateControl(frm.Name, acLabel, acHeader)
            ctl.Name = "lblMaterialDeliveryID"
            ctl.Caption = "Material CTRL Nr."
            CopyProperties frm, "lblMaterialDeliveryID", "LabelControl"
            ctl.InSelection = True
            ctl.TextAlign = 2

'            frm("lblService").Visible = False
'            frm("lblService").Left = 0
'            frm("lblService").Top = 0
'
            
            frm("MaterialID").AfterUpdate = "=contOrderAssignments_MaterialID_AfterUpdate([Form])"
            frm("ActualQuantity").Visible = False
            frm("OutToSubsupplierDate").Enabled = True
            frm.Section(acDetail).Height = 0
            
            frm("OrderAssignmentOrder").AfterUpdate = "=contOrderAssignments_OrderAssignmentOrder_AfterUpdate([Form])"
    End Select

End Function

Public Function Get_cmdQC_ControlSource(CustomerOrderID, OrderAssignmentID, OutToSubsupplierDate) As String
    
    If isFalse(CustomerOrderID) Then
        Exit Function
    End If
    
    
    Dim FirstOrderAssignmentID: FirstOrderAssignmentID = ELookup("tblOrderAssignments", "CustomerOrderID = " & _
        CustomerOrderID, "OrderAssignmentID", "OrderAssignmentOrder ASC, OrderAssignmentID ASC")
    
    ''This is the first or there's no existing orderassignment
    If CStr(OrderAssignmentID) = FirstOrderAssignmentID Or isFalse(FirstOrderAssignmentID) Then
        Get_cmdQC_ControlSource = "QC"
        Exit Function
    End If
    
    If Not isFalse(OutToSubsupplierDate) Then
        Get_cmdQC_ControlSource = "QC"
    End If
    
End Function

Public Function Get_cmdWH_ControlSource(frm As Form) As String
    
    Dim isCustomerOrder: isCustomerOrder = frm.Name = "contCustomerOrders"
    
    If Not isCustomerOrder Then
        If frm("lblQC") = "" Then Exit Function
        Dim QualityControlStatus: QualityControlStatus = frm("QualityControlStatus")
        If isFalse(QualityControlStatus) Then Exit Function
        Get_cmdWH_ControlSource = "WH"
    Else
        ''Check if at least one OrderAssignment from the CustomerOrderID has undergone Quality Control
        Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
        If isFalse(CustomerOrderID) Then Exit Function
        If isPresent("tblOrderAssignments", "CustomerOrderID = " & CustomerOrderID & " AND Not isFalse(QualityControlStatus)") Then
            Get_cmdWH_ControlSource = "WH"
        End If
    End If
    
End Function

Public Function contOrderAssignments_MaterialID_AfterUpdate(frm As Form)
    
    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    If isFalse(OrderAssignmentID) Then Exit Function
    
    frm("MaterialQuantity") = Null
    RunSQL "DELETE FROM tblOrderAssignmentMaterialDeliveries WHERE OrderAssignmentID = " & OrderAssignmentID
    frm("lblcmdMaterialDeliveryID").Requery

End Function

Public Function GetQualityControlCaption(OrderAssignmentID) As String
    
    If isFalse(OrderAssignmentID) Then Exit Function
    Dim rs2 As Recordset: Set rs2 = ReturnRecordset("SELECT * FROM qryOrderAssignments WHERE OrderAssignmentID = " & OrderAssignmentID)
    If rs2.EOF Then Exit Function
    
    Dim QualityControlStatus: QualityControlStatus = rs2.fields("QualityControlStatus")
    Dim TransferredOutDate: TransferredOutDate = rs2.fields("TransferredOutDate")
    Dim WHTWarehousePlace: WHTWarehousePlace = rs2.fields("WHTWarehousePlace")
    Dim WarehousePlace: WarehousePlace = Coalesce(rs2.fields("WarehousePlace"), "Quarantined")
    Dim WHTQty: WHTQty = Coalesce(rs2.fields("WHTQty"), 0)
    Dim ActualQuantity: ActualQuantity = rs2.fields("ActualQuantity")
    Dim ScrapQty: ScrapQty = Coalesce(rs2.fields("ScrapQty"), 0)
    Dim DCQty: DCQty = Coalesce(rs2.fields("DCQty"), 0)
    Dim ScrapConfirmation: ScrapConfirmation = rs2.fields("ScrapConfirmation")
    Dim CustomerOrderID: CustomerOrderID = rs2.fields("CustomerOrderID")
    Dim TransferredOutQty: TransferredOutQty = Coalesce(rs2.fields("TransferredOutQty"), 0)
    
    Dim OrderStatus: OrderStatus = ELookup("tblcustomerOrders", "CustomerOrderID = " & CustomerOrderID, "OrderStatus")
    
    If OrderStatus = "Closed" Then Exit Function

    If isFalse(QualityControlStatus) Then Exit Function
    ''If Not isFalse(TransferredOutDate) Then Exit Function
    
    Dim rs As Recordset: Set rs = GetNextOrderAssignment(OrderAssignmentID)
    
    Dim Warehouses As New clsArray
    
    If rs.EOF Then
        ActualQuantity = ActualQuantity - ScrapQty - DCQty
        
        If ActualQuantity = 0 Then Exit Function
        ''GetQualityControlCaption = GetQualityControlCaption_Warehouses(WarehousePlace, ActualQuantity, WHTWarehousePlace, WHTQty, DCQty, TransferredOutQty)
        GetQualityControlCaption = GetScrapWarehousePlaces(OrderAssignmentID)
        Exit Function
    End If
    
    If Not rs.EOF Then
        GetQualityControlCaption = GetQualityControlCaption_Warehouses(WarehousePlace, ActualQuantity, WHTWarehousePlace, WHTQty, DCQty, TransferredOutQty, ScrapQty)
'        QualityControlStatus = rs.fields("QualityControlStatus")
'        If isFalse(QualityControlStatus) Then
'            GetQualityControlCaption = GetQualityControlCaption_Warehouses(WarehousePlace, ActualQuantity, WHTWarehousePlace, WHTQty, DCQty)
'        End If
    End If
    
End Function

Private Function GetQualityControlCaption_Warehouses(WarehousePlace, ActualQuantity, WHTWarehousePlace, WHTQty, DCQty, TransferredOutQty, ScrapQty)
    
    Dim Warehouses As New clsArray
    Dim RemainingQTY: RemainingQTY = 0
    
    If (WHTQty = 0) Then
        RemainingQTY = ActualQuantity - TransferredOutQty - ScrapQty - DCQty
        If RemainingQTY > 0 Then
            Warehouses.Add WarehousePlace & ": " & RemainingQTY
        End If
    Else
        RemainingQTY = ActualQuantity - WHTQty - ScrapQty
        If RemainingQTY > 0 Then
            Warehouses.Add WarehousePlace & ": " & RemainingQTY
        End If
        RemainingQTY = WHTQty - DCQty - TransferredOutQty - ScrapQty
        If RemainingQTY > 0 Then
            Warehouses.Add WHTWarehousePlace & ": " & RemainingQTY
        End If
    End If
    
    ''Dim WarehousePlace: WarehousePlace = rs2.fields("WarehousePlace")
    GetQualityControlCaption_Warehouses = Warehouses.JoinArr(" | ")
           
End Function

Public Function GetOrderAssignmentCaption(OrderAssignmentOrder, SupplierShortName) As String

    If isFalse(OrderAssignmentOrder) Or isFalse(SupplierShortName) Then Exit Function
    GetOrderAssignmentCaption = OrderAssignmentOrder & ". " & SupplierShortName
End Function

Public Function Open_frmSubSupplierServiceSelector(frm As Form)
    
    Dim SubSupplierID: SubSupplierID = frm("SubSupplierID")
    Dim ServiceID: ServiceID = frm("ServiceID")
    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    
    If isFalse(SubSupplierID) Then
        MsgBox "Please select a valid subsupplier first.", vbOKOnly
        Exit Function
    End If
    
    DoCmd.OpenForm "frmSubSupplierServiceSelector", , , "SubSupplierID = " & SubSupplierID
    
    If Not isFalse(ServiceID) Then
        Forms("frmSubSupplierServiceSelector")("ServiceID") = ServiceID
        
    End If
    
    If Not isFalse(OrderAssignmentID) Then
        Forms("frmSubSupplierServiceSelector")("OrderAssignmentID") = OrderAssignmentID
    End If
    
End Function

Public Function Open_frmOrderAssignmentsWithMaterialMain(frm As Form)
    
    If frm("lblQC") = "" Then Exit Function
    If Not areDataValid2(frm, "OrderAssignment") Then Exit Function
    
    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    Dim OutToSubsupplierDate: OutToSubsupplierDate = frm("OutToSubsupplierDate")
    
    If isFalse(OrderAssignmentID) Then Exit Function
    
    Dim FirstOrderAssignmentID: FirstOrderAssignmentID = ELookup("tblOrderAssignments", "CustomerOrderID = " & _
        CustomerOrderID, "OrderAssignmentID", "OrderAssignmentOrder ASC, OrderAssignmentID ASC")
    
    If CStr(OrderAssignmentID) = FirstOrderAssignmentID Or isFalse(FirstOrderAssignmentID) Then
'        If isFalse(OutToSubsupplierDate) Then
'            MsgBox "Please enter Datum zum Sublieferant manually first.", vbCritical + vbOKOnly
'            Exit Function
'        End If
        GoTo CanBeOpened
    End If
    
    If isFalse(OutToSubsupplierDate) Then
        MsgBox "The product isn't in this sub-supplier yet. Send it from the previous sub-supplier first.", vbCritical + vbOKOnly
        Exit Function
    End If
    
    ''Check first if the previous OrderAssignment has QualityControl = OK
    Dim rs As Recordset: Set rs = GetPreviousOrderAssignment(OrderAssignmentID)
    
    ''If there's a previous check the QualityControlStatus
    If Not rs.EOF Then
        Dim QualityControlStatus: QualityControlStatus = rs.fields("QualityControlStatus")
        Dim SentBackDate: SentBackDate = rs.fields("SentBackDate")
        Dim TransferredOutDate: TransferredOutDate = rs.fields("TransferredOutDate")
        
        If QualityControlStatus = "STOP" And Not isFalse(TransferredOutDate) Then
            GoTo CanBeOpened
        End If
        
        If (QualityControlStatus <> "OK" Or isFalse(QualityControlStatus)) And isFalse(SentBackDate) Then
            MsgBox "Previous subsupplier didn't pass the Quality Control yet.", vbCritical + vbOKOnly
            Exit Function
        End If
    End If
    
CanBeOpened:
    DoCmd.RunCommand acCmdSaveRecord
    
    DoCmd.OpenForm "frmOrderAssignmentsWithMaterialMain"
    Set frm = Forms("frmOrderAssignmentsWithMaterialMain")
    
    frm("fltrCommissionNumber") = CustomerOrderID
    frmOrderAssignmentsWithMaterial_fltrCommissionNumber_AfterUpdate frm
    
    ''frm("fltrSupplierShortName") = OrderAssignmentID
    If Not isFalse(OrderAssignmentID) Then
        FindFirst frm("subform").Form, "OrderAssignmentID = " & OrderAssignmentID
        SetNavigationData frm, False, , "subsupplier"
    End If
    ''frmOrderAssignmentsWithMaterial_fltrSupplierShortName_AfterUpdate frm
    
End Function

Public Function Open_frmWarehouseManagement(frm As Form, Optional hide As Boolean = False)
    
    Dim isCustomerOrder: isCustomerOrder = frm.Name = "contCustomerOrders"
    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    
    If frm("lblWH") = "" Then Exit Function
    
    If Not isCustomerOrder Then
        If Not areDataValid2(frm, "OrderAssignment") Then Exit Function
    
        Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
        
        If isFalse(OrderAssignmentID) Then Exit Function
        
        Dim sqlStr: sqlStr = "SELECT * FROM tblOrderAssignments WHERE OrderAssignmentID = " & OrderAssignmentID
        Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
        
        If rs.EOF Then
            MsgBox "Product haven't undergone Quality Control yet.", vbCritical + vbOKOnly
            Exit Function
        End If
        
        Dim QualityControlStatus: QualityControlStatus = rs.fields("QualityControlStatus")
        If isFalse(QualityControlStatus) Then
            MsgBox "Product haven't undergone Quality Control yet.", vbCritical + vbOKOnly
            Exit Function
        End If
        
        DoCmd.RunCommand acCmdSaveRecord
    Else
        ''Check if at least one OrderAssignment from the CustomerOrderID has undergone Quality Control
        If isFalse(CustomerOrderID) Then Exit Function
        
        If Not areDataValid2(frm, "CustomerOrder") Then Exit Function
        
        If Not isPresent("tblOrderAssignments", "CustomerOrderID = " & CustomerOrderID & " AND Not isFalse(QualityControlStatus)") Then
            ShowError "Products haven't undergone Quality Control yet."
            Exit Function
        End If
    End If
    
    DoCmd.OpenForm "frmWarehouseManagement", , , , , IIf(hide, acHidden, acWindowNormal)
    Set frm = Forms("frmWarehouseManagement")
    
    frm("fltrCommissionNumber") = CustomerOrderID
    frmWarehouseManagement_fltrCommissionNumber_AfterUpdate frm
    
End Function

Public Function contOrderAssignments_OrderAssignmentOrder_AfterUpdate(frm As Form)
    
    DoCmd.RunCommand acCmdSaveRecord
    frm.Requery
    
End Function

Public Function contOrderAssignments_ActualQuantity_Set_Default(frm As Form)
    
    Dim QTY: QTY = frm.parent.Form("Qty")
    If Not isFalse(QTY) Then
        frm("ActualQuantity").DefaultValue = QTY
    End If
    
End Function

Public Function contOrderAssignments_ServiceID_SetRowSource(frm As Form)
    
    Dim SubSupplierID: SubSupplierID = frm("SubSupplierID")
    Dim sqlStr
    If Not isFalse(SubSupplierID) Then
        sqlStr = "SELECT ServiceID,Service FROM qrySubSupplierServices WHERE SubSupplierID = " & SubSupplierID & " ORDER BY Service"
    Else
        sqlStr = "SELECT ServiceID,Service AS MainField FROM tblServices ORDER BY Service"
    End If
    frm("ServiceID").RowSource = sqlStr
    
End Function

Public Function dshtOrderAssignments_remove_focus(frm As Form)
    
    frm.parent("ShortName").SetFocus
    
End Function

Public Function dshtOrderAssignments_Cost_AfterUpdate(frm As Form)

    Dim Cost: Cost = frm("Cost")
    frm("ActualCost") = Cost
    
End Function

'Public Function dshtOrderAssignments_MaterialQuantity_AfterUpdate(frm As Form)
'
'    Dim MaterialQuantity: MaterialQuantity = frm("MaterialQuantity")
'    frm("ActualQuantity") = MaterialQuantity
'
'End Function

Public Function OrderAssignmentValidation(frm As Form) As Boolean

    ''Dim MaterialDeliveryID: MaterialDeliveryID = frm("MaterialDeliveryID")
    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    Dim MaterialQuantity: MaterialQuantity = frm("MaterialQuantity")
    Dim MaterialID: MaterialID = frm("MaterialID")
    Dim HasMaterialDeliveries: HasMaterialDeliveries = isPresent("tblOrderAssignmentMaterialDeliveries", "OrderAssignmentID = " & _
        OrderAssignmentID)
    
    If Not isFalse(MaterialID) Then
        If Not HasMaterialDeliveries Then
            OrderAssignmentValidation = False
            MsgBox "Material CTRL Nr. is required when Material is provided.", vbCritical
            frm.Undo
            DoCmd.CancelEvent
            Exit Function
        End If
        
        If isFalse(MaterialQuantity) Then
            OrderAssignmentValidation = False
            MsgBox "Material Qty is required when Material is provided.", vbCritical
            Exit Function
        End If
        
    End If
    
    Dim QualityControlStatus: QualityControlStatus = frm("QualityControlStatus")
    If QualityControlStatus = "STOP" Then
        Dim DescriptionOfFailure: DescriptionOfFailure = frm("DescriptionOfFailure")
        If isFalse(DescriptionOfFailure) Then
            MsgBox "Description of Failure is a required field once quality control is ""STOP"".", vbCritical
            frm("Text44").SetFocus
            Exit Function
        End If
    End If
    
    If QualityControlStatus = "OK" Then
        Dim WarehousePlaceID: WarehousePlaceID = frm("WarehousePlaceID")
        If isFalse(WarehousePlaceID) Then
            MsgBox "Warehouse Place is a required field once quality control is ""OK"".", vbCritical
            frm("txtWarehousePlaceID").SetFocus
            Exit Function
        End If
    End If
    
    If frm.Name = "contOrderAssignments" Then
        ''Peek at the previous OrderAssignment
        OrderAssignmentID = frm("OrderAssignmentID")
        Dim OutToSubsupplierDate: OutToSubsupplierDate = frm("OutToSubsupplierDate")
        Dim OldOutToSubsupplierDate: OldOutToSubsupplierDate = frm("OutToSubsupplierDate").oldValue
        
        Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
        
        Dim FirstOrderAssignmentID: FirstOrderAssignmentID = ELookup("tblOrderAssignments", "CustomerOrderID = " & _
        CustomerOrderID, "OrderAssignmentID", "OrderAssignmentOrder ASC, OrderAssignmentID ASC")
    
        If CStr(OrderAssignmentID) <> FirstOrderAssignmentID And Not isFalse(FirstOrderAssignmentID) Then
            If IsNull(OldOutToSubsupplierDate) And Not IsNull(OutToSubsupplierDate) Then
                MsgBox "Datum zum Sublieferant can't be changed from blank when it's not the first sub-supplier.", vbCritical + vbOKOnly
                OrderAssignmentValidation = False
                frm("OutToSubsupplierDate") = Null
                DoCmd.CancelEvent
                Exit Function
            End If
        End If
    
        Dim rs As Recordset: Set rs = GetPreviousOrderAssignment(OrderAssignmentID)
        
        If rs Is Nothing Then
            OrderAssignmentValidation = True
            Exit Function
        End If
        
        If rs.EOF Then
            OrderAssignmentValidation = True
            Exit Function
        End If
        
        Dim SentBackDate: SentBackDate = rs.fields("SentBackDate")
        QualityControlStatus = rs.fields("QualityControlStatus")
        Dim TransferredOutDate: TransferredOutDate = rs.fields("TransferredOutDate")
        
        If QualityControlStatus = "STOP" And Not isFalse(TransferredOutDate) Then
            OrderAssignmentValidation = True
            Exit Function
        End If
        
        If (QualityControlStatus <> "OK" Or isFalse(QualityControlStatus)) And isFalse(SentBackDate) And Not isFalse(OutToSubsupplierDate) Then
            MsgBox "Previous subsupplier didn't pass the Quality Control yet.", vbCritical + vbOKOnly
            OrderAssignmentValidation = False
            DoCmd.CancelEvent
            frm.Undo
            Exit Function
        End If
    End If

    OrderAssignmentValidation = True
    
End Function

Public Function AddNewOrderAssignment(frm As Form)
    
    If frm.NewRecord Then
        MsgBox "You can't add order assignment on empty records.", vbCritical
        Exit Function
    End If
    
    ''Check if the last item in this OrderAssignments is not STOP
    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    If isFalse(CustomerOrderID) Then Exit Function
    
    Dim OrderStatus: OrderStatus = frm("OrderStatus")
    If OrderStatus = "Closed" Then
        ShowError "Komm. Nr. is already closed. Addition not allowed."
        Exit Function
    End If
    
    If isPresent("tblOrderAssignments", "CustomerOrderID = " & CustomerOrderID & " AND DCConfirmation") Then
        MsgBox "Product already delivered to customer. Addition not allowed.", vbCritical + vbOKOnly
        Exit Function
    End If
    
    If Not isPresent("tblOrderAssignments", "CustomerOrderID = " & CustomerOrderID) Then
        GoTo Allowed:
    End If
    
'    Dim QualityControlStatus: QualityControlStatus = ELookup("tblOrderAssignments", "CustomerOrderID = " & CustomerOrderID, "QualityControlStatus", "OrderAssignmentOrder DESC, OrderAssignmentID DESC")
'    If QualityControlStatus = "" Then
'        MsgBox "You can't add a new subsupplier since the last item is ""Waiting for next step"".", vbCritical
'        Exit Function
'    End If
    
'    If Not Is_tblOrderAssignmentAdditionAllowed(frm) Then
'        MsgBox "You can't add a new subsupplier since there are still items on ""Quarantine"".", vbCritical
'        Exit Function
'    End If
 
Allowed:

    Dim MaxOrderAssignmentOrder: MaxOrderAssignmentOrder = Emax("tblOrderAssignments", "CustomerOrderID = " & CustomerOrderID, "OrderAssignmentOrder")
    frm("subOrderAssignments").SetFocus
    Set frm = frm("subOrderAssignments").Form
    frm.AllowAdditions = True
    
    frm.OrderAssignmentOrder.DefaultValue = "=" & EscapeString((MaxOrderAssignmentOrder + 1), "tblOrderAssignments", "OrderAssignmentOrder")
    
    contOrderAssignments_GoToNewRecord frm
    
    
End Function

Public Function contOrderAssignments_GoToNewRecord(frm As Form)
    
    Dim rs As Recordset: Set rs = frm.Recordset
    rs.addNew
    frm("SubSupplierID").SetFocus
    
End Function

Public Function contOrderAssignments_AfterUpdate(frm As Form)

    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    'Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblOrderAssignments WHERE CustomerOrderID = " & CustomerOrderID & _
        " AND OrderAssignmentID <> " & OrderAssignmentID & "  ORDER BY OrderAssignmentOrder DESC,OrderAssignmentID DESC")
    
'    If Not rs.EOF Then
'        rs.Edit
'        rs.fields("OutToSubsupplierDate") = Date
'        rs.fields("QualityControlCaption") = Null
'        rs.Update
'    End If
    
    frm.AllowAdditions = False
    ''This should make the previous QualityControlCaption to be null if this is a new record
    
    If IsFormOpen("frmOrderAssignmentsWithMaterialMain") Then
        Set frm = Forms("frmOrderAssignmentsWithMaterialMain")
        frmOrderAssignmentsWithMaterial_fltrCommissionNumber_AfterUpdate frm
    End If
    
    DoCmd.RunCommand acCmdSaveRecord
    contOrderAssignments_Recalculate CustomerOrderID
      
End Function

Public Function frmSubSupplierManagement_cmdDelete_OnClick(frm As Form)
    
    ''Check wether the OutToSubsupplierDate is still null, QualityControlStatus is null
    Set frm = frm("subOrderAssignments").Form
    
    If frm.NewRecord Then Exit Function
On Error GoTo ErrHandler:
    Dim OutToSubsupplierDate: OutToSubsupplierDate = frm("OutToSubsupplierDate")
    Dim QualityControlStatus: QualityControlStatus = frm("QualityControlStatus")
    
    If isFalse(OutToSubsupplierDate) And isFalse(QualityControlStatus) Then
        Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
        RunSQL "DELETE FROM tblOrderAssignments WHERE OrderAssignmentID = " & OrderAssignmentID
        frm.Requery
    Else
        MsgBox "An action was already performed on this phase.", vbCritical + vbOKOnly, "Deletion Failed.."
        Exit Function
    End If
    
    frm.AllowAdditions = False
    Exit Function
ErrHandler:
    If Err.Number = 2427 Then
        ShowError "There's no record to delete."
        frm.AllowAdditions = False
        Exit Function
    End If
    
End Function

Public Function SendBackToPreviousSubSupplier(frm As Form, TransferredOutQty)
    
    If Not areDataValid(frm, "OrderAssignment") Then Exit Function
    
    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    If isFalse(CustomerOrderID) Then Exit Function
    
    Dim OrderAssignmentOrder: OrderAssignmentOrder = EscapeString(frm("OrderAssignmentOrder"), "tblOrderAssignments", "OrderAssignmentOrder")
    Dim OrderAssignmentID: OrderAssignmentID = frm("OrderAssignmentID")
    Dim QualityControlStatus: QualityControlStatus = frm("QualityControlStatus")
    Dim SentBackDate: SentBackDate = frm("SentBackDate")
    
    If QualityControlStatus <> "STOP" Or isFalse(QualityControlStatus) Then
        ''MsgBox "You're not allowed to send back when not quarantined.", vbOKOnly + vbCritical
        Exit Function
    End If
    
    If Not IsNull(SentBackDate) Then
        Exit Function
    End If
    ''Check first if the one being sent back is on quarantine
'    Dim QualityControlStatus: QualityControlStatus = frm("QualityControlStatus")
'    If QualityControlStatus <> "STOP" Or IsNull(QualityControlStatus) Then
'        MsgBox "Product is not on quarantine so it couldn't be sent back to previous supplier.", vbCritical
'        Exit Function
'    End If
    
    ''Check first if there's a previous subsupplier
'    If Not isPresent("tblOrderAssignments", "CustomerOrderID = " & CustomerOrderID & " AND ((OrderAssignmentOrder = " & OrderAssignmentOrder & " AND OrderAssignmentID < " & OrderAssignmentID & _
'        ") OR OrderAssignmentOrder < 0)") Then
'        MsgBox "There's no previous subsupplier.", vbCritical
'        Exit Function
'    End If
    
    ''The only items that can be sent back are those that are at the end of the subsupplier management
'    Dim LastOrderAssignmentID: LastOrderAssignmentID = ELookup("tblOrderAssignments", "CustomerOrderID = " & CustomerOrderID, "OrderAssignmentID", "OrderAssignmentOrder DESC," & _
'        "OrderAssignmentID DESC")
'
'    If LastOrderAssignmentID <> CStr(OrderAssignmentID) Then
'        MsgBox "There's a subsupplier after this order assignment so this can't be sent back to the previous one", vbCritical
'        Exit Function
'    End If
    RunSQL "UPDATE tblOrderAssignments SET SentBackDate = " & EscapeString(Date, "tblOrderAssignments", "SentBackDate") & " WHERE OrderAssignmentID = " & OrderAssignmentID
'    RunSQL "UPDATE tblOrderAssignments SET OrderAssignmentOrder = OrderAssignmentOrder - 0.002 " & _
'        " WHERE CustomerOrderID = " & CustomerOrderID & " AND ((OrderAssignmentOrder = " & OrderAssignmentOrder & " AND OrderAssignmentID < " & OrderAssignmentID & _
'        ") OR OrderAssignmentOrder < 0)"
    ''If sent back, remove the QualityControlStatus related things for this OrderAssignment and the previous subsupplier (just the copy though)
    
    ''Update OutToSubSupplierDate of all OrderAsignment of this CustomerOrderID
    ''Update all the OrderAssignmentOrder of each OrderAssignment after this OrderAssignmentID by +0.001
    RunSQL "UPDATE tblOrderAssignments SET OrderAssignmentOrder = OrderAssignmentOrder + 0.001 WHERE CustomerOrderID = " & CustomerOrderID & _
        " AND ((OrderAssignmentOrder = " & OrderAssignmentOrder & " AND OrderAssignmentID > " & _
        OrderAssignmentID & ") OR OrderAssignmentOrder > " & OrderAssignmentOrder & ")"
        
'    RunSQL "UPDATE tblOrderAssignments SET QualityControlCaption = Null, OutToSubsupplierDate = Date() WHERE CustomerOrderID = " & CustomerOrderID & _
'        " AND Not isFalse(QualityControlCaption) AND OutToSubsupplierDate IS NULL"
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblOrderAssignments WHERE OrderAssignmentID = " & OrderAssignmentID)
    Dim fld As field, fieldNames As New clsArray, fieldValues As New clsArray
    For Each fld In rs.fields
        If fld.Name Like "Material*" Or fld.Name Like "WHT*" Or fld.Name Like "Transf*" Or fld.Name Like "Scrap*" Or fld.Name Like "DC*" Then
            GoTo NextField
        End If
        Select Case fld.Name
            Case "OrderAssignmentID", "Timestamp", "CreatedBy", "RecordImportID":
            Case "OrderAssignmentOrder":
                fieldNames.Add fld.Name
                fieldValues.Add EscapeString(fld.Value + 0.001, "tblOrderAssignments", fld.Name)
            Case "QualityControlStatus", "OutToSubsupplierDate", "QualityControlCaption", "WarehousePlaceID", "DescriptionOfFailure", "SentBackDate", "SentOutDate":
                fieldNames.Add fld.Name
                fieldValues.Add "Null"
            Case "ActualQuantity":
                fieldNames.Add fld.Name
                fieldValues.Add EscapeString(TransferredOutQty, "tblOrderAssignments", fld.Name)
            Case Else:
                fieldNames.Add fld.Name
                fieldValues.Add EscapeString(fld.Value, "tblOrderAssignments", fld.Name)
        End Select
NextField:
    Next fld
    
    RunSQL "INSERT INTO tblOrderAssignments (" & fieldNames.JoinArr(",") & ") VALUES (" & fieldValues.JoinArr(",") & ")"
    ''QualityControlStatus,OutToSubsupplierDate,QualityControlCaption,WarehousePlaceID,DescriptionOfFailure
    
'    frm("QualityControlStatus") = Null
'    frm("OutToSubsupplierDate") = Null
'    frm("QualityControlCaption") = Null
'    frm("WarehousePlaceID") = Null
'    frm("DescriptionOfFailure") = Null
    ''Reduce all CustomerOrderNumbers by 0.001 incrementing each time and copy the content of the previous subsupplier
    frm.Requery
    
    ''contOrderAssignments_Recalculate CustomerOrderID
    
    ''Set rs = frm.RecordsetClone
    ''rs.MoveNext
    ''frm.Bookmark = rs.Bookmark
    
End Function
