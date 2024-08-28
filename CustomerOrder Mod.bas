Attribute VB_Name = "CustomerOrder Mod"
Option Compare Database
Option Explicit

Public Function CustomerOrderCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            frm.OnCurrent = "=CustomerOrderOnCurrent([Form])"
            frm("ProductID").AfterUpdate = "=ProductIDAfterUpdate([Form])"
            frm("pgOrderDueDates").Caption = "Customer Due Date Changes"
            frm("pgNonConformities").Caption = "ABWEICHUNG"
            frm("pgOrderAnalysis").Caption = "Analyse"
            frm("CustomerDueDate").AfterUpdate = "=CustomerDueDate_AfterUpdate([Form])"
            Dim ctl As control
            For Each ctl In frm.Controls
                If ctl.ControlType = acSubform Then
                    ctl.Height = ctl.Height * 2 / 3
                End If
            Next ctl
            
            frm("tabCtl").Height = 0
            frm("tabCtl").Height = frm("tabCtl").Height + 60
            
            For Each ctl In frm.Controls
                If ctl.ControlType = acCommandButton Then
                    ctl.Top = frm("tabCtl").Top + frm("tabCtl").Height + InchToTwip(0.25)
                End If
            Next ctl
            
            frm("subOrderDueDates").Form.recordSource = "qryOrderDueDates"
            frm("OrderStatus").AfterUpdate = "=frmCustomerOrders_OrderStatus_AfterUpdate([Form])"
            
        Case 5, 8: ''Datasheet Form
            ''frm("ProductID").AfterUpdate = "=ProductIDAfterUpdate([Form])"
            frm("ProductID1").RowSource = "SELECT ProductID,SESEMSProductNumber FROM qryProducts ORDER BY SESEMSProductNumber"
            frm("ProductID2").RowSource = "SELECT ProductID,CustomerProdNumber FROM qryProducts ORDER BY CustomerProdNumber"
            frm("ProductID3").RowSource = "SELECT ProductID,ProductDescription FROM qryProducts ORDER BY ProductDescription"
            
            frm("ProductID1").AfterUpdate = "=dshtCustomerOrders_Products_AfterUpdate([Form]," & Esc("ProductID1") & " )"
            frm("ProductID2").AfterUpdate = "=dshtCustomerOrders_Products_AfterUpdate([Form]," & Esc("ProductID2") & " )"
            frm("ProductID3").AfterUpdate = "=dshtCustomerOrders_Products_AfterUpdate([Form]," & Esc("ProductID3") & " )"
            
            frm("DeliveryNotes").ControlSource = "=GetDeliveryNotes([CustomerOrderID])"
            frm("LastAgreedDueDate").Enabled = False
            frm("LastAgreedDueDate").Name = "txtLastAgreedDueDate"
            frm("CustomerOrderID2").Enabled = False
            
            frm("CustomerDueDate").AfterUpdate = "=dshtCustomerOrders_CustomerDueDate_AfterUpdate()"
            
            frm("lblPosition").Caption = "POS"
            
            If FormTypeID = 8 Then
                ''Create the SubSupplier Button.
                Dim maxX: maxX = GetMaxX(frm)
                Dim standardHeight: standardHeight = frm("ProductID1").Height
                
                ''Dim ControlSource: ControlSource = "=" & Esc("Sublieferante")
                Dim ControlSource: ControlSource = "=Get_lbl_cmdManageSubSupplierControlSource([Form])"
                CreateContinuousFormButton frm, standardHeight, ControlSource, "lbl_cmdManageSubSupplier", "cmdManageSubSupplier"
                frm("cmdManageSubSupplier").OnClick = "=Open_frmSubSupplierManagementMain([Form])"
                
                ''frm("OrderStatus").Enabled = False
                ''frm("OrderStatus").Locked = True
                
                frm("DeliveryDate").Enabled = False
                frm("DeliveryDate").ControlSource = "=GetDeliveryDate([Form])"
                
                CreateContinuousFormButton frm, standardHeight, ControlSource, "lblcmdProductID2", "cmdProductID2"
                CopyProperties frm, "lblcmdProductID2", "TextControlInTab"
    
                frm("cmdProductID2").OnClick = "=Open_frmCustomerProdNumberSelector([Form])"
                frm("lblcmdProductID2").ControlSource = "=[CustomerProdNumber]"
                
                frm("cmdProductID2").Left = frm("ProductID2").Left
                frm("cmdProductID2").Width = frm("ProductID2").Width
                frm("lblcmdProductID2").Left = frm("ProductID2").Left
                frm("lblcmdProductID2").Width = frm("ProductID2").Width
                frm("ProductID2").Visible = False
                
                CreateContinuousFormButton frm, standardHeight, ControlSource, "lblcmdProductID3", "cmdProductID3"
                CopyProperties frm, "lblcmdProductID3", "TextControlInTab"
                frm("cmdProductID3").OnClick = "=Open_frmProductDescriptionSelector([Form])"
                frm("lblcmdProductID3").ControlSource = "=[ProductDescription]"
                
                frm("cmdProductID3").Left = frm("ProductID3").Left
                frm("cmdProductID3").Width = frm("ProductID3").Width
                frm("lblcmdProductID3").Left = frm("ProductID3").Left
                frm("lblcmdProductID3").Width = frm("ProductID3").Width
                frm("ProductID3").Visible = False
                
                CreateContinuousFormButton frm, standardHeight, "=Get_cmdWH_ControlSource([Form])", "lblWH", "cmdWH"
                frm("cmdWH").OnClick = "=Open_frmWarehouseManagement([Form])"
                frm("cmdWH").Width = frm("cmdWH").Width / 2
                frm("lblWH").Width = frm("lblWH").Width / 2
                
                CreateContinuousFormButton frm, standardHeight, "=Get_contCustomerOrders_lblDelete_ControlSource([Form])", "lblDelete", "cmdDelete"
                frm("cmdDelete").OnClick = "=contCustomerOrders_cmdDelete_OnClick([Form])"
                frm("cmdDelete").Width = frm("cmdDelete").Width / 2
                frm("lblDelete").Width = frm("lblDelete").Width / 2
                CopyProperties frm, "lblDelete", "ReverseTextControlDanger", False
                
                Set ctl = frm("Symbol")
                ctl.ControlSource = "=GetOrderStatusSymbol([OrderStatus])"
                ctl.FontName = "Wingdings"
                ctl.fontSize = 11
                ctl.TextAlign = 2
                
                Dim cond As FormatCondition
                Set cond = ctl.FormatConditions.Add(acExpression, acEqual, "[OrderStatus] = ""Closed""")
                cond.ForeColor = vbGreen
                
                Set cond = ctl.FormatConditions.Add(acExpression, acEqual, "[OrderStatus] = ""Open""")
                cond.ForeColor = vbRed
                
                frm("lblSymbol").Caption = ""
                
                For Each ctl In frm.Controls
                    If ctl.ControlType = acLabel And ctl.ForeColor = 12928318 Then
                        ctl.Height = ctl.Height * 2
                    End If
                Next ctl
                
            End If
            
            frm.Section(acDetail).Height = 0
            
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8:
            
            
    End Select

End Function

Public Function contCustomerOrders_cmdDelete_OnClick(frm As Form)
    
    If frm.NewRecord Then
        frm.Undo
        Exit Function
    End If
    
    Dim lblDelete: lblDelete = frm("lblDelete")
    If isFalse(lblDelete) Then Exit Function
    
    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    If isFalse(CustomerOrderID) Then Exit Function
    
    RunSQL "DELETE FROM tblCustomerOrders WHERE CustomerOrderID = " & CustomerOrderID
    
    frmCustomerOrders_Racalculate
    
End Function

Public Function Get_contCustomerOrders_lblDelete_ControlSource(frm As Form) As String
    
    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    If isFalse(CustomerOrderID) Then Exit Function
    
    If Not isPresent("tblOrderAssignments", "CustomerOrderID = " & CustomerOrderID) Then
        Get_contCustomerOrders_lblDelete_ControlSource = "Delete"
    End If

End Function

Public Function GetOrderStatusSymbol(OrderStatus) As String
    
    If isFalse(OrderStatus) Then Exit Function
    
    GetOrderStatusSymbol = "6"
    If OrderStatus = "Closed" Then
        GetOrderStatusSymbol = "l"
    End If

End Function

Public Function Get_lbl_cmdManageSubSupplierControlSource(frm As Form) As String
    
    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    If isFalse(CustomerOrderID) Then Exit Function

    Get_lbl_cmdManageSubSupplierControlSource = "Sublieferante"
    
End Function

Public Function CustomerOrderValidation(frm As Form) As Boolean
    
    Dim QTY: QTY = frm("Qty")
    If QTY <= 0 Then
        ShowError "Stück should be greater than 0."
        frm("Qty").SetFocus
        Exit Function
    End If
    
    CustomerOrderValidation = True
    
End Function

Public Function frmCustomerProdNumberSelector_OnCurrent(frm As Form)

    Dim CustomerID: CustomerID = frm("CustomerID")
    Dim filterStr: filterStr = "ProductID = 0"
    
    If Not isFalse(CustomerID) Then
       filterStr = "CustomerID = " & CustomerID
    End If
    
    Dim sqlStr: sqlStr = "SELECT ProductID,CustomerProdNumber FROM qryProducts WHERE " & filterStr & " ORDER BY CustomerProdNumber"
    frm("ProductID").RowSource = sqlStr
    
End Function

Public Function frmProductDescriptionSelector_OnCurrent(frm As Form)

    Dim CustomerID: CustomerID = frm("CustomerID")
    Dim filterStr: filterStr = "ProductID = 0"
    
    If Not isFalse(CustomerID) Then
       filterStr = "CustomerID = " & CustomerID
    End If
    
    Dim sqlStr: sqlStr = "SELECT ProductID,ProductDescription FROM qryProducts WHERE " & filterStr & " ORDER BY ProductDescription"
    frm("ProductID").RowSource = sqlStr
    
End Function

Public Function frmCustomerProdNumberSelector_cmdConfirm_OnClick(frm As Form)

    Dim ProductID: ProductID = frm("ProductID")
    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    
    If isFalse(CustomerOrderID) Then Exit Function
    If isFalse(ProductID) Then
        MsgBox "Select a valid product.", vbCritical + vbOKOnly
        Exit Function
    End If
    
    If IsFormOpen("frmCustomerOrders") Then
        Forms("frmCustomerOrders")("subform1").Form("ProductID2") = ProductID
        dshtCustomerOrders_Products_AfterUpdate Forms("frmCustomerOrders")("subform1").Form, "ProductID2"
    Else
        RunSQL "UPDATE tblCustomerOrders SET ProductID2 = " & ProductID & " WHERE CustomerOrderID = " & CustomerOrderID
    End If
    
    DoCmd.Close acForm, frm.Name, acSaveNo
    
End Function

Public Function frmProductDescriptionSelector_cmdConfirm_OnClick(frm As Form)

    Dim ProductID: ProductID = frm("ProductID")
    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    
    If isFalse(CustomerOrderID) Then Exit Function
    If isFalse(ProductID) Then
        MsgBox "Select a valid product.", vbCritical + vbOKOnly
        Exit Function
    End If
    
    If IsFormOpen("frmCustomerOrders") Then
        Forms("frmCustomerOrders")("subform1").Form("ProductID3") = ProductID
        dshtCustomerOrders_Products_AfterUpdate Forms("frmCustomerOrders")("subform1").Form, "ProductID3"
    Else
        RunSQL "UPDATE tblCustomerOrders SET ProductID3 = " & ProductID & " WHERE CustomerOrderID = " & CustomerOrderID
    End If
    
    DoCmd.Close acForm, frm.Name, acSaveNo
    
End Function

Public Function Open_frmCustomerProdNumberSelector(frm As Form)
    
    Dim CustomerID: CustomerID = frm("CustomerID")
    Dim ProductID: ProductID = frm("ProductID")
    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    
    If isFalse(CustomerID) Then
        MsgBox "Create a position first.", vbOKOnly
        Exit Function
    End If
    
    DoCmd.OpenForm "frmCustomerProdNumberSelector", , , "CustomerID = " & CustomerID
    
    If Not isFalse(ProductID) Then
        Forms("frmCustomerProdNumberSelector")("ProductID") = ProductID
        
    End If
    
    If Not isFalse(CustomerOrderID) Then
        Forms("frmCustomerProdNumberSelector")("CustomerOrderID") = CustomerOrderID
    End If
    
End Function

Public Function Open_frmProductDescriptionSelector(frm As Form)
    
    Dim CustomerID: CustomerID = frm("CustomerID")
    Dim ProductID: ProductID = frm("ProductID")
    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    
    If isFalse(CustomerID) Then
        MsgBox "Create a position first.", vbOKOnly
        Exit Function
    End If
    
    DoCmd.OpenForm "frmProductDescriptionSelector", , , "CustomerID = " & CustomerID
    
    If Not isFalse(ProductID) Then
        Forms("frmProductDescriptionSelector")("ProductID") = ProductID
        
    End If
    
    If Not isFalse(CustomerOrderID) Then
        Forms("frmProductDescriptionSelector")("CustomerOrderID") = CustomerOrderID
    End If
    
End Function


Public Function GetDeliveryDate(frm As Form) As Variant
        
    GetDeliveryDate = Null
    
    Dim CustomerOrderID
    
    If frm.Name = "contTempWarehouseTransferToSubSuppliers" Then
        Set frm = GetForm("frmWarehouseManagement")
        If Not frm Is Nothing Then
            CustomerOrderID = frm("fltrCommissionNumber")
        End If
    Else
        CustomerOrderID = frm("CustomerOrderID")
    End If
    
    If isFalse(CustomerOrderID) Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM qryDeliveryToCustomers WHERE CustomerOrderID = " & CustomerOrderID & _
        " ORDER BY OrderAssignmentID"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim dateArr As New clsArray
    
    Do Until rs.EOF
        Dim DeliveryDate: DeliveryDate = CStr(rs.fields("DeliveryDate"))
        dateArr.Add DeliveryDate, True
        rs.MoveNext
    Loop
    
    GetDeliveryDate = dateArr.JoinArr(",")
    
End Function

Public Sub frmCustomerOrders_Racalculate(Optional RunEvenWhenClosed As Boolean = False)
    
    Dim frmName: frmName = "frmCustomerOrders"
    Dim frm As Form
    If IsFormOpen(frmName) Then
        Set frm = Forms(frmName)
    Else
        If RunEvenWhenClosed Then
            DoCmd.OpenForm frmName
            Set frm = Forms(frmName)
        End If
    End If
    
    If frm Is Nothing Then
        Exit Sub
    End If
    
    frm.Requery
    frmCustomerOrders_fltrOrderStatus_AfterUpdate frm
    
End Sub

Private Function Set_fltrCommissionNumber_RowSource(frm As Form)
    
    Dim fltrOrderStatus: fltrOrderStatus = frm("fltrOrderStatus")
    Dim fltrStr: fltrStr = "CustomerOrderMainID > 0"
    
    If fltrOrderStatus = "Open" Then
        fltrStr = "OrderMainStatus = ""Open"""
    End If
    
    Dim CustomerOrderMainID: CustomerOrderMainID = frm("fltrCommissionNumber")
    frm("fltrCommissionNumber").RowSource = "SELECT CustomerOrderMainID FROM tblCustomerOrderMains WHERE " & fltrStr & " ORDER BY CustomerOrderMainID"
    
    If fltrOrderStatus = "Open" And Not isFalse(CustomerOrderMainID) Then
        fltrStr = fltrStr & " AND CustomerOrderMainID = " & CustomerOrderMainID
    End If
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT CustomerOrderMainID FROM tblCustomerOrderMains WHERE " & fltrStr & " ORDER BY CustomerOrderMainID")
    If rs.EOF Then
        frm("fltrCommissionNumber") = Null
    Else
        frm("fltrCommissionNumber") = CustomerOrderMainID
    End If
    
    frm("fltrCommissionNumber").Requery
    
End Function

Private Function Set_fltrCustomerOrderNumber_RowSource(frm As Form)
    
    Dim fltrOrderStatus: fltrOrderStatus = frm("fltrOrderStatus")
    Dim fltrStr: fltrStr = "CustomerOrderMainID > 0"
    
    If fltrOrderStatus = "Open" Then
        fltrStr = "OrderMainStatus = ""Open"""
    End If
    
    Dim CustomerOrderMainID: CustomerOrderMainID = frm("fltrCustomerOrderNumber")
    frm("fltrCustomerOrderNumber").RowSource = "SELECT CustomerOrderMainID,CustomerOrderNumber FROM tblCustomerOrderMains WHERE " & fltrStr & " ORDER BY CustomerOrderNumber"
    
    If fltrOrderStatus = "Open" And Not isFalse(CustomerOrderMainID) Then
        fltrStr = fltrStr & " AND CustomerOrderMainID = " & CustomerOrderMainID
    End If
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT CustomerOrderMainID,CustomerOrderNumber FROM tblCustomerOrderMains WHERE " & fltrStr & " ORDER BY CustomerOrderMainID")
    
    If rs.EOF Then
        frm("fltrCustomerOrderNumber") = Null
    Else
        frm("fltrCustomerOrderNumber") = CustomerOrderMainID
    End If
    
    frm("fltrCustomerOrderNumber").Requery
    
End Function

Public Function frmCustomerOrders_fltrOrderStatus_AfterUpdate(frm As Form)
    
    Set_fltrCommissionNumber_RowSource frm
    Set_fltrCustomerOrderNumber_RowSource frm
    frmCustomerOrders_FindFirst frm
    
End Function

Public Function Open_frmSubSupplierManagementMain(frm As Form)
    
    If Not areDataValid2(frm, "CustomerOrder") Then Exit Function
    
    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    If isFalse(CustomerOrderID) Then Exit Function
    
    DoCmd.RunCommand acCmdSaveRecord
    
    DoCmd.OpenForm "frmSubSupplierManagementMain"
    Set frm = Forms("frmSubSupplierManagementMain")
    
    frm("fltrCommissionNumber") = CustomerOrderID
    
    frmSubSupplierManagementMain_fltrCustomer_AfterUpdate frm
    
End Function

Public Function dshtCustomerOrders_CustomerDueDate_AfterUpdate()
    
    If IsFormOpen("frmSubSupplierManagementMain") Then
        Forms("frmSubSupplierManagementMain")("subform").LastAgreedDueDate.Requery
    End If

End Function

Public Function GetCustomerOrderID(CustomerOrderID, vTimestamp) As String
    
    If isFalse(CustomerOrderID) Then Exit Function
    
    Dim vYear
    If isFalse(vTimestamp) Then
        vYear = Format$(Date, "YY")
    Else
        vYear = Format$(vTimestamp, "YY")
    End If
    
    GetCustomerOrderID = vYear & CustomerOrderID
    
End Function

Public Function dshtCustomerOrders_Products_AfterUpdate(frm As Form, controlName)
    
    Dim ProductControl, ProductControls As New clsArray: ProductControls.arr = "ProductID1,ProductID2,ProductID3"
    Dim ProductID: ProductID = frm(controlName)
    
    If Not isFalse(ProductID) Then
        Dim CustomerPrice: CustomerPrice = ELookup("tblProducts", "ProductID = " & ProductID, "CustomerPrice")
        frm("Price") = CustomerPrice
    End If
    
    frm("ProductID") = ProductID
    For Each ProductControl In ProductControls.arr
        If ProductControl <> controlName Then
            frm(ProductControl) = ProductID
        End If
    Next ProductControl
    
End Function

Public Function frmCustomerOrders_SaveRecord(frm As Form, Optional mode = "") As Boolean
    
    If areDataValid2(frm, "CustomerOrderMain") Then
        Dim IsSaved: IsSaved = frm("IsSaved")
        
        If Not IsSaved Then
           Dim resp: resp = MsgBox("Damit wird der Eintrag gespeichert, samt allen Lieferterminen. Fortfahren?", vbYesNo)
           If resp = vbNo Then
            Exit Function
           End If
        End If
        
        frm("IsSaved") = True
        Select Case mode
            Case "new":
                DoCmd.GoToRecord acDataForm, frm.Name, acNewRec
            Case "close":
                DoCmd.Close acForm, frm.Name, acSaveNo
        End Select
    End If
    
    frmCustomerOrders_SaveRecord = True
    
End Function

Private Function frmCustomerOrders_FindFirst(frm As Form)
    
On Error GoTo ErrHandler
    Dim fltrCommissionNumber: fltrCommissionNumber = frm("fltrCommissionNumber")
    Dim fltrCustomerOrderNumber: fltrCustomerOrderNumber = frm("fltrCustomerOrderNumber")
    
    Dim CustomerOrderMainID: CustomerOrderMainID = fltrCommissionNumber
    If isFalse(CustomerOrderMainID) Then
        CustomerOrderMainID = fltrCustomerOrderNumber
    End If
    
    If isFalse(CustomerOrderMainID) Then
        DoCmd.GoToRecord acDataForm, frm.Name, acNewRec
        Exit Function
    End If
    
    FindFirst frm, "CustomerOrderMainID = " & CustomerOrderMainID
ErrHandler:
    If Err.Number = 2486 Then
        Exit Function
    End If
    
End Function

Public Function frmCustomerOrders_fltrCommissionNumber_AfterUpdate(frm As Form, Optional OppositeControlName = "fltrCustomerOrderNumber")
    
    frm(OppositeControlName) = Null
    frmCustomerOrders_FindFirst frm
   
End Function

Public Function frmCustomerOrders_fltrCommissionNumber_BeforeUpdate(frm As Form)
    
'    If Not frmCustomerOrders_SaveRecord(frm) Then
'        DoCmd.CancelEvent
'    End If
   
End Function

Public Function frmCustomerOrders_OrderStatus_AfterUpdate(frm As Form)

    Dim OrderStatus: OrderStatus = frm("OrderStatus")
    frm("ManualOrderStatus") = OrderStatus
    
End Function

Public Function GetDeliveryNotes(CustomerOrderID) As String

    If isFalse(CustomerOrderID) Then Exit Function
    
    Dim notesArr As New clsArray
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM qryDeliveryToCustomers WHERE " & _
        "CustomerOrderID = " & CustomerOrderID & " ORDER BY DeliveryToCustomerID")
    
    Do Until rs.EOF
        Dim DeliveryNote: DeliveryNote = rs.fields("DeliveryNote")
        notesArr.Add DeliveryNote
        rs.MoveNext
    Loop
    
    If notesArr.count > 0 Then
        GetDeliveryNotes = notesArr.JoinArr(vbCrLf)
    End If
    
End Function

Public Function CustomerDueDate_AfterUpdate(frm As Form)
    
    If frm.NewRecord Then
        frm("LastAgreedDueDate") = frm("CustomerDueDate")
        GoTo SaveRecord
    End If
    
    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    If isFalse(CustomerOrderID) Then
        frm("LastAgreedDueDate") = frm("CustomerDueDate")
        GoTo SaveRecord
    End If
    
    If Not isPresent("tblOrderDueDates", "CustomerOrderID = " & CustomerOrderID) Then
        frm("LastAgreedDueDate") = frm("CustomerDueDate")
        GoTo SaveRecord
    End If
    
    frm("LastAgreedDueDate") = GetLastAgreedDueDate(CustomerOrderID)
    
SaveRecord:
'    If IsFormOpen("frmCustomerOrders") Then
'        Forms("frmCustomerOrders")("subform").Form.Requery
'    End If
    DoCmd.RunCommand acCmdSaveRecord
    
End Function

Public Function CustomerOrderOnCurrent(frm As Form)
    
    SetFocusOnForm frm, "OrderDate"
    
End Function

Public Function ProductIDAfterUpdate(frm As Form)
    
    Dim ProductID: ProductID = frm("ProductID")
    If Not isFalse(ProductID) Then
        Dim CustomerPrice: CustomerPrice = ELookup("tblProducts", "ProductID = " & ProductID, "CustomerPrice")
        frm("Price") = CustomerPrice
    End If
    
End Function

Public Function frmCustomerOrderItems_AfterUpdate(frm As Form)

    ''Requery the parent's fltrCommissionNumber combo box
    frm.parent("fltrCommissionNumber").Requery
    ''Update the parent's txtRecordNumber and
    
End Function
