Attribute VB_Name = "FinancialReportOrder Mod"
Option Compare Database
Option Explicit

Public Function FinancialReportOrderCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        
            FormatFormAsReport frm, FormTypeID, "CustomerFullName"
            
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function

Public Function frmFinancialReports_OnLoad(frm As Form)

    Set_subform_RecordSource frm
    
End Function

Private Sub Set_subform_RecordSource(frm As Form)

    Dim CustomerOrderID: CustomerOrderID = frm("CustomerOrderID")
    Dim filterStr: filterStr = "CustomerOrderID = 0"
    If Not isFalse(CustomerOrderID) Then
        filterStr = "CustomerOrderID = " & CustomerOrderID
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM qryCustomerOrders WHERE " & filterStr & " ORDER BY CustomerOrderID2"
    SetQueryDefSQL "qryFinancialReports", sqlStr
    
    frm("subform").SourceObject = "Report.rptFinancialReports"
    
End Sub

Public Sub SetUp_frmFinancialReports()
    
    Dim frmName: frmName = "frmFinancialReports"
    
    If Not IsFormOpen(frmName) Then
        DoCmd.OpenForm frmName, acDesign
    End If
    
    Dim frm As Form: Set frm = Forms(frmName)
    
    CreateBannerControls frm
    
    Dim i As Integer: i = 16
    
    For i = 16 To 23
        Dim TextBoxName: TextBoxName = "Text" & i
        If DoesPropertyExists(frm, TextBoxName) Then
            DeleteControl frm.Name, TextBoxName
        End If
    Next i
    
End Sub


Public Function frmFinancialReports_SetCustomerOrderID(frm As Form)

    Dim CustomerOrderID: CustomerOrderID = frm("fltrCommissionNumber")
    
    frm("CustomerOrderID") = CustomerOrderID
    
End Function

Public Function frmFinancialReports_SetfltrCustomerShortName(frm As Form)

    Dim CustomerID: CustomerID = frm("fltrCommissionNumber").Column(2)
    
    frm("fltrCustomerShortName") = CustomerID
    
End Function

Public Function frmFinancialReports_fltrCommissionNumber_AfterUpdate(frm As Form)
    
    frmFinancialReports_SetCustomerOrderID frm
    frmFinancialReports_SetfltrCustomerShortName frm
    frmFinancialReports_fltrOrderDate_SetRowSource frm
    frmFinancialReports_SetfltrOrderDate frm
    Set_subform_RecordSource frm
    ''frmFinancialReports_fltrOrderDate_AfterUpdate frm
    
End Function

Public Function frmFinancialReports_SetfltrOrderDate(frm As Form)

    Dim CustomerOrderID: CustomerOrderID = frm("fltrCommissionNumber")
    
    frm("fltrOrderDate") = CustomerOrderID
    
End Function

Private Function frmFinancialReports_fltrOrderDate_SetRowSource(frm As Form)

    Dim CustomerID: CustomerID = frm("fltrCustomerShortName")
    
    Dim filterStr: filterStr = "CustomerOrderID = 0"
    If Not isFalse(CustomerID) Then
        filterStr = "CustomerID = " & CustomerID
    End If
    
    Dim sqlStr: sqlStr = "SELECT CustomerOrderID, OrderDate FROM qryCustomerOrders WHERE " & filterStr & " ORDER BY OrderDate,CustomerOrderID"
    
    frm("fltrOrderDate").RowSource = sqlStr
    
End Function

Public Function frmFinancialReports_fltrCustomerShortName_AfterUpdate(frm As Form)
    
    frmFinancialReports_fltrOrderDate_SetRowSource frm
    frm("fltrOrderDate") = Null
    
End Function

Public Function frmFinancialReports_fltrOrderDate_AfterUpdate(frm As Form)
    
    Dim CustomerOrderID: CustomerOrderID = frm("fltrOrderDate")
    frm("fltrCommissionNumber") = CustomerOrderID
    
    frmFinancialReports_SetCustomerOrderID frm
    Set_subform_RecordSource frm
    
End Function


