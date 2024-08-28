Attribute VB_Name = "MaterialReport Mod"
Option Compare Database
Option Explicit

Public Function MaterialReportCreate(frm As Object, FormTypeID)

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

Public Function frmMaterialReports_cmdPrint_OnClick(frm As Form)
    
    Dim rs As Recordset: Set rs = ReturnRecordset("qryMaterialReportMaterialDeliveries2")
    If CountRecordset(rs) = 0 Then
        ShowError "There is no record to print."
        Exit Function
    End If
    
    DoCmd.OpenReport "rptMaterialReports", acViewPreview
    
End Function

Public Sub rptMaterialReports_Create()
    
    Const HEADER = "Material Report"
    Const REPORT_NAME = "rptMaterialReports"
    Const RECORDSOURCE_NAME = "qryMaterialReports"
    
    Dim rpt As Report: Set rpt = CreateReport()
    SetCommonReportProperties rpt
    rpt.recordSource = RECORDSOURCE_NAME
    rpt.Caption = HEADER
    
    Dim ctl As control
    Set ctl = CreateLabelControl(rpt, HEADER, "Header", "Heading1", acPageHeader)
     
    ''MATERIAL --> MaterialReportMaterials
    Set ctl = CreateLabelControl(rpt, "MATERIAL", "MATERIAL", "Heading3")
    Offset_ctlPositions rpt, ctl, , InchToTwip(0.25)
    Set ctl = CreateSubformControl(rpt, "srptMaterialReportMaterials", "subMaterialReportMaterials", , "MaterialID")
    
    
    ''MATERIAL DELIVERIES --> MaterialReportMaterialDeliveries
    Set ctl = CreateLabelControl(rpt, "MATERIAL DELIVERIES", "MATERIALDELIVERIES", "Heading3")
    Offset_ctlPositions rpt, ctl, , InchToTwip(0.25)
    Set ctl = CreateSubformControl(rpt, "srptMaterialReportMaterialDeliveries", "subMaterialReportMaterialDeliveries", , "MaterialID")
    
    CreateBannerControls rpt
    
    CleanUpReportProperties rpt
    
    RenameFormOrReport rpt.Name, REPORT_NAME
    
    GetFormOrReport REPORT_NAME, True, True
    
End Sub

Public Function frmMaterialReports_fltrInStock_AfterUpdate(frm As Form)
    
    Set_fltrMaterialName_RowSource frm
    Set_subform_RecordSource frm
    
End Function

Public Sub Set_fltrMaterialName_RowSource(frm As Form)

    Dim fltrInStock: fltrInStock = frm("fltrInStock")
    Dim sqlStr: sqlStr = "SELECT MaterialID,MaterialName FROM tblMaterials ORDER BY MaterialName"
    If fltrInStock Then
        sqlStr = "SELECT MaterialID,MaterialName FROM qryMaterialReportMaterialDeliveries WHERE QuantityOnStock > 0" & _
            " GROUP BY MaterialID,MaterialName ORDER BY MaterialName,MaterialID"
    End If
    
    frm("fltrMaterialName").RowSource = sqlStr
    
    Dim MaterialID: MaterialID = frm("fltrMaterialName")
    
    If Not isFalse(MaterialID) Then
        Dim rs As Recordset: Set rs = ReturnRecordset(Replace(sqlStr, " GROUP BY", " AND MaterialID = " & MaterialID & " GROUP BY"))
        If rs.EOF Then frm("fltrMaterialName") = Null
    End If
End Sub

Private Sub Set_subCustomerOrderReportSubSuppliers_Filter(frm As Form)
    
    Dim fltrInStock: fltrInStock = frm("fltrInStock")
    
    Dim filterStr: filterStr = "MaterialDeliveryID > 0"
    
    If fltrInStock Then
        filterStr = "QuantityOnStock <> 0"
    End If
    
    frm("subCustomerOrderReportSubSuppliers").Form.Filter = filterStr
    frm("subCustomerOrderReportSubSuppliers").Form.FilterOn = True
    
End Sub

Public Sub SetUp_frmMaterialReports()
    
    Dim frmName: frmName = "frmMaterialReports"
    
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

Private Function Set_subform_RecordSource(ByVal frm As Form)
    
    Dim MaterialID: MaterialID = frm("MaterialID")
    
    Dim fltrMaterialName: fltrMaterialName = frm("fltrMaterialName")
    Dim fltrMaterialQuality: fltrMaterialQuality = frm("fltrMaterialQuality")
    Dim fltrInStock: fltrInStock = frm("fltrInStock")

    Dim filterStr: filterStr = "MaterialID = 0"
    
    Dim filterArr As New clsArray
    If Not isFalse(MaterialID) Then
        filterArr.Add "MaterialID = " & MaterialID
        filterStr = "MaterialID = " & MaterialID
    End If
    
    If Not isFalse(fltrInStock) Then
        filterArr.Add "QuantityOnStock > 0"
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblMaterials WHERE " & filterStr & " ORDER BY MaterialID"
    SetQueryDefSQL "qryMaterialReports", sqlStr
    
    filterStr = "MaterialDeliveryID = 0"
    If filterArr.count > 0 And Not isFalse(MaterialID) Then
        filterStr = filterArr.JoinArr(" AND ")
    End If
    
    sqlStr = "SELECT * FROM qryMaterialReportMaterialDeliveries WHERE " & filterStr
    
    SetQueryDefSQL "qryMaterialReportMaterialDeliveries2", sqlStr
    
    frm("subform").SourceObject = "Report.rptMaterialReports"
    
End Function

Public Function frmMaterialReports_OnLoad(frm As Form)

    Set_subform_RecordSource frm
    
    
End Function

Public Function frmMaterialReports_Filter_AfterUpdate(frm As Form, controlName)
    
    Set_MaterialID frm, controlName
    If controlName = "fltrMaterialQuality" Then
        Set_fltrMaterialName frm
    Else
        Set_fltrMaterialQuality frm
    End If
    Set_subform_RecordSource frm
    
End Function


Private Sub Set_MaterialID(frm As Form, controlName)
    
    frm("MaterialID") = frm(controlName)
    
End Sub


Private Sub Set_fltrMaterialQuality(frm As Form)
    
    Dim MaterialID: MaterialID = frm("MaterialID")
    frm("fltrMaterialQuality") = MaterialID
    
End Sub

Private Sub Set_fltrMaterialName(frm As Form)
    
    Dim MaterialID: MaterialID = frm("MaterialID")
    frm("fltrMaterialName") = MaterialID
    
End Sub
