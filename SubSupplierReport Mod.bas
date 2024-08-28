Attribute VB_Name = "SubSupplierReport Mod"
Option Compare Database
Option Explicit

Public Function SubSupplierReportCreate(frm As Object, FormTypeID)

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

Public Function frmSubSupplierReports_cmdPrint_OnClick(frm As Form)
    
    Dim rs As Recordset: Set rs = ReturnRecordset("qrySubSupplierReportOrders2")
    If CountRecordset(rs) = 0 Then
        ShowError "There is no record to print."
        Exit Function
    End If
    
    DoCmd.OpenReport "rptSubSupplierReports", acViewPreview
    
End Function


Public Sub rptSubSupplierReports_Create()
    
    Const HEADER = "Sub Supplier Report"
    Const REPORT_NAME = "rptSubSupplierReports"
    Const RECORDSOURCE_NAME = "qrySubSupplierReports"
    
    Dim rpt As Report: Set rpt = CreateReport()
    SetCommonReportProperties rpt
    rpt.recordSource = RECORDSOURCE_NAME
    rpt.Caption = HEADER
    
    Dim ctl As control
    Set ctl = CreateLabelControl(rpt, HEADER, "Header", "Heading1", acPageHeader)
     ''SubLieferant Name VOLL -> FullName
    Set ctl = CreateTextboxControl(rpt, """SubLieferant Name VOLL:  "" & FullName", "txtFullName", "Heading3", acPageHeader)
    
    Set ctl = CreateSubformControl(rpt, "srptSubSupplierReportOrders", "subSubSupplierReportOrders", , "SubSupplierID")
    Offset_ctlPositions rpt, ctl, , InchToTwip(0.25)
    CreateBannerControls rpt
    
    CleanUpReportProperties rpt
    
    RenameFormOrReport rpt.Name, REPORT_NAME
    
    GetFormOrReport REPORT_NAME, True, True
    
End Sub

Private Function Set_SubSupplierID(frm As Form)

    Dim ctl As control
    Set ctl = frm("fltrSubSuppliers")
    frm("SubSupplierID") = ctl
    
End Function

Private Function Set_FullName(frm As Form)

    Dim ctl As control
    Set ctl = frm("fltrSubSuppliers")
    
    If Not isFalse(ctl) Then
        frm("FullName") = ELookup("tblSubSuppliers", "SubSupplierID = " & ctl, "FullName")
    Else
        frm("FullName") = ""
    End If
    
End Function

Public Function frmSubSupplierReports_OnLoad(frm As Form)
    
    Set_subform_RecordSource frm

End Function

Private Function Set_subform_RecordSource(ByVal frm As Form)
    
    Dim SubSupplierID: SubSupplierID = frm("SubSupplierID")
    
    Dim fltrSubSuppliers: fltrSubSuppliers = frm("fltrSubSuppliers")
    Dim fltrSubDueDateFrom: fltrSubDueDateFrom = frm("fltrSubDueDateFrom")
    Dim fltrSubDeliveryDateFrom: fltrSubDeliveryDateFrom = frm("fltrSubDeliveryDateFrom")
    Dim fltrSubDueDateTo: fltrSubDueDateTo = frm("fltrSubDueDateTo")
    Dim fltrSubDeliveryDateTo: fltrSubDeliveryDateTo = frm("fltrSubDeliveryDateTo")
    
    Dim filterStr: filterStr = "SubSupplierID = 0"
    
    Dim filterArr As New clsArray
    If Not isFalse(fltrSubSuppliers) Then
        filterArr.Add "SubSupplierID = " & fltrSubSuppliers
        filterStr = "SubSupplierID = " & SubSupplierID
    End If
    
    If Not isFalse(fltrSubDueDateFrom) Then
        filterArr.Add "SubDueDate <= #" & SQLDate(fltrSubDueDateFrom) & "#"
    End If
    
    If Not isFalse(fltrSubDeliveryDateFrom) Then
        filterArr.Add "SubDeliveryDate <= #" & SQLDate(fltrSubDeliveryDateFrom) & "#"
    End If
    
    If Not isFalse(fltrSubDueDateTo) Then
        filterArr.Add "SubDueDate >= #" & SQLDate(fltrSubDueDateTo) & "#"
    End If
    
    If Not isFalse(fltrSubDeliveryDateTo) Then
        filterArr.Add "SubDeliveryDate >= #" & SQLDate(fltrSubDeliveryDateTo) & "#"
    End If

    Dim sqlStr: sqlStr = "SELECT * FROM tblSubSuppliers WHERE " & filterStr & " ORDER BY SubSupplierID"
    SetQueryDefSQL "qrySubSupplierReports", sqlStr
    
    filterStr = "OrderAssignmentID = 0"
    If filterArr.count > 0 And Not isFalse(fltrSubSuppliers) Then
        filterStr = filterArr.JoinArr(" AND ")
    End If
    
    sqlStr = "SELECT * FROM qrySubSupplierReportOrders WHERE " & filterStr
    
    SetQueryDefSQL "qrySubSupplierReportOrders2", sqlStr
    
    frm("subform").SourceObject = "Report.rptSubSupplierReports"
    
End Function

Public Function frmSubSupplierReports_DateFilters_AfterUpdate(frm As Form, baseName, Optional DateMode = "From")
    
    If DateMode = "From" Then
        CopyFromToToDate frm, baseName
    Else
        CopyFromToIfEarlier frm, baseName
    End If
    
    Set_subform_RecordSource frm
    
End Function

Public Function frmSubSupplierReports_fltrSubSuppliers_AfterUpdate(frm As Form)
    
    Set_SubSupplierID frm
    Set_subform_RecordSource frm
    
End Function

Public Sub SetUp_frmSubSupplierReports()
    
    Dim frmName: frmName = "frmSubSupplierReports"
    
    If Not IsFormOpen(frmName) Then
        DoCmd.OpenForm frmName, acDesign
    End If
    
    Dim frm As Form: Set frm = Forms(frmName)
    
    frm.OnLoad = "=frmSubSupplierReports_OnLoad([Form])"
    
    Dim ctl As control, bannerCtl As TextBox
    
    Dim sqlStr: sqlStr = "SELECT SubSupplierID,ShortName FROM tblSubSuppliers ORDER BY ShortName"
    Set ctl = frm("fltrSubSuppliers")
    ctl.RowSource = sqlStr
    ctl.AfterUpdate = "=frmSubSupplierReports_fltrSubSuppliers_AfterUpdate([Form])"
    
    Dim baseName
    For Each ctl In frm.Controls
        If ctl.Name Like "fltr*From" Then
            baseName = Replace(ctl.Name, "From", "")
            ctl.Format = "Short Date"
            ctl.AfterUpdate = "=frmSubSupplierReports_DateFilters_AfterUpdate([Form]," & Esc(baseName) & ",""From"")"
        End If
        
        If ctl.Name Like "fltr*To" Then
            baseName = Replace(ctl.Name, "To", "")
            ctl.Format = "Short Date"
            ctl.AfterUpdate = "=frmSubSupplierReports_DateFilters_AfterUpdate([Form]," & Esc(baseName) & ",""To"")"
        End If
        
    Next ctl
    
    CreateBannerControls frm
    
End Sub
