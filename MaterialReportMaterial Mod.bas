Attribute VB_Name = "MaterialReportMaterial Mod"
Option Compare Database
Option Explicit

Public Function MaterialReportMaterialCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            FormatFormAsReport frm, 4, "MaterialName"
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
            FormatFormAsReport frm, 8
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function
