Attribute VB_Name = "Material Mod"
Option Compare Database
Option Explicit

Public Function MaterialCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
           frm.OnLoad = "=frmMaterials_OnLoad([Form])"
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function frmMaterials_OnLoad(frm As Form)
    
    DefaultFormLoad frm, "MaterialID"
    frm("subMaterialSupplierMaterials").Form.AllowEdits = False
    frm("subMaterialSupplierMaterials").Form.AllowAdditions = False
    
End Function


