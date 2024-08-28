Attribute VB_Name = "MaterialSupplierMaterial Mod"
Option Compare Database
Option Explicit

Public Function MaterialSupplierMaterialCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
    End Select

End Function

Public Function dshtMaterialSupplierMaterials_AfterUpdate(frm As Form)
    

End Function
