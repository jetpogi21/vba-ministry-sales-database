Attribute VB_Name = "OrderAnalysis Mod"
Option Compare Database
Option Explicit

Public Function OrderAnalysisCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
            frm.AllowAdditions = False
            frm.AllowDeletions = False
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function frmOrderAnalysis_remove_focus(frm As Form)
    
    frm.parent("OrderDate").SetFocus
    
End Function
