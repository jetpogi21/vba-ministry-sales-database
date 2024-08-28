Attribute VB_Name = "Ministry Mod"
Option Compare Database
Option Explicit

Public Function MinistryCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
            
            Create_mainForm_CloseButton frm
            
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function

Public Function Create_mainForm_CloseButton(frm As Object)
    
    CreateButtonControl frm, "Main Menu", "cmdMainMenu", "=CloseThisForm([Form])"
            
    Dim ctl As control
    Set ctl = frm("cmdMainMenu")
    
    Dim Top, Left, Width
    Top = frm("cmdDelete").Top
    Left = GetRight(frm("cmdDelete")) + 50
    
    ctl.Width = frm("cmdDelete").Width
    ctl.Top = Top
    ctl.Left = Left
    
    frm.Section(acDetail).Height = GetBottom(frm("subform")) + 100
      
End Function

