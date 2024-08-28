Attribute VB_Name = "ModelField Mod"
Option Compare Database
Option Explicit

Public Function ModelFieldCreate(frm As Object, FormTypeID)
    
    If FormTypeID = 5 Then ModelFieldDSCreate frm
    Select Case FormTypeID
        Case 4, 5:
            AttachFunctions frm
            'frm("ModelID").RowSource = "SELECT ModelID, Model FROM tblModels WHERE UserQueryFields = 0"
    End Select
    
End Function

Private Sub ModelFieldDSCreate(frm As Form)
    
    frm.OrderBy = "FieldOrder ASC, ModelFieldID ASC"
    frm.SubPageOrder.DefaultValue = ""
    
End Sub

Private Sub AttachFunctions(frm As Form)

    frm.ParentModelID.AfterUpdate = "=ModelFieldParentModelIDChange([Form])"

End Sub

Private Sub UpdateFields(frm As Form, ctl As control)

    frm("VerboseName") = AddSpaces(ctl.Column(1))
    frm("ForeignKey") = ctl.Column(1)
    frm("IsIndexed") = -1
    frm("FieldTypeID") = dbLong
    frm("ModelField") = ctl.Column(1) & "ID"
    
End Sub

Private Function IsParentModelIDNull(frm As Form) As Boolean
    Dim ParentModelID As Variant
    ParentModelID = frm("ParentModelID")
    IsParentModelIDNull = IsNull(ParentModelID)
End Function


Public Function ModelFieldParentModelIDChange(frm As Form)
    
     Dim ctl As control
    Set ctl = frm("ParentModelID")

    If Not IsParentModelIDNull(frm) Then
        UpdateFields frm, ctl
    End If

End Function
