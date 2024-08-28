Attribute VB_Name = "ModelButton Mod"
Option Compare Database
Option Explicit

Public Function ModelButtonCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function ModelButton_ModelButton_AfterUpdate(frm As Form)
    
    Dim ModelButton: ModelButton = frm("ModelButton"): If ExitIfTrue(isFalse(ModelButton), "ModelButton is empty..") Then Exit Function
    
    If Not isFalse(frm("FunctionName")) Then Exit Function
    Dim FunctionName
    FunctionName = Replace(ModelButton, ".", "_")
    FunctionName = StrConv(FunctionName, vbProperCase)
    frm("FunctionName") = Replace(FunctionName, " ", "")
    
End Function

Public Function OpenButtonModule(frm As Form)

    Dim FunctionName: FunctionName = frm("FunctionName"): If ExitIfTrue(isFalse(FunctionName), "FunctionName is empty..") Then Exit Function
    DoCmd.OpenModule , FunctionName
    
End Function

Public Function ModelButton_FunctionName_AfterUpdate(frm As Form)
    
    Dim FunctionName: FunctionName = frm("FunctionName"): If ExitIfTrue(isFalse(FunctionName), "FunctionName is empty..") Then Exit Function
    
    If Not isFalse(frm("ModelButton")) Then Exit Function
    
    frm("ModelButton") = GetButtonCaptionFromFunctionName(FunctionName)
    
End Function

Public Function GetButtonCaptionFromFunctionName(FunctionName) As String
    
    Dim separatedWords As New clsArray
    separatedWords.arr = SeparateWords(FunctionName)
    GetButtonCaptionFromFunctionName = StrConv(separatedWords.JoinArr(" "), vbProperCase)
    
End Function

Public Function ModelButtonCreateFunction(frm As Form)
    
    Dim ModelButtonID: ModelButtonID = frm("ModelButtonID"): If ExitIfTrue(isFalse(ModelButtonID), "ModelButtonID is empty..") Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM qryModelButtons WHERE ModelButtonID = " & ModelButtonID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim FunctionName: FunctionName = rs.fields("FunctionName"): If ExitIfTrue(isFalse(FunctionName), "FunctionName is empty..") Then Exit Function
    Dim ModelID: ModelID = rs.fields("ModelID"): If ExitIfTrue(isFalse(ModelID), "ModelID is empty..") Then Exit Function
    Dim TableName: TableName = rs.fields("TableName"): If ExitIfTrue(isFalse(TableName), "TableName is empty..") Then Exit Function
    Dim Model: Model = rs.fields("Model"): If ExitIfTrue(isFalse(Model), "Model is empty..") Then Exit Function
    Dim templateName: templateName = rs.fields("TemplateName")
    Dim ModelButton: ModelButton = rs.fields("ModelButton"): If ExitIfTrue(isFalse(ModelButton), "ModelButton is empty..") Then Exit Function
    
    Dim PrimaryKey: PrimaryKey = GetPrimaryKeyFromTable(ModelID)
    
    Dim strFunction As String
    strFunction = "''Command Name: " & ModelButton & vbCrLf & _
                    "Public Function " & FunctionName & "(frm As Form, Optional " & PrimaryKey & " = """")" & vbCrLf & _
                    vbCrLf & _
                    "    RunCommandSaveRecord" & vbCrLf & _
                    vbCrLf & _
                    "    If isFalse(" & PrimaryKey & ") Then" & vbCrLf & _
                    "        " & PrimaryKey & " = frm(""" & PrimaryKey & """)" & vbCrLf & _
                    "        If ExitIfTrue(isFalse(" & PrimaryKey & "), """ & PrimaryKey & " is empty.."") Then Exit Function" & vbCrLf & _
                    "    End If" & vbCrLf & _
                    vbCrLf & _
                    "    Dim lines As New clsArray" & vbCrLf & _
                    "    Dim sqlStr: sqlStr = ""SELECT * FROM " & TableName & " WHERE " & PrimaryKey & " = "" & " & PrimaryKey & vbCrLf & _
                    "    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)" & vbCrLf & _
                    vbCrLf & _
                    "    " & FunctionName & " = GetReplacedTemplate(rs, """ & templateName & """)" & vbCrLf & _
                    "    " & FunctionName & " = GetGeneratedByFunctionSnippet(" & FunctionName & "," & Esc(FunctionName) & "," & Esc(templateName) & ")" & vbCrLf & _
                    "    CopyToClipboard " & FunctionName & vbCrLf & _
                    "End Function"
    
    Dim moduleName: moduleName = Model & " Mod"
    AddFunctionToModule moduleName, Replace(strFunction, "tbl", "qry")
    
    OpenModule moduleName, FunctionName
    
End Function

Public Sub AddFunctionToModule(moduleName, strFunction)
    
    Dim code As String
    Dim lineNum As Long
    Dim modObject As Module
    
    ' Set the code for the new function
    code = strFunction
    
'    Dim item
'    For Each item In Application.Modules
'        Debug.Print item.name
'    Next item
    
    Set modObject = Application.Modules(moduleName)
    ' Find the last line of code in the module
    lineNum = modObject.CountOfLines

    ' Insert the new function code into the module
    modObject.InsertLines lineNum + 1, code
    
End Sub
