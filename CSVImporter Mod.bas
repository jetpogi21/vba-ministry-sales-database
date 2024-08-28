Attribute VB_Name = "CSVImporter Mod"
Option Compare Database
Option Explicit

Public Function CSVImporterCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function GetCSVFilePath(CSVDirectory, CSVFileName) As String
    
    If isFalse(CSVDirectory) Or isFalse(CSVFileName) Then Exit Function
    GetCSVFilePath = CSVDirectory & CSVFileName & ".csv"
    
End Function

Public Function Set_tblCSVImporters_CSVDirectory(frm As Form)
    
    SelectDEDirectory frm, "CSVDirectory"
    Dim CSVDirectory: CSVDirectory = frm("CSVDirectory")
    Dim CSVImporterID: CSVImporterID = frm("CSVImporterID")
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblCSVImporters WHERE CSVImporterID <> " & CSVImporterID & " ORDER BY CSVImporterID")
    Do Until rs.EOF
        
        Dim CSVFileName: CSVFileName = rs.fields("CSVFileName")
        Dim CSVFilePath: CSVFilePath = GetCSVFilePath(CSVDirectory, CSVFileName)
        rs.Edit
        rs.fields("CSVDirectory") = CSVDirectory
        rs.fields("CSVFilePath") = CSVFilePath
        rs.Update
        
        rs.MoveNext
    Loop
    
    CSVFileName = frm("CSVFileName")
    frm("CSVFilePath") = GetCSVFilePath(CSVDirectory, CSVFileName)
    
End Function
''=SelectDEDirectory([Form],"CSVDirectory")

Public Function Set_frmCSVImporters_CSVFilePath(frm As Form)
    
    Dim CSVDirectory: CSVDirectory = frm("CSVDirectory")
    Dim CSVFileName: CSVFileName = frm("CSVFileName")
    
    frm("CSVFilePath") = GetCSVFilePath(CSVDirectory, CSVFileName)
    
End Function

Public Function GenerateCSVFields(frm As Form)
    
    DoCmd.RunCommand acCmdSaveRecord
    
    Dim CSVImporterID: CSVImporterID = frm("CSVImporterID"): If isFalse(CSVImporterID) Then Exit Function
    Dim CSVFilePath: CSVFilePath = frm("CSVFilePath"): If ExitIfTrue(isFalse(CSVFilePath), "CSV FilePath is empty..") Then Exit Function
    
    ImportCSVToTable CSVFilePath
    
    ''RunSQL "DELETE FROM tblCSVImporterFields WHERE CSVImporterID = " & CSVImporterID
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblCSVData")
    Dim fld As field
    For Each fld In rs.fields
        RunSQL "INSERT INTO tblCSVImporterFields (CSVImporterID,CSVField) VALUES (" & CSVImporterID & "," & Esc(fld.Name) & ")"
    Next fld
    
    frm("subCSVImporterFields").Form.Requery
    
End Function

Public Function GenerateTargetFields(frm As Form)
    
    DoCmd.RunCommand acCmdSaveRecord
    
    Dim CSVImporterID: CSVImporterID = frm("CSVImporterID"): If isFalse(CSVImporterID) Then Exit Function
    Dim DestinationDatabase: DestinationDatabase = frm("DestinationDatabase"): If ExitIfTrue(isFalse(DestinationDatabase), "Destination Database is empty..") Then Exit Function
    Dim TargetTable: TargetTable = frm("TargetTable"): If ExitIfTrue(isFalse(TargetTable), "Target Table is empty..") Then Exit Function
    
    Dim db As DAO.Database: Set db = OpenDatabase(DestinationDatabase)
    Dim rs As Recordset: Set rs = db.OpenRecordset("SELECT * FROM [" & TargetTable & "]")
    
    Dim fld As field, fieldNames As New clsArray
    For Each fld In rs.fields
        RunSQL "INSERT INTO tblTargetFieldLists (TableName,FieldName,FieldTypeID) VALUES (" & Esc(TargetTable) & "," & Esc(fld.Name) & "," & fld.Type & ")"
    Next fld
    
    frm("subCSVImporterFields").Form("TargetField").RowSource = "SELECT FieldName FROM tblTargetFieldLists WHERE TableName = " & Esc(TargetTable) & " ORDER BY FieldName"
    
End Function

Public Function BatchImport_tblCSVDataToTargetTable()
    
    Dim frm As Form
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblCSVImporters ORDER BY ImportOrder")
    
    Do Until rs.EOF
        Dim CSVImporterID: CSVImporterID = rs.fields("CSVImporterID")
        DoCmd.OpenForm "frmCSVImporters", , , "CSVImporterID = " & CSVImporterID
        Set frm = Forms("frmCSVImporters")
        
        Import_tblCSVDataToTargetTable frm, True
        
        rs.MoveNext
    Loop
    
    MsgBox "Data Files imported successfully"
    
End Function

'' DeleteDataBasedOnCSV(Forms("frmCSVImporters"))
Public Function DeleteDataBasedOnCSV(frm As Form, Optional batchMode As Boolean = False)
    
    Dim DestinationDatabase: DestinationDatabase = frm("DestinationDatabase"): If ExitIfTrue(isFalse(DestinationDatabase), "Destination Database is empty..") Then Exit Function
    Dim CSVFileName: CSVFileName = frm("CSVFileName"): If ExitIfTrue(isFalse(CSVFileName), "CSVFileName is empty..") Then Exit Function
    Dim CSVFilePath: CSVFilePath = frm("CSVFilePath"): If ExitIfTrue(isFalse(CSVFilePath), "CSV File Path is empty..") Then Exit Function
    Dim TargetConnector: TargetConnector = frm("TargetConnector"): If ExitIfTrue(isFalse(TargetConnector), "TargetConnector is empty..") Then Exit Function
    Dim TargetTable: TargetTable = frm("TargetTable"): If ExitIfTrue(isFalse(TargetTable), "TargetTable is empty..") Then Exit Function
    Dim SourceConnector: SourceConnector = frm("SourceConnector"): If ExitIfTrue(isFalse(SourceConnector), "SourceConnector is empty..") Then Exit Function
    
    CSVFilePath = Replace(CSVFilePath, CSVFileName, "deleted_" & CSVFileName)
    
    ''Validate FilePath if exists
    If Not fileExists(CSVFilePath) Then
        If Not batchMode Then
            MsgBox Esc(CSVFilePath) & " does not exist."
        End If
        Exit Function
    End If
    
    ImportCSVToTable CSVFilePath
    
    DoCmd.SetWarnings False
    DoCmd.CopyObject DestinationDatabase, "tblCSVData", acTable, "tblCSVData"
    DoCmd.SetWarnings True
    
    Dim pkType: pkType = GetFieldType(TargetTable, TargetConnector)
    Dim pkField: pkField = "tblCSVData.pk_value"
    
    If pkType = "dbDate" Then
        pkField = "CDate(" & pkField & ")"
    ElseIf pkType = "dbText" Then
        pkField = "CStr(" & pkField & ")"
    End If
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblCSVData"
          .fields = pkField & " AS " & TargetConnector
          sqlStr = .sql
    End With
    
    Dim deleteSQL
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "DELETE"
          .Source = "[" & TargetTable & "]"
          .joins.Add GenerateJoinObj(sqlStr, TargetConnector, "temp")
          deleteSQL = .sql
    End With
    
    RunSQLOnBackend DestinationDatabase, deleteSQL
    
    If Not batchMode Then MsgBox TargetTable & " successfully updated."
    
End Function

Private Function Create_tblCSVDataTemp(SourceConnector, DeletedFilePath)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
     .Source = "tblCSVData"
     .fields = "tblCSVData.*, ParseISO8601ToDateTime(tblCSVData.created_at) as valid_created_at"
     sqlStr = .sql
    End With
    
    ''tblCSVAccompanyingData is the deleted table (which may or may not exist)
    
    Dim sqlStr2
    If fileExists(DeletedFilePath) Then
        Set sqlObj = New clsSQL
        With sqlObj
         .Source = "tblCSVAccompanyingData"
         .fields = "tblCSVAccompanyingData.*, ParseISO8601ToDateTime(tblCSVAccompanyingData.created_at) as valid_created_at"
         sqlStr2 = .sql
        End With
    End If
    
    deleteTableIfExists "tblCSVDataTemp"
    
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "MAKE"
          .Source = sqlStr
          .SourceAlias = "temp"
          .MakeTable = "tblCSVDataTemp"
          .fields = "temp.*"
          If Not isFalse(sqlStr2) Then
            .joins.Add GenerateJoinObj(sqlStr2, SourceConnector, "temp2", "pk_value", "LEFT")
            .AddFilter "(temp.valid_created_at > temp2.valid_created_at) OR (temp2.valid_created_at IS NULL and NOT temp.valid_created_at IS NULL)"
          End If
          .Run
    End With
    
End Function

Public Function CreateCSVDataToDestinationDatabase(DestinationDatabase, SourceConnector)
    
    deleteTableIfExists "tblCSVData", DestinationDatabase
    
    Dim DatabasePath: DatabasePath = CurrentProject.path & "\" & CurrentProject.Name
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "MAKE"
          .Source = "tblCSVDataTemp IN " & Esc(DatabasePath)
          .MakeTable = "tblCSVData"
          .fields = "tblCSVDataTemp.*"
          sqlStr = .sql
    End With
    
    RunSQLOnBackend DestinationDatabase, sqlStr
    
End Function

Public Function Import_tblCSVDataToTargetTable(frm As Form, Optional batchMode As Boolean = False)

    DoCmd.RunCommand acCmdSaveRecord

    Dim CSVImporterID: CSVImporterID = frm("CSVImporterID"): If isFalse(CSVImporterID) Then Exit Function
    Dim DestinationDatabase: DestinationDatabase = frm("DestinationDatabase"): If ExitIfTrue(isFalse(DestinationDatabase), "Destination Database is empty..") Then Exit Function
    Dim TargetTable: TargetTable = frm("TargetTable"): If ExitIfTrue(isFalse(TargetTable), "Target Table is empty..") Then Exit Function
    Dim SourceConnector: SourceConnector = frm("SourceConnector"): If ExitIfTrue(isFalse(SourceConnector), "Source Connector is empty..") Then Exit Function
    Dim TargetConnector: TargetConnector = frm("TargetConnector"): If ExitIfTrue(isFalse(TargetConnector), "Target Connector is empty..") Then Exit Function
    Dim CSVFilePath: CSVFilePath = frm("CSVFilePath"): If ExitIfTrue(isFalse(CSVFilePath), "CSV File Path is empty..") Then Exit Function
    Dim CSVFileName: CSVFileName = frm("CSVFileName"): If ExitIfTrue(isFalse(CSVFileName), "CSVFileName is empty..") Then Exit Function
    
    DeleteDataBasedOnCSV frm, batchMode
    
    ''Validate FilePath if exists
    If Not fileExists(CSVFilePath) Then
        If Not batchMode Then
            MsgBox Esc(CSVFilePath) & " does not exist."
        End If
        Exit Function
    End If
    
    Dim DeletedFilePath: DeletedFilePath = Replace(CSVFilePath, CSVFileName, "deleted_" & CSVFileName)
    
    ImportCSVToTable CSVFilePath
    
    If fileExists(DeletedFilePath) Then
        ImportCSVToTable DeletedFilePath, "tblCSVAccompanyingData"
    End If
    
    Create_tblCSVDataTemp SourceConnector, DeletedFilePath
    
    DoCmd.SetWarnings False
    CreateCSVDataToDestinationDatabase DestinationDatabase, SourceConnector
    ''DoCmd.CopyObject DestinationDatabase, "tblCSVData", acTable, "tblCSVData"
    DoCmd.SetWarnings True
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblCSVData"
          .fields = GenerateSelectFields(CSVImporterID)
          sqlStr = .sql
    End With
    
    GenerateSQLJoins sqlStr, CSVImporterID
    
    ''Run any update first (present)
    Dim updateSQL
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "UPDATE"
        .Source = "[" & TargetTable & "]"
        .SetStatement = GenerateSetStatements(CSVImporterID)
        .joins.Add GenerateJoinObj(sqlStr, TargetConnector, "temp")
        updateSQL = .sql
    End With
    
    RunSQLOnBackend DestinationDatabase, updateSQL
    
    ''Run any insert (not present)
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = sqlStr
          .fields = "temp.*"
          .AddFilter "[" & TargetTable & "]!" & TargetConnector & " IS NULL"
          .joins.Add GenerateJoinObj("[" & TargetTable & "]", TargetConnector, , , "LEFT")
          .SourceAlias = "temp"
          sqlStr = .sql
    End With
    
    Dim insertSQL
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "[" & TargetTable & "]"
          .fields = GenerateInsertFields(CSVImporterID)
          .insertSQL = sqlStr
          .InsertFilterField = GenerateInsertFields(CSVImporterID)
          insertSQL = .sql
    End With
    
    RunSQLOnBackend DestinationDatabase, insertSQL
    
    If Not batchMode Then MsgBox TargetTable & " successfully updated."
    
End Function

Private Function GenerateSQLJoins(ByRef sqlStr, CSVImporterID)
    
    Dim additionalFields As New clsArray
    Dim sqlObj As clsSQL, joinObj As clsJoin, rowsAffected
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = sqlStr
          .SourceAlias = "temp"
    End With

    Dim rs As Recordset: Set rs = ReturnRecordset("Select * FROM qryCSVImporterFields where CSVImporterID = " & CSVImporterID & " AND NOT LookupTable IS NULL")
    Do Until rs.EOF
        Dim CSVField: CSVField = rs.fields("CSVField"): If ExitIfTrue(isFalse(CSVField), "CSVField is empty..") Then Exit Function
        Dim TargetTable: TargetTable = rs.fields("TargetTable"): If ExitIfTrue(isFalse(TargetTable), "TargetTable is empty..") Then Exit Function
        Dim TargetField: TargetField = rs.fields("TargetField"): If ExitIfTrue(isFalse(TargetField), "Target Field is empty..") Then Exit Function
        Dim LookupTable: LookupTable = rs.fields("LookupTable"): If ExitIfTrue(isFalse(LookupTable), "LookupTable is empty..") Then Exit Function
        Dim LookupField: LookupField = rs.fields("LookupField"): If ExitIfTrue(isFalse(LookupField), "LookupField is empty..") Then Exit Function
        Dim FieldToReturn: FieldToReturn = rs.fields("FieldToReturn"): If ExitIfTrue(isFalse(FieldToReturn), "FieldToReturn is empty..") Then Exit Function
        additionalFields.Add "[" & LookupTable & "]!" & FieldToReturn & " AS " & TargetField
        sqlObj.joins.Add GenerateJoinObj(LookupTable, "temp" & TargetField, , LookupField, "LEFT")
        rs.MoveNext
    Loop
    
    If additionalFields.count > 0 Then
        sqlObj.fields = "temp.*," & additionalFields.JoinArr(",")
    Else
        sqlObj.fields = "temp.*"
    End If
    
    sqlStr = sqlObj.sql
    
End Function

Public Function GetFieldType(TableName, fieldName)

    GetFieldType = ELookup("qryTargetFieldLists", "TableName = " & Esc(TableName) & " AND FieldName = " & Esc(fieldName), "FieldTypeEnum")
    
End Function

Private Function GenerateSetStatements(CSVImporterID) As String
    
    Dim fields As New clsArray
    Dim rs As Recordset: Set rs = ReturnRecordset("Select * FROM qryCSVImporterFields where CSVImporterID = " & CSVImporterID & " AND NOT TargetField IS NULL")
    Dim FieldType
    Do Until rs.EOF
        Dim TargetTable: TargetTable = rs.fields("TargetTable"): If ExitIfTrue(isFalse(TargetTable), "TargetTable is empty..") Then Exit Function
        Dim TargetField: TargetField = rs.fields("TargetField"): If ExitIfTrue(isFalse(TargetField), "Target Field is empty..") Then Exit Function
        fields.Add "[" & TargetTable & "]!" & TargetField & " = temp!" & TargetField
        rs.MoveNext
    Loop
    
    If fields.count = 0 Then Exit Function
    
    GenerateSetStatements = fields.JoinArr(",")
    
End Function


Private Function GenerateInsertFields(CSVImporterID) As String
    
    Dim fields As New clsArray
    Dim rs As Recordset: Set rs = ReturnRecordset("Select * FROM qryCSVImporterFields where CSVImporterID = " & CSVImporterID & " AND NOT TargetField IS NULL")
    Dim FieldType
    Do Until rs.EOF
        
        Dim TargetField: TargetField = rs.fields("TargetField"): If ExitIfTrue(isFalse(TargetField), "Target Field is empty..") Then Exit Function
        fields.Add "[" & TargetField & "]"
        rs.MoveNext
    Loop
    
    If fields.count = 0 Then Exit Function
    
    GenerateInsertFields = fields.JoinArr(",")
    
End Function

Private Function GenerateSelectFields(CSVImporterID) As String
    
    Dim fields As New clsArray
    Dim rs As Recordset: Set rs = ReturnRecordset("Select * FROM qryCSVImporterFields where CSVImporterID = " & CSVImporterID & " AND NOT TargetField IS NULL")
    Dim FieldType
    Do Until rs.EOF
    
        Dim CSVField: CSVField = rs.fields("CSVField"): If ExitIfTrue(isFalse(CSVField), "CSVField is empty..") Then Exit Function
        Dim TargetField: TargetField = rs.fields("TargetField"): If ExitIfTrue(isFalse(TargetField), "Target Field is empty..") Then Exit Function
        Dim TargetTable: TargetTable = rs.fields("TargetTable"): If ExitIfTrue(isFalse(TargetTable), "TargetTable is empty..") Then Exit Function
        Dim FieldToReturn: FieldToReturn = rs.fields("FieldToReturn")
        Dim LookupTable: LookupTable = rs.fields("LookupTable")
        Dim LookupField: LookupField = rs.fields("LookupField")
        
        CSVField = "[" & CSVField & "]"
        FieldType = GetFieldType(TargetTable, TargetField)
        
        Select Case FieldType
            Case "dbDate":
                CSVField = "DateValue(tblCSVData!" & CSVField & ")"
        End Select
        
        If TargetField = "RecordImportID" Then
            CSVField = "CStr(tblCSVData!" & CSVField & ")"
        End If
        
        If Not isFalse(LookupTable) And LookupField = "RecordImportID" Then
            CSVField = "iif(isNull(tblCSVData!" & CSVField & "),null,CStr(tblCSVData!" & CSVField & ")) AS temp" & TargetField
            fields.Add CSVField
            GoTo NextRecord
        End If
        
        TargetField = "[" & TargetField & "]"
        If CSVField <> TargetField Then
            fields.Add CSVField & " AS " & TargetField
        Else
            fields.Add CSVField
        End If
        
        
NextRecord:
        rs.MoveNext
    Loop
    
    If fields.count = 0 Then Exit Function
    
    GenerateSelectFields = fields.JoinArr(",")
    
End Function


