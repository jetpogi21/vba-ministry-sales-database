Attribute VB_Name = "FixModel Mod"
Option Compare Database
Option Explicit

Public Function FixModelCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Sub FixModel()
    ''in hero skills convert SkillType a and p to Active and Passive respectively
    Dim valueArr As New clsArray: valueArr.arr = "a,p"
    Dim updatedValueArr As New clsArray: updatedValueArr.arr = "Active,Passive"
    
    RunDatabaseUpdate valueArr, updatedValueArr, "tblHeroSkills", "SkillType"
    
    ''in cards CardType c,w,p,t
    valueArr.arr = "c,w,p,t"
    updatedValueArr.arr = "Character,Weapon,Power,Tactic"
    RunDatabaseUpdate valueArr, updatedValueArr, "tblCards", "CardType"
    
    ''in cards BattleStyle a,g,s
    valueArr.arr = "a,g,s"
    updatedValueArr.arr = "Attack,Guardian,Support"
    RunDatabaseUpdate valueArr, updatedValueArr, "tblCards", "BattleStyle"
    
End Sub

Private Function RunDatabaseUpdate(valueArr As clsArray, updatedValueArr As clsArray, tblName, fldName)
    
    Dim i As Integer
    Dim Value, updateValue
    For i = 0 To valueArr.count - 1
        Debug.Print valueArr.count
        Value = valueArr.items(i)
        updateValue = updatedValueArr.items(i)
        ''Debug.Print "UPDATE " & tblName & " SET " & fldName & " = " & EscapeString(updateValue) & " WHERE " & fldName & " = " & EscapeString(value)
        RunSQL "UPDATE " & tblName & " SET " & fldName & " = " & EscapeString(updateValue) & " WHERE " & fldName & " = " & EscapeString(Value)
    Next i
    
End Function
