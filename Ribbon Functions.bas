Attribute VB_Name = "Ribbon Functions"
Option Compare Database
Option Explicit

Public Sub OpenFormFromRibbon(ctl As IRibbonControl)
    DoCmd.OpenForm ctl.id
End Sub

Public Sub changeGlobal(ctl As IRibbonControl)

    'UPDATE STATEMENT
    Dim sqlObj As New clsSQL, fltrObj As New clsArray
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "UPDATE"
        .Source = "tblGlobalSettings"
        .SetStatement = "GlobalSettingValue = " & EscapeString(CurrentProject.path, "tblGlobalSettings", "GlobalSettingValue")
        .AddFilter "GlobalSetting = ""systemProductImages_FilePath"""
        .Run
    End With
    
    fltrObj.arr = "Application_ImportCSV_FilePath,rptShelfLocationLabels,rptPackSheets_FilePath,rptPickSheets_FilePath," & _
        "rptIntermediateLabels_FilePath,rptPrintH_FilePath"
    
    Dim arrItem
    For Each arrItem In fltrObj.arr
    
        Set sqlObj = New clsSQL
        With sqlObj
            .SQLType = "UPDATE"
            .Source = "tblGlobalSettings"
            .SetStatement = "GlobalSettingValue = " & EscapeString("C:\Users\user\Desktop\Printables\")
            .AddFilter "GlobalSetting = " & EscapeString(arrItem)
            .Run
        End With
    
    Next arrItem
    
End Sub

Public Sub TurnOnFilePrompt(ctl As IRibbonControl)
    
    'NoHasWriteToFilePrompt = False
    
End Sub

Public Sub CustomPrintPreviewPrint(control As IRibbonControl)
    ' Add your custom code here
    ' This code will be executed when the Print button is clicked
    MsgBox "I'm clicked"
End Sub
