Attribute VB_Name = "qryCustomerOrderReportNCs Mod"
Option Compare Database
Option Explicit

Public Function qryCustomerOrderReportNCsCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
            FormatFormAsReport frm, 8
            
            If IsObjectAReport(frm) Then
                frm("lblOrderAnalysisText").Caption = ""
                frm("OrderAnalysisText").TextFormat = 1
                
                Dim ctl As control
                For Each ctl In frm.Controls
                    If ctl.ControlType = acTextBox Then
                        ctl.CanGrow = True
                    End If
                Next ctl
                
                For Each ctl In frm.Controls
                    ctl.InSelection = True
                Next ctl
                
                DoCmd.RunCommand acCmdTabularLayout
                DoCmd.RunCommand acCmdControlPaddingNone
                
            End If
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function

Public Function GetOrderAnalysisText(ClaimRootCause, ClaimResolution, CorrectiveMeasures)

    Dim claimArr As New clsArray
    
    If Not isFalse(ClaimRootCause) Then
        claimArr.Add "<b>Root Cause: </b><br/>" & ClaimRootCause
    End If
    
    If Not isFalse(ClaimResolution) Then
        claimArr.Add "<b>Korrektur: </b><br/>" & ClaimResolution
    End If
    
    If Not isFalse(CorrectiveMeasures) Then
        claimArr.Add "<b>Massnahmen: </b><br/>" & CorrectiveMeasures
    End If

    GetOrderAnalysisText = claimArr.JoinArr("<br/>")
End Function
