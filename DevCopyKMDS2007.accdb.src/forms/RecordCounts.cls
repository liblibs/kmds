Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdPurgeItems_Click()
    Dim strSQL As String
    If gbAllowDeleteDocument Then

        If MsgBox("This will purge items with 0 onhands with no activity for the last 500 days." & vbCrLf & "Are you sure?", vbCritical + vbYesNo) = vbYes Then
        'First show the candidates by setting avgturns to 999
            SetOptions False
            SQLPassThru ("sp_PurgeItems_Step_1_Set_AvgTurns_0")
            SQLPassThru ("sp_PurgeItems_Step_2_Delete_Involine_0_Qty")
            'setting the purge flag to 999 (AvgTurns)
            SQLPassThru ("sp_PurgeItems_Step_3_Set_AvgTurns_999")
            DoCmd.OpenQuery "qryPurgeItemsExceptKits"
            'Make sure none of the system items get purged by resetting the purge flag (AvgTurns)
            strSQL = "UPDATE Item, Sysfile SET Item.AvgTurns = 0 WHERE (((Item.ID)=[SYSFILE].[CashDropItemID]))"
            DoCmd.RunSQL strSQL
            strSQL = "UPDATE Item, Sysfile SET Item.AvgTurns = 0 WHERE (((Item.ID)=[SYSFILE].[CashLoanItemID]))"
            DoCmd.RunSQL strSQL
            strSQL = "UPDATE Item, Sysfile SET Item.AvgTurns = 0 WHERE (((Item.ID)=[SYSFILE].[ChargeDropItemID]))"
            DoCmd.RunSQL strSQL
            strSQL = "UPDATE Item, Sysfile SET Item.AvgTurns = 0 WHERE (((Item.ID)=[SYSFILE].[CheckDropItemID]))"
            DoCmd.RunSQL strSQL
            strSQL = "UPDATE Item, Sysfile SET Item.AvgTurns = 0 WHERE (((Item.ID)=[SYSFILE].[ClubDiscntItemID]))"
            DoCmd.RunSQL strSQL
            strSQL = "UPDATE Item, Sysfile SET Item.AvgTurns = 0 WHERE (((Item.ID)=[SYSFILE].[CouponItemID]))"
            DoCmd.RunSQL strSQL
            strSQL = "UPDATE Item, Sysfile SET Item.AvgTurns = 0 WHERE (((Item.ID)=[SYSFILE].[DiscountItemID]))"
            DoCmd.RunSQL strSQL
            strSQL = "UPDATE Item, Sysfile SET Item.AvgTurns = 0 WHERE (((Item.ID)=[SYSFILE].[LotteryPaidOutItemID]))"
            DoCmd.RunSQL strSQL
            strSQL = "UPDATE Item, Sysfile SET Item.AvgTurns = 0 WHERE (((Item.ID)=[SYSFILE].[NonClubDiscntItemID]))"
            DoCmd.RunSQL strSQL
            strSQL = "UPDATE Item, Sysfile SET Item.AvgTurns = 0 WHERE (((Item.ID)=[SYSFILE].[NonClubInvItemID]))"
            DoCmd.RunSQL strSQL
            strSQL = "UPDATE Item, Sysfile SET Item.AvgTurns = 0 WHERE (((Item.ID)=[SYSFILE].[DepositItemID]))"
            DoCmd.RunSQL strSQL
            SQLPassThru ("sp_PurgeItems_Step_4_Delete_AvgTurns_999")
            SetOptions True
            MsgBox "Task Completed"
            'do NOT USE this because the the system items may get deleted
            'SQLPassThru ("sp_PurgeItems")
        End If
    Else
        MsgBox "This user is Not allowed to delete documents", vbCritical, glsCashierName
    End If


End Sub

Private Sub Command4_Click()
    Dim txtb As Control
    For Each txtb In Me.Form.Controls
        If txtb.ControlType = acTextBox Then
            If Not gcfHandleErrors Then Debug.Print txtb.Name
            GetRecordCount txtb.Name
            Me.Repaint
        End If
    Next txtb
End Sub

Public Sub GetRecordCount(TableName As String)

If gcfHandleErrors Then On Error GoTo Err_Command4_Click
On Error GoTo Err_Command4_Click
Dim rstemp As Recordset

    Set rstemp = CurrentDb().OpenRecordset(TableName) ', dbOpenForwardOnly, dbSeeChanges)
    'On Error Resume Next
    rstemp.MoveLast
    'If gcfHandleErrors Then On Error GoTo Err_Command4_Click

    Me(TableName) = rstemp.RecordCount

    rstemp.Close
    Set rstemp = Nothing


Exit_Command4_Click:
    Exit Sub

Err_Command4_Click:
    'MsgBox Err.Description
    Resume Exit_Command4_Click

End Sub
