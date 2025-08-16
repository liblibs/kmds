Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdClose_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdClose_Click


    DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
End Sub

Private Sub Frame3_AfterUpdate()
'With Forms("CompareScanItemToPurchaseOrderSub")
    With [CompareScanItemToPurchaseOrderSubForm].Form
        Select Case Frame3.Value
        Case 1 'all
            .FilterOn = False
            .Requery
            txtTotals.Caption = "All Items"

        Case 2 'Full cases
            .Filter = "(((CompareScanItemToPurchaseOrder.FullCase)=True))"
            .FilterOn = True
            .Requery
            txtTotals.Caption = "Full Cases"

        Case 3 'Repacks
            .Filter = "(((CompareScanItemToPurchaseOrder.FullCase)=false))"
            .FilterOn = True
            .Requery
            txtTotals.Caption = "Repacks"

        End Select
    End With

End Sub
