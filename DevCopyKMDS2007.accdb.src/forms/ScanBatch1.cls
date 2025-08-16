Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdAppendToAudit_Click()
    'Dim rsAudit As New ADODB.Recordset
    Dim rsAudit As Object
    Set rsAudit = CreateObject("ADODB.Recordset")

    If gcfHandleErrors Then On Error GoTo cmdMove_Click_Err
    SetOptions False
    ArchiveReport "65 Scan Batch Variance"

    rsAudit.Open "Audit", CurrentProject.Connection, adOpenDynamic, adLockOptimistic


    If rsAudit.EOF And rsAudit.BOF Then
        DoCmd.OpenQuery "qryScanBatchAppendToAudit", acViewNormal, acEdit
        DoCmd.OpenQuery "qryScanBatchAppendToAuditRecalcVariance", acViewNormal, acEdit
        DoCmd.RunSQL "Delete * from ScanBatchLineCounts Where ScanBatchHeaderID = " & glngScanBatchHeaderID
        DoCmd.Close acForm, "ScanBatch"
        DoCmd.Close , ""

    Else
        MsgBox "The D-7 Audit Inventory screen already has a batch pending; this batch must be posted or deleted before this is allowed!", vbCritical
    End If

cmdMove_Click_Exit:
    SetOptions True
    Exit Sub

cmdMove_Click_Err:
    MsgBox Error$
    Resume cmdMove_Click_Exit

End Sub

Private Sub cmdBatchScanReport_Click()
    If gcfHandleErrors Then On Error GoTo cmdBatchScan_Click_Err

    DoCmd.OpenReport "Batch Scan", acViewPreview, "", "", acNormal
    DoCmd.RunCommand acCmdPrint

cmdBatchScan_Click_Exit:
    Exit Sub

cmdBatchScan_Click_Err:
    MsgBox Error$
    Resume cmdBatchScan_Click_Exit
End Sub

Private Sub cmdEraseBatch_Click()
    If gcfHandleErrors Then On Error GoTo ErrLog
    If MsgBox("Erase the totals from this batch?", vbYesNo, "Delete this batch!") = vbYes Then
        If MsgBox("Are you sure?", vbYesNo, "Delete this batch!") = vbYes Then
            SetOptions False
            DoCmd.RunSQL "Delete * from ScanBatchLineCounts Where ScanBatchHeaderID = " & glngScanBatchHeaderID
            'mrsScanBatchLineCounts.Requery
            'mrsScanBatchLineCounts.Open "ScanBatchLineCounts", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
            DoCmd.Close acForm, "ScanBatch"
            DoCmd.Close , ""
            Exit Sub
        End If
    End If
    ScanBatchTotalsSub.Requery
    'txtCode.SetFocus


ExitSub:
    SetOptions True
    Exit Sub
ErrLog:
    MsgBox "Error: " & Err.Number & " " & Err.Description & "  cmdEraseBatch_Click", vbRed
    Resume ExitSub

End Sub
'------------------------------------------------------------
' Command2_Click
'
'------------------------------------------------------------
Private Sub Command2_Click()
If gcfHandleErrors Then On Error GoTo Command2_Click_Err

    DoCmd.Close , ""


Command2_Click_Exit:
    Exit Sub

Command2_Click_Err:
    MsgBox Error$
    Resume Command2_Click_Exit

End Sub


'------------------------------------------------------------
' cmdPrint65ScanBatch_Click
'
'------------------------------------------------------------
Private Sub cmdPrint65ScanBatch_Click()
If gcfHandleErrors Then On Error GoTo cmdPrint65ScanBatch_Click_Err

    DoCmd.OpenReport "65 Scan Batch Variance", acViewPreview, "", "", acNormal


cmdPrint65ScanBatch_Click_Exit:
    Exit Sub

cmdPrint65ScanBatch_Click_Err:
    MsgBox Error$
    Resume cmdPrint65ScanBatch_Click_Exit

End Sub
