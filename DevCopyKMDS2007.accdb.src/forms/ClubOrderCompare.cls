Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub List0_BeforeUpdate(Cancel As Integer)

End Sub

Private Sub cmdClose_Click()
    DoCmd.Close
End Sub

Private Sub Form_Load()
    Me.InsideHeight = boxScreen.Height + 100
    Me.InsideWidth = boxScreen.Width + 500
End Sub

Private Sub lstInvoices_2_AfterUpdate()
    txtPurchaseOrderID_2 = lstInvoices_2.Column(0)
    ComparePurchaseOrders

End Sub


Private Sub lstInvoices_AfterUpdate()
    txtPurchaseOrderID_1 = lstInvoices.Column(0)
    With lstInvoices_2
        .Requery
        .Visible = True
        .SetFocus
    End With
End Sub


Private Sub ComparePurchaseOrders()
    Dim lngVendorID As Long
    Dim strVendor As String
    If gcfHandleErrors Then On Error GoTo ErrLog
    With lstPurchaseOrders
        'lngPurchaseOrderID = .ItemData(.ListIndex + 1)
        'txtOrderID = .ItemData(.ListIndex + 1)
        'strVendor = lstVendor ' .List(.ListIndex, 1)
        If MsgBox("Compare Licensee order " & txtPurchaseOrderID_1 & " to Purchase order  " & txtPurchaseOrderID_2, vbYesNo) = vbNo Then
        Else
'            If CreatePurchaseOrderFromScanBatch(lngVendorID) > 0 Then
                SetOptions False
                On Error Resume Next
                DoCmd.DeleteObject acTable, "CompareScanItemToPurchaseOrder"
                DoCmd.DeleteObject acTable, "CompareScanItemToPurchaseOrderScanCounts"
                If gcfHandleErrors Then On Error GoTo ErrLog

                DoCmd.OpenQuery ("qryCompareLicenseeOrders_01_MakeTable")
                DoCmd.OpenQuery ("qryCompareLicenseeOrders_02_PO2")
                DoCmd.OpenQuery ("qryCompareScanItemToPurchaseOrder_03_UpdateQuantityScanned")
                DoCmd.OpenQuery ("qryCompareScanItemToPurchaseOrder_04_Append")
                DoCmd.OpenQuery ("qryCompareScanItemToPurchaseOrder_04a_Repacks")
'                ExportScanBatchLineCounts
'                DoCmd.RunSQL "Delete * from ScanBatchLineCounts Where ScanBatchHeaderID = 3" '& txtScanBatchHeaderId
'                Me.Requery
               ' SetupScreens
                SetOptions True
                If gcfHandleErrors Then On Error GoTo ErrLog
                Me.Visible = False
'                MsgBox "Purchase order created; the purchase order screen will be displayed."
                DoCmd.OpenForm "CompareScanItemToPurchaseOrder", acNormal, "", "", , acNormal
'
'                'DoCmd.Close acForm, "ScanBatch", acSavePrompt
'                'lstLocation.Visible = True
'                'lstLocation.SetFocus
'                Exit Sub
'
'            End If
            'txtCode.SetFocus
           ' .Visible = False
           ' cmdfrmClose.Visible = True


        End If
    End With

ExitSub:
    'txtCode.SetFocus
    Exit Sub
ErrLog:
    MsgBox "Error: " & Err.Number & " " & Err.Description & "  lstVendor Exit", vbRed, True
    Resume ExitSub
End Sub
