Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'simple down & dirty inventory input screen designed to use a bluetooth scanner and allow
'user to run program 'from a distance'  (Keep keyboard entries to a bare minimum)
'input card with quantities will be provided so the count can be entered with a scanner

'mrsScanBatch is the lookup query based on vwUPCItemsWithCurrentPricing
'mrsScanBatchLineCounts is a table that holds the counts

Dim mrsScanBatch As Object
Dim mrsScanBatchSub As Object

'Dim rstemp As New ADODB.Recordset
Dim mrsScanBatchLineCounts As Object

Dim mrsSBLC As Recordset

Dim mrsItemTotals As Recordset
Dim mbShowMoreOptions As Boolean
Dim mbBackRoom As Boolean
Dim nMaxListItems As Integer



Private Sub cmd65VarReport_Click()
    If gcfHandleErrors Then On Error GoTo cmd65VarReport_Click_Err
    If MsgBox("Do you want to print the report? " & vbCr & vbCr & "Yes = Print report" & vbCr & " No = View report", vbYesNo, "Print Report?") = vbYes Then
        DoCmd.OpenReport "65 Scan Batch Variance", acViewNormal, "", "", acNormal
    Else
        DoCmd.OpenReport "65 Scan Batch Variance", acViewReport, "", "", acNormal
    End If

cmd65VarReport_Click_Exit:
    Exit Sub

cmd65VarReport_Click_Err:
    MsgBox Error$
    Resume cmd65VarReport_Click_Exit

End Sub

Private Sub cmdAppendToAudit_Click()
    'Dim rstemp As New ADODB.Recordset
    Dim rstemp As Object
    Set rstemp = CreateObject("ADODB.Recordset")

    If gcfHandleErrors Then On Error GoTo cmdMove_Click_Err
    SetOptions False
    ArchiveReport "65 Scan Batch Variance"

    rstemp.Open "Audit", CurrentProject.Connection, adOpenDynamic, adLockOptimistic


    If rstemp.EOF And rstemp.BOF Then
        DoCmd.OpenQuery "qryScanBatchAppendToAudit", acViewNormal, acEdit
        DoCmd.OpenQuery "qryScanBatchAppendToAuditRecalcVariance", acViewNormal, acEdit
        DoCmd.OpenQuery "qryScanBatchAppendToInvoline", acViewNormal, acEdit
        ExportScanBatchLineCounts
        DoCmd.RunSQL "Delete * from ScanBatchLineCounts Where ScanBatchHeaderID = " & glngScanBatchHeaderID '& txtScanBatchHeaderId
        DoCmd.Close acForm, "ScanBatch"
    Else
        StatusLine "The D-7 Audit Inventory screen already has a batch pending; this batch must be posted or deleted before this is allowed!", vbRed, True
        MsgBox "The D-7 Audit Inventory screen already has a batch pending; this batch must be posted or deleted before this is allowed!", vbCritical
    End If

cmdMove_Click_Exit:
    SetOptions True
    Exit Sub

cmdMove_Click_Err:
    MsgBox Error$
    Resume cmdMove_Click_Exit

End Sub

Private Sub cmdBatchReport_Click()
    If gcfHandleErrors Then On Error GoTo cmdBatchReport_Click_Err


    If MsgBox("Do you want to print the report? " & vbCr & vbCr & "Yes = Print report" & vbCr & " No = View report", vbYesNo, "Print Report?") = vbYes Then
        DoCmd.OpenReport "66 Scan Batch Detail By Item", acViewNormal, "", "", acNormal
    Else
        DoCmd.OpenReport "66 Scan Batch Detail By Item", acViewReport, "", "", acNormal
    End If

cmdBatchReport_Click_Exit:
    Exit Sub

cmdBatchReport_Click_Err:
    MsgBox Error$
    Resume cmdBatchReport_Click_Exit

End Sub

Private Sub cmdBatchScanReport_Click()
    If gcfHandleErrors Then On Error GoTo cmdBatchScan_Click_Err

    DoCmd.OpenReport "Batch Scan", acViewNormal, "", "", acNormal

cmdBatchScan_Click_Exit:
    Exit Sub

cmdBatchScan_Click_Err:
    MsgBox Error$
    Resume cmdBatchScan_Click_Exit

End Sub

Private Sub cmdChangeLocations_Click()
    lstLocation.Visible = True
    lstLocation.SetFocus
End Sub

Private Sub cmdComparePO_Click()
    Dim strSQL As String

    If gcfHandleErrors Then On Error GoTo ErrLog
    With lstPurchaseOrders
        .Visible = True
        .Height = txtCode.Height * 10
        .BoundColumn = 1
        .ColumnHeads = True
        .SetFocus
    End With
    StatusLine "Compare these scans to a Purchase order; Select from the list and then hit enter", vbYellow, True
ExitSub:
    SetOptions True
    Exit Sub
ErrLog:
    MsgBox "Error: " & Err.Number & " " & Err.Description & " Create PO Click", vbRed
    Resume ExitSub

End Sub

Private Sub cmdCreateClaim_Click()
    Dim strSQL As String

    If gcfHandleErrors Then On Error GoTo ErrLog
    With lstClaim
        .Visible = True
        .Height = txtCode.Height * 10
        .BoundColumn = 1
        .ColumnHeads = True
        .SetFocus
    End With
    StatusLine "Create a Store Claim; Select from the list and then hit enter", vbYellow, True
ExitSub:
    SetOptions True
    Exit Sub
ErrLog:
    MsgBox "Error: " & Err.Number & " " & Err.Description & " Create PO Click", vbRed
    Resume ExitSub
End Sub

Private Sub cmdCreateClubOrder_Click()
    Dim strSQL As String

    If gcfHandleErrors Then On Error GoTo ErrLog
    If lstClubs.Visible Then
    Else
        strSQL = "SELECT Name, ID From Club "
        strSQL = strSQL & "WHERE (((ID) > 0)) "
        strSQL = strSQL & "ORDER BY Name"

        With lstClubs
            .Visible = True
            .Value = Null
            .ColumnCount = 2
            .ColumnHeads = True
            .RowSourceType = "Table/Query"
            .RowSource = strSQL
            .Height = txtCode.Height * 10
            .BoundColumn = 2
            .Requery
        End With
    End If
    StatusLine "Create a Club order; Select a Club from the list and then hit enter", vbYellow, True
    lstClubs.Visible = True
    lstClubs.SetFocus
ExitSub:
    SetOptions True
    Exit Sub
ErrLog:
    MsgBox "Error: " & Err.Number & " " & Err.Description & " Create PO Click", vbRed
    Resume ExitSub

End Sub

Private Sub cmdCreatePO_Click()
    Dim strSQL As String

    If gcfHandleErrors Then On Error GoTo ErrLog
    If lstVendor.Visible Then
    Else
        strSQL = "SELECT Name, ID From Vendor "
        strSQL = strSQL & "WHERE (((Vendor.ID) > 0)) "
        strSQL = strSQL & "ORDER BY Vendor.Name"

        With lstVendor
            .Visible = True
            .Value = Null
            .ColumnCount = 2
            .ColumnHeads = True
            .RowSourceType = "Table/Query"
            .RowSource = strSQL
            .Height = txtCode.Height * 10
            .BoundColumn = 2
            .Requery
            .SetFocus
        End With
    End If
    StatusLine "Create a purchase order; Select a vendor from the list and then hit enter", vbYellow, True
    'lstVendor.SetFocus
ExitSub:
    SetOptions True
    Exit Sub
ErrLog:
    MsgBox "Error: " & Err.Number & " " & Err.Description & " Create PO Click", vbRed
    Resume ExitSub

End Sub

Private Sub cmdCreateTransfer_Click()
    Dim strSQL As String

    If gcfHandleErrors Then On Error GoTo ErrLog
    With lstAgents
        .Visible = True
        .Height = txtCode.Height * 10
        .BoundColumn = 1
        .ColumnHeads = True
        .SetFocus
    End With

    StatusLine "Create a Store Transfer; Select an agent from the list and then hit enter", vbYellow, True
ExitSub:
    SetOptions True
    Exit Sub
ErrLog:
    MsgBox "Error: " & Err.Number & " " & Err.Description & " Create PO Click", vbRed
    Resume ExitSub

End Sub

Private Sub cmdEraseBatch_Click()
    If gcfHandleErrors Then On Error GoTo ErrLog
    If MsgBox("Erase the totals from this batch?", vbYesNo, "Delete this batch!") = vbYes Then
        If MsgBox("Are you sure?", vbYesNo, "Delete this batch!") = vbYes Then
            SetOptions False
            ExportScanBatchLineCounts
            DoCmd.RunSQL "Delete * from ScanBatchLineCounts Where ScanBatchHeaderID = " & glngScanBatchHeaderID '& txtScanBatchHeaderId
            SetOptions True
            DoCmd.Close acForm, "ScanBatch"
        End If
    End If

ExitSub:
    Exit Sub
ErrLog:
    MsgBox "Error: " & Err.Number & " " & Err.Description & "  cmdEraseBatch_Click", vbCritical
    StatusLine "Error: " & Err.Number & " " & Err.Description & "  Erase Batch Click", vbRed, True
    Resume ExitSub
End Sub

Private Sub cmdfrmClose_Click()
    DoCmd.Close acForm
End Sub




Private Sub cmdRestoreBatch_Click()
    If gcfHandleErrors Then On Error GoTo ErrLog
    On Error Resume Next
    DoCmd.OpenForm "ScanBatchRestore", acNormal, "", "", , acNormal
    DoCmd.Close acForm, "ScanBatch", acSavePrompt
ExitSub:
    Exit Sub
ErrLog:
    MsgBox "Error: " & Err.Number & " " & Err.Description & "  cmdEraseBatch_Click", vbCritical
    StatusLine "Error: " & Err.Number & " " & Err.Description & "  Erase Batch Click", vbRed, True
    Resume ExitSub
End Sub

Private Sub Form_Load()
    'Dim rsScanBatchHeader As New ADODB.Recordset
    'Dim rstemp As New ADODB.Recordset
    Dim rstemp As Object
    Set rstemp = CreateObject("ADODB.Recordset")

    If gcfHandleErrors Then On Error GoTo ErrLog

    nMaxListItems = 50

    'makesure the list box showing previous scans is emptied
    SetupScreens
    StatusLine "Opening Tables", vbYellow, False


    With rstemp  ' have to load 3 records in ScanBatchHeader table so parent row exists
        .Open "ScanBatchHeader", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
        If .EOF And .BOF Then
            .AddNew
            !CashierID = 1
            .Update
            .AddNew
            !CashierID = 2
            .Update
            .AddNew
            !CashierID = 3
            .Update
        End If
        .Close
    End With
    Set rstemp = Nothing


    Set mrsScanBatch = CurrentDb().OpenRecordset("qryScanBatch")
    txtOnShelfDate = Now ' this is when the above recordset was created

    Set mrsScanBatchLineCounts = CreateObject("ADODB.Recordset")
    mrsScanBatchLineCounts.Open "ScanBatchLineCounts", CurrentProject.Connection, adOpenDynamic, adLockOptimistic

    Me.Caption = "ScanBatch with User : " & glnCashierID & " " & glsCashierName

    lstLocation.Visible = True
    lstLocation.SetFocus

ExitSub:
    mbShowMoreOptions = False
    Exit Sub
ErrLog:
    StatusLine "Error: " & Err.Number & " " & Err.Description & "  Form Load", vbRed, True
    Resume ExitSub
End Sub

Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)

End Sub

Private Sub lstAgents_Exit(Cancel As Integer)
    Dim lngAgentID As Long
    Dim lngTransferID As Long
    Dim strAgentName As String
    Dim bTransferIn As Boolean
    Dim strInOrOut As String
    If gcfHandleErrors Then On Error GoTo ErrLog
    With lstAgents
        lngAgentID = .ItemData(.ListIndex + 1)
        'strVendor = lstVendor ' .List(.ListIndex, 1)
        If MsgBox("Create a Store transfer for Agent # " & lngAgentID & " " & strAgentName & "?", vbYesNo) = vbNo Then
        Else
            If MsgBox("Is this a Transfer Out? (Leaving your store?)", vbYesNo) = vbYes Then
                bTransferIn = False
                strInOrOut = "O"
            Else
                bTransferIn = True
                strInOrOut = "I"
            End If
            lngTransferID = CreateTransfer(lngAgentID, strInOrOut)



            If CreateStoreTransferFromScanBatch(lngAgentID, lngTransferID, strInOrOut) > 0 Then
                SetOptions False
                ExportScanBatchLineCounts
                DoCmd.RunSQL "Delete * from ScanBatchLineCounts Where ScanBatchHeaderID = " & glngScanBatchHeaderID
                Me.Requery
                Me.ScanBatchSub.Requery
                SetupScreens
                SetOptions True
                If gcfHandleErrors Then On Error GoTo ErrLog
                Me.Visible = False
                MsgBox "Store Transfer created; the Transfer screen will be displayed."
                If bTransferIn Then
                    DoCmd.OpenForm "AgentTransferIn", acNormal, "", "", , acNormal
                Else
                    DoCmd.OpenForm "AgentTransferOut", acNormal, "", "", , acNormal
                End If

            End If
            cmdfrmClose.Visible = True


        End If
        txtCode.SetFocus
        .Visible = False
    End With

ExitSub:
    txtCode.SetFocus
    Exit Sub
ErrLog:
    StatusLine "Error: " & Err.Number & " " & Err.Description & "  lstAgents Exit", vbRed, True
    Resume ExitSub

End Sub

Private Sub lstClaim_Exit(Cancel As Integer)
    Dim lngAgentID As Long
    Dim lngTransferID As Long
    Dim strAgentName As String
    Dim bTransferIn As Boolean
    Dim strInOrOut As String
    If gcfHandleErrors Then On Error GoTo ErrLog
    With lstClaim
        lngAgentID = .ItemData(.ListIndex + 1)
        If MsgBox("Create a CLAIM for " & lngAgentID & " " & strAgentName & "?", vbYesNo) = vbNo Then
        Else
            bTransferIn = False
            strInOrOut = "C"
            lngTransferID = CreateTransfer(lngAgentID, strInOrOut)



            If CreateStoreTransferFromScanBatch(lngAgentID, lngTransferID, strInOrOut) > 0 Then
                SetOptions False
                ExportScanBatchLineCounts
                DoCmd.RunSQL "Delete * from ScanBatchLineCounts Where ScanBatchHeaderID = " & glngScanBatchHeaderID
                Me.Requery
                SetupScreens
                SetOptions True
                If gcfHandleErrors Then On Error GoTo ErrLog
                Me.Visible = False
                MsgBox "Claim created; the OLCC Claims screen will be displayed."
                DoCmd.OpenForm "AgentTransferOutClaims", acNormal, "", "", , acNormal

'                'DoCmd.Close acForm, "ScanBatch", acSavePrompt
                'lstLocation.Visible = True
                'lstLocation.SetFocus
                'Exit Sub
'
            End If
            cmdfrmClose.Visible = True


        End If
        txtCode.SetFocus
        .Visible = False
    End With

ExitSub:
    txtCode.SetFocus
    Exit Sub
ErrLog:
    StatusLine "Error: " & Err.Number & " " & Err.Description & "  lstClaim Exit", vbRed, True
    Resume ExitSub
End Sub

Private Sub lstClubs_DblClick(Cancel As Integer)
    MsgBox "Please Hit enter to select the Club!"

End Sub

Private Sub lstClubs_Exit(Cancel As Integer)
    Dim lngClubID As Long
    Dim strClub As String
    If gcfHandleErrors Then On Error GoTo ErrLog
    On Error GoTo ErrLog
    With lstClubs
        lngClubID = .ItemData(.ListIndex + 1)
        'strClub = lstVendor ' .List(.ListIndex, 1)
        If MsgBox("Create an order for Club # " & lngClubID, vbYesNo) = vbNo Then
        Else
            If CreateClubOrderFromScanBatch(lngClubID) > 0 Then
                SetOptions False
                ExportScanBatchLineCounts
                DoCmd.RunSQL "Delete * from ScanBatchLineCounts Where ScanBatchHeaderID = " & glngScanBatchHeaderID
                Me.Requery
                SetupScreens
                SetOptions True
                If gcfHandleErrors Then On Error GoTo ErrLog
                Me.Visible = False
                MsgBox "Club order created; the Club order screen will be displayed."
                DoCmd.OpenForm "ClubOrderCreate", acNormal, "", "", , acNormal

                'DoCmd.Close acForm, "ScanBatch", acSavePrompt
                'lstLocation.Visible = True
                'lstLocation.SetFocus
                Exit Sub

            End If
            txtCode.SetFocus
            lstClubs.Visible = False
            cmdfrmClose.Visible = True


        End If
    End With

ExitSub:
    txtCode.SetFocus
    Exit Sub
ErrLog:
    StatusLine "Error: " & Err.Number & " " & Err.Description & "  lstClubs Exit", vbRed, True
    Resume ExitSub

End Sub



Private Sub lstLocation_DblClick(Cancel As Integer)
    MsgBox "Please Hit enter to select the location!"
End Sub

Private Sub lstLocation_Enter()
    StatusLine "Please select the location and Press Enter", vbYellow, True
End Sub

Private Sub lstLocation_Exit(Cancel As Integer)
    If gcfHandleErrors Then On Error GoTo ErrLog
    StatusLine "Location: " & lstLocation, vbRed, True
    Me.Caption = "Scan Batch at " & lstLocation
    txtLocation = lstLocation
    If lstLocation = "2. Backroom" Then
        mbBackRoom = True
    Else
        mbBackRoom = False
    End If
    txtCode.SetFocus
ExitSub:
    Exit Sub
ErrLog:
    StatusLine "Error: " & Err.Number & " " & Err.Description & "  lstLocation Exit", vbRed, True
    Resume ExitSub
End Sub



Private Sub lstPurchaseOrders_Exit(Cancel As Integer)
    Dim lngVendorID As Long
    Dim strVendor As String
    If gcfHandleErrors Then On Error GoTo ErrLog
    With lstPurchaseOrders
        glngPurchaseOrderID = .ItemData(.ListIndex + 1)
        txtOrderID = .ItemData(.ListIndex + 1)
        'strVendor = lstVendor ' .List(.ListIndex, 1)
        If MsgBox("Compare Scans to Purchase order # " & glngPurchaseOrderID, vbYesNo) = vbNo Then
        Else
'            If CreatePurchaseOrderFromScanBatch(lngVendorID) > 0 Then
                SetOptions False
                DoCmd.DeleteObject acTable, "CompareScanItemToPurchaseOrder"
                DoCmd.DeleteObject acTable, "CompareScanItemToPurchaseOrderScanCounts"

                DoCmd.OpenQuery ("qryCompareScanItemToPurchaseOrder_01_MakeTable")
                DoCmd.OpenQuery ("qryCompareScanItemToPurchaseOrder_02_ScanTotals")
                DoCmd.OpenQuery ("qryCompareScanItemToPurchaseOrder_03_UpdateQuantityScanned")
                DoCmd.OpenQuery ("qryCompareScanItemToPurchaseOrder_04_Append")
                DoCmd.OpenQuery ("qryCompareScanItemToPurchaseOrder_04a_Repacks")
                ExportScanBatchLineCounts
                DoCmd.RunSQL "Delete * from ScanBatchLineCounts Where ScanBatchHeaderID = " & glngScanBatchHeaderID
                Me.Requery
                SetupScreens
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
            txtCode.SetFocus
            .Visible = False
            cmdfrmClose.Visible = True


        End If
    End With

ExitSub:
    txtCode.SetFocus
    Exit Sub
ErrLog:
    StatusLine "Error: " & Err.Number & " " & Err.Description & "  lstVendor Exit", vbRed, True
    Resume ExitSub
End Sub

Private Sub lstVendor_DblClick(Cancel As Integer)
    MsgBox "Please Hit enter to select the Vendor!"

End Sub

Private Sub lstVendor_Exit(Cancel As Integer)
    Dim lngVendorID As Long
    Dim strVendor As String
    If gcfHandleErrors Then On Error GoTo ErrLog
    With lstVendor
        lngVendorID = .ItemData(.ListIndex + 1)
        'strVendor = lstVendor ' .List(.ListIndex, 1)
        If MsgBox("Create a purchase order for Vendor # " & lngVendorID, vbYesNo) = vbNo Then
        Else
            If CreatePurchaseOrderFromScanBatch(lngVendorID) > 0 Then
                SetOptions False
                ExportScanBatchLineCounts
                DoCmd.RunSQL "Delete * from ScanBatchLineCounts Where ScanBatchHeaderID = " & glngScanBatchHeaderID
                Me.Requery
                SetupScreens
                SetOptions True
                If gcfHandleErrors Then On Error GoTo ErrLog
                Me.Visible = False
                MsgBox "Purchase order created; the purchase order screen will be displayed."
                DoCmd.OpenForm "PurchaseOrderCreate", acNormal, "", "", , acNormal

                'DoCmd.Close acForm, "ScanBatch", acSavePrompt
                'lstLocation.Visible = True
                'lstLocation.SetFocus
                Exit Sub

            End If
            txtCode.SetFocus
            lstVendor.Visible = False
            cmdfrmClose.Visible = True


        End If
    End With

ExitSub:
    txtCode.SetFocus
    Exit Sub
ErrLog:
    StatusLine "Error: " & Err.Number & " " & Err.Description & "  lstVendor Exit", vbRed, True
    Resume ExitSub
End Sub

Private Sub ScanBatchLineCountsSub_Exit(Cancel As Integer)
    'lstEntries.Visible = True
    'lstEntries.Visible = False
End Sub



Private Sub optAutoEnter1Qty_Click()
    StatusLine "If '2. Backroom' is selected then Case Quantity is entered.", vbYellow, False
End Sub



'Private Sub txtCode_DblClick(Cancel As Integer)
'    DoCmd.OpenForm "ScanBatch1"
'End Sub

Private Sub txtCode_Enter()
    mbShowMoreOptions = False
    ShowControls
    txtCode = ""
    StatusLine "Scan on! ", vbYellow, False
End Sub

Private Sub txtCode_Exit(Cancel As Integer)
    Dim strSaveDesc As String
    Dim strResponse As String
   ' Find the record that matches the control.
    If gcfHandleErrors Then On Error GoTo ErrLog
    If Not IsNull(txtCode) Then txtCode = CheckScanSheet(txtCode)
    If Len(txtCode) = 0 Then
        Cancel = True
        Beep
        StatusLine "Scan Code can not be empty! (Press 99998 to bypass)", vbRed, True
        Exit Sub
    End If

    If Len(txtCode) < 4 Then 'the user wants to add quantity to last count
        Cancel = True
        If Not IsNumeric(txtCode) Then 'quantity has to be a number
            Beep
            StatusLine "Adding quantity to last scanned item.. must be numeric", vbRed, True
            BlankInfoTextBoxes
            Exit Sub
        End If

        If CLng(txtCode) > 100 Then
            StatusLine "Count must be less than 101!", vbRed, True
            Beep
            BlankInfoTextBoxes
            Exit Sub
        End If

        ScanBatchSub.Form.Requery
        Set mrsScanBatchSub = ScanBatchSub.Form.RecordsetClone
        With mrsScanBatchSub
            'empty list; nothing to add to
            If .RecordCount < 1 Then
                Beep
                StatusLine "Nothing in list to add!", vbRed, True
                BlankInfoTextBoxes
                Exit Sub
            End If

            'the first record is the most recent
            .MoveFirst
            txtCount = CInt(txtCode)
            txtCode = !ScanCode
            txtItemID = !ItemID
            txtDescription = !Description
            txtPrice = !Price
            txtOnShelf = !OnShelf
            txtInputMethod = "Add quantity"
        End With

        AddEntryToListBox
        StatusLine "Scan On!", vbYellow, True
        Exit Sub
    End If

    Select Case txtCode
    Case 99999 '- User wants to remove last Item
        Cancel = True
        ScanBatchSub.Form.Requery
        Set mrsScanBatchSub = ScanBatchSub.Form.RecordsetClone

        With mrsScanBatchSub
            'empty list; nothing to remove
            If .RecordCount < 1 Then
                Beep
                StatusLine "There are no more items to remove!", vbRed, True
                BlankInfoTextBoxes
                Exit Sub
            End If

            'the first record is the most recent
            .MoveFirst
            txtCount = 0 - !CountQuantity
            txtCode = !ScanCode
            txtItemID = !ItemID
            txtDescription = !Description
            txtPrice = !Price
            txtOnShelf = !OnShelf
            txtInputMethod = "Removed"
            strSaveDesc = txtDescription
            .Edit
            !InputMethod = "Removed"
            .Update
        End With
        AddEntryToListBox
        StatusLine "Removing Last item! " & strSaveDesc, vbYellow, True
    Case 99998 ' allow me to click somewhere
        mbShowMoreOptions = True
        ShowControls
        txtDescription.SetFocus
    Case Else
        With mrsScanBatch
            .FindFirst "[UPCCode] = '" & Me![txtCode] & "'"
             If Not .NoMatch Then
                txtItemID = !ID
                txtDescription = !Description
                txtPrice = Format(!Price, "#,##0.00")
                txtOnShelf = !OnShelf
                txtInputMethod = "Normal entry"
                If optPriceCheckMode Then
                    StatusLine txtPrice, vbYellow, True
                    StatusLine "Price Check Mode.. Data is not saved!", vbRed, False
                    txtInputMethod = "PRICE CHECK"
                    txtOnShelfLast = txtOnShelf
                    AddEntryToListBox
                    txtCount = "PRICE CHECK"
                    Cancel = True
                    'StatusLine txtPrice, vbYellow, True
                Else
                    If optAutoEnter1Qty Then
                        Cancel = True
                        If mbBackRoom Then
                            txtCount = !UnitsPerCase
                        Else
                            txtCount = 1
                        End If

                        'txtCode = lstEntries.Column(0, 1)
                        'txtItemID = lstEntries.Column(1, 1)
                        'txtDescription = lstEntries.Column(2, 1)
                        'txtPrice = lstEntries.Column(3, 1)
                        'txtOnShelf = lstEntries.Column(5, 1)
                        txtInputMethod = "AutoEnter 1 Q"
                        AddEntryToListBox
                        StatusLine "Scan On!", vbYellow, True
                        Exit Sub
                    Else
                        txtCount.SetFocus
                    End If
                End If
            Else
                StatusLine "Item not found", vbRed, True
                PlayNotOnFile
                Beep
'                Sleep (250)
'                Beep
'                Sleep (250)
'                Beep
                Cancel = True
                'added 1/30/2013
                strResponse = "."
                Do While strResponse <> "20" And strResponse <> "" 'cancel is hit..
                    Beep
                    strResponse = InputBox(txtCode & " not found!" & vbCrLf & vbCrLf & "Enter or Scan '20' to confirm", "NOT ON FILE!! Confirmation required", ".")

                    If Not IsNull(strResponse) Then strResponse = CheckScanSheet(strResponse)
                Loop

                If optPriceCheckMode Then
                    txtItemID = "000000"
                    txtDescription = "Not on file"
                    txtPrice = Format(0, "#,##0.00")
                    StatusLine "Price Check Mode.. Data is not saved!", vbRed, False
                    txtInputMethod = "Price Check Mode"
                    AddEntryToListBox
                    StatusLine "Item not found", vbRed, True
                    Cancel = True
                End If
                BlankInfoTextBoxes

                'txtCode = ""
                'txtDescription = ""
                'txtPrice = ""
            End If
        End With
    End Select
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "txtCode_Exit"
    StatusLine Err.Number & " " & Err.Description & " txtCode_Exit", vbRed, True
    Resume ExitSub
End Sub

Private Sub txtCount_Enter()
    StatusLine "Quantity? " & txtDescription, vbYellow, False
End Sub

Private Sub txtCount_Exit(Cancel As Integer)
    If gcfHandleErrors Then On Error GoTo ErrLog

    'is this off of the scan sheet?
    If Not IsNull(txtCount) Then txtCount = CheckScanSheet(txtCount)
    If Not IsNumeric(txtCount) Then
        Beep
        StatusLine "Count must be a number!", vbRed, True
        txtCount = ""
        Cancel = True
        Exit Sub
    End If

'    '- User wants to remove last Item
'    If txtCount = 99999 And lstEntries.ListCount > 1 Then
'        txtCode = lstEntries.Column(0, 1)
'        txtItemID = lstEntries.Column(1, 1)
'        txtDescription = lstEntries.Column(2, 1)
'        txtCount = 0 - lstEntries.Column(3, 1)
'        txtOnShelf = lstEntries.Column(4, 1)
'        txtInputMethod = "Remove last entry"
'        AddEntryToListBox
'        lstEntries.RemoveItem (1)
'        lstEntries.RemoveItem (1)
'        StatusLine "Removing Last item!", vbYellow, True
'        Exit Sub
'    End If

   If CLng(txtCount) > 100 Then
        StatusLine "Count must be less than 101!", vbRed, True
        Beep
        txtCount = ""
        Cancel = True
        Exit Sub
    End If

'    If CLng(txtCount) = 0 Then 'don't add a '0' count
'        txtCode = ""
'        txtDescription = ""
'        txtPrice = ""
'        txtCount = ""
'        Exit Sub
'    End If

    StatusLine "Adding scanned item to counts", vbYellow, False
    AddEntryToListBox

ExitSub:

    Exit Sub
ErrLog:
    StatusLine Err.Number & " " & Err.Description & " txtCount_Exit", vbRed, True
    Resume ExitSub
End Sub



Private Sub txtCountLast_Click()
    ShowScanBatchLineCounts
End Sub

'Private Sub txtScanCounts_Click()
'    DoCmd.OpenForm "ScanBatchTotalsSub"
'End Sub

Private Sub txtScanQ_Click()
    txtCountLast_Click
End Sub


Private Sub AddEntryToListBox()

    Dim strADD As String
    If gcfHandleErrors Then On Error GoTo ErrLog
    StatusLine "Adding to list", vbYellow, False

    If AddLineCounts Then
        ShowLastScanInfo
        strADD = txtCode & ";"
        strADD = strADD & txtItemID & ";"
        strADD = strADD & txtDescription & ";"
        strADD = strADD & txtPrice & ";"
        strADD = strADD & txtCount & ";"
        strADD = strADD & txtOnShelf & ";"
        strADD = strADD & txtLocation & ";"
        strADD = strADD & txtInputMethod & ";"
'        lstEntries.AddItem (strADD), 1
'        If lstEntries.ListIndex > nMaxListItems Then 'keep the list count low... it stops working at 1000 or so
'            lstEntries.RemoveItem (nMaxListItems)
'        End If
'
        txtCode = ""
        txtDescription = ""
        txtPrice = ""
        txtCount = ""
    End If
    ScanBatchSub.Requery
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "AddEntryToListBox"
    StatusLine Err.Number & " " & Err.Description & " AddEntryToListBox", vbRed, True
    Resume ExitSub
End Sub


Private Function AddLineCounts() As Boolean
    Dim nItemID As Double
    nItemID = CDbl(txtItemID)
    If gcfHandleErrors Then On Error GoTo ErrLog
    AddLineCounts = True
    If Not optPriceCheckMode Then
        With mrsScanBatchLineCounts
            .AddNew
            !ScanBatchHeaderID = glngScanBatchHeaderID
            !ItemID = nItemID
            !CashierID = glnCashierID
            !InputMethod = txtInputMethod
            !Location = txtLocation
            !ScanCode = txtCode
            !CountQuantity = txtCount
            !OnShelf = txtOnShelf
            !OnShelfDate = txtOnShelfDate
            .Update
        End With
    End If
ExitSub:
    Exit Function
ErrLog:
    AddLineCounts = False
    Beep
    Sleep 250
    Beep
    Sleep 250
    Beep
    Sleep 250

    StatusLine "SCAN NOT SAVED! " & Err.Number & " " & Err.Description & " AddLineCounts", vbRed, True
    MsgBox "SCAN NOT SAVED! " & Err.Number & " " & Err.Description, , "AddLineCounts"
    Resume ExitSub
End Function


Private Sub BlankInfoTextBoxes()
    txtCode = ""
    txtDescription = ""
    txtPrice = ""
    txtCount = ""
End Sub

Private Function CheckScanSheet(InputString As String) As String
    Dim intItemId As Double '07/17/2021
    If gcfHandleErrors Then On Error GoTo ErrLog
    CheckScanSheet = InputString
    'copyright KMDS :)
    If Len(InputString) = 12 And Left(InputString, 6) = "000000" Then
        CheckScanSheet = CStr(CLng(Mid(InputString, 7, 5)))
    End If
    'check for olcc label 001402
    If Len(InputString) = 6 And Left(InputString, 1) = "0" And IsNumeric(InputString) Then
        intItemId = CLng(InputString) 'effectively trim leading zeros from inputstring
        CheckScanSheet = CStr(intItemId)
    End If

ExitFunction:
    Exit Function
ErrLog:
    StatusLine Err.Number & " " & Err.Description & " Check Scan Sheet", vbRed, True
    Resume ExitFunction

End Function


Private Sub Sayit(message As String)


    Dim strLanguage    As String
    If gcfHandleErrors Then On Error GoTo ErrLog
    strLanguage = "EN"

    Dim strURL As String
    If optVoice Then
        strURL = Replace("http://translate.google.com/translate_tts?tl=" & strLanguage & "&q=" & message & "" & "&ie=UTF-8", "<div>", "")
        strURL = Replace(strURL, "</div>", "")
        WMP1.URL = strURL
    End If
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description & " Sayit"
    Resume ExitSub

End Sub

Private Sub SetupScreens()
    Dim nI As Integer
    Dim strADD As String
    If gcfHandleErrors Then On Error GoTo ErrLog

    'first populate the list box with previous entries
'    For ni = 0 To lstEntries.ListCount - 1
'        lstEntries.RemoveItem (0)
'    Next
    strADD = "UPC Code" & ";"
    strADD = strADD & "Item Id" & ";"
    strADD = strADD & "Description" & ";"
    strADD = strADD & "Price" & ";"
    strADD = strADD & "Count" & ";"
    strADD = strADD & "Shelf" & ";"
    strADD = strADD & "Location" & ";"
    strADD = strADD & "InputMethod" & ";"
'    lstEntries.AddItem (strADD), 0
    'cmbEntries.AddItem (strAdd), 0
    Set mrsSBLC = Me.RecordsetClone
    With mrsSBLC
        If .RecordCount > 0 Then
            .MoveFirst
            txtFirstScanDate = !SetupDate
            .MoveLast
            txtLastScanDate = !SetupDate
            txtScanCounts = .RecordCount
            '.MoveFirst
            Do While Not .BOF
                strADD = !ScanCode & ";"
                strADD = strADD & !ItemID & ";"
                strADD = strADD & !Description & ";"
                strADD = strADD & Format(!Price, "#,##0.00") & ";"
                strADD = strADD & !CountQuantity & ";"
                strADD = strADD & !OnShelf & ";"
                strADD = strADD & !Location & ";"
                strADD = strADD & !InputMethod & ";"
'                lstEntries.AddItem (strADD) ', 1
                'cmbEntries.AddItem (strAdd), 1
                '.MoveNext
                .MovePrevious
            Loop
        End If
    End With

    '2. format the hidden form so it is exactly beneath list box
'    With lstEntries
'        .Top = 3200
'        .Left = 330
'        .Height = 3120
'        .Width = 14000
'
'
''        ScanBatchLineCountsSub.Top = .Top
''        ScanBatchLineCountsSub.Left = .Left
''        ScanBatchLineCountsSub.Height = .Height
''        ScanBatchLineCountsSub.Width = .Width
'    End With

    '3. hide and seek with the 2 controls
'    lstEntries.Visible = True
'    lstEntries.Visible = False

'    ScanBatchLineCountsSub.Visible = False
'    ScanBatchLineCountsSubLabel.Visible = False

    '4. locations are stored in the settings tables
    For nI = 0 To lstLocation.ListCount - 1
        lstLocation.RemoveItem (0)
    Next

    'up to 8 locations.. first 4 are provided
    lstLocation.AddItem "Select a Location"
    strADD = GetKMDSSettings("ScanBatchLoc1", "1. Storeshelf")
    lstLocation.AddItem strADD
    strADD = GetKMDSSettings("ScanBatchLoc2", "2. Backroom")
    lstLocation.AddItem strADD
    strADD = GetKMDSSettings("ScanBatchLoc3", "3. Overflow")
    lstLocation.AddItem strADD
    strADD = GetKMDSSettings("ScanBatchLoc4", "4. Display")
    lstLocation.AddItem strADD
'    strAdd = GetKMDSSettings("ScanBatchLoc5", "5. Not Assigned")
'    If strAdd <> "5. Not Assigned" Then lstLocation.AddItem strAdd
'    strAdd = GetKMDSSettings("ScanBatchLoc6", "6. Not Assigned")
'    If strAdd <> "6. Not Assigned" Then lstLocation.AddItem strAdd
'    strAdd = GetKMDSSettings("ScanBatchLoc7", "7. Not Assigned")
'    If strAdd <> "7. Not Assigned" Then lstLocation.AddItem strAdd
'    strAdd = GetKMDSSettings("ScanBatchLoc8", "8. Not Assigned")
'    If strAdd <> "8. Not Assigned" Then lstLocation.AddItem strAdd

    '6. Set the screen size
    Me.InsideHeight = Box67.Height + FormFooter.Height + 200
    Me.InsideWidth = Box67.Width + 200

    '7 the location list box ???
    With lstLocation
        .Top = txtCode.Top
        .Width = txtDescription.Width
        .Height = txtCode.Height * 4
        lstVendor.Top = .Top
        lstVendor.Left = Label1.Left
        lstClubs.Top = .Top
        lstClubs.Left = Label1.Left
    End With
    With lstAgents
        .Top = txtCode.Top
        .Width = txtDescription.Width
        .Height = txtCode.Height * 4
        lstVendor.Top = .Top
        lstVendor.Left = Label1.Left
        lstClubs.Top = .Top
        lstClubs.Left = Label1.Left
    End With

    lstVendor.Visible = False

ExitSub:
    Exit Sub
ErrLog:
    StatusLine Err.Number & " " & Err.Description & " ClearLog", vbRed, True
    Resume ExitSub
End Sub

Private Sub ShowControls()
    If gcfHandleErrors Then On Error GoTo ErrLog
    lstLocation.Visible = False
    lstVendor.Visible = False
    lstClubs.Visible = False
    optPriceCheckMode.Visible = mbShowMoreOptions
    optVoice.Visible = mbShowMoreOptions
    optAutoEnter1Qty.Visible = mbShowMoreOptions
    txtCountLast.Visible = Not optPriceCheckMode
    txtVar.Visible = Not optPriceCheckMode
    txtScanCounts.Visible = mbShowMoreOptions
    txtScanQ.Visible = mbShowMoreOptions
    txtFirstScanDate.Visible = mbShowMoreOptions
    txtLastScanDate.Visible = mbShowMoreOptions

    cmd65VarReport.Visible = mbShowMoreOptions
    'cmdBatchScanReport.Visible = mbShowMoreOptions
    cmdBatchReport.Visible = mbShowMoreOptions
    cmdEraseBatch.Visible = mbShowMoreOptions
    If cmdEraseBatch.Visible Then
        cmdEraseBatch.Visible = gbAllowDeleteDocument
    End If
    cmdAppendToAudit.Visible = mbShowMoreOptions
    cmdCreatePO.Visible = mbShowMoreOptions
    cmdCreateClubOrder.Visible = mbShowMoreOptions
    cmdComparePO.Visible = mbShowMoreOptions
    cmdCreateTransfer.Visible = mbShowMoreOptions
    cmdCreateClaim.Visible = mbShowMoreOptions

    cmdfrmClose.Visible = mbShowMoreOptions

    'lstEntries.Visible = Not mbShowMoreOptions
    'lstEntries.Visible = False

    ScanBatchSub.Visible = Not mbShowMoreOptions

    Set mrsSBLC = Me.RecordsetClone
    With mrsSBLC
        If .RecordCount = 0 Then
            StatusLine "The current batch is empty!", vbYellow, True

            cmd65VarReport.Visible = False
            cmdBatchReport.Visible = False
            'cmdBatchScanReport.Visible = False
            cmdEraseBatch.Visible = False
            cmdAppendToAudit.Visible = False
            cmdCreatePO.Visible = False
            cmdComparePO.Visible = False
            cmdCreateClubOrder.Visible = False
            cmdCreateTransfer.Visible = False
            cmdCreateClaim.Visible = False

            cmdRestoreBatch.Visible = True
        Else
            cmdRestoreBatch.Visible = False
        End If
    End With
ExitSub:
    Exit Sub
ErrLog:
    StatusLine Err.Number & " " & Err.Description & " Show controls", vbRed, True
    Resume ExitSub


End Sub
Private Sub ShowLastScanInfo()

    Dim strSQL As String
    Dim rs As Recordset
    If gcfHandleErrors Then On Error GoTo ErrLog
    txtItemIDLast = txtItemID
    txtCodeLast = txtCode
    txtDescriptionLast = txtDescription
    txtLastScanDate = Now
    txtPriceLast = txtPrice
    StatusLine "ShowLastScanInfo - RecordCounts", vbYellow, False

    With mrsSBLC
        .Requery
        If .RecordCount > 0 Then .MoveLast
        txtScanCounts = .RecordCount
    End With

    StatusLine "ShowLastScanInfo - ShowVariance", vbYellow, False
    strSQL = "SELECT ItemID, Sum(CountQuantity) AS CountQ, Last(OnShelf) AS OS, Count(CountQuantity) AS ScanQ "
    strSQL = strSQL & "From ScanBatchLineCounts GROUP BY ScanBatchLineCounts.ItemID "
    strSQL = strSQL & "HAVING (ItemID= " & CDbl(txtItemID) & " )"
    Set rs = CurrentDb().OpenRecordset(strSQL)

    With rs
        If .RecordCount Then
            .MoveFirst
            txtCountLast = !Countq
            txtOnShelfLast = !OS
            txtScanQ = !ScanQ
            txtVar = !Countq - !OS
            If txtVar = 0 Then txtVar.BackColor = vbWhite
            If txtVar < 0 Then txtVar.BackColor = vbRed
            If txtVar > 0 Then txtVar.BackColor = vbGreen
            lblScanQ.Caption = txtItemID
            lblCountLast.Caption = "Count (" & !ScanQ & ")"
        End If
    End With
    rs.Close

    Set rs = Nothing

    StatusLine "ShowLastScanInfo - ScanCounts", vbYellow, False
    strSQL = "SELECT Sum(CountQuantity) AS CountQ "
    strSQL = strSQL & "From ScanBatchLineCounts "
    Set rs = CurrentDb().OpenRecordset(strSQL)

    With rs
        If .RecordCount Then
            .MoveFirst
            txtTotalQty = !Countq
        Else
            txtTotalQty = "0"

        End If
    End With
    rs.Close
    Set rs = Nothing

ExitSub:
    Exit Sub
ErrLog:
    StatusLine Err.Number & " " & Err.Description & " Show Last Scan Info", vbRed, True
    Resume ExitSub
End Sub

Private Sub ShowScanBatchLineCounts()
    If gcfHandleErrors Then On Error GoTo ErrLog
    Set mrsSBLC = Me.RecordsetClone
    With mrsSBLC
        txtCodeLast = txtCode
        .FindFirst "[ItemID] = " & txtItemIDLast
        If .EOF Then
            Beep
            Beep
            Beep
        Else
            Me.Bookmark = .Bookmark
            'lstEntries.Visible = False
'            ScanBatchLineCountsSub.Visible = True
'            ScanBatchLineCountsSubLabel.Visible = True
'            ScanBatchLineCountsSub.SetFocus
        End If
    End With
ExitSub:
    Exit Sub
ErrLog:
    StatusLine Err.Number & " " & Err.Description & " Show Scan Batch Line Counts", vbRed, True
    Resume ExitSub
End Sub

Private Sub StatusLine(message As String, Color As Long, PleaseSayit As Boolean)
    If gcfHandleErrors Then On Error GoTo ErrLog
    txtStatus = message
    txtStatus.BackColor = Color
    If PleaseSayit Then Sayit (message)
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description & " StatusLine"
    Resume ExitSub
End Sub


Public Sub ExportScanBatchLineCounts()
Dim strOutFname As String
On Error Resume Next
MkDir ("C:\kmds\SBLC")
If gcfHandleErrors Then On Error GoTo ErrLog

DoCmd.TransferText acExportDelim, , "qryScanBatchLineCountsExport", "C:\KMDS\SBLC\sblc" & Format(Now, "yyyy-mm-dd-hh-nn-ss") & ".txt", True
ExitSub:
    Exit Sub
ErrLog:
    StatusLine Err.Number & " " & Err.Description & " Export Scan Batch Line Counts", vbRed, True
    Resume ExitSub
End Sub

Private Sub PlayNotOnFile()
'    If Dir("C:\KMDS\Sounds\Laser.Wav") <> "" Then
'        With WMP1
'            .URL = "C:\KMDS\Sounds\Laser.Wav"
'            .
'        End With
'
'    End If
''    Sleep 1500
    If Dir("C:\KMDS\Sounds\NotOnFile.Wav") <> "" Then
        With WMP1
            .URL = "C:\KMDS\Sounds\NotOnFile.Wav"

        End With

    End If

End Sub
