Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Option Explicit

Private Sub Check81_Click()

If gcfHandleErrors Then On Error GoTo ErrLog
    If Check81 Then
        Me.AgentsTransfersSub.Top = boxPurchasedFrom.Top
        Me.AgentsTransfersSub.Height = FormFooter.Height + BoxBody.Height - 250
        FormFooter.Visible = False
    Else
        Me.AgentsTransfersSub.Height = boxSubFormNormal.Height - 250
        Me.AgentsTransfersSub.Top = boxSubFormNormal.Top
        FormFooter.Visible = True
    End If
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Check81_Click"
    Resume ExitSub
End Sub

Private Sub chkDisplayCost_Click()
    ShowColumns
End Sub

Private Sub chkDisplayDescription_Click()
    ShowColumns
End Sub

Private Sub chkDisplayExtCost_Click()
    ShowColumns
End Sub

Private Sub chkDisplayExtPrice_Click()
    ShowColumns
End Sub

Private Sub chkDisplayGP_Click()
    ShowColumns
End Sub

Private Sub chkDisplayItemID_Click()
    ShowColumns
End Sub

Private Sub chkDisplayLine_Click()
    ShowColumns
End Sub

Private Sub chkDisplayPrice_Click()
    ShowColumns
End Sub

Private Sub ShowColumns()
If gcfHandleErrors Then On Error GoTo ErrLog
    Me.AgentsTransfersSub.Form![LineID].ColumnHidden = Not chkDisplayLine
    Me.AgentsTransfersSub.Form![QuantityShipped].ColumnHidden = Not chkDisplayQuantity
    Me.AgentsTransfersSub.Form![ItemID].ColumnHidden = Not chkDisplayItemID
    Me.AgentsTransfersSub.Form![ItemDescription].ColumnHidden = Not chkDisplayDescription
    Me.AgentsTransfersSub.Form![Size].ColumnHidden = Not chkDisplaySize
    Me.AgentsTransfersSub.Form![Price].ColumnHidden = Not chkDisplayPrice
    Me.AgentsTransfersSub.Form![ExtPrice].ColumnHidden = Not chkDisplayExtPrice

    Me.AgentsTransfersSub.Form![LineID].ColumnOrder = 1
    Me.AgentsTransfersSub.Form![QuantityShipped].ColumnOrder = 2
    Me.AgentsTransfersSub.Form![ItemID].ColumnOrder = 3
    Me.AgentsTransfersSub.Form![ClaimCode].ColumnOrder = 4
    Me.AgentsTransfersSub.Form![ItemDescription].ColumnOrder = 5
    Me.AgentsTransfersSub.Form![Size].ColumnOrder = 6
    Me.AgentsTransfersSub.Form![Price].ColumnOrder = 7
    Me.AgentsTransfersSub.Form![ExtPrice].ColumnOrder = 8


    'ADDED 2007-12-29
    Select Case Status
    Case "O" 'TRANSFER OUT
        Me.AgentsTransfersSub.Form![ClaimCode].ColumnHidden = True
    Case "I" 'TRANSFER IN
        Me.AgentsTransfersSub.Form![ClaimCode].ColumnHidden = True
    Case "C" ' OLCC CLAIM
        Me.AgentsTransfersSub.Form![ClaimCode].ColumnHidden = False
'        Me.AgentsTransfersSub.Form![Cost].ColumnHidden = True
'        Me.AgentsTransfersSub.Form![ExtCost].ColumnHidden = True
    End Select
    Me.AgentsTransfersSub.Form![LineID].ColumnWidth = 650
    Me.AgentsTransfersSub.Form![ItemID].ColumnWidth = 1500
    Me.AgentsTransfersSub.Form![ItemDescription].ColumnWidth = 4000



ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "ShowColumns"
    Resume ExitSub
End Sub

Private Sub chkDisplayQuantity_Click()
    ShowColumns

End Sub


Private Sub chkDisplayRepack_Click()
    ShowColumns
End Sub

Private Sub chkDisplaySize_Click()
    ShowColumns
End Sub

'Private Sub cmbPurchaseOrderId_AfterUpdate()
'    ' Find the record that matches the control.
'    Dim rs As Object
'    If gcfHandleErrors Then On Error GoTo ErrLog
'    Set rs = Me.Recordset.Clone
'    rs.FindFirst "[ID] = " & Str(Me![cmbPurchaseOrderId])
'    Me.Bookmark = rs.Bookmark
'ExitSub:
'    Exit Sub
'
'ErrLog:
'    MsgBox Err.Description
'    Resume ExitSub
'End Sub


Private Sub cmdfrmDeleteRecord_Click()

If gcfHandleErrors Then On Error GoTo Err_cmdDelOrder_Click

    If gbAllowDeleteDocument Then
        DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
        DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
        DoCmd.Close
    Else
        MsgBox "This user is Not allowed to delete documents", vbCritical, glsCashierName
    End If

Exit_cmdDelOrder_Click:
    Exit Sub

Err_cmdDelOrder_Click:
    MsgBox Err.Description
    Resume Exit_cmdDelOrder_Click


End Sub


Private Sub cmdClose_Click()

If gcfHandleErrors Then On Error GoTo Err_cmdClose_Click


    DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click

End Sub



Private Sub cmdVendorPOHistory_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdVendorPOHistory_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "VendorPOHistory"

    stLinkCriteria = "[ID]=" & Me![VendorID]
    DoCmd.OpenForm stDocName, , , stLinkCriteria


Exit_cmdVendorPOHistory_Click:
    Exit Sub

Err_cmdVendorPOHistory_Click:
    MsgBox Err.Description
    Resume Exit_cmdVendorPOHistory_Click
End Sub

'Private Sub cmdPrint_Click()
'If gcfHandleErrors Then On Error GoTo ErrLog
'
'    If ID > 0 Then
'        DoCmd.OpenForm "PurchaseOrderChoices", acNormal, , , , acDialog
'        'MsgBox "Response = " & Forms("PurchaseOrderChoices").optChoices
'        Select Case Forms("PurchaseOrderChoices").optChoices
'        Case 0
'        Case 1
'            'do nothing, just unload the form
'        Case 2
'            'print the po and then unload the form
'            PrintPO False
'        Case 3
'            'post the order
'            If MsgBox("Post this order?", vbYesNo) = vbYes Then
'                'archive po print
'                'post order!
'            End If
'        Case 4
'            Cancel = True
'        Case 5
'            'print the po and then unload the form
'            PrintPO True
'
'        Case Else
'            MsgBox "Option " & Forms("PurchaseOrderChoices").optChoices & " not handled!"
'        End Select
'        DoCmd.Close acForm, "PurchaseOrderChoices", acSaveYes
'    End If
'ExitSub:
'    Exit Sub

'ErrLog:
'    MsgBox Err.Description
'    Resume ExitSub
'End Sub

Private Sub Form_Current()
If gcfHandleErrors Then On Error GoTo ErrLog
    If Me.NewRecord Then
    Else
        With cmbAgentID
            txtContact = .Column(2)
            txtAdd1 = .Column(3)
            txtAdd2 = .Column(4) & ", " & .Column(5) & " " & .Column(6)
            txtPhone1 = .Column(7)
            'chkSupplyAllItems = .Column(10) ' PrintUpcOnPo
        End With
        'cmbPurchaseOrderId = ID
        AgentsTransfersSub.Form.ItemDescription.Requery
        Select Case Status
        Case "I" 'in
            lblTransferINOUT.Caption = "Transfer IN"


        Case "O" 'out
            lblTransferINOUT.Caption = "Transfer OUT"
        Case "C" 'out
            lblTransferINOUT.Caption = "OLCC CLAIMS - Transfer OUT "
        Case Else 'error!
            MsgBox "Status not In or Out, please delete and start over!"
            lblTransferINOUT.Caption = "Delete this Transfer"
        End Select


    End If
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Current"
    Resume ExitSub
End Sub

Private Sub Form_Delete(Cancel As Integer)
    If gbAllowDeleteDocument Then
    Else
        MsgBox "This user is Not allowed to delete documents", vbCritical, glsCashierName
        Cancel = True
    End If


End Sub

Private Sub Form_Load()

If gcfHandleErrors Then On Error GoTo ErrLog

'
    Me.Form.Width = BoxBody.Width

'    If Me.OpenArgs = "PurchaseOrderCreate" Then 'just work on the order created by clubordercreate form
'        cmbPurchaseOrderId.Visible = False
'    Else
'        cmbPurchaseOrderId.Visible = True
'
'    End If

    Check81_Click
    ShowColumns

    'cmbPurchaseOrderId.BackColor = glnColorDropDown

    cmbAgentID.BackColor = glnColorInfo
    txtContact.BackColor = glnColorInfo
    txtAdd1.BackColor = glnColorInfo
    txtAdd2.BackColor = glnColorInfo
    txtAdd3.BackColor = glnColorInfo
    ID.BackColor = glnColorInfo
    OrdDate.BackColor = glnColorInfo
    txtPhone1.BackColor = glnColorInfo
    Status.BackColor = glnColorInfo

    txtTotalPrice.BackColor = glnColorTotals
    txtTotalCost.BackColor = glnColorTotals
    'txtTotalGP.BackColor = glnColorTotals
    txtTotalQ.BackColor = glnColorTotals



    'Me.Form.InsideHeight = BoxBody.Height + Me.FormFooter.Height + 125
    optSortBy = 2
    optSortBy = 1

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Load"
    Resume ExitSub
End Sub

Private Sub Form_Open(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog
    'Cancel = FormLogin(glnCashierID, Me)
    SetFormSize Me
    If Me.Recordset.RecordCount = 0 Then
        MsgBox "There are no outstanding orders!"
        Cancel = True
    End If

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Open"
    Resume ExitSub
End Sub

Private Sub Form_Resize()
'    If gcfHandleErrors Then On Error Resume Next
'        AgentsTransfersSub.Width = Me.Form.InsideWidth - 750
'        AgentsTransfersSub.Form.InsideHeight = Me.Form.InsideHeight - (Me.FormHeader.Height + Me.FormFooter.Height + 4000)
'

End Sub

Private Sub PrintPO(PrintPreView As Boolean)
If gcfHandleErrors Then On Error GoTo ErrLog

    Dim stDocName As String

    stDocName = "PurchaseOrder"
    If cmbAgentID.Column(9) <> "" Then
        stDocName = cmbAgentID.Column(9)
    End If
    If PrintPreView Then
        DoCmd.OpenReport stDocName, acPreview
    Else
        DoCmd.OpenReport stDocName, acViewNormal
    End If

ExitSub:
    Exit Sub

ErrLog:
    MsgBox Err.Description
    Resume ExitSub
End Sub




Private Sub optSortBy_AfterUpdate()
If gcfHandleErrors Then On Error GoTo ErrLog

    Select Case optSortBy
    Case 1 'line id
        Me.AgentsTransfersSub.Form.OrderBy = "LineID"
    Case 2 'Item id
            Me.AgentsTransfersSub.Form.OrderBy = "ItemID"
    Case 3 'description
        Me.AgentsTransfersSub.Form.OrderBy = "Lookup_ItemDescription.Description"
    Case 4 'quantity
        Me.AgentsTransfersSub.Form.OrderBy = "QuantityShipped"
    Case 5 'cost
        Me.AgentsTransfersSub.Form.OrderBy = "Cost"
    Case 6 'ExtCost
        Me.AgentsTransfersSub.Form.OrderBy = "ExtCost"
    Case 7 'Price
        Me.AgentsTransfersSub.Form.OrderBy = "Price"
    Case 8 'ExtPrice
        Me.AgentsTransfersSub.Form.OrderBy = "ExtPrice"
    Case 9 'Repack
        Me.AgentsTransfersSub.Form.OrderBy = "Repack DESC, ItemID"
    Case 10 'gp
        Me.AgentsTransfersSub.Form.OrderBy = "GP, ItemID"
    End Select
    Me.AgentsTransfersSub.Form.OrderByOn = True
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "optSortBy_AfterUpdate"
    Resume ExitSub
End Sub


Private Sub cmdPrintDocs_Click()

If gcfHandleErrors Then On Error GoTo Err_cmdPrintDocs_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "AgentsTransfersChoices"
    'used for building ftp script
    stLinkCriteria = "[AgentNumber]=" & Me![AgentID]
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , ID ' cmbPurchaseOrderId
    DoCmd.Close acForm, Me.Name
Exit_cmdPrintDocs_Click:
    Exit Sub

Err_cmdPrintDocs_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrintDocs_Click

End Sub

Private Sub cmdRefresh_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdRefresh_Click

    Me.AgentsTransfersSub("ItemID").Requery
    Me.AgentsTransfersSub("ItemDescription").Requery
   DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70

Exit_cmdRefresh_Click:
    Exit Sub

Err_cmdRefresh_Click:
    MsgBox Err.Description
    Resume Exit_cmdRefresh_Click

End Sub
