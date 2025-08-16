Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Option Explicit

Dim bOKToExit As Boolean


Private Sub Check81_Click()

If gcfHandleErrors Then On Error GoTo ErrLog
    If Check81 Then
        Me.PurchaseOrdersOnFileSub.Top = BoxBody.Top + 750
        Me.PurchaseOrdersOnFileSub.Height = BoxBody.Height - 1000
        'Me.PurchaseOrdersOnFileSub.Height = FormFooter.Height + BoxBody.Height - 250
        'FormFooter.Visible = False
    Else
        Me.PurchaseOrdersOnFileSub.Height = boxSubFormNormal.Height - 250
        Me.PurchaseOrdersOnFileSub.Top = boxSubFormNormal.Top
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
    Me.PurchaseOrdersOnFileSub.Form![LineID].ColumnHidden = Not chkDisplayLine
    Me.PurchaseOrdersOnFileSub.Form![QuantityShipped].ColumnHidden = Not chkDisplayQuantity
    Me.PurchaseOrdersOnFileSub.Form![ItemID].ColumnHidden = Not chkDisplayItemID
    Me.PurchaseOrdersOnFileSub.Form![ItemDescription].ColumnHidden = Not chkDisplayDescription
    Me.PurchaseOrdersOnFileSub.Form![Size].ColumnHidden = Not chkDisplaySize
    Me.PurchaseOrdersOnFileSub.Form![Price].ColumnHidden = Not chkDisplayPrice
    Me.PurchaseOrdersOnFileSub.Form![ExtPrice].ColumnHidden = Not chkDisplayExtPrice
    Me.PurchaseOrdersOnFileSub.Form![Cost].ColumnHidden = Not chkDisplayCost
    Me.PurchaseOrdersOnFileSub.Form![ExtCost].ColumnHidden = Not chkDisplayExtCost
    Me.PurchaseOrdersOnFileSub.Form![Repack].ColumnHidden = Not chkDisplayRepack
    Me.PurchaseOrdersOnFileSub.Form![GP].ColumnHidden = Not chkDisplayGP
    Me.PurchaseOrdersOnFileSub.Form![PrintLabels].ColumnHidden = Not chkDisplayPrintLabels

    Me.PurchaseOrdersOnFileSub.Form![LineID].ColumnOrder = 1
    Me.PurchaseOrdersOnFileSub.Form![QuantityShipped].ColumnOrder = 2
    Me.PurchaseOrdersOnFileSub.Form![ItemID].ColumnOrder = 3
    Me.PurchaseOrdersOnFileSub.Form![ItemDescription].ColumnOrder = 4
    Me.PurchaseOrdersOnFileSub.Form![Size].ColumnOrder = 5
    Me.PurchaseOrdersOnFileSub.Form![Price].ColumnOrder = 6
    Me.PurchaseOrdersOnFileSub.Form![Sl].ColumnOrder = 7
    Me.PurchaseOrdersOnFileSub.Form![ExtPrice].ColumnOrder = 8
    Me.PurchaseOrdersOnFileSub.Form![Cost].ColumnOrder = 9
    Me.PurchaseOrdersOnFileSub.Form![ExtCost].ColumnOrder = 10
    Me.PurchaseOrdersOnFileSub.Form![GP].ColumnOrder = 11
    Me.PurchaseOrdersOnFileSub.Form![Repack].ColumnOrder = 12
    Me.PurchaseOrdersOnFileSub.Form![PrintLabels].ColumnOrder = 13



    Me.PurchaseOrdersOnFileSub.Form![ExtPriceRepack].ColumnHidden = True
    Me.PurchaseOrdersOnFileSub.Form![QtyRepack].ColumnHidden = True
    Me.PurchaseOrdersOnFileSub.Form![ExtPriceCase].ColumnHidden = True
    Me.PurchaseOrdersOnFileSub.Form![QtyCase].ColumnHidden = True

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "ShowColumns"
    Resume ExitSub
End Sub

Private Sub chkDisplayPrintLabels_Click()
    ShowColumns
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

Private Sub chkShowChart_AfterUpdate()
    PurchaseOrdersOnFileChartSalesTotalSub.Visible = chkShowChart
End Sub

Private Sub cmdFixUPC_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    AppendUPCItemsWithItemID
ExitSub:
    SetOptions True
    Exit Sub
ErrLog:
    EMsgBox Me, "cmdFixUPC_Click" & vbCrLf & Err.Number & " " & Err.Description
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


    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
    DoCmd.Close
Exit_cmdDelOrder_Click:
    Exit Sub

Err_cmdDelOrder_Click:
    MsgBox Err.Description
    Resume Exit_cmdDelOrder_Click


End Sub


Private Sub cmdClose_Click()

If gcfHandleErrors Then On Error Resume Next ' GoTo Err_cmdClose_Click

    bOKToExit = True

    DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click

End Sub



Private Sub cmdPrintLabels_Click()
If gcfHandleErrors Then On Error GoTo ErrLog
    If Not chkDisplayPrintLabels Then
        MsgBox "Setting the Labels column visible" & vbCrLf & "Use Labels Checkbox to print labels!", vbInformation
        chkDisplayPrintLabels.Value = True
        chkDisplayRepack.Value = False
        ShowColumns
        Exit Sub
    End If

    Dim stDocName As String
    Dim strSQL As String

    'Dim comQDF As New Command
    Dim comQDF As Object
    Set comQDF = CreateObject("ADODB.COMMAND")


    SetOptions False
    'delete work table
    If gcfHandleErrors Then On Error Resume Next
    DoCmd.DeleteObject acTable, "ItemLabels"

    If gcfHandleErrors Then On Error GoTo ErrLog
    'append all items to work table
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "qryItemLabelMakeTable"
    End With
    comQDF.Execute
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")


    'set the current record to 'selected'
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "qryItemLabelsUpdateLabelsFromPO"
        .Parameters.Append .CreateParameter("parItemID", adBigInt, adParamInput)
        .Parameters("parItemID") = ID
    End With
    comQDF.Execute
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    SetOptions True

    'get busy with label printing!
    DoCmd.OpenForm "ItemLabels"

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "CreateLabels" & vbCrLf & " Error #" & Err.Number & " " & Err.Description
    Resume ExitSub
End Sub

Private Sub cmdShowDetail_Click()
If gcfHandleErrors Then On Error GoTo Err_Command144_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "PrintInventoryReport"

    stLinkCriteria = "[ReportName]=" & "'18 Item Movement'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , cmdShowDetail.Tag

Exit_Command144_Click:
    Exit Sub

Err_Command144_Click:
    MsgBox Err.Description
    Resume Exit_Command144_Click
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

Private Sub Command124_Click()

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

        With cmbVendorID
            txtContact = .Column(2)
            txtAdd1 = .Column(3)
            txtAdd2 = .Column(4)
            txtAdd3 = .Column(5) & ", " & .Column(6) & " " & .Column(7)
            txtPhone1 = .Column(8)
            chkSupplyAllItems = .Column(10) ' PrintUpcOnPo
        End With
        'cmbPurchaseOrderId = ID
        PurchaseOrdersOnFileSub.Form.ItemDescription.Requery
    End If
    If cmbVendorID = 1 Then
        chkDisplayCost = False
        chkDisplayExtCost = False
        chkDisplayRepack = False

        Check81_Click
        ShowColumns

    End If

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Current"
    Resume ExitSub
End Sub

Private Sub Form_Load()

If gcfHandleErrors Then On Error GoTo ErrLog

    bOKToExit = False
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

    cmbVendorID.BackColor = glnColorInfo
    txtContact.BackColor = glnColorInfo
    txtAdd1.BackColor = glnColorInfo
    txtAdd2.BackColor = glnColorInfo
    txtAdd3.BackColor = glnColorInfo
    ID.BackColor = glnColorInfo
    OrdDate.BackColor = glnColorInfo
    txtPhone1.BackColor = glnColorInfo

    txtTotalPrice.BackColor = glnColorTotals
    txtTotalCost.BackColor = glnColorTotals
    txtTotalGP.BackColor = glnColorTotals
    txtTotalQ.BackColor = glnColorTotals

    txtTotalPriceRepacks.BackColor = glnColorTotals
    txtTotalQRepacks.BackColor = glnColorTotals
    txtTotalPriceCases.BackColor = glnColorTotals
    txtTotalQCases.BackColor = glnColorTotals

    'Me.Form.InsideHeight = BoxBody.Height + Me.FormFooter.Height + 125
    If cmbVendorID = 1 Then
        optSortBy = 2
    Else
    End If
    optSortBy_AfterUpdate

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
'        PurchaseOrdersOnFileSub.Width = Me.Form.InsideWidth - 750
'        PurchaseOrdersOnFileSub.Form.InsideHeight = Me.Form.InsideHeight - (Me.FormHeader.Height + Me.FormFooter.Height + 4000)
'

End Sub

Private Sub PrintPO(PrintPreView As Boolean)
If gcfHandleErrors Then On Error GoTo ErrLog

    Dim stDocName As String

    stDocName = "PurchaseOrder"
    If cmbVendorID.Column(9) <> "" Then
        stDocName = cmbVendorID.Column(9)
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




Private Sub Form_Unload(Cancel As Integer)
    If gbRunMailChanges Then MailChanges False

    If bOKToExit Then
        Exit Sub
    End If
    'If MsgBox("Exit Purchase Order Program?", vbYesNo) = vbNo Then
    '    Cancel = True
    'End If
End Sub

Private Sub optSortBy_AfterUpdate()
If gcfHandleErrors Then On Error GoTo ErrLog

    Select Case optSortBy
    Case 1 'line id
        Me.PurchaseOrdersOnFileSub.Form.OrderBy = "LineID"
    Case 2 'Item id
            Me.PurchaseOrdersOnFileSub.Form.OrderBy = "ItemID"
    Case 3 'description
        Me.PurchaseOrdersOnFileSub.Form.OrderBy = "Lookup_ItemDescription.Description"
    Case 4 'quantity
        Me.PurchaseOrdersOnFileSub.Form.OrderBy = "QuantityShipped"
    Case 5 'cost
        Me.PurchaseOrdersOnFileSub.Form.OrderBy = "Cost"
    Case 6 'ExtCost
        Me.PurchaseOrdersOnFileSub.Form.OrderBy = "ExtCost"
    Case 7 'Price
        Me.PurchaseOrdersOnFileSub.Form.OrderBy = "Price"
    Case 8 'ExtPrice
        Me.PurchaseOrdersOnFileSub.Form.OrderBy = "ExtPrice"
    Case 9 'Repack
        Me.PurchaseOrdersOnFileSub.Form.OrderBy = "Repack DESC, ItemID"
    Case 10 'gp
        Me.PurchaseOrdersOnFileSub.Form.OrderBy = "GP, ItemID"
    'changed 08/31/2004
    Case 11 'size
        Me.PurchaseOrdersOnFileSub.Form.OrderBy = "Lookup_Size.PkgSize, ItemID"
    Case 12 'print Labels
        Me.PurchaseOrdersOnFileSub.Form.OrderBy = "PrintLabels DESC, ItemID"
    End Select
    Me.PurchaseOrdersOnFileSub.Form.OrderByOn = True
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
    bOKToExit = True

    stDocName = "PurchaseOrderChoices"

    stLinkCriteria = "[ID]=" & Me![VendorID]
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , ID ' cmbPurchaseOrderId
    DoCmd.Close acForm, Me.Name
Exit_cmdPrintDocs_Click:
    Exit Sub

Err_cmdPrintDocs_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrintDocs_Click

End Sub

Private Sub cmdRefresh_Click()
    Dim rs
    Dim strCode As String

    strCode = Me.PurchaseOrdersOnFileSubItemSetup.Form.ID

    optSortBy_AfterUpdate

    Set rs = Me.PurchaseOrdersOnFileSub.Form.RecordsetClone
    rs.FindFirst "[ItemID] = " & Str(strCode)
    Me.PurchaseOrdersOnFileSub.Form.Bookmark = rs.Bookmark
    rs.Close
    Set rs = Nothing




'OnError GoTo Err_cmdRefresh_Click
'
'
'    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70
'
'Exit_cmdRefresh_Click:
'    Exit Sub
'
'Err_cmdRefresh_Click:
'    MsgBox Err.Description
'    Resume Exit_cmdRefresh_Click
'
End Sub

Private Sub Show_Detail_Click()
If gcfHandleErrors Then On Error GoTo Err_Show_Detail_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "AdjustReasons"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Show_Detail_Click:
    Exit Sub

Err_Show_Detail_Click:
    MsgBox Err.Description
    Resume Exit_Show_Detail_Click

End Sub
Private Sub Command125_Click()
If gcfHandleErrors Then On Error GoTo Err_Command125_Click

    AppendUPCItemsWithItemID

Exit_Command125_Click:
    Exit Sub

Err_Command125_Click:
    MsgBox Err.Description
    Resume Exit_Command125_Click

End Sub
