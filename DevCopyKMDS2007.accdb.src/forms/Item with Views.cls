Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Dim gbRunMailChanges As Boolean
Dim mbTabCycle As Boolean
Dim mnLastComboBox As Integer
Dim mbPriceGroupUpdate As Boolean



Private Sub AverageDailySales_DblClick(Cancel As Integer)
    DoCmd.OpenForm "PurchaseOrderPercents", , , , , acDialog

    Me.Refresh


End Sub

Private Sub cmbItemByCode_AfterUpdate()
    ' Find the record that matches the control.
Dim rs As Object
Dim nsaveID As Double '07/17/2021
    If gcfHandleErrors Then On Error GoTo ErrLog

    nsaveID = cmbItemByCode
    Me.OrderBy = "Item.ID"
    Me.OrderByOn = True
    Me.Caption = "Items by Code"

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(nsaveID)
    Me.Bookmark = rs.Bookmark

    mnLastComboBox = 1
    TabCombos
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmbItemByCode_AfterUpdate"
    Resume ExitSub
End Sub

Private Sub cmbItemByDesc_AfterUpdate()
    Dim rs As Object
    Dim nsaveID As Double '07/17/2021

If gcfHandleErrors Then On Error GoTo ErrLog
    nsaveID = cmbItemByDesc
    Me.OrderBy = "Item.Description,Item.ID"
    Me.OrderByOn = True
    Me.Caption = "Items by Description by Code"

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(nsaveID)
    Me.Bookmark = rs.Bookmark
    mnLastComboBox = 2
    TabCombos
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmbItemByDesc_AfterUpdate"
    Resume ExitSub
End Sub

Private Sub cmbItemByUPC_AfterUpdate()
    Dim rs As Object

If gcfHandleErrors Then On Error GoTo ErrLog
    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(Me![cmbItemByUPC])
    Me.Bookmark = rs.Bookmark
    mnLastComboBox = 3
    TabCombos


ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmbItemByUPC_AfterUpdate"
    Resume ExitSub
End Sub



Private Sub cmbItemBySalesDept_AfterUpdate()
    Dim rs As Object
    Dim nsaveID As Double '07/17/2021

If gcfHandleErrors Then On Error GoTo ErrLog
    nsaveID = cmbItemBySalesDept
    Me.OrderBy = "SalesDepartmentID,Item.Description,Item.ID"
    Me.OrderByOn = True

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(nsaveID)
    Me.Bookmark = rs.Bookmark
    Me.Caption = "Items by Sales Department by Description by Code"

    mnLastComboBox = 4
    TabCombos

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmbItemBySalesDept_AfterUpdate"
    Resume ExitSub
End Sub

Private Sub cmbItemsByVendor_AfterUpdate()

    Dim rs As Object
    Dim nsaveID As Double '07/17/2021

    If gcfHandleErrors Then On Error GoTo ErrLog
    nsaveID = cmbItemsByVendor
    Me.OrderBy = "VendorID,Item.Description,Item.ID"
    Me.OrderByOn = True

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(nsaveID)
    Me.Bookmark = rs.Bookmark
    Me.Caption = "Items by Vendor by Description by Code"
    mnLastComboBox = 5
    TabCombos

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmbItemsByVendor_AfterUpdate"
    Resume ExitSub
End Sub

Private Sub cmbItemsByPriceGroup_AfterUpdate()
    Dim rs As Object
    Dim nsaveID As Double '07/17/2021

If gcfHandleErrors Then On Error GoTo ErrLog
    nsaveID = cmbItemsByPriceGroup
    Me.OrderBy = "PriceGroupID,Item.Description,Item.ID"
    Me.OrderByOn = True

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(nsaveID)
    Me.Bookmark = rs.Bookmark
    Me.Caption = "Items by Price Group by Description by Code"
    mnLastComboBox = 6
    TabCombos

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmbItemsByPriceGroup_AfterUpdate"
    Resume ExitSub
End Sub


Private Sub cmbItemByLocation_AfterUpdate()
    Dim rs As Object
    Dim nsaveID As Double '07/17/2021

If gcfHandleErrors Then On Error GoTo ErrLog
    nsaveID = cmbItemByLocation
    Me.OrderBy = "Location,Item.Description,Item.ID"
    Me.OrderByOn = True

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(nsaveID)
    Me.Bookmark = rs.Bookmark
    Me.Caption = "Items by Location by Description by Code"
    mnLastComboBox = 7
    TabCombos

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmbItemByLocation_AfterUpdate"
    Resume ExitSub
End Sub

Private Sub cmbItemBySubCategoryID_AfterUpdate()
    Dim rs As Object
    Dim nsaveID As Double '07/17/2021

If gcfHandleErrors Then On Error GoTo ErrLog
    nsaveID = cmbItemBySubCategoryID
    Me.OrderBy = "SubCategoryID,Item.Description,Item.ID"
    Me.OrderByOn = True

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(nsaveID)
    Me.Bookmark = rs.Bookmark
    Me.Caption = "Items by Category by Description by Code"
    mnLastComboBox = 8
    TabCombos

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmbItemBySubCategoryID_AfterUpdate"
    Resume ExitSub
End Sub


Private Sub cmdfrmAddNew_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    'DoCmd.GoToRecord , , acNewRec
    Dim rs As Object

    Dim stDocName As String
    Dim stLinkCriteria As String

    'load the new item screen
    stDocName = "ItemSetup"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, acFormAdd, acDialog
    'requery/refresh to include new records

    Me.Requery
    Me.Refresh

    'move the the last record
    Set rs = Me.Recordset.Clone
    rs.MoveLast
    Me.Bookmark = rs.Bookmark

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmAddNew_Click"
    Resume ExitSub
End Sub

Private Sub cmdfrmClose_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    DoCmd.Close

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmClose_Click"
    Resume ExitSub
End Sub

Private Sub cmdfrmDeleteRecord_Click()

If gcfHandleErrors Then On Error GoTo ErrLog
Dim bOKToDelete As Boolean
Dim rs As Recordset

        If MsgBox("This will delete the item." & vbCrLf & "Are you sure?", vbCritical + vbYesNo) = vbYes Then
            DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
            DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
        End If

'set focus to combo
'   cboID.Visible = True
'   cboID.SetFocus

ExitSub:
    Exit Sub

ErrLog:
    Select Case Err.Number
    Case 2501 'action was cancelled (delete failed)
    'Case 2501 'action was cancelled (delete failed)
    Case Else
        EMsgBox Me, "cmdfrmDeleteRecord_Click" & vbCrLf & Err.Number & " " & Err.Description
    End Select
    Resume ExitSub

End Sub

Private Sub cmdfrmSaveRecord_Click()
If gcfHandleErrors Then On Error GoTo ErrLog


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    cmdfrmAddNew.SetFocus
    cmdfrmUndo.Visible = False
    cmdfrmSaveRecord.Visible = False
    lblEditing.Visible = False

    'refresh combo box,
'   cboID.Requery
'   cboID = ID
'   cboID.Visible = True
'   cboID.SetFocus

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmSaveRecord_Click"
    Resume ExitSub
End Sub

Private Sub cmdfrmUndo_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
    cmdfrmAddNew.SetFocus
    cmdfrmUndo.Visible = False
    cmdfrmSaveRecord.Visible = False
    lblEditing.Visible = False

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmUndo_Click"
    Resume ExitSub
End Sub

Private Sub Command139_Click()

End Sub


Private Sub cmdPrintLabels_Click()
    CreateLabels
End Sub

Private Sub CreateLabels()
If gcfHandleErrors Then On Error GoTo ErrLog
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
        .CommandText = "qryItemLabelsUpdateLabels"
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
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , ID

Exit_Command144_Click:
    Exit Sub

Err_Command144_Click:
    MsgBox Err.Description
    Resume Exit_Command144_Click
End Sub

Private Sub Cost_DblClick(Cancel As Integer)
    TabStopSkip Cost

End Sub

Private Sub Cost_Exit(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog
    CalcGP
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Cost_Exit"
    Resume ExitSub
End Sub


Private Sub Deposit_DblClick(Cancel As Integer)
    TabStopSkip Deposit
End Sub



Private Sub Description_DblClick(Cancel As Integer)
    TabStopSkip Description
End Sub

Private Sub Form_AfterUpdate()
    'cmbItemByDesc.Requery

Dim nsaveID As Double '07/17/2021
Dim rs As Object


If gcfHandleErrors Then On Error GoTo ErrLog
    gbRunMailChanges = True
    lblEditing.Visible = False
    lblSaving.Visible = False
    Me.Repaint

    If mbPriceGroupUpdate Then
        PriceGroupUpdate PriceGroupID, Price, QtyPack, QtyPackPrice, Cost
        nsaveID = ID
        'to reflect the new prices on the 'other' items in the price group
        Me.Requery
        Set rs = Me.Recordset.Clone
        rs.FindFirst "[ID] = " & Str(nsaveID)
        Me.Bookmark = rs.Bookmark
    End If

If gcfHandleErrors Then On Error GoTo ErrLog
    Me![cmbItemByCode].Requery
    Me![cmbItemByDesc].Requery
    Me![cmbItemByUPC].Requery
    Me![cmbItemBySalesDept].Requery
    Me![cmbItemsByPriceGroup].Requery
    Me![cmbItemByLocation].Requery
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_AfterUpdate"
    Resume ExitSub
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Dim strSQL As String

    'added 12/27/2004
    'liquor items are len = 4 and 5
    Select Case Len(ID)
    Case 4
        'changed for "CRO's"
        'If VendorID <> 1 Then
        '    VendorID = 1
        'End If
        If SalesDepartmentID <> 1 Then
            SalesDepartmentID = 1
        End If
    Case 5
        'changed for "CRO's"
        'If VendorID <> 1 Then
        '    VendorID = 1
        'End If
        If SalesDepartmentID <> 1 Then
            SalesDepartmentID = 1
        End If
    Case Else
        If VendorID = 1 Then
            VendorID = 0
        End If
        If SalesDepartmentID = 1 Then
            SalesDepartmentID = 0
        End If
    End Select



    mbPriceGroupUpdate = False
    If OrderAt > OrderTo Then
        MsgBox "Order at can not be greater than Order to!"
        lblSaveRecord.Visible = False
        Cancel = True
        If OrderAt.Locked Then
        Else
            OrderAt.SetFocus
        End If
        Exit Sub
    End If
    If Not IsNull(PriceGroupID) And PriceGroupID > 0 And Price.Value <> Price.OldValue Then
        strSQL = "This will change every item in this Price Group (" & PriceGroupID.Column(1) & ")" & vbCrLf
        strSQL = strSQL & "to " & Format(Price, "Currency") & vbCrLf & vbCrLf
        strSQL = strSQL & "Are you sure?"
        If MsgBox(strSQL, vbYesNo) = vbNo Then
            lblSaveRecord.Visible = False
            Cancel = True
            Exit Sub
        Else
            mbPriceGroupUpdate = True
        End If

    End If
    If Me.Onhand.OldValue <> Me.Onhand Then
        DoCmd.OpenForm "AdjustReasons", acNormal, , , acFormEdit, acDialog

        AuditChange "Onhand", ID, Onhand.OldValue, Onhand.Value, Price
    End If

    If Me.Price.OldValue <> Me.Price Then
        'DoCmd.OpenForm "AdjustReasons", acNormal, , , acFormEdit, acDialog

        AuditPriceChange "Price", ID, Price.OldValue, Price.Value, Onhand
    End If

    lblSaving.Visible = True
    Me.Repaint
    ChangeDate = Now

End Sub

Private Sub Form_Current()
If gcfHandleErrors Then On Error GoTo ErrLog
    If Me.NewRecord Then
        'ID.Visible = True
        'cboID.Visible = False
    Else
        Me![cmbItemByCode] = ID
        Me![cmbItemByDesc] = ID
        Me![cmbItemByUPC] = ID
        Me![cmbItemBySalesDept] = ID
        Me![cmbItemsByPriceGroup] = ID
        Me![cmbItemByLocation] = ID
        Me![cmbItemBySubCategoryID] = ID
        Me![cmbItemsByVendor] = ID
        CalcGP

        'cboID.Visible = True
        'ID.Visible = False
        'cboID = ID
        cmdfrmDeleteRecord.Visible = True
        txtAvailOnTheShelf = [Onhand] + [TransferredToday] - [SoldTodayCr1] - [SoldTodayCr2] - [SoldTodayCr3] - [SoldTodayCr4] - Nz(OnReserve.Column(1), 0)
    End If

    cmdfrmAddNew.Visible = Not Me.NewRecord
    cmdfrmDeleteRecord.Visible = Not Me.NewRecord
    cmdfrmSaveRecord.Visible = False
    cmdfrmUndo.Visible = False
    lblEditing.Visible = False
'    txtRecordCount = Me.Recordset.RecordCount
'    txtAbsPos = Me.Recordset.AbsolutePosition
    'Forms("ItemChartSalesTotal").Requery
    PLUList.Locked = False
    TPR.Locked = False
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Current"
    Resume ExitSub
End Sub

Private Sub Form_Delete(Cancel As Integer)

'don't allow delete if there are onhands....
'move transactions to a catch all item to preseve sales history totals

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")


Dim nSaveCatchAllItemID As Long
Dim bInvolinesFound  As Boolean
Dim strSQL As String

    If gcfHandleErrors Then On Error GoTo ErrLog


    bInvolinesFound = False

    'are there any onhands?
    If Onhand <> 0 Then
        MsgBox "Onhands are not zero!", vbInformation, "Delete not allowed"
        Cancel = True
        Onhand.SetFocus
        Exit Sub
    End If

    '1. get the catch all item
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qrySalesDepartmentCatchAllItemsByID"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parSalesDepartmentID", adInteger, adParamInput)
        .Parameters("parSalesDepartmentID") = SalesDepartmentID
    End With

    With rstemp
        .Open comQDF, , adOpenDynamic, adLockOptimistic
        If .BOF And .EOF Then  'no catch all item set up for this department
            MsgBox "There needs to be a Catch All Item Setup for department #" & SalesDepartmentID, vbCritical
            MsgBox "Do this in Menu Option E-1", vbCritical
            Cancel = True
            Exit Sub
        Else
            .MoveFirst
            If ID = !CatchAllItemID Then 'trying to delete a catch all item
                MsgBox "Deleting this item is not allowed!"
                MsgBox "This is the Catch All Item # for SalesDepartment #" & SalesDepartmentID, vbCritical
                Cancel = True
                Exit Sub
            Else 'save the CatchAllItemID if history exists....
                nSaveCatchAllItemID = !CatchAllItemID
            End If
        End If
        .Close
    End With
    Set rstemp = Nothing
    Set rstemp = CreateObject("ADODB.Recordset")

    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    '2. is there is any history for this item
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryInvoLineCountByCode"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parCode", adInteger, adParamInput)
        .Parameters("parCode") = ID
    End With

    With rstemp
        .Open comQDF, , adOpenDynamic, adLockOptimistic
        If .BOF And .EOF Then 'no history found for this item
            'bOKToDelete = True
        Else  'Set flag to move history
            .MoveFirst
            strSQL = "Moving transaction history of this item to " & nSaveCatchAllItemID & vbCrLf
            strSQL = strSQL & !Count & " transactions to move"
            bInvolinesFound = True
        End If
        .Close
    End With
    Set rstemp = Nothing

    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    If bInvolinesFound Then 'move the history of this item to the catch all item for this department
        If MsgBox(strSQL, vbOKCancel) <> vbOK Then
            Cancel = True
            Exit Sub
        End If
        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryInvoLineSetCodeToCatchAllCode"
            .CommandType = adCmdStoredProc
            'ADO Numeric Datatypes are very particular
            .Parameters.Append .CreateParameter("parCode", adInteger, adParamInput)
            .Parameters.Append .CreateParameter("parCodeCatchAll", adInteger, adParamInput)
            .Parameters("parCode") = ID
            .Parameters("parCodeCatchAll") = nSaveCatchAllItemID
            .Execute
        End With
        Set comQDF = Nothing
    End If
    'if the code executed to here then the delete will happen!

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Delete"
    Cancel = True
    Resume ExitSub
End Sub



Private Sub Form_Dirty(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog
    lblEditing.Visible = True

    cmdfrmSaveRecord.Visible = True
    cmdfrmUndo.Visible = True
    lblEditing.Visible = True
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Dirty"
    Resume ExitSub
End Sub

Private Sub Form_Load()
If gcfHandleErrors Then On Error GoTo ErrLog
    Me.OrderByOn = False

    gbRunMailChanges = False
    cmbItemByCode.SetFocus

    SetBackColor Price, gbAllowChangePrice
    SetBackColor Onhand, gbAllowChangeOnhand
    SetBackColor Description, gbAllowChangeDescription

    SetBackColor Cost, gbAllowChangeCost
    SetBackColor OrderTo, gbAllowChangeOrderTo
    SetBackColor OrderAt, gbAllowChangeOrderAt

    Me.InsideWidth = boxScreen.Width + 500
    Me.InsideHeight = Detail.Height + Me.FormFooter.Height '+ 500


    Me![cmbItemByCode].BackColor = glnColorDropDown
    Me![cmbItemByDesc].BackColor = glnColorDropDown
    Me![cmbItemByUPC].BackColor = glnColorDropDown
    Me![cmbItemBySalesDept].BackColor = glnColorDropDown
    Me![cmbItemsByPriceGroup].BackColor = glnColorDropDown
    Me![cmbItemByLocation].BackColor = glnColorDropDown
    Me![cmbItemBySubCategoryID].BackColor = glnColorDropDown

    txtGP.BackColor = glnColorLockedForAll
    SoldTodayCr1.BackColor = glnColorLockedForAll
    SoldTodayCr2.BackColor = glnColorLockedForAll
    SoldTodayCr3.BackColor = glnColorLockedForAll
    SoldTodayCr4.BackColor = glnColorLockedForAll
    TransferredToday.BackColor = glnColorLockedForAll
    OnReserve.BackColor = glnColorLockedForAll
    OnOrder.BackColor = glnColorLockedForAll
    SetupDate.BackColor = glnColorLockedForAll
    ChangeDate.BackColor = glnColorLockedForAll
    DateLastSold.BackColor = glnColorLockedForAll

    'txtAvailOnhand.BackColor = glnColorTotals
    txtAvailSoldToday.BackColor = glnColorTotals
    txtAvailTransferred.BackColor = glnColorTotals
    txtAvailOnTheShelf.BackColor = glnColorTotals
    OnReserve.BackColor = glnColorTotals


    TabStopSkipInit (False)
    optTabSkip.Locked = False
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Load"
    Resume ExitSub
End Sub

Private Sub SetBackColor(myTextBox As TextBox, Allowed As Boolean)
    With myTextBox
        .Locked = Not Allowed
        .Enabled = Allowed

        If .Enabled Then
            .BackColor = vbWhite ' SetupDate.BackColor
        Else
            .BackColor = glnColorViewOnly
        End If
    End With

End Sub

Private Sub Form_Open(Cancel As Integer)
    'Cancel = FormLogin(glnCashierID, Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog
    If gbRunMailChanges Then
        MailChanges (True)
    End If
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Unload"
    Resume ExitSub
End Sub

Private Sub Location_DblClick(Cancel As Integer)
    TabStopSkip Location
End Sub

Private Sub Onhand_DblClick(Cancel As Integer)
    TabStopSkip Onhand
End Sub

Private Sub OnOrder_DblClick(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "ItemOnOrder"

    stLinkCriteria = "[ItemID]=" & Me![ID]
    DoCmd.OpenForm stDocName, acFormDS, , stLinkCriteria, acFormReadOnly, acDialog, Description

ExitSub:
    Exit Sub

ErrLog:
    MsgBox Err.Description
    Resume ExitSub
End Sub

Private Sub OnReserve_DblClick(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "ItemOnReserve"

    stLinkCriteria = "[ItemID]=" & Me![ID]
    DoCmd.OpenForm stDocName, acFormDS, , stLinkCriteria, acFormReadOnly, acDialog, Description

ExitSub:
    Exit Sub

ErrLog:
    MsgBox Err.Description
    Resume ExitSub
End Sub

Private Sub optTabSkip_Click()
    Dim strSQL As String
If gcfHandleErrors Then On Error GoTo ErrLog
    If optTabSkip Then
        strSQL = "Do you want to enable the skipped fields option?" & vbCrLf & vbCrLf
        strSQL = strSQL & "Yes = Sets all fields to skip (Bold Font)" & vbCrLf
        strSQL = strSQL & "       Double click with left mouse button on each field you want to stop at." & vbCrLf & vbCrLf
        strSQL = strSQL & "No  = Do Not Skip any fields"

        If MsgBox(strSQL, vbYesNo) = vbYes Then
            TabStopSkipInit (True)
        End If
    Else
        strSQL = "Do you want to DISABLE the skipped fields option?" & vbCrLf & vbCrLf
        strSQL = strSQL & "Yes = Sets all fields to NOT skip"
        If MsgBox(strSQL, vbYesNo) = vbYes Then
            TabStopSkipInit (False)
        End If

    End If

ExitSub:
    Exit Sub

ErrLog:
    Select Case Err.Number
    Case 438
        Resume Next
    Case Else
        MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "optTabSkip_Click"
        Resume ExitSub
    End Select
End Sub

Private Sub OrderAt_DblClick(Cancel As Integer)
    TabStopSkip OrderAt
End Sub

Private Sub OrderTo_DblClick(Cancel As Integer)
    TabStopSkip OrderTo
End Sub

Private Sub Price_DblClick(Cancel As Integer)
    TabStopSkip Price
End Sub

Private Sub Price_Exit(Cancel As Integer)
    CalcGP
End Sub

Private Sub CalcGP()
If gcfHandleErrors Then On Error GoTo ErrLog

        If gcfHandleErrors Then On Error GoTo ErrLog
        txtGP = IIf(Price <= 1, 1, (Price - Cost) / Price)

        'If not gcfHandleErrors then Debug.Print Price; txtGP
ExitSub:
    Exit Sub
ErrLog:
    txtGP = "*"
    Resume ExitSub
End Sub

Private Sub cmdNewItem_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdNewItem_Click
    Dim rs As Object

    Dim stDocName As String
    Dim stLinkCriteria As String

    'load the new item screen
    stDocName = "ItemSetup"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, acFormAdd, acDialog
    'requery/refresh to include new records

    Me.Requery
    Me.Refresh

    'move the the last record
    Set rs = Me.Recordset.Clone
    rs.MoveLast
    Me.Bookmark = rs.Bookmark


Exit_cmdNewItem_Click:
    Exit Sub

Err_cmdNewItem_Click:
    MsgBox Err.Description
    Resume Exit_cmdNewItem_Click

End Sub




Private Sub PriceGroupID_DblClick(Cancel As Integer)
    TabStopSkip PriceGroupID
End Sub

Private Sub PriceGroupID_NotInList(NewData As String, Response As Integer)
    If MsgBox(NewData & " not found; Would you like to set this up?", vbYesNo) = vbYes Then
        DoCmd.OpenForm "PriceGroup", , , , acFormAdd, acDialog, NewData
        PriceGroupID.Undo
        PriceGroupID.Requery
        PriceGroupID.Text = NewData
        Response = acDataErrAdded
    Else
        Response = acDataErrContinue
        PriceGroupID.Undo

    End If
End Sub

Private Sub QtyPack_DblClick(Cancel As Integer)
    TabStopSkip QtyPack

End Sub

Private Sub QtyPackPrice_DblClick(Cancel As Integer)
    TabStopSkip QtyPackPrice

End Sub

Private Sub SalesDepartmentID_DblClick(Cancel As Integer)
   TabStopSkip SalesDepartmentID

End Sub

Private Sub SalesDepartmentID_NotInList(NewData As String, Response As Integer)
    If MsgBox(NewData & " not found; Would you like to set this item up?", vbYesNo) = vbYes Then
        DoCmd.OpenForm "SalesGroup", , , , acFormAdd, acDialog, NewData
        SalesDepartmentID.Undo
        SalesDepartmentID.Requery
        SalesDepartmentID.Text = NewData
        Response = acDataErrAdded
    Else
        Response = acDataErrContinue
        SalesDepartmentID.Undo

    End If
End Sub

Private Sub SizeID_DblClick(Cancel As Integer)
    TabStopSkip SizeID
End Sub

Private Sub SizeID_NotInList(NewData As String, Response As Integer)
    If MsgBox(NewData & " not found; Would you like to set this item up?", vbYesNo) = vbYes Then
        DoCmd.OpenForm "Sizes", , , , acFormAdd, acDialog, NewData
        SizeID.Undo
        SizeID.Requery
        SizeID.Text = NewData
        Response = acDataErrAdded
    Else
        Response = acDataErrContinue
        SizeID.Undo

    End If
End Sub

Private Sub StockNumber_DblClick(Cancel As Integer)
    TabStopSkip StockNumber
End Sub

Private Sub SubCategoryID_DblClick(Cancel As Integer)
    TabStopSkip SubCategoryID
End Sub

Private Sub SubPack_DblClick(Cancel As Integer)
    TabStopSkip SubPack

End Sub


Private Sub Text154_DblClick(Cancel As Integer)
    AverageDailySales_DblClick False
End Sub

Private Sub UnitsPerCase_DblClick(Cancel As Integer)
    TabStopSkip UnitsPerCase

End Sub

Private Sub UnitsPerOrder_DblClick(Cancel As Integer)
    TabStopSkip UnitsPerOrder

End Sub

Private Sub UnitsPerPack_DblClick(Cancel As Integer)
    TabStopSkip UnitsPerPack

End Sub

Private Sub UPCItem_Exit(Cancel As Integer)
    mbTabCycle = True



End Sub

Private Sub VendorID_DblClick(Cancel As Integer)
    TabStopSkip VendorID
End Sub

Private Sub VendorID_Exit(Cancel As Integer)
 If gcfHandleErrors Then On Error GoTo ErrLog
   If SalesDepartmentID = 1 And VendorID <> glnOLCCVendor Then
        'VendorID = glnOLCCVendor
    End If
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "VendorID_Exit"
    Resume ExitSub
End Sub


'Private Sub Form_Load()
'    Form.InsideHeight = boxScreen.Height + 500 + cmdGoToFirstRecord.Height
'    Form.InsideWidth = boxScreen.Width + 500
'
'End Sub
'

Private Sub cmdGoToNew_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdGoToNew_Click


    DoCmd.GoToRecord , , acNewRec
    Description.SetFocus
Exit_cmdGoToNew_Click:
    Exit Sub

Err_cmdGoToNew_Click:
    EMsgBox Me, "cmdGoToNew_Click"
    Resume Exit_cmdGoToNew_Click

End Sub
Private Sub cmdDeleteRecord_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdDeleteRecord_Click


    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
    Description.SetFocus

Exit_cmdDeleteRecord_Click:
    Exit Sub

Err_cmdDeleteRecord_Click:
    Select Case Err.Number
    Case 2501
    Case Else
        EMsgBox Me, "cmdDeleteRecord_Click"
    End Select
    Resume Exit_cmdDeleteRecord_Click

End Sub

Private Sub cmdGoToLastRecord_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdGoToLastRecord_Click


    DoCmd.GoToRecord , , acLast
    Description.SetFocus

Exit_cmdGoToLastRecord_Click:
    Exit Sub

Err_cmdGoToLastRecord_Click:
    EMsgBox Me, "cmdGoToLastRecord_Click"
    Resume Exit_cmdGoToLastRecord_Click

End Sub
Private Sub cmdGoToNext_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdGoToNext_Click


    DoCmd.GoToRecord , , acNext
    Description.SetFocus
Exit_cmdGoToNext_Click:
    Exit Sub

Err_cmdGoToNext_Click:
    Select Case Err.Number
    Case 2105
        Beep
    Case 2110 ' description can't receive focus!
        Beep
    Case Else
        EMsgBox Me, "cmdGoToNext_Click"
    End Select
    Resume Exit_cmdGoToNext_Click

End Sub
Private Sub cmdGoToPreviousRecord_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdGoToPreviousRecord_Click


    DoCmd.GoToRecord , , acPrevious
    Description.SetFocus

Exit_cmdGoToPreviousRecord_Click:
    Exit Sub

Err_cmdGoToPreviousRecord_Click:
    Select Case Err.Number
    Case 2105
        Beep
    Case Else
        EMsgBox Me, "cmdGoToPreviousRecord_Click"
    End Select
    Resume Exit_cmdGoToPreviousRecord_Click

End Sub

Private Sub cmdGoToFirstRecord_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdGoToFirstRecord_Click


    DoCmd.GoToRecord , , acFirst
    Description.SetFocus

Exit_cmdGoToFirstRecord_Click:
    Exit Sub

Err_cmdGoToFirstRecord_Click:
    EMsgBox Me, "cmdGoToFirstRecord_Click"
    Resume Exit_cmdGoToFirstRecord_Click

End Sub
Private Sub cmdSaveRecord_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdSaveRecord_Click


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70

Exit_cmdSaveRecord_Click:
    Exit Sub

Err_cmdSaveRecord_Click:
    MsgBox Err.Description
    Resume Exit_cmdSaveRecord_Click

End Sub
Private Sub cmdUndo_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdUndo_Click


    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70

Exit_cmdUndo_Click:
    Exit Sub

Err_cmdUndo_Click:
    MsgBox Err.Description
    Resume Exit_cmdUndo_Click

End Sub

Public Sub TabStopSkip(myControl As Control)
    If optTabSkip Then
        myControl.TabStop = Not myControl.TabStop
        myControl.FontBold = Not myControl.TabStop
'
'        If myControl.TabStop Then
'            myControl.FontWeight = Normal
'        Else
'            myControl.FontWeight = vbbold
'        End If
    End If

End Sub

Public Sub TabStopSkipInit(SetToOn As Boolean)

    Dim myControl As Control
    If gcfHandleErrors Then On Error GoTo ErrLog
    For Each myControl In Me.Controls
        'If not gcfHandleErrors then Debug.Print myControl.Name
        If myControl.Tag = "TabStopSkip" Then
            myControl.TabStop = Not SetToOn
            myControl.FontBold = SetToOn
        End If
        If Not IsNull(myControl.BackColor) And myControl.BackColor = glnColorLockedForAll Then
            'If not gcfHandleErrors then Debug.Print "Locking "; myControl.BackColor; myControl.Name
            myControl.Locked = True

            'myControl.Enabled = False
        End If
    Next myControl
    optTabSkip.Locked = False
    UPCItem.Locked = False
    UPCItem.TabStop = Not SetToOn

ExitSub:
    Exit Sub

ErrLog:
    Select Case Err.Number
    Case 438
        Resume Next
    Case Else
        MsgBox myControl.Name & vbCrLf & "Error #" & Err.Number & " " & Err.Description, vbInformation, "TabStopSkipInit"
        Resume ExitSub
    End Select
End Sub

Private Sub cmdOnOrder_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdOnOrder_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "ItemOnOrder"

    stLinkCriteria = "[ItemID]=" & Me![ID]
    DoCmd.OpenForm stDocName, acFormDS, , stLinkCriteria

Exit_cmdOnOrder_Click:
    Exit Sub

Err_cmdOnOrder_Click:
    MsgBox Err.Description
    Resume Exit_cmdOnOrder_Click

End Sub

Private Sub VendorID_NotInList(NewData As String, Response As Integer)
    If MsgBox(NewData & " not found; Would you like to set this up?", vbYesNo) = vbYes Then
        DoCmd.OpenForm "Vendors", , , , acFormAdd, acDialog, NewData
        VendorID.Undo
        VendorID.Requery
        VendorID.Text = NewData
        Response = acDataErrAdded
    Else
        Response = acDataErrContinue
        VendorID.Undo

    End If
End Sub

Private Sub TabCombos()
    cmbItemByCode.TabStop = False
    cmbItemByDesc.TabStop = False
    cmbItemByUPC.TabStop = False
    cmbItemBySalesDept.TabStop = False
    cmbItemsByVendor.TabStop = False
    cmbItemsByPriceGroup.TabStop = False
    cmbItemByLocation.TabStop = False
    cmbItemBySubCategoryID.TabStop = False

    Select Case mnLastComboBox
    Case 1
        cmbItemByCode.TabStop = True
    Case 2
        cmbItemByDesc.TabStop = True
    Case 3
        cmbItemByUPC.TabStop = True
    Case 4
        cmbItemBySalesDept.TabStop = True
    Case 5
        cmbItemsByVendor.TabStop = True
    Case 6
        cmbItemsByPriceGroup.TabStop = True
    Case 7
        cmbItemByLocation.TabStop = True
    Case 8
        cmbItemBySubCategoryID.TabStop = True
    End Select
End Sub
Private Sub cmdShowChart_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdShowChart_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "ItemWithSalesTotalChart"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, Me.Name
Exit_cmdShowChart_Click:
    Exit Sub

Err_cmdShowChart_Click:
    MsgBox Err.Description
    Resume Exit_cmdShowChart_Click

End Sub
Private Sub Command138_Click()
If gcfHandleErrors Then On Error GoTo Err_Command138_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "ItemWithSalesTotalChart"

    stLinkCriteria = "[ID]=" & Me![ID]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command138_Click:
    Exit Sub

Err_Command138_Click:
    MsgBox Err.Description
    Resume Exit_Command138_Click

End Sub

Private Sub Command143_Click()
If gcfHandleErrors Then On Error GoTo Err_Command143_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "PrintInventoryReport"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command143_Click:
    Exit Sub

Err_Command143_Click:
    MsgBox Err.Description
    Resume Exit_Command143_Click

End Sub
Private Sub Command144_Click()


End Sub
