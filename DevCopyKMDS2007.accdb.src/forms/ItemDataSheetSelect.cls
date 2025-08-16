Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim nColumn As Integer
Dim nSaveColumn As Integer
Dim nSortBy As Integer

Private Sub chkCategory_Click()
    CC chkSubCategoryID, txtSubCategoryID
End Sub

Private Sub chkChangeDate_click()
    CC chkChangeDate, txtChangeDate
End Sub

Private Sub chkCost_Click()
    CC chkCost, txtCost
End Sub

Private Sub chkDateLastSold_Click()
   CC chkDateLastSold, txtDateLastSold
End Sub

Private Sub chkDeposit_click()
    CC chkDeposit, txtDeposit
End Sub

Private Sub chkDescription_Click()
    CC chkDescription, txtDescription

End Sub

Private Sub chkID_Click()
    CC chkID, txtID
End Sub

Private Sub chkLocation_click()
    CC chkLocation, txtLocation
End Sub

Private Sub chkOnhand_click()
    CC chkOnhand, txtOnhand
End Sub

Private Sub chkOrderAt_click()
    CC chkOrderAt, txtOrderAt
End Sub

Private Sub chkOrderTo_click()
    CC chkOrderTo, txtOrderTo
End Sub

Private Sub chkPLUList_click()
    CC chkPLUList, txtPLUList
End Sub

Private Sub chkPrice_click()
    CC chkPrice, txtPrice
End Sub



Private Sub chkPriceGroupID_Click()
    CC chkPriceGroupID, txtPriceGroup

End Sub

Private Sub chkQtyPack_click()
    CC chkQtyPack, txtQtyPack
End Sub

Private Sub chkSalesDepartmentID_click()
    CC chkSalesDepartmentID, txtSalesDpt
End Sub

Private Sub chkSetupDate_click()
    CC chkSetupDate, txtSetupDate
End Sub

Private Sub chkSizeID_click()
    CC chkSizeID, txtSizeID
End Sub

Private Sub chkStockNumber_click()
    CC chkStockNumber, txtStockNumber
End Sub

Private Sub chkSubCategoryID_Click()
    CC chkSubCategoryID, txtSubCategoryID
End Sub

Private Sub chkSubPack_click()
    CC chkSubPack, txtSubPack
End Sub

Private Sub chkUnitsPerCase_click()
    CC chkUnitsPerCase, txtUnitsPerCase
End Sub

Private Sub chkUnitsPerOrder_click()
    CC chkUnitsPerOrder, txtUnitsPerOrder
End Sub

Private Sub chkUnitsPerPack_click()
    CC chkUnitsPerPack, txtUnitsPerPack
End Sub

Private Sub chkVendorID_click()
    CC chkVendorID, txtVendor
End Sub

Private Sub cmbmyDescription_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[myID] = " & Str(Nz(Me![cmbmyDescription], 0))
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
End Sub

Private Sub cmdfrmAddNew_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    DoCmd.GoToRecord , , acNewRec

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmAddNew_Click"
    Resume ExitSub
End Sub

Private Sub cmdfrmClose_Click()
If gcfHandleErrors Then On Error GoTo ErrLog
    DoCmd.Close acForm, "ItemDataSheet"
    DoCmd.Close

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmDeleteRecord_Click"
    Resume ExitSub
End Sub

Private Sub cmdfrmDeleteRecord_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
    'set focus to combo
'   cboID.Visible = True
'   cboID.SetFocus

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmDeleteRecord_Click"
    Resume ExitSub

End Sub

Private Sub cmdfrmSaveRecord_Click()
If gcfHandleErrors Then On Error GoTo ErrLog


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    cmdfrmAddNew.SetFocus
    cmdfrmUndo.Visible = False
    cmdfrmSaveRecord.Visible = False

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

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmUndo_Click"
    Resume ExitSub
End Sub




Private Sub Form_AfterUpdate()
    cmbmyDescription.Requery
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Dim nResp As Integer
    nResp = MsgBox("Do you want to save your changes?", vbYesNoCancel, "Criteria has changed...")
    Select Case nResp
    Case vbYes
        'check to make sure myDescription has a value
        If IsNull(myDescription) Then
            MsgBox "Please type in a Name for this data sheet!"
            myDescription.SetFocus
            Cancel = True
            Exit Sub
        Else
            'do nothing; allow save to continue
        End If

    Case vbNo
        'undo changes
        'Cancel = True
        Me.Undo
    Case vbCancel
        Cancel = True
    End Select
End Sub

Private Sub Form_Current()
    nColumn = 0
    nSaveColumn = 0
    nSortBy = 0
    If gcfHandleErrors Then On Error GoTo ErrLog
    [Forms]![itemdatasheet]![ItemDataSheetSub].Form.Filter = ""
    SetCheckBoxes
    ShowColumns
    SetFilter
    cmbmyDescription = myID
    With [Forms]![itemdatasheet]
        .Caption = Nz(myDescription, "Not Set")
        .txtReportTitle = myDescription
        .Move Me.InsideWidth, 0
    End With

    cmdfrmUndo.Visible = False
    cmdfrmSaveRecord.Visible = False
    cmdfrmDeleteRecord.Visible = Not Me.NewRecord
    cmdfrmAddNew.Visible = Not Me.NewRecord
ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "Form_Current"
    Resume ExitSub
End Sub

Private Sub Form_Dirty(Cancel As Integer)
    cmdfrmSaveRecord.Visible = True
    cmdfrmUndo.Visible = True
End Sub

Private Sub Form_Load()
    If gcfHandleErrors Then On Error GoTo ErrLog

    cmdClearAll_click
    DoCmd.OpenForm "ItemDataSheet", acNormal
    Me.OrderByOn = True
    Me.OrderBy = "myDescription"
    Me.Move 0, 0
    Me.InsideWidth = boxScreen.Width + 500
    Me.InsideHeight = Me.FormHeader.Height + Me.Detail.Height + Me.FormFooter.Height
ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "Form_Load"
    Resume ExitSub
End Sub

Private Sub txtCategorySB_Click()

    csb txtSubCategoryIDSB
End Sub

Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)
End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub


Private Sub Form_Resize()
'    If Me.windowstate = 1 Then

    If gcfHandleErrors Then On Error Resume Next 'GoTo ErrLog
'    Forms!ItemDataSheet.Move 0, 0
ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "Form_Resize"
    Resume ExitSub
End Sub

Private Sub txtChangeDateSB_Click()

    csb txtChangeDateSB
End Sub

Private Sub txtCostSB_Click()

    csb txtCostSB
End Sub

Private Sub txtDateLastSoldSB_Click()

    csb txtDateLastSoldSB
End Sub

Private Sub txtDepositSB_Click()

    csb txtDepositSB
End Sub

Private Sub txtDescription_Click()
    csb txtDescription
End Sub

Private Sub txtDescriptionSB_Click()
    csb txtDescriptionSB

End Sub


Private Sub txtFilter_BeforeUpdate(Cancel As Integer)
    SetFilter

End Sub

Private Sub txtID_BeforeUpdate(Cancel As Integer)
    csb txtID
End Sub

Private Sub txtIDSB_Click()
    csb txtIDSB
End Sub

Private Sub txtLocationSB_Click()

    csb txtLocationSB
End Sub

Private Sub txtMinSB_Click()

    csb txtOrderAtSB
End Sub



Private Sub txtOnhandSB_Click()

    csb txtOnhandSB
End Sub

Private Sub txtOrderAtSB_Click()
    csb txtOrderAtSB
End Sub

Private Sub txtOrderToSB_Click()

    csb txtOrderToSB
End Sub

Private Sub txtPLUListSB_Click()

    csb txtPLUListSB
End Sub

Private Sub txtPriceGroupSB_Click()

    csb txtPriceGroupSB
End Sub

Private Sub txtPriceSB_Click()

    csb txtPriceSB
End Sub

Private Sub txtQtyPackSB_Click()

    csb txtQtyPackSB
End Sub

Private Sub txtSalesDptSB_Click()

    csb txtSalesDptSB
End Sub

Private Sub txtSetupDateSB_Click()

    csb txtSetupDateSB
End Sub


Private Sub txtSizeIDSB_Click()
    csb txtSizeIDSB

End Sub

Private Sub txtStockNumberSB_Click()
    csb txtStockNumberSB
End Sub

Private Sub txtSubCategoryIDSB_Click()
    csb txtSubCategoryIDSB

End Sub

Private Sub txtSubPackSB_Click()

    csb txtSubPackSB
End Sub

Private Sub txtUnitsPerCaseSB_Click()
    csb txtUnitsPerCaseSB
End Sub

Private Sub txtUnitsPerOrderSB_Click()

    csb txtUnitsPerOrderSB
End Sub

Private Sub txtUnitsPerPackSB_Click()
    csb txtUnitsPerPackSB
End Sub

Private Sub txtVendorSB_Click()
    csb txtVendorSB
End Sub


Private Function csb(ctlTextBox As TextBox) As Byte
    Dim ctl As Control
    Form_Dirty 0

    If IsNull(ctlTextBox) Or ctlTextBox = "" Or ctlTextBox = 0 Then

        nSortBy = nSortBy + 1
        If nSortBy < 1 Then nSortBy = 1
        ctlTextBox = nSortBy
    Else
        nSortBy = nSortBy - 1
        ctlTextBox = 0
    End If

    ShowColumns

End Function

Private Sub ShowColumns()
Dim ctl As Control
    Dim strSortBy As String
    Dim strSortBy1 As String
    Dim strSortBy2 As String
    Dim strSortBy3 As String
If gcfHandleErrors Then On Error GoTo ErrLog
    With [Forms]![itemdatasheet]![ItemDataSheetSub].[Form]

        ![ID].ColumnHidden = Not chkID
        ![Description].ColumnHidden = Not chkDescription
        ![SizeID].ColumnHidden = Not chkSizeID
        ![VendorID].ColumnHidden = Not chkVendorID
        ![PriceGroupID].ColumnHidden = Not chkPriceGroupID
        ![SalesDepartmentID].ColumnHidden = Not chkSalesDepartmentID
        ![SubCategoryID].ColumnHidden = Not chkSubCategoryID
        ![Onhand].ColumnHidden = Not chkOnhand
        ![OrderAt].ColumnHidden = Not chkOrderAt
        ![OrderTo].ColumnHidden = Not chkOrderTo
        ![Price].ColumnHidden = Not chkPrice
        ![Cost].ColumnHidden = Not chkCost
        ![Deposit].ColumnHidden = Not chkDeposit
        ![UnitsPerCase].ColumnHidden = Not chkUnitsPerCase
        ![UnitsPerOrder].ColumnHidden = Not chkUnitsPerOrder
        ![UnitsPerPack].ColumnHidden = Not chkUnitsPerPack
        ![SubPack].ColumnHidden = Not chkSubPack
        ![Location].ColumnHidden = Not chkLocation
        ![QtyPack].ColumnHidden = Not chkQtyPack
        ![QtyPackPrice].ColumnHidden = Not chkQtyPack
        ![StockNumber].ColumnHidden = Not chkStockNumber
        ![PLUList].ColumnHidden = Not chkPLUList
        ![DateLastSold].ColumnHidden = Not chkDateLastSold
        ![SetupDate].ColumnHidden = Not chkSetupDate
        ![ChangeDate].ColumnHidden = Not chkChangeDate


        If txtID > 0 Then ![ID].ColumnOrder = txtID
        If txtDescription > 0 Then ![Description].ColumnOrder = txtDescription
        If txtSizeID > 0 Then ![SizeID].ColumnOrder = txtSizeID
        If txtVendor > 0 Then ![VendorID].ColumnOrder = txtVendor
        If txtPriceGroup > 0 Then ![PriceGroupID].ColumnOrder = txtPriceGroup
        If txtSalesDpt > 0 Then ![SalesDepartmentID].ColumnOrder = txtSalesDpt
        If txtSubCategoryID > 0 Then ![SubCategoryID].ColumnOrder = txtSubCategoryID
        If txtOnhand > 0 Then ![Onhand].ColumnOrder = txtOnhand
        If txtOrderAt > 0 Then ![OrderAt].ColumnOrder = txtOrderAt
        If txtOrderAt > 0 Then ![OrderTo].ColumnOrder = txtOrderTo
        If txtPrice > 0 Then ![Price].ColumnOrder = txtPrice
        If txtCost > 0 Then ![Cost].ColumnOrder = txtCost
        If txtDeposit > 0 Then ![Deposit].ColumnOrder = txtDeposit
        If txtUnitsPerCase > 0 Then ![UnitsPerCase].ColumnOrder = txtUnitsPerCase
        If txtUnitsPerOrder > 0 Then ![UnitsPerOrder].ColumnOrder = txtUnitsPerOrder
        If txtUnitsPerPack > 0 Then ![UnitsPerPack].ColumnOrder = txtUnitsPerPack
        If txtSubPack > 0 Then ![SubPack].ColumnOrder = txtSubPack
        If txtLocation > 0 Then ![Location].ColumnOrder = txtLocation
        If txtQtyPack > 0 Then ![QtyPack].ColumnOrder = txtQtyPack
        If txtQtyPack > 0 Then ![QtyPackPrice].ColumnOrder = txtQtyPack
        If txtStockNumber > 0 Then ![StockNumber].ColumnOrder = txtStockNumber
        If txtPLUList > 0 Then ![PLUList].ColumnOrder = txtPLUList
        If txtDateLastSold > 0 Then ![DateLastSold].ColumnOrder = txtDateLastSold
        If txtSetupDate > 0 Then ![SetupDate].ColumnOrder = txtSetupDate
        If txtChangeDate > 0 Then ![ChangeDate].ColumnOrder = txtChangeDate

        ![ID].ColumnWidth = -2
        ![Description].ColumnWidth = -2
        ![SizeID].ColumnWidth = -2
        ![VendorID].ColumnWidth = -2
        ![PriceGroupID].ColumnWidth = -2
        ![SalesDepartmentID].ColumnWidth = -2
        ![SubCategoryID].ColumnWidth = -2
        ![Onhand].ColumnWidth = -2
        ![OrderAt].ColumnWidth = -2
        ![OrderTo].ColumnWidth = -2
        ![Price].ColumnWidth = -2
        ![Cost].ColumnWidth = -2
        ![Deposit].ColumnWidth = -2
        ![UnitsPerCase].ColumnWidth = -2
        ![UnitsPerOrder].ColumnWidth = -2
        ![UnitsPerPack].ColumnWidth = -2
        ![SubPack].ColumnWidth = -2
        ![Location].ColumnWidth = -2
        ![QtyPack].ColumnWidth = -2
        ![QtyPackPrice].ColumnWidth = -2
        ![StockNumber].ColumnWidth = -2
        ![PLUList].ColumnWidth = -2
        ![DateLastSold].ColumnWidth = -2
        ![SetupDate].ColumnWidth = -2
        ![ChangeDate].ColumnWidth = -2



    End With

    For Each ctl In Me.Detail.Controls
        Select Case ctl.ControlType
        Case acTextBox
            If Right(ctl.Name, 2) = "SB" Then

                If ctl.Value = 1 Then
                    strSortBy1 = ctl.Tag
                End If
                If ctl.Value = 2 Then strSortBy2 = ctl.Tag
                If ctl.Value = 3 Then strSortBy3 = ctl.Tag
            End If

        End Select
    Next ctl

    strSortBy = strSortBy1
    If (strSortBy2 <> "") Then strSortBy = strSortBy & "," & strSortBy2
    If (strSortBy3 <> "") Then strSortBy = strSortBy & "," & strSortBy3

    [Forms]![itemdatasheet]![ItemDataSheetSub].Form.OrderBy = strSortBy
    [Forms]![itemdatasheet]![ItemDataSheetSub].Form.OrderByOn = True
    SetFilter
ExitSub:
    Exit Sub

ErrLog:
    Select Case Err.Number
    Case 2450

        DoCmd.OpenForm "ItemDataSheet", acNormal
    Case Else

        MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "ShowColumns"
    End Select

        Resume ExitSub
End Sub

Private Sub cmdClearAll_click()
    Dim ctl As Control
    For Each ctl In Me.Detail.Controls
        Select Case ctl.ControlType
        Case acTextBox
            'CTL.Value = 0
'            If Right(ctl.Name, 2) = "SB" Then
'                ctl.Value = 1
'            Else
'                ctl.Value = 0
'            End If
        Case acCheckBox
            ctl.Value = 0
        End Select
    Next ctl
    nColumn = 0
    nSortBy = 0
    'txtColumns = nColumn
    'txtOrderBy = ""
    'If Me.Dirty Then Me.Dirty = False
End Sub

Private Sub RemoveColumn(ColumnStart As Integer)
    Dim ctl As Control
    For Each ctl In Me.Controls
        Select Case ctl.ControlType
        Case acTextBox
            If IsNumeric(ctl) And ctl > ColumnStart Then
                ctl = CInt(ctl) - 1
            End If
        End Select
    Next
End Sub

Private Function CC(ctlCheckBox As CheckBox, ctlTextBox As TextBox) As Byte
    Form_Dirty 0
    If Not ctlCheckBox Then
        nColumn = nColumn - 1
        RemoveColumn (CInt(ctlTextBox))
        ctlTextBox = 0
        'cc = 0
    Else
        nColumn = nColumn + 1
        'cc = nColumn
        If nColumn < 1 Then nColumn = 1
        ctlTextBox = nColumn
    End If
    ShowColumns

End Function


Public Sub SetCheckBoxes()
    Dim ctl As Control
    For Each ctl In Me.Detail.Controls
        Select Case ctl.ControlType
        Case acCheckBox
            ctl = 0

        End Select
    Next
    nColumn = 0

    If txtID > 0 Then
        chkID = True
        nColumn = nColumn + 1
    End If

    If txtDescription > 0 Then
        chkDescription = True
        nColumn = nColumn + 1
    End If

    If txtSizeID > 0 Then
        chkSizeID = True
        nColumn = nColumn + 1
    End If

    If txtVendor > 0 Then
        chkVendorID = True
        nColumn = nColumn + 1
    End If

    If txtPriceGroup > 0 Then
        chkPriceGroupID = True
        nColumn = nColumn + 1
    End If

    If txtSalesDpt > 0 Then
        chkSalesDepartmentID = True
        nColumn = nColumn + 1
    End If

    If txtSubCategoryID > 0 Then
        chkSubCategoryID = True
        nColumn = nColumn + 1
    End If

    If txtOnhand > 0 Then
        chkOnhand = True
        nColumn = nColumn + 1
    End If

    If txtOrderAt > 0 Then
        chkOrderAt = True
        nColumn = nColumn + 1
    End If

    If txtOrderTo > 0 Then
        chkOrderTo = True
        nColumn = nColumn + 1
    End If

    If txtPrice > 0 Then
        chkPrice = True
        nColumn = nColumn + 1
    End If

    If txtCost > 0 Then
        chkCost = True
        nColumn = nColumn + 1
    End If

    If txtDeposit > 0 Then
        chkDeposit = True
        nColumn = nColumn + 1
    End If

    If txtUnitsPerCase > 0 Then
        chkUnitsPerCase = True
        nColumn = nColumn + 1
    End If

    If txtUnitsPerOrder > 0 Then
        chkUnitsPerOrder = True
        nColumn = nColumn + 1
    End If

    If txtUnitsPerPack > 0 Then
        chkUnitsPerPack = True
        nColumn = nColumn + 1
    End If

    If txtSubPack > 0 Then
        chkSubPack = True
        nColumn = nColumn + 1
    End If

    If txtLocation > 0 Then
        chkLocation = True
        nColumn = nColumn + 1
    End If

    If txtQtyPack > 0 Then
        chkQtyPack = True
        nColumn = nColumn + 1
    End If

    If txtStockNumber > 0 Then
        chkStockNumber = True
        nColumn = nColumn + 1
    End If

    If txtPLUList > 0 Then
        chkPLUList = True
        nColumn = nColumn + 1
    End If

    If txtDateLastSold > 0 Then
        chkDateLastSold = True
        nColumn = nColumn + 1
    End If

    If txtSetupDate > 0 Then
        chkSetupDate = True
        nColumn = nColumn + 1
    End If

    If txtChangeDate > 0 Then
        chkChangeDate = True
        nColumn = nColumn + 1
    End If

'
End Sub

Private Sub SetFilter()
    [Forms]![itemdatasheet]![ItemDataSheetSub].Form.Filter = Nz(Me.txtFilter, "")
    [Forms]![itemdatasheet]![ItemDataSheetSub].Form.FilterOn = True
End Sub
