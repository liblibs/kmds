Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub Cost_AfterUpdate()
'    If MsgBox("Do you want to change the Cost of this item?", vbYesNo) = vbYes Then
'        MsgBox ("The Cost must be changed in the Item Screen!")
'        EditItem
'    End If
End Sub

Private Sub ExtCost_DblClick(Cancel As Integer)
Dim strResp As String
    If Me.QuantityShipped = 0 Then
        MsgBox "Please put in quantity first!"
        Me.QuantityShipped.SetFocus
        Cancel = True
    Else
        strResp = InputBox("Extended Cost")
        If IsNumeric(strResp) Then
            Me.Cost = strResp / Me.QuantityShipped
        Else
            MsgBox "Please input a number!"
        End If
    End If
End Sub

Private Sub ExtPrice_DblClick(Cancel As Integer)
Dim strResp As String
    If Me.QuantityShipped = 0 Then
        MsgBox "Please put in quantity first!"
        Me.QuantityShipped.SetFocus
        Cancel = True
    Else
        strResp = InputBox("Extended Price")
        If IsNumeric(strResp) Then
            Me.Price = strResp / Me.QuantityShipped
            Price_AfterUpdate
        Else
            MsgBox "Please input a number!"
        End If
    End If
End Sub

Private Sub Form_Activate()
'    LineID.ColumnWidth = -2
'    QuantityShipped.ColumnWidth = -2
'    ItemID.ColumnWidth = -2
'    ItemDescription.ColumnWidth = -2
'    Price.ColumnWidth = -2
'    ExtPrice.ColumnWidth = -2
'    Cost.ColumnWidth = -2
'    ExtCost.ColumnWidth = -2
'    GP.ColumnWidth = -2
'    Repack.ColumnWidth = -2
End Sub

Private Sub Form_AfterUpdate()
    Dim nSavePos As Long
    Dim rs
    If gcfHandleErrors Then On Error GoTo ErrLog
    'If Me.Recordset.AbsolutePosition = Me.Recordset.Me.NewRecord Then
    '    nSavePos = ItemID
    '    Set rs = Me.Recordset.Clone
    '    rs.FindFirst "[ItemID] = " & Str(ItemID)
    '    Me.Bookmark = rs.Bookmark
    'Else
        'nSavePos = Me.Recordset.AbsolutePosition
        'Me.Requery
        'Me.Recordset.AbsolutePosition = nSavePos
    'End If
ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "Form_AfterUpdate"
    Resume ExitSub
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
Dim nCode As Long
Dim rs As Object

    Beep
    If gcfHandleErrors Then On Error GoTo ErrLog
    Select Case DataErr
    Case 3022 'duplicate records
        nCode = ItemID

        If MsgBox("Duplicate records not allowed; Do you want to go to line with code =" & nCode, vbYesNo) = vbYes Then
            Me.Form.Undo
            Me.FilterOn = False
            If gcfHandleErrors Then On Error GoTo ErrLog
            Set rs = Me.Recordset.Clone
            rs.FindFirst "[ItemID] = " & Str(nCode)
            Me.Bookmark = rs.Bookmark
        End If
        Response = acDataErrContinue
    Case 3146 'duplicate records
        nCode = ItemID

        If MsgBox("Duplicate records not allowed; Do you want to go to line with code =" & nCode, vbYesNo) = vbYes Then
            Me.Form.Undo
            Me.FilterOn = False
            If gcfHandleErrors Then On Error GoTo ErrLog
            Set rs = Me.Recordset.Clone
            rs.FindFirst "[ItemID] = " & Str(nCode)
            Me.Bookmark = rs.Bookmark
        End If
        Response = acDataErrContinue
    Case 2237
        If MsgBox("Not In List... do you want to set up item #" & ItemID, vbYesNo) = vbYes Then
            DoCmd.OpenForm "ItemSetup", acNormal, , , acFormAdd, acDialog, CStr(ItemID)
            Me.Refresh
        End If
        Response = acDataErrContinue

    Case Else
        'MsgBox "Unhandled Form Error " & DataErr & " " & Err.Number & " " & Err.Description
    End Select
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "ClubOrderSub:Form_Error"
    Resume ExitSub

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Not gcfHandleErrors Then Debug.Print KeyCode
    Select Case KeyCode

    Case 119 '{F8}
        KeyCode = 0
        cmdAddNew_Click

    Case 120 '{F9}
        KeyCode = 0
        QuantityShipped_DblClick False

    Case 121 '{F10}
        KeyCode = 0
        ExtCost_DblClick False

    Case 122 '{F11}
        KeyCode = 0
        ExtPrice_DblClick False

    Case Else
    End Select
End Sub

Private Sub Form_Load()
If gcfHandleErrors Then On Error GoTo ErrLog
    LineID.ColumnOrder = 1
    QuantityShipped.ColumnOrder = 2
    ItemID.ColumnOrder = 3
    ItemDescription.ColumnOrder = 4
    cmbShelf.ColumnOrder = 5
    Price.ColumnOrder = 6
    Sl.ColumnOrder = 7
    ExtPrice.ColumnOrder = 8
    Cost.ColumnOrder = 9
    ExtCost.ColumnOrder = 10
    GP.ColumnOrder = 11
    Repack.ColumnOrder = 12
    PrintLabels.ColumnOrder = 13

    LineID.ColumnWidth = 465
    QuantityShipped.ColumnWidth = 690
    ItemID.ColumnWidth = 1000
    ItemDescription.ColumnWidth = 3465
    cmbShelf.ColumnWidth = 600
    Price.ColumnWidth = 870
    Sl.ColumnWidth = 500
    ExtPrice.ColumnWidth = 1000
    Cost.ColumnWidth = 900
    ExtCost.ColumnWidth = 1000
    GP.ColumnWidth = 840
    Repack.ColumnWidth = 840
    PrintLabels.ColumnWidth = 840
    Me.DatasheetBackColor = LineID.BackColor
    'QuantityShipped.TabStop = False 'Not Me.Parent!chkSupplyAllItems

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Load"
    Resume ExitSub

End Sub

Private Sub ItemDescription_AfterUpdate()
    ItemID_AfterUpdate

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "ItemDescription_AfterUpdate"
    Resume ExitSub
End Sub

Private Sub ItemDescription_BeforeUpdate(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog
    If CLng(ItemDescription.Column(8)) <> Me.Parent.VendorID And Not (Me.Parent!chkSupplyAllItems) Then
        MsgBox "vendor not allowed!"
        Cancel = True
        Me.Undo
    End If
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "ItemDescription_BeforeUpdate"
    Resume ExitSub
End Sub

Private Sub ItemDescription_DblClick(Cancel As Integer)
    EditItem
End Sub

Private Sub ItemID_AfterUpdate()
If gcfHandleErrors Then On Error GoTo ErrLog
    Price = ItemID.Column(4)
    Cost = ItemID.Column(3)

    'added 2007-08-13
    If QuantityShipped = 0 Then ' load it with a case qty!
        QuantityShipped = ItemID.Column(7)
    End If
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "ItemID_AfterUpdate"
    Resume ExitSub
End Sub

Private Sub ItemID_BeforeUpdate(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog
    If CLng(ItemID.Column(9)) <> Me.Parent.VendorID And Not (Me.Parent!chkSupplyAllItems) Then
        MsgBox "vendor not allowed!"
        Cancel = True
        Me.Undo
    End If
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "ItemID_BeforeUpdate"
    Resume ExitSub

End Sub


Private Sub Form_BeforeInsert(Cancel As Integer)


'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim nNextLine As Long

If gcfHandleErrors Then On Error GoTo ErrLog


     'get the largest line number in this order...
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryPurchaseOrderLineMaxID"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parPurchaseOrderID", adBigInt, adParamInput)
        .Parameters("parPurchaseOrderID") = PurchaseOrderID
        Set rstemp = .Execute
        rstemp.MoveFirst
        If IsNull(rstemp!MaxLineID) Then
            nNextLine = 1
        Else
            nNextLine = rstemp!MaxLineID + 1
        End If
        rstemp.Close
        Set rstemp = Nothing
    End With
    Set comQDF = Nothing


    LineID = nNextLine 'Me.Recordset.RecordCount + 1

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_BeforeInsert"
    Resume ExitSub

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
Dim nUCase As Integer
If gcfHandleErrors Then On Error GoTo ErrLog
    ' added 8/24/2005 'invalid use of null
    If IsNull(ItemID) Then
        Cancel = True
        ItemDescription.SetFocus
        Exit Sub
    End If
    nUCase = ItemID.Column(7)
    If nUCase = 0 Then nUCase = 1

    Repack = IIf([QuantityShipped] - (([QuantityShipped] \ [nUCase]) * [nUCase]) = 0, 0, -1)
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_BeforeUpdate"
    Resume ExitSub

End Sub

Private Sub Form_Current()

If gcfHandleErrors Then On Error GoTo ErrLog

    If Not Me.NewRecord Then
        If Me.Parent.FormFooter.Visible Then

            Me.Parent("PurchaseOrdersOnFileSubItemMovement").Form.Filter = "((ItemMovement.Code)=" & ItemID & ")"
            Me.Parent("PurchaseOrdersOnFileSubItemMovement").Form.FilterOn = True

            Me.Parent("PurchaseOrdersOnFileSubItemSetup").Form.Filter = "((Item.ID)=" & ItemID & ")"
            Me.Parent("PurchaseOrdersOnFileSubItemSetup").Form.FilterOn = True
            Me.Parent("PurchaseOrdersOnFileChartSalesTotalSub").Requery
        End If
        If Me.Parent("PurchaseOrdersOnFileChartSalesTotalSub").Visible Then
            Me.Parent("PurchaseOrdersOnFileChartSalesTotalSub").Requery
        End If
        Me.Parent("cmdShowDetail").Tag = ItemID
        'added 2007-08-13
        QuantityShipped.SetFocus
    Else
        'added 2007-12-24
        If Me.Parent("chkSkipQuantityOnNewItems") Then
            'added 2007-08-13
            ItemID.SetFocus
        End If
    End If


    ItemID.TabStop = Me.NewRecord
    'ExtPrice.BackColor = vbCyan
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Current"
    Resume ExitSub
End Sub

Private Sub ItemID_DblClick(Cancel As Integer)
    EditItem


End Sub

Public Sub EditItem()
    OpenItemForm (CStr(ItemID))
    cmdRefreshx_Click
    ItemID_AfterUpdate
End Sub
Private Sub ItemID_NotInList(NewData As String, Response As Integer)
        Dim nSaveData As Long
        If gcfHandleErrors Then On Error GoTo ErrLog
        Beep
        If MsgBox("Not In List... do you want to set up item #" & NewData, vbYesNo) = vbYes Then
            DoCmd.OpenForm "ItemSetup", acNormal, , , acFormAdd, acDialog, CStr(NewData)
            DoCmd.Close acForm, "ItemSetup"

            nSaveData = NewData
            'Me.Refresh
            ItemID.Undo
            'Me.Refresh
            ItemID.Requery
            ItemID = nSaveData
            ItemID_AfterUpdate
        End If
        Response = acDataErrContinue
        Me.Refresh

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "ItemID_NotInList"
    Resume ExitSub
End Sub
Private Sub cmdOpenItem_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdOpenItem_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Item"

    stLinkCriteria = "[ID]=" & Me![ItemID]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdOpenItem_Click:
    Exit Sub

Err_cmdOpenItem_Click:
    MsgBox Err.Description
    Resume Exit_cmdOpenItem_Click

End Sub

Private Sub Price_AfterUpdate()
    If MsgBox("Do you want to change the PRICE of this item to " & Format(Price, "Currency") & "?", vbYesNo) = vbYes Then
        MsgBox ("The PRICE must be changed in the Item Screen!")
        EditItem
    End If
End Sub

Private Sub QuantityShipped_DblClick(Cancel As Integer)
    Me.FilterOn = Not Me.FilterOn
    Me.Filter = "QuantityShipped<>0"
End Sub
Private Sub cmdAddNew_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdAddNew_Click


    DoCmd.GoToRecord , , acNewRec
    ItemID.SetFocus

Exit_cmdAddNew_Click:
    Exit Sub

Err_cmdAddNew_Click:
    MsgBox Err.Description
    Resume Exit_cmdAddNew_Click

End Sub
Private Sub cmdRefresh_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdRefresh_Click

    Dim rs
    Dim strCode As String

    strCode = Me.ID

    Me.Parent.optSortBy_AfterUpdate

    Set rs = Me.RecordsetClone
    rs.FindFirst "[ItemID] = " & Str(strCode)
    Me.Bookmark = rs.Bookmark
    rs.Close
    Set rs = Nothing


Exit_cmdRefresh_Click:
    Exit Sub

Err_cmdRefresh_Click:
    MsgBox Err.Description
    Resume Exit_cmdRefresh_Click

End Sub
Private Sub cmdRefreshx_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdRefreshx_Click


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70

Exit_cmdRefreshx_Click:
    Exit Sub

Err_cmdRefreshx_Click:
    MsgBox Err.Description
    Resume Exit_cmdRefreshx_Click

End Sub
