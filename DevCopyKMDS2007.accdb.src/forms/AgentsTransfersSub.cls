Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




Private Sub Cost_AfterUpdate()
    If MsgBox("Do you want to change the COST of this item?", vbYesNo) = vbYes Then
        MsgBox ("The COST must be changed in the Item Screen!")
        EditItem
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
    QuantityOrdered.ColumnWidth = 0

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
Dim nCode As Double '07/17/2021
Dim rs As Object

    If gcfHandleErrors Then On Error GoTo ErrLog
    Select Case DataErr
    Case 3022 'duplicate records
        nCode = ItemID

        If MsgBox("Duplicate records not allowed; Do you want to go to line with code =" & nCode, vbYesNo) = vbYes Then
            Me.Form.Undo
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
            If gcfHandleErrors Then On Error GoTo ErrLog
            Set rs = Me.Recordset.Clone
            rs.FindFirst "[ItemID] = " & Str(nCode)
            Me.Bookmark = rs.Bookmark
        End If
        Response = acDataErrContinue
    Case 2237
        MsgBox "Please select an item from the list"
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

Private Sub Form_Load()
If gcfHandleErrors Then On Error GoTo ErrLog
    LineID.ColumnOrder = 1
    QuantityShipped.ColumnOrder = 2
    ItemID.ColumnOrder = 3
    ClaimCode.ColumnOrder = 4
    ItemDescription.ColumnOrder = 5
    Price.ColumnOrder = 6
    ExtPrice.ColumnOrder = 7
    'Repack.ColumnOrder = 10

    LineID.ColumnWidth = 1000
    QuantityShipped.ColumnWidth = 690
    ItemID.ColumnWidth = 1500

    ClaimCode.ColumnWidth = 1500

    ItemDescription.ColumnWidth = 3465
    Price.ColumnWidth = 870
    ExtPrice.ColumnWidth = 1000
    'Repack.ColumnWidth = 840
    Me.DatasheetBackColor = LineID.BackColor
    Select Case Me.Parent.Status
    Case "O"
        ClaimCode.Visible = False
    Case "I"
        ClaimCode.Visible = False
    Case "C"
        ClaimCode.Visible = True
    End Select



ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Load"
    Resume ExitSub

End Sub

Private Sub ItemDescription_AfterUpdate()
If gcfHandleErrors Then On Error GoTo ErrLog
    'Price = ItemDescription.Column(4)
    'Cost = ItemDescription.Column(3)
    Cost = ItemID.Column(4)
    Price = ItemID.Column(5)


ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "ItemDescription_AfterUpdate"
    Resume ExitSub
End Sub

Private Sub ItemDescription_BeforeUpdate(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog
    'If CLng(ItemDescription.Column(6)) <> Me.Parent.VendorID And Not (Nz(Me.Parent!chkSupplyAllItems, 0)) Then
    '    MsgBox "vendor not allowed!"
    '    Cancel = True
    '    Me.Undo
    'End If
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
    Cost = ItemID.Column(4)
    Price = ItemID.Column(5)
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "ItemID_AfterUpdate"
    Resume ExitSub
End Sub

Private Sub ItemID_BeforeUpdate(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "ItemID_BeforeUpdate"
    Resume ExitSub

End Sub


Private Sub Form_BeforeInsert(Cancel As Integer)

If gcfHandleErrors Then On Error GoTo ErrLog

    LineID = Me.Recordset.RecordCount + 1
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_BeforeInsert"
    Resume ExitSub

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
Dim nUCase As Integer

If gcfHandleErrors Then On Error GoTo ErrLog
'ver1.145
    If IsNull(QuantityOrdered) Then QuantityOrdered = 0
    If IsNull(QuantityShipped) Then QuantityShipped = 0
    nUCase = ItemID.Column(7)
    If nUCase = 0 Then nUCase = 1
    Select Case Me.Parent.Status
    Case "O" 'transfer out
        QuantityOrdered = 0 - QuantityShipped
    Case "C"  'claim
        QuantityOrdered = 0 - QuantityShipped
    Case Else ' transfer in
        QuantityOrdered = QuantityShipped
    End Select
    'Repack = IIf([QuantityShipped] - (([QuantityShipped] \ [nUCase]) * [nUCase]) = 0, 0, -1)
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

'            Me.Parent("PurchaseOrdersOnFileSubItemMovement").Form.Filter = "((ItemMovement.Code)=" & ItemID & ")"
'            Me.Parent("PurchaseOrdersOnFileSubItemMovement").Form.FilterOn = True
'
'            Me.Parent("PurchaseOrdersOnFileSubItemSetup").Form.Filter = "((Item.ID)=" & ItemID & ")"
'            Me.Parent("PurchaseOrdersOnFileSubItemSetup").Form.FilterOn = True
        End If
    End If

    ItemID.TabStop = Me.NewRecord
ExtPrice.BackColor = vbCyan
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Current"
    Resume ExitSub

End Sub

Private Sub ItemID_DblClick(Cancel As Integer)
    EditItem
    'Me.Parent.cmdRefresh_Click
End Sub

Private Sub ItemID_NotInList(NewData As String, Response As Integer)
        Dim nSaveData As Double '07/17/2021
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
Private Sub cmdRefresh_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdRefresh_Click


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70

Exit_cmdRefresh_Click:
    Exit Sub

Err_cmdRefresh_Click:
    MsgBox Err.Description
    Resume Exit_cmdRefresh_Click

End Sub

Private Sub Price_AfterUpdate()
    If MsgBox("Do you want to change the PRICE of this item?", vbYesNo) = vbYes Then
        MsgBox ("The PRICE must be changed in the Item Screen!")
        EditItem
    End If

End Sub



Public Sub EditItem()
    OpenItemForm (CStr(ItemID))
    cmdRefresh_Click
    ItemID_AfterUpdate
End Sub
