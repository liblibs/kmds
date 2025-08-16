Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmbItemID_AfterUpdate()
    RefreshFields
        PhysicalCount = OnShelf
        VarianceAmount = PhysicalCount - OnShelf

End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    RefreshFields

End Sub

Private Sub RefreshFields()
    If Not IsNull(cmbItemID.Column(5)) Then
        OnShelf = cmbItemID.Column(2)
        Price = cmbItemID.Column(3)
        Cost = cmbItemID.Column(4)
    End If
End Sub

Private Sub Form_Current()
Dim nI As Integer
    nI = -2
    cmbItemID.ColumnWidth = nI
    ItemDescription.ColumnWidth = nI
    Size.ColumnWidth = nI
    OnShelf.ColumnWidth = nI
    PhysicalCount.ColumnWidth = nI
    VarianceAmount.ColumnWidth = nI
    Price.ColumnWidth = nI
    Cost.ColumnWidth = nI
    ExtPrice.ColumnWidth = nI
    ExtCost.ColumnWidth = nI
    ChangeDate.ColumnWidth = nI

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog
    Select Case DataErr
    Dim nCode As Double '07/17/2021
    Dim rs As Object
    Case 3022 'duplicate records
        nCode = cmbItemID

        If MsgBox("Duplicate records not allowed; Do you want to go to line with code =" & nCode, vbYesNo) = vbYes Then
            Me.Form.Undo
            If gcfHandleErrors Then On Error GoTo ErrLog
            Set rs = Me.Recordset.Clone
            rs.FindFirst "[ItemID] = " & Str(nCode)
            Me.Bookmark = rs.Bookmark
            PhysicalCount.SetFocus
        End If
        Response = acDataErrContinue
    Case 3621 'duplicate records
        nCode = cmbItemID

        If MsgBox("Duplicate records not allowed; Do you want to go to line with code =" & nCode, vbYesNo) = vbYes Then
            Me.Form.Undo
            If gcfHandleErrors Then On Error GoTo ErrLog
            Set rs = Me.Recordset.Clone
            rs.FindFirst "[ItemID] = " & Str(nCode)
            Me.Bookmark = rs.Bookmark
            PhysicalCount.SetFocus
        End If
        Response = acDataErrContinue
    Case 3146 'duplicate records
        nCode = cmbItemID

        If MsgBox("Duplicate records not allowed; Do you want to go to line with code =" & nCode, vbYesNo) = vbYes Then
            Me.Form.Undo
            If gcfHandleErrors Then On Error GoTo ErrLog
            Set rs = Me.Recordset.Clone
            rs.FindFirst "[ItemID] = " & Str(nCode)
            Me.Bookmark = rs.Bookmark
            PhysicalCount.SetFocus
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 119 '{F8}
        KeyCode = 0
        cmdAddNew_Click
    Case 120 '{F9}
        KeyCode = 0
        'QuantityShipped_DblClick False




    Case Else
    End Select
End Sub

Private Sub Form_Load()
Dim nI As Integer
    nI = nI + 1
    cmbItemID.ColumnOrder = nI
    nI = nI + 1
    ItemDescription.ColumnOrder = nI
    nI = nI + 1
    Size.ColumnOrder = nI
    nI = nI + 1
    OnShelf.ColumnOrder = nI
    nI = nI + 1
    PhysicalCount.ColumnOrder = nI
    nI = nI + 1
    VarianceAmount.ColumnOrder = nI
    nI = nI + 1
    Price.ColumnOrder = nI
    nI = nI + 1
    Cost.ColumnOrder = nI
    nI = nI + 1
    ExtPrice.ColumnOrder = nI
    nI = nI + 1
    ExtCost.ColumnOrder = nI
    nI = nI + 1
    ChangeDate.ColumnOrder = nI

    nI = -2
    cmbItemID.ColumnWidth = nI
    ItemDescription.ColumnWidth = nI
    Size.ColumnWidth = nI
    OnShelf.ColumnWidth = nI
    PhysicalCount.ColumnWidth = nI
    VarianceAmount.ColumnWidth = nI
    Price.ColumnWidth = nI
    Cost.ColumnWidth = nI
    ExtPrice.ColumnWidth = nI
    ExtCost.ColumnWidth = nI
    ChangeDate.ColumnWidth = nI


End Sub

Private Sub ItemDescription_AfterUpdate()
    RefreshFields
    'If Me.NewRecord Then
        PhysicalCount = OnShelf
        VarianceAmount = PhysicalCount - OnShelf

    'End If

End Sub

Private Sub PhysicalCount_AfterUpdate()
    VarianceAmount = PhysicalCount - OnShelf
    'If VarianceAmount < 0 Then

End Sub

Private Sub VarianceAmount_AfterUpdate()
    PhysicalCount = VarianceAmount + OnShelf

End Sub

Private Sub cmdAddNew_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdAddNew_Click


    DoCmd.GoToRecord , , acNewRec
    cmbItemID.SetFocus

Exit_cmdAddNew_Click:
    Exit Sub

Err_cmdAddNew_Click:
    MsgBox Err.Description
    Resume Exit_cmdAddNew_Click

End Sub
Private Sub cmdSave_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdSave_Click


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70

Exit_cmdSave_Click:
    Exit Sub

Err_cmdSave_Click:
    MsgBox Err.Description
    Resume Exit_cmdSave_Click

End Sub
