Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Back_AfterUpdate()
    CountAfterUpdate
End Sub

Private Sub cmbItemID_AfterUpdate()

    RefreshFields
    PhysicalCount = OnShelf

    CountFront = 0
    CountBack = 0
    CountDisplay = 0
    CountOverflow = 0

    Select Case Me.Parent.optLocations.Value
    Case 1 ' front
        CountFront = OnShelf
    Case 2 'back
        CountBack = OnShelf
    Case 3 'display
        CountDisplay = OnShelf
    Case Else ' Overflow
        CountOverflow = OnShelf
    End Select


    VarianceAmount = PhysicalCount - OnShelf

ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "cmbItemID_AfterUpdate"
    Resume ExitSub

End Sub

Private Sub cmbItemID_DblClick(Cancel As Integer)
    OpenItemForm (CStr(cmbItemID))

End Sub

Private Sub CountAfterUpdate()
    If gcfHandleErrors Then On Error GoTo ErrLog

    PhysicalCount = Nz(CountFront) + Nz(CountBack) + Nz(CountDisplay) + Nz(CountOverflow)
    VarianceAmount = PhysicalCount - OnShelf

ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "CountAfterUpdate"
    If IsNull(ItemID) Then
        ItemID.SetFocus
    End If
    Resume ExitSub

End Sub


Private Sub Display_AfterUpdate()
CountAfterUpdate
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    RefreshFields

End Sub

Private Sub RefreshFields()
    If gcfHandleErrors Then On Error GoTo ErrLog
    If Not IsNull(cmbItemID.Column(5)) Then
        OnShelf = cmbItemID.Column(2)
        Price = cmbItemID.Column(3)
        Cost = cmbItemID.Column(4)
        Me.Parent.txtDispItem = cmbItemID
        Me.Parent.txtDispDescription = ItemDescription.Column(0)
        Me.Parent.txtDispOnShelf = OnShelf
    End If
ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "RefreshFields"
    Resume ExitSub

End Sub

Private Sub Form_Current()
Dim nI As Integer
Dim nCol As Integer

    If gcfHandleErrors Then On Error GoTo ErrLog

    If Not Me.NewRecord Then

        Me.Parent.txtDispItem = cmbItemID
        Me.Parent.txtDispDescription = ItemDescription.Column(0)
        Me.Parent.txtDispOnShelf = OnShelf
    Else
        SetColumnWidth
        cmbItemID.SetFocus
    End If
ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "Form_Current"
    Resume ExitSub

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog
    Select Case DataErr

    Case 3022 'duplicate records
        DuplicateRecord
        Response = acDataErrContinue
    Case 3621 'duplicate reco140rds
        DuplicateRecord
        Response = acDataErrContinue
    Case 3146 'duplicate records
        DuplicateRecord
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

Private Sub DuplicateRecord()

    Dim nCode As Double
    Dim strSaveDesc As String

    Dim rs As Object
    Dim nFrontOldvalue As Integer
    Dim nBackOldvalue As Integer
    Dim nDisplayOldvalue As Integer
    Dim nOverflowOldvalue As Integer
    If gcfHandleErrors Then On Error GoTo ErrLog
    nCode = cmbItemID
    strSaveDesc = ItemDescription.Column(0)

        'If MsgBox("Duplicate records not allowed; Do you want to go to line with code =" & nCode, vbYesNo) = vbYes Then

    nFrontOldvalue = Front
    nBackOldvalue = Back
    nDisplayOldvalue = Display
    nOverflowOldvalue = Overflow

    Me.Form.Undo
    If gcfHandleErrors Then On Error GoTo ErrLog
    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ItemID] = " & Str(nCode)
    Me.Bookmark = rs.Bookmark
    Select Case Me.Parent.optCountBy
    Case 1
        PhysicalCount.SetFocus
    Case 2
        Var.SetFocus
    Case Else
        Select Case Me.Parent.optLocations.Value
        Case 1 ' front
            If Front > 0 And nFrontOldvalue <> Front Then
                If MsgBox("Overwrite existing Front count of " & Front & " with " & nFrontOldvalue & "?", vbYesNo, strSaveDesc & " Front has been counted!") = vbNo Then
                    Exit Sub
                End If
            End If

            Front = nFrontOldvalue
            Front.SetFocus

        Case 2 'back
            If Back > 0 And nBackOldvalue <> Back Then
                If MsgBox("Overwrite existing Back count of " & Back & " with " & nBackOldvalue & "?", vbYesNo, strSaveDesc & " Back has been counted!") = vbNo Then
                    Exit Sub
                End If
            End If
            Back = nBackOldvalue
            Back.SetFocus
        Case 3 'display
            If Display > 0 And nDisplayOldvalue <> Display Then
                If MsgBox("Overwrite existing Display count of " & Display & " with " & nDisplayOldvalue & "?", vbYesNo, strSaveDesc & " Display has been counted!") = vbNo Then
                    Exit Sub
                End If
            End If
            Display = nDisplayOldvalue
            Display.SetFocus
        Case Else ' Overflow
            If Overflow > 0 And nOverflowOldvalue <> Overflow Then
                If MsgBox("Overwrite existing Overflow count of " & Overflow & " with " & nOverflowOldvalue & "?", vbYesNo, strSaveDesc & " Overflow has been counted!") = vbNo Then
                    Exit Sub
                End If
            End If
            Overflow = nOverflowOldvalue
            Overflow.SetFocus
        End Select
    End Select
    CountAfterUpdate
        'End If
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "ClubOrderSub:DuplicateRecord"
    Resume ExitSub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If gcfHandleErrors Then On Error GoTo ErrLog
    Select Case KeyCode
    Case 119 '{F8}
        KeyCode = 0
        cmdAddNew_Click
    Case 120 '{F9}
        KeyCode = 0
        'QuantityShipped_DblClick False




    Case Else
    End Select
ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "Form_KeyDown"
    Resume ExitSub

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

    Front.ColumnOrder = nI
    nI = nI + 1
    Back.ColumnOrder = nI
    nI = nI + 1
    Display.ColumnOrder = nI
    nI = nI + 1
    Overflow.ColumnOrder = nI
    nI = nI + 1


    PhysicalCount.ColumnOrder = nI
    nI = nI + 1
    VarianceAmount.ColumnOrder = nI
    nI = nI + 1
    Price.ColumnOrder = nI
    nI = nI + 1
    ExtPrice.ColumnOrder = nI
    nI = nI + 1
    Cost.ColumnOrder = nI
    nI = nI + 1
    ExtCost.ColumnOrder = nI
    nI = nI + 1
    ChangeDate.ColumnOrder = nI
    SetColumnWidth
ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "cmbItemID_AfterUpdate"
    Resume ExitSub

End Sub

Private Sub SetColumnWidth()

    Dim nI As Integer
    If gcfHandleErrors Then On Error GoTo ErrLog
    nI = -2
    cmbItemID.ColumnWidth = nI
    ItemDescription.ColumnWidth = nI
    Size.ColumnWidth = nI
    OnShelf.ColumnWidth = nI
    Front.ColumnWidth = nI
    Back.ColumnWidth = nI
    Overflow.ColumnWidth = nI
    Display.ColumnWidth = nI

    PhysicalCount.ColumnWidth = nI
    VarianceAmount.ColumnWidth = nI
    Price.ColumnWidth = nI
    Cost.ColumnWidth = nI
    ExtPrice.ColumnWidth = nI
    ExtCost.ColumnWidth = nI
    ChangeDate.ColumnWidth = nI

ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "SetColumnWidth"
    Resume ExitSub


End Sub

Private Sub Front_AfterUpdate()
CountAfterUpdate
End Sub

Private Sub ItemDescription_AfterUpdate()
    If gcfHandleErrors Then On Error GoTo ErrLog
    RefreshFields

    Count = OnShelf
    CountFront = 0
    CountBack = 0
    CountDisplay = 0
    CountOverflow = 0

    Select Case Me.Parent.optLocations.Value
    Case 1 ' front
        CountFront = OnShelf
    Case 2 'back
        CountBack = OnShelf
    Case 3 'display
        CountDisplay = OnShelf
    Case Else ' Overflow
        CountOverflow = OnShelf
    End Select

    VarianceAmount = Count - OnShelf


ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "ItemDescription_AfterUpdate"
    Resume ExitSub


End Sub

Private Sub Overflow_AfterUpdate()
    CountAfterUpdate
End Sub

Private Sub PhysicalCount_AfterUpdate()
    If gcfHandleErrors Then On Error GoTo ErrLog
    VarianceAmount = PhysicalCount - OnShelf
    CountFront = PhysicalCount
    CountBack = 0
    CountDisplay = 0
    CountOverflow = 0
ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "PhysicalCount_AfterUpdate"
    Resume ExitSub

End Sub

Private Sub VarianceAmount_AfterUpdate()
    If gcfHandleErrors Then On Error GoTo ErrLog
    CountFront = VarianceAmount + OnShelf
    CountBack = 0
    CountDisplay = 0
    CountOverflow = 0
    PhysicalCount = CountFront

    'PhysicalCount = VarianceAmount + OnShelf
ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "VarianceAmount_AfterUpdate"
    Resume ExitSub


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
