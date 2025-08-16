Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub cmdfrmAddNew_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdfrmAddNew

    DoCmd.GoToRecord , , acNewRec

Exit_cmdfrmAddNew:
    Exit Sub

Err_cmdfrmAddNew:
    MsgBox Err.Description
    Resume Exit_cmdfrmAddNew
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

    DoCmd.DoMenuItem acFormBar, acEditMenu, acSelectRecord, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, acDelete, , acMenuVer70



ExitSub:
    Exit Sub

ErrLog:
    Select Case Err.Number
    Case 2501
        MsgBox "There are related items for this record"
    Case Else
    EMsgBox Me, "cmdfrmDeleteRecord_Click"
    End Select
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

Private Sub cmdRecalcUnitsPerOrder_Click()
    Dim stDocName As String
    If gcfHandleErrors Then On Error GoTo ErrLog1
    If MsgBox("Do you want to have the system recalculate the Units Per Order?", vbYesNo) = vbYes Then

        stDocName = "Items After Recalc of Units Per Order"

        SQLPassThru ("sp_ItemSetUnitsPerOrder")

        ArchiveReport stDocName
        DoCmd.OpenReport stDocName, acPreview
        cmdRecalcUnitsPerOrderUNDO.Visible = True
        cmdRecalcUnitsPerOrderUNDO.SetFocus
        cmdRecalcUnitsPerOrder.Visible = False
    End If

SetLocation:
    If gcfHandleErrors Then On Error GoTo ErrLog
    If MsgBox("Do you want to have the Location reset?" & vbCrLf & "This will reset the Location field!", vbYesNo) = vbYes Then
        SetOptions False
        stDocName = "qryItemSetLocation_1_BACK"
        DoCmd.OpenQuery stDocName
        stDocName = "qryItemSetLocation_2_SHELF"
        DoCmd.OpenQuery stDocName
        stDocName = "qryItemSetLocation_3_OFFICE"
        DoCmd.OpenQuery stDocName

        SQLPassThru ("sp_ItemSetUnitsPerOrder")

        SetOptions True

    End If

ExitSub:
    Exit Sub

ErrLog1:
    EMsgBox Me, "cmdRecalcUnitsPerOrder_Click"
    Resume SetLocation

ErrLog:
    EMsgBox Me, "cmdRecalcLocation_Click"
    Resume ExitSub


End Sub

Private Sub cmdRecalcUnitsPerOrderUNDO_Click()
    If MsgBox("Do you want to UNDO the changes made?", vbYesNo) = vbYes Then
        SQLPassThru ("sp_ItemSetUnitsPerOrderUNDO")
    End If

End Sub

Private Sub cmdReset_Click()
Month1 = (1 / 13)
Month2 = (1 / 13)
Month3 = (1 / 13)
Month4 = (1 / 13)
Month5 = (1 / 13)
Month6 = (1 / 13)
Month7 = (1 / 13)
Month8 = (1 / 13)
Month9 = (1 / 13)
Month10 = (1 / 13)
Month11 = (1 / 13)
Month12 = (1 / 13)
CurrentMonth = (1 / 13)
CheckDirty
End Sub

Private Sub CheckDirty()
    Dim nCurMM As Integer
    If gcfHandleErrors Then On Error GoTo ErrLog
    If Me.Form.Dirty Then '
        ChangeDate = Now
        Me.Form.Dirty = False
        Me.Refresh
    End If
    nCurMM = Month(Now)
    Select Case nCurMM
    Case 1 'January
        CurrMM = CurrentMonth
        Jan = Month12
        Feb = Month11
        Mar = Month10
        Apr = Month9
        May = Month8
        Jun = Month7
        Jul = Month6
        Aug = Month5
        Sep = Month4
        Oct = Month3
        Nov = Month2
        Dec = Month1
    Case 2 'FEB
        CurrMM = CurrentMonth
        Jan = Month1
        Feb = Month12
        Mar = Month11
        Apr = Month10
        May = Month9
        Jun = Month8
        Jul = Month7
        Aug = Month6
        Sep = Month5
        Oct = Month4
        Nov = Month3
        Dec = Month2
    Case 3 'MAR
        CurrMM = CurrentMonth
        Jan = Month2
        Feb = Month1
        Mar = Month12
        Apr = Month11
        May = Month10
        Jun = Month9
        Jul = Month8
        Aug = Month7
        Sep = Month6
        Oct = Month5
        Nov = Month4
        Dec = Month3
    Case 4 'APR
        CurrMM = CurrentMonth
        Jan = Month3
        Feb = Month2
        Mar = Month1
        Apr = Month12
        May = Month11
        Jun = Month10
        Jul = Month9
        Aug = Month8
        Sep = Month7
        Oct = Month6
        Nov = Month5
        Dec = Month4
    Case 5 'MAY
        CurrMM = CurrentMonth
        Jan = Month4
        Feb = Month3
        Mar = Month2
        Apr = Month1
        May = Month12
        Jun = Month11
        Jul = Month10
        Aug = Month9
        Sep = Month8
        Oct = Month7
        Nov = Month6
        Dec = Month5
    Case 6 'JUN
        CurrMM = CurrentMonth
        Jan = Month5
        Feb = Month4
        Mar = Month3
        Apr = Month2
        May = Month1
        Jun = Month12
        Jul = Month11
        Aug = Month10
        Sep = Month9
        Oct = Month8
        Nov = Month7
        Dec = Month6
    Case 7 'JUL
        CurrMM = CurrentMonth
        Jan = Month6
        Feb = Month5
        Mar = Month4
        Apr = Month3
        May = Month2
        Jun = Month1
        Jul = Month12
        Aug = Month11
        Sep = Month10
        Oct = Month9
        Nov = Month8
        Dec = Month7
    Case 8 'AUG
        CurrMM = CurrentMonth
        Jan = Month7
        Feb = Month6
        Mar = Month5
        Apr = Month4
        May = Month3
        Jun = Month2
        Jul = Month1
        Aug = Month12
        Sep = Month11
        Oct = Month10
        Nov = Month9
        Dec = Month8
    Case 9 'SEP
        CurrMM = CurrentMonth
        Jan = Month8
        Feb = Month7
        Mar = Month6
        Apr = Month5
        May = Month4
        Jun = Month3
        Jul = Month2
        Aug = Month1
        Sep = Month12
        Oct = Month11
        Nov = Month10
        Dec = Month9
    Case 10 'OCT
        CurrMM = CurrentMonth
        Jan = Month9
        Feb = Month8
        Mar = Month7
        Apr = Month6
        May = Month5
        Jun = Month4
        Jul = Month3
        Aug = Month2
        Sep = Month1
        Oct = Month12
        Nov = Month11
        Dec = Month10
    Case 11 'NOV
        CurrMM = CurrentMonth
        Jan = Month10
        Feb = Month9
        Mar = Month8
        Apr = Month7
        May = Month6
        Jun = Month5
        Jul = Month4
        Aug = Month3
        Sep = Month2
        Oct = Month1
        Nov = Month12
        Dec = Month11
    Case 12 'DEC
        CurrMM = CurrentMonth
        Jan = Month11
        Feb = Month10
        Mar = Month9
        Apr = Month8
        May = Month7
        Jun = Month6
        Jul = Month5
        Aug = Month4
        Sep = Month3
        Oct = Month2
        Nov = Month1
        Dec = Month12
    Case Else
        MsgBox "Month not found!" & Month(Now)
    End Select

    If Me.Dirty Then
        Me.Dirty = False
        Me.Refresh
    End If


ExitSub:
    Exit Sub
ErrLog:
    Select Case Err.Number
    Case 3146 'description null
        EMsgBox Me, "Description can not be blank!"
        Me.Undo
        Description.SetFocus
        Resume ExitSub
    Case Else
        EMsgBox Me, "Check Dirty" & vbCrLf & Err.Number & " " & Err.Description
    End Select
    Resume ExitSub





End Sub

Private Sub Command32_Click()
If gcfHandleErrors Then On Error GoTo Err_Command32_Click
    lblPleaseWait.Visible = True
    Me.Repaint
    SetOptions False
    AssignDaysSupplyGroupIDs
    lblPleaseWait.Visible = False
Exit_Command32_Click:
    SetOptions True

    DoCmd.Close

    Exit Sub

Err_Command32_Click:
    MsgBox Err.Description
    Resume Exit_Command32_Click

End Sub

Private Sub CurrentMonth_Exit(Cancel As Integer)
            CheckDirty

End Sub


Private Sub Description_Exit(Cancel As Integer)
'    CurrentMonth.SetFocus
End Sub

Private Sub Form_AfterUpdate()
    Combo6.Requery
    Combo6 = ID

End Sub

Private Sub Form_Current()
    If Me.NewRecord Then
        Description.SetFocus
    Else
        Combo6 = ID
    End If
    cmdfrmAddNew.Visible = True 'Not Me.NewRecord
    cmdfrmDeleteRecord.Visible = Not Me.NewRecord
    cmdfrmSaveRecord.Visible = Me.NewRecord
    cmdfrmUndo.Visible = Me.NewRecord
End Sub

Private Sub Form_Dirty(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog
    cmdfrmSaveRecord.Visible = True
    cmdfrmUndo.Visible = True
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Dirty"
    Resume ExitSub
End Sub

Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
    Call DoMouseWheel(Me, Count)

End Sub

Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)
End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub

Private Sub Month1_Exit(Cancel As Integer)
            CheckDirty

End Sub

Private Sub Month10_Exit(Cancel As Integer)
            CheckDirty

End Sub

Private Sub Month11_Exit(Cancel As Integer)
            CheckDirty

End Sub

Private Sub Month12_Exit(Cancel As Integer)
            CheckDirty

End Sub

Private Sub Month2_Exit(Cancel As Integer)
            CheckDirty

End Sub

Private Sub Month3_Exit(Cancel As Integer)
            CheckDirty

End Sub

Private Sub Month4_Exit(Cancel As Integer)
            CheckDirty

End Sub

Private Sub Month5_Exit(Cancel As Integer)
            CheckDirty

End Sub

Private Sub Month6_Exit(Cancel As Integer)
            CheckDirty

End Sub

Private Sub Month7_Exit(Cancel As Integer)
            CheckDirty

End Sub

Private Sub Month8_Exit(Cancel As Integer)
            CheckDirty

End Sub

Private Sub Month9_Exit(Cancel As Integer)
            CheckDirty

End Sub
Private Sub Combo6_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(Nz(Me![Combo6], 0))
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
End Sub
