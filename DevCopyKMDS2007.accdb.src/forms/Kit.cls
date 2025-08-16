Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub cboID_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object
    If gcfHandleErrors Then On Error GoTo ErrLog

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(Nz(Me![cboID], 0))
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cboID_AfterUpdate"
    Resume ExitSub

End Sub

Private Sub cmbItemByDesc_AfterUpdate()
    ItemID = cmbItemByDesc
    Me![cmbItemByUPC] = cmbItemByDesc

End Sub

Private Sub cmbItemByUPC_AfterUpdate()
    Dim rs As Object

    If gcfHandleErrors Then On Error GoTo ErrLog
    'Set RS = Me.Recordset.Clone
    'RS.FindFirst "[ID] = " & Str(Me![cmbItemByUPC])
    'Me.Bookmark = RS.Bookmark
    'mnLastComboBox = 3
    'TabCombos
    ItemID = cmbItemByUPC
    Me![cmbItemByDesc] = cmbItemByUPC


ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmbItemByUPC_AfterUpdate"
    Resume ExitSub
End Sub



Private Sub cmdfrmAddNew_Click()
Dim rs As Object
If gcfHandleErrors Then On Error GoTo ErrLog

    '    DoCmd.GoToRecord , , acNewRec
    DoCmd.OpenForm "KitNewItem", acNormal, , , acFormAdd, acDialog
    If GetKMDSSettings("KitLastItem", "NA") = "NA" Then
'        MsgBox "not saved correctly"
    Else
'        MsgBox ("OK " & GetKMDSSettings("KitLastItem", "NA"))
        cboID.Requery
        cboID = GetKMDSSettings("KitLastItem", "NA")
        Me.Requery
        Me.Refresh

        'move the the last record
        Set rs = Me.Recordset.Clone
        rs.MoveLast
        Me.Bookmark = rs.Bookmark

    End If


ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmbItemByUPC_AfterUpdate"
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

    DoCmd.DoMenuItem acFormBar, acEditMenu, acSelectRecord, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, acDelete, , acMenuVer70
    SetOptions False
ExitSub:
    SetOptions True
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



Private Sub Form_AfterDelConfirm(Status As Integer)
    cboID.Requery
End Sub

Private Sub Form_AfterInsert()
    cboID.Requery
End Sub

Private Sub Form_AfterUpdate()
    cboID.Requery
    gbRunMailChanges = True
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    ChangeDate = Now
End Sub

Private Sub Form_Current()
    If gcfHandleErrors Then On Error Resume Next

    cboID.Visible = Not Me.NewRecord
    cmdfrmAddNew.Visible = Not Me.NewRecord
    cmdfrmDeleteRecord.Visible = Not Me.NewRecord
    KitISubForm.Visible = Not Me.NewRecord

    cmbItemByDesc.Visible = Me.NewRecord
    cmbItemByUPC.Visible = Me.NewRecord
    cmdfrmSaveRecord.Visible = Me.NewRecord
    cmdfrmUndo.Visible = Me.NewRecord

    If gcfHandleErrors Then On Error GoTo ErrLog

    cboID = ID
    'cmbItemByUPC_AfterUpdate
    If Me.NewRecord Then
        lblStatus.Caption = "Setting up a new Kit."
        lblInitialInstruction.Caption = "Select the Kit's Item from either of the dropdowns below"

    Else
        lblStatus.Caption = "Changing an existing Kit."
        lblInitialInstruction.Caption = "Select a Kit to change   OR    Click the 'Add New' To set up a new Kit."
        CheckPrice
'        If Price < txtLiquorValue Then
'            MsgBox "Price can not be less than the liquor value"
'            Price = txtLiquorValue
'        End If

    End If

    'Me![cmbItemByDesc] = ItemID
    'Me![cmbItemByUPC] = ItemID
    'cboID = ID

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "Form_Current"
    Resume ExitSub

End Sub

Private Sub Form_Dirty(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog
    cmdfrmSaveRecord.Visible = True
    cmdfrmUndo.Visible = True
    KitISubForm.Visible = True
    Me.Repaint
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Dirty"
    Resume ExitSub
End Sub

Private Sub Form_Load()
    'Me.Filter = "((ID>1))"
    'Me.FilterOn = True
If gcfHandleErrors Then On Error GoTo ErrLog
    cboID.ColumnHidden = True
    SetupDate.ColumnHidden = True
    ChangeDate.ColumnHidden = True
    txtSumPrice.BackColor = glnColorLockedForAll
    txtDiscount.BackColor = glnColorLockedForAll
    txtDiscPercent.BackColor = glnColorLockedForAll
    'cboID.ColumnHidden = False
'    MsgBox "Load"
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Load"
    Resume ExitSub
End Sub

Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
    Call DoMouseWheel(Me, Count)

End Sub

Private Sub Form_Open(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog
    Cancel = FormLogin(glnCashierID, Me)
    If Not Cancel Then
        SQLPassThru "sp_KitItemWithPriceChange"
        Me.Requery
        If Me.OpenArgs <> "" Then
            'Me.Description.SetFocus
            'Me.Description.Text = Me.OpenArgs
        End If
    End If
'    MsgBox "open"
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Open"
    Resume ExitSub
End Sub


Private Sub Form_Close()
    FormLogout glnCashierID, Me
    If gbRunMailChanges Then
        MailChanges (False)
    End If

End Sub



Private Sub Price_AfterUpdate()
    CheckPrice
End Sub

Private Sub txtLiquorValue_AfterUpdate()
'    If Price < txtLiquorValue Then
'        MsgBox "Price can not be less than the liquor value"
'        Price = txtLiquorValue
'    End If
    CheckPrice

End Sub

Private Sub CheckPrice()
If gcfHandleErrors Then On Error GoTo ErrLog
    If Form.CurrentView = 1 Then ' Form View.... Data sheet view has a problem with calculated fields..
        If Price < txtLiquorValue Then
            MsgBox "Kit Price (" & Format(Price, "CURRENCY") & ") can not be less than the liquor value (" & Format(txtLiquorValue, "CURRENCY") & ")."
            Price = txtLiquorValue
        End If
    End If
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "CheckPrice"
    Resume ExitSub
End Sub
