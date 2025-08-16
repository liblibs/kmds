Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Dim gbRunMailChanges As Boolean

Private Sub cboID_AfterUpdate()
   ' Find the record that matches the control.
    Dim rs As Object
    If gcfHandleErrors Then On Error GoTo ErrLog
    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(Me![cboID])
    Me.Bookmark = rs.Bookmark
ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "cboID_AfterUpdate"
    Resume ExitSub

End Sub


Private Sub cmdCashRegisterSystemStatus_Click()
    Dim strstatus As String
    strstatus = CashStatus()
    MsgBox strstatus
End Sub

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


Private Sub Form_AfterUpdate()
    cboID.Requery
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    gbRunMailChanges = True
    ChangeDate = Now
End Sub

Private Sub Form_Current()
    If gcfHandleErrors Then On Error GoTo ErrLog

    ID.Visible = Me.NewRecord
    cboID.Visible = Not Me.NewRecord
    cboID = ID
    cmdfrmAddNew.Visible = Not Me.NewRecord
    cmdfrmDeleteRecord.Visible = Not Me.NewRecord
    cmdfrmSaveRecord.Visible = Me.NewRecord
    cmdfrmUndo.Visible = Me.NewRecord
    If Not Me.NewRecord Then
        If IsNull(chkActive) Then chkActive = False
        tabCashsys.Enabled = chkActive
        If tabCashsys.Enabled Then
            ComputerName.BackColor = Desc.BackColor
        Else
            ComputerName.BackColor = vbRed

        End If

    End If
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Current"
    Resume ExitSub
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

Private Sub Form_Error(DataErr As Integer, Response As Integer)


    If gcfHandleErrors Then On Error GoTo ErrLog
    Select Case DataErr
    Case 3146 'description = blank
        Me.Undo
        DoCmd.GoToRecord , , acPrevious
        Response = acDataErrContinue

    Case Else
        'MsgBox "Unhandled Form Error " & DataErr & " " & Err.Number & " " & Err.Description
    End Select
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "Sizes:Form_Error"
    Resume ExitSub

End Sub


Private Sub cmdGoToNew_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdGoToNew_Click


    DoCmd.GoToRecord , , acNewRec

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

Exit_cmdDeleteRecord_Click:
    Exit Sub

Err_cmdDeleteRecord_Click:
    EMsgBox Me, "cmdDeleteRecord_Click"
    Resume Exit_cmdDeleteRecord_Click

End Sub

Private Sub cmdGoToLastRecord_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdGoToLastRecord_Click


    DoCmd.GoToRecord , , acLast

Exit_cmdGoToLastRecord_Click:
    Exit Sub

Err_cmdGoToLastRecord_Click:
    EMsgBox Me, "cmdGoToLastRecord_Click"
    Resume Exit_cmdGoToLastRecord_Click

End Sub
Private Sub cmdGoToNext_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdGoToNext_Click


    DoCmd.GoToRecord , , acNext

Exit_cmdGoToNext_Click:
    Exit Sub

Err_cmdGoToNext_Click:
    Select Case Err.Number
    Case 2105
        Beep
    Case Else
        EMsgBox Me, "cmdGoToNext_Click"
    End Select
    Resume Exit_cmdGoToNext_Click

End Sub
Private Sub cmdGoToPreviousRecord_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdGoToPreviousRecord_Click


    DoCmd.GoToRecord , , acPrevious

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

Exit_cmdGoToFirstRecord_Click:
    Exit Sub

Err_cmdGoToFirstRecord_Click:
    EMsgBox Me, "cmdGoToFirstRecord_Click"
    Resume Exit_cmdGoToFirstRecord_Click

End Sub
Private Sub cmdUndoRecord_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdUndoRecord_Click


    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70

Exit_cmdUndoRecord_Click:
    Exit Sub

Err_cmdUndoRecord_Click:
    MsgBox Err.Description
    Resume Exit_cmdUndoRecord_Click

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


Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
    Call DoMouseWheel(Me, Count)

End Sub

Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)
    gbRunMailChanges = False
End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gbRunMailChanges Then
        MailChanges (False)
    End If

End Sub

Private Sub PrinterPort_Exit(Cancel As Integer)
    DrawerPort = PrinterPort
End Sub
