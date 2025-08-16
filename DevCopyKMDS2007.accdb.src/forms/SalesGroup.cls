Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Dim gbRunMailChanges As Boolean

Private Sub cmbDescription_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(Nz(Me![cmbDescription], 0))
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark


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


Private Sub Form_BeforeUpdate(Cancel As Integer)
    gbRunMailChanges = True
End Sub

Private Sub Form_Current()
    cmbDescription = ID
    cmbDescription.Visible = Not Me.NewRecord
    Description.Visible = Me.NewRecord

    cmdfrmAddNew.Visible = Not Me.NewRecord
    cmdfrmDeleteRecord.Visible = Not Me.NewRecord
    cmdfrmSaveRecord.Visible = Me.NewRecord
    cmdfrmUndo.Visible = Me.NewRecord

End Sub

Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
    Call DoMouseWheel(Me, Count)

End Sub

Private Sub Form_Open(Cancel As Integer)
    gbRunMailChanges = False
   Cancel = FormLogin(glnCashierID, Me)
    If Not Cancel Then

        If Me.OpenArgs <> "" Then
            Me.Description.SetFocus
            Me.Description.Text = Me.OpenArgs
        End If
    End If
End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gbRunMailChanges Then MailChanges False

End Sub
Private Sub cmdSalesdepartmentGroupID_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdSalesdepartmentGroupID_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "SalesDepartmentGroupID"

    stLinkCriteria = "[ID]=" & cmdSalesdepartmentGroupID.Tag


    DoCmd.OpenForm stDocName, , , stLinkCriteria, , acDialog
    Me.Requery
    Me.Refresh


Exit_cmdSalesdepartmentGroupID_Click:
    Exit Sub

Err_cmdSalesdepartmentGroupID_Click:
    MsgBox Err.Description
    Resume Exit_cmdSalesdepartmentGroupID_Click

End Sub
'------------------------------------------------------------
' cmdfrmDeleteRecord_Click
'
''------------------------------------------------------------
'Private Sub cmdfrmDeleteRecord_Click()
'On Error GoTo cmdfrmDeleteRecord_Click_Err
'
'    On Error Resume Next
'    DoCmd.GoToControl Screen.PreviousControl.Name
'    Err.Clear
'    If (Not Form.NewRecord) Then
'        DoCmd.RunCommand acCmdDeleteRecord
'    End If
'    If (Form.NewRecord And Not Form.Dirty) Then
'        Beep
'    End If
'    If (Form.NewRecord And Form.Dirty) Then
'        DoCmd.RunCommand acCmdUndo
'    End If
'    If (MacroError <> 0) Then
'        Beep
'        MsgBox MacroError.Description, vbOKOnly, ""
'    End If
'    Call SetRunMailChangesTrue
'
'
'cmdfrmDeleteRecord_Click_Exit:
'    Exit Sub
'
'cmdfrmDeleteRecord_Click_Err:
'    MsgBox Error$
'    Resume cmdfrmDeleteRecord_Click_Exit
'
'End Sub


'------------------------------------------------------------
' cmdfrmSaveRecord_Click
'
''------------------------------------------------------------
'Private Sub cmdfrmSaveRecord_Click()
'On Error GoTo cmdfrmSaveRecord_Click_Err
'
'    On Error Resume Next
'    DoCmd.RunCommand acCmdSaveRecord
'    If (MacroError <> 0) Then
'        Beep
'        MsgBox MacroError.Description, vbOKOnly, ""
'    End If
'
'
'cmdfrmSaveRecord_Click_Exit:
'    Exit Sub
'
'cmdfrmSaveRecord_Click_Err:
'    MsgBox Error$
'    Resume cmdfrmSaveRecord_Click_Exit
'
'End Sub


'------------------------------------------------------------
' cmdfrmUndo_Click
'
'------------------------------------------------------------
'Private Sub cmdfrmUndo_Click()
'On Error GoTo cmdfrmUndo_Click_Err
'
'    On Error Resume Next
'    DoCmd.RunCommand acCmdUndo
'    If (MacroError <> 0) Then
'        Beep
'        MsgBox MacroError.Description, vbOKOnly, ""
'    End If
'
'
'cmdfrmUndo_Click_Exit:
'    Exit Sub
'
'cmdfrmUndo_Click_Err:
'    MsgBox Error$
'    Resume cmdfrmUndo_Click_Exit
'
'End Sub


'------------------------------------------------------------
' Command0_Click
'
'------------------------------------------------------------
Private Sub Command0_Click()
On Error GoTo Command0_Click_Err

    DoCmd.Close , ""


Command0_Click_Exit:
    Exit Sub

Command0_Click_Err:
    MsgBox Error$
    Resume Command0_Click_Exit

End Sub


'------------------------------------------------------------
' cmdfrmAddNew_Click
'
'------------------------------------------------------------
'Private Sub cmdfrmAddNew_Click()
'On Error GoTo cmdfrmAddNew_Click_Err
'
'    On Error Resume Next
'    DoCmd.GoToRecord , "", acNewRec
'    If (MacroError <> 0) Then
'        Beep
'        MsgBox MacroError.Description, vbOKOnly, ""
'    End If
'
'
'cmdfrmAddNew_Click_Exit:
'    Exit Sub
'
'cmdfrmAddNew_Click_Err:
'    MsgBox Error$
'    Resume cmdfrmAddNew_Click_Exit
'
'End Sub


'------------------------------------------------------------
' cmdPrintPreview_Click
'
'------------------------------------------------------------
Private Sub cmdPrintPreview_Click()
On Error GoTo cmdPrintPreview_Click_Err

    Call PrintPreView(Form)


cmdPrintPreview_Click_Exit:
    Exit Sub

cmdPrintPreview_Click_Err:
    MsgBox Error$
    Resume cmdPrintPreview_Click_Exit

End Sub


'------------------------------------------------------------
' cmdDataSheet_Click
'
'------------------------------------------------------------
Private Sub cmdDataSheet_Click()
On Error GoTo cmdDataSheet_Click_Err

    Call DatasheetView(Form)


cmdDataSheet_Click_Exit:
    Exit Sub

cmdDataSheet_Click_Err:
    MsgBox Error$
    Resume cmdDataSheet_Click_Exit

End Sub


'------------------------------------------------------------
' cmdZeroNegativeOnhands_Click
'
'------------------------------------------------------------
Private Sub cmdZeroNegativeOnhands_Click()
On Error GoTo cmdZeroNegativeOnhands_Click_Err
Dim strSQL
    If cmdZeroNegativeOnhands.Tag > 2 Then
        strSQL = "This will reset the negative onhands for Department # " & cmdZeroNegativeOnhands.Tag & vbCrLf
        strSQL = strSQL & "Are you sure?"
        If MsgBox(strSQL, vbYesNo) = vbYes Then
            strSQL = "UPDATE Item SET Item.Onhand = 0 "
            strSQL = strSQL & "WHERE (((Item.SalesDepartmentID)=" & cmdZeroNegativeOnhands.Tag & " "
            strSQL = strSQL & ")) and (((Item.Onhand)<0))"
            Debug.Print strSQL
            DoCmd.RunSQL strSQL

        End If
    End If


cmdZeroNegativeOnhands_Click_Exit:
    Exit Sub

cmdZeroNegativeOnhands_Click_Err:
    MsgBox Error$
    Resume cmdZeroNegativeOnhands_Click_Exit

End Sub
