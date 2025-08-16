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

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(Nz(Me![cboID], 0))
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark

End Sub

Private Sub cmbSalesDepartment_Enter()
    If Me.Dirty Then
        'FORCE SAVE
        Me.Dirty = False
    End If
End Sub

Private Sub cmdAddItems_Click()
Dim nI As Integer
Dim lngItem As Double '07/17/2021
Dim rsLines As Recordset
Dim nLcount As Integer
    If gcfHandleErrors Then On Error GoTo ErrLog
    If IsNumeric(txtAmount) = False Or txtAmount <= 0 Then
        MsgBox "Please enter the Discount Amount"
        txtAmount.SetFocus
        Exit Sub
    End If


    Set rsLines = Me.BuyDownItemsSub.Form.RecordsetClone
    If rsLines.RecordCount = 0 Then
        nLcount = 1
    Else
        rsLines.MoveLast
        nLcount = rsLines!Line + 1
    End If

    With lstSelectBuyDown
        For nI = 0 To .ListCount - 1
            If .Selected(nI) Then
                lngItem = .Column(0, nI)
                With rsLines
                    .AddNew
                    !BuyDownID = ID
                    !Line = nLcount
                    !ItemsCode = lngItem
                    !SpcPrice = txtAmount
                    !SetupDate = Now
                    !ChangeDate = Now
                    .Update
                    nLcount = nLcount + 1
                End With
                .Selected(nI) = False
            End If
        Next nI
        rsLines.Close
        Set rsLines = Nothing
        .Requery

    End With

ExitSub:

    cmbSalesDepartment.SetFocus
    cmdAddItems.Visible = False
    Exit Sub
ErrLog:
    EMsgBox Me, "cmdAddItems"
    Resume ExitSub

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
Dim rs
If gcfHandleErrors Then On Error GoTo ErrLog

    DoCmd.DoMenuItem acFormBar, acEditMenu, acSelectRecord, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, acDelete, , acMenuVer70

    Set rs = Me.Recordset.Clone
    rs.MoveFirst
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark

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

Private Sub cmdListRequery_Click()
    lstSelectBuyDown.Requery

End Sub

Private Sub Form_AfterDelConfirm(Status As Integer)
    cboID.Requery
End Sub

Private Sub Form_AfterInsert()
    cboID.Requery
End Sub

Private Sub Form_AfterUpdate()
    cboID.Requery
    cmdfrmDeleteRecord.Visible = True
    gbRunMailChanges = True

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    ChangeDate = Now
End Sub

Private Sub Form_Current()
    cboID.Visible = Not Me.NewRecord
    'ID.Visible = Me.NewRecord
    cboID = ID

    cmdfrmAddNew.Visible = Not Me.NewRecord
    cmdfrmDeleteRecord.Visible = Not Me.NewRecord
    'cmdfrmSaveRecord.Visible = Me.NewRecord
    'cmdfrmUndo.Visible = Me.NewRecord

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


Private Sub Form_Load()

    Me.InsideHeight = boxScreen.Height + FormFooter.Height + 100
    Me.InsideWidth = boxScreen.Width + 500

End Sub

Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
    Call DoMouseWheel(Me, Count)

End Sub

Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)
    Text27.BackColor = glnColorTotals
    Text29.BackColor = glnColorTotals
    gbRunMailChanges = False
End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gbRunMailChanges Then MailChanges False

End Sub

Private Sub lstSelectBuyDown_AfterUpdate()
    cmdAddItems.Visible = True

End Sub

Private Sub cmdPrintBuydown_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdPrintBuydown_Click

    Dim stDocName As String

    stDocName = "qryBuydownDetail"
    DoCmd.OpenReport stDocName, acPreview

Exit_cmdPrintBuydown_Click:
    Exit Sub

Err_cmdPrintBuydown_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrintBuydown_Click

End Sub
