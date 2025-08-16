Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'------------------------------------------------------------
' Combo2_AfterUpdate
'
'------------------------------------------------------------
Private Sub Combo2_AfterUpdate()
On Error GoTo Combo2_AfterUpdate_Err

    DoCmd.SearchForRecord , "", acFirst, "[Description] = " & "'" & Screen.ActiveControl & "'"


Combo2_AfterUpdate_Exit:
    Exit Sub

Combo2_AfterUpdate_Err:
    MsgBox Error$
    Resume Combo2_AfterUpdate_Exit

End Sub


'------------------------------------------------------------
' Combo4_AfterUpdate
'
'------------------------------------------------------------
Private Sub Combo4_AfterUpdate()
On Error GoTo Combo4_AfterUpdate_Err

    DoCmd.SearchForRecord , "", acFirst, "[ID] = " & Str(Nz(Screen.ActiveControl, 0))


Combo4_AfterUpdate_Exit:
    Exit Sub

Combo4_AfterUpdate_Err:
    MsgBox Error$
    Resume Combo4_AfterUpdate_Exit

End Sub


Private Sub cmbSalesdepartment_AfterUpdate()
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(Nz(Me![cmbSalesDepartment], 0))
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
    cmdCreateWorkTable.Visible = True
        txtRecordCount.Visible = True

End Sub

Private Sub cmdCreateWorkTable_Click()
On Error GoTo cmdCreateWorkTable_Click_Err

    SetOptions False
    DoCmd.OpenQuery "qryPricingWorkTable_1_Delete", acViewNormal, acEdit
    DoCmd.OpenQuery "qryPricingWorkTable_2_AppendSalesDepartment", acViewNormal, acEdit
    SetOptions True
    DoCmd.Close , ""
    DoCmd.OpenForm "PricingTool", acNormal, "", "", , acNormal

cmdCreateWorkTable_Click_Exit:

    Exit Sub

cmdCreateWorkTable_Click_Err:
    MsgBox Error$
    Resume cmdCreateWorkTable_Click_Exit
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

Private Sub Form_Current()
    cmbSalesDepartment = ID
    On Error GoTo ErrNoRecords
    'keep this debug code.. it is what generates the error when there are no records@@@
    Debug.Print txtRecordCount
    'cmdCreateWorkTable.Caption = "2. Make work table based on these items."
    cmdCreateWorkTable.Visible = True
    txtRecordCount.Visible = True
ExitSub:
    Exit Sub


ErrNoRecords:
    cmdCreateWorkTable.Visible = False
    txtRecordCount.Visible = False
    Resume ExitSub
End Sub
