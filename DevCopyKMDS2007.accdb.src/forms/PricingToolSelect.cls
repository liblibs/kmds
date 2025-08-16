Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



'------------------------------------------------------------
' Combo13_AfterUpdate
'
'------------------------------------------------------------


Private Sub cmbVendorID_AfterUpdate()
On Error GoTo cmbVendorID_AfterUpdate_Err

    DoCmd.SearchForRecord , "", acFirst, "[ID] = " & Str(Nz(Screen.ActiveControl, 0))

cmbVendorID_AfterUpdate_Exit:
    Exit Sub

cmbVendorID_AfterUpdate_Err:
    MsgBox Error$
    Resume cmbVendorID_AfterUpdate_Exit

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
    cmbVendorID = ID
    On Error GoTo ErrNoRecords
    'keep this debug code.. it is what generates the error when there are no records@@@
    Debug.Print [PricingToolSubForm].[Form]![txtRecordCount]
    'cmdCreateWorkTable.Caption = "2. Make work table based on these items."
    cmdCreateWorkTable.Visible = True

ExitSub:
    Exit Sub


ErrNoRecords:
    cmdCreateWorkTable.Visible = False
    Resume ExitSub
End Sub

'------------------------------------------------------------
' cmdCreateWorkTable_Click
'
'------------------------------------------------------------
Private Sub cmdCreateWorkTable_Click()
On Error GoTo cmdCreateWorkTable_Click_Err

    SetOptions False
    DoCmd.OpenQuery "qryPricingWorkTable_1_Delete", acViewNormal, acEdit
    DoCmd.OpenQuery "qryPricingWorkTable_2_Append", acViewNormal, acEdit
    SetOptions True
    DoCmd.Close , ""
    DoCmd.OpenForm "PricingTool", acNormal, "", "", , acNormal

cmdCreateWorkTable_Click_Exit:

    Exit Sub

cmdCreateWorkTable_Click_Err:
    MsgBox Error$
    Resume cmdCreateWorkTable_Click_Exit

End Sub




'------------------------------------------------------------
' cmdSelectBySalesDepartment_Click
'
'------------------------------------------------------------
Private Sub cmdSelectBySalesDepartment_Click()
On Error GoTo cmdSelectBySalesDepartment_Click_Err

    DoCmd.Close , ""
    DoCmd.OpenForm "PricingToolSelectBySalesDepartment", acNormal, "", "", , acNormal



cmdSelectBySalesDepartment_Click_Exit:
    Exit Sub

cmdSelectBySalesDepartment_Click_Err:
    MsgBox Error$
    Resume cmdSelectBySalesDepartment_Click_Exit

End Sub


'------------------------------------------------------------
' Command23_Click
'
'------------------------------------------------------------
Private Sub Command23_Click()
On Error GoTo Command23_Click_Err

    DoCmd.Close , ""
    DoCmd.OpenForm "PricingToolSelectBySalesDepartment", acNormal, "", "", , acNormal


Command23_Click_Exit:
    Exit Sub

Command23_Click_Err:
    MsgBox Error$
    Resume Command23_Click_Exit

End Sub
