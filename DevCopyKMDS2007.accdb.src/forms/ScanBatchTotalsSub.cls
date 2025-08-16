Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'------------------------------------------------------------
' cmbItemID_AfterUpdate
'
'------------------------------------------------------------
Private Sub cmbItemID_AfterUpdate()
On Error GoTo cmbItemID_AfterUpdate_Err

    DoCmd.SearchForRecord , "", acFirst, "[ItemID] = " & Str(Nz(Screen.ActiveControl, 0))

cmbItemID_AfterUpdate_Exit:
    Exit Sub

cmbItemID_AfterUpdate_Err:
    MsgBox Error$
    Resume cmbItemID_AfterUpdate_Exit

End Sub


Private Sub Form_Current()
    cmbItemID = ItemID

End Sub
