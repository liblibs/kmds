Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmbItemByDesc_AfterUpdate()
    cmbItemByUPC_AfterUpdate
End Sub

Private Sub cmbItemByUPC_AfterUpdate()
    Price = cmbItemByUPC.Column(5)
End Sub

Private Sub Form_AfterInsert()
    If txtSumLiquor > Me.Parent("txtLiquorValue") Then
        MsgBox "Over!"
    End If
End Sub

Private Sub Form_AfterUpdate()
    If gcfHandleErrors Then On Error GoTo ErrLog
    gbRunMailChanges = True
    Me.Recalc
    If txtSumLiquor > Me.Parent("Price") Then
        MsgBox "CHANGING Kit Price BECAUSE LIQUOR AMOUNT IS GREATER THAN Kit Price!", vbCritical, "Selling Price is too low!"
        Me.Parent("Price") = txtSumLiquor
        Me.Parent.Dirty = False
    End If
    DoCmd.GoToRecord , , acNewRec

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_AfterUpdate"
    Resume ExitSub

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
'    MsgBox DataErr & " " & Response & " " & Err.Number & " " & Err.Description
    Select Case DataErr
    Case 3162  'null
        Me.Undo
    Case 2237   ' not in list
        Me.Undo
    Case Else
    End Select


End Sub
