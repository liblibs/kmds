Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmbID_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(Nz(Me![cmbID], 0))
    If Not rs.EOF Then
        Me.Bookmark = rs.Bookmark
        TPR = True
        Me.Dirty = False
'        Me.SubItemsTPRSub.Form.Filter = "((Not Item.TPR=0))"
'        Me.SubItemsTPRSub.Form.FilterOn = True
'        Me.SubItemsTPRSub.Form.Requery
    End If
End Sub
Private Sub cmdClose_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdClose_Click


    DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click

End Sub
Private Sub Form_AfterUpdate()
    If gcfHandleErrors Then On Error Resume Next
        Me.SubItemsTPRSub.Form.Filter = "((Not Item.TPR=0))"
        Me.SubItemsTPRSub.Form.FilterOn = True
        Me.SubItemsTPRSub.Form.Requery

End Sub

Private Sub Form_Current()
    cmbID = ID
End Sub

Private Sub Form_Load()
    Form_AfterUpdate
End Sub
