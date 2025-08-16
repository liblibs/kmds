Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmbDesc_AfterUpdate()
    LookupItem (cmbDesc)
End Sub

'Private Sub cmbItemID_afterUpdate(Cancel As Integer)
 '   LookupItem cmbItemID

'End Sub

Private Sub cmbUPC_AfterUpdate()
    LookupItem cmbUPC
End Sub

Private Sub LookupItem(ItemID As Long)
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & ItemID
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark

End Sub

Private Sub Form_Current()
    If Me.NewRecord Then
    Else
        cmbUPC = ID
        cmbDesc = ID
        cmbItemID = ID

    End If

End Sub
