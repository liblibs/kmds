Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub cmbCashier_AfterUpdate()
    Me.Refresh
    Me.CashiersAndProgramsSwitchBoardSub.Requery
End Sub

Private Sub List10_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ItemText] = '" & Me![List10] & "'"
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
End Sub
