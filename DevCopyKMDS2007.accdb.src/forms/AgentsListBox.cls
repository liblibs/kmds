Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub


Private Sub Form_Load()
    SetOptions False
    Me.InsideHeight = boxScreen.Height + 100
    Me.InsideWidth = boxScreen.Width + 500

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetOptions True
End Sub

Private Sub List0_BeforeUpdate(Cancel As Integer)

End Sub

Private Sub List0_DblClick(Cancel As Integer)
   strSQL = "UPDATE Agents "
    strSQL = strSQL & "SET Agents.ActiveForTransfer = False "
    strSQL = strSQL & " WHERE (((AgentNumber)=" & CLng(List0.ItemData(List0.ListIndex)) & "))"


    DoCmd.RunSQL strSQL, False
    'MsgBox CLng(List2.ItemData(List2.ListIndex))
    List2.Requery
    List0.Requery
    Me.Refresh

End Sub

Private Sub List2_DblClick(Cancel As Integer)
    strSQL = "UPDATE Agents "
    strSQL = strSQL & "SET Agents.ActiveForTransfer = True "
    strSQL = strSQL & " WHERE (((AgentNumber)=" & CLng(List2.ItemData(List2.ListIndex)) & "))"


    DoCmd.RunSQL strSQL, False
    'MsgBox CLng(List2.ItemData(List2.ListIndex))
    List2.Requery
    List0.Requery
    Me.Refresh

End Sub
