Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub AgentNumber_DblClick(Cancel As Integer)
    UpdateAgentNumberTrue
End Sub

Private Sub FirstName_DblClick(Cancel As Integer)
    UpdateAgentNumberTrue
End Sub

Private Sub Form_Load()
    Me.DatasheetBackColor = Detail.BackColor
    Me.DatasheetAlternateBackColor = Detail.BackColor
End Sub

Private Sub LastName_DblClick(Cancel As Integer)
    UpdateAgentNumberTrue
End Sub

Private Sub UpdateAgentNumberTrue()

Dim intSavePosition As Integer

    'save position to restore after requery (prevents pointer from going to top of screen)
    If gcfHandleErrors Then On Error GoTo ErrLog

    intSavePosition = Me.Recordset.AbsolutePosition

    strSQL = "UPDATE Agents "
    strSQL = strSQL & "SET Agents.ActiveForTransfer = True "
    strSQL = strSQL & " WHERE (((AgentNumber)=" & AgentNumber & "))"


    DoCmd.RunSQL strSQL, False

    Forms![AgentsListBox]![AgentsActiveForTransferTrue].Form.Requery
    Forms![AgentsListBox]![AgentsActiveForTransferFalse].Form.Requery

    Me.Recordset.AbsolutePosition = intSavePosition
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "UpdateAgentNumberTrue"
    Resume ExitSub
End Sub
