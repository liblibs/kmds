Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdShowLines_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdShowLines_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "AgentsPreviousOrderWithLines"

    stLinkCriteria = "[InvoHeadID]=" & Me![InvoHeadID]
    DoCmd.OpenForm stDocName, acFormDS, , stLinkCriteria, acFormReadOnly

Exit_cmdShowLines_Click:
    Exit Sub

Err_cmdShowLines_Click:
    MsgBox Err.Description
    Resume Exit_cmdShowLines_Click

End Sub

Private Sub InvoHeadID_Click()
cmdShowLines_Click
End Sub
