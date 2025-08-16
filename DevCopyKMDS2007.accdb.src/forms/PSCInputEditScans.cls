Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cmdSetupNotOnFile_Click()
    If MsgBox("Setup the 'Not on File' Items?", vbYesNo) = vbYes Then
        'set them up!
    End If
End Sub
Private Sub cmdClubOrderCreate_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdClubOrderCreate_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "ClubOrderCreate"
    DoCmd.Close

    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , "PDA"

Exit_cmdClubOrderCreate_Click:
    Exit Sub

Err_cmdClubOrderCreate_Click:
    MsgBox Err.Description
    Resume Exit_cmdClubOrderCreate_Click

End Sub
Private Sub cmdAudit_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdAudit_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    AuditFileAppendByPDA
    stDocName = "Audit"
    'delete the work file
    SetOptions False
    DoCmd.OpenQuery "qryPSCInput_Step1_DeleteRecords"
    SetOptions True

    DoCmd.Close
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdAudit_Click:
    Exit Sub

Err_cmdAudit_Click:
    MsgBox Err.Description
    Resume Exit_cmdAudit_Click

End Sub

Private Sub Form_Load()
    cmdAudit.Visible = gbAllowChangeOnhand
End Sub
Private Sub Command7_Click()
If gcfHandleErrors Then On Error GoTo Err_Command7_Click


    DoCmd.Close

Exit_Command7_Click:
    Exit Sub

Err_Command7_Click:
    MsgBox Err.Description
    Resume Exit_Command7_Click

End Sub
