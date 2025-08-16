Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdQueryName_AfterUpdate()
Dim qdfTemp As QueryDef
    Set qdfTemp = CurrentDb().QueryDefs(cmdQueryName)
    txtSQL = qdfTemp.SQL
    qdfTemp.Close
    Set qdfTemp = Nothing
End Sub

Private Sub Command13_Click()
If gcfHandleErrors Then On Error GoTo Err_Command13_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = cmbForm
    DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria

Exit_Command13_Click:
    Exit Sub

Err_Command13_Click:
    MsgBox Err.Description
    Resume Exit_Command13_Click

End Sub

Private Sub Command2_Click()
If gcfHandleErrors Then On Error GoTo Err_Command2_Click

    Dim stDocName As String

    stDocName = cmdQueryName
    DoCmd.OpenQuery stDocName, acNormal, acEdit

Exit_Command2_Click:
    Exit Sub

Err_Command2_Click:
    MsgBox Err.Description
    Resume Exit_Command2_Click

End Sub

Private Sub Command7_Click()
Dim qdfTemp As QueryDef

      DoCmd.OpenQuery txtSQL

End Sub

Private Sub Command8_Click()
    DoCmd.RunCommand txtSQL

End Sub

Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)
    If Not Cancel Then
        If InputBox("DOB:") <> "07271956" Then
            Cancel = True
            Exit Sub
        End If
        If InputBox("ODL:") <> "3112245" Then
            Cancel = True
            Exit Sub
        End If
    End If
End Sub


Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub

Private Sub Command5_Click()
If gcfHandleErrors Then On Error GoTo Err_Command5_Click

    Dim stDocName As String


    DoCmd.RunSQL txtSQL

Exit_Command5_Click:
    Exit Sub

Err_Command5_Click:
    MsgBox Err.Description
    Resume Exit_Command5_Click

End Sub


Private Sub Command12_Click()
If gcfHandleErrors Then On Error GoTo Err_Command12_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = cmbForm
    DoCmd.OpenForm stDocName, acFormDS, , stLinkCriteria

Exit_Command12_Click:
    Exit Sub

Err_Command12_Click:
    MsgBox Err.Description
    Resume Exit_Command12_Click

End Sub
Private Sub Command14_Click()
    SQLPassThru txtSQL


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
Private Sub cmdMacRun_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdMacRun_Click

    Dim stDocName As String

    stDocName = txtSQL '"macMailSQLBackup"
    DoCmd.RunMacro stDocName

Exit_cmdMacRun_Click:
    Exit Sub

Err_cmdMacRun_Click:
    MsgBox Err.Description
    Resume Exit_cmdMacRun_Click

End Sub
