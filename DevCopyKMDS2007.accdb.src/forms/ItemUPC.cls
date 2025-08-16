Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_AfterDelConfirm(Status As Integer)
    gbRunMailChanges = True

End Sub

Private Sub Form_AfterUpdate()
    gbRunMailChanges = True

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    gbRunMailChanges = True
    ChangeDate = Now
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
Dim strMsg As String
    strMsg = "Form_Error " & vbCrLf
    strMsg = strMsg & "Data Error #" & DataErr & vbCrLf & vbCrLf
    Select Case DataErr
    Case 3146
        strMsg = strMsg & "This indicates a duplicate record!"
    Case 3162
        strMsg = strMsg & "This indicates a blank record"

    Case Else

    End Select

    MsgBox strMsg
    Response = acDataErrContinue
End Sub

Private Sub ID_DblClick(Cancel As Integer)

Dim strNextUpcCode As String
'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")


    If MsgBox("Setup up a CUSTOM barcode for this item?", vbYesNo) = vbYes Then
        rstemp.Open "qryUPCItemMax999999", CurrentProject.Connection, , , adCmdStoredProc

        If rstemp.BOF And rstemp.EOF Then
            strNextUpcCode = CheckDigit11("99999900001")
        Else
            strNextUpcCode = CDbl(Left(rstemp!maxofid, 11) + 1)
            strNextUpcCode = CheckDigit11(strNextUpcCode)
        End If
        rstemp.Close
        Set rstemp = Nothing
        ID = strNextUpcCode
    End If

End Sub
