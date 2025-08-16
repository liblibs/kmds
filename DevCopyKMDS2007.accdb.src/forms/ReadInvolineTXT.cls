Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command2_Click()
    If gcfHandleErrors Then On Error GoTo ErrLog
    Dim fsoTemp As FileSystemObject
Dim fldrTemp As Folder
Dim filTemp As File
Dim txtstrmTemp As TextStream
Dim fieldTemp As ADODB.Field
Dim strTemp As String
Dim nI As Integer


'Dim rstemp As New ADODB.Recordset
Dim rsInvoLineLD As Object
Set rsInvoLineLD = CreateObject("ADODB.Recordset")


    Open Text0 For Input As #1

    With rsInvoLineLD
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Open "InvoLineLD", , , , adCmdTable
        Do While Not EOF(1)
            .AddNew
            For nI = 0 To 14
                Line Input #1, strTemp
                .Fields(nI) = strTemp
            Next nI
            .Update
        Loop
        .Close
    End With

    Set rsInvoLineLD = Nothing
ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "command2" & vbCrLf & Err.Number & Err.Description
    Close #1
    Resume ExitSub

End Sub
