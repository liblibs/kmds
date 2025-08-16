Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub Form_Close()
    FormLogout glnCashierID, Me

End Sub

Private Sub Form_Open(Cancel As Integer)
    '2006-12-5 'created crview.ins
    '               used by crview.exe for valid registers
    '10/03/2004 added "confirm" Parameter
    'Dim rstemp As New ADODB.Recordset

    Dim rstemp As Object
    Set rstemp = CreateObject("ADODB.Recordset")



    Dim bAtLeastOneRegisterCheckedOut
    Dim strValidRegisters As String

    Dim fso As FileSystemObject
    Dim ts As TextStream
    If gcfHandleErrors Then On Error GoTo ExitSub

    'Cancel = FormLogin(glnCashierID, Me)

    Me.Repaint
    gbRunMailChanges = False

    bAtLeastOneRegisterCheckedOut = False
    rstemp.Open "qryCashRegisters", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    With rstemp
        Do While Not .EOF
            Select Case !ID
            Case 1
                If !Cr1CheckOut = True Then bAtLeastOneRegisterCheckedOut = True
            Case 2
                If !Cr2CheckOut = True Then bAtLeastOneRegisterCheckedOut = True
            Case 3
                If !Cr3CheckOut = True Then bAtLeastOneRegisterCheckedOut = True
            Case 4
                If !Cr4CheckOut = True Then bAtLeastOneRegisterCheckedOut = True
            Case 5
                If !Cr5CheckOut = True Then bAtLeastOneRegisterCheckedOut = True
            End Select
            .MoveNext
        Loop
        .Close
    End With
    Set rstemp = Nothing

    If bAtLeastOneRegisterCheckedOut Then
        EndOFDayStream
        'added 04-07-2009
        'If Get_User_Name = True Then
            Shell "c:\kmds\kmdslogout.bat", vbNormalFocus
            'Shell "shutdown -l", vbNormalFocus
            DoCmd.Quit 'Exit Sub
        'End If
    Else
        MsgBox "There are no registers checked out!"
        Cancel = True
        Exit Sub
    End If

ExitSub:

    Cancel = True
    DoCmd.Quit
    Exit Sub

ErrLog:
    MsgBox "Error end of day" & vbCrLf & "error #" & Err.Number & " " & Err.Description
    Resume ExitSub

End Sub
