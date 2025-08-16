Attribute VB_Name = "modRegisterReports"
Option Compare Database
Option Explicit

Public Sub GetAllRegistersDetailFiles()
On Error GoTo ErrLog

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")


    If gcfHandleErrors Then On Error GoTo ErrLog
    'Delete records in the work table
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "Delete * from InvoLineAllRegisters"
        .CommandType = adCmdText
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    With rstemp
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic

        .Open "qryRegistersInstalledAndOnLine", , , , adCmdStoredProc
        Do While Not .EOF
            Select Case !ID
            Case 1
                If !Cr1Installed And !Cr1Online Then
                    ImportTextFile !ComputerName
                End If
            Case 2
                If !Cr2Installed And !Cr2Online Then
                    ImportTextFile !ComputerName
                End If
            Case 3
                If !Cr3Installed And !Cr3Online Then
                    ImportTextFile !ComputerName
                End If
            Case 4
                If !Cr4Installed And !Cr4Online Then
                    ImportTextFile !ComputerName
                End If
            Case 5
                If !Cr5Installed And !Cr5Online Then
                    ImportTextFile !ComputerName
                End If
            Case Else
            End Select
            .MoveNext
        Loop
        .Close
    End With

ExitSub:
    Set rstemp = Nothing
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "RentrakExport"
    Resume ExitSub
End Sub


Public Sub ImportTextFile(ComputerName As String)
On Error GoTo ErrLog

Dim fsoTemp As FileSystemObject
Dim fldrTemp As Folder
Dim filTemp As File
Dim txtstrmTemp As TextStream

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

Dim fieldTemp As ADODB.Field
'Dim fieldTemp As Object
'Set fieldTemp = CreateObject("ADODB.FIELD")




Dim strFname As String


    If gcfHandleErrors Then On Error GoTo ErrLog

    strFname = "\\" & ComputerName & "\KMDS\CRDetail\InvoLine.Txt"

    Set fsoTemp = CreateObject("Scripting.FileSystemObject")
    If fsoTemp.FileExists(strFname) Then
        Set txtstrmTemp = fsoTemp.OpenTextFile(strFname, ForReading, False, TristateFalse)

        With rstemp
            Set .ActiveConnection = CurrentProject.Connection
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockPessimistic
            .Open "InvoLineAllRegisters", , , , adCmdTable


            Do While txtstrmTemp.AtEndOfStream <> True
    '            ''If not gcfHandleErrors then Debug.Print !Code
                .AddNew
                For Each fieldTemp In .Fields
                    fieldTemp = Nz(txtstrmTemp.ReadLine, 0)
                    ''If not gcfHandleErrors then Debug.Print fieldTemp.Name; " "; fieldTemp.Value
                Next
                .Update

                '.MoveNext
            Loop
            txtstrmTemp.Close
            Set txtstrmTemp = Nothing


        End With
        Set rstemp = Nothing
    End If
    Set fsoTemp = Nothing

ExitSub:
    Exit Sub
ErrLog:

    MsgBox "Error #" & Err.Number & " " & Err.Description, , "RentrakExport"
    Resume ExitSub

End Sub
