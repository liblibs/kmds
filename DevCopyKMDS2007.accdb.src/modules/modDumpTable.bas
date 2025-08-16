Attribute VB_Name = "modDumpTable"
Option Compare Database
Option Explicit

Public Sub DumpSQLTables()
'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

If gcfHandleErrors Then On Error GoTo ErrLog

     'delete original
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qrySQLTables"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        Set rstemp = .Execute
        Do While Not rstemp.EOF
            ExportTable rstemp!TableName
            rstemp.MoveNext
        Loop
        rstemp.Close
        Set rstemp = Nothing
    End With
    Set comQDF = Nothing

ExitSub:
    Exit Sub
ErrLog:
    fMsgBox "DumpSQLTables"
    EOLog "Error #" & Err.Number & " " & Err.Description & " DumpSQLTables"
    Resume ExitSub

End Sub

Public Sub ExportTable(TableName As String)

Dim fso As New FileSystemObject

Dim strFolderName As String

If gcfHandleErrors Then On Error GoTo ErrLog

    strFolderName = "C:\KMDS\Backup"

    If Not fso.FolderExists(strFolderName) Then
        fso.CreateFolder strFolderName
    End If

    strFolderName = "C:\KMDS\Backup\" & Format(Now, "dd")

    If Not fso.FolderExists(strFolderName) Then
        fso.CreateFolder strFolderName
    End If

    strFolderName = strFolderName & "\" & TableName & ".txt"


    EOLog "  " & TableName & " to " & strFolderName
    DoCmd.TransferText acExportDelim, , TableName, strFolderName, True

ExitSub:
    Exit Sub

ErrLog:
    fMsgBox "ExportTable"
    EOLog "Error #" & Err.Number & " " & Err.Description & " ExportTable"
    Resume ExitSub
End Sub

Public Sub MakeSQLBackup()
Dim strPassString

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

If gcfHandleErrors Then On Error GoTo ErrLog
     'delete original
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "Select CashsysInsLoc from Sysfile"
        .CommandType = adCmdText
        'ADO Numeric Datatypes are very particular
        Set rstemp = .Execute
        'If rsTemp.BOF Then
            rstemp.MoveFirst
            strPassString = "c:\kmds\kmdsDUMP.bat " & rstemp!CashsysInsLoc & " " & Format(Now, "yyyy-dd-mm$hh-nn")
            WaitForExecCmd (strPassString)
        'End If
    End With
    rstemp.Close
    Set rstemp = Nothing
    Set comQDF = Nothing

ExitSub:
    Exit Sub


ErrLog:
    fMsgBox "MakeSQLBackup"
    Resume ExitSub
End Sub


Public Sub BackUpSQL()
Dim strPassString
'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

If gcfHandleErrors Then On Error GoTo ErrLog
     'delete original
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "Select CashsysInsLoc from Sysfile"
        .CommandType = adCmdText
        'ADO Numeric Datatypes are very particular
        Set rstemp = .Execute
        'If rsTemp.BOF Then
            rstemp.MoveFirst
            strPassString = "c:\kmds\kmdsback.bat " & rstemp!CashsysInsLoc & " " & Format(Now, "yyyy-dd-mm$hh-nn")
            WaitForExecCmd (strPassString)
        'End If
    End With
    rstemp.Close
    Set rstemp = Nothing
    Set comQDF = Nothing

ExitSub:
    Exit Sub


ErrLog:
    fMsgBox "BackUpSQL"
    Resume ExitSub
End Sub


Public Sub RestoreSQL()
Dim strPassString
'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

If gcfHandleErrors Then On Error GoTo ErrLog
     'delete original
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "Select CashsysInsLoc from Sysfile"
        .CommandType = adCmdText
        'ADO Numeric Datatypes are very particular
        Set rstemp = .Execute
        'If rsTemp.BOF Then
            rstemp.MoveFirst
            strPassString = "c:\kmds\kmdsRest.bat " & rstemp!CashsysInsLoc
            Shell strPassString, vbNormalFocus

        'End If
    End With
    rstemp.Close
    Set rstemp = Nothing
    Set comQDF = Nothing

ExitSub:
    Exit Sub

ErrLog:
    fMsgBox "BackUpSQL"
    Resume ExitSub
End Sub

Public Function MailSQLBackup()

Dim strPassString

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim strSourceFileName As String
Dim strTargetFileName As String
Dim bCopy As Boolean

    'create backup of sql database!
    BackUpSQL
    EOLog "Backup Complete"

    If gcfHandleErrors Then On Error GoTo Err_Command1_Click
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "Select CashsysInsLoc from Sysfile"
        .CommandType = adCmdText
        Set rstemp = .Execute
        rstemp.MoveFirst
        strPassString = rstemp!CashsysInsLoc
        strSourceFileName = "c:\kmds\backup\KMDSBACK.zip"
        '\\%1\c\kmds\backup\sql
    End With
    rstemp.Close
    Set rstemp = Nothing
    Set rstemp = CreateObject("ADODB.Recordset")
    Set comQDF = Nothing

    rstemp.Open "qryCashRegisters", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

    With rstemp
        Do While Not .EOF
            bCopy = False
            Select Case !ID
            Case 1
                bCopy = !Cr1Installed
            Case 2
                bCopy = !Cr2Installed
            Case 3
                bCopy = !Cr3Installed
            Case 4
                bCopy = !Cr4Installed
            End Select
            If bCopy = True Then
                strTargetFileName = "\\" & !ComputerName & "\kmds\backup\KMDSSQL.ZIP"
                If !ComputerName = gstrComputerName Then
                    EOLog "SkipCopying " & strSourceFileName & " to " & strTargetFileName
                Else
                    EOLog "Copying " & strSourceFileName & " to " & strTargetFileName
                    FileCopy strSourceFileName, strTargetFileName
                End If
            End If
            .MoveNext
        Loop
        .Close
    End With

Exit_Command1_Click:
    SysCmd acSysCmdClearStatus
    Exit Function

Err_Command1_Click:
    fMsgBox "MailSQLBackup" & vbCrLf & Err.Number & " " & Err.Description

    Select Case Err.Number
    Case 76
        Resume Next
    Case Else
        Resume Exit_Command1_Click
    End Select

End Function



Public Sub FetchJournals()



'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

    Dim strSource As String
    Dim strshell As String

    Dim bFetch As Boolean

    Dim strTarget As String
    strTarget = gstrMailDir & "Alljournals\"

    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "qryCashRegisters"
        Set rstemp = .Execute
    End With

    With rstemp
        Do While Not .EOF
            bFetch = False
            Select Case !ID
            Case 1
                If !Cr1Installed Then
                    bFetch = True
                End If
            Case 2
                If !Cr2Installed Then
                    bFetch = True
                End If
            Case 3
                If !Cr3Installed Then
                    bFetch = True
                End If
            Case 4
                If !Cr4Installed Then
                    bFetch = True
                End If
            Case 5
                If !Cr5Installed Then
                    bFetch = True
                End If
            Case Else
            End Select
            If bFetch = True Then
                strSource = "\\" & !ComputerName & "\KMDS\Journal\CR" & !ID & "*.txt"

                If Dir(gstrMailDir & "Alljournals", vbDirectory) = "" Then
                    MkDir (gstrMailDir & "Alljournals")
                End If
                strshell = "XCOPY " & strSource & " " & strTarget & " /M"
                WaitForExecCmd (strshell)
            End If
            .MoveNext
        Loop
        .Close
    End With
    Set rstemp = Nothing
    Set comQDF = Nothing

    strshell = "C:\KMDS\DIRDATER.EXE " & strTarget
    WaitForExecCmd (strshell)

End Sub
