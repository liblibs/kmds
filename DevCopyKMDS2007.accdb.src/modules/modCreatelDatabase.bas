Attribute VB_Name = "modCreatelDatabase"
Option Compare Database
Option Explicit
Dim mstrDB As String
Dim mstrDBSource As String



Public Function CreateNewLocalDatabase() As Boolean

'    Dim catSource As New ADOX.Catalog
'    Dim catTarget As New ADOX.Catalog
'    Dim tblSource As New ADOX.Table
'    Dim tblTarget As New ADOX.Table
'    Dim colTemp As New ADOX.Column
'    Dim idxSource As New ADOX.Index
'    Dim idxTarget As New ADOX.Index
'    Dim fldTemp As ADODB.Field
'    Dim propTemp As ADOX.Property

    AppendUPCItemsWithItemID

    'Dim comQDF As New Command
    Dim comQDF As Object
    Set comQDF = CreateObject("ADODB.COMMAND")

    'Dim rstemp As New ADODB.Recordset
    Dim rstemp As Object
    Set rstemp = CreateObject("ADODB.Recordset")


    Dim strSQL As String


    Dim stDocName As String
    Dim varReturn As Variant

    If gcfHandleErrors Then On Error GoTo ErrLog

    mstrDB = "C:\KMDS\NewDB.mdb"
    mstrDBSource = "C:\KMDS\NewDBTemplate.mdb"
    varReturn = SysCmd(acSysCmdSetStatus, "Creating " & mstrDB)

    FileCopy mstrDBSource, mstrDB
    'Set wsp = DBEngine(0)
    'Set dbs = wsp.CreateDatabase("Mydb")

    varReturn = SysCmd(acSysCmdSetStatus, "Deleting " & mstrDB)
    'If Dir(mstrDB) <> "" Then Kill (mstrDB)
'    catTarget.Create "Provider = Microsoft.Jet.OLEDB.4.0;" _
'       & "Data Source = " & mstrDB

'    catTarget.Create "Provider = Microsoft.ACE.OLEDB.12.0;" _
'       & "Data Source = " & mstrDB

'    catSource.ActiveConnection = CurrentProject.Connection

'    stDocName = "qryCRObjects"
    strSQL = "SELECT CRObjects.* FROM CRObjects "
    strSQL = strSQL & "ORDER BY CRObjects.Type DESC , "
    strSQL = strSQL & "         CRObjects.Name;"


    rstemp.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    With rstemp
        .MoveLast
        .MoveFirst

        Do While Not .EOF
            If gcfHandleErrors Then On Error GoTo ErrLogName


            varReturn = SysCmd(acSysCmdInitMeter, "Creating object: " & .AbsolutePosition & " of " & .RecordCount, .RecordCount)
            varReturn = SysCmd(acSysCmdUpdateMeter, .AbsolutePosition)
            Forms("MailChangesForm").lblProcess.Caption = "Creating object: " & .AbsolutePosition & " of " & .RecordCount & " " & rstemp!Name
            Forms("MailChangesForm").Repaint
            Select Case !Type
            Case "Table"
                'varReturn = SysCmd(acSysCmdSetStatus, "Deleting Records from z" & rsTemp!Name)
                '1. delete the records in the temporary work table
                With comQDF
                    .ActiveConnection = CurrentProject.Connection
                    .CommandType = adCmdText
                    .CommandText = "Delete * from z" & rstemp!Name
                    .Execute
                End With
                Set comQDF = Nothing
                Set comQDF = CreateObject("ADODB.COMMAND")

                'varReturn = SysCmd(acSysCmdSetStatus, "Inserting Records from " & rsTemp!Name)
                '2. copy records from SQL Table to Access Table
                If rstemp!Name <> "InvoLine" And rstemp!Name <> "InvoHead" Then
                    With comQDF
                        .ActiveConnection = CurrentProject.Connection
                        strSQL = "INSERT INTO z" & rstemp!Name
                        strSQL = strSQL & " SELECT * "
                        strSQL = strSQL & " FROM [" & rstemp!Name & "]"
                        .CommandType = adCmdText

                        .CommandText = strSQL
                        .Execute
                    End With
                    Set comQDF = Nothing
                    Set comQDF = CreateObject("ADODB.COMMAND")
                End If

                If UCase(rstemp!Name) = "SalesPictures" Then
                    'KEEP THE PICTURES IN THE LOCAL DRIVE TO MINIMIZE NETWORK TRAFFIC
                    With comQDF
                        .ActiveConnection = CurrentProject.Connection
                        strSQL = "UPDATE zSalesPictures SET zSalesPictures.FileName = "
                        strSQL = strSQL & "Replace([filename], '\\" & gstrCashsysInsLoc & "','C:')"
                        .CommandType = adCmdText

                        .CommandText = strSQL
                        .Execute
                    End With
                    Set comQDF = Nothing
                    Set comQDF = CreateObject("ADODB.COMMAND")
                End If

                'DoCmd.TransferDatabase acExport, "Microsoft Access", mstrDB, acTable, !Name, !Name & "Foreign", False
                'varReturn = SysCmd(acSysCmdSetStatus, "Table Transfer: " & rsTemp!Name) & " to " & mstrDB
                '3. Use the TransferDatabase command to populate Register MDB
                DoCmd.TransferDatabase acExport, "Microsoft Access", mstrDB, acTable, "z" & !Name, !Name, False

                '4. delete the records in the temporary work table to reduce size
                With comQDF
                    .ActiveConnection = CurrentProject.Connection
                    .CommandType = adCmdText
                    .CommandText = "Delete * from z" & rstemp!Name
                    .Execute
                End With
                Set comQDF = Nothing
                Set comQDF = CreateObject("ADODB.COMMAND")

            Case "Query"
                'varReturn = SysCmd(acSysCmdSetStatus, "Query Transfer " & rsTemp!Name) & " to " & mstrDB

                DoCmd.TransferDatabase acExport, "Microsoft Access", mstrDB, acQuery, !Name, !Name, True
            Case Else
                MsgBox "unknown type : " & !Type
            End Select
            .MoveNext
        Loop

'        .MoveFirst
'        Do While Not .EOF
'            Select Case !Type
'            Case "Table"
'                For Each tblSource In catSource.Tables
'                    If tblSource.Name = !Name Then Exit For
'                Next tblSource
'
'                Set tblTarget = Nothing
'
'                tblTarget.Name = !Name
'                For Each colTemp In tblSource.Columns
'                    tblTarget.Columns.Append colTemp.Name, colTemp.Type, colTemp.DefinedSize
'                    'tblTarget.Columns(colTemp.Name).Properties("Allow Zero Length") = True
'                    For Each propTemp In colTemp.Properties
'                        If not gcfHandleErrors then Debug.Print propTemp.Name; propTemp.Type; propTemp.Value; propTemp.Attributes
''                        tblTarget.Columns(colTemp.Name).Properties.Name = "Nullable"
'
'                    Next propTemp
'                    'tblTarget.Columns.Item (colTemp)
'                Next colTemp
'
'                For Each idxSource In tblSource.Indexes
'                    Set idxTarget = Nothing
'
'                    idxTarget.Name = idxSource.Name
'                    idxTarget.PrimaryKey = idxSource.PrimaryKey
'                    idxTarget.Unique = idxSource.Unique
'
'                    For Each colTemp In idxSource.Columns
'                        idxTarget.Columns.Append colTemp.Name ', colTemp.Type, colTemp.DefinedSize
'                    Next colTemp
'                    tblTarget.Indexes.Append idxTarget
'
'                Next idxSource
'
'
'                catTarget.Tables.Append tblTarget
'
'                For Each colTemp In tblTarget.Columns
'                    'tblTarget.Columns(colTemp.Name).Properties("Allow Zero Length") = True
'                    For Each propTemp In colTemp.Properties
'                        If not gcfHandleErrors then Debug.Print propTemp.Name; propTemp.Type; propTemp.Value; propTemp.Attributes
'                        tblTarget.Columns(colTemp.Name).Properties("Nullable") = True
'                        'tblTarget.Columns(colTemp.Name).Properties("Allow Zero Length") = True
'                    Next propTemp
'                    'tblTarget.Columns.Item (colTemp)
'                Next colTemp
'
'
'            Case "Query"
'            Case Else
'                MsgBox "unknown type : " & !Type
'            End Select
'            .MoveNext
'        Loop

    End With

    rstemp.Close

    Set rstemp = Nothing
    CurrentProject.Connection.Close
    CreateNewLocalDatabase = True

ExitFunction:
    varReturn = SysCmd(acSysCmdClearStatus)
        varReturn = SysCmd(acSysCmdSetStatus, " ")
    Exit Function

ErrLogName:
    MsgBox rstemp!Name & vbCrLf & mstrDB & vbCrLf & Err.Number & " " & Err.Description, , "CreateNewDataBase " & mstrDB
    CreateNewLocalDatabase = False
    Resume ExitFunction

ErrLog:

    MsgBox mstrDB & " " & Err.Number & " " & Err.Description, , "CreateNewDataBase " & mstrDB
    CreateNewLocalDatabase = False
    Resume ExitFunction
End Function



Public Function MailChanges(Confirm As Boolean)
    '2006-12-5 'created crview.ins
    '               used by crview.exe for valid registers
    '10/03/2004 added "confirm" Parameter
    '02/11/2009 added cashsys.ins copy

    'Dim rstemp As New ADODB.Recordset
    Dim rstemp As Object
    Set rstemp = CreateObject("ADODB.Recordset")

    Dim bAtLeastOneRegisterInstalled
    Dim strValidRegisters As String

    Dim fso As FileSystemObject
    Dim ts As TextStream

    If gcfHandleErrors Then On Error GoTo Err_Command1_Click
    gbRunMailChanges = False
    bAtLeastOneRegisterInstalled = False
    rstemp.Open "qryCashRegisters", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     With rstemp
        Do While Not .EOF
            Select Case !ID
            Case 1
                If !Cr1Installed = True Then bAtLeastOneRegisterInstalled = True
            Case 2
                If !Cr2Installed = True Then bAtLeastOneRegisterInstalled = True
            Case 3
                If !Cr3Installed = True Then bAtLeastOneRegisterInstalled = True
            Case 4
                If !Cr4Installed = True Then bAtLeastOneRegisterInstalled = True
            Case 5
                If !Cr5Installed = True Then bAtLeastOneRegisterInstalled = True
            End Select
            .MoveNext
        Loop
        .MoveFirst
    End With

    If bAtLeastOneRegisterInstalled = False Then
        'MsgBox ("There are no registers defined")
        Exit Function
    End If


    If Confirm Then
        If MsgBox("Do you want to send the changes to the Cash Register now?", vbYesNo) = vbNo Then
            Exit Function
        End If
    End If
    DoCmd.OpenForm "MailChangesForm"
    Forms("MailChangesForm").Repaint

'    Set rsTemp = CurrentDb().OpenRecordset("qryCashRegisters")


    If CreateNewLocalDatabase = True Then
        With rstemp
            Do While Not .EOF
                Select Case !ID
                Case 1
                    Forms("MailChangesForm").lblProcess.Caption = "Copying " & mstrDB & " to " & "\\" & !ComputerName & "\KMDS\Mail\In\KMDSLocal.mdb"
                    Forms("MailChangesForm").Repaint
                    SysCmd acSysCmdSetStatus, "Copying " & mstrDB & " to " & "\\" & !ComputerName & "\KMDS\Mail\In\KMDSLocal.mdb"
                    If !Cr1Installed = True Then
                        FileCopy mstrDB, "\\" & !ComputerName & "\KMDS\Mail\In\KMDSLocal.mdb"
                        FileCopy mstrDB, "\\" & !ComputerName & "\KMDS\CASHSYS.INS"
                        strValidRegisters = strValidRegisters & !ComputerName & vbCrLf
                    End If
                Case 2
                    Forms("MailChangesForm").lblProcess.Caption = "Copying " & mstrDB & " to " & "\\" & !ComputerName & "\KMDS\Mail\In\KMDSLocal.mdb"
                    Forms("MailChangesForm").Repaint
                    SysCmd acSysCmdSetStatus, "Copying " & mstrDB & " to " & "\\" & !ComputerName & "\KMDS\Mail\In\KMDSLocal.mdb"
                    If !Cr2Installed = True Then
                        FileCopy mstrDB, "\\" & !ComputerName & "\KMDS\Mail\In\KMDSLocal.mdb"
                        FileCopy mstrDB, "\\" & !ComputerName & "\KMDS\CASHSYS.INS"
                        strValidRegisters = strValidRegisters & !ComputerName & vbCrLf
                    End If
                Case 3
                    Forms("MailChangesForm").lblProcess.Caption = "Copying " & mstrDB & " to " & "\\" & !ComputerName & "\KMDS\Mail\In\KMDSLocal.mdb"
                    Forms("MailChangesForm").Repaint
                    SysCmd acSysCmdSetStatus, "Copying " & mstrDB & " to " & "\\" & !ComputerName & "\KMDS\Mail\In\KMDSLocal.mdb"
                    If !Cr3Installed = True Then
                        FileCopy mstrDB, "\\" & !ComputerName & "\KMDS\Mail\In\KMDSLocal.mdb"
                        FileCopy mstrDB, "\\" & !ComputerName & "\KMDS\CASHSYS.INS"
                        strValidRegisters = strValidRegisters & !ComputerName & vbCrLf
                    End If

                Case 4
                    Forms("MailChangesForm").lblProcess.Caption = "Copying " & mstrDB & " to " & "\\" & !ComputerName & "\KMDS\Mail\In\KMDSLocal.mdb"
                    Forms("MailChangesForm").Repaint
                    SysCmd acSysCmdSetStatus, "Copying " & mstrDB & " to " & "\\" & !ComputerName & "\KMDS\Mail\In\KMDSLocal.mdb"
                    If !Cr4Installed = True Then
                        FileCopy mstrDB, "\\" & !ComputerName & "\KMDS\Mail\In\KMDSLocal.mdb"
                        FileCopy mstrDB, "\\" & !ComputerName & "\KMDS\CASHSYS.INS"
                        strValidRegisters = strValidRegisters & !ComputerName & vbCrLf
                    End If

                Case 5
                    Forms("MailChangesForm").lblProcess.Caption = "Copying " & mstrDB & " to " & "\\" & !ComputerName & "\KMDS\Mail\In\KMDSLocal.mdb"
                    Forms("MailChangesForm").Repaint
                    SysCmd acSysCmdSetStatus, "Copying " & mstrDB & " to " & "\\" & !ComputerName & "\KMDS\Mail\In\KMDSLocal.mdb"
                    If !Cr5Installed = True Then
                        FileCopy mstrDB, "\\" & !ComputerName & "\KMDS\Mail\In\KMDSLocal.mdb"
                        FileCopy mstrDB, "\\" & !ComputerName & "\KMDS\CASHSYS.INS"
                        strValidRegisters = strValidRegisters & !ComputerName & vbCrLf
                    End If
                End Select
                .MoveNext
            Loop
        End With

        If strValidRegisters <> "" Then
            SysCmd acSysCmdSetStatus, "Creating valid register file... \\" & gstrCashsysInsLoc & "\kmds\CRView.ins"
            Set fso = New FileSystemObject
            Set ts = fso.OpenTextFile("\\" & gstrCashsysInsLoc & "\kmds\CRView.ins", ForWriting, True, TristateFalse) '8 is append
            ts.Write strValidRegisters
            ts.Close
            Set ts = Nothing
            Set fso = Nothing
        End If

    If GetKMDSSettings("IOWSub", "False") = "True" Then
        ExportFilesForWeb
    End If




    Else
        MsgBox "error"
    End If

Exit_Command1_Click:
    SysCmd acSysCmdClearStatus
    DoCmd.Close acForm, "MailChangesForm"

    Exit Function


Err_Command1_Click:

    Select Case Err.Number
    Case 76
        If gcfHandleErrors Then On Error Resume Next
        fMsgBox "MailChanges" & vbCrLf & Err.Number & " " & Err.Description & vbCrLf & rstemp!ComputerName & " NOT ONLINE!"
        Resume Next
    Case Else
        fMsgBox "MailChanges" & vbCrLf & Err.Number & " " & Err.Description
        Resume Exit_Command1_Click
    End Select

End Function



'Public Sub ZTables()
'
'    Dim catSource As New ADOX.Catalog
'    Dim catTarget As New ADOX.Catalog
'    Dim tblSource As New ADOX.Table
'    Dim tblTarget As New ADOX.Table
'    Dim colTemp As New ADOX.Column
'    Dim idxSource As New ADOX.Index
'    Dim idxTarget As New ADOX.Index
'    Dim comQDF As New ADODB.Command
'
'
'    Dim strSQL As String
'
'    Dim fldTemp As ADODB.Field
'
'    Dim propTemp As ADOX.Property
'
'    Dim stDocName As String
'    Dim rstemp As New ADODB.Recordset
'    Dim varReturn As Variant
'    If gcfHandleErrors Then On Error GoTo ErrLog
'    mstrDB = "C:\KMDS\NewDB.mdb"
'
'    'Set wsp = DBEngine(0)
'    'Set dbs = wsp.CreateDatabase("Mydb")
'
'    varReturn = SysCmd(acSysCmdSetStatus, "Deleting " & mstrDB)
'    If Dir(mstrDB) <> "" Then Kill (mstrDB)
'    varReturn = SysCmd(acSysCmdSetStatus, "Creating " & mstrDB)
'    catTarget.Create "Provider = Microsoft.Jet.OLEDB.4.0;" _
'       & "Data Source = " & mstrDB
'
'    catSource.ActiveConnection = CurrentProject.Connection
'
''    stDocName = "qryCRObjects"
'    strSQL = "SELECT CRObjects.* FROM CRObjects "
'    strSQL = strSQL & "ORDER BY CRObjects.Type DESC , "
'    strSQL = strSQL & "         CRObjects.Name;"
'
'    Set comQDF = Nothing
'
'    rstemp.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'    With rstemp
'        .MoveLast
'        .MoveFirst
'
'        Do While Not .EOF
'            If gcfHandleErrors Then On Error GoTo ErrLogName
'
'
'            Select Case !Type
'            Case "Table"
'                'varReturn = SysCmd(acSysCmdSetStatus, "Deleting Records from z" & rsTemp!Name)
'                '1. delete the records in the temporary work table
'                With comQDF
'                    .ActiveConnection = CurrentProject.Connection
'                    .CommandType = adCmdText
'                    .CommandText = "Delete * from z" & rstemp!Name
'                    .Execute
'                End With
'                Set comQDF = Nothing
'
'            End Select
'            .MoveNext
'        Loop
'
'    End With
'
'    rstemp.Close
'
'    Set rstemp = Nothing
'    CurrentProject.Connection.Close
'
'ExitFunction:
'    varReturn = SysCmd(acSysCmdClearStatus)
'        varReturn = SysCmd(acSysCmdSetStatus, " ")
'    Exit Sub
'
'ErrLogName:
'    MsgBox Err.Number & " " & Err.Description
'    Resume ExitFunction
'
'ErrLog:
'    MsgBox Err.Number & " " & Err.Description
'
'    Resume ExitFunction
'End Sub



Sub SetConnectForInvolineX()

'good idea... but too cumbersome...
'easier to set a drive letter K: =  \\fs1\kmds and map register files to k:\mail\cr?.mdb
'
''dbsTemp As Database, _
''   strTable As String, strConnect As String, _
''   strSourceTable As String)
'
'    Dim tdfL As TableDef
'    Dim intTemp As Integer
'    Dim fsotemp As New FileSystemObject
'    Dim bDeleteTemp As Boolean
'    bDeleteTemp = False
'    Dim strTarg As String
'
'    LoadDefaults
'    ' Create a new TableDef, set its Connect and
'    ' SourceTableName properties based on the passed
'    ' arguments, and append it to the TableDefs collection.
'
'    If gcfHandleErrors Then On Error Resume Next
'    CurrentDb.TableDefs.Delete ("Involine1")
'    CurrentDb.TableDefs.Delete ("Involine2")
'    CurrentDb.TableDefs.Delete ("Involine3")
'    CurrentDb.TableDefs.Delete ("Involine4")
'
'
'    If gcfHandleErrors Then On Error GoTo ErrLog
'    'CR1 create linked table
'    bDeleteTemp = False
'    strTarg = "\\" & gstrCashsysInsLoc & "\kmds\mail\cr1.mdb"
'    If Not fsotemp.FileExists(strTarg) Then
'        bDeleteTemp = True
'        fsotemp.CopyFile "CashRegisterDetailTemplate.mdb", strTarg
'    End If
'    Set tdfL = CurrentDb().CreateTableDef("InvoLine1")
'    tdfL.Connect = ";DATABASE=" & strTarg
'    tdfL.SourceTableName = "Involine"
'    CurrentDb().TableDefs.Append tdfL
'    Set tdfL = Nothing
'    If bDeleteTemp Then fsotemp.DeleteFile strTarg, True
'
'    'CR2 create linked table
'    bDeleteTemp = False
'    strTarg = "\\" & gstrCashsysInsLoc & "\kmds\mail\cr2.mdb"
'    If Not fsotemp.FileExists(strTarg) Then
'        bDeleteTemp = True
'        fsotemp.CopyFile "CashRegisterDetailTemplate.mdb", strTarg
'    End If
'    Set tdfL = CurrentDb().CreateTableDef("InvoLine2")
'    tdfL.Connect = ";DATABASE=" & strTarg
'    tdfL.SourceTableName = "Involine"
'    CurrentDb().TableDefs.Append tdfL
'    Set tdfL = Nothing
'    If bDeleteTemp Then fsotemp.DeleteFile strTarg, True
'
'    'CR3 create linked table
'    bDeleteTemp = False
'    strTarg = "\\" & gstrCashsysInsLoc & "\kmds\mail\cr3.mdb"
'    If Not fsotemp.FileExists(strTarg) Then
'        bDeleteTemp = True
'        fsotemp.CopyFile "CashRegisterDetailTemplate.mdb", strTarg
'    End If
'    Set tdfL = CurrentDb().CreateTableDef("InvoLine3")
'    tdfL.Connect = ";DATABASE=" & strTarg
'    tdfL.SourceTableName = "Involine"
'    CurrentDb().TableDefs.Append tdfL
'    Set tdfL = Nothing
'    If bDeleteTemp Then fsotemp.DeleteFile strTarg, True
'
'    'CR4 create linked table
'    bDeleteTemp = False
'    strTarg = "\\" & gstrCashsysInsLoc & "\kmds\mail\cr4.mdb"
'    If Not fsotemp.FileExists(strTarg) Then
'        bDeleteTemp = True
'        fsotemp.CopyFile "CashRegisterDetailTemplate.mdb", strTarg
'    End If
'    Set tdfL = CurrentDb().CreateTableDef("InvoLine4")
'    tdfL.Connect = ";DATABASE=" & strTarg
'    tdfL.SourceTableName = "Involine"
'    CurrentDb().TableDefs.Append tdfL
'    Set tdfL = Nothing
'    If bDeleteTemp Then fsotemp.DeleteFile strTarg, True
'    ' Delete the linked table because this is a demonstration.
ExitFunction:
    Exit Sub

ErrLogName:
    MsgBox Err.Number & " " & Err.Description
    Resume ExitFunction

ErrLog:
    MsgBox Err.Number & " " & Err.Description
    Resume ExitFunction
End Sub

'        ConnectOutput dbsTemp, _
'            "JetTable", _
'            ";DATABASE=C:\My Documents\Northwind.mdb", _
'            "Employees"

'Sub ConnectOutput(dbsTemp As Database, _
'   strTable As String, strConnect As String, _
'   strSourceTable As String)
'
'   Dim tdfLinked As TableDef
'   Dim rstLinked As Recordset
'   Dim intTemp As Integer
'
'   ' Create a new TableDef, set its Connect and
'   ' SourceTableName properties based on the passed
'   ' arguments, and append it to the TableDefs collection.
'   Set tdfLinked = dbsTemp.CreateTableDef(strTable)
'
'   tdfLinked.Connect = strConnect
'   tdfLinked.SourceTableName = strSourceTable
'   dbsTemp.TableDefs.Append tdfLinked
'
'   Set rstLinked = dbsTemp.OpenRecordset(strTable)
'
'   If not gcfHandleErrors then Debug.Print "Data from linked table:"
'
'   ' Display the first three records of the linked table.
'   intTemp = 1
'   With rstLinked
'      Do While Not .EOF And intTemp <= 3
'         If not gcfHandleErrors then Debug.Print , .Fields(0), .Fields(1)
'         intTemp = intTemp + 1
'         .MoveNext
'      Loop
'      If Not .EOF Then If not gcfHandleErrors then Debug.Print , "[additional records]"
'      .Close
'   End With
'
'   ' Delete the linked table because this is a demonstration.
'   dbsTemp.TableDefs.Delete strTable
'
'End Sub

Public Function MailPictures(Confirm As Boolean)

    'Dim rstemp As New ADODB.Recordset
    Dim rstemp As Object
    Set rstemp = CreateObject("ADODB.Recordset")

    Dim bAtLeastOneRegisterInstalled


    Dim strValidRegisters As String
    Dim strServerDir As String
    Dim strWorkStationDir As String


    If gcfHandleErrors Then On Error GoTo ErrMailPictures
    On Error GoTo ErrMailPictures
    strServerDir = "\\" & gstrCashsysInsLoc & "\KMDS\Pictures"
    bAtLeastOneRegisterInstalled = False
    rstemp.Open "qryCashRegisters", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     With rstemp
        Do While Not .EOF And Not bAtLeastOneRegisterInstalled
            Select Case !ID
            Case 1
                If !Cr1Installed = True Then bAtLeastOneRegisterInstalled = True
            Case 2
                If !Cr2Installed = True Then bAtLeastOneRegisterInstalled = True
            Case 3
                If !Cr3Installed = True Then bAtLeastOneRegisterInstalled = True
            Case 4
                If !Cr4Installed = True Then bAtLeastOneRegisterInstalled = True
            Case 5
                If !Cr5Installed = True Then bAtLeastOneRegisterInstalled = True
            End Select
            .MoveNext
        Loop
        .MoveFirst
    End With

    If bAtLeastOneRegisterInstalled = False Then
        'MsgBox ("There are no registers defined")
        Exit Function
    End If


    If Confirm Then
        If MsgBox("Do you want to send the changes to the Cash Register now?", vbYesNo) = vbNo Then
            Exit Function
        End If
    End If

'    Set rsTemp = CurrentDb().OpenRecordset("qryCashRegisters")


        With rstemp
            Do While Not .EOF
                strWorkStationDir = "\\" & !ComputerName & "\KMDS\Pictures"
                Select Case !ID
                Case 1
                    If !Cr1Installed = True Then CopyMyFolder strServerDir, strWorkStationDir
                Case 2
                    If !Cr2Installed = True Then CopyMyFolder strServerDir, strWorkStationDir
                Case 3
                    If !Cr3Installed = True Then CopyMyFolder strServerDir, strWorkStationDir
                Case 4
                    If !Cr4Installed = True Then CopyMyFolder strServerDir, strWorkStationDir
                Case 5
                    If !Cr5Installed = True Then CopyMyFolder strServerDir, strWorkStationDir
                End Select
                .MoveNext
            Loop
        End With

ExitFunction:
    SysCmd acSysCmdClearStatus
    Exit Function

ErrMailPictures:
    MsgBox "ErrMailPictures: Error #" & Err.Number & " " & Err.Description
    Resume ExitFunction
End Function

Public Sub CopyMyFolder(SourceFolder As String, TargetFolder As String)
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    If gcfHandleErrors Then On Error GoTo ErrLog
    If SourceFolder <> TargetFolder Then

        With Forms("SalesPicturesRefresh")
            .txtStatus.Visible = True
            .txtStatus.SetFocus
            .txtStatus = "Copying " & SourceFolder & " to " & TargetFolder
            .tabControl.Visible = False

            .Refresh
            .Repaint
        End With
        SysCmd acSysCmdSetStatus, "Copying " & SourceFolder & " to " & TargetFolder
        If Not fso.FolderExists(TargetFolder) Then
            fso.CreateFolder (TargetFolder)
        End If
        fso.CopyFolder SourceFolder, TargetFolder, True
    End If
ExitSub:
    SysCmd acSysCmdClearStatus
    Exit Sub

ErrLog:
    MsgBox "CopyMyFolder: Error #" & Err.Number & " " & Err.Description
    Resume ExitSub
End Sub
