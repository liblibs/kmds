Attribute VB_Name = "modEOD"
Option Compare Database
Option Explicit


Public Function EndOFDayStream()
Dim strTargetPath As String
Dim strTargetFile As String
Dim strSourceFileName As String


Dim fso As FileSystemObject
Dim ts As TextStream
Dim f As Folder
Dim strMsg As String

Dim strSQL As String
Dim colFiles As New Collection

    If gcfHandleErrors Then On Error GoTo ErrLog
    gbEOProcessError = False
    EOLog "========================================================="
    EOLog "End of Day - Start"
    EOLog "========================================================="

    If gstrMailDir = "" Then
        EOLog "gstrMailDir =''; calling Main"
        LoadDefaults
    End If
    If gstrNetworkDatabase = "" Then
        EOLog "gstrNetworkDatabase =''; calling Main"
        LoadDefaults
    End If

    EOLog "gstrMailDir = " & gstrMailDir
    EOLog "gstrNetworkDatabase = " & gstrNetworkDatabase

    'create the directory for today e.g. "c:\KMDS\reports\2003\12\25"
    Set fso = New FileSystemObject

    strTargetPath = gstrMailDir & "Backup\"
    If Not (fso.FolderExists(strTargetPath)) Then
        Set f = fso.CreateFolder(strTargetPath)
        EOLog "Creating Folder : " & strTargetPath
    End If

    strTargetPath = strTargetPath & Format(Now, "yyyy") & "\"
    If Not (fso.FolderExists(strTargetPath)) Then
        Set f = fso.CreateFolder(strTargetPath)
        EOLog "Creating Folder : " & strTargetPath
    End If

    strTargetPath = strTargetPath & Format(Now, "mm") & "\"
    If Not (fso.FolderExists(strTargetPath)) Then
        Set f = fso.CreateFolder(strTargetPath)
        EOLog "Creating Folder : " & strTargetPath
    End If

    strTargetPath = strTargetPath & Format(Now, "dd") & "\"
    If Not (fso.FolderExists(strTargetPath)) Then
        Set f = fso.CreateFolder(strTargetPath)
        EOLog "Creating Folder : " & strTargetPath
    End If

'added 10/05/2004
'    the gstrNetworkDatabase is a static file and copies exist else where...
'    therefore no reason to make duplicates of it every night...

'    strTargetFile = "OLCC" & Format(Now, "yyyy-mm-dd hh-nn-ss") & ".mdb"
'    EOLog "Copying " & gstrNetworkDatabase & " to " & strTargetPath & strTargetFile
'    fso.CopyFile gstrNetworkDatabase, strTargetPath & strTargetFile
'
'    If Not fso.FileExists(strTargetPath & strTargetFile) Then
'        strMsg = "Error Copying " & vbcrlf
'        strMsg = strMsg & gstrNetworkDatabase & vbcrlf
'        strMsg = strMsg & "to" & vbcrlf
'        strMsg = strMsg & strTargetPath & strTargetFile & vbcrlf
'        EOLog "ERROR " & strMsg
'        MsgBox strMsg
'        Exit Function
'    End If
'
    'do the cash register files exist?
    If Not fso.FileExists(gstrMailDir & "Mail\CR1.MDB") Then
        EOLog "ERROR " & gstrMailDir & "Mail\CR1.MDB not found"
        MsgBox gstrMailDir & "Mail\CR1.MDB not found", vbCritical, "EndOfDay"
        gbEOProcessError = True
        GoTo ExitSub
    End If
    If Not fso.FileExists(gstrMailDir & "Mail\CR2.MDB") Then
        EOLog "ERROR " & gstrMailDir & "Mail\CR2.MDB not found"
        MsgBox gstrMailDir & "Mail\CR2.MDB not found", vbCritical, "EndOfDay"
        gbEOProcessError = True
        GoTo ExitSub
    End If
    If Not fso.FileExists(gstrMailDir & "Mail\CR3.MDB") Then
        EOLog "ERROR " & gstrMailDir & "Mail\CR3.MDB not found"
        MsgBox gstrMailDir & "Mail\CR3.MDB not found", vbCritical, "EndOfDay"
        gbEOProcessError = True
        GoTo ExitSub
    End If
    If Not fso.FileExists(gstrMailDir & "Mail\CR4.MDB") Then
        EOLog "ERROR " & gstrMailDir & "Mail\CR4.MDB not found"
        MsgBox gstrMailDir & "Mail\CR4.MDB not found", vbCritical, "EndOfDay"
        gbEOProcessError = True
        GoTo ExitSub
    End If
    If Not fso.FileExists(gstrMailDir & "Mail\CR5.MDB") Then
        EOLog "ERROR " & gstrMailDir & "Mail\CR5.MDB not found"
        MsgBox gstrMailDir & "Mail\CR4.MDB not found", vbCritical, "EndOfDay"
        gbEOProcessError = True
        GoTo ExitSub
    End If

    EOLog "Dumping SQL Tables"
    'DumpSQLTables
    'changed 2007-04-07 -
    ' just do the dump at this point...
    ' archive and copy archive as last step below
    EOLog "MakeSQLBackup"
    MakeSQLBackup

    'added 2008-12-11 ... build a list of files to e-mail at end of stream
    EOLog "Delete : C:\kmds\EOD File Transmit List.txt"
    If fso.FileExists("C:\kmds\EOD File Transmit List.txt") Then
        fso.DeleteFile "C:\kmds\EOD File Transmit List.txt", True
    End If

    'added 2008-12-11 ... to expedite reprint lastnite
    EOLog "Delete : K:LastNiteFiles.BAT"
    If fso.FileExists("K:LastNiteFiles.BAT") Then
        fso.DeleteFile "K:LastNiteFiles.BAT", True
    End If

    GetKMDSSettings "DaysToRetain", "500"

'*********************************************************************************************************************

    If EOProcess("qryEOD", "DayEnd") = True Then
        EOLog "Moving :" & gstrMailDir & "Mail\CR1.MDB to " & strTargetPath
        fso.MoveFile gstrMailDir & "Mail\CR1.MDB", strTargetPath & "CR1" & Format(Now, "yyyy-mm-dd hh-nn-ss") & ".mdb"
        EOLog "Moving :" & gstrMailDir & "Mail\CR2.MDB to " & strTargetPath
        fso.MoveFile gstrMailDir & "Mail\CR2.MDB", strTargetPath & "CR2" & Format(Now, "yyyy-mm-dd hh-nn-ss") & ".mdb"
        EOLog "Moving :" & gstrMailDir & "Mail\CR3.MDB to " & strTargetPath
        fso.MoveFile gstrMailDir & "Mail\CR3.MDB", strTargetPath & "CR3" & Format(Now, "yyyy-mm-dd hh-nn-ss") & ".mdb"
        EOLog "Moving :" & gstrMailDir & "Mail\CR4.MDB to " & strTargetPath
        fso.MoveFile gstrMailDir & "Mail\CR4.MDB", strTargetPath & "CR4" & Format(Now, "yyyy-mm-dd hh-nn-ss") & ".mdb"
        EOLog "Moving :" & gstrMailDir & "Mail\CR5.MDB to " & strTargetPath
        fso.MoveFile gstrMailDir & "Mail\CR5.MDB", strTargetPath & "CR5" & Format(Now, "yyyy-mm-dd hh-nn-ss") & ".mdb"
    End If

    'moved from bottom of routine to here as per EFitz request.

    If GetKMDSSettings("Fintech.Active", "False") = "True" Then
        EOLog "Fintech.Active = True"
        FinTechExport
        EOLog "Fintech.Active Finished"

    Else
        EOLog "Fintech = False"

    End If

    'added 2007-01-21
    EOLog "Create Daily Transmit files"
    'If gbOLCCFTPEndOfDaySend Then
    OLCCCreateDailyTransmitFiles

    EOLog "Creating transmit file for inventory levels..."

    OLCCInventFile CLng(glnAgentNumber), True

    'added 05/14/2005
    If gbTransferCartonsToPacks Then
        EOLog "Transfer Cartons To Packs Begin"
        TransferCartonsToPacks
        EOLog "Transfer Cartons To Packs End"
    End If

    'added 05/14/2005
    If gbRentrakInstalled Then
        EOLog "Rentrak Export Begin"
        RentrakExport
        EOLog "RentrakExport End"
    End If

'    'added 08/30/2005
'
'    EOLog "Append Time Clock Punches Begin"
'    TimeClockPunchAppend
'    EOLog "Append Time Clock Punches End"
'

    EOLog "Append Lottery Tickets Recon Info Begin"
    SQLPassThru "sp_LotteryScratch"
    'LotteryTicketsAppend
    EOLog "Append Lottery Tickets Recon Info End"

    'ADDED 2005-12-04
    EOLog "Reindex Tables Begin"
    ReindexTables
    EOLog "Reindex Tables End"

    EOLog "Updating ItemMovement"
    SQLPassThru "sp_UpdateItemMovement"
    'added 2008-01-12
    EOLog "Updating ItemCost"
    SQLPassThru "sp_UpdateItemCost"


    'UpdateMovement False
    EOLog "Updating Average Daily Sales and Assign DaysSupplyGroupIDs to Item Table..."
    AssignDaysSupplyGroupIDs


    'End If
'   'removed 2009-05-19
'    'added 2006-12-06
'    EOLog "Purge pcCharge"
'    EOLog " Shell('\\' & gstrCashsysInsLoc & '\KMDS\pcP.exe', 1)"
'    Call Shell("\\" & gstrCashsysInsLoc & "\KMDS\pcP.exe", 1)

    'Added to 31415
    If VerifyCurrentKMDSCustomer Then
        EOLog "Sending OLCC Daily Transmit Files"
        'OLCCDailyTransmit
        If gbOLCCFTPEndOfDaySend Then
            OLCCDailyTransmit
            EOLog "...Complete Sending OLCC Daily Transmit Files"
        Else
            EOLog "Sending KMDS Daily Transmit Files (Send to OLCC not checked in G-5)"
            WinSCPFTPTransmitKMDS
            EOLog "... Complete"
        End If
    Else
        EOLog "... NOT SENT: Please contact KMDS Support to update your KMDS Subscription"
        EOLog "Sending KMDS Daily Transmit Files (not current customer)"
        WinSCPFTPTransmitKMDS

    End If

    EOLog "MailSQLBackup"
    MailSQLBackup

    EOLog "EODTextMessage"
    EODTextMessage

    EOLog "EODeMailReports"

    If fso.FileExists("C:\kmds\EOD File Transmit List.txt") Then
        Set ts = fso.OpenTextFile("C:\kmds\EOD File Transmit List.txt", ForReading, False)
        Do While Not ts.AtEndOfStream
            strMsg = ts.ReadLine
            If fso.FileExists(strMsg) Then
                EODeMailFiles.Add strMsg
                EOLog "  " & strMsg
            End If
        Loop
        ts.Close
        Set ts = Nothing
        If EODeMailFiles.Count > 0 Then
            KMDSeMail GetKMDSSettings("KMDSMail EODeMail", "info@kmds.us"), EODeMailFiles, "EOD - " & gstrStoreName
        End If
    End If

    EOLog "Make Online Backup Batch File"
    MakeOnlineBackupBatchFile

    EOLog "Append UPCItems With ItemID"
    AppendUPCItemsWithItemID


    'added 11/5/2013
    If GetKMDSSettings("DropBoxInstalled", "False") = "True" Then
        EOLog "DropBoxInstalled = True"
        strSourceFileName = "c:\kmds\backup\KMDSBACK.zip"
        strTargetPath = GetKMDSSettings("DropBoxFolder", "N/A")
        EOLog "DropBoxFolder = " & strTargetPath
        If Not (fso.FolderExists(strTargetPath)) Then
            EOLog strTargetPath & " does not exist"
        Else
            EOLog "Copying " & strSourceFileName & " to " & strTargetPath & "\KMDSSQL.ZIP"
            FileCopy strSourceFileName, strTargetPath & "\KMDSSQL.ZIP"
        End If
    Else
        EOLog "DropBoxInstalled = False"

    End If

    EOLog "Getting OLCC ICE Item Files"
    If GetKMDSSettings("GetItemICEatEOD", "True") = "True" Then
        GetOLCCItemFiles
    End If
    EOLog ".... complete!"
    EOLog "========================================================="
    EOLog "End of Day = Finished Success"
    EOLog "========================================================="

ExitSub:
    If gbEOProcessError Then
        strMsg = gstrMailDir & "Log\"
        strMsg = strMsg & Format(Now, "yyyy") & "\"
        strMsg = strMsg & Format(Now, "mm") & "\"
        strMsg = strMsg & Format(Now, "dd") & "\"
        strMsg = strMsg & "EOLog.Txt"
        StartDoc strMsg
        'added

        If InStr(GetKMDSSettings("KMDSMail supportemail", "info@kmds.us"), "@") <> 0 Then 'suppress sending support email by taking out the @ sign in
            colFiles.Add strMsg
            'added 2009.05.07
'            If WhatIsMyIp Then
'                colFiles.Add "c:\kmds\kmdsgetweb.txt"
'            End If
            KMDSeMail GetKMDSSettings("KMDSMail supportemail", "info@kmds.us"), colFiles, "EOD Support " & gstrStoreName
        End If
    Else
        'MsgBox "End of Day Finished Successfully"
    End If
    Exit Function

ErrLog:
    gbEOProcessError = True
    EOLog "Error #" & Err.Number & " " & Err.Description & " End of Day"
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbCritical, "EndOfDay"
    EOLog "========================================================="
    EOLog "End of Day = Finished with Errors"
    EOLog "========================================================="
    Resume ExitSub
End Function

Public Sub EOLog(Optional LogEntry As String)

Dim fso As FileSystemObject
Dim f As Folder
Dim ts As Object
Dim fldr As String
    If gcfHandleErrors Then On Error GoTo ErrLog
    If gstrMailDir = "" Then LoadDefaults

    'create the directory for today e.g. "c:\KMDS\reports\2003\12\25"
    Set fso = New FileSystemObject

    fldr = gstrMailDir & "Log\"
    If Not (fso.FolderExists(fldr)) Then
          Set f = fso.CreateFolder(fldr)
    End If

    fldr = fldr & Format(Now, "yyyy") & "\"
    If Not (fso.FolderExists(fldr)) Then
          Set f = fso.CreateFolder(fldr)
    End If

    fldr = fldr & Format(Now, "mm") & "\"
    If Not (fso.FolderExists(fldr)) Then
          Set f = fso.CreateFolder(fldr)
    End If

    fldr = fldr & Format(Now, "dd") & "\"
    If Not (fso.FolderExists(fldr)) Then
          Set f = fso.CreateFolder(fldr)
    End If

    fldr = fldr & "EOLog.Txt"
    Set f = Nothing

    Set ts = fso.OpenTextFile(fldr, 8, True, TristateFalse) '8 is append
    ts.Write Format(Now, "hh:nn:ss ") & LogEntry & vbCrLf
    ts.Close
    Set ts = Nothing
'    With Forms("EndOFDay")
'        .txtStatus = .txtStatus & LogEntry & vbcrlf
'        .Repaint
'        .txtStatus.SetFocus
'    End With

ExitSub:
    Exit Sub
ErrLog:
    fMsgBox "EOLog"
    Resume ExitSub
End Sub

Public Sub EndOFMonth()
    Dim fsoTemp As New FileSystemObject

    If gcfHandleErrors Then On Error GoTo ErrLog
    If Dir("C:\npos\NEWPRICE_ICExpansion.fil") <> "" Then
        If Dir("C:\npos\NEWPRICE_ICExpansion.txt") <> "" Then
            Kill ("C:\npos\NEWPRICE_ICExpansion.txt")
        End If
        Name "C:\npos\NEWPRICE_ICExpansion.FIL" As "C:\npos\NEWPRICE_ICExpansion.TXT"
    End If

    'added 31415
    If (GetKMDSSettings("UseWinSCPForOLASTransmits", "False")) = True Then
        'build a Dayend files
        '1. WinSCPDaily.txt 'script file used by :
        '2. WinSCPDaily.bat

        CreateWinSCPtoOLASFTPScriptFile ("Newprice")
        WaitForExecCmd "c:\npos\WinSCPNewprice.BAT"

    End If

    CurrentDb().Execute ("delete * from NewPrice")
    DoCmd.TransferText acImportFixed, "NewPriceImportSpecification", "NewPrice", "C:\npos\NEWPRICE_ICExpansion.txt", False
    EOProcess "qryEOM", "MonthEnd"
    'update the whole file!
    'changed 2005-12-08
    SQLPassThru "sp_UpdateItemMovement"


    'UpdateMovement False
    fsoTemp.CopyFile "C:\npos\NEWPRICE_ICExpansion.txt", "c:\npos\archive\NEWPRICE_ICExpansion" & Format(Now, "yyyy-mm-dd-nn-ss") & ".txt"
    fsoTemp.DeleteFile "C:\npos\NEWPRICE_ICExpansion.txt"
ExitSub:
    Exit Sub

ErrLog:
    EOLog "Error #" & Err.Number & " " & Err.Description & " : EndofMonth"
    fMsgBox "EndofMonth"
    Resume ExitSub
End Sub


Public Function EOProcess(QuerySequencePrefix As String, ShareName As String) As Boolean

    'Dim rstemp As New ADODB.Recordset
    Dim rstemp As Object
    Set rstemp = CreateObject("ADODB.Recordset")

    'Dim comQDF As New Command
    Dim comQDF As Object
    Set comQDF = CreateObject("ADODB.COMMAND")

    Dim strSQL As String

    If gcfHandleErrors Then On Error GoTo ErrLog
    EOProcess = True
    strSQL = "SELECT MSysObjects.Name, "
    strSQL = strSQL & "MSysObjects.Type "
    strSQL = strSQL & "FROM MSysObjects "
    strSQL = strSQL & "WHERE (((MSysObjects.Name) Like '" & QuerySequencePrefix & "%')) "
    strSQL = strSQL & "ORDER BY MSysObjects.Name"

    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandType = adCmdText
        .CommandText = strSQL
    End With

    Set rstemp = comQDF.Execute

    If gcfHandleErrors Then On Error GoTo ErrLogWithName
    With rstemp
        Do While Not .EOF

            EOLog !Name & " " & !Type
            Select Case !Type
            Case 5 '
                QueryExecute !Name
            Case -32764
                ReportExecute !Name, ShareName
            Case Else
                gbEOProcessError = True
                MsgBox "Unknown qryEOD type: " & !Type & " for " & !Name
            End Select
            .MoveNext
        Loop
        .Close
    End With
    Set rstemp = Nothing

ExitSub:
    Exit Function

ErrLogWithName:
    strSQL = rstemp!Name
ErrLog:
    EOProcess = False
    EOLog strSQL & vbCrLf & "Error #" & Err.Number & " " & Err.Description & " : EOProcess"
    fMsgBox "EOProcess " & strSQL
    gbEOProcessError = True
    Resume ExitSub
End Function

Public Sub QueryExecute(QueryName As String)

If Not gcfHandleErrors Then Debug.Print QueryName

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

If gcfHandleErrors Then On Error GoTo ErrLog
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = QueryName
        .CommandType = adCmdStoredProc
        .CommandTimeout = 60000
        .Execute

    End With
    Set comQDF = Nothing
ExitSub:
    Exit Sub
ErrLog:
    EOLog "Error #" & Err.Number & " " & Err.Description & " : QueryExecute : " & QueryName
    fMsgBox "QueryExecute : " & QueryName
    gbEOProcessError = True
    Resume ExitSub
End Sub



Public Sub ReportExecute(ReportName As String, ShareName As String)
'If Not gcfHandleErrors Then Debug.Print ReportName
Dim nNumberOfCopies As Integer
Dim strEmailFileName As String
Dim beMail As Boolean
Dim fsoTemp As New FileSystemObject
Dim ts As TextStream


    'look up report in EOReport for number copies
    'Dim rstemp As New ADODB.Recordset
    Dim rstemp As Object
    Set rstemp = CreateObject("ADODB.Recordset")

    If gcfHandleErrors Then On Error GoTo ErrLog
    With rstemp
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "EOReport", CurrentProject.Connection, , , adCmdTable
    End With

    rstemp.Find "ReportName='" & ReportName & "'"  ', adSeekFirstEQ

    If rstemp.EOF Then
        EOLog "Setting up new reportname : " & ReportName
        'remove 2006-03-02
        'MsgBox "Setting up new reportname : " & ReportName
        rstemp.AddNew
        rstemp!ReportName = ReportName
        'changed 2103-02-08
        'default to 0 instead of 1
        rstemp!NumberOfCopies = 0
        rstemp.Update
        nNumberOfCopies = 0
    Else
        nNumberOfCopies = rstemp!NumberOfCopies
    End If
    beMail = rstemp!EMail

    rstemp.Close
    Set rstemp = Nothing
    On Error Resume Next
    DoCmd.OpenReport ReportName, acViewPreview
    strEmailFileName = ArchiveReportShare(ReportName, ShareName)

    If nNumberOfCopies > 0 Then
        DoCmd.PrintOut acPrintAll, , , , nNumberOfCopies
        'added 2009-02-09
        If ShareName = "DayEnd" Then
            Set ts = fsoTemp.OpenTextFile("K:LastNiteFiles.Bat", ForAppending, True, TristateMixed)
            ts.WriteLine Chr(34) & strEmailFileName & Chr(34)
            ts.Close
            Set ts = Nothing
        End If
    End If


    DoCmd.Close acReport, ReportName

    'build a list of files to transmit
    If beMail Then
        Set ts = fsoTemp.OpenTextFile("C:\kmds\EOD File Transmit List.txt", ForAppending, True, TristateMixed)
        ts.WriteLine strEmailFileName
        ts.Close
        Set ts = Nothing
    End If

    'added 31416
    If gbKMDSFTPActive Then 'send files to ftp.kmds.us
        If ReportName = "qryEOD040SalesByDepartment" Then
            Dim strt As String
            strt = "K:\Mail\KMDSFTP\OUT\ADS " & glnAgentNumber & " " & Format(Now, "yyyy-mm-dd hh-nn-ss")
            If gbSaveReportsAsPDF Then
                strt = strt & ".pdf"
            Else
                strt = strt & ".snp"
            End If
            fsoTemp.CopyFile strEmailFileName, strt, True
        End If
    End If

ExitSub:
    Exit Sub

ErrLog:
    Select Case Err.Number
    Case 2501 ' cancel error (No data)
        EOLog "  NO DATA TO REPORT!"
    Case Else
        EOLog "Error #" & Err.Number & " " & Err.Description & " : ReportExecute : " & ReportName
        fMsgBox "ReportExecute : " & ReportName & vbCrLf & Err.Number & " " & Err.Description
        gbEOProcessError = True
    End Select
    Resume ExitSub
End Sub



Public Function UpdateMovement(CurrentMonthOnly As Boolean)
Dim varReturn As Variant

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")


    If gcfHandleErrors Then On Error GoTo ErrLog

    If Not gcfHandleErrors Then Debug.Print "Step 1 "; Now

    'This will 'reprice' all orders with vendor #1
    'to fix the price for orders input in the previous month
    varReturn = SysCmd(acSysCmdSetStatus, "Update Item Movement Step # of ")

    If gcfHandleErrors Then On Error GoTo ErrLog
    SetOptions False
    varReturn = SysCmd(acSysCmdSetStatus, "Update Item Movement")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "sp_UpdateItemMovement"
        .CommandType = adCmdStoredProc
        .Execute
    End With
    Set comQDF = Nothing



ExitSub:
    SetOptions True
    varReturn = SysCmd(acSysCmdClearStatus)
    If Not gcfHandleErrors Then Debug.Print Now

    Exit Function

ErrLog:
    fMsgBox "UpdateMovement" & vbCrLf & Err.Number & " " & Err.Description
    gbEOProcessError = True
    Resume ExitSub

End Function

Public Function UpdateMovementByADO(CurrentMonthOnly As Boolean)
Dim varReturn As Variant


'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

    If gcfHandleErrors Then On Error GoTo ErrLog

    If Not gcfHandleErrors Then Debug.Print "Step 1 "; Now

    'This will 'reprice' all orders with vendor #1
    'to fix the price for orders input in the previous month
    varReturn = SysCmd(acSysCmdSetStatus, "Update Item Movement Step # of ")

    If gcfHandleErrors Then On Error GoTo ErrLog
    SetOptions False
    If Not CurrentMonthOnly Then 'this only needs to be done at month end
        varReturn = SysCmd(acSysCmdSetStatus, "Update Item Movement Step #1 of 10")
        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryIMStep01DeleteItemMovement"
            .CommandType = adCmdStoredProc
            .Execute
        End With
        Set comQDF = Nothing
        Set comQDF = CreateObject("ADODB.COMMAND")

        If Not gcfHandleErrors Then Debug.Print "Step 2a "; Now
        varReturn = SysCmd(acSysCmdSetStatus, "Update Item Movement Step #2a of 10")
        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryIMStep02aAppendItemMovement"
            .CommandType = adCmdStoredProc
            .Execute
        End With
        Set comQDF = Nothing
        Set comQDF = CreateObject("ADODB.COMMAND")

        If Not gcfHandleErrors Then Debug.Print "Step 2b "; Now
        varReturn = SysCmd(acSysCmdSetStatus, "Update Item Movement Step #2b of 10")
        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryIMStep02bAppendItemMovementPacks"
            .CommandType = adCmdStoredProc
            .Execute
        End With
        Set comQDF = Nothing
        Set comQDF = CreateObject("ADODB.COMMAND")

    End If

    If Not gcfHandleErrors Then Debug.Print "Step 3 "; Now
    varReturn = SysCmd(acSysCmdSetStatus, "Update Item Movement Step #3 of 10")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryIMStep03DeleteItemMovementCM"
        .CommandType = adCmdStoredProc
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")


    If Not gcfHandleErrors Then Debug.Print "Step 4 "; Now
    varReturn = SysCmd(acSysCmdSetStatus, "Update Item Movement Step #4 of 10")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryIMStep04AppendItemMovementCM"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("parDate", adDate, adParamInput)
        .Parameters("parDate") = CDate(Month(Now) & "/01/" & Year(Now))
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")


    'DoCmd.RunSQL (CurrentDb().QueryDefs("qryIM Step 4 AppendItemMovementCM").SQL)

    If Not gcfHandleErrors Then Debug.Print "Step 5 "; Now
    varReturn = SysCmd(acSysCmdSetStatus, "Update Item Movement Step #5 of 10")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryIMStep05UpdateItemMovementCM"
        .CommandType = adCmdStoredProc
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")


    If Not gcfHandleErrors Then Debug.Print "Step 6 "; Now
    varReturn = SysCmd(acSysCmdSetStatus, "Update Item Movement Step #6 of 10")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryIMStep06AppendItemMovementNewCM"
        .CommandType = adCmdStoredProc
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")


    If Not gcfHandleErrors Then Debug.Print "Step 7 "; Now
    varReturn = SysCmd(acSysCmdSetStatus, "Update Item Movement Step #7 of 10")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryIMStep07DeleteItemMovementSalesTotal"
        .CommandType = adCmdStoredProc
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")


    If Not gcfHandleErrors Then Debug.Print "Step 8 "; Now
    varReturn = SysCmd(acSysCmdSetStatus, "Update Item Movement Step #8 of 10")
'    With comqdf
 '       .ActiveConnection = CurrentProject.Connection
 '       .CommandText = "qryIMStep08AppendIMSalesTotal"
 '       .CommandType = adCmdStoredProc
 '       .Execute
 '   End With
    DoCmd.RunSQL (CurrentDb().QueryDefs("qryIMStep08AppendIMSalesTotal").SQL)

'    Set comQDF = Nothing

    'DoCmd.RunSQL (CurrentDb().QueryDefs("qryIM Step 9 ReversePurchases").SQL)

    If Not gcfHandleErrors Then Debug.Print "Step 10 "; Now
    varReturn = SysCmd(acSysCmdSetStatus, "Update Item Movement Step #10 of 10")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryIMStep10AppendAllTransactions"
        .CommandType = adCmdStoredProc
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")


ExitSub:
    SetOptions True
    varReturn = SysCmd(acSysCmdClearStatus)
    If Not gcfHandleErrors Then Debug.Print Now

    Exit Function

ErrLog:
    fMsgBox "UpdateMovement" & vbCrLf & Err.Number & " " & Err.Description
    gbEOProcessError = True
    Resume ExitSub

End Function



Public Function testsp()
'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "sp_PassThru"
        .CommandType = adCmdStoredProc
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    If Not gcfHandleErrors Then Debug.Print Now
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qt"
        .CommandType = adCmdStoredProc
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    If Not gcfHandleErrors Then Debug.Print Now

    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "sp_PassThru"
        .CommandType = adCmdStoredProc
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    If Not gcfHandleErrors Then Debug.Print Now
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryIMStep02aAppendItemMovement"
        .CommandType = adCmdStoredProc
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    If Not gcfHandleErrors Then Debug.Print Now

End Function


Public Function EOReports() As Boolean
'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

    Dim strSQL As String

    If gcfHandleErrors Then On Error GoTo ErrLog
    'EOProcess = True
    strSQL = "SELECT MSysObjects.Name, "
    strSQL = strSQL & "MSysObjects.Type "
    strSQL = strSQL & "FROM MSysObjects "
    strSQL = strSQL & "WHERE (((MSysObjects.Name) Like 'qryEOD%')) "
    strSQL = strSQL & "ORDER BY MSysObjects.Name"

    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandType = adCmdText
        .CommandText = strSQL
    End With

    Set rstemp = comQDF.Execute

    If gcfHandleErrors Then On Error GoTo ErrLogWithName
    With rstemp
        Do While Not .EOF

            EOLog !Name & " " & !Type
            Select Case !Type
            Case 5 '
'                QueryExecute !Name
            Case -32764
                ReportExecute !Name, "EOD"

            Case Else
                gbEOProcessError = True
                MsgBox "Unknown qryEOD type: " & !Type & " for " & !Name
            End Select
            .MoveNext
        Loop
        .Close
    End With
    Set rstemp = Nothing

ExitSub:
    Exit Function

ErrLogWithName:
    strSQL = rstemp!Name
ErrLog:
    'EOProcess = False
    EOLog strSQL & vbCrLf & "Error #" & Err.Number & " " & Err.Description & " : EOProcess"
    fMsgBox "EOProcess " & strSQL
    gbEOProcessError = True
    Resume ExitSub
End Function


Public Function TransferCartonsToPacks()
    If gbTransferCartonsToPacks Then
        ReportExecute "TransferCartonsToPacks", "EOD"
        QueryExecute "qryItemsCartonsToPacks_Step2_UpdatePacks"
        QueryExecute "qryItemsCartonsToPacks_Step3_UpdateCartons"
    End If
End Function



Public Sub SQLPassThru(SQLStatement As String)

Dim strMsg As String
Dim varReturn As Variant

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

'Dim adoCnOLCNetwork As New ADODB.Connection
Dim adoCnOLCNetwork As Object
Set adoCnOLCNetwork = CreateObject("ADODB.Connection")

If gcfHandleErrors Then On Error GoTo ErrLog


    If gcfHandleErrors Then On Error GoTo ErrLog
    SysCmd acSysCmdClearStatus
    varReturn = SysCmd(acSysCmdSetStatus, " SQL SERVER IS Executing: " & SQLStatement)
'    With adoCnOLCNetwork
'        .ConnectionTimeout = 10
'        .Provider = "sqloledb"
'        'http://www.microsoft.com/en-us/download/confirmation.aspx?id=23734
'        '.Provider = "Microsoft.ACE.OLEDB.12.0"
'        .Properties("Data Source").Value = gstrCashsysInsLoc 'adoCnOLCNetworkg adoCnOLCNetworkrsSysfile!CashsysInsLoc ' "KMDSSQL"
'        .Properties("Initial Catalog").Value = "KMDSSQL7"
'        'integrated security .Properties("User ID").Value = "sa"
'        'integrated security .Properties("Password").Value = ""
'        .ConnectionString = .ConnectionString & "; Trusted_Connection=Yes"
'
'        .Open
'    End With

   '4/11/2014
    With adoCnOLCNetwork
        .ConnectionTimeout = 10
        .ConnectionString = "Provider=SQLNCLI.1;" _
        & "SERVER=" & gstrCashsysInsLoc & ";" _
        & "Database=KMDSSQL7;" _
        & "DataTypeCompatibility=80;" _
        & "Trusted_Connection=Yes;"
        .Open
    End With


    With comQDF
        .ActiveConnection = adoCnOLCNetwork
        .CommandText = SQLStatement
        .CommandType = adCmdText ' = adCmdStoredProc
        .CommandTimeout = 60000
        .Execute

    End With
    Set comQDF = Nothing
    adoCnOLCNetwork.Close
    Set adoCnOLCNetwork = Nothing

ExitSub:

    SysCmd acSysCmdClearStatus
    Exit Sub
ErrLog:
    strMsg = "Error #" & Err.Number & " " & Err.Description & " : SQLPassThru : " & SQLStatement
    EOLog strMsg
    fMsgBox strMsg
    gbEOProcessError = True
    Resume ExitSub
End Sub




Public Function LotteryTicketsAppend()
    SQLPassThru "sp_LotteryScratch"

'Dim comQDF As New ADODB.Command
'Dim adoCnOLCNetwork As New ADODB.Connection
'
'
'    If gcfHandleErrors Then On Error GoTo ErrLog
'    With adoCnOLCNetwork
'        .ConnectionTimeout = 10
'        .Provider = "sqloledb"
'        .Properties("Data Source").Value = gstrCashsysInsLoc 'adoCnOLCNetworkg adoCnOLCNetworkrsSysfile!CashsysInsLoc ' "KMDSSQL"
'        .Properties("Initial Catalog").Value = "KMDSSQL"
'        'integrated security .Properties("User ID").Value = "sa"
'        'integrated security .Properties("Password").Value = ""
'        .ConnectionString = .ConnectionString & "; Trusted_Connection=Yes"
'        .Open
'    End With
'
'
'    With comQDF
'        .ActiveConnection = adoCnOLCNetwork
'        .CommandText = "sp_LotteryScratch"
'        .CommandType = adCmdStoredProc
'        .Execute
'    End With
'
'    adoCnOLCNetwork.Close
'
'ExitLog:
'    Set adoCnOLCNetwork = Nothing
'    Set comQDF = Nothing
'    Exit Function
'
'ErrLog:
'    MsgBox "Error #" & Err.Number & " " & Err.Description & vbcrlf & "LotteryTicketsAppend"
'    Resume ExitLog
'
End Function

Public Function MakeOnlineBackupBatchFile()

'added 2010-07-01 ... Create "OnlineBackup.bat for KMDSOLB.BAT online backup parameters "
    Dim rs As Recordset
    Dim strSQL As String
    Dim strFileName As String
    Dim strText As String
    Dim txtstrTemp As TextStream
    Dim fsoTemp As New FileSystemObject

    If gcfHandleErrors Then On Error GoTo ErrLog
    strSQL = "SELECT Command from [Switchboard Items]"
    strSQL = strSQL & "WHERE (((SwitchboardID)=1) AND ((ItemNumber)=0))"

    Set rs = CurrentDb().OpenRecordset(strSQL)
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        glnVersionNumber = rs!Command
        rs.Close
    End If
    Set rs = Nothing

    If gcfHandleErrors Then On Error GoTo ErrLog
    strFileName = "\\" & gstrCashsysInsLoc & "\KMDS\OnlineBackup.bat"

    If fsoTemp.FileExists(strFileName) Then
        fsoTemp.DeleteFile strFileName, True
    End If
    strText = "Rem Created by End of Day Process" & vbCrLf
    strText = strText & "c:\kmds\KMDSOLB.bat " & gstrCashsysInsLoc & " " & glnAgentNumber & " " & glnVersionNumber


    Set txtstrTemp = fsoTemp.OpenTextFile(strFileName, ForWriting, True, TristateFalse)
    txtstrTemp.WriteLine (strText)

    txtstrTemp.Close
    Set txtstrTemp = Nothing

'added 31416
    strFileName = "\\" & gstrCashsysInsLoc & "\KMDS\OnlineBackupWinSCP.bat"

    If fsoTemp.FileExists(strFileName) Then
        fsoTemp.DeleteFile strFileName, True
    End If
    strText = "Rem Created by End of Day Process" & vbCrLf
    strText = strText & "c:\kmds\KMDSOLBWinSCP.bat " & gstrCashsysInsLoc & " " & glnAgentNumber & " " & glnVersionNumber


    Set txtstrTemp = fsoTemp.OpenTextFile(strFileName, ForWriting, True, TristateFalse)
    txtstrTemp.WriteLine (strText)

    txtstrTemp.Close
    Set txtstrTemp = Nothing

    If gcfHandleErrors Then On Error GoTo ErrLog
    strFileName = "\\" & gstrCashsysInsLoc & "\KMDS\OnlineBackupWinSCP.txt"

    If fsoTemp.FileExists(strFileName) Then
        fsoTemp.DeleteFile strFileName, True
    End If

    strText = gstrftpLogonWinSCPKMDS & vbCrLf
    strText = strText & "cd /AgentOnlineBackup" & vbCrLf
    strText = strText & "send c:\kmds\backup\olbsql" & glnAgentNumber & glnVersionNumber & ".zip" & vbCrLf
    strText = strText & "exit"

    Set txtstrTemp = fsoTemp.OpenTextFile(strFileName, ForWriting, True, TristateFalse)
    txtstrTemp.WriteLine (strText)

    txtstrTemp.Close
    Set txtstrTemp = Nothing

'end added 31416


    Set fsoTemp = Nothing

ExitFunc:
    Exit Function
ErrLog:
    Resume ExitFunc


End Function


Public Function ReindexTables()
Dim varReturn As Variant
Dim strReindex As String

    If gcfHandleErrors Then On Error GoTo ErrLog
    SQLPassThru "dbcc DBREINDEX  ('ClubOrderGuides','',0)"
    SQLPassThru "dbcc DBREINDEX  ('PurchaseOrderLine','',0)"
    SQLPassThru "dbcc DBREINDEX  ('OrderLine','',0)"
    SQLPassThru "dbcc DBREINDEX  ('Item','',0)"
    SQLPassThru "dbcc DBREINDEX  ('UPCItem','',0)"
    SQLPassThru "dbcc DBREINDEX  ('Involine','',0)"

ExitSub:
    Exit Function

ErrLog:
    EOLog "          Error #" & Err.Number & " " & Err.Description & " : ReindexTables"
    MsgBox "Error #" & Err.Number & " " & Err.Description & vbCrLf & "ReindexTables"
    gbEOProcessError = True
    Resume ExitSub

End Function



Public Sub LookForDayEndFiles()

    '2006-12-5 'created crview.ins
    '               used by crview.exe for valid registers
    '10/03/2004 added "confirm" Parameter
    'Dim rstemp As New ADODB.Recordset
    Dim rstemp As Object
    Set rstemp = CreateObject("ADODB.Recordset")

    Dim bAtLeastOneRegisterCheckedOut
    Dim strValidRegisters As String

    Dim strCR1 As String
    Dim strCR2 As String
    Dim strCR3 As String
    Dim strCR4 As String
    Dim strCR5 As String

    Dim strMsg As String

    Dim fsoTemp As New FileSystemObject
    Dim filTemp As File

    strCR1 = "Not Active"
    strCR2 = "Not Active"
    strCR3 = "Not Active"
    strCR4 = "Not Active"
    strCR5 = "Not Active"

    If gcfHandleErrors Then On Error GoTo ExitSub

    'Cancel = FormLogin(glnCashierID, Me)

    bAtLeastOneRegisterCheckedOut = False
    rstemp.Open "qryCashRegisters", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    With rstemp
        Do While Not .EOF
            Select Case !ID
            Case 1
                If !Cr1CheckOut = True Then bAtLeastOneRegisterCheckedOut = True
                If !Cr1Installed = True Then
                    strCR1 = "\\" & !ComputerName & "\kmds\archive\"
                End If

            Case 2
                If !Cr2CheckOut = True Then bAtLeastOneRegisterCheckedOut = True
                If !Cr2Installed = True Then
                    strCR2 = "\\" & !ComputerName & "\kmds\archive\"
                End If
            Case 3
                If !Cr3CheckOut = True Then bAtLeastOneRegisterCheckedOut = True
                If !Cr3Installed = True Then
                    strCR3 = "\\" & !ComputerName & "\kmds\archive\"
                End If
            Case 4
                If !Cr4CheckOut = True Then bAtLeastOneRegisterCheckedOut = True
                If !Cr4Installed = True Then
                    strCR4 = "\\" & !ComputerName & "\kmds\archive\"
                End If
            Case 5
                If !Cr5CheckOut = True Then bAtLeastOneRegisterCheckedOut = True
                If !Cr5Installed = True Then
                    strCR5 = "\\" & !ComputerName & "\kmds\archive\"
                End If
            End Select
            .MoveNext
        Loop
        .Close
    End With
    Set rstemp = Nothing

    If Not bAtLeastOneRegisterCheckedOut Then
        MsgBox "There no registers that have been checked out!"
        Exit Sub
    End If
    strMsg = CheckRegister("1", strCR1) & "     " & strCR1 & vbCrLf
    strMsg = strMsg & CheckRegister("2", strCR2) & "     " & strCR2 & vbCrLf
    strMsg = strMsg & CheckRegister("3", strCR3) & "     " & strCR3 & vbCrLf
    strMsg = strMsg & CheckRegister("4", strCR4) & "     " & strCR4 & vbCrLf
    strMsg = strMsg & CheckRegister("5", strCR5) & "     " & strCR5 & vbCrLf


    MsgBox strMsg, , "LookForDayEndFiles"

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "LookForDayEndFiles"
    Resume ExitSub

End Sub


Public Function CheckRegister(RegisterNumber As String, RegisterArchiveDir As String) As String

    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim fil As File
    Dim strFname As String
    If gcfHandleErrors Then On Error GoTo ExitSub
    strFname = "CR" & RegisterNumber & ".MDB"
    'REGISTER 4
    If Not fso.FileExists("k:mail\" & strFname) Then
        If RegisterArchiveDir = "Not Active" Then
        'copy shell file
            If MsgBox(strFname & " not found; copy shell file?", vbYesNo) = vbYes Then
                CheckRegister = "CopyFile C:\KMDS\CashRegisterDetailTemplate.MDB to K:MAIL\" & strFname
                EOLog "CopyFile C:\KMDS\CashRegisterDetailTemplate.MDB to K:MAIL\" & strFname
                fso.CopyFile "C:\KMDS\CashRegisterDetailTemplate.MDB", "K:MAIL\" & strFname
            End If
        Else
            'copy archived file..
            Shell "Explorer " & RegisterArchiveDir, vbNormalFocus
            Set fil = fso.GetFile(RegisterArchiveDir & strFname)
            MsgBox RegisterArchiveDir & strFname & vbCrLf & "     Last Modified:" & fil.DateLastModified & vbCrLf & "      Current date: " & Now, vbCritical, "Register detail file not found!"
            If MsgBox("Do you want to copy this file to K:\MAIL ?", vbYesNo) = vbYes Then
                CheckRegister = "Copyfile " & RegisterArchiveDir & strFname & " to K:MAIL\" & strFname & "   .. Modified: " & fil.DateLastAccessed
                EOLog "Copyfile " & RegisterArchiveDir & strFname & " to K:MAIL\" & strFname & "   .. Modified: " & fil.DateLastAccessed
                fso.CopyFile RegisterArchiveDir & strFname, "K:MAIL\" & strFname
            End If
        End If
    Else
        CheckRegister = "  File found: k:kmds\" & strFname
    End If

ExitSub:
    Exit Function

ErrLog:
    CheckRegister = "Error #" & Err.Number & " " & Err.Description
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "CheckRegister"
    Resume ExitSub
End Function

Public Function WhatIsMyIp() As Boolean

Dim stAppName As String
Dim strFileName As String
Dim strWebSiteForIPAddress As String
Dim fsoTemp As New FileSystemObject

    If gcfHandleErrors Then On Error GoTo ErrLog
    stAppName = "C:\KMDS\KMDSGetWeb.exe"
    strFileName = "c:\kmds\kmdsgetweb.txt"
    If fsoTemp.FileExists(stAppName) Then
        If fsoTemp.FileExists(strFileName) Then
            fsoTemp.DeleteFile strFileName, True
        End If
        'stAppName = stAppName & " " & GetKMDSSettings("WebSiteForIPAddress", "http://www.whatismyip.com/automation/n09230945.asp")
        stAppName = stAppName & " " & GetKMDSSettings("WebSiteForIPAddress", "http://automation.whatismyip.com/n09230945.asp")
        WaitForExecCmd stAppName
        If fsoTemp.FileExists(strFileName) Then
            WhatIsMyIp = True
        Else
            MsgBox strFileName & " output file not found!", , "WhatIsMyIp"
            WhatIsMyIp = False
        End If
    Else
        MsgBox stAppName & " not found!", , "WhatIsMyIp"
        WhatIsMyIp = False
    End If

ExitSub:
    Exit Function

ErrLog:
    WhatIsMyIp = False
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "WhatIsMyIp"
    Resume ExitSub

End Function
