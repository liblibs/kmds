Attribute VB_Name = "modOLCTransmit"
Option Compare Database
Option Explicit

Public Sub OLCCInventFile(nAgentNumber As Long, Archive As Boolean)
Dim nOrderID As Long
Dim stDocName As String
Dim stLinkCriteria As String
'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")
'Dim nAgentNumber As Long
Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject
Dim strTemp As String
Dim nLineCount As Integer
Dim mstrNow As String

    If gcfHandleErrors Then On Error GoTo ErrLog
'    nAgentNumber = AgentNumber
    mstrNow = Format(Now, "yymmddhhnn")

    If nAgentNumber > 1000 Then nAgentNumber = nAgentNumber - 1000

    If Not fsoTemp.FolderExists("C:\NPOS") Then
        fsoTemp.CreateFolder ("C:\NPOS")
    End If
    If Not fsoTemp.FolderExists("C:\NPOS\Archive") Then
        fsoTemp.CreateFolder ("C:\NPOS\Archive")
    End If

    'add begin 31415
    If gbKMDSFTPActive Then ' if KMDSFTPActive = true - G-5 -> Settings -> KMDSFTPActive
        If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP") Then
            fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP")
        End If
        If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT") Then
            fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT")
        End If
        If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\Archive") Then
            fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\Archive")
        End If
    End If
    'add end 31415

    If Archive Then
        stDocName = "C:\NPOS\Archive\Invent " & Format(Date, "mm-dd") & "." & CStr(Format(nAgentNumber, "000"))
        'Set txtstrtemp = fsotemp.OpenTextFile("C:\NPOS\Archive\Invent " & Format(Date, "mm-dd") & "." & CStr(Format(nAgentNumber, "000")), ForWriting, True, TristateFalse)
    Else
        stDocName = "C:\NPOS\Invent." & CStr(Format(nAgentNumber, "000"))
        'Set txtstrtemp = fsotemp.OpenTextFile("C:\NPOS\Invent." & CStr(Format(nAgentNumber, "000")), ForWriting, True, TristateFalse)
    End If
    Debug.Print stDocName
    Set txtstrTemp = fsoTemp.OpenTextFile(stDocName, ForWriting, True, TristateFalse)

    nAgentNumber = nAgentNumber + 1000

    txtstrTemp.WriteLine (Build21Header(nAgentNumber, mstrNow))
    txtstrTemp.WriteLine (Build84Header(nAgentNumber, mstrNow))



    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryCreateTransmitInventory"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        Set rstemp = .Execute
    End With

    With rstemp
        If rstemp.RecordCount = 0 Then
            MsgBox "No Records found for Sales department 1"
            GoTo ExitSub
        End If

        Do While Not .EOF
            nLineCount = nLineCount + 1
            'If not gcfHandleErrors then Debug.Print !ItemID
            txtstrTemp.WriteLine BuildInvLine(nAgentNumber, !ID, !Onhand, mstrNow)
            .MoveNext
        Loop
    End With
    rstemp.Close
    txtstrTemp.WriteLine (Build84Trailer(nAgentNumber, nLineCount, mstrNow))
    txtstrTemp.WriteLine (Build21Trailer(nAgentNumber, nLineCount, mstrNow))
    txtstrTemp.Close
    'Shell ("Notepad C:\NPOS\Invent." & CStr(Format(nAgentNumber, "000")))

    If Archive Then
    'added 31415
        If gbKMDSFTPActive Then fsoTemp.CopyFile stDocName, "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\MonthlyInventAsOF" & Format(Now, "yy-mm-dd-hh-nn") & ".txt", True
    'end added
    Else
        'stDocName = "C:\NPOS\Invent." & CStr(Format(nAgentNumber, "000"))
        'Set txtstrtemp = fsotemp.OpenTextFile("C:\NPOS\Invent." & CStr(Format(nAgentNumber, "000")), ForWriting, True, TristateFalse)
        fsoTemp.CopyFile stDocName, "C:\NPOS\Inventory\Invent.txt", True
        If gbKMDSFTPActive Then fsoTemp.CopyFile stDocName, "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\", True

    End If

ExitSub:

    Set comQDF = Nothing
    Set rstemp = Nothing
    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing

    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "OLCCInventFile"
    Resume ExitSub
End Sub

Public Sub OLCCInventFileAfterNewMonthHasBegun(nAgentNumber As Long, Archive As Boolean)
Dim nOrderID As Long
Dim stDocName As String
Dim stLinkCriteria As String
'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")
'Dim nAgentNumber As Long
Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject
Dim strTemp As String
Dim nLineCount As Integer
Dim mstrNow As String

    If gcfHandleErrors Then On Error GoTo ErrLog
'    nAgentNumber = AgentNumber

    SetOptions False
    'Build a table of this month's transactions
    DoCmd.OpenQuery "qryRME_01_Delete_ResendMonthEnd_Table"
    DoCmd.OpenQuery "qryRME_02_Append_Current_Month_Transactions"

    mstrNow = Format(Now, "yymm01hhnn") 'force it to be the first day of current month

    If nAgentNumber > 1000 Then nAgentNumber = nAgentNumber - 1000

    If Not fsoTemp.FolderExists("C:\NPOS") Then
        fsoTemp.CreateFolder ("C:\NPOS")
    End If
    If Not fsoTemp.FolderExists("C:\NPOS\Archive") Then
        fsoTemp.CreateFolder ("C:\NPOS\Archive")
    End If

    'add begin 31415
    If gbKMDSFTPActive Then ' if KMDSFTPActive = true - G-5 -> Settings -> KMDSFTPActive
        If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP") Then
            fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP")
        End If
        If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT") Then
            fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT")
        End If
        If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\Archive") Then
            fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\Archive")
        End If
    End If
    'add end 31415

    If Archive Then
        stDocName = "C:\NPOS\Archive\Invent " & Format(Date, "mm-dd") & "." & CStr(Format(nAgentNumber, "000"))
        'Set txtstrtemp = fsotemp.OpenTextFile("C:\NPOS\Archive\Invent " & Format(Date, "mm-dd") & "." & CStr(Format(nAgentNumber, "000")), ForWriting, True, TristateFalse)
    Else
        stDocName = "C:\NPOS\Invent." & CStr(Format(nAgentNumber, "000"))
        'Set txtstrtemp = fsotemp.OpenTextFile("C:\NPOS\Invent." & CStr(Format(nAgentNumber, "000")), ForWriting, True, TristateFalse)
    End If
    Debug.Print stDocName
    Set txtstrTemp = fsoTemp.OpenTextFile(stDocName, ForWriting, True, TristateFalse)

    nAgentNumber = nAgentNumber + 1000

    txtstrTemp.WriteLine (Build21Header(nAgentNumber, mstrNow))
    txtstrTemp.WriteLine (Build84Header(nAgentNumber, mstrNow))



    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryRME_03_Last_Month_End"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        Set rstemp = .Execute
    End With

    With rstemp
        If rstemp.RecordCount = 0 Then
            MsgBox "No Records found for Sales department 1"
            GoTo ExitSub
        End If

        Do While Not .EOF
            nLineCount = nLineCount + 1
            'If not gcfHandleErrors then Debug.Print !ItemID
            txtstrTemp.WriteLine BuildInvLine(nAgentNumber, !ID, !MonthEndOnhand, mstrNow)
            .MoveNext
        Loop
    End With
    rstemp.Close
    txtstrTemp.WriteLine (Build84Trailer(nAgentNumber, nLineCount, mstrNow))
    txtstrTemp.WriteLine (Build21Trailer(nAgentNumber, nLineCount, mstrNow))
    txtstrTemp.Close
    'Shell ("Notepad C:\NPOS\Invent." & CStr(Format(nAgentNumber, "000")))

    If Archive Then
    'added 31415
        If gbKMDSFTPActive Then fsoTemp.CopyFile stDocName, "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\MonthlyInventAsOF" & Format(Now, "yy-mm-dd-hh-nn") & ".txt", True
    'end added
    Else
        'stDocName = "C:\NPOS\Invent." & CStr(Format(nAgentNumber, "000"))
        'Set txtstrtemp = fsotemp.OpenTextFile("C:\NPOS\Invent." & CStr(Format(nAgentNumber, "000")), ForWriting, True, TristateFalse)
        fsoTemp.CopyFile stDocName, "C:\NPOS\Inventory\Invent.txt", True
        If gbKMDSFTPActive Then fsoTemp.CopyFile stDocName, "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\", True

    End If

ExitSub:

    Set comQDF = Nothing
    Set rstemp = Nothing
    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing

    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "OLCCInventFile"
    Resume ExitSub
End Sub



Public Sub OLCCDailyInventFile(nAgentNumber As Long, Archive As Boolean)
Dim nOrderID As Long
Dim stDocName As String
Dim stLinkCriteria As String

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

'Dim nAgentNumber As Long
Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject
Dim strTemp As String
Dim nLineCount As Integer
Dim mstrNow As String

Dim mstrFileNameArchive As String
Dim mstrFileNameKMDSFTP As String

    If gcfHandleErrors Then On Error GoTo ErrLog
'    nAgentNumber = AgentNumber
    mstrNow = Format(Now, "yymmddhhnn")

    If nAgentNumber > 1000 Then nAgentNumber = nAgentNumber - 1000

    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL")
    End If
    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT")
    End If
    'added 31417
    If fsoTemp.FileExists("k:\mail\Out\Log") Then
        fsoTemp.DeleteFile ("k:\mail\Out\Log")
    End If
    If Not fsoTemp.FolderExists("k:\mail\Out\Log") Then
        fsoTemp.CreateFolder ("k:\mail\Out\Log")
    End If

    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\ARCHIVE") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\ARCHIVE")
    End If


    If Archive Then
        'Set txtstrtemp = fsotemp.OpenTextFile("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\ARCHIVE\DailyInv " & mstrNow & ".txt", ForWriting, True, TristateFalse)
        'above code commented out in 31415 version replaced with:

        mstrFileNameArchive = "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\ARCHIVE\DailyInv " & mstrNow & ".txt"
        mstrFileNameKMDSFTP = "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\DailyInventory " & mstrNow & ".txt"
        Set txtstrTemp = fsoTemp.OpenTextFile(mstrFileNameArchive, ForWriting, True, TristateFalse)
        'so as to store the archive file name which will be copied at the end of this routine...

    Else
        Set txtstrTemp = fsoTemp.OpenTextFile("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\DailyInv" & Left(mstrNow, 6) & ".txt", ForWriting, True, TristateFalse)
    End If


    nAgentNumber = nAgentNumber + 1000

    txtstrTemp.WriteLine (Build21Header(nAgentNumber, mstrNow))
    txtstrTemp.WriteLine (Build63Header(nAgentNumber, mstrNow))



    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryCreateTransmitInventory"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        Set rstemp = .Execute
    End With

    With rstemp
        If rstemp.RecordCount = 0 Then
            MsgBox "No Records found for Sales department 1"
            GoTo ExitSub
        End If

        Do While Not .EOF
            nLineCount = nLineCount + 1
            'If not gcfHandleErrors then Debug.Print !ItemID
            txtstrTemp.WriteLine BuildDailyInvLine(nAgentNumber, !ID, !Onhand, mstrNow)
            .MoveNext
        Loop
    End With
    rstemp.Close
    txtstrTemp.WriteLine (Build63Trailer(nAgentNumber, nLineCount, mstrNow))
    txtstrTemp.WriteLine (Build21Trailer(nAgentNumber, nLineCount, mstrNow))
    txtstrTemp.Close
    'Shell ("Notepad C:\NPOS\Invent." & CStr(Format(nAgentNumber, "000")))

    ' BEGIN added     version 31405
    If Archive Then
        If gbKMDSFTPActive Then fsoTemp.CopyFile mstrFileNameArchive, mstrFileNameKMDSFTP
    End If
    'END
ExitSub:

    Set comQDF = Nothing
    Set rstemp = Nothing
    Set txtstrTemp = Nothing

    Set fsoTemp = Nothing

    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "OLCCDailyInventFile"
    Resume ExitSub
End Sub

Public Function OLCCDailySalesFile(nAgentNumber As Long, Archive As Boolean)
Dim nOrderID As Long
Dim stDocName As String
Dim stLinkCriteria As String

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

'Dim nAgentNumber As Long
Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject
Dim strTemp As String
Dim nLineCount As Integer
Dim mstrNow As String
Dim Build32 As String
Dim TenderCode As String


Dim mstrFileNameArchive As String
Dim mstrFileNameKMDSFTP As String

'filename is same as business date so you must inspect the
    If gcfHandleErrors Then On Error GoTo ErrLog
'    nAgentNumber = AgentNumber
    mstrNow = Format(Now, "yymmddhhnn")

    If nAgentNumber > 1000 Then nAgentNumber = nAgentNumber - 1000

    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL")
    End If
    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT")
    End If
    If Not fsoTemp.FolderExists("k:\mail\Out\Log") Then
        fsoTemp.CreateFolder ("k:\mail\Out\Log")
    End If

    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\ARCHIVE") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\ARCHIVE")
    End If
    If Archive Then
        'Set txtstrtemp = fsotemp.OpenTextFile("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\ARCHIVE\Sales " & mstrNow & ".txt", ForWriting, True, TristateFalse)
        'above code commented out in 31415 version replaced with:

        mstrFileNameArchive = "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\ARCHIVE\Sales " & mstrNow & ".txt"
        mstrFileNameKMDSFTP = "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\DailySales " & mstrNow & ".txt"
        Set txtstrTemp = fsoTemp.OpenTextFile(mstrFileNameArchive, ForWriting, True, TristateFalse)
        'so as to store the archive file name which will be copied at the end of this routine...

    Else
        Set txtstrTemp = fsoTemp.OpenTextFile("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\Sales" & Left(mstrNow, 6) & ".txt", ForWriting, True, TristateFalse)
    End If

    nAgentNumber = nAgentNumber + 1000

    txtstrTemp.WriteLine (Build21Header(nAgentNumber, mstrNow))
    txtstrTemp.WriteLine (Build32Header(nAgentNumber, mstrNow))

    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryCreateTransmitDailySales"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        Set rstemp = .Execute
    End With

    With rstemp
        If rstemp.RecordCount = 0 Then
            MsgBox "No Records found for Sales department 1"
            GoTo ExitSub
        End If

        Do While Not .EOF
            nLineCount = nLineCount + 1
            'If not gcfHandleErrors then Debug.Print !ItemID
            TenderCode = !TenderCode
            If IsNull(!TenderCode) Then
                TenderCode = 1
            Else
                If !TenderCode > 3 Then
                    TenderCode = 1
                End If
            End If
            Build32 = "32"
            Build32 = Build32 & !CreateDate & !CreateTime & nAgentNumber
            Build32 = Build32 & !OLCCItemCode & !OLCCQuantity & !RegisterDate & !RegisterTime
            Build32 = Build32 & !CashSysID & !TransactionNumber & IIf(!PN = "", "          ", !PN)
            Build32 = Build32 & IIf(!lin = "", "          ", !lin) & !UPCCode & !SELLINGPRICE & TenderCode
            txtstrTemp.WriteLine Build32
            .MoveNext
        Loop
    End With
    rstemp.Close
    txtstrTemp.WriteLine (Build32Trailer(nAgentNumber, nLineCount, mstrNow))
    txtstrTemp.WriteLine (Build21Trailer(nAgentNumber, nLineCount, mstrNow))
    txtstrTemp.Close
    ' BEGIN added     version 31405
    If Archive Then
        If gbKMDSFTPActive Then fsoTemp.CopyFile mstrFileNameArchive, mstrFileNameKMDSFTP
    End If
    'END
    'Shell ("Notepad C:\NPOS\Invent." & CStr(Format(nAgentNumber, "000")))
ExitSub:

    Set comQDF = Nothing
    Set rstemp = Nothing
    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing

    Exit Function

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "OLCCDailySalesFile"
    Resume ExitSub
End Function


Public Function OLCCDailySalesFileReCreate(nAgentNumber As Long, Archive As Boolean, ReCreateDate As Date)
Dim nOrderID As Long
Dim stDocName As String
Dim stLinkCriteria As String
Dim mstrFileNameArchive As String
Dim mstrFileNameKMDSFTP As String



'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")


'Dim nAgentNumber As Long
Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject
Dim strTemp As String
Dim nLineCount As Integer
Dim mstrNow As String
Dim Build32 As String
Dim TenderCode As String

'filename is same as business date so you must inspect the
    If gcfHandleErrors Then On Error GoTo ErrLog
'    nAgentNumber = AgentNumber
    mstrNow = Format(ReCreateDate, "yymmddhhnn")

    If nAgentNumber > 1000 Then nAgentNumber = nAgentNumber - 1000

    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL")
    End If
    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT")
    End If
    If Not fsoTemp.FolderExists("k:\mail\Out\Log") Then
        fsoTemp.CreateFolder ("k:\mail\Out\Log")
    End If

    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\ARCHIVE") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\ARCHIVE")
    End If

    If Archive Then
        'Set txtstrtemp = fsotemp.OpenTextFile("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\ARCHIVE\Sales " & mstrNow & ".txt", ForWriting, True, TristateFalse)
        'above code commented out in 31415 version replaced with:

        mstrFileNameArchive = "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\ARCHIVE\Sales " & mstrNow & ".txt"
        mstrFileNameKMDSFTP = "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\DailySales " & mstrNow & ".txt"
        Set txtstrTemp = fsoTemp.OpenTextFile(mstrFileNameArchive, ForWriting, True, TristateFalse)
        'so as to store the archive file name which will be copied at the end of this routine...
    Else
        Set txtstrTemp = fsoTemp.OpenTextFile("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\Sales" & Left(mstrNow, 6) & ".txt", ForWriting, True, TristateFalse)
    End If

    nAgentNumber = nAgentNumber + 1000

    txtstrTemp.WriteLine (Build21Header(nAgentNumber, mstrNow))
    txtstrTemp.WriteLine (Build32Header(nAgentNumber, mstrNow))

    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryCreateTransmitDailySales"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        Set rstemp = .Execute
    End With

    With rstemp
        If rstemp.RecordCount = 0 Then
            MsgBox "No Records found for Sales department 1"
            GoTo ExitSub
        End If

        Do While Not .EOF
            nLineCount = nLineCount + 1
            'If not gcfHandleErrors then Debug.Print !ItemID
            TenderCode = !TenderCode
            If IsNull(!TenderCode) Then
                TenderCode = 1
            Else
                If !TenderCode > 3 Then
                    TenderCode = 1
                End If
            End If
            Build32 = "32"
            Build32 = Build32 & !CreateDate & !CreateTime & nAgentNumber
            Build32 = Build32 & !OLCCItemCode & !OLCCQuantity & !RegisterDate & !RegisterTime
            Build32 = Build32 & !CashSysID & !TransactionNumber & IIf(!PN = "", "          ", !PN)
            Build32 = Build32 & IIf(!lin = "", "          ", !lin) & !UPCCode & !SELLINGPRICE & TenderCode
            txtstrTemp.WriteLine Build32
            .MoveNext
        Loop
    End With
    rstemp.Close
    txtstrTemp.WriteLine (Build32Trailer(nAgentNumber, nLineCount, mstrNow))
    txtstrTemp.WriteLine (Build21Trailer(nAgentNumber, nLineCount, mstrNow))
    txtstrTemp.Close
        ' BEGIN added     version 31405
    If Archive Then
        CreateWinSCPtoKMDSFTPScriptFile
        If gbKMDSFTPActive Then fsoTemp.CopyFile mstrFileNameArchive, mstrFileNameKMDSFTP
    End If
    'END

    'Shell ("Notepad C:\NPOS\Invent." & CStr(Format(nAgentNumber, "000")))
ExitSub:


    Set comQDF = Nothing
    Set rstemp = Nothing
    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing

    Exit Function

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "OLCCDailySalesFile"
    Resume ExitSub
End Function



Public Function Build21Header(AgentNumber As Long, mstrNow As String) As String
'0004030514441103OLF00021.FIL  00
    Build21Header = "00" & mstrNow & Format(CStr(AgentNumber), "0000")
    Build21Header = Build21Header & "OLF00021.FIL  00"
End Function

Public Function Build15Header(AgentNumber As Long, mstrNow As String) As String
'0004030514441103OLF00015.FIL  00
    Build15Header = "00" & mstrNow & Format(CStr(AgentNumber), "0000")
    Build15Header = Build15Header & "OLF00015.FIL  00"
End Function

Public Function Build32Header(AgentNumber As Long, mstrNow As String) As String
'0004030514441103OLF00015.FIL  00
    Build32Header = "00" & mstrNow & Format(CStr(AgentNumber), "0000")
    Build32Header = Build32Header & "OLF00032.FIL  00"
End Function
Public Function Build63Header(AgentNumber As Long, mstrNow As String) As String
'0004030514441103OLF00015.FIL  00
    Build63Header = "00" & mstrNow & Format(CStr(AgentNumber), "0000")
    Build63Header = Build63Header & "OLF00063.FIL  00"
End Function


Public Function Build84Header(AgentNumber As Long, mstrNow As String) As String
'0004030514441103OLF00015.FIL  00
    Build84Header = "00" & mstrNow & Format(CStr(AgentNumber), "0000")
    Build84Header = Build84Header & "OLF00084.FIL  00"
End Function

Public Function Build15Trailer(AgentNumber As Long, LineCount As Integer, mstrNow As String) As String
'9904030514441103OLF00015.FIL00022600000099
    Build15Trailer = "99" & mstrNow & Format(CStr(AgentNumber), "0000")
    Build15Trailer = Build15Trailer & "OLF00015.FIL"
    Build15Trailer = Build15Trailer & Format(LineCount, "000000")
    Build15Trailer = Build15Trailer & "00000099"
End Function

Public Function Build21Trailer(AgentNumber As Long, LineCount As Integer, mstrNow As String) As String
'9904030514441103OLF00021.FIL00022600000099
    Build21Trailer = "99" & mstrNow & Format(CStr(AgentNumber), "0000")
    Build21Trailer = Build21Trailer & "OLF00021.FIL"
    Build21Trailer = Build21Trailer & Format(LineCount, "000000")
    Build21Trailer = Build21Trailer & "00000099"
End Function

Public Function Build32Trailer(AgentNumber As Long, LineCount As Integer, mstrNow As String) As String
    Build32Trailer = "99" & mstrNow & Format(CStr(AgentNumber), "0000")
    Build32Trailer = Build32Trailer & "OLF00032.FIL"
    Build32Trailer = Build32Trailer & Format(LineCount, "000000")
    Build32Trailer = Build32Trailer & "00000099"
End Function

Public Function Build63Trailer(AgentNumber As Long, LineCount As Integer, mstrNow As String) As String
    Build63Trailer = "99" & mstrNow & Format(CStr(AgentNumber), "0000")
    Build63Trailer = Build63Trailer & "OLF00063.FIL"
    Build63Trailer = Build63Trailer & Format(LineCount, "000000")
    Build63Trailer = Build63Trailer & "00000099"
End Function

Public Function Build84Trailer(AgentNumber As Long, LineCount As Integer, mstrNow As String) As String
'9904030514441103OLF00015.FIL00022600000099
    Build84Trailer = "99" & mstrNow & Format(CStr(AgentNumber), "0000")
    Build84Trailer = Build84Trailer & "OLF00084.FIL"
    Build84Trailer = Build84Trailer & Format(LineCount, "000000")
    Build84Trailer = Build84Trailer & "00000099"
End Function


Public Function BuildPOLine(AgentNumber As Long, ItemID As Double, Quantity As Integer, mstrNow As String) As String
'05040305144411030013120001200000
'05040305144411030013220001200000
   BuildPOLine = "05" & mstrNow & Format(CStr(AgentNumber), "0000")
   BuildPOLine = BuildPOLine & Format(CStr(ItemID), "00000000000")
   BuildPOLine = BuildPOLine & Format(CStr(Quantity), "00000")
   BuildPOLine = BuildPOLine & "00000"

End Function

Public Function BuildInvLine(AgentNumber As Long, ItemID As Double, Quantity As Integer, mstrNow As String) As String
'05040305144411030013120001200000
'05040305144411030013220001200000
   BuildInvLine = "84" & mstrNow & Format(CStr(AgentNumber), "0000")
   BuildInvLine = BuildInvLine & Format(CStr(ItemID), "00000000000")
   BuildInvLine = BuildInvLine & Format(CStr(Quantity), "00000")
End Function

Public Function BuildDailyInvLine(AgentNumber As Long, ItemID As Double, Quantity As Integer, mstrNow As String) As String
   BuildDailyInvLine = "63" & mstrNow & Format(CStr(AgentNumber), "0000")
   BuildDailyInvLine = BuildDailyInvLine & Format(CStr(ItemID), "00000000000")
   BuildDailyInvLine = BuildDailyInvLine & Format(CStr(Quantity), "00000")
   BuildDailyInvLine = BuildDailyInvLine & Left(mstrNow, 6)
End Function

Public Sub OLCCCreateDailyTransmitFiles()
Dim strTemp As String
Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject

'filename is same as business date so you must inspect the
    If gcfHandleErrors Then On Error GoTo ErrLog

    If glnAgentNumber = 0 Then LoadDefaults
    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL")
    End If
    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT")
    End If
    If Not fsoTemp.FolderExists("k:\mail\Out\Log") Then
        fsoTemp.CreateFolder ("k:\mail\Out\Log")
    End If



    Set txtstrTemp = fsoTemp.OpenTextFile("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\FTPSales.txt", ForWriting, True, TristateFalse)
    strTemp = gstrFTPLogon
    strTemp = strTemp & "MPUT" & vbTab & "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\" & vbTab & "Sales*.txt" & vbTab & IIf(gbOLCCFTPAscii, 1, 0) & vbTab & 1
    txtstrTemp.WriteLine strTemp
    txtstrTemp.Close

    Set txtstrTemp = fsoTemp.OpenTextFile("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\FTPDailyInv.txt", ForWriting, True, TristateFalse)
    strTemp = gstrFTPLogon
    strTemp = strTemp & "MPUT" & vbTab & "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\" & vbTab & "DailyInv*.txt" & vbTab & IIf(gbOLCCFTPAscii, 1, 0) & vbTab & 1
    txtstrTemp.WriteLine strTemp
    txtstrTemp.Close

    'added for 31415
    CreateWinSCPtoKMDSFTPScriptFile

    OLCCDailyInventFile CLng(glnAgentNumber), False
    OLCCDailyInventFile CLng(glnAgentNumber), True
    OLCCDailySalesFile CLng(glnAgentNumber), False
    OLCCDailySalesFile CLng(glnAgentNumber), True
ExitSub:

    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing

    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "OLCCDailySalesFile"
    Resume ExitSub

End Sub



Public Sub OLCCDailyTransmit()
    If (GetKMDSSettings("UseWinSCPForOLASTransmits", "False")) = True Then
    'begin Added to 31415
        If gbKMDSFTPActive Then
            EOLog "   Sending KMDS Daily Transmit Files"
            WinSCPFTPTransmitKMDS
            EOLog "...Complete Sending KMDS "
        End If
        'build Dayend files
            '1. WinSCPDaily.txt 'script file used by :
            '2. WinSCPDaily.bat
        CreateWinSCPtoOLASFTPScriptFile ("DayEnd")
        WaitForExecCmd "k:\MAIL\OUT\WinSCPDaily.BAT"
    Else
        WaitForExecCmd "\\" & gstrCashsysInsLoc & "\kmds\kmdsftp.exe \\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\FTPSales.txt" ', vbNormalFocus
        WaitForExecCmd "\\" & gstrCashsysInsLoc & "\kmds\kmdsftp.exe \\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\FTPDailyInv.txt"    ', vbNormalFocus
    End If
End Sub



Public Sub OLCCDailyInventFileRecreate(nAgentNumber As Long, Archive As Boolean, ReCreateDate As Date)
Dim nOrderID As Long
Dim stDocName As String
Dim stLinkCriteria As String

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")


'Dim nAgentNumber As Long
Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject
Dim strTemp As String
Dim nLineCount As Integer
Dim mstrNow As String
Dim mstrFileNameArchive As String
Dim mstrFileNameKMDSFTP As String


    If gcfHandleErrors Then On Error GoTo ErrLog
'    nAgentNumber = AgentNumber
    mstrNow = Format(ReCreateDate, "yymmddhhnn")

    If nAgentNumber > 1000 Then nAgentNumber = nAgentNumber - 1000

    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL")
    End If
    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT")
    End If
    If Not fsoTemp.FolderExists("k:\mail\Out\Log") Then
        fsoTemp.CreateFolder ("k:\mail\Out\Log")
    End If

    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\ARCHIVE") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\ARCHIVE")
    End If

    If Archive Then
        'Set txtstrTemp = fsoTemp.OpenTextFile("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\ARCHIVE\DailyInv " & mstrNow & ".txt", ForWriting, True, TristateFalse)
        'above code commented out in 31415 version replaced with:

        mstrFileNameArchive = "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\ARCHIVE\DailyInv " & mstrNow & ".txt"
        mstrFileNameKMDSFTP = "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\DailyInventory " & mstrNow & ".txt"
        Set txtstrTemp = fsoTemp.OpenTextFile(mstrFileNameArchive, ForWriting, True, TristateFalse)
        'so as to store the archive file name which will be copied at the end of this routine...

    Else
        Set txtstrTemp = fsoTemp.OpenTextFile("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\DailyInv" & Left(mstrNow, 6) & ".txt", ForWriting, True, TristateFalse)
    End If


    nAgentNumber = nAgentNumber + 1000

    txtstrTemp.WriteLine (Build21Header(nAgentNumber, mstrNow))
    txtstrTemp.WriteLine (Build63Header(nAgentNumber, mstrNow))



    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryCreateTransmitInventoryRecreate"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        Set rstemp = .Execute
    End With

    With rstemp
        If rstemp.RecordCount = 0 Then
            MsgBox "No Records found for Sales department 1"
            GoTo ExitSub
        End If

        Do While Not .EOF
            nLineCount = nLineCount + 1
            'If not gcfHandleErrors then Debug.Print !ItemID
            txtstrTemp.WriteLine BuildDailyInvLine(nAgentNumber, !ID, !AvgTurns, mstrNow)
            .MoveNext
        Loop
    End With
    rstemp.Close
    txtstrTemp.WriteLine (Build63Trailer(nAgentNumber, nLineCount, mstrNow))
    txtstrTemp.WriteLine (Build21Trailer(nAgentNumber, nLineCount, mstrNow))
    txtstrTemp.Close

        ' BEGIN added     version 31405
    If Archive Then
        If gbKMDSFTPActive Then fsoTemp.CopyFile mstrFileNameArchive, mstrFileNameKMDSFTP
    End If
    'END
    'Shell ("Notepad C:\NPOS\Invent." & CStr(Format(nAgentNumber, "000")))
ExitSub:

    Set comQDF = Nothing
    Set rstemp = Nothing
    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing

    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "OLCCDailyInventFile"
    Resume ExitSub
End Sub

Public Sub OLCCCreateOLCCItemsTransmitFiles()
Dim strTemp As String
Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject

'filename is same as business date so you must inspect the
    If gcfHandleErrors Then On Error GoTo ErrLog

    If glnAgentNumber = 0 Then LoadDefaults
    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL")
    End If
    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\IN") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\IN")
    End If

    Set txtstrTemp = fsoTemp.OpenTextFile("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\FTPOLCCItems.txt", ForWriting, True, TristateFalse)
    strTemp = gstrFTPLogon
    strTemp = strTemp & "MGET" & vbTab & "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\IN\" & vbTab & "Sales*.txt" & vbTab & IIf(gbOLCCFTPAscii, 1, 0) & vbTab & 1
    txtstrTemp.WriteLine strTemp
    txtstrTemp.Close

    Set txtstrTemp = fsoTemp.OpenTextFile("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\FTPDailyInv.txt", ForWriting, True, TristateFalse)
    strTemp = gstrFTPLogon
    strTemp = strTemp & "MPUT" & vbTab & "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\" & vbTab & "DailyInv*.txt" & vbTab & IIf(gbOLCCFTPAscii, 1, 0) & vbTab & 1
    txtstrTemp.WriteLine strTemp
    txtstrTemp.Close

    OLCCDailyInventFile CLng(glnAgentNumber), False
    OLCCDailyInventFile CLng(glnAgentNumber), True
    OLCCDailySalesFile CLng(glnAgentNumber), False
    OLCCDailySalesFile CLng(glnAgentNumber), True
ExitSub:

    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing

    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "OLCCCreateOLCCItemsTransmitFiles"
    Resume ExitSub

End Sub


Public Sub GetOLCCItemFiles()
'OLCC provides the master files on the ftp website..
' this routine gets the files and imports them into SQL tables

    Dim varReturn As Variant


    If gcfHandleErrors Then On Error Resume Next
    If Not gcfHandleErrors Then On Error Resume Next
    SetOptions False
    'step 1 delete local tables
    SysCmd acSysCmdClearStatus
    varReturn = SysCmd(acSysCmdSetStatus, "Step 1. Delete Local Tables")

    DoCmd.RunSQL ("Delete * from Item_ICExpansion")
    DoCmd.RunSQL ("Delete * from OLCCScanCode")

    If gcfHandleErrors Then On Error GoTo ErrLog
    'step 2 get source files from olcc ftp server
    SysCmd acSysCmdClearStatus
    varReturn = SysCmd(acSysCmdSetStatus, "Step 2. get source files from olcc ftp server")
    'added 31415
    If (GetKMDSSettings("UseWinSCPForOLASTransmits", "False")) = True Then
        'build a Dayend files
        '1. WinSCPDaily.txt 'script file used by :
        '2. WinSCPDaily.bat
        CreateWinSCPtoOLASFTPScriptFile ("ItemFile")
        WaitForExecCmd "c:\npos\WinSCPItemFile.BAT"
    Else
        If Not OLCCFTPGet("item_ICExpansion.txt", "c:\NPOS\item_ICExpansion.txt") Then Exit Sub
        'If Not OLCCFTPGet("scancode_ICExpansion.txt", "c:\NPOS\scancode_ICExpansion.txt") Then Exit Sub

    End If


    'step 3 transfer source files to local tables
    SysCmd acSysCmdClearStatus
    varReturn = SysCmd(acSysCmdSetStatus, "Step 3 transfer source files to local tables")

    DoCmd.TransferText acImportFixed, "OLCCItemImportSpecification", "Item_ICExpansion", "c:\NPOS\item_ICExpansion.txt", False
    DoCmd.OpenQuery "qryqryOLCCItemAppendRemoveApostrophe"

    '31508 this never worked and took too long
    'DoCmd.TransferText acImportFixed, "OLCCScanCodeImportSpecification", "OLCCScanCode", "c:\NPOS\scancode_ICExpansion.txt", False

    'step 4 delete sql rows from sql tables
    SysCmd acSysCmdClearStatus
    varReturn = SysCmd(acSysCmdSetStatus, "Step 4 delete sql rows from sql tables")

    SQLPassThru "sp_OLCCItems_Step_1_Delete"
    '31508 this never worked and took too long
    'SQLPassThru "sp_OLCCItems_Step_2_Delete_UPC"

    'step 5 transfer local tables to SQL Tables
    SysCmd acSysCmdClearStatus
    varReturn = SysCmd(acSysCmdSetStatus, "Step 5 transfer local tables to SQL Tables")

    SetOptions False
    DoCmd.OpenQuery "qryOLCCItemAppend"
    '31508 this never worked and took too long
    'DoCmd.OpenQuery "qryOLCCItemAppendUPC"
    varReturn = SysCmd(acSysCmdSetStatus, "Step 6 Parse UPC Codes")

    '31508 this never worked and took too long
    'DoCmd.OpenQuery "qrydbo_OLCCUPCItem_01_Append_12_From_10"
    '31508 this never worked and took too long
    'DoCmd.OpenQuery "qrydbo_OLCCUPCItem_02_Append_8_From_6"
    '31508 this never worked and took too long
    'DoCmd.OpenQuery "qrydbo_OLCCUPCItem_03_Update_UPCCodeKMDS"
    '31508 this never worked and took too long
    'DoCmd.OpenQuery "qrydbo_OLCCUPCItemUpdate_04_DeleteTheRest"

    'step 6 Assign Items stock number with stock status
    SQLPassThru "sp_OLCCItems_Step_3_Assign_Stock_Status"

    'step 7 Assign Items Vendor = 0 with stock number "X" or "S"
    'REM Disable due to unpopular demand
    'SQLPassThru "sp_OLCCItems_Step_4_Assign_Vendor_0"

    If GetKMDSSettings("UpdateUnitsPerCaseFromICE", "True") = "True" Then
        DoCmd.OpenQuery "qryOLCCitemUpdateUnitsPerCase"
    End If

ExitSub:
    SetOptions True
    SysCmd acSysCmdClearStatus
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, vbCritical, "GetOLCCItemFiles"
    Resume ExitSub

End Sub



Public Sub GetNabcaOLCCItemFiles()
'OLCC provides the master files on the ftp website..
' this routine gets the files and imports them into SQL tables

    Dim varReturn As Variant


    If gcfHandleErrors Then On Error Resume Next
    If Not gcfHandleErrors Then On Error Resume Next
     SetOptions False

    'step 1 delete local tables
    SysCmd acSysCmdClearStatus
    varReturn = SysCmd(acSysCmdSetStatus, "Step 1. Delete Local Tables")
    DoCmd.OpenQuery "qryNabcaOLCCItemImport_01_Delete_NabcaOlccItem", acViewNormal, acEdit
'    DoCmd.DeleteObject acTable, "OLCCItem"

'    DoCmd.DeleteObject acTable, "OLCCScanCode"

'    If gcfHandleErrors Then On Error GoTo ErrLog
'    'step 2 get source files from olcc ftp server
'    SysCmd acSysCmdClearStatus
'    varReturn = SysCmd(acSysCmdSetStatus, "Step 2. get source files from olcc ftp server")
'    'added 31415
'    If (GetKMDSSettings("UseWinSCPForOLASTransmits", "False")) = True Then
'        'build a Dayend files
'        '1. WinSCPDaily.txt 'script file used by :
'        '2. WinSCPDaily.bat
'        CreateWinSCPtoOLASFTPScriptFile ("ItemFile")
'        WaitForExecCmd "c:\npos\WinSCPItemFile.BAT"
'    Else
'        If Not OLCCFTPGet("item.txt", "\\" & gstrCashsysInsLoc & "\NPOS\item.txt") Then Exit Sub
'        If Not OLCCFTPGet("scancode.txt", "\\" & gstrCashsysInsLoc & "\NPOS\scancode.txt") Then Exit Sub
'
'    End If


    'step 3 transfer source files to local tables
'    SysCmd acSysCmdClearStatus
'    varReturn = SysCmd(acSysCmdSetStatus, "Step 3 transfer source files to local tables")
'
    DoCmd.TransferText acImportFixed, "NabcaOLCCItemImportSpecification", "NabcaOLCCitem", "\\" & gstrCashsysInsLoc & "\NPOS\NABCAitem.txt", False
    'update NabcaOLCCItem.ItemID 140B -> 1402
    DoCmd.OpenQuery "qryNabcaOLCCItemImport_02_Update_A", acViewNormal, acEdit
    DoCmd.OpenQuery "qryNabcaOLCCItemImport_03_Update_B", acViewNormal, acEdit
    DoCmd.OpenQuery "qryNabcaOLCCItemImport_04_Update_D", acViewNormal, acEdit
    DoCmd.OpenQuery "qryNabcaOLCCItemImport_05_Update_E", acViewNormal, acEdit
    DoCmd.OpenQuery "qryNabcaOLCCItemImport_06_Update_F", acViewNormal, acEdit
    DoCmd.OpenQuery "qryNabcaOLCCItemImport_07_Update_H", acViewNormal, acEdit

''    DoCmd.TransferText acImportFixed, "OLCCScanCodeImportSpecification", "OLCCScanCode", "\\" & gstrCashsysInsLoc & "\NPOS\scancode.txt", False
'
'    'step 4 delete sql rows from sql tables
'    SysCmd acSysCmdClearStatus
'    varReturn = SysCmd(acSysCmdSetStatus, "Step 4 delete sql rows from sql tables")
'
'    SQLPassThru "sp_OLCCItems_Step_1_Delete"
'    SQLPassThru "sp_OLCCItems_Step_2_Delete_UPC"
'
'    'step 5 transfer local tables to SQL Tables
'    SysCmd acSysCmdClearStatus
'    varReturn = SysCmd(acSysCmdSetStatus, "Step 5 transfer local tables to SQL Tables")
'
'    SetOptions False
'    DoCmd.OpenQuery "qryOLCCItemAppend"
'    DoCmd.OpenQuery "qryOLCCItemAppendUPC"
'    varReturn = SysCmd(acSysCmdSetStatus, "Step 6 Parse UPC Codes")
'
'    SetOptions False
'
'    DoCmd.OpenQuery "qrydbo_OLCCUPCItem_01_Append_12_From_10"
'    DoCmd.OpenQuery "qrydbo_OLCCUPCItem_02_Append_8_From_6"
'    DoCmd.OpenQuery "qrydbo_OLCCUPCItem_03_Update_UPCCodeKMDS"
'    DoCmd.OpenQuery "qrydbo_OLCCUPCItemUpdate_04_DeleteTheRest"
'
'    'step 6 Assign Items stock number with stock status
'    SQLPassThru "sp_OLCCItems_Step_3_Assign_Stock_Status"
'
'    'step 7 Assign Items Vendor = 0 with stock number "X" or "S"
'    'REM Disable due to unpopular demand
'    'SQLPassThru "sp_OLCCItems_Step_4_Assign_Vendor_0"

ExitSub:
    SetOptions True
    SysCmd acSysCmdClearStatus
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, vbCritical, "GetOLCCItemFiles"
    Resume ExitSub

End Sub

Public Sub NabcaOLCCInventFile(nAgentNumber As Long, Archive As Boolean)
Dim nOrderID As Long
Dim stDocName As String
Dim stLinkCriteria As String
'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")
'Dim nAgentNumber As Long
Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject
Dim strTemp As String
Dim nLineCount As Integer
Dim mstrNow As String

    If gcfHandleErrors Then On Error GoTo ErrLog
'    nAgentNumber = AgentNumber
    mstrNow = Format(Now, "yymmddhhnn")

    If nAgentNumber > 1000 Then nAgentNumber = nAgentNumber - 1000

    If Not fsoTemp.FolderExists("C:\NPOS") Then
        fsoTemp.CreateFolder ("C:\NPOS")
    End If
    If Not fsoTemp.FolderExists("C:\NPOS\Archive") Then
        fsoTemp.CreateFolder ("C:\NPOS\Archive")
    End If

    'add begin 31415
    If gbKMDSFTPActive Then ' if KMDSFTPActive = true - G-5 -> Settings -> KMDSFTPActive
        If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP") Then
            fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP")
        End If
        If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT") Then
            fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT")
        End If
        If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\Archive") Then
            fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\Archive")
        End If
    End If
    'add end 31415

    If Archive Then
        stDocName = "C:\NPOS\Archive\NInvent " & Format(Date, "mm-dd") & "." & CStr(Format(nAgentNumber, "000"))
        'Set txtstrtemp = fsotemp.OpenTextFile("C:\NPOS\Archive\Invent " & Format(Date, "mm-dd") & "." & CStr(Format(nAgentNumber, "000")), ForWriting, True, TristateFalse)
    Else
        stDocName = "C:\NPOS\NInvent." & CStr(Format(nAgentNumber, "000"))
        'Set txtstrtemp = fsotemp.OpenTextFile("C:\NPOS\Invent." & CStr(Format(nAgentNumber, "000")), ForWriting, True, TristateFalse)
    End If
    Set txtstrTemp = fsoTemp.OpenTextFile(stDocName, ForWriting, True, TristateFalse)

    nAgentNumber = nAgentNumber + 1000

    txtstrTemp.WriteLine (Build21Header(nAgentNumber, mstrNow))
    txtstrTemp.WriteLine (Build84Header(nAgentNumber, mstrNow))



    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryCreateTransmitInventoryNABCA"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        Set rstemp = .Execute
    End With

    With rstemp
        If rstemp.RecordCount = 0 Then
            MsgBox "No Records found for Sales department 1"
            GoTo ExitSub
        End If

        Do While Not .EOF
            nLineCount = nLineCount + 1
            'If not gcfHandleErrors then Debug.Print !ItemID
            txtstrTemp.WriteLine BuildInvLineNabca(nAgentNumber, !Code11, !Onhand, mstrNow)
            .MoveNext
        Loop
    End With
    rstemp.Close
    txtstrTemp.WriteLine (Build84Trailer(nAgentNumber, nLineCount, mstrNow))
    txtstrTemp.WriteLine (Build21Trailer(nAgentNumber, nLineCount, mstrNow))
    txtstrTemp.Close
    'Shell ("Notepad C:\NPOS\Invent." & CStr(Format(nAgentNumber, "000")))

    If Archive Then
    'added 31415
        If gbKMDSFTPActive Then fsoTemp.CopyFile stDocName, "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\MonthlyInventAsOF" & Format(Now, "yy-mm-dd-hh-nn") & ".txt", True
    'end added
    Else
        'stDocName = "C:\NPOS\Invent." & CStr(Format(nAgentNumber, "000"))
        'Set txtstrtemp = fsotemp.OpenTextFile("C:\NPOS\Invent." & CStr(Format(nAgentNumber, "000")), ForWriting, True, TristateFalse)
        fsoTemp.CopyFile stDocName, "C:\NPOS\Inventory\NInvent.txt", True
        If gbKMDSFTPActive Then fsoTemp.CopyFile stDocName, "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\", True

    End If

ExitSub:

    Set comQDF = Nothing
    Set rstemp = Nothing
    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing

    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "NabcaOLCCInventFile"
    Resume ExitSub
End Sub

Public Function BuildInvLineNabca(AgentNumber As Long, ItemID As String, Quantity As Integer, mstrNow As String) As String
'05040305144411030013120001200000
'05040305144411030013220001200000
   BuildInvLineNabca = "84" & mstrNow & Format(CStr(AgentNumber), "0000")
'   BuildInvLine = BuildInvLine & Format(CStr(ItemID), "000000")
   BuildInvLineNabca = BuildInvLineNabca & ItemID ' formatted in calling routine
   BuildInvLineNabca = BuildInvLineNabca & Format(CStr(Quantity), "00000")
End Function
'Public Sub NabcaOLCCInventFile(nAgentNumber As Long, Archive As Boolean)
'Dim nOrderID As Long
'Dim stDocName As String
'Dim stLinkCriteria As String
''Dim rstemp As New ADODB.Recordset
'Dim rstemp As Object
'Set rstemp = CreateObject("ADODB.Recordset")
'
''Dim comQDF As New Command
'Dim comQDF As Object
'Set comQDF = CreateObject("ADODB.COMMAND")
''Dim nAgentNumber As Long
'Dim txtstrTemp As TextStream
'Dim fsoTemp As New FileSystemObject
'Dim strTemp As String
'Dim nLineCount As Integer
'Dim mstrNow As String
'
'    If gcfHandleErrors Then On Error GoTo ErrLog
''    nAgentNumber = AgentNumber
'    mstrNow = Format(Now, "yymmddhhnn")
'
'    If nAgentNumber > 1000 Then nAgentNumber = nAgentNumber - 1000
'
'    If Not fsoTemp.FolderExists("C:\NPOS") Then
'        fsoTemp.CreateFolder ("C:\NPOS")
'    End If
'    If Not fsoTemp.FolderExists("C:\NPOS\Archive") Then
'        fsoTemp.CreateFolder ("C:\NPOS\Archive")
'    End If
'
'    'add begin 31415
'    If gbKMDSFTPActive Then ' if KMDSFTPActive = true - G-5 -> Settings -> KMDSFTPActive
'        If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP") Then
'            fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP")
'        End If
'        If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT") Then
'            fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT")
'        End If
'        If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\Archive") Then
'            fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\Archive")
'        End If
'    End If
'    'add end 31415
'
'    If Archive Then
'        stDocName = "C:\NPOS\Archive\NInvent " & Format(Date, "mm-dd") & "." & CStr(Format(nAgentNumber, "000"))
'        'Set txtstrtemp = fsotemp.OpenTextFile("C:\NPOS\Archive\Invent " & Format(Date, "mm-dd") & "." & CStr(Format(nAgentNumber, "000")), ForWriting, True, TristateFalse)
'    Else
'        stDocName = "C:\NPOS\NInvent." & CStr(Format(nAgentNumber, "000"))
'        'Set txtstrtemp = fsotemp.OpenTextFile("C:\NPOS\Invent." & CStr(Format(nAgentNumber, "000")), ForWriting, True, TristateFalse)
'    End If
'    Set txtstrTemp = fsoTemp.OpenTextFile(stDocName, ForWriting, True, TristateFalse)
'
'    nAgentNumber = nAgentNumber + 1000
'
'    txtstrTemp.WriteLine (Build21Header(nAgentNumber, mstrNow))
'    txtstrTemp.WriteLine (Build84Header(nAgentNumber, mstrNow))
'
'
'
'    With comQDF
'        .ActiveConnection = CurrentProject.Connection
'        .CommandText = "qryCreateTransmitInventoryNABCA"
'        .CommandType = adCmdStoredProc
'        'ADO Numeric Datatypes are very particular
'        Set rstemp = .Execute
'    End With
'
'    With rstemp
'        If rstemp.RecordCount = 0 Then
'            MsgBox "No Records found for Sales department 1"
'            GoTo ExitSub
'        End If
'
'        Do While Not .EOF
'            nLineCount = nLineCount + 1
'            'If not gcfHandleErrors then Debug.Print !ItemID
'            txtstrTemp.WriteLine BuildInvLineNabca(nAgentNumber, !Code11, !Onhand, mstrNow)
'            .MoveNext
'        Loop
'    End With
'    rstemp.Close
'    txtstrTemp.WriteLine (Build84Trailer(nAgentNumber, nLineCount, mstrNow))
'    txtstrTemp.WriteLine (Build21Trailer(nAgentNumber, nLineCount, mstrNow))
'    txtstrTemp.Close
'    'Shell ("Notepad C:\NPOS\Invent." & CStr(Format(nAgentNumber, "000")))
'
'    If Archive Then
'    'added 31415
'        If gbKMDSFTPActive Then fsoTemp.CopyFile stDocName, "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\MonthlyInventAsOF" & Format(Now, "yy-mm-dd-hh-nn") & ".txt", True
'    'end added
'    Else
'        'stDocName = "C:\NPOS\Invent." & CStr(Format(nAgentNumber, "000"))
'        'Set txtstrtemp = fsotemp.OpenTextFile("C:\NPOS\Invent." & CStr(Format(nAgentNumber, "000")), ForWriting, True, TristateFalse)
'        fsoTemp.CopyFile stDocName, "C:\NPOS\Inventory\NInvent.txt", True
'        If gbKMDSFTPActive Then fsoTemp.CopyFile stDocName, "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\", True
'
'    End If
'
'ExitSub:
'
'    Set comQDF = Nothing
'    Set rstemp = Nothing
'    Set txtstrTemp = Nothing
'    Set fsoTemp = Nothing
'
'    Exit Sub
'
'ErrLog:
'    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "NabcaOLCCInventFile"
'    Resume ExitSub
'End Sub
'
