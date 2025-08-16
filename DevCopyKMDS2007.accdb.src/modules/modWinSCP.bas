Attribute VB_Name = "modWinSCP"
Option Compare Database
Option Explicit

Public Function WinSCPFTPTransmitKMDS() As Boolean

'    WaitForExecCmd "\\" & gstrCashsysInsLoc & "\kmds\kmdsftp.exe \\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\KMDSFTP\FTPSalesKMDS.txt" ', vbNormalFocus
'    WaitForExecCmd "\\" & gstrCashsysInsLoc & "\kmds\kmdsftp.exe \\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\KMDSFTP\FTPDailyInvKMDS.txt"    ', vbNormalFocus



Dim fsoTemp As New FileSystemObject
Dim f As Folder
Dim txtstrTemp As TextStream

Dim strTargetFile As String
Dim stAppName As String


If gcfHandleErrors Then On Error GoTo Err_Funtion
    WinSCPFTPTransmitKMDS = False


'    DEL C:\kmds\kmdsftp\WinSCPTestScript.log
'"C:\Program Files (x86)\WinSCP\WinSCP.exe" /log="C:\kmds\kmdsftp\WinSCPTestScript.log" /ini=nul /script="C:\kmds\kmdsftp\WinsSCPTestScript.txt
'C:\kmds\kmdsftp\WinSCPTestScript.log

    strTargetFile = "k:\mail\kmdsftp\out\WinSCPFTPTransmitKMDS.bat"

    Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
    With txtstrTemp
        .WriteLine ("k:")
        .WriteLine ("cd \mail\kmdsftp\out")
        .WriteLine ("move /y k:\mail\kmdsftp\out\*.log k:\mail\kmdsftp\out\Log")
        .Write ("""C:\Program Files (x86)\WinSCP\WinSCP.exe"" ")
        .Write ("/Log=""" & "k:\mail\kmdsftp\out\WinSCPFTPTransmitKMDS" & Format(Now, "yyyy-mm-dd-hh-nn") & ".log"" ")
        .Write ("/ini=nul ")
        .WriteLine ("/script=""" & "k:\mail\kmdsftp\out\WinSCPFTPTransmitKMDS.txt"" ")
        .WriteLine ("move /y k:\mail\kmdsftp\out\*DailySales* k:\mail\kmdsftp\out\archive")
        .WriteLine ("move /y k:\mail\kmdsftp\out\*DailyInventory* k:\mail\kmdsftp\out\archive")
        .WriteLine ("move /y k:\mail\kmdsftp\out\Invent* k:\mail\kmdsftp\out\archive")
        .WriteLine ("move /y k:\mail\kmdsftp\out\Monthly* k:\mail\kmdsftp\out\archive")
        .WriteLine ("move /y k:\mail\kmdsftp\out\Order* k:\mail\kmdsftp\out\archive")
        .WriteLine ("move /y k:\mail\kmdsftp\out\ADS* k:\mail\kmdsftp\out\archive")
        .Close
    End With
    Set txtstrTemp = Nothing

    WaitForExecCmd strTargetFile

    WinSCPFTPTransmitKMDS = True

Exit_Function:
    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing
    Exit Function

Err_Funtion:
    MsgBox Err.Description
    Resume Exit_Function


End Function


Public Sub CreateWinSCPtoKMDSFTPScriptFile()
    'Begin 'added for 31415
    'ftp -> kmds uses  ...https://winscp.net/eng/index.php
    '. following code creates the script file used by WinSCP.exe


    Dim strFTPServer As String
    Dim strFTPUser As String
    Dim strFTPPassword As String
    Dim strTargetFile As String
    Dim strFTPLogon As String
    Dim nDoubleQuote As Integer

    Dim fsoTemp As New FileSystemObject
    Dim txtstrTemp As TextStream

    On Error GoTo ErrLog
    If Not gbKMDSFTPActive Then 'jump out if not active - G-5 -> Settings -> KMDSFTPActive
        Exit Sub
    End If
    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL")
    End If
    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP")
    End If
    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT")
    End If
    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\Archive") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\Archive")
    End If
    If fsoTemp.FileExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\Log") Then
        fsoTemp.DeleteFile ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\Log"), True
    End If
    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\Log") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\Log")
    End If

    strTargetFile = "k:\MAIL\KMDSFTP\OUT\WinSCPFTPTransmitKMDS.TXT"


    Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)

'     here is a working example of the text file
'    open ftp://Agents:Ag3nt%24@ftp.kmds.us/ -rawsettings ProxyPort=0
'    lcd K:\Mail\KMDSFTP\OUT
'    cd /agent1057/DailyInventory
'    put *DailyInventory*
'    cd /agent1057/DailySales
'    put *DailySales*
'    cd /agent1057/MonthlyInventory
'    put *Invent*
'    exit
    With txtstrTemp
        'the MONTH END inventory file format is different from the DAILY inventory..
        'so the file name is different too.. MONTH END  begins with "Invent".. DAILY begins with the agent number...

        .WriteLine gstrftpLogonWinSCPKMDS 'set in LoadDefaults routine
        .WriteLine "lcd K:\Mail\KMDSFTP\OUT"
        '.WriteLine "cd /agent" & glnAgentNumber
        .WriteLine "cd /agent" & glnAgentNumber & "/DailyInventory"
        .WriteLine "put " & Chr(34) & "DailyInventory*" & Chr(34)
        .WriteLine "cd /agent" & glnAgentNumber & "/DailySales"
        .WriteLine "put " & Chr(34) & "DailySales*" & Chr(34)
        .WriteLine "cd /agent" & glnAgentNumber & "/MonthlyInventory"
        .WriteLine "put " & Chr(34) & "MonthlyInvent*" & Chr(34)
        .WriteLine "put " & Chr(34) & "Invent*" & Chr(34)
        .WriteLine "cd /agent" & glnAgentNumber & "/PurchaseOrder"
        .WriteLine "put " & Chr(34) & "Order*" & Chr(34)
        .WriteLine "cd /AgentDailySales"
        .WriteLine "put " & Chr(34) & "ADS*" & Chr(34)
        .WriteLine "exit"
        .Close
    End With
    Set txtstrTemp = Nothing

'end 31415

'begin 31416 added to implement KillSwitch -

    strTargetFile = "c:\npos\WinSCPKS.TXT"

    Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
    With txtstrTemp
        .WriteLine gstrftpLogonWinSCPKMDS
        .WriteLine "lcd K:\Mail\In"
        .WriteLine "get KMDSSW.txt"
        .WriteLine "exit"
        .Close
    End With
    Set txtstrTemp = Nothing

    strTargetFile = "c:\npos\WinSCPKS.bat"
    Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
    With txtstrTemp
        .WriteLine "Echo off"
        .WriteLine "cls"
        .WriteLine "c:"
        .WriteLine "cd \npos"
        .Write ("""C:\Program Files (x86)\WinSCP\WinSCP.exe"" ")
        .Write ("/Log=""" & "k:\mail\out\WinSCPKS" & Format(Now, "yyyy-mm-dd-hh-nn") & ".log"" ")
        .Write ("/ini=nul ")
        .WriteLine ("/script=""" & "c:\npos\WinSCPKS.txt"" ")
        .Close
    End With
    Set txtstrTemp = Nothing
'end 31416

ExitSub:

    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Create WinSCP to KMDSFTP Script File"
    Resume ExitSub
End Sub

Public Sub CreateWinSCPtoOLASFTPScriptFile(DayendOrderMonthendNewpriceItemFileLicenseeOrEPS As String, Optional FileName As String)
    'Begin 'added for 31415
    'ftp -> kmds uses  ...https://winscp.net/eng/index.php
    '. following code creates the script file used by WinSCP.exe

    Dim fsoTemp As New FileSystemObject
    Dim txtstrTemp As TextStream

    Dim strFTPServer As String
    Dim strFTPUser As String
    Dim strFTPPassword As String
    Dim strTargetFile As String
    Dim strOpenCommand As String
    Dim strOpenParameters As String
    On Error GoTo ErrLog

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


'     here is a working example of the text file
'    open ftp://Agent1001:Agent1001@ftp.olcc.state.or.us/ -rawsettings ProxyPort=0
'
'    # Your command 1
'    # Your command 2
'
'    exit
'
'    # Execute the script using a command like:
'    # "C:\Program Files (x86)\WinSCP\WinSCP.exe" /log="C:\writable\path\to\log\WinSCP.log" /ini=nul /script="C:\path\to\script\script.txt"

    DayendOrderMonthendNewpriceItemFileLicenseeOrEPS = UCase(DayendOrderMonthendNewpriceItemFileLicenseeOrEPS)

    Select Case DayendOrderMonthendNewpriceItemFileLicenseeOrEPS
    Case "DAYEND"
        strTargetFile = "k:\MAIL\OUT\WinSCPDaily.TXT"
        Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
        With txtstrTemp
            .WriteLine gstrFTPLogonWinSCPOLCC ' Set in LoadDefaults
            .WriteLine "lcd K:\Mail\OUT"
            .WriteLine "put DailyInv*"
            .WriteLine "put Sales*"
            .WriteLine "exit"
            .Close
        End With
        Set txtstrTemp = Nothing
        strTargetFile = "k:\MAIL\OUT\WinSCPDaily.BAT"
        Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
        With txtstrTemp
            .WriteLine "k:"
            .WriteLine "cd \mail\out"
            .WriteLine "move /y k:\mail\out\*.log k:\mail\out\Log"
            .Write ("""C:\Program Files (x86)\WinSCP\WinSCP.exe"" ")
            .Write ("/Log=""" & "k:\MAIL\OUT\WinSCPDaily" & Format(Now, "yyyy-mm-dd-hh-nn") & ".log"" ")
            .Write ("/ini=nul ")
            .WriteLine ("/script=""" & "k:\MAIL\OUT\WinSCPDaily.txt"" ")
            .WriteLine "move /y k:\mail\out\Sales* k:\mail\out\archive"
            .WriteLine "move /y k:\mail\out\DailyInv* k:\mail\out\archive"
            .Close
        End With
        Set txtstrTemp = Nothing


    Case "ORDER"
        'SEND PURCHASE ORDER TO OLCC
        strTargetFile = "c:\npos\WinSCPOrderOLCC.TXT"
        Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
        With txtstrTemp
            .WriteLine gstrFTPLogonWinSCPOLCC
            .WriteLine "lcd c:\npos\order"
            .WriteLine "send order.txt"
            '.WriteLine "put Sales*"
            .WriteLine "exit"
            .Close
        End With
        Set txtstrTemp = Nothing

        strTargetFile = "c:\npos\WinSCPOrderOLCC.bat"
        Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
        With txtstrTemp
            .WriteLine "c:"
            .WriteLine "cd \npos"
            .WriteLine "move /y k:\mail\out\*.log k:\mail\out\Log"
            .Write ("""C:\Program Files (x86)\WinSCP\WinSCP.exe"" ")
            .Write ("/Log=""" & "k:\mail\out\WinSCPOrderOLCC" & Format(Now, "yyyy-mm-dd-hh-nn") & ".log"" ")
            .Write ("/ini=nul ")
            .WriteLine ("/script=""" & "c:\npos\WinSCPOrderOLCC.txt"" ")
            .Close
        End With
        Set txtstrTemp = Nothing


    Case "MONTHEND"
        'SEND MONTH END INVENTORY TO OLCC
        strTargetFile = "c:\npos\WinSCPMonthEndOLCC.TXT"
        Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
        With txtstrTemp
            .WriteLine gstrFTPLogonWinSCPOLCC
            .WriteLine "lcd c:\npos\inventory"
            .WriteLine "send invent.txt"
            .WriteLine "exit"
            .Close
        End With
        Set txtstrTemp = Nothing

        strTargetFile = "c:\npos\WinSCPMonthEndOLCC.bat"
        Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
        With txtstrTemp
            .WriteLine "c:"
            .WriteLine "cd \npos"
            .WriteLine "move /y k:\mail\out\*.log k:\mail\out\Log"
            .Write ("""C:\Program Files (x86)\WinSCP\WinSCP.exe"" ")
            .Write ("/Log=""" & "k:\mail\out\WinSCPMonthEndOLCC" & Format(Now, "yyyy-mm-dd-hh-nn") & ".log"" ")
            .Write ("/ini=nul ")
            .WriteLine ("/script=""" & "c:\npos\WinSCPMonthEndOLCC.txt"" ")
            .Close
        End With
        Set txtstrTemp = Nothing
    Case "NEWPRICE"
        strTargetFile = "c:\npos\WinSCPNewprice.TXT"
        Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
        With txtstrTemp
            .WriteLine gstrFTPLogonWinSCPOLCC
            .WriteLine "lcd c:\npos"
            .WriteLine "get NEWPRICE_ICExpansion.txt"
            .WriteLine "get full_pricelist_ICExpansion.txt"
            .WriteLine "exit"
            .Close
        End With
        Set txtstrTemp = Nothing

        strTargetFile = "c:\npos\WinSCPNewprice.bat"
        Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
        With txtstrTemp
            .WriteLine "c:"
            .WriteLine "cd \npos"
            .WriteLine "move /y k:\mail\out\*.log k:\mail\out\Log"
            .Write ("""C:\Program Files (x86)\WinSCP\WinSCP.exe"" ")
            .Write ("/Log=""" & "k:\mail\out\WinSCPNewprice" & Format(Now, "yyyy-mm-dd-hh-nn") & ".log"" ")
            .Write ("/ini=nul ")
            .WriteLine ("/script=""" & "c:\npos\WinSCPNewprice.txt"" ")
            .Close
        End With
        Set txtstrTemp = Nothing

    Case "ITEMFILE"
        strTargetFile = "c:\npos\WinSCPItemFile.TXT"
        Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
        With txtstrTemp
            .WriteLine gstrFTPLogonWinSCPOLCC
            .WriteLine "lcd c:\npos"
            .WriteLine "get item_ICExpansion.txt"
            '.WriteLine "get scancode_ICExpansion.txt"
            .WriteLine "exit"
            .Close
        End With
        Set txtstrTemp = Nothing

        strTargetFile = "c:\npos\WinSCPItemFile.bat"
        Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
        With txtstrTemp
            .WriteLine "c:"
            .WriteLine "cd \npos"
            .WriteLine "move /y k:\mail\out\*.log k:\mail\out\Log"
            .Write ("""C:\Program Files (x86)\WinSCP\WinSCP.exe"" ")
            .Write ("/Log=""" & "k:\mail\out\WinSCPItemFile" & Format(Now, "yyyy-mm-dd-hh-nn") & ".log"" ")
            .Write ("/ini=nul ")
            .WriteLine ("/script=""" & "c:\npos\WinSCPItemFile.txt"" ")
            .Close
        End With
        Set txtstrTemp = Nothing
    Case "EPS"
        strTargetFile = "c:\npos\WinSCPEPS.TXT"
        Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
        With txtstrTemp
            .WriteLine gstrFTPLogonWinSCPOLCC
            .WriteLine "lcd c:\npos\mail\in"
            .WriteLine "get Packinglist*"
            .WriteLine "exit"
            .Close
        End With
        Set txtstrTemp = Nothing

        strTargetFile = "c:\npos\WinSCPEPS.bat"
        Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
        With txtstrTemp
            .WriteLine "c:"
            .WriteLine "cd \npos\mail\in"
            .WriteLine "move /y k:\mail\out\*.log k:\mail\out\Log"
            .Write ("""C:\Program Files (x86)\WinSCP\WinSCP.exe"" ")
            .Write ("/Log=""" & "k:\mail\out\WinSCPEPS" & Format(Now, "yyyy-mm-dd-hh-nn") & ".log"" ")
            .Write ("/ini=nul ")
            .WriteLine ("/script=""" & "c:\npos\WinSCPEPS.txt"" ")
            .Close
        End With
        Set txtstrTemp = Nothing
    Case "LICENSEE"
        strTargetFile = "c:\npos\WinSCPLicensee.TXT"

        Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
        With txtstrTemp
            .WriteLine gstrFTPLogonWinSCPOLCC
            .WriteLine "lcd \\" & gstrCashsysInsLoc & "\npos"
            .WriteLine "get licensee.txt*"
            .WriteLine "exit"
            .Close
        End With
        Set txtstrTemp = Nothing

        strTargetFile = "c:\npos\WinSCPLicensee.bat"
        Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
        With txtstrTemp
            .WriteLine "c:"
            .WriteLine "cd \npos\mail\in"
            .WriteLine "move /y k:\mail\out\*.log k:\mail\out\Log"
            .Write ("""C:\Program Files (x86)\WinSCP\WinSCP.exe"" ")
            .Write ("/Log=""" & "k:\mail\out\WinSCPLicensee" & Format(Now, "yyyy-mm-dd-hh-nn") & ".log"" ")
            .Write ("/ini=nul ")
            .WriteLine ("/script=""" & "c:\npos\WinSCPLicensee.txt"" ")
            .Close
        End With
        Set txtstrTemp = Nothing


    Case Else
        MsgBox "Unhandled Case: " & DayendOrderMonthendNewpriceItemFileLicenseeOrEPS, , "CreateWinSCPtoOLASFTPScriptFile"
    End Select


ExitSub:

    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Create WinSCP to KMDSFTP Script File"
    Resume ExitSub
End Sub


'
Public Sub WinSCPtoKMDSFTPScriptFileXFR(TransferInOrOut As String, Optional FileName As String, Optional OtherAgentnumber As Integer)
    'Begin 'added for 31415
    'ftp -> kmds uses  ...https://winscp.net/eng/index.php
    '. following code creates the script file used by WinSCP.exe

    Dim fsoTemp As New FileSystemObject
    Dim txtstrTemp As TextStream

    Dim strFTPServer As String
    Dim strFTPUser As String
    Dim strFTPPassword As String
    Dim strTargetFile As String
    Dim strFTPLogon As String
    Dim strWinscpLoglevel As String
    Dim nDoubleQuote As Integer


    On Error GoTo ErrLog
    If Not gbKMDSFTPActive Then 'jump out if not active - G-5 -> Settings -> KMDSFTPActive
        Exit Sub
    End If
    If Not fsoTemp.FolderExists("c:\KMDS\FTP") Then
        fsoTemp.CreateFolder ("c:\KMDS\FTP")
    End If
    If Not fsoTemp.FolderExists("c:\KMDS\FTP\IN") Then
        fsoTemp.CreateFolder ("c:\KMDS\FTP\IN")
    End If
    If Not fsoTemp.FolderExists("c:\KMDS\FTP\Archive") Then
        fsoTemp.CreateFolder ("c:\KMDS\FTP\Archive")
    End If
    If Not fsoTemp.FolderExists("c:\KMDS\FTP\Log") Then
        fsoTemp.CreateFolder ("c:\KMDS\FTP\Log")
    End If
    strWinscpLoglevel = GetKMDSSettings("WinSCPLogLevel", "-1") 'false

    TransferInOrOut = UCase(TransferInOrOut)
    Select Case TransferInOrOut
    Case "IN"
        strTargetFile = "k:\MAIL\KMDSFTP\In\WinSCPtoKMDSFTPScriptFileXFRIn.TXT"
        Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)

    '     here is a working example of the text file
    '    open ftp://Agents:Ag3nt%24@ftp.kmds.us/ -rawsettings ProxyPort=0
    '    lcd K:\Mail\KMDSFTP\OUT
    '    cd /agent1057/DailyInventory
    '    put *DailyInventory*
    '    cd /agent1057/DailySales
    '    put *DailySales*
    '    cd /agent1057/MonthlyInventory
    '    put *Invent*
    '    exit
        With txtstrTemp
            'the MONTH END inventory file format is different from the DAILY inventory..
            'so the file name is different too.. MONTH END  begins with "Invent".. DAILY begins with the agent number...

            .WriteLine gstrftpLogonWinSCPKMDS 'set in LoadDefaults routine
            .WriteLine "lcd K:\Mail\KMDSFTP\IN"
            .WriteLine "cd /agent" & glnAgentNumber & "/In"
            .WriteLine "get " & Chr(34) & "*.csv" & Chr(34)
            ' rm = delete
            .WriteLine "rm " & Chr(34) & "*.csv" & Chr(34)

            .WriteLine "exit"
            .Close
        End With
        Set txtstrTemp = Nothing
        strTargetFile = "k:\mail\kmdsftp\in\WinSCPtoKMDSFTPScriptFileXFRIn.bat"

        Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
        With txtstrTemp
            .WriteLine ("K:")
            .WriteLine ("cd \Mail\KMDSFTP\IN")
            .WriteLine ("move /y *.log k:\mail\kmdsftp\in\Log")
            .Write ("""C:\Program Files (x86)\WinSCP\WinSCP.exe"" ")
            .Write ("/Log=""" & "k:\mail\kmdsftp\in\WinSCPtoKMDSFTPScriptFileXFRIn" & Format(Now, "yyyy-mm-dd-hh-nn") & ".log"" ")
            .Write ("/Loglevel=" & strWinscpLoglevel & " ")
            .Write ("/ini=nul ")
            .WriteLine ("/script=""" & "k:\mail\kmdsftp\in\WinSCPtoKMDSFTPScriptFileXFRIn.txt"" ")
            .WriteLine ("xcopy  k:\mail\kmdsftp\in\*.csv c:\kmds\ftp\in\*.csv /y")
            .WriteLine ("move /y k:\mail\kmdsftp\in\*.csv k:\mail\kmdsftp\in\archive")
            .Close
        End With
        Set txtstrTemp = Nothing

        WaitForExecCmd strTargetFile

    Case "OUT"
        strTargetFile = "k:\MAIL\KMDSFTP\OUT\WinSCPtoKMDSFTPScriptFileXFROut.TXT"
        Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)

    '     here is a working example of the text file
    '    open ftp://Agents:Ag3nt%24@ftp.kmds.us/ -rawsettings ProxyPort=0
    '    lcd K:\Mail\KMDSFTP\OUT
    '    cd /agent1057/DailyInventory
    '    put *DailyInventory*
    '    cd /agent1057/DailySales
    '    put *DailySales*
    '    cd /agent1057/MonthlyInventory
    '    put *Invent*
    '    exit
        With txtstrTemp
            'the MONTH END inventory file format is different from the DAILY inventory..
            'so the file name is different too.. MONTH END  begins with "Invent".. DAILY begins with the agent number...

            .WriteLine gstrftpLogonWinSCPKMDS 'set in LoadDefaults routine
            .WriteLine "lcd K:\Mail\KMDSFTP\OUT"
            .WriteLine "cd /agent" & OtherAgentnumber & "/In"
            .WriteLine "send " & FileName
            .WriteLine "exit"
            .Close
        End With
        Set txtstrTemp = Nothing
        strTargetFile = "k:\mail\kmdsftp\Out\WinSCPtoKMDSFTPScriptFileXFROut.bat"

        Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
        With txtstrTemp
            .WriteLine ("K:")
            .WriteLine ("cd \Mail\KMDSFTP\OUT")
            .WriteLine ("move /y *.log k:\mail\kmdsftp\in\Log")
            .Write ("""C:\Program Files (x86)\WinSCP\WinSCP.exe"" ")
            .Write ("/Log=""" & "k:\mail\kmdsftp\OUT\WinSCPtoKMDSFTPScriptFileXFROut" & Format(Now, "yyyy-mm-dd-hh-nn") & ".log"" ")
            .Write ("/ini=nul ")
            .WriteLine ("/script=""" & "k:\mail\kmdsftp\OUT\WinSCPtoKMDSFTPScriptFileXFROut.txt"" ")
            .WriteLine ("move /y " & FileName & " k:\mail\kmdsftp\out\archive")
            .Close
        End With
        Set txtstrTemp = Nothing

        WaitForExecCmd strTargetFile
    Case Else
        MsgBox "Unhandled CASE: " & TransferInOrOut & "Routine: WinSCPtoKMDSFTPScriptFileXFR"
    End Select
'end 31415
ExitSub:

    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Create WinSCP to KMDSFTP Script File"
    Resume ExitSub
End Sub



Public Function VerifyCurrentKMDSCustomer() As Boolean
    'Begin 'added for 31416
    Dim rstemp As Object
    Dim fsoTemp As New FileSystemObject

    VerifyCurrentKMDSCustomer = True

    'creates WinSCPKS .bat and .txt
    CreateWinSCPtoKMDSFTPScriptFile
    'get KMDSSW.TXT from FTP.KMDS.TXT
    If fsoTemp.FileExists("K:\Mail\In\KMDSSW.TXT") Then
        fsoTemp.DeleteFile "K:\Mail\In\KMDSSW.TXT"
    End If
    WaitForExecCmd "c:\npos\WinSCPKS.BAT"
    If fsoTemp.FileExists("K:\Mail\In\KMDSSW.TXT") Then
        'refresh KMKDSSW table with KMDSSW.TXT downloaded in previous command
        SetOptions False
        DoCmd.RunSQL "Delete * from KMDSKS"
        DoCmd.TransferText acImportFixed, "KMDSSWImportSpecification", "KMDSKS", "K:\Mail\In\KMDSSW.TXT", False
        SetOptions True

        'create a recordset joining Sysfile/KMDSSW tables
        Set rstemp = CreateObject("ADODB.Recordset")
        rstemp.Open "qrySysfileKMDSKS", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
        If rstemp.BOF And rstemp.EOF Then 'recordset is empty... sysfile.AgentNumber NOT found in KMDSSW table
            EOLog "... VerifyCurrentKMDSCustomer: True"
            VerifyCurrentKMDSCustomer = True
        Else 'recordset has been found
            VerifyCurrentKMDSCustomer = False
            EOLog "... VerifyCurrentKMDSCustomer: False Please update your support subscription to resume transmits"
        End If
        rstemp.Close
    Else
        EOLog "... VerifyCurrentKMDSCustomer: ks file not found"
    End If
ExitFunction:
    If fsoTemp.FileExists("K:\Mail\In\KMDSSW.TXT") Then
        fsoTemp.DeleteFile "K:\Mail\In\KMDSSW.TXT"
    End If
    Set rstemp = Nothing
    Set fsoTemp = Nothing

    Exit Function

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "VerifyCurrentKMDSCustomer"
    Resume ExitFunction
End Function
