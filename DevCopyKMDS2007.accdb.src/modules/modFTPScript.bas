Attribute VB_Name = "modFTPScript"
Option Compare Database
Option Explicit

Public Sub MakeFTPScriptFiles()

Dim nOrderID As Long
Dim stDocName As String
Dim stLinkCriteria As String

Dim nAgentNumber As Long
Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject
Dim strTemp As String
Dim nLineCount As Integer
Dim strSourceFile As String

Dim strTargetFile As String

Dim f As Folder


    If gcfHandleErrors Then On Error GoTo ErrLog

    If glnAgentNumber < 1000 Then glnAgentNumber = glnAgentNumber + 1000

    If Not fsoTemp.FolderExists("C:\NPOS") Then
        fsoTemp.CreateFolder ("C:\NPOS")
    End If



    strTargetFile = "C:\NPOS\FTPOrd.scr"

'open ftp.olcc.state.or.us
'Agent1116
'Agent1116
'Dir
'send order.116 order.txt
'Dir
'get newprice.txt newprice.fil
'Quit

    Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
    txtstrTemp.WriteLine ("open ftp.olcc.state.or.us")
    txtstrTemp.WriteLine ("Agent" & CStr(glnAgentNumber))
    txtstrTemp.WriteLine ("Agent" & CStr(glnAgentNumber))
    'txtstrtemp.WriteLine ("binary")
    'txtstrtemp.WriteLine ("Dir")
    txtstrTemp.WriteLine ("send order." & Format(CStr((glnAgentNumber - 1000)), "000") & " order.txt")
    'txtstrtemp.WriteLine ("Dir")
    txtstrTemp.WriteLine ("get NEWPRICE_ICExpansion.txt")
    txtstrTemp.WriteLine ("get full_pricelist_ICExpansion.txt")

    txtstrTemp.WriteLine ("Quit")
    txtstrTemp.Close
    Set txtstrTemp = Nothing

'open ftp.olcc.state.or.us
'Agent1116
'Agent1116
'Dir
'send invent.116 invent.txt
'Dir
'get newprice.txt newprice.fil
'Quit

    strTargetFile = "C:\NPOS\FTPInv.scr"
    Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
    txtstrTemp.WriteLine ("open ftp.olcc.state.or.us")
    txtstrTemp.WriteLine ("Agent" & CStr(glnAgentNumber))
    txtstrTemp.WriteLine ("Agent" & CStr(glnAgentNumber))
    'txtstrtemp.WriteLine ("binary")
    'txtstrtemp.WriteLine ("Dir")
    'txtstrTemp.WriteLine ("send invent." & CStr((glnAgentNumber - 1000)) & " invent.txt")
    txtstrTemp.WriteLine ("send invent." & Format(CStr((glnAgentNumber - 1000)), "000") & " invent.txt")

    'txtstrtemp.WriteLine ("Dir")
    txtstrTemp.WriteLine ("get full_pricelist_ICExpansion.txt")
    txtstrTemp.WriteLine ("get NEWPRICE_ICExpansion.txt")
    txtstrTemp.WriteLine ("Quit")
    txtstrTemp.Close
    Set txtstrTemp = Nothing


    '2015-01-11
    strTargetFile = "C:\NPOS\FTPEPS.scr" ' Electronic Packing List
    Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
    txtstrTemp.WriteLine ("open ftp.olcc.state.or.us")
    txtstrTemp.WriteLine ("Agent" & CStr(glnAgentNumber))
    txtstrTemp.WriteLine ("Agent" & CStr(glnAgentNumber))
    'txtstrtemp.WriteLine ("binary")
    'txtstrtemp.WriteLine ("Dir")
    txtstrTemp.WriteLine ("prompt")
    txtstrTemp.WriteLine ("mget PackingList." & CStr(glnAgentNumber) & ".*.txt")
    txtstrTemp.WriteLine ("Quit")


ExitSub:
    txtstrTemp.Close
    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmdCreate"
    Resume ExitSub
End Sub


Public Sub ImportFulPriceTable()
  If gcfHandleErrors Then On Error GoTo ErrLog
    If Dir("C:\npos\full_pricelist_ICExpansion.txt") = "" Then
        MsgBox "C:\npos\full_pricelist_ICExpansion.txt not found"
        Exit Sub
    End If
    SetOptions False
    CurrentDb().Execute ("delete * from FulPrice")
    DoCmd.TransferText acImportFixed, "FulPrice Import Specification", "FulPrice", "C:\npos\full_pricelist_ICExpansion.txt", False

    SetOptions False

ExitSub:
    Exit Sub

ErrLog:
    EOLog "Error #" & Err.Number & " " & Err.Description & " : EndofMonth"
    fMsgBox "EndofMonth"
    Resume ExitSub

End Sub

Public Function MakeXFRScript(GetOrSend As String, XFRAgent As String, Optional stDocName As String)

Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject
Dim strTemp As String

Dim strTargetFile As String

Dim f As Folder


    If gcfHandleErrors Then On Error GoTo ErrLog

    If glnAgentNumber < 1000 Then glnAgentNumber = glnAgentNumber + 1000

    If Not fsoTemp.FolderExists("C:\KMDS\FTP") Then
        fsoTemp.CreateFolder ("C:\KMDS\FTP")
    End If

    If Not fsoTemp.FolderExists("C:\KMDS\FTP\OUT") Then
        fsoTemp.CreateFolder ("C:\KMDS\FTP\OUT")
    End If
    If Not fsoTemp.FolderExists("C:\KMDS\FTP\IN") Then
        fsoTemp.CreateFolder ("C:\KMDS\FTP\IN")
    End If

    strTargetFile = "C:\KMDS\FTP\XFR.scr"

'open ftp.olcc.state.or.us
'Agent1116
'Agent1116
'Dir
'send order.116 order.txt
'Dir
'get newprice.txt newprice.fil
'Quit
    Dim strFTPServer As String
    Dim strFTPUser As String
    Dim strFTPPassword As String
    strFTPServer = GetKMDSSettings("KMDSFTPServer", "ftp.kmds.us")
    strFTPUser = GetKMDSSettings("KMDSFTPUser", "Agents")
    strFTPPassword = GetKMDSSettings("KMDSFTPPassword", "Ag3nt$")



    Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
'    txtstrTemp.WriteLine ("open ftp.kmds.us")
'    'txtstrTemp.WriteLine ("open ftp.ipower.com")
'    txtstrTemp.WriteLine ("Agents")
'    txtstrTemp.WriteLine ("Ag3nt$")

    txtstrTemp.WriteLine "open " & (strFTPServer)
    txtstrTemp.WriteLine (strFTPUser)
    txtstrTemp.WriteLine (strFTPPassword)

    txtstrTemp.WriteLine ("binary")
    txtstrTemp.WriteLine ("cd agent" & CStr(XFRAgent))

    txtstrTemp.WriteLine ("Dir")

    Select Case UCase(GetOrSend)
    Case "GET"
        txtstrTemp.WriteLine ("lcd IN")

        txtstrTemp.WriteLine ("prompt")
        txtstrTemp.WriteLine ("mget *.csv")
        txtstrTemp.WriteLine ("mdelete *.csv")

    Case "SEND"
        txtstrTemp.WriteLine ("lcd OUT")
        txtstrTemp.WriteLine ("send " & stDocName)
    Case Else
        MsgBox "UNKNOWN CASE : " & GetOrSend, , "MakeXFR"
    End Select

    txtstrTemp.WriteLine ("Dir")
    txtstrTemp.WriteLine ("quit")

    txtstrTemp.Close
    Set txtstrTemp = Nothing


    strTargetFile = "C:\KMDS\FTP\KMDSXFR.scr"

'open ftp.olcc.state.or.us
'Agent1116
'Agent1116
'Dir
'send order.116 order.txt
'Dir
'get newprice.txt newprice.fil
'Quit

    Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
'    'txtstrTemp.WriteLine ("ftp.kmds.us")
'    txtstrTemp.WriteLine ("ftp.ipower.us")
'    txtstrTemp.WriteLine ("Agents")
'    txtstrTemp.WriteLine ("Agents")

    txtstrTemp.WriteLine "open " & (strFTPServer)
    txtstrTemp.WriteLine (strFTPUser)
    txtstrTemp.WriteLine (strFTPPassword)


    txtstrTemp.WriteLine ("RCD agent" & CStr(XFRAgent))
    Select Case UCase(GetOrSend)
    Case "GET"
        txtstrTemp.WriteLine ("RCD IN")
        txtstrTemp.WriteLine ("mget *.csv")
        txtstrTemp.WriteLine ("mdelete *.csv")

    Case "SEND"
        txtstrTemp.WriteLine ("RCD OUT")
        txtstrTemp.WriteLine ("PUT " & stDocName)
    Case Else
        MsgBox "UNKNOWN CASE : " & GetOrSend, , "MakeXFR"
    End Select

    txtstrTemp.WriteLine ("Dir")
    txtstrTemp.WriteLine ("quit")

    txtstrTemp.Close
    Set txtstrTemp = Nothing



ExitSub:
    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing
    Exit Function

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "MakeXFROutScript"
    Resume ExitSub
End Function



Public Function FTPXfr() As Boolean


Dim fsoTemp As New FileSystemObject
Dim f As Folder
Dim txtstrTemp As TextStream

Dim strTargetFile As String
Dim stAppName As String


If gcfHandleErrors Then On Error GoTo Err_FTPXfr
    FTPXfr = False
    If Not fsoTemp.FolderExists("C:\KMDS\FTP") Then
        fsoTemp.CreateFolder ("C:\KMDS\FTP")
    End If

    strTargetFile = "C:\KMDS\FTP\FTPXfr.bat"

    Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
    txtstrTemp.WriteLine ("c:")
    txtstrTemp.WriteLine ("cd \kmds\ftp")
    txtstrTemp.WriteLine ("ftp -s:XFR.scr > XFR.log")
    txtstrTemp.Close
    Set txtstrTemp = Nothing


    If MsgBox("Are you logged on to the Internet?", vbYesNo) = vbYes Then


        stAppName = "C:\KMDS\FTP\FTPXfr.bat"
        WaitForExecCmd stAppName
        FTPXfr = True
    Else
        MsgBox "Please connect to the internet and then start the process over"
    End If

Exit_FTPXfr:
    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing
    Exit Function

Err_FTPXfr:
    MsgBox Err.Description
    Resume Exit_FTPXfr
End Function


Public Sub ProcessFTPXFRIn()
Dim varDir As Variant
Dim strInDir As String
'Dim cmdtemp3 As New Command
Dim cmdtemp3 As Object
Set cmdtemp3 = CreateObject("ADODB.COMMAND")

''Dim cmdtemp4 As New Command
'Dim cmdtemp4 As Object
'Set cmdtemp4 = CreateObject("ADODB.COMMAND")

'Dim cmdtemp1 As New Command
Dim cmdtemp1 As Object
Set cmdtemp1 = CreateObject("ADODB.COMMAND")

'Dim cmdtemp2 As New Command
Dim cmdtemp2 As Object
Set cmdtemp2 = CreateObject("ADODB.COMMAND")
'If GetKMDSSettings("UseWinSCPForOLASTransmits", "False") = True Then
'    strInDir = "K:\Mail\KMDSFTP\IN\"
'Else
    strInDir = "C:\KMDS\FTP\IN\"
'End If

Dim txtCurFile As String
Dim txtCurFileTXT As String

    If gcfHandleErrors Then On Error GoTo ErrLog
    varDir = Dir(strInDir & "*.csv")
    If varDir <> "" Then

        Do While varDir <> ""
            SetOptions False

            DoCmd.RunSQL "Delete * from AgentTransferFTPImport", False
            SetOptions True

            txtCurFile = strInDir & varDir


            If gcfHandleErrors Then On Error GoTo ErrLogx
            DoCmd.TransferText acImportDelim, "AgentTransferFTPImportSpec", "AgentTransferFTPImport", txtCurFile, True
            If gcfHandleErrors Then On Error GoTo ErrLog

            'map related items by UPC Codes
            With cmdtemp3
                .ActiveConnection = CurrentProject.Connection
                .CommandText = "qryAgentTransferFTP_Step1_Related"
                .CommandType = adCmdStoredProc
                .Execute
            End With

            'this will create a record in the AgentTransferHead table
            With cmdtemp1
                .ActiveConnection = CurrentProject.Connection
                .CommandText = "qryAgentTransferFTP_Step2_AppendHead"
                .CommandType = adCmdStoredProc
                .Execute
            End With

            'append lines in AgentTransferLines
            With cmdtemp2
                .ActiveConnection = CurrentProject.Connection
                .CommandText = "qryAgentTransferFTP_Step3_AppendLine"
                .CommandType = adCmdStoredProc
                .Execute

            End With
            'rename input file for 2 reasons,
            '1. don't reprocess
            '2. dirdater.exe looks for .txt files
            txtCurFileTXT = Replace(txtCurFile, ".CSV", ".TXT")
            Name txtCurFile As txtCurFileTXT

            varDir = Dir
        Loop

    End If


ExitSub:
    SetOptions True
    Set cmdtemp1 = Nothing
    Set cmdtemp2 = Nothing
    Exit Sub

ErrLogx:
    MsgBox Err.Number & " " & Err.Description & vbCrLf & txtCurFile, vbCritical, "ProcessFTPXFRIn"
    txtCurFileTXT = Replace(txtCurFile, ".CSV", ".err.TXT")
    Name txtCurFile As txtCurFileTXT
    'Kill txtCurFile
    Resume Next

ErrLog:
    MsgBox Err.Number & " " & Err.Description & vbCrLf & txtCurFile, vbCritical, "ProcessFTPXFRIn"
    txtCurFileTXT = Replace(txtCurFile, ".CSV", ".err.TXT")
    Name txtCurFile As txtCurFileTXT
    Resume ExitSub

End Sub


Public Function OLCCFTPGet(SourceFileName As String, TargetFile As String) As Boolean
Dim strTemp As String
Dim fsoTemp As New FileSystemObject
Dim txtstrTemp As TextStream
Dim nLineCount As Integer
Dim mstrNow As String
Dim strFname As String
    OLCCFTPGet = True

    'builds a short input file for the KMDSFtp program and then
    'executes the program
    'deletes the input file

    If gcfHandleErrors Then On Error GoTo ErrLog
    strFname = "\\" & gstrCashsysInsLoc & "\NPOS\OLCCFTPGet.TXT"

    strTemp = gstrFTPLogon
    strTemp = strTemp & "GET" & vbTab & SourceFileName & vbTab & TargetFile & vbTab & IIf(gbOLCCFTPAscii, 1, 0)

'    nAgentNumber = AgentNumber
    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\NPOS") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\NPOS")
    End If

    Set txtstrTemp = fsoTemp.OpenTextFile(strFname, ForWriting, True, TristateFalse)
    txtstrTemp.WriteLine strTemp
    txtstrTemp.Close
    'Shell "k:\kmds\kmdsftp.exe " & strFName, vbNormalFocus
    WaitForExecCmd "\\" & gstrCashsysInsLoc & "\kmds\kmdsftp.exe " & strFname

    'fsotemp.DeleteFile strFName

ExitSub:

    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing
    Exit Function

ErrLog:
    OLCCFTPGet = False
    MsgBox "Error #" & Err.Number & " " & Err.Description & vbCrLf & "\\" & gstrCashsysInsLoc & "\NPOS", vbInformation, "OLCCFTPGet"
    Resume ExitSub
End Function




Public Function KMDSUpdateGet()
Dim strTemp As String
Dim fsoTemp As New FileSystemObject
Dim txtstrTemp As TextStream
Dim nLineCount As Integer
Dim mstrNow As String
Dim strFname As String

    Dim strFTPServer As String
    Dim strFTPUser As String
    Dim strFTPPassword As String

    strFTPServer = GetKMDSSettings("KMDSFTPServer", "ftp.kmds.us")
    strFTPUser = GetKMDSSettings("KMDSFTPUser", "Agents")
    strFTPPassword = GetKMDSSettings("KMDSFTPPassword", "Ag3nt$")

    'builds a short input file for the KMDSFtp program and then
    'executes the program
    'deletes the input file

    If gcfHandleErrors Then On Error GoTo ErrLog
    strFname = "\\" & gstrCashsysInsLoc & "\KMDS\Update\KMDSUpdateGet.TXT"

    strTemp = "open " & strFTPServer & vbCrLf
    strTemp = strTemp & strFTPUser & vbCrLf
    strTemp = strTemp & strFTPPassword & vbCrLf


    strTemp = strTemp & "GET" & vbTab & "KMDSUpdate.zip" & vbTab & "KMDSUpdate.zip" & vbTab & IIf(gbOLCCFTPAscii, 1, 0)

'    nAgentNumber = AgentNumber
    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\Update") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\Update")
    End If

    Set txtstrTemp = fsoTemp.OpenTextFile(strFname, ForWriting, True, TristateFalse)
    txtstrTemp.WriteLine strTemp
    txtstrTemp.Close
    'Shell "k:\kmds\kmdsftp.exe " & strFName, vbNormalFocus
    WaitForExecCmd "\\" & gstrCashsysInsLoc & "\kmds\kmdsftp.exe " & strFname

    'fsotemp.DeleteFile strFName

ExitSub:

    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing
    Exit Function

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description & vbCrLf & "\\" & gstrCashsysInsLoc & "\NPOS", vbInformation, "OLCCFTPGet"
    Resume ExitSub
End Function
