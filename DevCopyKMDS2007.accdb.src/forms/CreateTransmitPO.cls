Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mstrNow As String

Private Sub chkFTPFZActive_Click()
    SetKMDSSettings "FTPFZActive", chkFTPFZActive
End Sub

Private Sub cmdEditThisOrder_Click()
Dim nOrderID As Long
Dim stDocName As String
Dim stLinkCriteria As String

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")


Dim nAgentNumber As Long
Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject
Dim strTemp As String
Dim nLineCount As Integer
Dim strSourceFile As String

Dim strTargetFile As String

Dim f As Folder
    If Not VerifyCurrentKMDSCustomer Then
        GoTo ExitNotCurrentKMDS
    End If
    MakeFTPScriptFiles

    If gcfHandleErrors Then On Error GoTo ErrLog
    nAgentNumber = AgentNumber
    mstrNow = Format(Now, "yymmddhhnn")

    If nAgentNumber > 1000 Then nAgentNumber = nAgentNumber - 1000

    If Not fsoTemp.FolderExists("C:\NPOS") Then
        fsoTemp.CreateFolder ("C:\NPOS")
    End If

    If Not fsoTemp.FolderExists("C:\NPOS\Inventory") Then
        fsoTemp.CreateFolder ("C:\NPOS\Inventory")
    End If
    If Not fsoTemp.FolderExists("C:\NPOS\Order") Then
        fsoTemp.CreateFolder ("C:\NPOS\Order")
    End If

    If Not fsoTemp.FolderExists("C:\NPOS\Price File") Then
        fsoTemp.CreateFolder ("C:\NPOS\Price File")
    End If


    strTemp = "C:\NPOS\Mail"

    'lets build a structure ..
    If Not fsoTemp.FolderExists(strTemp) Then
        fsoTemp.CreateFolder (strTemp)
    End If

    If Not fsoTemp.FolderExists(strTemp & "\In") Then
        fsoTemp.CreateFolder (strTemp & "\In")
    End If
    If Not fsoTemp.FolderExists(strTemp & "\In\Archive") Then
        fsoTemp.CreateFolder (strTemp & "\In\Archive")
    End If
    If Not fsoTemp.FolderExists(strTemp & "\In\ERROR") Then
        fsoTemp.CreateFolder (strTemp & "\In\ERROR")
    End If
    If Not fsoTemp.FolderExists(strTemp & "\In\Temp") Then
        fsoTemp.CreateFolder (strTemp & "\In\Temp")
    End If




    strTargetFile = "C:\NPOS\"
    strTargetFile = strTargetFile & Format(Now, "yyyy") & "\"
    If Not (fsoTemp.FolderExists(strTargetFile)) Then
          Set f = fsoTemp.CreateFolder(strTargetFile)
    End If

    strTargetFile = strTargetFile & Format(Now, "mm") & "\"
    If Not (fsoTemp.FolderExists(strTargetFile)) Then
          Set f = fsoTemp.CreateFolder(strTargetFile)
    End If

    strTargetFile = strTargetFile & Format(Now, "dd") & "\"
    If Not (fsoTemp.FolderExists(strTargetFile)) Then
          Set f = fsoTemp.CreateFolder(strTargetFile)
    End If

    strSourceFile = "Order." & CStr(Format(nAgentNumber, "000"))

    strTargetFile = strTargetFile & Format(Now, "hh nn ss ") & strSourceFile

    Set txtstrTemp = fsoTemp.OpenTextFile("C:\NPOS\" & strSourceFile, ForWriting, True, TristateFalse)
    nAgentNumber = nAgentNumber + 1000

    txtstrTemp.WriteLine (Build21Header(nAgentNumber, mstrNow))
    txtstrTemp.WriteLine (Build15Header(nAgentNumber, mstrNow))



    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryCreateTransmitPurchaseOrder"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parPurchaseOrderID", adBigInt, adParamInput)
        .Parameters("parPurchaseOrderID") = CLng(lstInvoices)
        Set rstemp = .Execute
    End With

    With rstemp
        If rstemp.RecordCount = 0 Then
            MsgBox "No Records found for Order #" & lstInvoices
            GoTo ExitSub
        End If

        Do While Not .EOF
            nLineCount = nLineCount + 1
            If Not gcfHandleErrors Then Debug.Print !LineID; !ItemID
            txtstrTemp.WriteLine BuildPOLine(nAgentNumber, !ItemID, !QuantityShipped, mstrNow)
            .MoveNext
        Loop
    End With
    rstemp.Close
    txtstrTemp.WriteLine (Build15Trailer(nAgentNumber, nLineCount, mstrNow))
    txtstrTemp.WriteLine (Build21Trailer(nAgentNumber, nLineCount, mstrNow))

ExitSub:

    txtstrTemp.Close
    fsoTemp.CopyFile "C:\NPOS\" & strSourceFile, strTargetFile, True
    fsoTemp.CopyFile "C:\NPOS\" & strSourceFile, "C:\NPOS\Order\Order.txt", True
    'added begin 31415
    If gbKMDSFTPActive Then
        CreateWinSCPtoKMDSFTPScriptFile
        fsoTemp.CopyFile "C:\NPOS\Order\Order.txt", "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\KMDSFTP\OUT\Order " & Format(Now, "yyyymmddhhnn") & ".txt", True
    End If
    'added end 31415

    MsgBox "Transmit File Created!  " '& vbcrlf & "Next click on the Oregon Liquor Control Commission to get on the website"
    Set comQDF = Nothing
    Set rstemp = Nothing
    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing
    cmdFTPSend_Click
'    DoCmd.Close acForm, Me.Name
'    cmdOLAS.Visible = True
'    cmdOLAS.SetFocus
'    cmdEditThisOrder.Visible = False
ExitNotCurrentKMDS:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmdCreate"
    Resume ExitSub
End Sub


Private Sub Command37_Click()

End Sub

Private Sub cmdOLAS_Click()
    MsgBox "Log on to the Website and then return to this screen..." & vbCrLf & vbCrLf & "Press Enter to continue.."
    cmdFTPSend.Visible = True
    cmdFTPSend.SetFocus
    cmdOLAS.Visible = False
End Sub

Private Sub Form_Load()
If gcfHandleErrors Then On Error GoTo ErrLog
    Form.InsideHeight = boxScreen.Height + 500 '+ cmdGoToFirstRecord.Height
    Form.InsideWidth = boxScreen.Width + 500
    chkFTPFZActive = GetKMDSSettings("FTPFZActive", "0") 'false
    chkUseWinSCPForOLASTransmits = GetKMDSSettings("UseWinSCPForOLASTransmits", "False") 'false

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Load"
    Resume ExitSub
End Sub

Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)
End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub


Private Sub lstInvoices_AfterUpdate()
If gcfHandleErrors Then On Error GoTo ErrLog
    txtSelectedInvoice = lstInvoices.Column(0) & " " & lstInvoices.Column(1)

    txtSelectedInvoice.Visible = Not IsNull(lstInvoices)
    cmdEditThisOrder.Visible = Not IsNull(lstInvoices)
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmbClubID_AfterUpdate"
    Resume ExitSub
End Sub



Private Sub lstInvoices_DblClick(Cancel As Integer)
 cmdEditThisOrder_Click
End Sub


Private Sub cmdClose_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdClose_Click

    DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click

End Sub
Private Sub cmdFTPSend_Click()
    If gcfHandleErrors Then On Error GoTo Err_cmdFTPSend_Click

    Dim stAppName As String
    Dim strFZDir As String
    Dim strFZExe As String
    If MsgBox("Are you connected to the internet?", vbYesNo) = vbYes Then
        If chkFTPFZActive Then

            strFZDir = GetKMDSSettings("FTPFZDir", "C:\Program Files\FileZilla FTP Client\") 'false
            If strFZDir = "C:\Program Files\FileZilla FTP Client\" Then
                SetKMDSSettings "FTPFZDir", "C:\Program Files (x86)\FileZilla FTP Client\"
            End If
            strFZDir = GetKMDSSettings("FTPFZDir", "C:\Program Files (x86)\FileZilla FTP Client\") 'false
            strFZExe = GetKMDSSettings("FTPFZExe", "FileZilla.exe") 'false
            strFZDir = Trim(strFZDir)
            strFZExe = strFZExe
            'stAppName = strFZDir & strFZExe & " -c " & Chr(34) & "0/OLCC FTP" & Chr(34)
            stAppName = strFZDir & strFZExe & " --site=" & Chr(34) & "0/OLCC FTP"
            stAppName = stAppName & "/Weekly Order" & Chr(34)

            'MsgBox "Using Filezilla to send Weekly Order" & vbCrLf & vbCrLf & "Choose 'Bookmarks -> Weekly Order" & vbCrLf & "Then double-click on order.txt"
            MsgBox "Using Filezilla to send Weekly Order" & vbCrLf & vbCrLf & "Double-click on order.txt"
'            DoCmd.OpenForm "Filezilla"
            Call Shell(stAppName, 1)
            MsgBox "Please verify your order was received by the OLCC by checking on the Website"
            MsgBox "NOTE: It may take up to 5 minutes before your order can be checked..."
            DoCmd.Close acForm, Me.Name
            'DoCmd.Quit
        Else
            'added 31415
            If gbKMDSFTPActive Then
                EOLog "Send order file to ftp.kmds.us"
                WinSCPFTPTransmitKMDS
                EOLog "... Complete"
            End If

            If GetKMDSSettings("UseWinSCPForOLASTransmits", "False") = "True" Then
                'build a po transmit files
                '1. WinSCPDaily.txt 'script file used by :
                '2. WinSCPDaily.bat

                CreateWinSCPtoOLASFTPScriptFile ("Order")
                WaitForExecCmd "c:\npos\WinSCPOrderOLCC.BAT"

            Else
                stAppName = "c:\npos\ftporder.bat"
                Call Shell(stAppName, 1)
                MsgBox "Please verify your order was received by the OLCC by checking on the Website"
                MsgBox "NOTE: It may take up to 5 minutes before your order can be checked..."
                DoCmd.Close acForm, Me.Name
                DoCmd.Quit
            End If
            'END 31415
        End If
    Else
        MsgBox "Please connect to the internet and then start the process over"
        DoCmd.Close acForm, Me.Name
    End If



Exit_cmdFTPSend_Click:
    Exit Sub

Err_cmdFTPSend_Click:
    MsgBox Err.Description
    Resume Exit_cmdFTPSend_Click

End Sub
