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
    Dim fsoTemp As New FileSystemObject
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

    MakeFTPScriptFiles

    OLCCInventFile CLng(AgentNumber), False
    'create an archive of the file... increment agent number so tonite's close won't overwrite the xmit file...
    OLCCInventFile CLng(AgentNumber + 1), True


    MsgBox "Transmit File Created!  " '& vbcrlf & "Next click on the Oregon Liquor Control Commission to get on the website"
    cmdFTPSend_Click
'    DoCmd.Close acForm, Me.Name
    'cmdOLAS.Visible = True
    'cmdOLAS.SetFocus
    'cmdEditThisOrder.Visible = False

End Sub


Private Sub Command32_Click()
    OLCCInventFile CLng(AgentNumber), False
    'create an archive of the file... increment agent number so tonite's close won't overwrite the xmit file...
    OLCCInventFile CLng(AgentNumber + 1), True
    MsgBox "Transmit File Created!"
End Sub

Private Sub cmdFTPSend_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdFTPSend_Click
    Dim strFZDir As String
    Dim strFZExe As String

    Dim stAppName As String
    If MsgBox("Are you logged on to the Internet?", vbYesNo) = vbYes Then

        If gbKMDSFTPActive Then
            'MsgBox "KMDS FTP Active; sending to kmds"
            CreateWinSCPtoKMDSFTPScriptFile
            WinSCPFTPTransmitKMDS
        End If
        If chkFTPFZActive Then
            strFZDir = GetKMDSSettings("FTPFZDir", "C:\Program Files (x86)\FileZilla FTP Client\") 'false
            If strFZDir = "C:\Program Files\FileZilla FTP Client\" Then
                SetKMDSSettings "FTPFZDir", "C:\Program Files (x86)\FileZilla FTP Client\"
            End If
            strFZDir = GetKMDSSettings("FTPFZDir", "C:\Program Files (x86)\FileZilla FTP Client\") 'false
            strFZExe = GetKMDSSettings("FTPFZExe", "FileZilla.exe") 'false
            strFZDir = Trim(strFZDir)
            strFZExe = strFZExe
            'stAppName = strFZDir & strFZExe & " -c " & Chr(34) & "0/OLCC FTP" & Chr(34)
            stAppName = strFZDir & strFZExe & " --site=" & Chr(34) & "0/OLCC FTP"
            stAppName = stAppName & "/Monthly Inventory" & Chr(34)

            'MsgBox "Using Filezilla to send Monthly Inventory " & vbCrLf & vbCrLf & "Choose 'Bookmarks -> Monthly Inventory" & vbCrLf & "Then double-click on invent.txt"
            MsgBox "Using Filezilla to send Monthly Inventory " & vbCrLf & vbCrLf & "Double-click on invent.txt"
            Call Shell(stAppName, 1)
            MsgBox "Please verify your order was received by the OLCC by checking on the Website"
            MsgBox "NOTE: It may take up to 5 minutes before your order can be checked..."
            DoCmd.Close acForm, Me.Name
            'DoCmd.Quit
        Else
            If GetKMDSSettings("UseWinSCPForOLASTransmits", "False") = "True" Then
                'build a po transmit files
                '1. WinSCPDaily.txt 'script file used by :
                '2. WinSCPDaily.bat

                CreateWinSCPtoOLASFTPScriptFile ("MonthEnd")
                WaitForExecCmd "c:\npos\WinSCPMonthendOLCC.BAT"

            Else

                stAppName = "c:\npos\ftpInvent.bat"
                Call Shell(stAppName, 1)
                MsgBox "Please verify your INVENTORY was received by the OLCC by checking on the Website"
                MsgBox "NOTE: It may take up to 5 minutes before your INVENTORY can be checked..."
                DoCmd.Close acForm, Me.Name
                DoCmd.Quit
            End If
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

Private Sub cmdOLAS_Click()
    MsgBox "Log on to the Website and then return to this screen..." & vbCrLf & vbCrLf & "Press Enter to continue.."
    cmdFTPSend.Visible = True
    cmdFTPSend.SetFocus
    cmdOLAS.Visible = False
End Sub

Private Sub cmdReTransmit_Click()
    Dim strMsg As String
    Dim strDate As String
    Dim dtDate As Date
    Dim fsoTemp As New FileSystemObject

    strDate = Month(Now)
    strDate = strDate & "/01/" & Year(Now)
    dtDate = CDate(strDate)
    dtDate = DateAdd("d", -1, dtDate)

    strMsg = "Use this option if you need to transmit the month end inventory AFTER the 1st of the month."
    strMsg = strMsg & vbCrLf & vbCrLf & "Choosing Yes will retransmit the inventory ending " & dtDate

    strMsg = strMsg & vbCrLf & vbCrLf & "Do you want to continue?"
    If MsgBox(strMsg, vbYesNo, "Retransmit " & dtDate) = vbNo Then
        Exit Sub
    End If
    'make sure all transactions have been rolled up into the item movement table..
    lblPatience.Visible = True
    Me.Repaint
    SQLPassThru "sp_UpdateItemMovement"

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

    MakeFTPScriptFiles

    OLCCInventFileAfterNewMonthHasBegun CLng(AgentNumber), False
    'create an archive of the file... increment agent number so tonite's close won't overwrite the xmit file...
    OLCCInventFileAfterNewMonthHasBegun CLng(AgentNumber + 1), True

    lblPatience.Visible = False

    MsgBox "Transmit File Created!  " '& vbcrlf & "Next click on the Oregon Liquor Control Commission to get on the website"
    cmdFTPSend_Click
'    DoCmd.Close acForm, Me.Name
    'cmdOLAS.Visible = True
    'cmdOLAS.SetFocus
    'cmdEditThisOrder.Visible = False


End Sub

Private Sub Form_Load()
If gcfHandleErrors Then On Error GoTo ErrLog
    Form.InsideHeight = boxScreen.Height + 500 '+ cmdGoToFirstRecord.Height
    Form.InsideWidth = boxScreen.Width + 500
    chkFTPFZActive = GetKMDSSettings("FTPFZActive", "0") 'false
    '    added 31415
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


Private Sub cmdClose_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdClose_Click


    DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click

End Sub
