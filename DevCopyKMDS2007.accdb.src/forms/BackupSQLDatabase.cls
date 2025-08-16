Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdBackup_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdBackup_Click
    FetchJournals
    Me.Repaint
    BackUpSQL


Exit_cmdBackup_Click:
    Exit Sub

Err_cmdBackup_Click:
    EMsgBox Me, "cmdBackup_Click"
    Resume Exit_cmdBackup_Click

End Sub

'Private Sub cmdBackup_Click()
'Dim strMsg As String
'If gcfHandleErrors Then On Error GoTo Err_cmdBackup_Click
'    FetchJournals
'    Me.Repaint
'    BackUpSQL
'    Me.Repaint
'    strMsg = "Database has been backed up on the Hard drive" & vbcrlf & vbcrlf
'    strMsg = strMsg & "Do you want to make a cd-backup of the data?"
'
'
'    If MsgBox(strMsg, vbYesNo) = vbYes Then
'        'MsgBox "Please Insert a new CD"
'        MsgBox "Choose the 'Write these files to CD' option under the CD Writing Tasks "
'        Shell "C:\KMDS\KMDSBurn.bat", vbNormalFocus
'        DoCmd.Close acForm, Me.Name
'        DoCmd.Quit
'    End If
'
'
'
'Exit_cmdBackup_Click:
'    Exit Sub
'
'Err_cmdBackup_Click:
'    EMsgBox Me, "cmdBackup_Click"
'    Resume Exit_cmdBackup_Click
'
'End Sub

Private Sub cmdfrmClose_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    DoCmd.Close

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmClose_Click"
    Resume ExitSub
End Sub

Private Sub cmdOnlineBackup_Click()

Dim stOutputName As String
Dim strFTP As String
Dim strTemp As String
Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject

Dim strPassString

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")


    If gcfHandleErrors Then On Error GoTo ErrLog

    If glnAgentNumber = 0 Then LoadDefaults
    If glnAgentNumber < 1000 Then glnAgentNumber = glnAgentNumber + 1000

    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "Select CashsysInsLoc from Sysfile"
        .CommandType = adCmdText
        'ADO Numeric Datatypes are very particular
        Set rstemp = .Execute
        'If rsTemp.BOF Then
            rstemp.MoveFirst
            strPassString = "c:\kmds\KMDSOLB.bat " & rstemp!CashsysInsLoc & " " & glnAgentNumber & " " & glnVersionNumber
        'End If
    End With
    rstemp.Close
    Set rstemp = Nothing
    Set comQDF = Nothing



    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL")
    End If
    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT")
    End If

    If Not gcfHandleErrors Then Debug.Print strPassString
    DoCmd.OutputTo acOutputQuery, "qryKMDGlobalUPCItems", acFormatXLS, "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\KMDGlobalUPCItems-" & glnAgentNumber & ".xls", False
    DoCmd.OutputTo acOutputQuery, "qryKMDGlobalInvolineUPC", acFormatXLS, "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT\KMDGlobalInvolineUPC-" & glnAgentNumber & ".xls", False


    Shell strPassString, vbNormalFocus


ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmdKMDSFTP_Click"
    Resume ExitSub
End Sub

Private Sub cmdRestore_Click()
Dim strMsg As String
If gcfHandleErrors Then On Error GoTo Err_cmdBackup_Click
    strMsg = "This will restore the database to the last time it was backed up!" & vbCrLf & vbCrLf

    strMsg = strMsg & "All the work since the last backup will need to be redone!" & vbCrLf & vbCrLf
    strMsg = strMsg & "Are you sure you want to continue?"
    If MsgBox(strMsg, vbCritical + vbYesNo) = vbYes Then
        RestoreSQL
        DoCmd.Quit

    End If

Exit_cmdBackup_Click:
    Exit Sub

Err_cmdBackup_Click:
    EMsgBox Me, "cmdRestore_Click"
    Resume Exit_cmdBackup_Click
End Sub

Private Sub Command10_Click()
    Dim strPassString As String
    strPassString = "c:\kmds\KMDSftpc.bat " & glnAgentNumber
    Shell strPassString, vbNormalFocus

End Sub

Private Sub Command12_Click()
    Dim strPassString As String
    strPassString = "c:\kmds\kmdsgetver.bat "
    Shell strPassString, vbNormalFocus

End Sub

Private Sub Detail_DblClick(Cancel As Integer)
cmdRestore.Visible = True
End Sub

Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)

End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub

'------------------------------------------------------------
' cmdChronJob_Click
'
'------------------------------------------------------------
Private Sub cmdChronJob_Click()
'added 31416
Dim fsoTemp As New FileSystemObject
Dim txtstrTemp As TextStream
Dim strTargetFile As String


On Error GoTo cmdChronJob_Click_Err

    If Not gstrCashsysInsLoc = MyComputerName Then
        MsgBox "Run on SERVER ONLY!!!!", vbCritical
    Else
        txtStartTime.Visible = True
        txtAgentNumber.Visible = True
        txtSCHTask.Visible = True
        'build the command line
        txtStartTime = 0
        txtStartTime = DateAdd("n", txtAgentNumber - 1000, 0)
        txtStartTime = Format(txtStartTime, "Short time")
        txtSCHTask = "SCHTASKS /CREATE /SC DAILY /TN " & Chr(34) & "KMDSTasks\OnlineBackup" & Chr(34)
        txtSCHTask = txtSCHTask & " /TR " & Chr(34) & "C:\kmds\OnlineBackupWinSCP.bat" & Chr(34)
        txtSCHTask = txtSCHTask & " /ST " & txtStartTime
        txtSCHTask = txtSCHTask & " /RU " & Chr(34) & gstrCashsysInsLoc & "\Owner" & Chr(34)
        txtSCHTask = txtSCHTask & " /RP " & Chr(34) & "1234567" & Chr(34)
        Debug.Print txtSCHTask


        'output to a batch file
        strTargetFile = "C:\KMDS\OnlineBackupSCHTask.bat"
        Set txtstrTemp = fsoTemp.OpenTextFile(strTargetFile, ForWriting, True, TristateFalse)
        txtstrTemp.WriteLine
        txtstrTemp.WriteLine txtSCHTask
        txtstrTemp.WriteLine "Pause"
        txtstrTemp.Close
        Set txtstrTemp = Nothing
        Set fsoTemp = Nothing
        'run the batch file
        WaitForExecCmd strTargetFile
        MsgBox "Task created!  Backup will be sent to ftp.kmds.us every day at " & txtStartTime
    End If


cmdChronJob_Click_Exit:
    Exit Sub

cmdChronJob_Click_Err:
    MsgBox Error$
    Resume cmdChronJob_Click_Exit

End Sub
