Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdEmail_Click()
Dim savStr As String
Dim coleMail As New Collection

If gcfHandleErrors Then On Error GoTo Err_cmdEmail_Click

    Dim stDocName As String
    Dim strMsg As String
    strMsg = "Please find attached a copy of the transfer from " & gstrStoreName
    If EMail <> "" Then
        stDocName = "AgentTransfer"
        savStr = ArchiveReportShare(stDocName, "XFR " & txtAgentTransferID)

        If gcfHandleErrors Then On Error Resume Next
        coleMail.Add savStr
        KMDSeMail EMail, coleMail, "Transfer from " & gstrStoreName

        'DoCmd.SendObject acReport, stDocName, acFormatSNP, EMail, , , "KMDS Transfer " & txtAgentTransferID & " from " & gstrStoreName, strMsg
    End If



Exit_cmdEmail_Click:
    Exit Sub

Err_cmdEmail_Click:
    MsgBox Err.Description
    Resume Exit_cmdEmail_Click

End Sub
Private Sub cmdKMDSFTP_Click()
    Dim stOutputName As String
    Dim strFTP As String
    Dim fsoTemp As New FileSystemObject

    If gcfHandleErrors Then On Error GoTo ErrLog

    If Not fsoTemp.FolderExists("C:\KMDS\FTP") Then
        fsoTemp.CreateFolder ("C:\KMDS\FTP")
    End If
   If Not fsoTemp.FolderExists("C:\KMDS\FTP\OUT") Then
        fsoTemp.CreateFolder ("C:\KMDS\FTP\OUT")
    End If
    Set fsoTemp = Nothing

    stOutputName = "C:\KMDS\FTP\OUT\" & AgentNumber & "-" & CStr(glnAgentNumber) & "-" & txtAgentTransferID & ".csv"
    DoCmd.TransferText acExportDelim, , "qryAgentTransferLinesAndSysfile", stOutputName, True

    If KMDScustomer Then
            'begin added 31415
            If GetKMDSSettings("UseWinSCPForOLASTransmits", "False") = "True" Then
                'build a po transmit files
                '1. WinSCPDaily.txt 'script file used by :
                '2. WinSCPDaily.bat

                WinSCPtoKMDSFTPScriptFileXFR "OUT", stOutputName, AgentNumber

            Else
                MakeXFRScript "Send", CStr(AgentNumber), stOutputName
                FTPXfr
                Name stOutputName As stOutputName & ".txt"
                WaitForExecCmd "c:\kmds\Dirdater.exe C:\KMDS\FTP\OUT\"
                MsgBox "Please notify the receiving agent to CHECK FOR TRANSFERS!"
            End If
            'END 31415

    End If
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmdKMDSFTP_Click"
    Resume ExitSub
End Sub

Private Sub Command11_Click()
    Dim stDocName As String
    Dim strOpenArgs As String
    If gcfHandleErrors Then On Error GoTo ErrLog

    'added 2007-12-29
    Select Case Status
    Case "C"
        stDocName = "AgentTransferClaim"
    Case Else
        stDocName = "AgentTransfer" '"PurchaseOrderLiquor"
    End Select

'    If InStr(stDocName, "(NoPrice)") > 0 Then
'        stDocName = Replace(stDocName, "(NoPrice)", "")
'        strOpenArgs = "NoPrice"
'    End If

    ArchiveReport stDocName, AgentNumber & " " & txtAgentTransferID

    Select Case optChoices.Value
    Case 1 ' print
        DoCmd.OpenReport stDocName, acViewPreview
        If IsNumeric(txtNumberOfCopies) And txtNumberOfCopies > 0 Then
            If txtNumberOfCopies > 5 Then txtNumberOfCopies = 5
            DoCmd.PrintOut , , , , CInt(txtNumberOfCopies)
        End If
        DoCmd.Close acReport, stDocName

        'DoCmd.OpenReport stDocName, acViewNormal, , , , strOpenArgs
    Case 2 'post
    'invoice print

        ArchiveReportDocument stDocName, "TransferOrder", AgentNumber, txtAgentTransferID
        If gcfHandleErrors Then On Error Resume Next

        DoCmd.OpenReport stDocName, acViewPreview
        If IsNumeric(txtNumberOfCopies) And txtNumberOfCopies > 0 Then
            If txtNumberOfCopies > 5 Then txtNumberOfCopies = 5
            DoCmd.PrintOut , , , , CInt(txtNumberOfCopies)
        End If
        DoCmd.Close acReport, stDocName

        If gcfHandleErrors Then On Error GoTo ErrLog
        PostTransfer CLng(txtAgentTransferID), LastName ', "I" ', Status


    Case Else
        MsgBox "Unhandled case : optchoices"
    End Select
    DoCmd.Close acForm, Me.Name

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Command11_Click"
    Resume ExitSub
End Sub
Private Sub cmdClose_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdClose_Click


        DoCmd.Close acForm, Me.Name

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click

End Sub
Private Sub cmdPrintPreview_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdPrintPreview_Click

    Dim stDocName As String
    Dim strOpenArgs As String

    'added 2007-12-29
    Select Case Status
    Case "C"
        stDocName = "AgentTransferClaim"
    Case Else

        stDocName = "AgentTransfer" '"PurchaseOrderLiquor"
    End Select
'    If InStr(stDocName, "(NoPrice)") > 0 Then
'        stDocName = Replace(stDocName, "(NoPrice)", "")
'        strOpenArgs = "NoPrice"
'    End If




    DoCmd.OpenReport stDocName, acPreview, , , , strOpenArgs

Exit_cmdPrintPreview_Click:
    Exit Sub

Err_cmdPrintPreview_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrintPreview_Click

End Sub


Private Sub Form_Load()
    cmdKMDSFTP.Visible = KMDScustomer
    If IsNull(EMail) Or EMail = "" Then
        cmdEmail.Visible = False
    Else
        cmdEmail.Visible = True
    End If
    chkUseWinSCPForOLASTransmits = GetKMDSSettings("UseWinSCPForOLASTransmits", "False") 'false

End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim strAgentTransferID As String
    Dim rs As DAO.Recordset

    txtAgentTransferID = Me.OpenArgs
    If IsNull(AgentNumber) Then
        MsgBox "Unknown Purchase Order Number, try printing again!"
        Cancel = True
    End If

    If Not IsNull(Me.OpenArgs) Then
        strAgentTransferID = Me.OpenArgs
        Set rs = CurrentDb().OpenRecordset("select * from AgentTransferHead where id = " & CStr(strAgentTransferID), dbOpenForwardOnly, dbSeeChanges)
        If Not rs.EOF Then
            Status = rs!Status

        End If
        rs.Close
        Set rs = Nothing
    End If


End Sub



Private Sub Command37_Click()
If gcfHandleErrors Then On Error GoTo Err_Command37_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command37_Click:
    Exit Sub

Err_Command37_Click:
    MsgBox Err.Description
    Resume Exit_Command37_Click

End Sub

Private Sub optChoices_AfterUpdate()
    Select Case optChoices.Value
    Case 1 ' print
        txtNumberOfCopies = 1
        'DoCmd.OpenReport stDocName, acViewNormal, , , , strOpenArgs
    Case 2 'post
    'invoice print
        Select Case Status
        Case "C"
            txtNumberOfCopies = GetKMDSSettings("TransferCopyClaims", 1)
        Case "O"
            txtNumberOfCopies = GetKMDSSettings("TransferCopyOUT", 1)
        Case "I"
            txtNumberOfCopies = GetKMDSSettings("TransferCopyIN", 1)
        Case Else
        End Select
    Case Else
        MsgBox "Unhandled case : optchoices"
    End Select

End Sub

Private Sub txtNumberOfCopies_AfterUpdate()
    Dim strSQL As String
    strSQL = "Save the number of copies (" & txtNumberOfCopies & ") as the default for "
    Select Case Status
    Case "C"
        strSQL = strSQL & "CLAIMS"
    Case "O"
        strSQL = strSQL & "Transfer OUTS"
    Case "I"
        strSQL = strSQL & "Transfer INS"
    Case Else
    End Select
    strSQL = strSQL & "?"
    If MsgBox(strSQL, vbYesNo, "Change default number of copies?") = vbYes Then
        Select Case Status
        Case "C"
            SetKMDSSettings "TransferCopyClaims", txtNumberOfCopies
        Case "O"
            SetKMDSSettings "TransferCopyOUT", txtNumberOfCopies
        Case "I"
            SetKMDSSettings "TransferCopyIN", txtNumberOfCopies
        Case Else
        End Select

    End If
End Sub
