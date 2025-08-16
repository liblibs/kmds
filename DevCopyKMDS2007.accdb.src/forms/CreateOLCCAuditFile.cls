Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mstrNow As String


Private Sub cmdEditThisOrder_Click()
Dim nOrderID As Long
Dim stDocName As String
Dim stLinkCriteria As String
Dim strSource As String
Dim strTarget As String

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
Dim nLineCount99 As Integer

    If gcfHandleErrors Then On Error GoTo ErrLog
    If Not fsoTemp.Drives.Item(txtDrive).IsReady Then
        If MsgBox("This requires a disk in " & txtDrive & vbCrLf & "Please make sure there is before you continue..", vbOKCancel) = vbCancel Then
            Exit Sub
        End If
    End If
    nAgentNumber = AgentNumber
    mstrNow = Format(Now, "yymmddhhnn")

    If nAgentNumber > 1000 Then nAgentNumber = nAgentNumber - 1000

    If Not fsoTemp.FolderExists("C:\NPOS") Then
        fsoTemp.CreateFolder ("C:\NPOS")
    End If
    If Not fsoTemp.FolderExists("C:\NPOS\AuditFile") Then
        fsoTemp.CreateFolder ("C:\NPOS\AuditFile")
    End If

    strSource = txtDrive & "\olf00021." & CStr(Format(nAgentNumber, "000"))
    strTarget = "C:\NPOS\AuditFile\olf00021-" & Format(Now, "yymmddhhnn") & "." & CStr(Format(nAgentNumber, "000"))

    Debug.Print strSource
    Debug.Print strTarget
    Set txtstrTemp = fsoTemp.OpenTextFile(strSource, ForWriting, True, TristateFalse)
    nAgentNumber = nAgentNumber + 1000

    txtstrTemp.WriteLine (Build21Header(nAgentNumber, mstrNow))

        nLineCount = 0
        'create inventory file
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
                nLineCount99 = nLineCount99 + 1
                'If not gcfHandleErrors then Debug.Print !ItemID
                txtstrTemp.WriteLine BuildInvLine(nAgentNumber, !ID, !Onhand, mstrNow)
                .MoveNext
            Loop
        End With
        rstemp.Close
        txtstrTemp.WriteLine (Build84Trailer(nAgentNumber, nLineCount, mstrNow))

    txtstrTemp.WriteLine (Build21Trailer(nAgentNumber, nLineCount99, mstrNow))
    txtstrTemp.Close

    MsgBox "Audit File Created!"
    SetKMDSSettings "AuditFileDrive", txtDrive

    DoCmd.Close , Me.Name

ExitSub:
    If gcfHandleErrors Then On Error Resume Next
    Set comQDF = Nothing
    Set rstemp = Nothing
    Set txtstrTemp = Nothing
    fsoTemp.CopyFile strSource, strTarget, True
    Set fsoTemp = Nothing
    DoCmd.Close acForm, Me.Name

    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmdCreate"
    Resume ExitSub
End Sub




Private Sub Form_Load()

Dim fsoTemp As New FileSystemObject
Dim nI As Integer
Dim nasc As Integer
    txtDrive = GetKMDSSettings("AuditFileDrive", "F:")

If gcfHandleErrors Then On Error GoTo ErrLog
    Form.InsideHeight = boxScreen.Height + 500 '+ cmdGoToFirstRecord.Height
    Form.InsideWidth = boxScreen.Width + 500
    nasc = Asc("A")
    For nI = 0 To 25
        If fsoTemp.DriveExists(Chr(nasc + nI)) Then
            If fsoTemp.Drives.Item(Chr(nasc + nI)).DriveType = Removable Then
                lstDrives.AddItem fsoTemp.GetDrive(Chr(nasc + nI))
            End If
        End If
    Next nI
    For nI = 0 To 25
        If fsoTemp.DriveExists(Chr(nasc + nI)) Then
            If fsoTemp.Drives.Item(Chr(nasc + nI)).DriveType <> Removable Then
                lstDrivesFixed.AddItem fsoTemp.GetDrive(Chr(nasc + nI))
            End If
        End If
    Next nI

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
Private Sub cmdEmail_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdEmail_Click

    Dim stDocName As String

    'Dim comQDF As New Command
    Dim comQDF As Object
    Set comQDF = CreateObject("ADODB.COMMAND")


    SetOptions False
    If gcfHandleErrors Then On Error Resume Next
    DoCmd.DeleteObject acTable, "HourlySales"

If gcfHandleErrors Then On Error GoTo Err_cmdEmail_Click
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qrySalesByHourMakeTable"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("BeginDate", adDate, adParamInput)
        .Parameters("BeginDate") = Format(Now - 60, "mm/dd/yyyy 00:00")
        .Parameters.Append .CreateParameter("EndDate", adDate, adParamInput)
        .Parameters("EndDate") = Format(Now, "mm/dd/yyyy 23:59")
        .Execute
    End With
    Set comQDF = Nothing
    DoCmd.SendObject acSendTable, "HourlySales", acFormatXLS, , , , "Hourly Sales from " & gstrStoreName, "Attached is the file you requested!"

Exit_cmdEmail_Click:
    Exit Sub

Err_cmdEmail_Click:
    MsgBox Err.Description
    Resume Exit_cmdEmail_Click

End Sub

Private Sub lstDrives_Click()
    txtDrive = lstDrives
End Sub

Private Sub lstDrivesFixed_Click()
    txtDrive = lstDrivesFixed


End Sub
'------------------------------------------------------------
' cmdAuditorsAuditFile_Click
'
'------------------------------------------------------------
Private Sub cmdAuditorsAuditFile_Click()
On Error GoTo cmdAuditorsAuditFile_Click_Err
    MsgBox "Please insert the flash drive from the auditor's before continuing!", vbCritical
    DoCmd.OpenForm "AuditorsAuditFile", acNormal, "", "", , acNormal
    DoCmd.Close acForm, Me.Name


cmdAuditorsAuditFile_Click_Exit:
    Exit Sub

cmdAuditorsAuditFile_Click_Err:
    MsgBox Error$
    Resume cmdAuditorsAuditFile_Click_Exit

End Sub
