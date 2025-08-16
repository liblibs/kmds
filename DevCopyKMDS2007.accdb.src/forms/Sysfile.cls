Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Dim gbRunMailChanges As Boolean


Private Sub CashsysInsLoc_BeforeUpdate(Cancel As Integer)
    If MsgBox("Only change this when directed by KMDS Support" & vbCrLf & "Do you want to change this?", vbYesNo) <> vbYes Then
        Cancel = True
        CashsysInsLoc.Undo
    Else
        cmdSetConnectInvoline.Visible = True
    End If
End Sub

Private Sub chkKMDSFTPActive_Click()

    If chkKMDSFTPActive Then
        SetKMDSSettings "KMDSFTPActive", "True"
    Else
        SetKMDSSettings "KMDSFTPActive", "False"
    End If
End Sub

Private Sub chkUseWinSCPForOLASTransmits_Click()
    If chkUseWinSCPForOLASTransmits Then
        SetKMDSSettings "UseWinSCPForOLASTransmits", "True"
    Else
        SetKMDSSettings "UseWinSCPForOLASTransmits", "False"
    End If
End Sub

Private Sub cmdCashCopy_Click()
    If MsgBox("Copy updates to all registers from " & gstrCashsysInsLoc & vbCrLf & " Are you sure?", vbYesNo) = vbYes Then
        CopyUpdates
    End If
End Sub

'Private Sub cmdColorDropDown_Click()
''
''    cmdlgColor.ShowColor
''    ColorDropDown = cmdlgColor.Color
''    ColorDropDown.BackColor = cmdlgColor.Color
'End Sub
'
'Private Sub cmdColorInfo_Click()
'    cmdlgColor.ShowColor
'    ColorInfo = cmdlgColor.Color
'    ColorInfo.BackColor = cmdlgColor.Color
'End Sub
'
'Private Sub cmdColorLockedForAll_Click()
'    cmdlgColor.ShowColor
'    ColorLockedForAll = cmdlgColor.Color
'    ColorLockedForAll.BackColor = cmdlgColor.Color
'End Sub
'
'Private Sub cmdColorTotals_Click()
'    cmdlgColor.ShowColor
'    ColorTotals = cmdlgColor.Color
'    ColorTotals.BackColor = cmdlgColor.Color
'End Sub
'
'Private Sub cmdColorViewOnly_Click()
'    cmdlgColor.ShowColor
'    ColorViewOnly = cmdlgColor.Color
'    ColorViewOnly.BackColor = cmdlgColor.Color
'End Sub

Private Sub cmdfrmAddNew_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdfrmAddNew

    DoCmd.GoToRecord , , acNewRec

Exit_cmdfrmAddNew:
    Exit Sub

Err_cmdfrmAddNew:
    MsgBox Err.Description
    Resume Exit_cmdfrmAddNew
End Sub

Private Sub cmdCopyRegisterInvoiceTemplate_Click()

    '2006-12-5 'created crview.ins
    '               used by crview.exe for valid registers
    '10/03/2004 added "confirm" Parameter
    'Dim rstemp As New ADODB.Recordset
    Dim rstemp As Object
    Set rstemp = CreateObject("ADODB.Recordset")

    Dim bAtLeastOneRegisterCheckedOut
    Dim strValidRegisters As String

    Dim strMsg As String

    Dim fsoTemp As New FileSystemObject
    Dim filTemp As File



    If gcfHandleErrors Then On Error GoTo ExitSub

    'Cancel = FormLogin(glnCashierID, Me)

    bAtLeastOneRegisterCheckedOut = False
    rstemp.Open "qryCashRegisters", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    With rstemp
        Do While Not .EOF
            Select Case !ID
            Case 1
                If !Cr1Installed = True Then
                    fsoTemp.CopyFile "k:\versions\CashRegisterDetailTemplate.mdb", "\\" & !ComputerName & "\kmds\CashRegisterDetailTemplate.mdb", True
                    strMsg = !ComputerName & " success" & vbCrLf
                End If

            Case 2
                If !Cr2Installed = True Then
                    fsoTemp.CopyFile "k:\versions\CashRegisterDetailTemplate.mdb", "\\" & !ComputerName & "\kmds\CashRegisterDetailTemplate.mdb", True
                    strMsg = strMsg & !ComputerName & " success" & vbCrLf
                End If
            Case 3
                If !Cr3Installed = True Then
                    fsoTemp.CopyFile "k:\versions\CashRegisterDetailTemplate.mdb", "\\" & !ComputerName & "\kmds\CashRegisterDetailTemplate.mdb", True
                    strMsg = strMsg & !ComputerName & " success" & vbCrLf
                End If
            Case 4
                If !Cr4Installed = True Then
                    fsoTemp.CopyFile "k:\versions\CashRegisterDetailTemplate.mdb", "\\" & !ComputerName & "\kmds\CashRegisterDetailTemplate.mdb", True
                    strMsg = strMsg & !ComputerName & " success" & vbCrLf
                End If
            Case 5
                If !Cr5Installed = True Then
                    fsoTemp.CopyFile "k:\versions\CashRegisterDetailTemplate.mdb", "\\" & !ComputerName & "\kmds\CashRegisterDetailTemplate.mdb", True
                    strMsg = strMsg & !ComputerName & " success" & vbCrLf
                End If
            End Select
            .MoveNext
        Loop
        .Close
    End With
    Set rstemp = Nothing

    MsgBox strMsg

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "cmdCopyRegisterInvoiceTemplate_Click"
    Resume ExitSub

End Sub



Private Sub cmdfrmClose_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    DoCmd.Close

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmClose_Click"
    Resume ExitSub
End Sub

Private Sub cmdfrmDeleteRecord_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    DoCmd.DoMenuItem acFormBar, acEditMenu, acSelectRecord, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, acDelete, , acMenuVer70



ExitSub:
    Exit Sub

ErrLog:
    Select Case Err.Number
    Case 2501
        MsgBox "There are related items for this record"
    Case Else
    EMsgBox Me, "cmdfrmDeleteRecord_Click"
    End Select
    Resume ExitSub
End Sub

Private Sub cmdfrmSaveRecord_Click()
If gcfHandleErrors Then On Error GoTo ErrLog


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    cmdfrmClose.SetFocus
    cmdfrmUndo.Visible = False
    cmdfrmSaveRecord.Visible = False

    'refresh combo box,
'   cboID.Requery
'   cboID = ID
'   cboID.Visible = True
'   cboID.SetFocus

ExitSub:
    Exit Sub

ErrLog:
    If Not gcfHandleErrors Then Debug.Print Err.Number & " " & Err.Description
    Select Case Err.Number
    Case 2501 ' SAVE WAS CANCELED
    Case Else
        MsgBox "Error #" & Err.Number & " " & Err.Description, , "cmdfrmSaveRecord_Click"
    End Select
    Resume ExitSub
End Sub

Private Sub cmdfrmUndo_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
    'cmdfrmAddNew.SetFocus
    Me.SetFocus
    cmdfrmUndo.Visible = False
    cmdfrmSaveRecord.Visible = False

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmUndo_Click"
    Resume ExitSub
End Sub

Private Sub cmdKMDSGet_Click()
    If MsgBox("Download Current Software", vbYesNo) = vbYes Then
        WaitForExecCmd ("K:KMDSGet.bat")
    End If

End Sub

Private Sub cmdLookForDayEndFiles_Click()
    If Me.Dirty Then
        Me.Dirty = False
    End If
    LookForDayEndFiles
End Sub

Private Sub cmdOLCCItems_Click()
    GetOLCCItemFiles
    MsgBox "Done!"


End Sub

Private Sub cmdReindex_Click()
    If MsgBox("Reindex Tables are you sure?", vbYesNo) = vbYes Then
        If MsgBox("This will lock the files for about 2 minutes... Make sure there is no activity at the registers... OK?", vbYesNo) = vbYes Then
            ReindexTables
        End If
    End If
End Sub

Private Sub Command290_Click()

End Sub

Private Sub cmdSaveEMailAddress_Click()
    If InStr(txtEmail, "@") = 0 Then
        MsgBox txtEmail & " invalid EMail... not defined", vbCritical, "cmdSaveEMailAddress"
        Exit Sub
    End If

    SetKMDSSettings "KMDSMail EODeMail", txtEmail

End Sub



Private Sub Form_BeforeUpdate(Cancel As Integer)
    'ColorLockedForAll must be unique!
    Dim bColorLockedForAll As Boolean
    bColorLockedForAll = False
    If ColorLockedForAll = ColorDropDown Then
        bColorLockedForAll = True
    End If
    If ColorLockedForAll = ColorInfo Then
        bColorLockedForAll = True
    End If
    If ColorLockedForAll = ColorTotals Then
        bColorLockedForAll = True
    End If
    If ColorLockedForAll = ColorViewOnly Then
        bColorLockedForAll = True
    End If
    If bColorLockedForAll Then
        MsgBox "Color Locked For All must be unique! "
        Cancel = True
        ColorLockedForAll.SetFocus
        Exit Sub
    End If


    ChangeDate = Now
    Me.Repaint
End Sub

Private Sub Form_Current()
    If gcfHandleErrors Then On Error GoTo ErrLog
    ColorDropDown.BackColor = ColorDropDown
    ColorInfo.BackColor = ColorInfo
    ColorTotals.BackColor = ColorTotals
    ColorViewOnly.BackColor = ColorViewOnly
    ColorLockedForAll.BackColor = ColorLockedForAll
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "Form_Current"
    Resume ExitSub
End Sub

Private Sub Form_Dirty(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog
    cmdfrmSaveRecord.Visible = True
    cmdfrmUndo.Visible = True
    gbRunMailChanges = True
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Dirty"
    Resume ExitSub
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
'Dim nCode As Long
Dim rs As Object

    If gcfHandleErrors Then On Error GoTo ErrLog
    Select Case DataErr
    Case 3146 'description = blank
        Me.Undo
        DoCmd.GoToRecord , , acPrevious
        Response = acDataErrContinue

    Case Else
        'MsgBox "Unhandled Form Error " & DataErr & " " & Err.Number & " " & Err.Description
    End Select
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "Sizes:Form_Error"
    Resume ExitSub

End Sub

Private Sub Form_Load()
    If gcfHandleErrors Then On Error GoTo ErrLog
    cmbCR1.BackColor = ColorLockedForAll
    cmbCR2.BackColor = ColorLockedForAll
    cmbCR3.BackColor = ColorLockedForAll
    cmbCR4.BackColor = ColorLockedForAll
    gbRunMailChanges = False
    'txtEmail = kmdsgetsettings
    txtEmail = GetKMDSSettings("KMDSMail EODeMail", "info@kmds.us")
    GetKMDSSettings "StockNumberOnReceipt", "False"
    GetKMDSSettings "UseOLCCScanItemsForLookup", "False"
    GetKMDSSettings "Fintech.Active", "False"
    GetKMDSSettings "Fintech.Division-ID", "0"
    GetKMDSSettings "Fintech.URL", "ftp.kmds.us"
    GetKMDSSettings "Fintech.UserName", "Agents"
    GetKMDSSettings "Fintech.Password", "Ag3nt$"
    GetKMDSSettings "FintechLiquorOnly", "True"
    GetKMDSSettings "Fintech.Division-IDRelated", "0"
    GetKMDSSettings "GetItemICEatEOD", "True"
    chkUseWinSCPForOLASTransmits = GetKMDSSettings("UseWinSCPForOLASTransmits", "False")
    chkKMDSFTPActive = GetKMDSSettings("KMDSFTPActive", "False")

    GetKMDSSettings "UpdateUnitsPerCaseFromICE", "True"
    GetKMDSSettings "RunTaskKillAtZOut", "True"
    GetKMDSSettings "PriceToolVisible", "False"

ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "Form_Load"
    Resume ExitSub
End Sub

Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)
End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If gcfHandleErrors Then On Error GoTo ErrLog
    glnColorInfo = ColorInfo
    glnColorDropDown = ColorDropDown
    glnColorViewOnly = ColorViewOnly
    glnColorTotals = ColorTotals
    glnColorLockedForAll = ColorLockedForAll
    LoadDefaults
    If gbRunMailChanges Then MailChanges False
ExitSub:
    Exit Sub
ErrLog:
    Resume ExitSub
End Sub

Private Sub Label306_DblClick(Cancel As Integer)
    MsgBox "append"
    SetOptions False
    DoCmd.OpenQuery "qryKMDSLinksAppend"
    SetOptions True
    KMDSLinks.Requery

End Sub

Private Sub NetworkDataBaseConnectString_DblClick(Cancel As Integer)
    NetworkDataBaseConnectString = "Data Source=KMDSSQL7;User ID=sa;Password=;"
End Sub
Private Sub cmdUpdateMovement_Click()
    CashsysInsLoc.SetFocus
    cmdUpdateMovement.Enabled = False
    SQLPassThru "sp_UpdateItemMovement"

    'UpdateMovement (False)
    MsgBox "Finished"
    cmdUpdateMovement.Enabled = True

End Sub
Private Sub Command263_Click()
If gcfHandleErrors Then On Error GoTo Err_Command263_Click

    Dim stDocName As String
    Dim MyForm As Form

    stDocName = "Audit"
    Set MyForm = Screen.ActiveForm
    DoCmd.SelectObject acTable, stDocName, True
    DoCmd.PrintOut
    DoCmd.SelectObject acForm, MyForm.Name, False

Exit_Command263_Click:
    Exit Sub

Err_Command263_Click:
    MsgBox Err.Description
    Resume Exit_Command263_Click

End Sub
Private Sub cmdCRReportsAppend_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdCRReportsAppend_Click

    Dim stDocName As String

    stDocName = "qrydbo_CRReportsAppend"
    DoCmd.OpenQuery stDocName, acNormal, acEdit
    stDocName = "qrydbo_CRReportsUpdate"
    DoCmd.OpenQuery stDocName, acNormal, acEdit
    MsgBox "refreshed!"
Exit_cmdCRReportsAppend_Click:
    Exit Sub

Err_cmdCRReportsAppend_Click:
    MsgBox Err.Description
    Resume Exit_cmdCRReportsAppend_Click

End Sub
Private Sub cmdCRReportsUpdate_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdCRReportsUpdate_Click

    Dim stDocName As String
    SetOptions False
    stDocName = "qrydbo_CRReportsAppend"
    DoCmd.OpenQuery stDocName, acNormal, acEdit
    stDocName = "qrydbo_CRReportsUpdate"
    DoCmd.OpenQuery stDocName, acNormal, acEdit
    MsgBox "Refreshed!"
Exit_cmdCRReportsUpdate_Click:
    SetOptions True
    Exit Sub

Err_cmdCRReportsUpdate_Click:
    MsgBox Err.Description
    Resume Exit_cmdCRReportsUpdate_Click

End Sub
Private Sub cmdTranTypes_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdTranTypes_Click

    Dim stDocName As String
    SetOptions False
    stDocName = "qryTranTypes_Step_1_Append"
    DoCmd.OpenQuery stDocName, acNormal, acEdit

    stDocName = "qryTranTypes_Step_2_Update"
    DoCmd.OpenQuery stDocName, acNormal, acEdit
    MsgBox "Refreshed!"
Exit_cmdTranTypes_Click:
    SetOptions False
    Exit Sub

Err_cmdTranTypes_Click:
    MsgBox Err.Description
    Resume Exit_cmdTranTypes_Click

End Sub
Private Sub cmdSetConnectInvoline_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdSetConnectInvoline_Click
    Me.Dirty = False

    SetConnectForInvolineX

Exit_cmdSetConnectInvoline_Click:
    Exit Sub

Err_cmdSetConnectInvoline_Click:
    MsgBox Err.Description
    Resume Exit_cmdSetConnectInvoline_Click

End Sub




Private Sub txtEmail_GotFocus()
    txtEmail.Text = GetKMDSSettings("KMDSMail EODeMail", "info@kmds.us")

End Sub
'------------------------------------------------------------
' cmdInstallFonts_Click
'
'------------------------------------------------------------
Private Sub cmdInstallFonts_Click()
On Error GoTo cmdInstallFonts_Click_Err
Dim filTemp As New FileSystemObject

'    filTemp.CopyFile "k:\versions\arialn.ttf", "C:\windows\fonts"

   ShellView ("K:\Versions\InstallFonts.vbs")



cmdInstallFonts_Click_Exit:
    Exit Sub

cmdInstallFonts_Click_Err:
    MsgBox Error$
    Resume cmdInstallFonts_Click_Exit

End Sub


'------------------------------------------------------------
' cmdFixKMDS_Click
'
'------------------------------------------------------------
Private Sub cmdFixKMDS_Click()
On Error GoTo cmdFixKMDS_Click_Err

    Debug.Print ("K:\Versions\FixKMDS.bat " & gstrCashsysInsLoc)
    WaitForExecCmd ("K:\Versions\FixKMDS.bat " & gstrCashsysInsLoc)


cmdFixKMDS_Click_Exit:
    Exit Sub

cmdFixKMDS_Click_Err:
    MsgBox Error$
    Resume cmdFixKMDS_Click_Exit

End Sub
