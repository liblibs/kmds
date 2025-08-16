Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdBackup_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdBackup_Click

    ShellView gstrReportsDir



Exit_cmdBackup_Click:
    Exit Sub

Err_cmdBackup_Click:
    EMsgBox Me, "cmdBackup_Click"
    Resume Exit_cmdBackup_Click

End Sub

Private Sub cmdCreateDailySalesFile_Click()

Dim stDocName As String

If gcfHandleErrors Then On Error GoTo ErrLog

    txtDateRangeFrom = Format(dtpDate, "mm/dd/yyyy 00:00:00")
    txtDateRangeTo = Format(dtpDate, "mm/dd/yyyy 23:59:59")
    txtDateRangeFrom = Replace(txtDateRangeFrom, "-", "/")
    txtDateRangeTo = Replace(txtDateRangeTo, "-", "/")
    SQLPassThru "spDeleteInvoLine5"
    SQLPassThru "sp_SalesReturnsByCashierAppend '" & txtDateRangeFrom & "' , '" & txtDateRangeTo & "'"

    OLCCDailySalesFileReCreate CLng(glnAgentNumber), False, dtpDate
    OLCCDailySalesFileReCreate CLng(glnAgentNumber), True, dtpDate

    SetOptions False
    stDocName = "qryRecreateTransmitFileOnhands_Step1_Set_AvgTurns_to_Onhand"
    DoCmd.OpenQuery stDocName

    stDocName = "qryRecreateTransmitFileOnhands_Step2a_Delete_Movement"
    DoCmd.OpenQuery stDocName

    stDocName = "qryRecreateTransmitFileOnhands_Step2b_Append_Movement"
    DoCmd.OpenQuery stDocName

    stDocName = "qryRecreateTransmitFileOnhands_Step3_Update_AvgTurns"
    DoCmd.OpenQuery stDocName

    OLCCDailyInventFileRecreate CLng(glnAgentNumber), False, dtpDate
    OLCCDailyInventFileRecreate CLng(glnAgentNumber), True, dtpDate


ExitSub:
    SetOptions True
    Exit Sub

ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "cmdCreateDailySalesFile_Click"
    Resume ExitSub


End Sub

Private Sub cmdEndofDay_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdBackup_Click
Dim datTemp As Date
Dim strPath As String

    datTemp = dtpDate
    strPath = gstrReportsDir & "EndOfDay\"
    strPath = strPath & Year(datTemp) & "\"
    strPath = strPath & Format(Month(datTemp), "00") & "\"
    strPath = strPath & Format(Day(datTemp), "00")


    ShellView strPath




Exit_cmdBackup_Click:
    Exit Sub

Err_cmdBackup_Click:
    EMsgBox Me, "cmdBackup_Click"
    Resume Exit_cmdBackup_Click

End Sub


Private Sub cmdExploreJournals_Click()
    Dim strTarget As String
    FetchJournals
    strTarget = gstrMailDir & "Alljournals\"
    ShellView strTarget

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

Private Sub Combo3_AfterUpdate()
    cmdExploreJournals.Visible = True
End Sub

Private Sub Combo3_BeforeUpdate(Cancel As Integer)

End Sub

Private Sub cmdLastNiteFiles_Click()
    Shell "K:LastNiteFiles.Bat"

End Sub

Private Sub cmdRentrak_Click()
Dim datTemp As Date

    If MsgBox("Recreate Rentrak transmit file for " & Format(dtpRentrak, "mm/dd/yyyy"), vbYesNo) = vbYes Then

        RentrakRetransmit Format(dtpRentrak, "mm/dd/yyyy")
    End If


End Sub

Private Sub Command21_Click()

End Sub

Private Sub dtpDate_Change()
    cmdCreateDailySalesFile.Caption = "Re-Create Daily Sales File for " & Format(dtpDate, "mm/dd/yyyy")
End Sub

Private Sub Form_Load()
    dtpDate.Value = Now - 1
    cmdCreateDailySalesFile.Caption = "Re-Create Daily Sales File for " & Format(dtpDate, "mm/dd/yyyy")
    dtpRentrak.Visible = gbRentrakInstalled
    cmdRentrak.Visible = gbRentrakInstalled
    cmdTransferCartons.Visible = gbTransferCartonsToPacks

End Sub

Private Sub Form_Open(Cancel As Integer)
       Cancel = FormLogin(glnCashierID, Me)
End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub

Private Sub cmdTransferCartons_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdTransferCartons_Click
    If MsgBox("This will move all of the Cartons onhands to the packs, are you sure?", vbYesNo) = vbYes Then
        TransferCartonsToPacks
    End If

Exit_cmdTransferCartons_Click:
    Exit Sub

Err_cmdTransferCartons_Click:
    MsgBox Err.Description
    Resume Exit_cmdTransferCartons_Click

End Sub
Private Sub cmdOLCCDailyTransmit_Click()
    If VerifyCurrentKMDSCustomer Then
        OLCCDailyTransmit
    Else
        MsgBox "Please contact KMDS Support to update your KMDS Subscription", vbCritical, "Support expired"
    End If
End Sub
Private Sub cmdRetransmitOLCCFiles_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdRetransmitOLCCFiles_Click

    Dim stAppName As String

    stAppName = "\KMDS\OLCCRetransmit.exe"
    Call Shell(stAppName, 1)
    cmdfrmClose_Click

Exit_cmdRetransmitOLCCFiles_Click:
    Exit Sub

Err_cmdRetransmitOLCCFiles_Click:
    MsgBox Err.Description
    Resume Exit_cmdRetransmitOLCCFiles_Click

End Sub
Private Sub cmdEmailRecap_Click()
    EODTextMessage
Exit_cmdEmailRecap_Click:
    Exit Sub

Err_cmdEmailRecap_Click:
    MsgBox Err.Description
    Resume Exit_cmdEmailRecap_Click

End Sub
