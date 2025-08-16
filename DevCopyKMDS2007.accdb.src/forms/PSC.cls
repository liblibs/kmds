Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub cboSalesDepartmentIDBegin_AfterUpdate()
    cboSalesDepartmentIDEnd = cboSalesDepartmentIDBegin

End Sub

Private Sub cmdGet_Click()
    MsgBox ("Make sure the PSC Device is in the cradle")
    DoCmd.Close acForm, Me.Name
    CreatePSCImportScannedData

End Sub

Private Sub cmdRestoreStep2_Click()
    If gcfHandleErrors Then On Error GoTo ErrLog
    SetOptions False
    DoCmd.OpenQuery "qryPSCInput_Step1_DeleteRecords"

    DoCmd.TransferText , "PSCINPUT Import Specification", "PSCInput", "c:\pdtFiles\PSCInput.txt", False

    DoCmd.OpenQuery "qryPSCInput_Step2_UpdateWithItemInfo"
    SetOptions True
    ArchivePSCDownload
    DoCmd.Close
    DoCmd.OpenForm "PSCInputEditScans"

ExitSub:
    Exit Sub

ErrLog:
    MsgBox Err.Number & " " & Err.Description & "cmdRestoreStep2_Click "
    Resume ExitSub
End Sub

Private Sub cmdSend_Click()
    cmdClose.SetFocus
    cmdSend.Visible = False
    Me.Repaint
    CreatePSCLookupFiles
    DoCmd.Close acForm, Me.Name

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
Private Sub cmdRestore_Click()

If gcfHandleErrors Then On Error GoTo Err_cmdRestore_Click
    MsgBox "The input file must be copied to the c:\pdtfiles "
    MsgBox "and copied to a file named   PSCInput.txt"

    ShellView gstrReportsDir & "PSC"



ExitSub:
    Exit Sub

cmdRestore_Click:
    EMsgBox Me, "cmdRestore_Click"
    Resume ExitSub



Exit_cmdRestore_Click:
    Exit Sub

Err_cmdRestore_Click:
    MsgBox Err.Description
    Resume Exit_cmdRestore_Click

End Sub

Private Sub Command4_Click()





End Sub

Private Sub cmdQuery_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdQuery_Click

    Dim stDocName As String

    stDocName = "qryPSCDownloadShowMe"
    DoCmd.OpenQuery stDocName, acNormal, acEdit

Exit_cmdQuery_Click:
    Exit Sub

Err_cmdQuery_Click:
    MsgBox Err.Description
    Resume Exit_cmdQuery_Click

End Sub
