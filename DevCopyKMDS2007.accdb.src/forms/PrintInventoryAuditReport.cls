Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub cmdReport_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdReport_Click

    Dim stDocName As String
    dtpDateBegin = Format(dtpDateBegin, "mm/dd/yyyy 00:00")
    dtpDateEnd = Format(dtpDateEnd, "mm/dd/yyyy 23:59")

    stDocName = "ItemOnhandChangeAudit"
    DoCmd.OpenReport stDocName, acPreview

Exit_cmdReport_Click:
    Exit Sub

Err_cmdReport_Click:
    MsgBox Err.Description
    Resume Exit_cmdReport_Click

End Sub

Private Sub Form_Load()
    dtpDateBegin = Now
    dtpDateEnd = dtpDateBegin + 1
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

Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)
End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub
