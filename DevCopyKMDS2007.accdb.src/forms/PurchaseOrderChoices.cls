Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdExport_Click()
If gcfHandleErrors Then On Error GoTo Err_Command28_Click

    Dim stDocName As String

    stDocName = POReportForm '"01 Inventory Value By Code"
    DoCmd.OutputTo acReport, stDocName

Exit_Command28_Click:
    Exit Sub

Err_Command28_Click:
    MsgBox Err.Description
    Resume Exit_Command28_Click

End Sub

Private Sub Command11_Click()
    Dim stDocName As String
    Dim strOpenArgs As String

    If gcfHandleErrors Then On Error GoTo ErrLog
    If Me.Dirty Then Me.Dirty = False
    stDocName = POReportForm
    If InStr(stDocName, "(NoPrice)") > 0 Then
        stDocName = Replace(stDocName, "(NoPrice)", "")
        strOpenArgs = "NoPrice"
    End If

    ArchiveReport stDocName, txtName & " " & txtPurchaseOrderID

    Select Case optChoices.Value
    Case 1 ' print
        DoCmd.OpenReport stDocName, acViewNormal, , , , strOpenArgs
    Case 2 'post
    'invoice print
        'stDocName = "PurchaseOrder"
        If MsgBox("This will APPLY the purchase order to the ONHANDS!" & vbCrLf & vbCrLf & "Are you sure?", vbYesNo) = vbNo Then
            Exit Sub
        End If

        'DoCmd.OpenReport stDocName, acViewPreview, , , , strOpenArgs
        'DoCmd.Close acReport, stDocName
        ArchiveReportDocument stDocName, "PurchaseOrder", txtName, txtPurchaseOrderID
        PostPO CLng(txtPurchaseOrderID), txtName
    Case 3 'print & post
    'invoice print
        If MsgBox("This will APPLY the purchase order to the ONHANDS!" & vbCrLf & vbCrLf & "Are you sure?", vbYesNo) = vbNo Then
            Exit Sub
        End If
        DoCmd.OpenReport stDocName, acViewNormal, , , , strOpenArgs
        ArchiveReportDocument stDocName, "PurchaseOrder", txtName, txtPurchaseOrderID
        PostPO CLng(txtPurchaseOrderID), txtName


    Case Else
        MsgBox "Unhandled case : optchoices"
    End Select
    DoCmd.Close
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Command11_Click"
    Resume ExitSub
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
Private Sub cmdPrintPreview_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdPrintPreview_Click

    Dim stDocName As String
    Dim strOpenArgs As String
    If Me.Dirty Then Me.Dirty = False

    stDocName = POReportForm '"PurchaseOrderLiquor"
    If InStr(stDocName, "(NoPrice)") > 0 Then
        stDocName = Replace(stDocName, "(NoPrice)", "")
        strOpenArgs = "NoPrice"
    End If




    DoCmd.OpenReport stDocName, acPreview, , , , strOpenArgs

Exit_cmdPrintPreview_Click:
    Exit Sub

Err_cmdPrintPreview_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrintPreview_Click

End Sub


Private Sub Command23_Click()

End Sub

Private Sub Form_Open(Cancel As Integer)
    txtPurchaseOrderID = Me.OpenArgs
    If IsNull(txtPurchaseOrderID) Then
        MsgBox "Unknown Purchase Order Number, try printing again!"
        Cancel = True
    End If
End Sub

Private Sub Command28_Click()
If gcfHandleErrors Then On Error GoTo Err_Command28_Click

    Dim stDocName As String

    stDocName = "01 Inventory Value By Code"
    DoCmd.OutputTo acReport, stDocName

Exit_Command28_Click:
    Exit Sub

Err_Command28_Click:
    MsgBox Err.Description
    Resume Exit_Command28_Click

End Sub
