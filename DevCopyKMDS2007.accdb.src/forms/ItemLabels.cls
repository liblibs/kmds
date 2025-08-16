Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdBrotherPriceTags_Click()
    BrotherLabels "Item"

End Sub

Private Sub cmdBrotherShelfTags_Click()
    BrotherLabels "Shelf"

End Sub

Private Sub cmdfrmClose_Click()
If gcfHandleErrors Then On Error GoTo ErrLog
    'If MsgBox("Your work will not be saved..." & vbcrlf & "Stop this program?", vbYesNo) = vbYes Then
        DoCmd.Close
    'End If
ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmClose_Click"
    Resume ExitSub
End Sub


Private Sub cmdItemLabels5267_Click()

If gcfHandleErrors Then On Error GoTo ErrLog

    Dim stDocName As String

    stDocName = "LabelItemLabels5267"
    DoCmd.OpenReport stDocName, acViewNormal

ExitLog:
    Exit Sub

ErrLog:
    Select Case Err.Number
    Case Else
        EMsgBox Me, "cmdLabelsItemsWithOUTBarCodes_Click"
    End Select
    Resume ExitLog
End Sub

Private Sub cmdItemLabels5267WithPrice_Click()

If gcfHandleErrors Then On Error GoTo ErrLog

    Dim stDocName As String

    stDocName = "LabelItemLabels5267WithPrice"
    DoCmd.OpenReport stDocName, acViewNormal

ExitLog:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdLabelsItemsWithOUTBarCodes_Click"
    Resume ExitLog
End Sub

Private Sub cmdLabelsItemsWithOUTBarCodes_Click()
    If gcfHandleErrors Then On Error GoTo ErrLog

    Dim stDocName As String

    stDocName = "LabelShelfTagsWithOutBarCodes8160"
    DoCmd.OpenReport stDocName, acViewNormal

ExitLog:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdLabelsItemsWithOUTBarCodes_Click"
    Resume ExitLog

End Sub


Private Sub cmdLabelsItemsWithBarCodes_Click()

If gcfHandleErrors Then On Error GoTo ErrLog

    Dim stDocName As String

    stDocName = "LabelShelfTagsWithBarcodes8160"
    DoCmd.OpenReport stDocName, acViewNormal

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdLabelsItemsWithBarCodes_Click"
    Resume ExitSub

End Sub

Private Sub cmdPreview_5260_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    Dim stDocName As String

    stDocName = "LabelShelfTagsWithOutBarCodes8160"
    DoCmd.OpenReport stDocName, acViewPreview


ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdLabelsItemsWithBarCodes_Click"
    Resume ExitSub
End Sub

Private Sub cmdPreview_5260_WithBarcodes_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    Dim stDocName As String

    stDocName = "LabelShelfTagsWithBarcodes8160"
    DoCmd.OpenReport stDocName, acPreview

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdLabelsItemsWithBarCodes_Click"
    Resume ExitSub
End Sub

Private Sub cmdRefresh_Click()
    Dim nAbsPos As Long
    If gcfHandleErrors Then On Error GoTo ErrLog

    nAbsPos = Me.ItemLabelsSubWithout.Form.Recordset.AbsolutePosition
   'MsgBox nAbsPos

    Me.ItemLabelsSubWithout.Form.Filter = "(( ItemLabels.Labels=0))"
    Me.ItemLabelsSubWithout.Form.FilterOn = False
    Me.ItemLabelsSubWithout.Form.FilterOn = True

    If gcfHandleErrors Then On Error GoTo ErrMove
    Me.ItemLabelsSubWithout.Form.Recordset.AbsolutePosition = nAbsPos

    Me.ItemLabelsSubWith.Form.Filter = "((Not ItemLabels.Labels=0))"

    Me.ItemLabelsSubWith.Form.FilterOn = False
    Me.ItemLabelsSubWith.Form.FilterOn = True
    Me.ItemLabelsSubWith.Form.OrderBy = "ItemLabels.ItemD"
    Me.ItemLabelsSubWith.Form.OrderByOn = True

    ItemLabelsSubWithout.SetFocus
ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdRefresh_Click"
    Resume ExitSub

ErrMove:
    Me.ItemLabelsSubWithout.Form.Recordset.MoveLast ' .AbsolutePosition = nAbsPos
    Resume Next
End Sub



Private Sub Form_Load()
    Dim bBrother As Boolean

    SetFormSize Me.Form
    Me.ItemLabelsSubWithout.Form.Filter = "(( ItemLabels.Labels=0))"
    Me.ItemLabelsSubWithout.Form.FilterOn = False
    Me.ItemLabelsSubWithout.Form.FilterOn = True

    Me.ItemLabelsSubWith.Form.Filter = "((Not ItemLabels.Labels=0))"
    Me.ItemLabelsSubWith.Form.FilterOn = False
    Me.ItemLabelsSubWith.Form.FilterOn = True

    Me.ItemLabelsSubWith.Form.OrderBy = "ItemLabels.ItemD"
    Me.ItemLabelsSubWith.Form.OrderByOn = True

    bBrother = True
    If GetKMDSSettings("BrotherLabelPrinterInstalled", "False") = "False" Then
        bBrother = False
    End If
    cmdBrotherPriceTags.Visible = bBrother
    cmdBrotherShelfTags.Visible = bBrother
    lblBrotherLabelPrinter.Visible = bBrother
    lblBrother.Visible = bBrother


End Sub
Private Sub cmdPreview_5267_WithPrice_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdPreview_5267_WithPrice_Click

    Dim stDocName As String

    stDocName = "LabelItemLabels5267WithPrice"

    DoCmd.OpenReport stDocName, acPreview

Exit_cmdPreview_5267_WithPrice_Click:
    Exit Sub

Err_cmdPreview_5267_WithPrice_Click:
    MsgBox Err.Description
    Resume Exit_cmdPreview_5267_WithPrice_Click

End Sub
Private Sub cmdPreview_5267_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdPreview_5267_Click

    Dim stDocName As String

    stDocName = "LabelItemLabels5267"
    DoCmd.OpenReport stDocName, acPreview

Exit_cmdPreview_5267_Click:
    Exit Sub

Err_cmdPreview_5267_Click:
    MsgBox Err.Description
    Resume Exit_cmdPreview_5267_Click

End Sub
'------------------------------------------------------------
' cmdExport_Click
'
'------------------------------------------------------------
Private Sub cmdExport_Click()
Dim strLabel As String
On Error GoTo cmdExport_Click_Err

    DoCmd.OutputTo acOutputReport, "ItemLabelsForBrother", acFormatXLS, "c:\kmds\mail\out\LabelOutput.xls", False, "", 0
     MsgBox "Label spreadsheet can be found in c:\kmds\mail\out\LabelOutput.xls"
'    strLabel = GetKMDSSettings("BrotherLabel", "K:\Brother\DK205 KMDS Shelf Tag.lbx")
'    strLabel = GetKMDSSettings("BrotherLabel", "K:\Brother\DK205 KMDS Shelf Tag.lbx")
'    StartDoc (strLabel)

cmdExport_Click_Exit:
    Exit Sub

cmdExport_Click_Err:
    MsgBox Error$
    Resume cmdExport_Click_Exit

End Sub


'------------------------------------------------------------
' cmdExportExcel_Click
'
'------------------------------------------------------------
Private Sub cmdExportExcel_Click()
On Error GoTo cmdExportExcel_Click_Err

    MsgBox "Label spreadsheet can be found in c:\kmds\mail\out\LabelOutput.xls"
    DoCmd.OutputTo acOutputReport, "ItemLabelsForBrother", acFormatXLS, "c:\kmds\mail\out\LabelOutput.xls", True, ""

cmdExportExcel_Click_Exit:
    Exit Sub

cmdExportExcel_Click_Err:
    MsgBox Error$
    Resume cmdExportExcel_Click_Exit

End Sub
