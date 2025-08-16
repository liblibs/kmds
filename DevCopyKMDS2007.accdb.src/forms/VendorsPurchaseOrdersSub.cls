Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Date_DblClick(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo Err_cmdPrintPreview_Click

    Dim stDocName As String

    stDocName = "PurchaseOrderWithPriceAndCostVendorScr"
    DoCmd.OpenReport stDocName, acPreview

Exit_cmdPrintPreview_Click:
    Exit Sub

Err_cmdPrintPreview_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrintPreview_Click

End Sub

Private Sub Form_Current()
    Me.Parent("txtInvoiceHeadID") = InvoHeadID
    Me.Parent("cmdPrintPreview").Caption = "Show Order " & InvoHeadID
End Sub
