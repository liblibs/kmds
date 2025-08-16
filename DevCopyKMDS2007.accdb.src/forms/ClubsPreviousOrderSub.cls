Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Text20_BeforeUpdate(Cancel As Integer)

End Sub

Private Sub Form_Current()
Forms!Clubs("cmdReprintInvoice").Caption = "Reprint InvoiceId #" & InvoiceNumber
Forms!Clubs("cmdReprintInvoice").Tag = InvoHeadID

End Sub

Private Sub Form_Load()
    PostDate.ColumnWidth = -2
    InvoHeadID.ColumnWidth = -2
    TotalQty.ColumnWidth = -2
'    ItemDescription.ColumnWidth = -2
'    Price.ColumnWidth = -2
'    ExtPrice.ColumnWidth = -2
'    Cost.ColumnWidth = -2
'    ExtCost.ColumnWidth = -2
'    GP.ColumnWidth = -2
'    Repack.ColumnWidth = -2
End Sub

Private Sub InvoiceNumber_Click()
PostDate_DblClick False

End Sub

Private Sub InvoiceNumber_DblClick(Cancel As Integer)
PostDate_DblClick False
End Sub

Private Sub PostDate_DblClick(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo Err_cmdReprintInvoice_Click

    Dim stDocName As String

    stDocName = "ClubInvoiceLiquorAndNonLiquorForClubScreen"
    DoCmd.OpenReport stDocName, acViewPreview


Exit_cmdReprintInvoice_Click:
    Exit Sub

Err_cmdReprintInvoice_Click:
    MsgBox Err.Description
    Resume Exit_cmdReprintInvoice_Click
End Sub
