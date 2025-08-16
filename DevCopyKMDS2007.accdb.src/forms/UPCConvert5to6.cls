Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdqryUPCConvert5to6AppendFromUPCItemsWith6DigitVendors_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdqryUPCConvert5to6AppendFromUPCItemsWith6DigitVendors_Click

    Dim stDocName As String

'    DoCmd.RunSQL "qryUPCConvert5to6AppendFromUPCItemsWith6DigitVendors"
    stDocName = "qryUPCConvert5to6AppendFromUPCItemsWith6DigitVendors"
    CurrentDb.Execute stDocName

'    DoCmd.RunSQL "qryUPCAppend6DigitVendors"
    stDocName = "qryUPCAppend6DigitVendors"
    CurrentDb.Execute stDocName


Exit_cmdqryUPCConvert5to6AppendFromUPCIt:
    Exit Sub

Err_cmdqryUPCConvert5to6AppendFromUPCItemsWith6DigitVendors_Click:
    MsgBox Err.Description
    Resume Exit_cmdqryUPCConvert5to6AppendFromUPCIt

End Sub
