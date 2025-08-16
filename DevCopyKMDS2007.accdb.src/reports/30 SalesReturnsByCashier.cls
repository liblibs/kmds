Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Close()
    SQLPassThru "spDeleteInvoLine5"



End Sub

Private Sub Report_Open(Cancel As Integer)





    SQLPassThru "spDeleteInvoLine5"
    SQLPassThru "sp_SalesReturnsByCashierAppend " & CDbl(CDate(Forms!PrintInventoryReport!txtDateRangeFrom)) - 2 & " , " & CDbl(CDate(Forms!PrintInventoryReport!txtDateRangeTo)) - 2




End Sub
