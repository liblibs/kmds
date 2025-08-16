Operation =1
Option =0
Begin InputTables
    Name ="qryInvoLine5SalesByCashierSummary"
    Name ="qrySalesReturnsByCashier"
End
Begin OutputColumns
    Expression ="qryInvoLine5SalesByCashierSummary.ID"
    Expression ="qryInvoLine5SalesByCashierSummary.Name"
    Expression ="qryInvoLine5SalesByCashierSummary.TotalTransactions"
    Expression ="qryInvoLine5SalesByCashierSummary.TotalSales"
    Expression ="qryInvoLine5SalesByCashierSummary.AvgSale"
    Expression ="qryInvoLine5SalesByCashierSummary.TotalNoSales"
    Expression ="qryInvoLine5SalesByCashierSummary.TotalCancelledSales"
    Expression ="qryInvoLine5SalesByCashierSummary.TotalIDScanCount"
    Alias ="ReturnItems"
    Expression ="IIf(IsNull([ReturnQty]),0,[ReturnQty])"
    Alias ="ReturnAmount"
    Expression ="IIf(IsNull([ReturnDollars]),0,[ReturnDollars])"
End
Begin Joins
    LeftTable ="qryInvoLine5SalesByCashierSummary"
    RightTable ="qrySalesReturnsByCashier"
    Expression ="qryInvoLine5SalesByCashierSummary.ID = qrySalesReturnsByCashier.CashierID"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="ReturnItems"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ReturnAmount"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1225
    Bottom =677
    Left =-1
    Top =-1
    Right =1209
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="qryInvoLine5SalesByCashierSummary"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =105
        Top =0
        Name ="qrySalesReturnsByCashier"
        Name =""
    End
End
