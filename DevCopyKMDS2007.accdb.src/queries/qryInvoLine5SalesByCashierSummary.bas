Operation =1
Option =0
Begin InputTables
    Name ="Cashier"
    Name ="qryInvoLine5SalesByCashierByInvoice"
End
Begin OutputColumns
    Expression ="Cashier.ID"
    Expression ="Cashier.Name"
    Alias ="TotalTransactions"
    Expression ="Sum(qryInvoLine5SalesByCashierByInvoice.Transactions)"
    Alias ="TotalSales"
    Expression ="Sum(qryInvoLine5SalesByCashierByInvoice.Sales)"
    Alias ="AvgSale"
    Expression ="[TotalSales]/[TotalTransactions]"
    Alias ="TotalNoSales"
    Expression ="Sum(qryInvoLine5SalesByCashierByInvoice.NoSales)"
    Alias ="TotalCancelledSales"
    Expression ="Sum(qryInvoLine5SalesByCashierByInvoice.CancelledSales)"
    Alias ="TotalIDScanCount"
    Expression ="Sum(qryInvoLine5SalesByCashierByInvoice.IDScanCount)"
End
Begin Joins
    LeftTable ="Cashier"
    RightTable ="qryInvoLine5SalesByCashierByInvoice"
    Expression ="Cashier.ID=qryInvoLine5SalesByCashierByInvoice.CashierID"
    Flag =1
End
Begin Groups
    Expression ="Cashier.ID"
    GroupLevel =0
    Expression ="Cashier.Name"
    GroupLevel =0
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
        dbText "Name" ="Cashier.ID"
        dbInteger "ColumnWidth" ="390"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cashier.Name"
        dbInteger "ColumnWidth" ="840"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgSale"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalSales"
        dbInteger "ColumnWidth" ="1170"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalTransactions"
        dbInteger "ColumnWidth" ="1755"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalNoSales"
        dbInteger "ColumnWidth" ="1335"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalCancelledSales"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalIDScanCount"
        dbInteger "ColumnWidth" ="1740"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =150
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="Cashier"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =314
        Bottom =135
        Top =0
        Name ="qryInvoLine5SalesByCashierByInvoice"
        Name =""
    End
End
