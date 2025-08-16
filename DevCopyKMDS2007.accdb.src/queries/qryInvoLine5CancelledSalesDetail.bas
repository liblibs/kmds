Operation =1
Option =0
Where ="(((InvoLine5.ID)<=[maxid]))"
Begin InputTables
    Name ="qryInvoLine5CancelledSales"
    Name ="Cashier"
    Name ="InvoLine5"
    Name ="Item"
End
Begin OutputColumns
    Expression ="qryInvoLine5CancelledSales.CashierID"
    Alias ="CashierName"
    Expression ="Cashier.Name"
    Expression ="qryInvoLine5CancelledSales.SetupDate"
    Expression ="InvoLine5.CashSysID"
    Expression ="InvoLine5.InvoHeadID"
    Expression ="InvoLine5.Code"
    Expression ="Item.Description"
    Expression ="InvoLine5.InputCode"
    Expression ="InvoLine5.Quantity"
    Expression ="InvoLine5.Price"
    Expression ="InvoLine5.ID"
End
Begin Joins
    LeftTable ="qryInvoLine5CancelledSales"
    RightTable ="Cashier"
    Expression ="qryInvoLine5CancelledSales.CashierID = Cashier.ID"
    Flag =1
    LeftTable ="qryInvoLine5CancelledSales"
    RightTable ="InvoLine5"
    Expression ="qryInvoLine5CancelledSales.InvoHeadID = InvoLine5.InvoHeadID"
    Flag =1
    LeftTable ="qryInvoLine5CancelledSales"
    RightTable ="InvoLine5"
    Expression ="qryInvoLine5CancelledSales.CashSysID = InvoLine5.CashSysID"
    Flag =1
    LeftTable ="qryInvoLine5CancelledSales"
    RightTable ="InvoLine5"
    Expression ="qryInvoLine5CancelledSales.SetupDate = InvoLine5.SetupDate"
    Flag =1
    LeftTable ="InvoLine5"
    RightTable ="Item"
    Expression ="InvoLine5.Code = Item.ID"
    Flag =1
End
Begin OrderBy
    Expression ="qryInvoLine5CancelledSales.CashierID"
    Flag =0
    Expression ="qryInvoLine5CancelledSales.SetupDate"
    Flag =0
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
        dbText "Name" ="CashierName"
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
    Bottom =220
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =19
        Top =27
        Right =115
        Bottom =237
        Top =0
        Name ="qryInvoLine5CancelledSales"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =94
        Top =0
        Name ="Cashier"
        Name =""
    End
    Begin
        Left =218
        Top =100
        Right =314
        Bottom =214
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =407
        Top =60
        Right =503
        Bottom =174
        Top =0
        Name ="Item"
        Name =""
    End
End
