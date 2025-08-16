Operation =1
Option =0
Where ="(((involine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom And F"
    "orms!PrintInventoryReport!txtDateRangeTo) And ((involine.Code)=5))"
Begin InputTables
    Name ="Cashier"
    Name ="involine"
    Name ="Tender"
    Name ="CheckCashing"
End
Begin OutputColumns
    Expression ="involine.CashierID"
    Alias ="CashierName"
    Expression ="Cashier.Name"
    Alias ="CR"
    Expression ="involine.CashSysID"
    Alias ="Invoice"
    Expression ="involine.InvoHeadID"
    Alias ="Time"
    Expression ="Format(involine.SetupDate,\"Short Time\")"
    Alias ="TranCount"
    Expression ="([Quantity])"
    Alias ="Amount"
    Expression ="involine.Price"
    Alias ="CustomerID"
    Expression ="involine.SalesDept"
    Alias ="Customer"
    Expression ="[FirstName] & \" \" & [LastName]"
    Expression ="CheckCashing.LastName"
    Alias ="CheckNumber"
    Expression ="involine.SalesGroup"
    Expression ="involine.SetupDate"
End
Begin Joins
    LeftTable ="involine"
    RightTable ="Tender"
    Expression ="involine.Code = Tender.ID"
    Flag =2
    LeftTable ="involine"
    RightTable ="CheckCashing"
    Expression ="involine.SalesDept = CheckCashing.ID"
    Flag =2
    LeftTable ="Cashier"
    RightTable ="involine"
    Expression ="Cashier.ID = involine.CashierID"
    Flag =1
End
Begin OrderBy
    Expression ="involine.CashierID"
    Flag =0
    Expression ="involine.Price"
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
        dbText "Name" ="TranCount"
        dbInteger "ColumnWidth" ="1140"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Amount"
        dbInteger "ColumnWidth" ="885"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashierName"
        dbInteger "ColumnWidth" ="705"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Time"
        dbInteger "ColumnWidth" ="630"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Customer"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CheckNumber"
        dbInteger "ColumnWidth" ="1560"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustomerID"
        dbInteger "ColumnWidth" ="930"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CR"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Invoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="involine.CashierID"
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
    Bottom =200
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =197
        Top =-10
        Right =293
        Bottom =104
        Top =0
        Name ="Cashier"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =135
        Bottom =203
        Top =0
        Name ="involine"
        Name =""
    End
    Begin
        Left =425
        Top =-5
        Right =521
        Bottom =109
        Top =0
        Name ="Tender"
        Name =""
    End
    Begin
        Left =420
        Top =114
        Right =603
        Bottom =223
        Top =0
        Name ="CheckCashing"
        Name =""
    End
End
