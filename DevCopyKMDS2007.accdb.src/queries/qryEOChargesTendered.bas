Operation =1
Option =0
Where ="(((InvoLine5.Code)=3))"
Begin InputTables
    Name ="Cashier"
    Name ="InvoLine5"
    Name ="Tender"
End
Begin OutputColumns
    Alias ="CR"
    Expression ="InvoLine5.CashSysID"
    Expression ="InvoLine5.CashierID"
    Alias ="CashierName"
    Expression ="Cashier.Name"
    Alias ="Invoice"
    Expression ="InvoLine5.InvoHeadID"
    Expression ="InvoLine5.SalesDept"
    Alias ="Time"
    Expression ="Format(InvoLine5.SetupDate,\"Short Time\")"
    Alias ="TranCount"
    Expression ="(0-[Quantity])"
    Alias ="Amount"
    Expression ="0-[Price]"
    Expression ="InvoLine5.SalesGroup"
End
Begin Joins
    LeftTable ="InvoLine5"
    RightTable ="Tender"
    Expression ="InvoLine5.Code = Tender.ID"
    Flag =2
    LeftTable ="Cashier"
    RightTable ="InvoLine5"
    Expression ="Cashier.ID = InvoLine5.CashierID"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLine5.CashSysID"
    Flag =0
    Expression ="InvoLine5.CashierID"
    Flag =0
    Expression ="InvoLine5.InvoHeadID"
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
        dbText "Name" ="InvoLine5.CashierID"
        dbInteger "ColumnWidth" ="1065"
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
    Bottom =217
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =140
        Top =93
        Right =236
        Bottom =207
        Top =0
        Name ="Cashier"
        Name =""
    End
    Begin
        Left =266
        Top =-6
        Right =362
        Bottom =183
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =419
        Top =-11
        Right =515
        Bottom =103
        Top =0
        Name ="Tender"
        Name =""
    End
End
