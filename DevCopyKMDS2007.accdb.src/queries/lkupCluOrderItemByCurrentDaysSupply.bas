Operation =1
Option =0
Begin InputTables
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Expression ="vwItemsWithCurrentPricing.ID"
    Alias ="CurDaysSupply"
    Expression ="IIf([AverageDailySales]=0,0,[OnShelf]/[AverageDailySales])"
    Expression ="vwItemsWithCurrentPricing.OnShelf"
    Expression ="vwItemsWithCurrentPricing.AverageDailySales"
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
        dbText "Name" ="vwItemsWithCurrentPricing.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurDaysSupply"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.OnShelf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.AverageDailySales"
        dbInteger "ColumnWidth" ="2205"
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
    Bottom =232
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =266
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
End
