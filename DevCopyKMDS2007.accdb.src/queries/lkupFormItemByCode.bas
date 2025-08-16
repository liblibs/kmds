Operation =1
Option =0
Begin InputTables
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Expression ="vwItemsWithCurrentPricing.ID"
    Expression ="vwItemsWithCurrentPricing.Description"
    Expression ="vwItemsWithCurrentPricing.PkgSize"
    Expression ="vwItemsWithCurrentPricing.OnShelf"
    Expression ="vwItemsWithCurrentPricing.CurPrice"
    Expression ="vwItemsWithCurrentPricing.Cost"
    Alias ="GP"
    Expression ="Format(([CurPrice]-[cost])/[CurPrice],\"Percent\")"
End
Begin OrderBy
    Expression ="vwItemsWithCurrentPricing.ID"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="GP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.PkgSize"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.OnShelf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.CurPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Cost"
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
    Right =1193
    Bottom =227
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =294
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
End
