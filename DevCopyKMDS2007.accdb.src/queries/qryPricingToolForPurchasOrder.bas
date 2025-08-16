Operation =1
Option =0
Where ="(((vwItemsWithCurrentPricing.Price)<>[PricingWorkTable].[Price]))"
Begin InputTables
    Name ="vwItemsWithCurrentPricing"
    Name ="PricingWorkTable"
End
Begin OutputColumns
    Expression ="PricingWorkTable.ID"
    Expression ="vwItemsWithCurrentPricing.Cost"
    Expression ="vwItemsWithCurrentPricing.Price"
    Expression ="vwItemsWithCurrentPricing.UnitsPerCase"
End
Begin Joins
    LeftTable ="vwItemsWithCurrentPricing"
    RightTable ="PricingWorkTable"
    Expression ="vwItemsWithCurrentPricing.ID=PricingWorkTable.ID"
    Flag =1
End
Begin OrderBy
    Expression ="PricingWorkTable.ID"
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
        dbText "Name" ="vwItemsWithCurrentPricing.Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PricingWorkTable.ID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-335
    Top =70
    Right =703
    Bottom =726
    Left =-1
    Top =-1
    Right =1006
    Bottom =322
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
    Begin
        Left =309
        Top =62
        Right =453
        Bottom =297
        Top =0
        Name ="PricingWorkTable"
        Name =""
    End
End
