Operation =1
Option =0
Begin InputTables
    Name ="ScanBatchLineCounts"
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Expression ="ScanBatchLineCounts.ItemID"
    Alias ="SumOfCountQuantity"
    Expression ="Sum(ScanBatchLineCounts.CountQuantity)"
    Alias ="FirstOfCost"
    Expression ="First(vwItemsWithCurrentPricing.Cost)"
    Alias ="FirstOfPrice"
    Expression ="First(vwItemsWithCurrentPricing.Price)"
    Expression ="vwItemsWithCurrentPricing.UnitsPerCase"
End
Begin Joins
    LeftTable ="ScanBatchLineCounts"
    RightTable ="vwItemsWithCurrentPricing"
    Expression ="ScanBatchLineCounts.ItemID = vwItemsWithCurrentPricing.ID"
    Flag =1
End
Begin OrderBy
    Expression ="ScanBatchLineCounts.ItemID"
    Flag =0
End
Begin Groups
    Expression ="ScanBatchLineCounts.ItemID"
    GroupLevel =0
    Expression ="vwItemsWithCurrentPricing.UnitsPerCase"
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
        dbText "Name" ="ScanBatchLineCounts.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfCountQuantity"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2775"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="FirstOfCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.[UnitsPerCase]"
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
    Bottom =225
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =215
        Bottom =244
        Top =0
        Name ="ScanBatchLineCounts"
        Name =""
    End
    Begin
        Left =292
        Top =1
        Right =519
        Bottom =258
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
End
