Operation =1
Option =0
Begin InputTables
    Name ="vwItemsWithCurrentPricing"
    Name ="qryMetrologicScanBatchTotalByItemID"
End
Begin OutputColumns
    Expression ="qryMetrologicScanBatchTotalByItemID.ItemID"
    Expression ="qryMetrologicScanBatchTotalByItemID.SumOfQuantity"
    Expression ="vwItemsWithCurrentPricing.Cost"
    Expression ="vwItemsWithCurrentPricing.Price"
    Expression ="vwItemsWithCurrentPricing.UnitsPerCase"
End
Begin Joins
    LeftTable ="qryMetrologicScanBatchTotalByItemID"
    RightTable ="vwItemsWithCurrentPricing"
    Expression ="qryMetrologicScanBatchTotalByItemID.ItemID = vwItemsWithCurrentPricing.ID"
    Flag =1
End
Begin OrderBy
    Expression ="qryMetrologicScanBatchTotalByItemID.ItemID"
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
        dbText "Name" ="qryMetrologicScanBatchTotalByItemID.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryMetrologicScanBatchTotalByItemID.SumOfQuantity"
        dbLong "AggregateType" ="-1"
    End
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
End
Begin
    State =0
    Left =115
    Top =1
    Right =1153
    Bottom =657
    Left =-1
    Top =-1
    Right =1006
    Bottom =373
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
        Left =488
        Top =104
        Right =632
        Bottom =248
        Top =0
        Name ="qryMetrologicScanBatchTotalByItemID"
        Name =""
    End
End
