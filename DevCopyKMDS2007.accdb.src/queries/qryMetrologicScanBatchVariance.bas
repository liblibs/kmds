Operation =1
Option =0
Begin InputTables
    Name ="qryMetrologicScanBatchTotalByItemID"
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Expression ="vwItemsWithCurrentPricing.SalesDptDescr"
    Expression ="vwItemsWithCurrentPricing.CategoryDescr"
    Expression ="vwItemsWithCurrentPricing.ID"
    Expression ="vwItemsWithCurrentPricing.Description"
    Expression ="vwItemsWithCurrentPricing.PkgSize"
    Alias ="Count"
    Expression ="qryMetrologicScanBatchTotalByItemID.SumOfQuantity"
    Expression ="vwItemsWithCurrentPricing.OnShelf"
    Alias ="Variance"
    Expression ="[SumOfQuantity]-[OnShelf]"
    Alias ="Variance Dollars"
    Expression ="[curPrice]*([SumOfQuantity]-[OnShelf])"
    Alias ="CashierID"
    Expression ="qryMetrologicScanBatchTotalByItemID.FirstOfCashierID"
    Expression ="qryMetrologicScanBatchTotalByItemID.InvoLineID"
End
Begin Joins
    LeftTable ="qryMetrologicScanBatchTotalByItemID"
    RightTable ="vwItemsWithCurrentPricing"
    Expression ="qryMetrologicScanBatchTotalByItemID.ItemID=vwItemsWithCurrentPricing.ID"
    Flag =1
End
Begin OrderBy
    Expression ="vwItemsWithCurrentPricing.SalesDptDescr"
    Flag =0
    Expression ="vwItemsWithCurrentPricing.CategoryDescr"
    Flag =0
    Expression ="vwItemsWithCurrentPricing.ID"
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
        dbText "Name" ="vwItemsWithCurrentPricing.ID"
        dbInteger "ColumnWidth" ="615"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Description"
        dbInteger "ColumnWidth" ="3075"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.OnShelf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="variance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Variance Dollars"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.PkgSize"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashierID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryMetrologicScanBatchTotalByItemID.InvoLineID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.SalesDptDescr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.CategoryDescr"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-231
    Top =113
    Right =1202
    Bottom =832
    Left =-1
    Top =-1
    Right =1401
    Bottom =351
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =254
        Bottom =353
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
    Begin
        Left =507
        Top =38
        Right =652
        Bottom =206
        Top =0
        Name ="qryMetrologicScanBatchTotalByItemID"
        Name =""
    End
End
