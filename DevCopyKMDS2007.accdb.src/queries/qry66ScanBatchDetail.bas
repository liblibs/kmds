Operation =1
Option =0
Begin InputTables
    Name ="ScanBatchLineCounts"
    Name ="vwItemsWithCurrentPricing"
    Name ="Cashier"
End
Begin OutputColumns
    Expression ="ScanBatchLineCounts.ID"
    Expression ="ScanBatchLineCounts.ItemID"
    Expression ="vwItemsWithCurrentPricing.Description"
    Expression ="ScanBatchLineCounts.SetupDate"
    Expression ="ScanBatchLineCounts.CountQuantity"
    Alias ="OnShelfAtTimeOfScan"
    Expression ="ScanBatchLineCounts.OnShelf"
    Expression ="ScanBatchLineCounts.OnShelfDate"
    Expression ="vwItemsWithCurrentPricing.Price"
    Expression ="ScanBatchLineCounts.Location"
    Expression ="vwItemsWithCurrentPricing.OnShelf"
    Expression ="Cashier.Name"
End
Begin Joins
    LeftTable ="ScanBatchLineCounts"
    RightTable ="vwItemsWithCurrentPricing"
    Expression ="ScanBatchLineCounts.ItemID = vwItemsWithCurrentPricing.ID"
    Flag =1
    LeftTable ="ScanBatchLineCounts"
    RightTable ="Cashier"
    Expression ="ScanBatchLineCounts.CashierID = Cashier.ID"
    Flag =1
End
Begin OrderBy
    Expression ="ScanBatchLineCounts.ID"
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
        dbText "Name" ="ScanBatchLineCounts.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OnShelfAtTimeOfScan"
        dbInteger "ColumnWidth" ="2490"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Description"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScanBatchLineCounts.CountQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScanBatchLineCounts.SetupDate"
        dbInteger "ColumnWidth" ="2145"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScanBatchLineCounts.OnShelfDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScanBatchLineCounts.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.OnShelf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScanBatchLineCounts.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cashier.Name"
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
    Bottom =256
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =163
        Top =4
        Right =503
        Bottom =268
        Top =0
        Name ="ScanBatchLineCounts"
        Name =""
    End
    Begin
        Left =527
        Top =-2
        Right =770
        Bottom =291
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
    Begin
        Left =818
        Top =12
        Right =962
        Bottom =156
        Top =0
        Name ="Cashier"
        Name =""
    End
End
