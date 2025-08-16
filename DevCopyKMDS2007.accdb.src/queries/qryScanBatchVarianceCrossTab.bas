Operation =1
Option =0
Begin InputTables
    Name ="ScanBatchLineCounts"
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Expression ="ScanBatchLineCounts.ItemID"
    Expression ="vwItemsWithCurrentPricing.Description"
    Alias ="FirstScan"
    Expression ="First(ScanBatchLineCounts.SetupDate)"
    Alias ="LastScan"
    Expression ="Last(ScanBatchLineCounts.SetupDate)"
    Alias ="SumOfCountQuantity"
    Expression ="Sum(ScanBatchLineCounts.CountQuantity)"
    Alias ="OnShelfAtTimeOfScan"
    Expression ="Last(ScanBatchLineCounts.OnShelf)"
    Alias ="OnShelfAtTimeOfScanNOW"
    Expression ="First(vwItemsWithCurrentPricing.OnShelf)"
    Alias ="LastOfOnShelfDate"
    Expression ="Last(ScanBatchLineCounts.OnShelfDate)"
    Alias ="Loc1"
    Expression ="Sum(IIf(Left([ScanBatchLineCounts].[Location],1)='1',[CountQuantity],0))"
    Alias ="Loc2"
    Expression ="Sum(IIf(Left([ScanBatchLineCounts].[Location],1)='2',[CountQuantity],0))"
    Alias ="Loc3"
    Expression ="Sum(IIf(Left([ScanBatchLineCounts.Location],1)='3',[CountQuantity],0))"
    Alias ="Loc4"
    Expression ="Sum(IIf(Left([ScanBatchLineCounts.Location],1)='4',[CountQuantity],0))"
    Expression ="vwItemsWithCurrentPricing.Price"
    Alias ="Scans"
    Expression ="Count(1)"
    Expression ="vwItemsWithCurrentPricing.Cost"
End
Begin Joins
    LeftTable ="ScanBatchLineCounts"
    RightTable ="vwItemsWithCurrentPricing"
    Expression ="ScanBatchLineCounts.ItemID = vwItemsWithCurrentPricing.ID"
    Flag =1
End
Begin Groups
    Expression ="ScanBatchLineCounts.ItemID"
    GroupLevel =0
    Expression ="vwItemsWithCurrentPricing.Description"
    GroupLevel =0
    Expression ="vwItemsWithCurrentPricing.Price"
    GroupLevel =0
    Expression ="vwItemsWithCurrentPricing.Cost"
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
        dbText "Name" ="FirstScan"
        dbInteger "ColumnWidth" ="2145"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastScan"
        dbInteger "ColumnWidth" ="2145"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OnShelfAtTimeOfScan"
        dbInteger "ColumnWidth" ="2490"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OnShelfAtTimeOfScanNOW"
        dbInteger "ColumnWidth" ="2985"
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
        dbText "Name" ="SumOfCountQuantity"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfOnShelfDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2190"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Loc1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Loc2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Loc3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Loc4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Scans"
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
    Right =1209
    Bottom =194
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =163
        Top =-12
        Right =591
        Bottom =268
        Top =0
        Name ="ScanBatchLineCounts"
        Name =""
    End
    Begin
        Left =682
        Top =-10
        Right =925
        Bottom =283
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
End
