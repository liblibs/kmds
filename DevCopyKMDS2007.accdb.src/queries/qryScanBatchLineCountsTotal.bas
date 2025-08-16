Operation =1
Option =0
Begin InputTables
    Name ="ScanBatchLineCounts"
End
Begin OutputColumns
    Expression ="ScanBatchLineCounts.ScanBatchHeaderID"
    Expression ="ScanBatchLineCounts.ItemID"
    Alias ="SumOfCountQuantity"
    Expression ="Sum(ScanBatchLineCounts.CountQuantity)"
    Alias ="LastOfOnShelf"
    Expression ="Last(ScanBatchLineCounts.OnShelf)"
End
Begin Groups
    Expression ="ScanBatchLineCounts.ScanBatchHeaderID"
    GroupLevel =0
    Expression ="ScanBatchLineCounts.ItemID"
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
        dbText "Name" ="ScanBatchLineCounts.ScanBatchHeaderID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScanBatchLineCounts.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfCountQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfOnShelf"
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
    Bottom =208
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =72
        Top =-1
        Right =316
        Bottom =255
        Top =0
        Name ="ScanBatchLineCounts"
        Name =""
    End
End
