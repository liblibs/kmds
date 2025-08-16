Operation =2
Name ="CompareScanItemToPurchaseOrderScanCounts"
Option =0
Begin InputTables
    Name ="ScanBatchLineCounts"
End
Begin OutputColumns
    Expression ="ScanBatchLineCounts.ItemID"
    Alias ="SumOfCountQuantity"
    Expression ="Sum(ScanBatchLineCounts.CountQuantity)"
End
Begin Groups
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
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="ScanBatchLineCounts.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfCountQuantity"
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
    Bottom =340
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =5
        Right =284
        Bottom =299
        Top =0
        Name ="ScanBatchLineCounts"
        Name =""
    End
End
