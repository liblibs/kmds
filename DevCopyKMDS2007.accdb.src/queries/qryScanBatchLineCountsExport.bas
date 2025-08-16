Operation =1
Option =0
Begin InputTables
    Name ="ScanBatchLineCounts"
End
Begin OutputColumns
    Expression ="ScanBatchLineCounts.ScanBatchHeaderID"
    Expression ="ScanBatchLineCounts.ItemID"
    Expression ="ScanBatchLineCounts.InputMethod"
    Expression ="ScanBatchLineCounts.CashierID"
    Expression ="ScanBatchLineCounts.Location"
    Expression ="ScanBatchLineCounts.ScanCode"
    Expression ="ScanBatchLineCounts.CountQuantity"
    Expression ="ScanBatchLineCounts.OnShelf"
    Expression ="ScanBatchLineCounts.SetupDate"
    Expression ="ScanBatchLineCounts.ChangeDate"
    Expression ="ScanBatchLineCounts.OnShelfDate"
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
        dbText "Name" ="ScanBatchLineCounts.OnShelfDate"
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
    Bottom =254
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="ScanBatchLineCounts"
        Name =""
    End
End
