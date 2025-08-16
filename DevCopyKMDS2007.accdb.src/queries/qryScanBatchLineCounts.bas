Operation =1
Option =0
Begin InputTables
    Name ="ScanBatchLineCounts"
    Name ="Item"
End
Begin OutputColumns
    Expression ="ScanBatchLineCounts.ID"
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
    Expression ="Item.Description"
    Expression ="Item.Price"
End
Begin Joins
    LeftTable ="ScanBatchLineCounts"
    RightTable ="Item"
    Expression ="ScanBatchLineCounts.ItemID=Item.ID"
    Flag =1
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
        dbText "Name" ="ScanBatchLineCounts.ChangeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScanBatchLineCounts.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScanBatchLineCounts.ScanBatchHeaderID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScanBatchLineCounts.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScanBatchLineCounts.InputMethod"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScanBatchLineCounts.CashierID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScanBatchLineCounts.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScanBatchLineCounts.ScanCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScanBatchLineCounts.CountQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScanBatchLineCounts.OnShelf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScanBatchLineCounts.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Price"
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
    Bottom =293
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =257
        Top =0
        Name ="ScanBatchLineCounts"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =278
        Top =0
        Name ="Item"
        Name =""
    End
End
