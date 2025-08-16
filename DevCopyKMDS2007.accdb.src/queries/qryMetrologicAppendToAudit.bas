Operation =3
Name ="Audit"
Option =0
Begin InputTables
    Name ="Item"
    Name ="qryMetrologicScanBatchVariance"
End
Begin OutputColumns
    Name ="ItemID"
    Expression ="Item.ID"
    Alias ="PhysicalCount"
    Name ="PhysicalCount"
    Expression ="qryMetrologicScanBatchVariance.Count"
    Name ="OnShelf"
    Expression ="qryMetrologicScanBatchVariance.OnShelf"
    Name ="Price"
    Expression ="Item.Price"
    Name ="Cost"
    Expression ="Item.Cost"
    Alias ="CountFront"
    Name ="CountFront"
    Expression ="qryMetrologicScanBatchVariance.Count"
    Alias ="Expr1"
    Name ="SetupDate"
    Expression ="Now()"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="qryMetrologicScanBatchVariance"
    Expression ="Item.ID = qryMetrologicScanBatchVariance.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="ScanBatchLineCounts.ItemID"
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
        dbText "Name" ="Item.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Audit.PhysicalCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfOnShelf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScanBatchLineCounts.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfSetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScanBatchLineCounts.ChangeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfChangeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Variance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PhysicalCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Audit.CountFront"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountFront"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryMetrologicScanBatchVariance.OnShelf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
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
    Right =1193
    Bottom =279
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =294
        Top =12
        Right =437
        Bottom =199
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =591
        Top =22
        Right =762
        Bottom =236
        Top =0
        Name ="qryMetrologicScanBatchVariance"
        Name =""
    End
End
