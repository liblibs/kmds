Operation =3
Name ="Audit"
Option =0
Begin InputTables
    Name ="ScanBatchLineCounts"
    Name ="Item"
End
Begin OutputColumns
    Name ="ItemID"
    Expression ="ScanBatchLineCounts.ItemID"
    Alias ="PhysicalCount"
    Name ="PhysicalCount"
    Expression ="Sum(ScanBatchLineCounts.CountQuantity)"
    Alias ="LastOfOnShelf"
    Name ="OnShelf"
    Expression ="Last(ScanBatchLineCounts.OnShelf)"
    Alias ="FirstOfPrice"
    Name ="Price"
    Expression ="First(Item.Price)"
    Alias ="FirstOfCost"
    Name ="Cost"
    Expression ="First(Item.Cost)"
    Alias ="CountFront"
    Name ="CountFront"
    Expression ="Sum(ScanBatchLineCounts.CountQuantity)"
    Alias ="FirstOfSetupDate"
    Name ="SetupDate"
    Expression ="First(ScanBatchLineCounts.SetupDate)"
    Alias ="LastOfChangeDate"
    Name ="ChangeDate"
    Expression ="Last(ScanBatchLineCounts.ChangeDate)"
End
Begin Joins
    LeftTable ="ScanBatchLineCounts"
    RightTable ="Item"
    Expression ="ScanBatchLineCounts.ItemID = Item.ID"
    Flag =1
End
Begin Groups
    Expression ="ScanBatchLineCounts.ItemID"
    GroupLevel =0
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
    Bottom =186
    Left =0
    Top =0
    ColumnsShown =655
    Begin
        Left =48
        Top =12
        Right =245
        Bottom =270
        Top =0
        Name ="ScanBatchLineCounts"
        Name =""
    End
    Begin
        Left =294
        Top =12
        Right =437
        Bottom =199
        Top =0
        Name ="Item"
        Name =""
    End
End
