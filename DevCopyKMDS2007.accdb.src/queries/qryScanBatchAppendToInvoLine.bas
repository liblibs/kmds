Operation =3
Name ="InvoLine"
Option =0
Begin InputTables
    Name ="ScanBatchLineCounts"
    Name ="Item"
    Name ="qryMetrologicScanBatchVariance"
End
Begin OutputColumns
    Name ="CashierID"
    Expression ="ScanBatchLineCounts.CashierID"
    Name ="Code"
    Expression ="ScanBatchLineCounts.ItemID"
    Name ="SalesDept"
    Expression ="ScanBatchLineCounts.OnShelf"
    Alias ="PhysicalCount"
    Name ="Quantity"
    Expression ="ScanBatchLineCounts.CountQuantity"
    Name ="Price"
    Expression ="Item.Price"
    Name ="Cost"
    Expression ="Item.Cost"
    Name ="SetupDate"
    Expression ="ScanBatchLineCounts.ChangeDate"
    Alias ="InputCode"
    Name ="InputCode"
    Expression ="\"ScanBatch\""
    Alias ="TransType"
    Name ="TransType"
    Expression ="\"Y\""
    Alias ="Expr1"
    Name ="InvoHeadID"
    Expression ="Format(Now(),\"yymmddhhnn\")"
    Alias ="Expr2"
    Name ="CashSysID"
    Expression ="50"
End
Begin Joins
    LeftTable ="ScanBatchLineCounts"
    RightTable ="Item"
    Expression ="ScanBatchLineCounts.ItemID = Item.ID"
    Flag =1
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
    Begin
        dbText "Name" ="ScanBatchLineCounts.CashierID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.InputCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InputCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.TransType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashSysID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Format(Now(),\"yymmddhhnn\")"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="50"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-43
    Top =306
    Right =1182
    Bottom =983
    Left =-1
    Top =-1
    Right =1193
    Bottom =386
    Left =0
    Top =0
    ColumnsShown =651
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
        Right =465
        Bottom =372
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =645
        Top =23
        Right =789
        Bottom =300
        Top =0
        Name ="qryMetrologicScanBatchVariance"
        Name =""
    End
End
