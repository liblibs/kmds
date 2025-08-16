Operation =3
Name ="InvoLine"
Option =0
Begin InputTables
    Name ="Item"
    Name ="qryMetrologicScanBatchVariance"
End
Begin OutputColumns
    Alias ="Expr1"
    Name ="InvoHeadID"
    Expression ="Format(Now(),\"yymmddhhnn\")"
    Name ="ID"
    Expression ="qryMetrologicScanBatchVariance.InvoLineID"
    Name ="CashierID"
    Expression ="qryMetrologicScanBatchVariance.CashierID"
    Name ="Code"
    Expression ="Item.ID"
    Name ="SalesDept"
    Expression ="qryMetrologicScanBatchVariance.OnShelf"
    Alias ="PhysicalCount"
    Name ="Quantity"
    Expression ="qryMetrologicScanBatchVariance.Count"
    Name ="Price"
    Expression ="Item.Price"
    Name ="Cost"
    Expression ="Item.Cost"
    Alias ="Expr3"
    Name ="SetupDate"
    Expression ="Now()"
    Alias ="InputCode"
    Name ="InputCode"
    Expression ="\"Metrologic\""
    Alias ="TransType"
    Name ="TransType"
    Expression ="\"Y\""
    Alias ="Expr2"
    Name ="CashSysID"
    Expression ="50"
End
Begin Joins
    LeftTable ="qryMetrologicScanBatchVariance"
    RightTable ="Item"
    Expression ="qryMetrologicScanBatchVariance.ID = Item.ID"
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
    Begin
        dbText "Name" ="qryMetrologicScanBatchVariance.CashierID"
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
        dbText "Name" ="qryMetrologicScanBatchVariance.OnShelf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr3"
        dbInteger "ColumnWidth" ="2130"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryMetrologicScanBatchVariance.InvoLineID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =70
    Top =60
    Right =1295
    Bottom =737
    Left =-1
    Top =-1
    Right =1193
    Bottom =352
    Left =0
    Top =0
    ColumnsShown =651
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
        Left =706
        Top =39
        Right =989
        Bottom =244
        Top =0
        Name ="qryMetrologicScanBatchVariance"
        Name =""
    End
End
