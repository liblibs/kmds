Operation =3
Name ="Audit"
Option =0
Begin InputTables
    Name ="OLCCAuditVariance"
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Name ="ItemID"
    Expression ="OLCCAuditVariance.ItemID"
    Name ="VarianceAmount"
    Expression ="OLCCAuditVariance.Variance"
    Name ="Price"
    Expression ="vwItemsWithCurrentPricing.Price"
    Name ="Cost"
    Expression ="vwItemsWithCurrentPricing.Cost"
    Name ="OnShelf"
    Expression ="vwItemsWithCurrentPricing.OnShelf"
    Alias ="Expr1"
    Name ="CountFront"
    Expression ="[OnShelf]+[Variance]"
    Alias ="Expr2"
    Name ="PhysicalCount"
    Expression ="[OnShelf]+[Variance]"
End
Begin Joins
    LeftTable ="OLCCAuditVariance"
    RightTable ="vwItemsWithCurrentPricing"
    Expression ="OLCCAuditVariance.ItemId = vwItemsWithCurrentPricing.ID"
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
        dbText "Name" ="OLCCAuditVariance.Variance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.OnShelf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Cost"
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
        dbText "Name" ="OLCCAuditVariance.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfVariance"
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
    Bottom =334
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =555
        Top =13
        Right =699
        Bottom =157
        Top =0
        Name ="OLCCAuditVariance"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =352
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
End
