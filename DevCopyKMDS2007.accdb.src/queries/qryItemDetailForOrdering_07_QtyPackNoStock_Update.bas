Operation =4
Option =0
Begin InputTables
    Name ="Item"
    Name ="OLASItemDetailForOrdering"
End
Begin OutputColumns
    Name ="Item.QtyPack"
    Expression ="[OLASItemDetailForOrdering].[QuantityAvailable]"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="OLASItemDetailForOrdering"
    Expression ="Item.ID = OLASItemDetailForOrdering.NewItemCode"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Item.AvgTurns"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLASItemDetailForOrdering.QuantityAvailable"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.QtyPack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLASItemDetailForOrdering.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLASItemDetailForOrdering.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLASItemDetailForOrdering.NoWillCalls"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.StockNumber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =117
    Top =184
    Right =845
    Bottom =829
    Left =-1
    Top =-1
    Right =696
    Bottom =276
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =223
        Bottom =316
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =290
        Top =8
        Right =490
        Bottom =297
        Top =0
        Name ="OLASItemDetailForOrdering"
        Name =""
    End
End
