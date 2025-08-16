Operation =4
Option =0
Where ="(((ScioInitialOrder.Price)<>[Item].[price]))"
Begin InputTables
    Name ="ScioInitialOrder"
    Name ="Item"
End
Begin OutputColumns
    Name ="Item.Price"
    Expression ="[ScioInitialOrder].[price]"
    Name ="Item.Cost"
    Expression ="[ScioInitialOrder].[price]"
End
Begin Joins
    LeftTable ="ScioInitialOrder"
    RightTable ="Item"
    Expression ="ScioInitialOrder.Code = Item.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="ScioInitialOrder.Price"
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
End
Begin
    State =0
    Left =0
    Top =40
    Right =1531
    Bottom =873
    Left =-1
    Top =-1
    Right =1508
    Bottom =563
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =391
        Top =0
        Name ="ScioInitialOrder"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =501
        Top =0
        Name ="Item"
        Name =""
    End
End
