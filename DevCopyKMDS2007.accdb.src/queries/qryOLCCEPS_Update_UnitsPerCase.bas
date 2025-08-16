Operation =4
Option =0
Where ="(((Item.UnitsPerCase)<>[OLCCEPS].[UnitsPerCase]))"
Begin InputTables
    Name ="OLCCEPS"
    Name ="Item"
End
Begin OutputColumns
    Name ="Item.UnitsPerCase"
    Expression ="[OLCCEPS].[UnitsPerCase]"
    Name ="Item.ChangeDate"
    Expression ="Now()"
End
Begin Joins
    LeftTable ="OLCCEPS"
    RightTable ="Item"
    Expression ="OLCCEPS.OLCCItem = Item.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="OLCCEPS.UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ChangeDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =42
    Top =120
    Right =1174
    Bottom =873
    Left =-1
    Top =-1
    Right =1109
    Bottom =483
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =214
        Bottom =315
        Top =0
        Name ="OLCCEPS"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =448
        Top =0
        Name ="Item"
        Name =""
    End
End
