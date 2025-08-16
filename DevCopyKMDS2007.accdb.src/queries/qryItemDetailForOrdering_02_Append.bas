Operation =3
Name ="OLASItemDetailForOrdering"
Option =0
Where ="(((ItemDetailForOrdering.F2) Like \"999*\"))"
Begin InputTables
    Name ="ItemDetailForOrdering"
End
Begin OutputColumns
    Name ="ItemCode"
    Expression ="ItemDetailForOrdering.F1"
    Name ="NewItemCode"
    Expression ="ItemDetailForOrdering.F2"
    Name ="Description"
    Expression ="ItemDetailForOrdering.F3"
    Name ="Category"
    Expression ="ItemDetailForOrdering.F4"
    Name ="Size"
    Expression ="ItemDetailForOrdering.F5"
    Name ="Status"
    Expression ="ItemDetailForOrdering.F6"
    Name ="NoWillCalls"
    Expression ="ItemDetailForOrdering.F7"
    Name ="MinOrderQty"
    Expression ="ItemDetailForOrdering.F8"
    Name ="MaxOrderQty"
    Expression ="ItemDetailForOrdering.F9"
    Name ="BottlesPerCase"
    Expression ="ItemDetailForOrdering.F10"
    Name ="PricePerBottle"
    Expression ="ItemDetailForOrdering.F11"
    Name ="QuantityAvailable"
    Expression ="ItemDetailForOrdering.F12"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="ItemDetailForOrdering.F2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F11"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-67
    Top =39
    Right =1074
    Bottom =780
    Left =-1
    Top =-1
    Right =1109
    Bottom =423
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="ItemDetailForOrdering"
        Name =""
    End
End
