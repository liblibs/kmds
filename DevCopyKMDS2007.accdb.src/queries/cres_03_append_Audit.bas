Operation =3
Name ="Audit"
Option =0
Where ="((([sumofquantity]-[onhand])<0))"
Begin InputTables
    Name ="CrescInventory"
    Name ="Item"
End
Begin OutputColumns
    Name ="ItemID"
    Expression ="CrescInventory.ID"
    Name ="PhysicalCount"
    Expression ="CrescInventory.SumOfQuantity"
    Name ="OnShelf"
    Expression ="Item.Onhand"
    Name ="Price"
    Expression ="Item.Price"
    Name ="Cost"
    Expression ="Item.Cost"
    Alias ="va"
    Name ="VarianceAmount"
    Expression ="[sumofquantity]-[onhand]"
    Name ="CountFront"
    Expression ="CrescInventory.sumofquantity"
End
Begin Joins
    LeftTable ="CrescInventory"
    RightTable ="Item"
    Expression ="CrescInventory.ID = Item.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="CrescInventory.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CrescInventory.SumOfQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Onhand"
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
        dbText "Name" ="va"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vb"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-18
    Top =50
    Right =1513
    Bottom =883
    Left =-1
    Top =-1
    Right =1508
    Bottom =531
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="CrescInventory"
        Name =""
    End
    Begin
        Left =222
        Top =106
        Right =366
        Bottom =391
        Top =0
        Name ="Item"
        Name =""
    End
End
