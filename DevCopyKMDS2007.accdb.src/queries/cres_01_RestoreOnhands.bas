Operation =4
Option =0
Where ="(((Item.SalesDepartmentID)>21))"
Begin InputTables
    Name ="CRES_ITEM"
    Name ="Item"
End
Begin OutputColumns
    Name ="Item.Onhand"
    Expression ="[CRES_ITEM].[Onhand]"
End
Begin Joins
    LeftTable ="CRES_ITEM"
    RightTable ="Item"
    Expression ="CRES_ITEM.ID = Item.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Onhand"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CRES_ITEM.Onhand"
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
        Bottom =375
        Top =0
        Name ="CRES_ITEM"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =350
        Top =0
        Name ="Item"
        Name =""
    End
End
