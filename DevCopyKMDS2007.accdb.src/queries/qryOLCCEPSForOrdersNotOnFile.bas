Operation =1
Option =0
Where ="(((Item.ID) Is Null) And ((OLCCEPS.OrderNumber)=Forms!OLCCPackingList!OrderNumbe"
    "r))"
Begin InputTables
    Name ="OLCCEPS"
    Name ="Item"
End
Begin OutputColumns
    Expression ="OLCCEPS.OLCCItem"
    Expression ="OLCCEPS.Description"
    Expression ="OLCCEPS.UnitsPerCase"
    Expression ="OLCCEPS.Shipped"
    Expression ="OLCCEPS.Price"
    Expression ="OLCCEPS.UnitsPerCase"
    Expression ="OLCCEPS.OrderNumber"
    Expression ="Item.ID"
    Expression ="OLCCEPS.OrderNumber"
End
Begin Joins
    LeftTable ="OLCCEPS"
    RightTable ="Item"
    Expression ="OLCCEPS.OLCCItem = Item.ID"
    Flag =2
End
Begin OrderBy
    Expression ="OLCCEPS.OLCCItem"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="OLCCEPS.OLCCItem"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCEPS.UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCEPS.OrderNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCEPS.Shipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCEPS.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCEPS.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1002"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1006"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =75
    Top =124
    Right =1300
    Bottom =801
    Left =-1
    Top =-1
    Right =1193
    Bottom =189
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =67
        Top =0
        Right =211
        Bottom =255
        Top =0
        Name ="OLCCEPS"
        Name =""
    End
    Begin
        Left =342
        Top =16
        Right =486
        Bottom =160
        Top =0
        Name ="Item"
        Name =""
    End
End
