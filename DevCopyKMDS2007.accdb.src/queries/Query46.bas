Operation =3
Name ="PurchaseOrderLine"
Option =0
Begin InputTables
    Name ="A1265"
    Name ="Item"
End
Begin OutputColumns
    Alias ="POID"
    Name ="PurchaseOrderID"
    Expression ="932"
    Alias ="VD"
    Name ="VendorID"
    Expression ="1"
    Name ="LineID"
    Expression ="A1265.ID"
    Name ="ItemID"
    Expression ="A1265.Code"
    Name ="QuantityOrdered"
    Expression ="A1265.qty"
    Name ="QuantityShipped"
    Expression ="A1265.qty"
    Name ="Price"
    Expression ="Item.Price"
    Name ="Cost"
    Expression ="Item.Cost"
End
Begin Joins
    LeftTable ="A1265"
    RightTable ="Item"
    Expression ="A1265.Code = Item.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="A1265.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A1265.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A1265.qty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="POID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1002"
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
    State =2
    Left =-4
    Top =-23
    Right =1553
    Bottom =877
    Left =-1
    Top =-1
    Right =1534
    Bottom =483
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="A1265"
        Name =""
    End
    Begin
        Left =299
        Top =73
        Right =443
        Bottom =400
        Top =0
        Name ="Item"
        Name =""
    End
End
