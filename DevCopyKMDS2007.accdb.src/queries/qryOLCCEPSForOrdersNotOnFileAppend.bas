Operation =3
Name ="item"
Option =0
Where ="(((OLCCEPS.OrderNumber)=[prmOrderNumber]) AND ((Item.ID) Is Null) AND ((OLCCEPS."
    "OrderDate)=[prmOrderDate]))"
Begin InputTables
    Name ="OLCCEPS"
    Name ="Item"
End
Begin OutputColumns
    Name ="ID"
    Expression ="OLCCEPS.OLCCItem"
    Alias ="SizeId"
    Name ="SizeID"
    Expression ="Right([OLCCItem],2)"
    Alias ="SalesDepartmentID"
    Name ="SalesDepartmentID"
    Expression ="1"
    Alias ="VendorID"
    Name ="VendorID"
    Expression ="1"
    Name ="Description"
    Expression ="OLCCEPS.Description"
    Name ="UnitsPerCase"
    Expression ="OLCCEPS.UnitsPerCase"
    Name ="Price"
    Expression ="OLCCEPS.Price"
    Name ="Cost"
    Expression ="OLCCEPS.price"
End
Begin Parameters
    Name ="prmOrderNumber"
    Flag =4
    Name ="prmOrderDate"
    Flag =8
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
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="OLCCEPS.OLCCItem"
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
        dbText "Name" ="OLCCEPS.Shipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCEPS.Price"
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
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCEPS.OrderDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =44
    Top =11
    Right =1273
    Bottom =791
    Left =-1
    Top =-1
    Right =1197
    Bottom =423
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =302
        Bottom =272
        Top =0
        Name ="OLCCEPS"
        Name =""
    End
    Begin
        Left =412
        Top =10
        Right =601
        Bottom =403
        Top =0
        Name ="Item"
        Name =""
    End
End
