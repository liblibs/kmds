Operation =3
Name ="PricingWorkTable"
Option =0
Where ="(((qryItemsByVendor.Price)>0) And ((qryItemsByVendor.SalesDeptID)=Forms!PricingT"
    "oolSelectBySalesDepartment!txtID))"
Begin InputTables
    Name ="qryItemsByVendor"
End
Begin OutputColumns
    Expression ="qryItemsByVendor.*"
    Alias ="Expr1"
    Name ="NewPrice"
    Expression ="qryItemsByVendor.Price"
    Alias ="Expr2"
    Name ="NewCost"
    Expression ="qryItemsByVendor.Cost"
    Alias ="NewGP"
    Name ="NewGP"
    Expression ="IIf(([price]>0),(([Price]-[cost])/[Price]),0)"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="qryItemsByVendor.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemsByVendor.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemsByVendor.Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewGP"
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
        dbText "Name" ="qryItemsByVendor.SalesDeptID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemsByVendor.Item.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemsByVendor.Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemsByVendor.Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemsByVendor.SalesDept"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemsByVendor.Category"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemsByVendor.Item.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemsByVendor.Item.Cost"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-567
    Top =71
    Right =876
    Bottom =877
    Left =-1
    Top =-1
    Right =1411
    Bottom =335
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =297
        Bottom =385
        Top =0
        Name ="qryItemsByVendor"
        Name =""
    End
End
