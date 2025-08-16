Operation =2
Name ="ScioNotOnFile"
Option =0
Where ="(((Item.Description) Is Null))"
Begin InputTables
    Name ="Item"
    Name ="OLCCitem"
    Name ="ScioInitialOrder"
End
Begin OutputColumns
    Alias ="ID"
    Expression ="ScioInitialOrder.Code"
    Alias ="SizeID"
    Expression ="Right([ScioInitialOrder].[code],1)"
    Alias ="VendorID"
    Expression ="1"
    Expression ="OLCCitem.Description"
    Expression ="ScioInitialOrder.Price"
    Alias ="Cost"
    Expression ="ScioInitialOrder.Price"
    Expression ="OLCCitem.UnitsPerCase"
    Alias ="UnitsPerOrder"
    Expression ="OLCCitem.UnitsPerCase"
    Alias ="SalesDepartmentId"
    Expression ="1"
End
Begin Joins
    LeftTable ="ScioInitialOrder"
    RightTable ="Item"
    Expression ="ScioInitialOrder.Code = Item.ID"
    Flag =2
    LeftTable ="ScioInitialOrder"
    RightTable ="OLCCitem"
    Expression ="ScioInitialOrder.Code = OLCCitem.Code"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScioInitialOrder.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCitem.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeID"
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
        dbText "Name" ="ScioInitialOrder.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCitem.UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsPerOrder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1002"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1001"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID"
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
    Bottom =547
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =628
        Top =13
        Right =772
        Bottom =508
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =499
        Top =0
        Name ="OLCCitem"
        Name =""
    End
    Begin
        Left =449
        Top =11
        Right =593
        Bottom =224
        Top =0
        Name ="ScioInitialOrder"
        Name =""
    End
End
