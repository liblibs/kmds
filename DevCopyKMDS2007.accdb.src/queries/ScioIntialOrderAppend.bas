Operation =3
Name ="PurchaseOrderLine"
Option =0
Where ="(((ScioInitialOrder.ID) Is Not Null))"
Begin InputTables
    Name ="ScioInitialOrder"
End
Begin OutputColumns
    Alias ="PurchaseOrderID"
    Name ="PurchaseOrderID"
    Expression ="[PurchaseOrderNumber]"
    Alias ="LineID"
    Name ="LineID"
    Expression ="ScioInitialOrder.ID"
    Alias ="VendorID"
    Name ="VendorID"
    Expression ="1"
    Alias ="QuantityOrdered"
    Name ="QuantityOrdered"
    Expression ="ScioInitialOrder.OrderQty"
    Alias ="QuantityShipped"
    Name ="QuantityShipped"
    Expression ="ScioInitialOrder.OrderQty"
    Name ="Price"
    Expression ="ScioInitialOrder.Price"
    Alias ="Cost"
    Name ="Cost"
    Expression ="ScioInitialOrder.Price"
    Name ="ItemID"
    Expression ="ScioInitialOrder.Code"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="PurchaseOrderLine.PurchaseOrderID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.LineID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LineID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScioInitialOrder.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScioInitialOrder.OrderQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuantityOrdered"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuantityShipped"
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
        dbText "Name" ="ScioInitialOrder.Code"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =20
    Top =19
    Right =1530
    Bottom =812
    Left =-1
    Top =-1
    Right =1487
    Bottom =491
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =89
        Top =54
        Right =277
        Bottom =340
        Top =0
        Name ="ScioInitialOrder"
        Name =""
    End
End
