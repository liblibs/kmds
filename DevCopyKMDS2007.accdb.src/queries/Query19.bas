Operation =3
Name ="PurchaseOrderLine"
Option =0
Begin InputTables
    Name ="Sheet1"
End
Begin OutputColumns
    Alias ="PurchaseOrderID"
    Name ="PurchaseOrderID"
    Expression ="2279"
    Alias ="VendorID"
    Name ="VendorID"
    Expression ="1"
    Alias ="LineID"
    Name ="LineID"
    Expression ="Sheet1.id"
    Alias ="ItemID"
    Name ="ItemID"
    Expression ="Sheet1.field1"
    Alias ="QuantityOrdered"
    Name ="QuantityOrdered"
    Expression ="Sheet1.field2"
    Alias ="Expr1"
    Name ="QuantityShipped"
    Expression ="Sheet1.Field2"
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
        dbText "Name" ="PurchaseOrderLine.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorID"
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
        dbText "Name" ="PurchaseOrderLine.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.QuantityOrdered"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuantityOrdered"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.QuantityShipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sheet1.Field4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sheet1.Field5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sheet1.Field3"
        dbInteger "ColumnWidth" ="3165"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sheet1.Field2"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1225
    Bottom =677
    Left =-1
    Top =-1
    Right =1209
    Bottom =304
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =204
        Top =0
        Name ="Sheet1"
        Name =""
    End
End
