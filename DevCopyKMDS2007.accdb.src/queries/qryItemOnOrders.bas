Operation =1
Option =0
Where ="(((PurchaseOrderLine.QuantityShipped)<>0))"
Begin InputTables
    Name ="PurchaseOrderLine"
    Name ="PurchaseOrder"
    Name ="Vendor"
End
Begin OutputColumns
    Expression ="PurchaseOrder.ID"
    Expression ="PurchaseOrder.OrderDate"
    Expression ="Vendor.ID"
    Expression ="Vendor.Name"
    Expression ="PurchaseOrderLine.ItemID"
    Expression ="PurchaseOrderLine.QuantityShipped"
    Expression ="PurchaseOrderLine.Price"
    Expression ="PurchaseOrderLine.Cost"
    Expression ="PurchaseOrderLine.SetupDate"
    Expression ="PurchaseOrderLine.ChangeDate"
    Alias ="ExtPrice"
    Expression ="[QuantityShipped]*[Price]"
    Alias ="ExtCost"
    Expression ="[QuantityShipped]*[Cost]"
End
Begin Joins
    LeftTable ="PurchaseOrderLine"
    RightTable ="PurchaseOrder"
    Expression ="PurchaseOrderLine.PurchaseOrderID = PurchaseOrder.ID"
    Flag =1
    LeftTable ="PurchaseOrder"
    RightTable ="Vendor"
    Expression ="PurchaseOrder.VendorID = Vendor.ID"
    Flag =1
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
        dbText "Name" ="ExtPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrder.ID"
        dbInteger "ColumnWidth" ="390"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrder.OrderDate"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.ID"
        dbInteger "ColumnWidth" ="975"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.Name"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.ItemID"
        dbInteger "ColumnWidth" ="810"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.ChangeDate"
        dbInteger "ColumnWidth" ="2175"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =453
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =62
        Top =7
        Right =158
        Bottom =241
        Top =0
        Name ="PurchaseOrderLine"
        Name =""
    End
    Begin
        Left =211
        Top =12
        Right =307
        Bottom =171
        Top =0
        Name ="PurchaseOrder"
        Name =""
    End
    Begin
        Left =370
        Top =6
        Right =460
        Bottom =240
        Top =0
        Name ="Vendor"
        Name =""
    End
End
