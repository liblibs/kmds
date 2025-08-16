Operation =3
Name ="CompareScanItemToPurchaseOrder"
Option =0
Where ="(((CompareScanItemToPurchaseOrder.ItemID) Is Null))"
Begin InputTables
    Name ="CompareScanItemToPurchaseOrder"
    Name ="CompareScanItemToPurchaseOrderScanCounts"
    Name ="Item"
End
Begin OutputColumns
    Name ="ItemID"
    Expression ="CompareScanItemToPurchaseOrderScanCounts.ItemID"
    Alias ="QuantityShipped"
    Name ="QuantityShipped"
    Expression ="0"
    Name ="Price"
    Expression ="Item.Price"
    Alias ="PurchaseOrderID"
    Name ="PurchaseOrderID"
    Expression ="-1"
    Name ="QuantityScanned"
    Expression ="CompareScanItemToPurchaseOrderScanCounts.SumOfCountQuantity"
    Alias ="FullCase"
    Name ="FullCase"
    Expression ="0"
End
Begin Joins
    LeftTable ="CompareScanItemToPurchaseOrderScanCounts"
    RightTable ="CompareScanItemToPurchaseOrder"
    Expression ="CompareScanItemToPurchaseOrderScanCounts.ItemID = CompareScanItemToPurchaseOrder"
        ".ItemID"
    Flag =2
    LeftTable ="CompareScanItemToPurchaseOrderScanCounts"
    RightTable ="Item"
    Expression ="CompareScanItemToPurchaseOrderScanCounts.ItemID = Item.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="CompareScanItemToPurchaseOrder.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompareScanItemToPurchaseOrderScanCounts.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompareScanItemToPurchaseOrderScanCounts.SumOfCountQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompareScanItemToPurchaseOrder.QuantityShipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuantityShipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompareScanItemToPurchaseOrder.PurchaseOrderID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompareScanItemToPurchaseOrder.FullCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FullCase"
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
    Bottom =323
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =341
        Top =109
        Right =515
        Bottom =308
        Top =0
        Name ="CompareScanItemToPurchaseOrder"
        Name =""
    End
    Begin
        Left =120
        Top =14
        Right =264
        Bottom =206
        Top =0
        Name ="CompareScanItemToPurchaseOrderScanCounts"
        Name =""
    End
    Begin
        Left =563
        Top =12
        Right =707
        Bottom =292
        Top =0
        Name ="Item"
        Name =""
    End
End
