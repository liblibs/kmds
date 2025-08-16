Operation =1
Option =0
Begin InputTables
    Name ="CompareScanItemToPurchaseOrder"
    Name ="Item"
End
Begin OutputColumns
    Expression ="CompareScanItemToPurchaseOrder.PurchaseOrderID"
    Expression ="CompareScanItemToPurchaseOrder.LineID"
    Expression ="CompareScanItemToPurchaseOrder.ItemID"
    Expression ="Item.Description"
    Expression ="CompareScanItemToPurchaseOrder.QuantityShipped"
    Expression ="CompareScanItemToPurchaseOrder.QuantityScanned"
    Alias ="Diff"
    Expression ="[QuantityScanned]-[QuantityShipped]"
    Expression ="Item.Price"
    Expression ="CompareScanItemToPurchaseOrder.FullCase"
End
Begin Joins
    LeftTable ="CompareScanItemToPurchaseOrder"
    RightTable ="Item"
    Expression ="CompareScanItemToPurchaseOrder.ItemID = Item.ID"
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
        dbText "Name" ="CompareScanItemToPurchaseOrder.PurchaseOrderID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompareScanItemToPurchaseOrder.LineID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompareScanItemToPurchaseOrder.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompareScanItemToPurchaseOrder.QuantityShipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompareScanItemToPurchaseOrder.QuantityScanned"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Diff"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompareScanItemToPurchaseOrder.FullCase"
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
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =204
        Bottom =288
        Top =0
        Name ="CompareScanItemToPurchaseOrder"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =471
        Bottom =284
        Top =0
        Name ="Item"
        Name =""
    End
End
