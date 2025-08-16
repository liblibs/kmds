Operation =2
Name ="CompareScanItemToPurchaseOrder"
Option =0
Where ="(((PurchaseOrderLine.PurchaseOrderID)=[Forms]![PurchaseOrderCompare]![txtPurchas"
    "eOrderID_1]))"
Begin InputTables
    Name ="PurchaseOrderLine"
End
Begin OutputColumns
    Expression ="PurchaseOrderLine.PurchaseOrderID"
    Expression ="PurchaseOrderLine.LineID"
    Expression ="PurchaseOrderLine.ItemID"
    Expression ="PurchaseOrderLine.QuantityShipped"
    Expression ="PurchaseOrderLine.Price"
    Alias ="QuantityScanned"
    Expression ="0"
    Alias ="FullCase"
    Expression ="0"
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
        dbText "Name" ="PurchaseOrderLine.LineID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.QuantityShipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuantityScanned"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FullCase"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-618
    Top =125
    Right =607
    Bottom =802
    Left =-1
    Top =-1
    Right =1193
    Bottom =289
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =180
        Top =12
        Right =436
        Bottom =235
        Top =0
        Name ="PurchaseOrderLine"
        Name =""
    End
End
