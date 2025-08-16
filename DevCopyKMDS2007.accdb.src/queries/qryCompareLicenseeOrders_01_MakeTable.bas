Operation =2
Name ="CompareScanItemToPurchaseOrder"
Option =0
Where ="(((OrderLine.OrdheadID)=[Forms]![ClubOrderCompare]![txtPurchaseOrderID_1]))"
Begin InputTables
    Name ="OrderLine"
End
Begin OutputColumns
    Alias ="PurchaseOrderID"
    Expression ="OrderLine.OrdheadID"
    Expression ="OrderLine.LineID"
    Expression ="OrderLine.ItemID"
    Expression ="OrderLine.QuantityShipped"
    Expression ="OrderLine.Price"
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
    Begin
        dbText "Name" ="OrderLine.PurchaseOrderID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderLine.LineID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderLine.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderLine.QuantityShipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderLine.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderLine.OrdheadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-22
    Top =162
    Right =1203
    Bottom =839
    Left =-1
    Top =-1
    Right =1193
    Bottom =238
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =484
        Top =12
        Right =628
        Bottom =266
        Top =0
        Name ="OrderLine"
        Name =""
    End
End
