Operation =4
Option =0
Where ="(((CompareScanItemToPurchaseOrder.QuantityShipped)>=[UnitsPerCase]))"
Begin InputTables
    Name ="CompareScanItemToPurchaseOrder"
    Name ="Item"
End
Begin OutputColumns
    Name ="CompareScanItemToPurchaseOrder.FullCase"
    Expression ="True"
End
Begin Joins
    LeftTable ="CompareScanItemToPurchaseOrder"
    RightTable ="Item"
    Expression ="CompareScanItemToPurchaseOrder.ItemID = Item.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="CompareScanItemToPurchaseOrder.FullCase"
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
        dbText "Name" ="Item.UnitsPerCase"
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
    Bottom =340
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =339
        Bottom =286
        Top =0
        Name ="CompareScanItemToPurchaseOrder"
        Name =""
    End
    Begin
        Left =506
        Top =15
        Right =698
        Bottom =361
        Top =0
        Name ="Item"
        Name =""
    End
End
