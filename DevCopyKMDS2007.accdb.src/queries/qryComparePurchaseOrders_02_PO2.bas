Operation =2
Name ="CompareScanItemToPurchaseOrderScanCounts"
Option =0
Begin InputTables
    Name ="qryPurchaseOrderCompare_2"
End
Begin OutputColumns
    Expression ="qryPurchaseOrderCompare_2.ItemID"
    Alias ="SumOfCountQuantity"
    Expression ="qryPurchaseOrderCompare_2.QuantityShipped"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="ScanBatchLineCounts.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfCountQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderCompare_2.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfQuantityShipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderCompare_2.QuantityShipped"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-454
    Top =77
    Right =771
    Bottom =754
    Left =-1
    Top =-1
    Right =1193
    Bottom =306
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =510
        Top =85
        Right =654
        Bottom =229
        Top =0
        Name ="qryPurchaseOrderCompare_2"
        Name =""
    End
End
