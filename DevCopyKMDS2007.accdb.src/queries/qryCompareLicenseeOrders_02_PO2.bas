Operation =2
Name ="CompareScanItemToPurchaseOrderScanCounts"
Option =0
Begin InputTables
    Name ="qryLicenseeOrderCompare_2"
End
Begin OutputColumns
    Expression ="qryLicenseeOrderCompare_2.ItemID"
    Alias ="SumOfCountQuantity"
    Expression ="qryLicenseeOrderCompare_2.QuantityShipped"
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
    Begin
        dbText "Name" ="qryLicenseeOrderCompare_2.ItemID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-446
    Top =133
    Right =779
    Bottom =810
    Left =-1
    Top =-1
    Right =1193
    Bottom =289
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qryLicenseeOrderCompare_2"
        Name =""
    End
End
