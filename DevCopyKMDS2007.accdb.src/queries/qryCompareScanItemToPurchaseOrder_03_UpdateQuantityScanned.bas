Operation =4
Option =0
Begin InputTables
    Name ="CompareScanItemToPurchaseOrder"
    Name ="CompareScanItemToPurchaseOrderScanCounts"
End
Begin OutputColumns
    Name ="CompareScanItemToPurchaseOrder.QuantityScanned"
    Expression ="[SumOfCountQuantity]"
End
Begin Joins
    LeftTable ="CompareScanItemToPurchaseOrderScanCounts"
    RightTable ="CompareScanItemToPurchaseOrder"
    Expression ="CompareScanItemToPurchaseOrderScanCounts.ItemID = CompareScanItemToPurchaseOrder"
        ".ItemID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="CompareScanItemToPurchaseOrder.QuantityScanned"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryCompareScanItemToPurchaseOrder_02_ScanTotals.SumOfCountQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompareScanItemToPurchaseOrderScanCounts.SumOfCountQuantity"
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
    Bottom =220
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =206
        Bottom =187
        Top =0
        Name ="CompareScanItemToPurchaseOrder"
        Name =""
    End
    Begin
        Left =265
        Top =12
        Right =512
        Bottom =156
        Top =0
        Name ="CompareScanItemToPurchaseOrderScanCounts"
        Name =""
    End
End
