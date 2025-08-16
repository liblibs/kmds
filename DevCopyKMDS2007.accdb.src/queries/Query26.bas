dbMemo "SQL" ="INSERT INTO PurchaseOrderLine ( PurchaseOrderID, VendorID, ItemID, QuantityOrder"
    "ed, QuantityShipped, LineID )\015\012SELECT 4113 AS PurchaseOrderID, 1 AS Vendor"
    "ID, Sheet1.ItemID, Sheet1.qty AS QuantityOrdered, Sheet1.qty AS QuantityShipped,"
    " Sheet1.ID\015\012FROM Sheet1\015\012WHERE (((Sheet1.qty)<>0));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Sheet1.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sheet1.description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KenoSalesHist.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sheet1.qty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sheet1.cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1001"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sheet1.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1002"
        dbLong "AggregateType" ="-1"
    End
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
        dbText "Name" ="QuantityShipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1004"
        dbLong "AggregateType" ="-1"
    End
End
