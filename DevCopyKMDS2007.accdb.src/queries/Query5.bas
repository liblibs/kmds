dbMemo "SQL" ="INSERT INTO PurchaseOrderLine ( ItemID, Price, VendorID, Cost, PurchaseOrderID, "
    "LineID, QuantityOrdered, QuantityShipped )\015\012SELECT TalentPO.ItemID, Talent"
    "PO.Price, TalentPO.VendorID, TalentPO.Cost, TalentPO.PurchaseOrderID, TalentPO.L"
    "ineID, Sheet1.Qty, Sheet1.Qty\015\012FROM TalentPO INNER JOIN Sheet1 ON TalentPO"
    ".ItemID=Sheet1.[Item Code];\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Sheet1.[Item Description]"
        dbInteger "ColumnWidth" ="3285"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
