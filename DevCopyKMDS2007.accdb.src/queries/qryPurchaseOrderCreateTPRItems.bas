dbMemo "SQL" ="PARAMETERS parVendorID Long;\015\012SELECT Item.VendorID, Item.ID AS ItemID, Ite"
    "m.Description, Item.Price, qryItemsSalePriceNow.SalePrice, Item.Cost, Item.TPR\015"
    "\012FROM (qryItemsSalePriceNow RIGHT JOIN Item ON qryItemsSalePriceNow.ItemID = "
    "Item.ID) INNER JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID\015\012WH"
    "ERE (((Item.VendorID)=[parVendorID]) AND ((Item.TPR)<>0) AND ((SubCategory.Descr"
    "iption)<>\"OBSOLETE\"))\015\012ORDER BY Item.ID;\015\012"
dbMemo "Connect" =""
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
        dbText "Name" ="Item.Description"
        dbInteger "ColumnWidth" ="2805"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Price"
        dbInteger "ColumnWidth" ="885"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.VendorID"
        dbLong "AggregateType" ="-1"
    End
End
