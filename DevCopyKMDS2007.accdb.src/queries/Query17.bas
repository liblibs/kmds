dbMemo "SQL" ="INSERT INTO UPCItem ( ItemID, ID )\015\012SELECT MeadowlandBigNumbers.ID, Meadow"
    "landInventoryReport13.NewUPC2\015\012FROM ((MeadowlandBigNumbers INNER JOIN Mead"
    "owlandInventoryValuex ON MeadowlandBigNumbers.ItemNo=MeadowlandInventoryValuex.I"
    "temNo) INNER JOIN MeadowlandInventoryReport13 ON MeadowlandInventoryValuex.ItemN"
    "o=MeadowlandInventoryReport13.Item) INNER JOIN SalesDepartment ON MeadowlandInve"
    "ntoryValuex.SalesDpt=SalesDepartment.ID\015\012WHERE (((MeadowlandInventoryRepor"
    "t13.NewUPC2) Is Not Null));\015\012"
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
        dbText "Name" ="MeadowlandBigNumbers.ID"
        dbInteger "ColumnWidth" ="2145"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeadowlandBigNumbers.AutoID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="newid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeadowlandInventoryReport13.NewUPC1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeadowlandInventoryReport13.NewUPC2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeadowlandInventoryValuex.description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeadowlandBigNumbers.ItemNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeadowlandInventoryReport13.Item"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeadowlandInventoryReport13.UPC1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeadowlandInventoryReport13.UPC2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeadowlandInventoryValuex.ItemNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeadowlandBigNumbers.SalesDpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeadowlandBigNumbers.description"
        dbInteger "ColumnWidth" ="2895"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeadowlandBigNumbers.Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeadowlandBigNumbers.Onhand"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeadowlandBigNumbers.Price"
        dbLong "AggregateType" ="-1"
    End
End
