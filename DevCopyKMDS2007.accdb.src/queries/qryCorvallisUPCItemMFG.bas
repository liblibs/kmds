dbMemo "SQL" ="INSERT INTO upcitem ( ID, ItemID )\015\012SELECT checkdigit11(Right([mfg part],1"
    "1)) AS UPCCode, ItemCorv1.sku\015\012FROM ItemCorv1 INNER JOIN Item ON ItemCorv1"
    ".sku=Item.ID\015\012WHERE (((Len([mfg part]))=12) AND ((ItemCorv1.[vendor number"
    "])=100));\015\012"
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
        dbText "Name" ="ItemCorv1.[scan number]"
        dbInteger "ColumnWidth" ="3810"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCCode"
        dbLong "AggregateType" ="-1"
    End
End
