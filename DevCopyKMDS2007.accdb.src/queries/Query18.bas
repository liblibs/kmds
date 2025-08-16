dbMemo "SQL" ="SELECT Item.ID, Item.Price, [01 Inventory Value By Code].Cost, [01 Inventory Val"
    "ue By Code].CurPrice, [01 Inventory Value By Code].ID, [01 Inventory Value By Co"
    "de].Description\015\012FROM [01 Inventory Value By Code] LEFT JOIN Item ON [01 I"
    "nventory Value By Code].ID = Item.ID\015\012WHERE (((Item.Price) Is Not Null));\015"
    "\012"
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
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[01 Inventory Value By Code].Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[01 Inventory Value By Code].CurPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[01 Inventory Value By Code].ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[01 Inventory Value By Code].Description"
        dbLong "AggregateType" ="-1"
    End
End
