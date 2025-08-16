dbMemo "SQL" ="INSERT INTO UPCItem ( ID, ItemID )\015\012SELECT upcatoe([id]) AS UPCE, UPCItem."
    "ItemID\015\012FROM UPCItem\015\012WHERE (((Len([id]))=10) AND ((Len(upcatoe([id]"
    ")))=8));\015\012"
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
        dbText "Name" ="UPCItem.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lend"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lupcid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCItem.ItemID"
        dbLong "AggregateType" ="-1"
    End
End
