dbMemo "SQL" ="UPDATE MeadowlandInventoryReport13 INNER JOIN Item ON MeadowlandInventoryReport1"
    "3.NewItemID=Item.ID SET Item.OrderAt = [NewMin], Item.OrderTo = [NewMax];\015\012"
dbMemo "Connect" =""
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
        dbText "Name" ="MeadowlandInventoryReport13.NewMin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeadowlandInventoryReport13.NewMax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.OrderAt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.OrderTo"
        dbLong "AggregateType" ="-1"
    End
End
