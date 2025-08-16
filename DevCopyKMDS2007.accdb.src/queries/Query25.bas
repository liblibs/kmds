dbMemo "SQL" ="UPDATE KenoSalesHist INNER JOIN Item ON KenoSalesHist.ItemID=Item.ID SET Item.Un"
    "itsPerCase = [ucase], Item.UnitsPerOrder = [uorder];\015\012"
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
        dbText "Name" ="Item.UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.UnitsPerOrder"
        dbLong "AggregateType" ="-1"
    End
End
