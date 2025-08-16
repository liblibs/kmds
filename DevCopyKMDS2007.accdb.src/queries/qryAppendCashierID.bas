dbMemo "SQL" ="INSERT INTO TimeClockDetail ( CashierID )\015\012SELECT [CashierID] AS Cash;\015"
    "\012"
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
        dbText "Name" ="Cash"
        dbLong "AggregateType" ="-1"
    End
End
