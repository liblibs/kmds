dbMemo "SQL" ="SELECT OrdHead.ID, InvoiceNo, Club.Name, OrdHead.InvDate, OrdHead.InvDollars, Or"
    "dHead.PaidFlag\015\012FROM OrdHead INNER JOIN Club ON OrdHead.ClubID=Club.ID;\015"
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
End
