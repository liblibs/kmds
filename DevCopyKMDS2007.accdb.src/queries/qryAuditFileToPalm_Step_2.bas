dbMemo "SQL" ="UPDATE InvoLineUPCMaxDate INNER JOIN UPCItem ON InvoLineUPCMaxDate.InputCode=UPC"
    "Item.ID SET UPCItem.ChangeDate = InvoLineUPCMaxDate.maxofsetupdate;\015\012"
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
        dbText "Name" ="UPCItem.ChangeDate"
        dbInteger "ColumnWidth" ="3180"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
