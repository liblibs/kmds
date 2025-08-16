dbMemo "SQL" ="UPDATE Item INNER JOIN InvoLineMinSetupDate ON Item.ID=InvoLineMinSetupDate.Code"
    " SET Item.SetupDate = [MinOfSetupDate]\015\012WHERE (((Item.SetupDate)>[MinOfSet"
    "upDate]));\015\012"
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
        dbText "Name" ="Item.SetupDate"
        dbInteger "ColumnWidth" ="2175"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
