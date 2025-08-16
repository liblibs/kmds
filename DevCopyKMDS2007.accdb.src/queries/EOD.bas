dbMemo "SQL" ="SELECT MSysObjects.Name, MSysObjects.Type\015\012FROM MSysObjects\015\012WHERE ("
    "((MSysObjects.Name) Like 'qryeod*'))\015\012ORDER BY MSysObjects.Name;\015\012"
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
        dbText "Name" ="MSysObjects.Name"
        dbInteger "ColumnWidth" ="4425"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
