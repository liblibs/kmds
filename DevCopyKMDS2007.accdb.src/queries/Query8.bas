dbMemo "SQL" ="SELECT ReportForms.ReportName, ReportForms.ReportType\015\012FROM ReportForms\015"
    "\012WHERE (((ReportForms.ReportType) Like \"KMDS*\"));\015\012"
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
        dbText "Name" ="ReportForms.ReportName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ReportForms.ReportType"
        dbInteger "ColumnWidth" ="3735"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
