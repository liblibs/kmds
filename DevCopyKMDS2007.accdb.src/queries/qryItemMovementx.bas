dbMemo "SQL" ="SELECT InvoLine.Code, InvoLine.TransType, Sum(InvoLine.Quantity) AS SumOfQuantit"
    "y, Sum(InvoLine.Price) AS SumOfPrice, Format(InvoLine.SetupDate,\"yyyy\") AS YYY"
    "Y, Format(InvoLine.SetupDate,\"mm\") AS MM\015\012FROM InvoLine\015\012WHERE ((("
    "InvoLine.SetupDate) Between #3/1/2004# And #3/1/2005#))\015\012GROUP BY InvoLine"
    ".Code, InvoLine.TransType, Format(InvoLine.SetupDate,\"yyyy\"), Format(InvoLine."
    "SetupDate,\"mm\")\015\012ORDER BY InvoLine.Code, Format(InvoLine.SetupDate,\"yyy"
    "y\"), Format(InvoLine.SetupDate,\"mm\");\015\012"
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
        dbText "Name" ="SumOfQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="YYYY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MM"
        dbLong "AggregateType" ="-1"
    End
End
