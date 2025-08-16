dbMemo "SQL" ="SELECT InvoLine.Code, InvoLine.InputCode AS Transact, Sum(InvoLine.Quantity) AS "
    "SumOfQuantity, Sum(InvoLine.Price) AS SumOfPrice\015\012FROM InvoLine INNER JOIN"
    " Tender ON InvoLine.Code=Tender.ID\015\012GROUP BY InvoLine.Code, InvoLine.Input"
    "Code, InvoLine.TransType\015\012HAVING (((InvoLine.Code)=1) AND ((InvoLine.Input"
    "Code)=\"Cash\") AND ((InvoLine.TransType)=\"T\")) OR (((InvoLine.Code)=1) AND (("
    "InvoLine.InputCode)=\"Change\") AND ((InvoLine.TransType)=\"T\"))\015\012ORDER B"
    "Y InvoLine.Code;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Transact"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPrice"
        dbLong "AggregateType" ="-1"
    End
End
