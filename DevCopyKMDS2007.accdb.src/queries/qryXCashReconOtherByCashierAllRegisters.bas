dbMemo "SQL" ="SELECT InvoLineAllRegisters.Code, InvoLineAllRegisters.InputCode AS Transact, Su"
    "m(InvoLineAllRegisters.Quantity) AS SumOfQuantity, Sum(InvoLineAllRegisters.Pric"
    "e) AS SumOfPrice, Cashier.ID, Cashier.Name\015\012FROM Cashier INNER JOIN (InvoL"
    "ineAllRegisters INNER JOIN Tender ON InvoLineAllRegisters.Code=Tender.ID) ON Cas"
    "hier.ID=InvoLineAllRegisters.CashierID\015\012WHERE (((InvoLineAllRegisters.Tran"
    "sType)='T'))\015\012GROUP BY InvoLineAllRegisters.Code, InvoLineAllRegisters.Inp"
    "utCode, Cashier.ID, Cashier.Name\015\012HAVING (((InvoLineAllRegisters.Code)=1) "
    "AND ((InvoLineAllRegisters.InputCode)<>\"Cash\" And (InvoLineAllRegisters.InputC"
    "ode)<>\"Change\") AND ((Cashier.ID)=[parCashierNumber]))\015\012ORDER BY InvoLin"
    "eAllRegisters.Code;\015\012"
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
