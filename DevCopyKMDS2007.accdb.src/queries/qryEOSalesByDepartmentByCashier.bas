dbMemo "SQL" ="PARAMETERS parCashier Long;\015\012SELECT SalesGroup.ID AS [Group], SalesGroup.D"
    "escription AS GroupDesc, SalesDepartment.ID AS Department, IIf(Involine5.Transty"
    "pe='C',\"Club Sales\",SalesDepartment.Description) AS DepartmentDesc, Sum(InvoLi"
    "ne5.Quantity) AS TotalQ, Sum(InvoLine5.Price) AS TotalP, Sum(IIf(InvoLine5.Cashs"
    "ysID=1,[Quantity],0)) AS Q1, Sum(IIf(InvoLine5.CashsysID=1,InvoLine5.Price,0)) A"
    "S D1, Sum(IIf(InvoLine5.CashsysID=2,[Quantity],0)) AS Q2, Sum(IIf(InvoLine5.Cash"
    "sysID=2,InvoLine5.Price,0)) AS D2, Sum(IIf(InvoLine5.CashsysID=3,[Quantity],0)) "
    "AS Q3, Sum(IIf(InvoLine5.CashsysID=3,InvoLine5.Price,0)) AS D3, Sum(IIf(InvoLine"
    "5.CashsysID=4,[Quantity],0)) AS Q4, Sum(IIf(InvoLine5.CashsysID=4,InvoLine5.Pric"
    "e,0)) AS D4, InvoLine5.TransType, InvoLine5.CashierID\015\012FROM SalesGroup INN"
    "ER JOIN (SalesDepartment INNER JOIN (InvoLine5 INNER JOIN Item ON InvoLine5.Code"
    "=Item.ID) ON SalesDepartment.ID=Item.SalesDepartmentID) ON SalesGroup.ID=SalesDe"
    "partment.SalesGroupID\015\012WHERE (((InvoLine5.TransType)<>'T'))\015\012GROUP B"
    "Y SalesGroup.ID, SalesGroup.Description, SalesDepartment.ID, IIf(Involine5.Trans"
    "type='C',\"Club Sales\",SalesDepartment.Description), InvoLine5.TransType, InvoL"
    "ine5.CashierID\015\012HAVING (((InvoLine5.CashierID)=[parCashier]))\015\012ORDER"
    " BY SalesGroup.ID, SalesDepartment.ID;\015\012"
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
        dbText "Name" ="TotalQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DepartmentDesc"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Department"
        dbLong "AggregateType" ="-1"
    End
End
