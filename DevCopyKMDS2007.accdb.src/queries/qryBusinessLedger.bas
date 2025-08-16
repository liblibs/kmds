dbMemo "SQL" ="SELECT Format(involine.SetupDate,\"mm/dd/yyyy\") AS BusDate, Sum(IIf([Transtype]"
    "=\"S\",Involine.Price,0)) AS SalesCounter, Sum(IIf([Transtype]=\"C\",Involine.Pr"
    "ice,0)) AS SalesClubs, Sum(IIf([Transtype]=\"P\",Involine.Price,0)) AS Purchases"
    ", Sum(IIf([Transtype]=\"X\" And Involine.Price>0,Involine.Price,0)) AS Transfers"
    "In, Sum(IIf([Transtype]=\"X\" And Involine.Price<0,Involine.Price,0)) AS Transfe"
    "rsOut, Sum(IIf([Transtype]=\"A\",Involine.Price,0)) AS Adjust\015\012FROM (((Inv"
    "oLine INNER JOIN TranTypes ON InvoLine.TransType=TranTypes.Code) INNER JOIN Item"
    " ON InvoLine.Code=Item.ID) INNER JOIN SalesDepartment ON Item.SalesDepartmentID="
    "SalesDepartment.ID) INNER JOIN SalesGroup ON SalesDepartment.SalesGroupID=SalesG"
    "roup.ID\015\012WHERE (((InvoLine.SetupDate) Between Forms!PrintInventoryReport!t"
    "xtDateRangeFrom And Forms!PrintInventoryReport!txtDateRangeTo) And ((SalesGroup."
    "Description)=\"Liquor\"))\015\012GROUP BY Format(involine.SetupDate,\"mm/dd/yyyy"
    "\");\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="BusDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesCounter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesClubs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Purchases"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Adjust"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransfersIn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransfersOut"
        dbLong "AggregateType" ="-1"
    End
End
