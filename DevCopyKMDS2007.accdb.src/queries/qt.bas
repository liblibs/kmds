dbMemo "SQL" ="INSERT INTO ItemMovement ( Code, Description, CurrMM, YTDQty, Jan, Feb, Mar, Apr"
    ", May, Jun, Jul, Aug, Sep, Oct, Nov, [Dec] )\015\012SELECT qryItemMovementByView"
    ".Code, TranTypes.Description, 0 AS CurrMM, Sum(qryItemMovementByView.SumOfQuanti"
    "ty) AS YTDQty, Sum(IIf([MM]=1,[SumOfQuantity],0)) AS Jan, Sum(IIf([MM]=2,[SumOfQ"
    "uantity],0)) AS Feb, Sum(IIf([MM]=3,[SumOfQuantity],0)) AS Mar, Sum(IIf([MM]=4,["
    "SumOfQuantity],0)) AS Apr, Sum(IIf([MM]=5,[SumOfQuantity],0)) AS May, Sum(IIf([M"
    "M]=6,[SumOfQuantity],0)) AS Jun, Sum(IIf([MM]=7,[SumOfQuantity],0)) AS Jul, Sum("
    "IIf([MM]=8,[SumOfQuantity],0)) AS Aug, Sum(IIf([MM]=9,[SumOfQuantity],0)) AS Sep"
    ", Sum(IIf([MM]=10,[SumOfQuantity],0)) AS Oct, Sum(IIf([MM]=11,[SumOfQuantity],0)"
    ") AS Nov, Sum(IIf([MM]=12,[SumOfQuantity],0)) AS [Dec]\015\012FROM qryItemMoveme"
    "ntByView LEFT JOIN TranTypes ON qryItemMovementByView.TransType=TranTypes.Code\015"
    "\012GROUP BY qryItemMovementByView.Code, TranTypes.Description, 0\015\012ORDER B"
    "Y TranTypes.Description;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbBoolean "UseTransaction" ="-1"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Jan"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Feb"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mar"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Apr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="May"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Jun"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Jul"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Aug"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sep"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Oct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Nov"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="YTDQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurrMM"
        dbLong "AggregateType" ="-1"
    End
End
