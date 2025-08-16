dbMemo "SQL" ="SELECT ScanBatchLineCounts.ItemID, Sum(ScanBatchLineCounts.CountQuantity) AS Cou"
    "ntQ, Last(ScanBatchLineCounts.OnShelf) AS OS, Count(ScanBatchLineCounts.CountQua"
    "ntity) AS ScanQ\015\012FROM ScanBatchLineCounts\015\012GROUP BY ScanBatchLineCou"
    "nts.ItemID\015\012HAVING (((ScanBatchLineCounts.ItemID)=1402));\015\012"
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
        dbText "Name" ="ScanBatchLineCounts.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScanQ"
        dbLong "AggregateType" ="-1"
    End
End
