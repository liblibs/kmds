dbMemo "SQL" ="SELECT AverageDaySupplyLevels.ID, AverageDaySupplyLevels.DaysSupply\015\012FROM "
    "AverageDaySupplyLevels\015\012ORDER BY AverageDaySupplyLevels.ID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="AverageDaySupplyLevels.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AverageDaySupplyLevels.DaysSupply"
        dbLong "AggregateType" ="-1"
    End
End
