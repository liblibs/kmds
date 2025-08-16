dbMemo "SQL" ="SELECT AgentTransferLines.ItemID, Sum(AgentTransferLines.QuantityOrdered) AS OnT"
    "ransfer\015\012FROM AgentTransferLines\015\012GROUP BY AgentTransferLines.ItemID"
    ";\015\012"
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
        dbText "Name" ="OnTransfer"
        dbLong "AggregateType" ="-1"
    End
End
