dbMemo "SQL" ="SELECT qryInvoline5Charges.Auth, qryInvoline5Charges.InvoHeadID, qryInvoline5Cha"
    "rges.CashSysID, qryInvoline5Charges.CashierID, qryInvoline5Charges.ShiftID, qryI"
    "nvoline5Charges.Code, qryInvoline5Charges.Quantity, qryInvoline5Charges.Price, p"
    "cChargeTrans.Amount, [Price]-[Amount] AS Diff, pcChargeTrans.TroutD\015\012FROM "
    "qryInvoline5Charges INNER JOIN pcChargeTrans ON qryInvoline5Charges.Auth=pcCharg"
    "eTrans.Auth\015\012WHERE ((([Price]-[Amount])<>0));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Diff"
        dbLong "AggregateType" ="-1"
    End
End
