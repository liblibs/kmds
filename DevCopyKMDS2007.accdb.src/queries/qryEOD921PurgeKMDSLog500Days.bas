dbMemo "SQL" ="DELETE *\015\012FROM dbo_KMDSLog\015\012WHERE ((( dbo_KMDSLog.SetupDate)<Now()-5"
    "00));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End
