Operation =1
Option =0
Where ="(((Len([OLCCCode]))>6))"
Begin InputTables
    Name ="OLCCItemsDetailList"
End
Begin OutputColumns
    Alias ="Lid"
    Expression ="Len([OLCCCode])"
End
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
        dbText "Name" ="OLCCItemsDetailList.OLCCCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lid"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =6
    Top =12
    Right =1492
    Bottom =845
    Left =-1
    Top =-1
    Right =1463
    Bottom =547
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =147
        Top =22
        Right =291
        Bottom =166
        Top =0
        Name ="OLCCItemsDetailList"
        Name =""
    End
End
