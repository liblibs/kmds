Operation =5
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
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
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
    Left =0
    Top =40
    Right =1486
    Bottom =873
    Left =-1
    Top =-1
    Right =1463
    Bottom =563
    Left =0
    Top =0
    ColumnsShown =771
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
