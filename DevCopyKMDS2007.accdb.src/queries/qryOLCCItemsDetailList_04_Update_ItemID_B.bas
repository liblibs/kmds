Operation =4
Option =0
Where ="(((Right([olcccode],1))=\"B\"))"
Begin InputTables
    Name ="OLCCItemsDetailList"
End
Begin OutputColumns
    Name ="OLCCItemsDetailList.iTEMID"
    Expression ="Left([OLCCCODE],Len([OLCCCODE])-1) & \"2\""
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
    Begin
        dbText "Name" ="lastch"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="iTEMID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCItemsDetailList.iTEMID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =21
    Top =8
    Right =1507
    Bottom =841
    Left =-1
    Top =-1
    Right =1463
    Bottom =487
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =147
        Top =22
        Right =291
        Bottom =448
        Top =0
        Name ="OLCCItemsDetailList"
        Name =""
    End
End
