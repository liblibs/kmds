Operation =3
Name ="dbo_OLCCScanCode"
Option =0
Where ="(((Len([UPCCode]))=10 Or (Len([UPCCode]))=6 Or (Len([UPCCode]))=14))"
Begin InputTables
    Name ="OLCCScanCode"
End
Begin OutputColumns
    Alias ="FirstOfCode"
    Name ="Code"
    Expression ="First(OLCCScanCode.Code)"
    Name ="UPCCode"
    Expression ="OLCCScanCode.UPCCode"
End
Begin Groups
    Expression ="OLCCScanCode.UPCCode"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="OLCCScanCode.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCScanCode.UPCCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfUPCCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Len([UPCCode])"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =8
    Top =44
    Right =1233
    Bottom =721
    Left =-1
    Top =-1
    Right =1202
    Bottom =288
    Left =0
    Top =0
    ColumnsShown =655
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="OLCCScanCode"
        Name =""
    End
End
