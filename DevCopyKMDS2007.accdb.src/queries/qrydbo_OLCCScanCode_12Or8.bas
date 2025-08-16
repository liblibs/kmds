Operation =1
Option =0
Where ="(((Len([UPCCodeKMDS]))=12 Or (Len([UPCCodeKMDS]))=8))"
Begin InputTables
    Name ="dbo_OLCCScanCode"
End
Begin OutputColumns
    Expression ="dbo_OLCCScanCode.Code"
    Expression ="dbo_OLCCScanCode.UPCCodeKMDS"
    Alias ="lid"
    Expression ="Len([UPCCodeKMDS])"
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
        dbText "Name" ="dbo_OLCCScanCode.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_OLCCScanCode.UPCCodeKMDS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lid"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1225
    Bottom =677
    Left =-1
    Top =-1
    Right =1209
    Bottom =254
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="dbo_OLCCScanCode"
        Name =""
    End
End
