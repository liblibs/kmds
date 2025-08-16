Operation =4
Option =0
Where ="(((Len([UPCCode]))=8 Or (Len([UPCCode]))=12))"
Begin InputTables
    Name ="dbo_OLCCScanCode"
End
Begin OutputColumns
    Name ="dbo_OLCCScanCode.UPCCodeKMDS"
    Expression ="[UPCCode]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="dbo_OLCCScanCode.UPCCode"
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
    Top =40
    Right =1246
    Bottom =886
    Left =-1
    Top =-1
    Right =1223
    Bottom =576
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =183
        Top =0
        Name ="dbo_OLCCScanCode"
        Name =""
    End
End
