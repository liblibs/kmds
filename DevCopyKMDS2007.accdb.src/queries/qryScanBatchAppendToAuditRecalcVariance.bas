Operation =4
Option =0
Begin InputTables
    Name ="Audit"
End
Begin OutputColumns
    Name ="Audit.VarianceAmount"
    Expression ="[PhysicalCount]-[OnShelf]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Audit.VarianceAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Audit.OnShelf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Audit.PhysicalCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Var"
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
    Bottom =347
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =264
        Bottom =236
        Top =0
        Name ="Audit"
        Name =""
    End
End
