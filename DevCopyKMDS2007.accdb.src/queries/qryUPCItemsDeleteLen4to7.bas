Operation =5
Option =0
Where ="(((Len([id])) Between 4 And 7))"
Begin InputTables
    Name ="UPCItem"
End
Begin OutputColumns
    Alias ="Lid"
    Expression ="Len([id])"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Lid"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =771
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="UPCItem"
        Name =""
    End
End
