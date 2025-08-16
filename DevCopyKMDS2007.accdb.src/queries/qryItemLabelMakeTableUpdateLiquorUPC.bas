Operation =4
Option =0
Where ="(((ItemLabels.ID) Between 1000 And 99999))"
Begin InputTables
    Name ="ItemLabels"
End
Begin OutputColumns
    Name ="ItemLabels.UPCCodeMax"
    Expression ="Format([id],\"000000\")"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="ItemLabels.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemLabels.UPCCodeMax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemLabels.SalesDpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewUPC"
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
    Bottom =311
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =47
        Top =12
        Right =192
        Bottom =276
        Top =0
        Name ="ItemLabels"
        Name =""
    End
End
