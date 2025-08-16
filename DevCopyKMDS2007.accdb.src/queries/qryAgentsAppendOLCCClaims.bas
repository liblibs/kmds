Operation =3
Name ="agents"
Option =0
Begin InputTables
End
Begin OutputColumns
    Alias ="AgentNumber"
    Name ="AgentNumber"
    Expression ="1000"
    Alias ="FirstName"
    Name ="LastName"
    Expression ="\"OLCC Claims\""
    Alias ="Active"
    Name ="ActiveForTransfer"
    Expression ="True"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Agents.AgentNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Agents.FirstName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AgentNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Active"
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
    Bottom =503
    Left =0
    Top =0
    ColumnsShown =651
End
