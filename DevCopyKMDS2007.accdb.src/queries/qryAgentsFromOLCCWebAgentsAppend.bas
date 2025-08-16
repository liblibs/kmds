Operation =3
Name ="AGENTS"
Option =0
Begin InputTables
    Name ="OLCCWebAgents"
End
Begin OutputColumns
    Name ="AgentNumber"
    Expression ="OLCCWebAgents.ID"
    Name ="LastName"
    Expression ="OLCCWebAgents.Name"
    Name ="Phone"
    Expression ="OLCCWebAgents.Phone"
    Name ="Address"
    Expression ="OLCCWebAgents.Address"
    Name ="City"
    Expression ="OLCCWebAgents.Name"
    Name ="Zip"
    Expression ="OLCCWebAgents.ZIP"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="OLCCWebAgents.Address"
        dbInteger "ColumnWidth" ="2595"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCWebAgents.Field6"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
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
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="OLCCWebAgents"
        Name =""
    End
End
