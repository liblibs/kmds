Operation =4
Option =0
Begin InputTables
    Name ="OLCCWebAgents"
    Name ="Agents"
End
Begin OutputColumns
    Name ="Agents.LastName"
    Expression ="OLCCWebAgents.name"
    Name ="Agents.Address"
    Expression ="OLCCWebAgents.address"
    Name ="Agents.City"
    Expression ="OLCCWebAgents.City"
    Name ="Agents.ST"
    Expression ="\"OR\""
    Name ="Agents.Zip"
    Expression ="OLCCWebAgents.ZIP"
    Name ="Agents.Phone"
    Expression ="OLCCWebAgents.PHONE"
End
Begin Joins
    LeftTable ="OLCCWebAgents"
    RightTable ="Agents"
    Expression ="OLCCWebAgents.ID = Agents.AgentNumber"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "FailOnError" ="0"
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
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="OLCCWebAgents"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="Agents"
        Name =""
    End
End
