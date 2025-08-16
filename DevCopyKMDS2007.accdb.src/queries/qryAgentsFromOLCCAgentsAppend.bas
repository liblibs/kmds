Operation =3
Name ="Agents"
Option =0
Where ="(((Agents.AgentNumber) Is Null))"
Begin InputTables
    Name ="Agents"
    Name ="OLCCAgentsXLS"
End
Begin OutputColumns
    Name ="AgentNumber"
    Expression ="OLCCAgentsXLS.F1"
    Name ="LastName"
    Expression ="OLCCAgentsXLS.F2"
    Name ="FirstName"
    Expression ="OLCCAgentsXLS.F6"
    Name ="Phone"
    Expression ="OLCCAgentsXLS.F3"
    Alias ="addy"
    Name ="Address"
    Expression ="Left([F9],(Len([f9])-5))"
    Name ="City"
    Expression ="OLCCAgentsXLS.F2"
    Alias ="zp"
    Name ="Zip"
    Expression ="Right([f9],5)"
    Alias ="state"
    Name ="ST"
    Expression ="\"OR\""
End
Begin Joins
    LeftTable ="Agents"
    RightTable ="OLCCAgentsXLS"
    Expression ="Agents.AgentNumber = OLCCAgentsXLS.F1"
    Flag =3
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="OLCCAgents.AgentNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCAgents.AgencyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCAgents.FirstName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="fnam"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCAgents.Phone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCAgents.Add1"
        dbInteger "ColumnWidth" ="1845"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCAgents.Add2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCAgents.City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCAgents.ZipCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="state"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Agents.AgentNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCAgentsXLS.AgentNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCAgentsXLS.F1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Agents.AgencyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCAgentsXLS.AgencyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCAgentsXLS.F2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCAgentsXLS.F6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCAgentsXLS.Phone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCAgentsXLS.F3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCAgentsXLS.Add1"
        dbInteger "ColumnWidth" ="1845"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCAgentsXLS.F9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="add"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCAgentsXLS.City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="zp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="addy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1001"
        dbInteger "ColumnWidth" ="2415"
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
    Right =1193
    Bottom =244
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =244
        Top =12
        Right =388
        Bottom =245
        Top =0
        Name ="Agents"
        Name =""
    End
    Begin
        Left =630
        Top =64
        Right =774
        Bottom =208
        Top =0
        Name ="OLCCAgentsXLS"
        Name =""
    End
End
