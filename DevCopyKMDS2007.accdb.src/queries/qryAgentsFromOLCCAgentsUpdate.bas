Operation =4
Option =0
Begin InputTables
    Name ="OLCCAgents"
    Name ="Agents"
    Name ="OLCCAgentsXLS"
End
Begin OutputColumns
    Name ="Agents.FirstName"
    Expression ="[OLCCAgentsXLS]![F6]"
    Name ="Agents.LastName"
    Expression ="[OLCCAgentsXLS]![F2]"
    Name ="Agents.Address"
    Expression ="Left([OLCCAgentsXLS].[F9],(Len([OLCCAgentsXLS].[f9])-5))"
    Name ="Agents.Zip"
    Expression ="Right([OLCCAgentsXLS].[f9],5)"
    Name ="Agents.Phone"
    Expression ="[OLCCAgentsXLS]![F3]"
    Name ="Agents.ChangeDate"
    Expression ="Now()"
End
Begin Joins
    LeftTable ="OLCCAgents"
    RightTable ="Agents"
    Expression ="OLCCAgents.AgentNumber = Agents.AgentNumber"
    Flag =1
    LeftTable ="OLCCAgents"
    RightTable ="OLCCAgentsXLS"
    Expression ="OLCCAgents.AgentNumber = OLCCAgentsXLS.F1"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
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
        dbText "Name" ="Agents.FirstName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Agents.LastName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Agents.Address"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3450"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Agents.City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Agents.Zip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Agents.Phone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Agents.ChangeDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-302
    Top =253
    Right =923
    Bottom =930
    Left =-1
    Top =-1
    Right =1193
    Bottom =244
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =52
        Top =9
        Right =196
        Bottom =284
        Top =0
        Name ="OLCCAgents"
        Name =""
    End
    Begin
        Left =244
        Top =12
        Right =388
        Bottom =156
        Top =0
        Name ="Agents"
        Name =""
    End
    Begin
        Left =520
        Top =56
        Right =664
        Bottom =200
        Top =0
        Name ="OLCCAgentsXLS"
        Name =""
    End
End
