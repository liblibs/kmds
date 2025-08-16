Operation =1
Option =0
Where ="(((AgentTransferLines.QuantityOrdered)<>0))"
Begin InputTables
    Name ="AgentTransferLines"
    Name ="AgentTransferHead"
    Name ="Agents"
End
Begin OutputColumns
    Expression ="AgentTransferHead.ID"
    Expression ="AgentTransferHead.OrdDate"
    Expression ="Agents.AgentNumber"
    Expression ="Agents.FirstName"
    Expression ="Agents.LastName"
    Expression ="AgentTransferLines.ItemID"
    Expression ="AgentTransferLines.QuantityOrdered"
    Expression ="AgentTransferLines.Price"
    Expression ="AgentTransferLines.Cost"
    Expression ="AgentTransferLines.SetupDate"
    Expression ="AgentTransferLines.ChangeDate"
    Alias ="ExtPrice"
    Expression ="[QuantityShipped]*[Price]"
    Alias ="ExtCost"
    Expression ="[QuantityShipped]*[Cost]"
    Alias ="InOrOut"
    Expression ="IIf([Status]=\"I\",\"In\",\"Out\")"
End
Begin Joins
    LeftTable ="AgentTransferLines"
    RightTable ="AgentTransferHead"
    Expression ="AgentTransferLines.AgentTransferHeadID = AgentTransferHead.ID"
    Flag =1
    LeftTable ="AgentTransferHead"
    RightTable ="Agents"
    Expression ="AgentTransferHead.AgentID = Agents.AgentNumber"
    Flag =1
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
        dbText "Name" ="AgentTransferHead.ID"
        dbInteger "ColumnWidth" ="390"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AgentTransferHead.OrdDate"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Agents.FirstName"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AgentTransferLines.ItemID"
        dbInteger "ColumnWidth" ="810"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AgentTransferLines.ChangeDate"
        dbInteger "ColumnWidth" ="2175"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Agents.AgentNumber"
        dbInteger "ColumnWidth" ="975"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InOrOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Agents.LastName"
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
    Bottom =436
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =62
        Top =258
        Right =158
        Bottom =372
        Top =0
        Name ="AgentTransferLines"
        Name =""
    End
    Begin
        Left =254
        Top =200
        Right =423
        Bottom =314
        Top =0
        Name ="AgentTransferHead"
        Name =""
    End
    Begin
        Left =454
        Top =168
        Right =662
        Bottom =281
        Top =0
        Name ="Agents"
        Name =""
    End
End
