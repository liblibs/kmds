Operation =1
Option =0
Where ="(((AgentTransferHead.ID)=[Forms]![AgentsTransfersChoices]![txtAgentTransferID]))"
Begin InputTables
    Name ="Agents"
    Name ="AgentTransferHead"
    Name ="Item"
    Name ="AgentTransferLines"
    Name ="qryUPCItemMax"
    Name ="Size"
End
Begin OutputColumns
    Expression ="AgentTransferHead.ID"
    Expression ="AgentTransferHead.AgentID"
    Expression ="AgentTransferHead.Order"
    Expression ="AgentTransferHead.OrdDate"
    Alias ="AgentName"
    Expression ="Agents.LastName"
    Expression ="Agents.Address"
    Alias ="CitySTZ"
    Expression ="[City] & \", \" & [ST] & \" \" & [Zip]"
    Expression ="Agents.Phone"
    Expression ="AgentTransferLines.LineID"
    Expression ="AgentTransferLines.ItemID"
    Expression ="Item.Description"
    Expression ="Item.SizeID"
    Expression ="AgentTransferLines.QuantityShipped"
    Expression ="AgentTransferLines.Price"
    Expression ="qryUPCItemMax.UPCCodeMax"
    Alias ="ExtPrice"
    Expression ="[QuantityShipped]*AgentTransferLines.Price"
    Expression ="AgentTransferHead.Status"
    Alias ="SizeDesc"
    Expression ="IIf(IsNull(Size.Description),\"N/A\",Size!Description)"
    Expression ="Item.Cost"
    Alias ="ExtCost"
    Expression ="[QuantityShipped]*Item.Cost"
    Expression ="AgentTransferLines.ClaimCode"
End
Begin Joins
    LeftTable ="Agents"
    RightTable ="AgentTransferHead"
    Expression ="Agents.AgentNumber = AgentTransferHead.AgentID"
    Flag =1
    LeftTable ="Item"
    RightTable ="AgentTransferLines"
    Expression ="Item.ID = AgentTransferLines.ItemID"
    Flag =1
    LeftTable ="Item"
    RightTable ="qryUPCItemMax"
    Expression ="Item.ID = qryUPCItemMax.ItemID"
    Flag =2
    LeftTable ="AgentTransferHead"
    RightTable ="AgentTransferLines"
    Expression ="AgentTransferHead.ID = AgentTransferLines.AgentTransferHeadID"
    Flag =1
    LeftTable ="Item"
    RightTable ="Size"
    Expression ="Item.SizeID = Size.ID"
    Flag =2
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
        dbText "Name" ="AgentName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CitySTZ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtCost"
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
    Bottom =202
    Left =0
    Top =20
    ColumnsShown =539
    Begin
        Left =38
        Top =-14
        Right =134
        Bottom =100
        Top =0
        Name ="Agents"
        Name =""
    End
    Begin
        Left =172
        Top =-14
        Right =268
        Bottom =100
        Top =0
        Name ="AgentTransferHead"
        Name =""
    End
    Begin
        Left =534
        Top =-7
        Right =630
        Bottom =197
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =320
        Top =-6
        Right =490
        Bottom =108
        Top =0
        Name ="AgentTransferLines"
        Name =""
    End
    Begin
        Left =668
        Top =-14
        Right =764
        Bottom =85
        Top =0
        Name ="qryUPCItemMax"
        Name =""
    End
    Begin
        Left =668
        Top =88
        Right =764
        Bottom =202
        Top =0
        Name ="Size"
        Name =""
    End
End
