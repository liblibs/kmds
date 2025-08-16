Operation =1
Option =0
Where ="(((AgentTransferHead.Status)=\"I\"))"
Begin InputTables
    Name ="qryAgentTransferFromFTPHead"
    Name ="AgentTransferHead"
End
Begin OutputColumns
    Expression ="AgentTransferHead.ID"
    Expression ="AgentTransferHead.Status"
End
Begin Joins
    LeftTable ="qryAgentTransferFromFTPHead"
    RightTable ="AgentTransferHead"
    Expression ="qryAgentTransferFromFTPHead.AgentNumber = AgentTransferHead.AgentID"
    Flag =1
    LeftTable ="qryAgentTransferFromFTPHead"
    RightTable ="AgentTransferHead"
    Expression ="qryAgentTransferFromFTPHead.Order = AgentTransferHead.Order"
    Flag =1
    LeftTable ="qryAgentTransferFromFTPHead"
    RightTable ="AgentTransferHead"
    Expression ="qryAgentTransferFromFTPHead.OrdDate = AgentTransferHead.OrdDate"
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
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AgentTransferHead.Status"
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
    Bottom =220
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =105
        Top =0
        Name ="qryAgentTransferFromFTPHead"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =564
        Bottom =234
        Top =0
        Name ="AgentTransferHead"
        Name =""
    End
End
