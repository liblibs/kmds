Operation =1
Option =0
Where ="(((Agents.AgentNumber)=[parAgentID]))"
Begin InputTables
    Name ="Agents"
End
Begin OutputColumns
    Expression ="Agents.AgentNumber"
    Expression ="Agents.OrderCount"
End
Begin Parameters
    Name ="parAgentID"
    Flag =4
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
    ColumnsShown =539
    Begin
        Left =71
        Top =5
        Right =167
        Bottom =119
        Top =0
        Name ="Agents"
        Name =""
    End
End
