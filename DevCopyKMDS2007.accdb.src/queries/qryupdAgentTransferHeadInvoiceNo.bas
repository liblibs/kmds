Operation =4
Option =0
Where ="(((AgentTransferHead.ID)=[parAgentTransferHead]))"
Begin InputTables
    Name ="Sysfile"
    Name ="AgentTransferHead"
End
Begin OutputColumns
    Name ="AgentTransferHead.InvoiceNo"
    Expression ="Sysfile.InvoiceCounter"
    Name ="AgentTransferHead.InvDate"
    Expression ="Now()"
End
Begin Parameters
    Name ="parAgentTransferHead"
    Flag =4
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
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
    Left =47
    Top =0
    ColumnsShown =579
    Begin
        Left =95
        Top =6
        Right =191
        Bottom =120
        Top =0
        Name ="Sysfile"
        Name =""
    End
    Begin
        Left =317
        Top =6
        Right =413
        Bottom =120
        Top =0
        Name ="AgentTransferHead"
        Name =""
    End
End
