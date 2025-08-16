Operation =4
Option =0
Where ="(((OrdHead.ID)=[parOrdheadID]))"
Begin InputTables
    Name ="OrdHead"
    Name ="Sysfile"
End
Begin OutputColumns
    Name ="OrdHead.InvoiceNo"
    Expression ="Sysfile.InvoiceCounter"
    Name ="OrdHead.InvDate"
    Expression ="Now()"
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
        Left =47
        Top =7
        Right =143
        Bottom =121
        Top =0
        Name ="OrdHead"
        Name =""
    End
    Begin
        Left =183
        Top =5
        Right =279
        Bottom =119
        Top =0
        Name ="Sysfile"
        Name =""
    End
End
