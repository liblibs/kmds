Operation =4
Option =0
Begin InputTables
    Name ="Sysfile"
End
Begin OutputColumns
    Name ="Sysfile.Cr1CheckOut"
    Expression ="0"
    Name ="Sysfile.Cr2CheckOut"
    Expression ="0"
    Name ="Sysfile.Cr3CheckOut"
    Expression ="0"
    Name ="Sysfile.Cr4CheckOut"
    Expression ="0"
    Name ="Sysfile.Cr5CheckOut"
    Expression ="0"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Sysfile.Cr1CheckOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Cr5CheckOut"
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
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="Sysfile"
        Name =""
    End
End
