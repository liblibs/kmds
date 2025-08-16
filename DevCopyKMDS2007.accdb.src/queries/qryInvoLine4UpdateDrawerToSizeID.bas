Operation =4
Option =0
Where ="(((IsNumeric([code]))=True))"
Begin InputTables
    Name ="InvoLine4"
End
Begin OutputColumns
    Alias ="Expr1"
    Name ="InvoLine4.Drawer"
    Expression ="Right([code],1)"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
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
        Right =210
        Bottom =120
        Top =0
        Name ="InvoLine4"
        Name =""
    End
End
