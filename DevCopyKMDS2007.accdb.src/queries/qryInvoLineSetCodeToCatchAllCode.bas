Operation =4
Option =0
Where ="(((InvoLine.Code)=[parCode]))"
Begin InputTables
    Name ="InvoLine"
End
Begin OutputColumns
    Name ="InvoLine.Code"
    Expression ="[parCodeCatchAll]"
End
Begin Parameters
    Name ="parCode"
    Flag =16
    Name ="parCodeCatchAll"
    Flag =16
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="InvoLine.Code"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-183
    Top =26
    Right =1042
    Bottom =703
    Left =-1
    Top =-1
    Right =1193
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
        Name ="InvoLine"
        Name =""
    End
End
