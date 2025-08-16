Operation =1
Option =0
Where ="((([ConvITEMS1].[SalesDept])=1))"
Begin InputTables
    Name ="ConvITEMS1"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="ConvITEMS1.Code"
    Alias ="Expr2"
    Expression ="ConvITEMS1.Price"
    Alias ="Expr3"
    Expression ="ConvITEMS1.NewPrice"
    Alias ="Expr4"
    Expression ="ConvITEMS1.Starts"
    Alias ="Expr5"
    Expression ="ConvITEMS1.Ends"
    Alias ="csPrice"
    Expression ="IIf(Not IsNull([Starts]) And CDate([Starts])>#2/28/2005#,[NewPrice],[Price])"
End
Begin OrderBy
    Expression ="ConvITEMS1.Starts"
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
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="csPrice"
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
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="ConvITEMS1"
        Name =""
    End
End
