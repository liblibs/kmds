Operation =3
Name ="ItemSalePrice"
Option =0
Where ="(((([Starts]))>#5/31/2004#) And ((ConvITEMS.Ends)>\"#05/31/2004#\")) Or (((([Sta"
    "rts])) Is Not Null) And ((ConvITEMS.Ends) Is Not Null))"
Begin InputTables
    Name ="ConvITEMS"
End
Begin OutputColumns
    Alias ="Expr1"
    Name ="ItemID"
    Expression ="ConvITEMS.Code"
    Alias ="Expr2"
    Name ="SalePrice"
    Expression ="ConvITEMS.NewPrice"
    Alias ="st"
    Name ="StartDate"
    Expression ="([Starts])"
    Alias ="Expr3"
    Name ="EndDate"
    Expression ="ConvITEMS.Ends"
    Alias ="Expr4"
    Name ="OldPrice"
    Expression ="ConvITEMS.Price"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
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
        dbText "Name" ="st"
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
    Bottom =152
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =45
        Top =0
        Name ="ConvITEMS"
        Name =""
    End
End
