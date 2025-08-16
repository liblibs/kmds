Operation =1
Option =0
Where ="((([ItemCorv1].[appendSKU])<>\"N\"))"
Begin InputTables
    Name ="ItemCorv1"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="ItemCorv1.sku"
    Alias ="Expr2"
    Expression ="ItemCorv1.[case q]"
    Alias ="Expr3"
    Expression ="ItemCorv1.appendSKU"
    Alias ="PackSKU"
    Expression ="[SKU]-1"
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
        dbText "Name" ="PackSKU"
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
        Left =172
        Top =6
        Right =268
        Bottom =45
        Top =0
        Name ="ItemCorv1"
        Name =""
    End
End
