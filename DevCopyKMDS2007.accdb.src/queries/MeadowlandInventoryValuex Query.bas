Operation =1
Option =0
Begin InputTables
    Name ="MeadowlandInventoryValuex"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="MeadowlandInventoryValuex.[ItemNo]"
    Alias ="Expr2"
    Expression ="MeadowlandInventoryValuex.[NewItem]"
    Alias ="Expr3"
    Expression ="MeadowlandInventoryValuex.[SalesDpt]"
    Alias ="Expr4"
    Expression ="MeadowlandInventoryValuex.[description]"
    Alias ="Expr5"
    Expression ="MeadowlandInventoryValuex.[Cost]"
    Alias ="Expr6"
    Expression ="MeadowlandInventoryValuex.[Price]"
    Alias ="Expr7"
    Expression ="MeadowlandInventoryValuex.[Onhand]"
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
        dbText "Name" ="MeadowlandInventoryValuex.[ItemNo]"
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
    Bottom =322
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="MeadowlandInventoryValuex"
        Name =""
    End
End
