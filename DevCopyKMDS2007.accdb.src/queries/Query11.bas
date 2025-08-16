Operation =1
Option =0
Where ="(((Len([NewUPC1]))=12) And ((Left([NewUPC1],1))=0) And ((MeadowlandInventoryRepo"
    "rt13.NewItemID)>0))"
Begin InputTables
    Name ="MeadowlandInventoryReport13"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="MeadowlandInventoryReport13.NewUPC1"
    Alias ="lid"
    Expression ="Len([NewUPC1])"
    Alias ="le"
    Expression ="Left([NewUPC1],1)"
    Alias ="Expr2"
    Expression ="MeadowlandInventoryReport13.NewItemID"
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
        dbText "Name" ="MeadowlandInventoryReport13.NewUPC1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="le"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeadowlandInventoryReport13.NewItemID"
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
    Bottom =432
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="MeadowlandInventoryReport13"
        Name =""
    End
End
