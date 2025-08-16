Operation =1
Option =0
Begin InputTables
    Name ="qryUPCItemMax"
End
Begin OutputColumns
    Expression ="qryUPCItemMax.ItemID"
    Alias ="UPCG"
    Expression ="UPCGlyph([UPCCodeMax])"
    Expression ="qryUPCItemMax.UPCCodeMax"
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
        dbText "Name" ="UPCG"
        dbInteger "ColumnWidth" ="1740"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =214
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =72
        Top =8
        Right =293
        Bottom =203
        Top =0
        Name ="qryUPCItemMax"
        Name =""
    End
End
