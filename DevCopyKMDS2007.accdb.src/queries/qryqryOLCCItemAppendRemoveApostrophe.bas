Operation =4
Option =0
Where ="(((Replace([Description],\"'\",\"\"))<>[Description]))"
Begin InputTables
    Name ="Item_ICExpansion"
End
Begin OutputColumns
    Name ="Item_ICExpansion.Description"
    Expression ="Replace([Description],\"'\",\"\")"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "OrderBy" ="[Query52].[Description]"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="Item_ICExpansion.Description"
        dbInteger "ColumnWidth" ="4305"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewDesc"
        dbInteger "ColumnWidth" ="3225"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1075
    Bottom =846
    Left =-1
    Top =-1
    Right =1043
    Bottom =505
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =281
        Bottom =312
        Top =0
        Name ="Item_ICExpansion"
        Name =""
    End
End
