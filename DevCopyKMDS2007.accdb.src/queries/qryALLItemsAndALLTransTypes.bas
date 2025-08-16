Operation =1
Option =0
Where ="(((TranTypes.Code)<>\"T\" And (TranTypes.Code)<>\"R\" And (TranTypes.Code)<>\"D\""
    "))"
Begin InputTables
    Name ="Item"
    Name ="TranTypes"
End
Begin OutputColumns
    Expression ="Item.ID"
    Expression ="TranTypes.Description"
    Expression ="TranTypes.Code"
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
        dbText "Name" ="TranTypes.Description"
        dbInteger "ColumnWidth" ="1590"
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
        Name ="Item"
        Name =""
    End
    Begin
        Left =151
        Top =9
        Right =247
        Bottom =108
        Top =0
        Name ="TranTypes"
        Name =""
    End
End
