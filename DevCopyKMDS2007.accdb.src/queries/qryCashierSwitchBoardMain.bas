Operation =1
Option =0
Where ="((([Switchboard Items].SwitchboardID)=1) AND (([Switchboard Items].Command)=1))"
Begin InputTables
    Name ="Switchboard Items"
End
Begin OutputColumns
    Expression ="[Switchboard Items].ItemText"
    Expression ="[Switchboard Items].SwitchboardID"
    Expression ="[Switchboard Items].ItemNumber"
    Expression ="[Switchboard Items].Command"
    Alias ="Arg"
    Expression ="CInt([Argument])"
End
Begin OrderBy
    Expression ="[Switchboard Items].ItemText"
    Flag =0
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
        dbText "Name" ="[Switchboard Items].ItemText"
        dbInteger "ColumnWidth" ="2580"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Arg"
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
    Bottom =190
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =174
        Bottom =150
        Top =0
        Name ="Switchboard Items"
        Name =""
    End
End
