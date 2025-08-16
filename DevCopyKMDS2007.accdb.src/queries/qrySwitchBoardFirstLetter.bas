Operation =1
Option =0
Where ="((([Switchboard Items].Command)=0))"
Begin InputTables
    Name ="Switchboard Items"
End
Begin OutputColumns
    Expression ="[Switchboard Items].SwitchboardID"
    Expression ="[Switchboard Items].ItemNumber"
    Expression ="[Switchboard Items].ItemText"
    Expression ="[Switchboard Items].Command"
    Expression ="[Switchboard Items].Argument"
    Alias ="FirstLetter"
    Expression ="Left([ItemText],1)"
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
        dbInteger "ColumnWidth" ="3045"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstLetter"
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
    Bottom =244
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =192
        Bottom =196
        Top =0
        Name ="Switchboard Items"
        Name =""
    End
End
