Operation =1
Option =0
Where ="((([Switchboard Items].SwitchboardID)=1) AND (([Switchboard Items].Command)=1))"
Begin InputTables
    Name ="Switchboard Items"
End
Begin OutputColumns
    Expression ="[Switchboard Items].SwitchboardID"
    Expression ="[Switchboard Items].ItemNumber"
    Expression ="[Switchboard Items].ItemText"
    Expression ="[Switchboard Items].Command"
    Expression ="[Switchboard Items].Argument"
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
        Name ="Switchboard Items"
        Name =""
    End
End
