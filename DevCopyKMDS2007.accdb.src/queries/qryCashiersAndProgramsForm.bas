Operation =1
Option =0
Where ="((([Switchboard Items].Command)=3))"
Begin InputTables
    Name ="Switchboard Items"
    Name ="qrySwitchBoardFirstLetter"
End
Begin OutputColumns
    Alias ="MenuOption"
    Expression ="[FirstLetter] & [Switchboard Items].ItemText"
    Expression ="[Switchboard Items].ItemText"
    Expression ="[Switchboard Items].Command"
    Expression ="[Switchboard Items].Argument"
End
Begin Joins
    LeftTable ="Switchboard Items"
    RightTable ="qrySwitchBoardFirstLetter"
    Expression ="[Switchboard Items].SwitchboardID = qrySwitchBoardFirstLetter.SwitchboardID"
    Flag =1
End
Begin OrderBy
    Expression ="[FirstLetter] & [Switchboard Items].ItemText"
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
        dbText "Name" ="MenuOption"
        dbInteger "ColumnWidth" ="2895"
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
    Bottom =195
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =359
        Top =42
        Right =562
        Bottom =186
        Top =0
        Name ="Switchboard Items"
        Name =""
    End
    Begin
        Left =38
        Top =48
        Right =206
        Bottom =162
        Top =0
        Name ="qrySwitchBoardFirstLetter"
        Name =""
    End
End
