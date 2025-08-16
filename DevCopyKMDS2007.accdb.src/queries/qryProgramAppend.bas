Operation =3
Name ="Programs"
Option =0
Where ="((([Switchboard Items].Command)=3) AND ((Programs.Name) Is Null))"
Begin InputTables
    Name ="Switchboard Items"
    Name ="Programs"
End
Begin OutputColumns
    Name ="Name"
    Expression ="[Switchboard Items].Argument"
End
Begin Joins
    LeftTable ="Switchboard Items"
    RightTable ="Programs"
    Expression ="[Switchboard Items].Argument = Programs.Name"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="[Switchboard Items].Argument"
        dbInteger "ColumnWidth" ="2745"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Switchboard Items].ItemText"
        dbInteger "ColumnWidth" ="3480"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Switchboard Items].Command"
        dbInteger "ColumnWidth" ="1065"
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
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="Switchboard Items"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =105
        Top =0
        Name ="Programs"
        Name =""
    End
End
