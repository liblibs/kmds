Operation =3
Name ="SubCategory"
Option =0
Begin InputTables
    Name ="OLCC Categories"
End
Begin OutputColumns
    Name ="Description"
    Expression ="[OLCC Categories].Category"
    Alias ="Cat"
    Name ="CategoryID"
    Expression ="1"
End
Begin Groups
    Expression ="[OLCC Categories].Category"
    GroupLevel =0
    Expression ="1"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="[OLCC Categories].Category"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cat"
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
    Bottom =295
    Left =0
    Top =0
    ColumnsShown =655
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="OLCC Categories"
        Name =""
    End
End
