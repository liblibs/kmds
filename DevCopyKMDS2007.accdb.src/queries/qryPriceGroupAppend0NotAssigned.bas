Operation =3
Name ="PriceGroup"
Option =0
Begin InputTables
End
Begin OutputColumns
    Alias ="ID"
    Name ="ID"
    Expression ="0"
    Alias ="DES"
    Name ="Description"
    Expression ="\" Not Assigned\""
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DES"
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
End
