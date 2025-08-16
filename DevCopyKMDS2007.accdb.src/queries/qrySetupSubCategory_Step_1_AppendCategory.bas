Operation =3
Name ="category"
Option =0
Begin InputTables
End
Begin OutputColumns
    Alias ="ID1"
    Name ="ID"
    Expression ="1"
    Alias ="SalesDepartmentID1"
    Name ="SalesDepartmentID"
    Expression ="1"
    Alias ="Description1"
    Name ="Description"
    Expression ="\"Liquor\""
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="ID1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartmentID1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Description1"
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
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =651
End
