Operation =3
Name ="ClubOrderGuideTemp"
Option =0
Begin InputTables
End
Begin OutputColumns
    Alias ="Club"
    Name ="ClubID"
    Expression ="[parClubID]"
    Alias ="Line"
    Name ="LineID"
    Expression ="[parLineID]"
End
Begin Parameters
    Name ="parClubID"
    Flag =4
    Name ="parLineID"
    Flag =4
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
        dbText "Name" ="Club"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Line"
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
