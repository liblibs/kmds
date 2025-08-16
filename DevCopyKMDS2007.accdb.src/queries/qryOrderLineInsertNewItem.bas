Operation =3
Name ="OrderLine"
Option =0
Begin InputTables
End
Begin OutputColumns
    Alias ="xLineID"
    Name ="LineID"
    Expression ="[parLineID]"
    Alias ="xOrdheadID"
    Name ="OrdheadID"
    Expression ="[parOrdheadID]"
    Alias ="xItemId"
    Name ="ItemID"
    Expression ="[parItemID]"
End
Begin Parameters
    Name ="parOrdheadID"
    Flag =4
    Name ="parLineID"
    Flag =3
    Name ="parItemID"
    Flag =7
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="qryOrderLineByLineIDDescending.LineID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LineID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryOrderLineByLineIDDescending.OrdheadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdheadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xLineID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xOrdheadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xItemId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-354
    Top =150
    Right =871
    Bottom =827
    Left =-1
    Top =-1
    Right =1193
    Bottom =238
    Left =0
    Top =0
    ColumnsShown =651
End
