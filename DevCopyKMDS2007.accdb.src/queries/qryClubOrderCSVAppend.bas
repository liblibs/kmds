Operation =3
Name ="OrderLine"
Option =0
Where ="(((ClubOrderCSV.LineCount)>0) AND ((ClubOrderCSV.QuantityOrdered)<>0))"
Begin InputTables
    Name ="ClubOrderCSV"
    Name ="Item"
End
Begin OutputColumns
    Alias ="OrdheadID"
    Name ="OrdheadID"
    Expression ="[parOrderNumber]"
    Alias ="LineID"
    Name ="LineID"
    Expression ="ClubOrderCSV.LineCount"
    Name ="ItemId"
    Expression ="ClubOrderCSV.ItemId"
    Name ="QuantityOrdered"
    Expression ="ClubOrderCSV.QuantityOrdered"
    Name ="QuantityShipped"
    Expression ="ClubOrderCSV.QuantityOrdered"
    Name ="Price"
    Expression ="Item.Price"
    Name ="Cost"
    Expression ="Item.Cost"
End
Begin Joins
    LeftTable ="ClubOrderCSV"
    RightTable ="Item"
    Expression ="ClubOrderCSV.ItemId = Item.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="OrdheadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubOrderCSV.LineCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LineID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubOrderCSV.ItemId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubOrderCSV.QuantityOrdered"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Cost"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =52
    Top =-7
    Right =1268
    Bottom =675
    Left =-1
    Top =-1
    Right =1193
    Bottom =366
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =274
        Top =0
        Name ="ClubOrderCSV"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =318
        Top =0
        Name ="Item"
        Name =""
    End
End
