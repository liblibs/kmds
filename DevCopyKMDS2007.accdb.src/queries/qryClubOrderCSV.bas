Operation =1
Option =0
Where ="(((ClubOrderCSV.LineCount)>0))"
Begin InputTables
    Name ="Item"
    Name ="ClubOrderCSV"
    Name ="Club"
End
Begin OutputColumns
    Alias ="LineID"
    Expression ="ClubOrderCSV.LineCount"
    Expression ="ClubOrderCSV.ItemID"
    Expression ="Item.Price"
    Expression ="Item.Cost"
    Expression ="ClubOrderCSV.ClubID"
    Alias ="ClubName"
    Expression ="Club.Name"
    Expression ="ClubOrderCSV.QuantityOrdered"
    Alias ="QuantityShipped"
    Expression ="ClubOrderCSV.QuantityOrdered"
End
Begin Joins
    LeftTable ="ClubOrderCSV"
    RightTable ="Item"
    Expression ="ClubOrderCSV.ItemId = Item.ID"
    Flag =1
    LeftTable ="ClubOrderCSV"
    RightTable ="Club"
    Expression ="ClubOrderCSV.ClubID = Club.ID"
    Flag =1
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
        dbText "Name" ="LineID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuantityShipped"
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
    Begin
        dbText "Name" ="ClubOrderCSV.ClubID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubOrderCSV.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubOrderCSV.QuantityOrdered"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubName"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =38
    Top =86
    Right =974
    Bottom =877
    Left =-1
    Top =-1
    Right =904
    Bottom =286
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =115
        Top =9
        Right =259
        Bottom =303
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =607
        Bottom =314
        Top =0
        Name ="ClubOrderCSV"
        Name =""
    End
    Begin
        Left =637
        Top =28
        Right =781
        Bottom =172
        Top =0
        Name ="Club"
        Name =""
    End
End
