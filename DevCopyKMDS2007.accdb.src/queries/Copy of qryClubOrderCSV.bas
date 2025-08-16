Operation =1
Option =0
Begin InputTables
    Name ="Item"
    Name ="ClubOrderCSV"
End
Begin OutputColumns
    Alias ="LineID"
    Expression ="ClubOrderCSV.line"
    Expression ="ClubOrderCSV.ItemID"
    Alias ="QuantityOrdered"
    Expression ="IIf(IsNumeric([Qty]),[Qty],0)"
    Alias ="QuantityShipped"
    Expression ="IIf(IsNumeric([Qty]),[Qty],0)"
    Expression ="Item.Price"
    Expression ="Item.Cost"
    Expression ="ClubOrderCSV.ClubID"
    Expression ="ClubOrderCSV.ClubName"
    Expression ="ClubOrderCSV.StoreName"
    Expression ="ClubOrderCSV.QuantityOrdered"
    Alias ="QuantityShipped"
    Expression ="ClubOrderCSV.QuantityOrdered"
End
Begin Joins
    LeftTable ="ClubOrderCSV"
    RightTable ="Item"
    Expression ="ClubOrderCSV.ItemId = Item.ID"
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
        dbText "Name" ="ClubOrderXL.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuantityOrdered"
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
        dbText "Name" ="ClubOrderXL.ClubID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubOrderXL.ClubName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubOrderXL.StoreName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubOrderCSV.ClubID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubOrderCSV.ClubName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubOrderCSV.StoreName"
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
End
Begin
    State =0
    Left =180
    Top =106
    Right =1405
    Bottom =783
    Left =-1
    Top =-1
    Right =1193
    Bottom =354
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =306
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =292
        Top =0
        Name ="ClubOrderCSV"
        Name =""
    End
End
