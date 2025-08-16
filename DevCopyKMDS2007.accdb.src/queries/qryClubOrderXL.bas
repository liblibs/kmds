Operation =1
Option =0
Begin InputTables
    Name ="ClubOrderXL"
    Name ="Item"
End
Begin OutputColumns
    Alias ="LineID"
    Expression ="ClubOrderXL.line"
    Expression ="ClubOrderXL.ItemID"
    Alias ="QuantityOrdered"
    Expression ="IIf(IsNumeric([Qty]),[Qty],0)"
    Alias ="QuantityShipped"
    Expression ="IIf(IsNumeric([Qty]),[Qty],0)"
    Expression ="Item.Price"
    Expression ="Item.Cost"
    Expression ="ClubOrderXL.ClubID"
    Expression ="ClubOrderXL.ClubName"
    Expression ="ClubOrderXL.StoreName"
End
Begin Joins
    LeftTable ="ClubOrderXL"
    RightTable ="Item"
    Expression ="ClubOrderXL.ItemID = Item.ID"
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
    Bottom =371
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =208
        Top =0
        Name ="ClubOrderXL"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =306
        Top =0
        Name ="Item"
        Name =""
    End
End
