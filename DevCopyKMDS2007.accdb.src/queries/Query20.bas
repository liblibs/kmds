Operation =1
Option =0
Where ="(((Item.ID) Is Null))"
Begin InputTables
    Name ="Item"
    Name ="Inv20091031"
End
Begin OutputColumns
    Expression ="Inv20091031.ItemID"
    Expression ="Inv20091031.Onhand"
    Expression ="Item.ID"
End
Begin Joins
    LeftTable ="Inv20091031"
    RightTable ="Item"
    Expression ="Inv20091031.ItemID = Item.ID"
    Flag =2
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
        dbText "Name" ="Inv20091031.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Inv20091031.Onhand"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
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
    Bottom =321
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Inv20091031"
        Name =""
    End
End
