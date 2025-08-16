Operation =1
Option =0
Where ="(((Item.ID)=[parItemID]))"
Begin InputTables
    Name ="Item"
    Name ="qryItemOnReserve"
End
Begin OutputColumns
    Expression ="Item.ID"
    Alias ="Avail"
    Expression ="[Onhand]+[TransferredToday]-[SoldTodayCr1]-[SoldTodayCr2]-[SoldTodayCr3]-[SoldTo"
        "dayCr4]-[Reserve]"
    Expression ="Item.Onhand"
    Alias ="Reserve"
    Expression ="IIf(IsNull([OnReserve]),0,[OnReserve])"
    Expression ="Item.TransferredToday"
    Expression ="Item.SoldTodayCr1"
    Expression ="Item.SoldTodayCr2"
    Expression ="Item.SoldTodayCr3"
    Expression ="Item.SoldTodayCr4"
End
Begin Parameters
    Name ="parItemID"
    Flag =4
End
Begin Joins
    LeftTable ="Item"
    RightTable ="qryItemOnReserve"
    Expression ="Item.ID = qryItemOnReserve.ItemID"
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
        dbText "Name" ="Avail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Reserve"
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
    Bottom =276
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =201
        Bottom =270
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =239
        Top =6
        Right =362
        Bottom =131
        Top =0
        Name ="qryItemOnReserve"
        Name =""
    End
End
