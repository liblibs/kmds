Operation =1
Option =0
Where ="(((Item.VendorID) Between [Forms]![PrintInventoryReport]![cboVendorIDBegin] And "
    "[Forms]![PrintInventoryReport]![cboVendorIDEnd]))"
Begin InputTables
    Name ="Item"
    Name ="Size"
    Name ="qryItemsSalePriceNow"
    Name ="qryItemOnReserve"
End
Begin OutputColumns
    Expression ="Item.VendorID"
    Expression ="Item.ID"
    Expression ="Item.StockNumber"
    Expression ="Item.Description"
    Alias ="SizeD"
    Expression ="Size.Description"
    Expression ="Item.Onhand"
    Expression ="Item.Cost"
    Alias ="TodaysPrice"
    Expression ="IIf(IsNull([SalePrice]),[Price],[SalePrice])"
    Expression ="qryItemOnReserve.OnReserve"
    Alias ="Avail"
    Expression ="[onhand]+[TransferredToday]-[SoldTodayCr1]-[SoldTodayCr2]-[SoldTodayCr3]-[SoldTo"
        "dayCr4]-(IIf(IsNull([OnReserve]),0,[OnReserve]))"
    Expression ="Item.OrderAt"
    Expression ="Item.OrderTo"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="Size"
    Expression ="Item.SizeID = Size.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="qryItemsSalePriceNow"
    Expression ="Item.ID = qryItemsSalePriceNow.ItemID"
    Flag =2
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
        dbText "Name" ="TodaysPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Avail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeD"
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
    Bottom =261
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =-1
        Top =13
        Right =165
        Bottom =262
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =359
        Top =232
        Right =455
        Bottom =346
        Top =0
        Name ="Size"
        Name =""
    End
    Begin
        Left =323
        Top =5
        Right =513
        Bottom =119
        Top =0
        Name ="qryItemsSalePriceNow"
        Name =""
    End
    Begin
        Left =330
        Top =125
        Right =481
        Bottom =209
        Top =0
        Name ="qryItemOnReserve"
        Name =""
    End
End
