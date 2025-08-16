Operation =1
Option =0
Where ="(((Item.Location) Between Forms!PrintInventoryReport!cboLocationBegin And Forms!"
    "PrintInventoryReport!cboLocationEnd) And ((Item.VendorID) Between Forms!PrintInv"
    "entoryReport!cboVendorIDBegin And Forms!PrintInventoryReport!cboVendorIDEnd) And"
    " ((Vendor.POMethodDefault)<>-1))"
Begin InputTables
    Name ="Item"
    Name ="Size"
    Name ="qryItemsSalePriceNow"
    Name ="qryItemOnReserve"
    Name ="Vendor"
End
Begin OutputColumns
    Expression ="Item.Location"
    Expression ="Item.ID"
    Expression ="Item.Description"
    Alias ="SizeD"
    Expression ="Size.Description"
    Expression ="Item.Onhand"
    Expression ="Item.Cost"
    Expression ="Item.Price"
    Expression ="qryItemsSalePriceNow.SalePrice"
    Alias ="TodaysPrice"
    Expression ="IIf(IsNull([SalePrice]),[Price],[SalePrice])"
    Expression ="qryItemOnReserve.OnReserve"
    Alias ="OnTheShelf"
    Expression ="[onhand]+[TransferredToday]-[SoldTodayCr1]-[SoldTodayCr2]-[SoldTodayCr3]-[SoldTo"
        "dayCr4]"
    Expression ="Item.VendorID"
    Alias ="VendorName"
    Expression ="Vendor.Name"
    Expression ="Vendor.POMethodDefault"
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
    LeftTable ="Item"
    RightTable ="Vendor"
    Expression ="Item.VendorID = Vendor.ID"
    Flag =1
End
Begin OrderBy
    Expression ="Item.Location"
    Flag =0
    Expression ="Item.Description"
    Flag =0
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
        dbText "Name" ="OnTheShelf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.POMethodDefault"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1422
    Bottom =677
    Left =-1
    Top =-1
    Right =1390
    Bottom =244
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
        Left =195
        Top =145
        Right =291
        Bottom =259
        Top =0
        Name ="Size"
        Name =""
    End
    Begin
        Left =369
        Top =13
        Right =559
        Bottom =127
        Top =0
        Name ="qryItemsSalePriceNow"
        Name =""
    End
    Begin
        Left =372
        Top =133
        Right =468
        Bottom =217
        Top =0
        Name ="qryItemOnReserve"
        Name =""
    End
    Begin
        Left =766
        Top =15
        Right =924
        Bottom =255
        Top =0
        Name ="Vendor"
        Name =""
    End
End
