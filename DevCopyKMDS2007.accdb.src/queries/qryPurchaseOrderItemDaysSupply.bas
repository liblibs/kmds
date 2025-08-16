Operation =1
Option =0
Where ="(((qryPurchaseOrderItemDaysSupplyADS.Sold)<>0) AND ((Item.Onhand) Is Not Null))"
Begin InputTables
    Name ="qryPurchaseOrderItemDaysSupplyADS"
    Name ="Item"
    Name ="Vendor"
    Name ="qryItemOnReserve"
    Name ="qryItemOnOrder"
    Name ="qryItemsSalePriceNow"
    Name ="Size"
End
Begin OutputColumns
    Expression ="qryPurchaseOrderItemDaysSupplyADS.VendorID"
    Alias ="VendorName"
    Expression ="Vendor.Name"
    Expression ="Item.SalesDepartmentID"
    Expression ="qryPurchaseOrderItemDaysSupplyADS.ItemId"
    Expression ="Item.Description"
    Alias ="SizeDesc"
    Expression ="Size.Description"
    Expression ="qryPurchaseOrderItemDaysSupplyADS.Sold"
    Expression ="qryPurchaseOrderItemDaysSupplyADS.SoldDays"
    Alias ="AverageDailySales"
    Expression ="([Sold]/[SoldDays])"
    Alias ="SupplyDays"
    Expression ="[parSupplyDays]"
    Alias ="SDOrderTo"
    Expression ="CInt(([Sold]/[SoldDays])*[parSupplyDays])"
    Expression ="Item.Onhand"
    Alias ="Ordered"
    Expression ="IIf(IsNull([OnOrder]),0,[OnOrder])"
    Alias ="Reserved"
    Expression ="IIf(IsNull([OnReserve]),0,[OnReserve])"
    Alias ="Avail"
    Expression ="([Onhand]+[Ordered]-[Reserved]-[SoldTodayCr1]-[SoldTodayCr2]-[SoldTodayCr3]-[Sol"
        "dTodayCr4]+[TransferredToday])"
    Alias ="SOQRaw"
    Expression ="((CInt(([Sold]/[SoldDays])*[parSupplyDays]))-[Avail])"
    Alias ="OrderMultiple"
    Expression ="Item.UnitsPerOrder"
    Alias ="SOQ"
    Expression ="CInt((([SOQRaw])/[UnitsPerOrder]))*[UnitsPerOrder]"
    Expression ="qryPurchaseOrderItemDaysSupplyADS.DateStart"
    Expression ="qryPurchaseOrderItemDaysSupplyADS.DateEnd"
    Expression ="Item.UnitsPerCase"
    Alias ="Repack"
    Expression ="IIf([soq] Mod [UnitsPerCase]<>0,True,False)"
    Expression ="Item.Price"
    Alias ="SPrice"
    Expression ="IIf(IsNull([SalePrice]),Item.Price,[SalePrice])"
    Alias ="SPriceStar"
    Expression ="IIf(IsNull([SalePrice]),\" \",\"*\")"
    Alias ="ExtPrice"
    Expression ="[SOQ]*[SPrice]"
    Expression ="Item.Cost"
    Alias ="ExtCost"
    Expression ="[SOQ]*[Cost]"
    Alias ="ItemOrderAt"
    Expression ="Item.OrderAt"
    Expression ="Item.OrderTo"
End
Begin Parameters
    Name ="parSupplyDays"
    Flag =3
End
Begin Joins
    LeftTable ="Item"
    RightTable ="Vendor"
    Expression ="Item.VendorID = Vendor.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="qryItemOnReserve"
    Expression ="Item.ID = qryItemOnReserve.ItemID"
    Flag =2
    LeftTable ="Item"
    RightTable ="qryItemOnOrder"
    Expression ="Item.ID = qryItemOnOrder.ItemID"
    Flag =2
    LeftTable ="Item"
    RightTable ="qryItemsSalePriceNow"
    Expression ="Item.ID = qryItemsSalePriceNow.ItemID"
    Flag =2
    LeftTable ="Item"
    RightTable ="Size"
    Expression ="Item.SizeID = Size.ID"
    Flag =1
    LeftTable ="qryPurchaseOrderItemDaysSupplyADS"
    RightTable ="Item"
    Expression ="qryPurchaseOrderItemDaysSupplyADS.ItemId = Item.ID"
    Flag =1
End
Begin OrderBy
    Expression ="Item.SalesDepartmentID"
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
        dbText "Name" ="Item.Onhand"
        dbInteger "ColumnWidth" ="855"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbInteger "ColumnWidth" ="2700"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SOQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemDaysSupplyADS.VendorID"
        dbInteger "ColumnWidth" ="975"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SOQRaw"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ordered"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Reserved"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Avail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Repack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SPriceStar"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AverageDailySales"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SupplyDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SDOrderTo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderMultiple"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemOrderAt"
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
    Bottom =374
    Left =0
    Top =20
    ColumnsShown =539
    Begin
        Left =75
        Top =56
        Right =240
        Bottom =200
        Top =0
        Name ="qryPurchaseOrderItemDaysSupplyADS"
        Name =""
    End
    Begin
        Left =291
        Top =60
        Right =387
        Bottom =354
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =549
        Top =327
        Right =680
        Bottom =411
        Top =0
        Name ="Vendor"
        Name =""
    End
    Begin
        Left =554
        Top =13
        Right =685
        Bottom =97
        Top =0
        Name ="qryItemOnReserve"
        Name =""
    End
    Begin
        Left =551
        Top =136
        Right =679
        Bottom =220
        Top =0
        Name ="qryItemOnOrder"
        Name =""
    End
    Begin
        Left =726
        Top =60
        Right =892
        Bottom =174
        Top =0
        Name ="qryItemsSalePriceNow"
        Name =""
    End
    Begin
        Left =552
        Top =227
        Right =683
        Bottom =311
        Top =0
        Name ="Size"
        Name =""
    End
End
