Operation =2
Name ="PurchaseOrderItemWADS"
Option =0
Where ="(((Vendor.ID)=[parVendorID]))"
Begin InputTables
    Name ="Vendor"
    Name ="Item"
    Name ="qryItemOnOrder"
    Name ="qryItemOnReserve"
    Name ="qryItemsSalePriceNow"
    Name ="AverageDaySupplyLevels"
End
Begin OutputColumns
    Alias ="VendorID"
    Expression ="Vendor.ID"
    Expression ="Vendor.Name"
    Alias ="ItemID"
    Expression ="Item.ID"
    Expression ="Item.Description"
    Expression ="Item.SizeID"
    Expression ="Item.Onhand"
    Alias ="Ordered"
    Expression ="IIf(IsNull([OnOrder]),0,[OnOrder])"
    Alias ="Reserved"
    Expression ="IIf(IsNull([OnReserve]),0,[OnReserve])"
    Alias ="Avail"
    Expression ="[Onhand]+[Ordered]-[Reserved]-[SoldTodayCr1]-[SoldTodayCr2]-[SoldTodayCr3]-[Sold"
        "TodayCr4]+[TransferredToday]"
    Expression ="Item.AverageDailySales"
    Expression ="AverageDaySupplyLevels.DaysSupply"
    Alias ="OrderMax"
    Expression ="CLng(IIf(IsNull([AverageDailySales]),0,[AverageDailySales])*[DaysSupply])"
    Alias ="SOQRaw"
    Expression ="([OrderMax]-[Avail]+0.01)"
    Alias ="OrderMultiple"
    Expression ="IIf([UnitsPerOrder]<1,1,[UnitsPerOrder])"
    Alias ="SOQ"
    Expression ="(CInt([SOQRaw]/[OrderMultiple])*[OrderMultiple])"
    Expression ="Item.UnitsPerCase"
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
    Expression ="Item.SalesDepartmentID"
    Alias ="Repack"
    Expression ="IIf([soq] Mod [UnitsPerCase]<>0,True,False)"
End
Begin Joins
    LeftTable ="Vendor"
    RightTable ="Item"
    Expression ="Vendor.ID = Item.VendorID"
    Flag =1
    LeftTable ="Item"
    RightTable ="qryItemOnOrder"
    Expression ="Item.ID = qryItemOnOrder.ItemID"
    Flag =2
    LeftTable ="Item"
    RightTable ="qryItemOnReserve"
    Expression ="Item.ID = qryItemOnReserve.ItemID"
    Flag =2
    LeftTable ="Item"
    RightTable ="qryItemsSalePriceNow"
    Expression ="Item.ID = qryItemsSalePriceNow.ItemID"
    Flag =2
    LeftTable ="Item"
    RightTable ="AverageDaySupplyLevels"
    Expression ="Item.DaysSupplyGroupID = AverageDaySupplyLevels.ID"
    Flag =1
End
Begin OrderBy
    Expression ="Item.ID"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="Vendor.Name"
        dbInteger "ColumnWidth" ="765"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ordered"
        dbInteger "ColumnOrder" ="4"
        dbInteger "ColumnWidth" ="990"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Reserved"
        dbInteger "ColumnOrder" ="16"
        dbInteger "ColumnWidth" ="990"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemID"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbInteger "ColumnOrder" ="5"
        dbInteger "ColumnWidth" ="1020"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SizeID"
        dbInteger "ColumnOrder" ="6"
        dbInteger "ColumnWidth" ="750"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Onhand"
        dbInteger "ColumnOrder" ="7"
        dbInteger "ColumnWidth" ="0"
        dbBoolean "ColumnHidden" ="-1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.UnitsPerCase"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Avail"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SOQRaw"
        dbInteger "ColumnOrder" ="12"
        dbInteger "ColumnWidth" ="1305"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SOQ"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderMultiple"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SPriceStar"
        dbInteger "ColumnWidth" ="1095"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="19"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SPrice"
        dbInteger "ColumnWidth" ="885"
        dbInteger "ColumnOrder" ="18"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtPrice"
        dbInteger "ColumnOrder" ="20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtCost"
        dbInteger "ColumnOrder" ="22"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Vendor].[Name]"
        dbInteger "ColumnWidth" ="225"
        dbInteger "ColumnOrder" ="9"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Item].[Description]"
        dbInteger "ColumnWidth" ="1590"
        dbInteger "ColumnOrder" ="10"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Item].[SizeID]"
        dbInteger "ColumnWidth" ="0"
        dbInteger "ColumnOrder" ="11"
        dbBoolean "ColumnHidden" ="-1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Item].[Onhand]"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderMax"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Price"
        dbInteger "ColumnWidth" ="885"
        dbInteger "ColumnOrder" ="17"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Cost"
        dbInteger "ColumnOrder" ="21"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbInteger "ColumnOrder" ="23"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Repack"
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
    Bottom =342
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =27
        Top =37
        Right =123
        Bottom =136
        Top =0
        Name ="Vendor"
        Name =""
    End
    Begin
        Left =177
        Top =69
        Right =361
        Bottom =348
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =428
        Top =8
        Right =615
        Bottom =92
        Top =0
        Name ="qryItemOnOrder"
        Name =""
    End
    Begin
        Left =424
        Top =103
        Right =587
        Bottom =187
        Top =0
        Name ="qryItemOnReserve"
        Name =""
    End
    Begin
        Left =426
        Top =201
        Right =609
        Bottom =300
        Top =0
        Name ="qryItemsSalePriceNow"
        Name =""
    End
    Begin
        Left =958
        Top =82
        Right =1054
        Bottom =196
        Top =0
        Name ="AverageDaySupplyLevels"
        Name =""
    End
End
