Operation =2
Name ="PurchaseOrderItemWADS"
Option =0
Where ="(((vwItemsWithCurrentPricing.ItemVendorID)=[parVendorID]) AND ((vwItemsWithCurre"
    "ntPricing.CategoryDescr)<>\"OBSOLETE\"))"
Begin InputTables
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Alias ="VendorID"
    Expression ="vwItemsWithCurrentPricing.ItemVendorID"
    Alias ="Name"
    Expression ="vwItemsWithCurrentPricing.VendorName"
    Alias ="ItemID"
    Expression ="vwItemsWithCurrentPricing.ID"
    Expression ="vwItemsWithCurrentPricing.Description"
    Expression ="vwItemsWithCurrentPricing.SizeID"
    Expression ="vwItemsWithCurrentPricing.Onhand"
    Expression ="vwItemsWithCurrentPricing.Ordered"
    Expression ="vwItemsWithCurrentPricing.Reserved"
    Alias ="Avail"
    Expression ="[OnShelf]+[Ordered]"
    Expression ="vwItemsWithCurrentPricing.AverageDailySales"
    Expression ="vwItemsWithCurrentPricing.DaysSupply"
    Alias ="OrderMax"
    Expression ="CLng(IIf(IsNull([AverageDailySales]),0,[AverageDailySales])*[DaysSupply])"
    Alias ="SOQRaw"
    Expression ="([OrderMax]-[Avail]+0.01)"
    Alias ="OrderMultiple"
    Expression ="IIf([UnitsPerOrder]<1,1,[UnitsPerOrder])"
    Alias ="SOQ"
    Expression ="(CInt([SOQRaw]/[OrderMultiple])*[OrderMultiple])"
    Expression ="vwItemsWithCurrentPricing.UnitsPerCase"
    Expression ="vwItemsWithCurrentPricing.Price"
    Expression ="vwItemsWithCurrentPricing.Sprice"
    Alias ="SPriceStar"
    Expression ="IIf(([SPrice]=0),\" \",\"*\")"
    Alias ="ExtPrice"
    Expression ="[SOQ]*[SPrice]"
    Expression ="vwItemsWithCurrentPricing.Cost"
    Alias ="ExtCost"
    Expression ="[SOQ]*[Cost]"
    Expression ="vwItemsWithCurrentPricing.SalesDepartmentID"
    Alias ="Repack"
    Expression ="IIf([soq] Mod [UnitsPerCase]<>0,True,False)"
    Expression ="vwItemsWithCurrentPricing.CategoryDescr"
End
Begin OrderBy
    Expression ="vwItemsWithCurrentPricing.ID"
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
        dbText "Name" ="vwItemsWithCurrentPricing.ItemVendorID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Name"
        dbInteger "ColumnWidth" ="765"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.VendorName"
        dbInteger "ColumnWidth" ="765"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Name"
        dbInteger "ColumnWidth" ="765"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Description"
        dbInteger "ColumnWidth" ="1020"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.SizeID"
        dbInteger "ColumnWidth" ="750"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Onhand"
        dbInteger "ColumnWidth" ="0"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="-1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Ordered"
        dbInteger "ColumnWidth" ="990"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Reserved"
        dbInteger "ColumnWidth" ="990"
        dbInteger "ColumnOrder" ="16"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Cost"
        dbInteger "ColumnOrder" ="21"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.SalesDepartmentID"
        dbInteger "ColumnOrder" ="23"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Sprice"
        dbInteger "ColumnWidth" ="885"
        dbInteger "ColumnOrder" ="18"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Repack"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =261
    Top =66
    Right =784
    Bottom =743
    Left =-1
    Top =-1
    Right =491
    Bottom =408
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =13
        Top =12
        Right =109
        Bottom =186
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
End
