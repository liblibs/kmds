Operation =1
Option =0
Where ="(((qryPurchaseOrderItemDaysSupplyADS.Sold)<>0))"
Begin InputTables
    Name ="Vendor"
    Name ="qryPurchaseOrderItemDaysSupplyADS"
    Name ="qryVWItemsWithCurrentPricing"
End
Begin OutputColumns
    Expression ="qryPurchaseOrderItemDaysSupplyADS.VendorID"
    Alias ="VendorName"
    Expression ="Vendor.Name"
    Expression ="qryPurchaseOrderItemDaysSupplyADS.ItemId"
    Expression ="qryVWItemsWithCurrentPricing.Description"
    Alias ="SizeDesc"
    Expression ="qryVWItemsWithCurrentPricing.PkgSize"
    Expression ="qryPurchaseOrderItemDaysSupplyADS.Sold"
    Expression ="qryPurchaseOrderItemDaysSupplyADS.SoldDays"
    Expression ="qryVWItemsWithCurrentPricing.Ordered"
    Expression ="qryVWItemsWithCurrentPricing.Reserved"
    Alias ="Avail"
    Expression ="[OnShelf]+[Ordered]"
    Alias ="SOQRaw"
    Expression ="qryPurchaseOrderItemDaysSupplyADS.Sold"
    Alias ="OrderMultiple"
    Expression ="qryVWItemsWithCurrentPricing.UnitsPerOrder"
    Alias ="SOQ"
    Expression ="CInt((([Sold])/[UnitsPerOrder]))*[UnitsPerOrder]"
    Expression ="qryPurchaseOrderItemDaysSupplyADS.DateStart"
    Expression ="qryPurchaseOrderItemDaysSupplyADS.DateEnd"
    Expression ="qryVWItemsWithCurrentPricing.UnitsPerCase"
    Alias ="Repack"
    Expression ="IIf([soq] Mod [UnitsPerCase]<>0,True,False)"
    Alias ="SPrice"
    Expression ="qryVWItemsWithCurrentPricing.CurPrice"
    Alias ="SPriceStar"
    Expression ="IIf([Price]=[CurPrice],\" \",\"*\")"
    Alias ="ExtPrice"
    Expression ="[SOQ]*[SPrice]"
    Expression ="qryVWItemsWithCurrentPricing.Cost"
    Alias ="ExtCost"
    Expression ="[SOQ]*[Cost]"
End
Begin Parameters
    Name ="parBeginDate"
    Flag =8
    Name ="parEndDate"
    Flag =8
    Name ="parVendorID"
    Flag =4
    Name ="parBeginSalesDepartmentID"
    Flag =4
    Name ="parEndSalesDepartmentID"
    Flag =4
End
Begin Joins
    LeftTable ="Vendor"
    RightTable ="qryPurchaseOrderItemDaysSupplyADS"
    Expression ="Vendor.ID = qryPurchaseOrderItemDaysSupplyADS.VendorID"
    Flag =1
    LeftTable ="qryPurchaseOrderItemDaysSupplyADS"
    RightTable ="qryVWItemsWithCurrentPricing"
    Expression ="qryPurchaseOrderItemDaysSupplyADS.ItemId = qryVWItemsWithCurrentPricing.ID"
    Flag =1
End
Begin OrderBy
    Expression ="qryVWItemsWithCurrentPricing.SalesDepartmentID"
    Flag =0
    Expression ="qryVWItemsWithCurrentPricing.Description"
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
        dbText "Name" ="qryVWItemsWithCurrentPricing.Description"
        dbInteger "ColumnWidth" ="2700"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="OrderMultiple"
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
    Bottom =253
    Left =0
    Top =51
    ColumnsShown =539
    Begin
        Left =46
        Top =14
        Right =177
        Bottom =323
        Top =0
        Name ="Vendor"
        Name =""
    End
    Begin
        Left =278
        Top =110
        Right =374
        Bottom =254
        Top =0
        Name ="qryPurchaseOrderItemDaysSupplyADS"
        Name =""
    End
    Begin
        Left =412
        Top =6
        Right =508
        Bottom =120
        Top =0
        Name ="qryVWItemsWithCurrentPricing"
        Name =""
    End
End
