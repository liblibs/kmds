Operation =1
Option =0
Begin InputTables
    Name ="vwUPCItemsWithCurrentPricing"
End
Begin OutputColumns
    Expression ="vwUPCItemsWithCurrentPricing.LenID"
    Expression ="vwUPCItemsWithCurrentPricing.UPCCode"
    Expression ="vwUPCItemsWithCurrentPricing.ID"
    Expression ="vwUPCItemsWithCurrentPricing.Description"
    Expression ="vwUPCItemsWithCurrentPricing.Onhand"
    Expression ="vwUPCItemsWithCurrentPricing.Cost"
    Expression ="vwUPCItemsWithCurrentPricing.Price"
    Expression ="vwUPCItemsWithCurrentPricing.CurPrice"
    Expression ="vwUPCItemsWithCurrentPricing.BuyDown"
    Expression ="vwUPCItemsWithCurrentPricing.ItemVendorID"
    Expression ="vwUPCItemsWithCurrentPricing.StockNumber"
    Expression ="vwUPCItemsWithCurrentPricing.UnitsPerCase"
    Expression ="vwUPCItemsWithCurrentPricing.UnitsPerOrder"
    Expression ="vwUPCItemsWithCurrentPricing.PkgSize"
    Expression ="vwUPCItemsWithCurrentPricing.SalesDepartmentID"
    Expression ="vwUPCItemsWithCurrentPricing.SoldTodayCr1"
    Expression ="vwUPCItemsWithCurrentPricing.SoldTodayCr2"
    Expression ="vwUPCItemsWithCurrentPricing.SoldTodayCr3"
    Expression ="vwUPCItemsWithCurrentPricing.SoldTodayCr4"
    Expression ="vwUPCItemsWithCurrentPricing.TransferredToday"
    Expression ="vwUPCItemsWithCurrentPricing.OnShelf"
    Expression ="vwUPCItemsWithCurrentPricing.Reserved"
    Expression ="vwUPCItemsWithCurrentPricing.Ordered"
    Expression ="vwUPCItemsWithCurrentPricing.Transfer"
    Expression ="vwUPCItemsWithCurrentPricing.TPR"
    Expression ="vwUPCItemsWithCurrentPricing.PriceGroupID"
    Expression ="vwUPCItemsWithCurrentPricing.SubCategoryID"
    Expression ="vwUPCItemsWithCurrentPricing.Location"
    Expression ="vwUPCItemsWithCurrentPricing.PLUList"
    Expression ="vwUPCItemsWithCurrentPricing.SalesDptDescr"
    Expression ="vwUPCItemsWithCurrentPricing.PriceGroupDescr"
    Expression ="vwUPCItemsWithCurrentPricing.CategoryDescr"
    Expression ="vwUPCItemsWithCurrentPricing.VendorName"
    Expression ="vwUPCItemsWithCurrentPricing.DaysSupply"
    Expression ="vwUPCItemsWithCurrentPricing.SizeID"
    Expression ="vwUPCItemsWithCurrentPricing.AverageDailySales"
    Expression ="vwUPCItemsWithCurrentPricing.DaysSupplyGroupID"
    Expression ="vwUPCItemsWithCurrentPricing.Sprice"
End
Begin OrderBy
    Expression ="vwUPCItemsWithCurrentPricing.LenID"
    Flag =0
    Expression ="vwUPCItemsWithCurrentPricing.UPCCode"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="vwUPCItemsWithCurrentPricing.LenID"
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
    Bottom =347
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="vwUPCItemsWithCurrentPricing"
        Name =""
    End
End
