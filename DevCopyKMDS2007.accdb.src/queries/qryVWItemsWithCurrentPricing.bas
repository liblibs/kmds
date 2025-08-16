Operation =1
Option =0
Begin InputTables
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Expression ="vwItemsWithCurrentPricing.ID"
    Expression ="vwItemsWithCurrentPricing.Description"
    Expression ="vwItemsWithCurrentPricing.Onhand"
    Expression ="vwItemsWithCurrentPricing.Cost"
    Expression ="vwItemsWithCurrentPricing.Price"
    Expression ="vwItemsWithCurrentPricing.CurPrice"
    Expression ="vwItemsWithCurrentPricing.BuyDown"
    Expression ="vwItemsWithCurrentPricing.ItemVendorID"
    Expression ="vwItemsWithCurrentPricing.StockNumber"
    Expression ="vwItemsWithCurrentPricing.UnitsPerCase"
    Expression ="vwItemsWithCurrentPricing.UnitsPerOrder"
    Expression ="vwItemsWithCurrentPricing.PkgSize"
    Expression ="vwItemsWithCurrentPricing.SalesDepartmentID"
    Expression ="vwItemsWithCurrentPricing.SoldTodayCr1"
    Expression ="vwItemsWithCurrentPricing.SoldTodayCr2"
    Expression ="vwItemsWithCurrentPricing.SoldTodayCr3"
    Expression ="vwItemsWithCurrentPricing.SoldTodayCr4"
    Expression ="vwItemsWithCurrentPricing.TransferredToday"
    Expression ="vwItemsWithCurrentPricing.OnShelf"
    Expression ="vwItemsWithCurrentPricing.Reserved"
    Expression ="vwItemsWithCurrentPricing.Ordered"
    Expression ="vwItemsWithCurrentPricing.Transfer"
    Expression ="vwItemsWithCurrentPricing.TPR"
    Expression ="vwItemsWithCurrentPricing.PriceGroupID"
    Expression ="vwItemsWithCurrentPricing.SubCategoryID"
    Expression ="vwItemsWithCurrentPricing.Location"
    Expression ="vwItemsWithCurrentPricing.PLUList"
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
        dbText "Name" ="vwItemsWithCurrentPricing.ID"
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
    Bottom =167
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
End
