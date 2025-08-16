Operation =1
Option =0
Where ="(((vwUPCItemsWithCurrentPricing.LenID)>2))"
Begin InputTables
    Name ="vwUPCItemsWithCurrentPricing"
End
Begin OutputColumns
    Alias ="CodeOrUPC"
    Expression ="vwUPCItemsWithCurrentPricing.UPCCode"
    Expression ="vwUPCItemsWithCurrentPricing.Description"
    Expression ="vwUPCItemsWithCurrentPricing.PkgSize"
    Expression ="vwUPCItemsWithCurrentPricing.OnShelf"
    Expression ="vwUPCItemsWithCurrentPricing.Cost"
    Expression ="vwUPCItemsWithCurrentPricing.CurPrice"
    Alias ="Star"
    Expression ="IIf([SPrice]<>0,\"*\",\" \")"
    Expression ="vwUPCItemsWithCurrentPricing.StockNumber"
    Expression ="vwUPCItemsWithCurrentPricing.UnitsPerCase"
    Expression ="vwUPCItemsWithCurrentPricing.Reserved"
    Expression ="vwUPCItemsWithCurrentPricing.Ordered"
    Expression ="vwUPCItemsWithCurrentPricing.SalesDepartmentID"
    Expression ="vwUPCItemsWithCurrentPricing.ID"
    Expression ="vwUPCItemsWithCurrentPricing.BuyDown"
    Expression ="vwUPCItemsWithCurrentPricing.ItemVendorID"
    Expression ="vwUPCItemsWithCurrentPricing.DaysSupplyGroupID"
    Expression ="vwUPCItemsWithCurrentPricing.AverageDailySales"
    Expression ="vwUPCItemsWithCurrentPricing.DaysSupply"
    Alias ="CurDaysSupply"
    Expression ="IIf([AverageDailySales]=0,0,[OnShelf]/[AverageDailySales])"
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
        dbText "Name" ="Star"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CodeOrUPC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwUPCItemsWithCurrentPricing.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwUPCItemsWithCurrentPricing.PkgSize"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwUPCItemsWithCurrentPricing.OnShelf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwUPCItemsWithCurrentPricing.Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwUPCItemsWithCurrentPricing.CurPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwUPCItemsWithCurrentPricing.StockNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwUPCItemsWithCurrentPricing.UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwUPCItemsWithCurrentPricing.Reserved"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwUPCItemsWithCurrentPricing.Ordered"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwUPCItemsWithCurrentPricing.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwUPCItemsWithCurrentPricing.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwUPCItemsWithCurrentPricing.BuyDown"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwUPCItemsWithCurrentPricing.ItemVendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwUPCItemsWithCurrentPricing.DaysSupplyGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwUPCItemsWithCurrentPricing.AverageDailySales"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwUPCItemsWithCurrentPricing.DaysSupply"
        dbInteger "ColumnWidth" ="1485"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurDaysSupply"
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
    Bottom =295
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =54
        Top =-9
        Right =300
        Bottom =470
        Top =0
        Name ="vwUPCItemsWithCurrentPricing"
        Name =""
    End
End
