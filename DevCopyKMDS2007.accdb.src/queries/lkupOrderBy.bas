dbMemo "SQL" ="SELECT vwUPCItemsWithCurrentPricing.UPCCode AS CodeOrUPC, vwUPCItemsWithCurrentP"
    "ricing.Description, vwUPCItemsWithCurrentPricing.PkgSize, vwUPCItemsWithCurrentP"
    "ricing.OnShelf, vwUPCItemsWithCurrentPricing.Cost, vwUPCItemsWithCurrentPricing."
    "CurPrice, IIf([CurPrice]<>[Price],\"*\",\" \") AS Star, vwUPCItemsWithCurrentPri"
    "cing.StockNumber, vwUPCItemsWithCurrentPricing.UnitsPerCase, vwUPCItemsWithCurre"
    "ntPricing.Reserved, vwUPCItemsWithCurrentPricing.Ordered, vwUPCItemsWithCurrentP"
    "ricing.SalesDepartmentID, vwUPCItemsWithCurrentPricing.ID, vwUPCItemsWithCurrent"
    "Pricing.BuyDown\015\012FROM vwUPCItemsWithCurrentPricing\015\012WHERE (((vwUPCIt"
    "emsWithCurrentPricing.LenID)>3))\015\012ORDER BY vwUPCItemsWithCurrentPricing.Le"
    "nID, vwUPCItemsWithCurrentPricing.UPCCode;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Star"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CodeOrUPC"
        dbLong "AggregateType" ="-1"
    End
End
