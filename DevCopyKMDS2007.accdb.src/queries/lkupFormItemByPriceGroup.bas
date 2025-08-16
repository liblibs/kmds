dbMemo "SQL" ="SELECT vwItemsWithCurrentPricing.PriceGroupDescr, vwItemsWithCurrentPricing.Desc"
    "ription, vwItemsWithCurrentPricing.PkgSize, vwItemsWithCurrentPricing.ID, vwItem"
    "sWithCurrentPricing.OnShelf, vwItemsWithCurrentPricing.CurPrice, vwItemsWithCurr"
    "entPricing.Cost, Format(([CurPrice]-[cost])/[CurPrice],\"Percent\") AS GP\015\012"
    "FROM vwItemsWithCurrentPricing\015\012ORDER BY vwItemsWithCurrentPricing.PriceGr"
    "oupDescr, vwItemsWithCurrentPricing.Description, vwItemsWithCurrentPricing.PkgSi"
    "ze;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="GP"
        dbLong "AggregateType" ="-1"
    End
End
