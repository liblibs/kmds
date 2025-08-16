dbMemo "SQL" ="SELECT vwItemsWithCurrentPricing.VendorName, vwItemsWithCurrentPricing.StockNumb"
    "er, vwItemsWithCurrentPricing.Description, vwItemsWithCurrentPricing.PkgSize, vw"
    "ItemsWithCurrentPricing.ID, vwItemsWithCurrentPricing.OnShelf, vwItemsWithCurren"
    "tPricing.CurPrice, vwItemsWithCurrentPricing.Cost, Format(([CurPrice]-[cost])/[C"
    "urPrice],\"Percent\") AS GP\015\012FROM vwItemsWithCurrentPricing\015\012ORDER B"
    "Y vwItemsWithCurrentPricing.VendorName, vwItemsWithCurrentPricing.StockNumber, v"
    "wItemsWithCurrentPricing.Description, vwItemsWithCurrentPricing.PkgSize;\015\012"
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
