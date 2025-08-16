dbMemo "SQL" ="SELECT vwItemsWithCurrentPricing.SalesDptDescr, vwItemsWithCurrentPricing.Descri"
    "ption, vwItemsWithCurrentPricing.PkgSize, vwItemsWithCurrentPricing.ID, vwItemsW"
    "ithCurrentPricing.OnShelf, vwItemsWithCurrentPricing.CurPrice, vwItemsWithCurren"
    "tPricing.Cost, Format(([CurPrice]-[cost])/[CurPrice],\"Percent\") AS GP\015\012F"
    "ROM vwItemsWithCurrentPricing\015\012ORDER BY vwItemsWithCurrentPricing.SalesDpt"
    "Descr, vwItemsWithCurrentPricing.Description, vwItemsWithCurrentPricing.PkgSize;"
    "\015\012"
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
