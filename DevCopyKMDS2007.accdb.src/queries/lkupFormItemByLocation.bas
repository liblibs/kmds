dbMemo "SQL" ="SELECT vwItemsWithCurrentPricing.Location, vwItemsWithCurrentPricing.Description"
    ", vwItemsWithCurrentPricing.PkgSize, vwItemsWithCurrentPricing.ID, vwItemsWithCu"
    "rrentPricing.OnShelf, vwItemsWithCurrentPricing.CurPrice, vwItemsWithCurrentPric"
    "ing.Cost, Format(([CurPrice]-[cost])/[CurPrice],\"Percent\") AS GP\015\012FROM v"
    "wItemsWithCurrentPricing\015\012ORDER BY vwItemsWithCurrentPricing.Location, vwI"
    "temsWithCurrentPricing.Description;\015\012"
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
