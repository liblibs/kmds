Operation =1
Option =0
Begin InputTables
    Name ="vw_WebLogErrorCode-1"
    Name ="OrdHead"
    Name ="OrderLine"
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Expression ="[vw_WebLogErrorCode-1].WebOrderID"
    Expression ="[vw_WebLogErrorCode-1].LicenseeName"
    Expression ="OrdHead.InvoiceNo"
    Expression ="OrderLine.ItemID"
    Expression ="vwItemsWithCurrentPricing.Description"
    Expression ="OrderLine.QuantityOrdered"
    Alias ="ExtPrice"
    Expression ="[Orderline].[Price]*[QuantityOrdered]"
    Expression ="OrderLine.QuantityShipped"
End
Begin Joins
    LeftTable ="vw_WebLogErrorCode-1"
    RightTable ="OrdHead"
    Expression ="[vw_WebLogErrorCode-1].WebOrderID = OrdHead.InvoiceNo"
    Flag =1
    LeftTable ="OrdHead"
    RightTable ="OrderLine"
    Expression ="OrdHead.ID = OrderLine.OrdheadID"
    Flag =1
    LeftTable ="OrderLine"
    RightTable ="vwItemsWithCurrentPricing"
    Expression ="OrderLine.ItemID = vwItemsWithCurrentPricing.ID"
    Flag =1
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
        dbText "Name" ="OrderLine.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderLine.QuantityOrdered"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderLine.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[vw_WebLogErrorCode-1].WebOrderID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[vw_WebLogErrorCode-1].LicenseeName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdHead.InvoiceNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderLine.QuantityShipped"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-111
    Top =67
    Right =1420
    Bottom =900
    Left =-1
    Top =-1
    Right =1508
    Bottom =515
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =241
        Top =0
        Name ="vw_WebLogErrorCode-1"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="OrdHead"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =329
        Top =0
        Name ="OrderLine"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =872
        Bottom =311
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
End
