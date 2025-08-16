Operation =1
Option =0
Begin InputTables
    Name ="vw_WebLogErrorCode-1"
    Name ="OrdHead"
    Name ="OrderLine"
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Expression ="[vw_WebLogErrorCode-1].LicenseeName"
    Expression ="[vw_WebLogErrorCode-1].WebOrderNotes"
    Expression ="OrdHead.InvoiceNo"
    Expression ="OrdHead.InvDate"
    Expression ="OrderLine.ItemID"
    Expression ="vwItemsWithCurrentPricing.Description"
    Expression ="OrderLine.QuantityOrdered"
    Expression ="OrderLine.Price"
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
        dbText "Name" ="[vw_WebLogErrorCode-1].LicenseeName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[vw_WebLogErrorCode-1].WebOrderNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdHead.InvoiceNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdHead.OrdDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdHead.InvDate"
        dbLong "AggregateType" ="-1"
    End
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
End
Begin
    State =0
    Left =82
    Top =141
    Right =1239
    Bottom =934
    Left =-1
    Top =-1
    Right =1134
    Bottom =523
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =394
        Top =0
        Name ="vw_WebLogErrorCode-1"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =455
        Top =0
        Name ="OrdHead"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =427
        Top =0
        Name ="OrderLine"
        Name =""
    End
    Begin
        Left =696
        Top =12
        Right =840
        Bottom =156
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
End
