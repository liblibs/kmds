Operation =1
Option =0
Where ="(((vwUPCItemsWithCurrentPricing.CurPrice)=0))"
Begin InputTables
    Name ="vwUPCItemsWithCurrentPricing"
End
Begin OutputColumns
    Expression ="vwUPCItemsWithCurrentPricing.UPCCode"
    Expression ="vwUPCItemsWithCurrentPricing.Description"
    Expression ="vwUPCItemsWithCurrentPricing.PkgSize"
    Expression ="vwUPCItemsWithCurrentPricing.ID"
    Expression ="vwUPCItemsWithCurrentPricing.OnShelf"
    Expression ="vwUPCItemsWithCurrentPricing.CurPrice"
    Expression ="vwUPCItemsWithCurrentPricing.Cost"
    Alias ="GP"
    Expression ="Format(([CurPrice]-[cost])/[CurPrice],\"Percent\")"
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
        dbText "Name" ="GP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwUPCItemsWithCurrentPricing.UPCCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwUPCItemsWithCurrentPricing.CurPrice"
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
    Bottom =77
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
