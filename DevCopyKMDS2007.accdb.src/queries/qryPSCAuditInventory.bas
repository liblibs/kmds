Operation =1
Option =0
Begin InputTables
    Name ="PSCInput"
    Name ="vwUPCItemsWithCurrentPricing"
End
Begin OutputColumns
    Expression ="PSCInput.SalesDepartmentID"
    Expression ="vwUPCItemsWithCurrentPricing.ID"
    Alias ="QuantityShipped"
    Expression ="Sum(PSCInput.Quantity)"
    Expression ="vwUPCItemsWithCurrentPricing.CurPrice"
    Expression ="vwUPCItemsWithCurrentPricing.Cost"
    Expression ="vwUPCItemsWithCurrentPricing.BuyDown"
    Expression ="vwUPCItemsWithCurrentPricing.OnShelf"
End
Begin Joins
    LeftTable ="PSCInput"
    RightTable ="vwUPCItemsWithCurrentPricing"
    Expression ="PSCInput.Code = vwUPCItemsWithCurrentPricing.UPCCode"
    Flag =1
End
Begin OrderBy
    Expression ="PSCInput.SalesDepartmentID"
    Flag =0
    Expression ="vwUPCItemsWithCurrentPricing.ID"
    Flag =0
End
Begin Groups
    Expression ="PSCInput.SalesDepartmentID"
    GroupLevel =0
    Expression ="vwUPCItemsWithCurrentPricing.ID"
    GroupLevel =0
    Expression ="vwUPCItemsWithCurrentPricing.CurPrice"
    GroupLevel =0
    Expression ="vwUPCItemsWithCurrentPricing.Cost"
    GroupLevel =0
    Expression ="vwUPCItemsWithCurrentPricing.BuyDown"
    GroupLevel =0
    Expression ="vwUPCItemsWithCurrentPricing.OnShelf"
    GroupLevel =0
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
        dbText "Name" ="QuantityShipped"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="PSCInput"
        Name =""
    End
    Begin
        Left =202
        Top =-1
        Right =450
        Bottom =113
        Top =0
        Name ="vwUPCItemsWithCurrentPricing"
        Name =""
    End
End
