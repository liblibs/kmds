Operation =1
Option =0
Where ="(((OLCCEPS.OrderNumber)=Forms!OLCCPackingList!OrderNumber) And ((OLCCEPS.Price)<"
    ">[CurPrice]))"
Begin InputTables
    Name ="OLCCEPS"
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Expression ="OLCCEPS.OrderNumber"
    Expression ="vwItemsWithCurrentPricing.ID"
    Expression ="vwItemsWithCurrentPricing.Description"
    Expression ="OLCCEPS.Shipped"
    Alias ="OLAS"
    Expression ="OLCCEPS.Price"
    Alias ="KMDS"
    Expression ="vwItemsWithCurrentPricing.CurPrice"
    Alias ="Diff"
    Expression ="([shipped]*([CURprice]-[olcceps].[Price]))"
End
Begin Joins
    LeftTable ="OLCCEPS"
    RightTable ="vwItemsWithCurrentPricing"
    Expression ="OLCCEPS.OLCCItem=vwItemsWithCurrentPricing.ID"
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
        dbText "Name" ="OLCCEPS.OrderNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCEPS.Shipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Diff"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLAS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KMDS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Description"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =36
    Top =129
    Right =2217
    Bottom =1369
    Left =-1
    Top =-1
    Right =2143
    Bottom =886
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =60
        Top =15
        Right =240
        Bottom =415
        Top =0
        Name ="OLCCEPS"
        Name =""
    End
    Begin
        Left =587
        Top =187
        Right =767
        Bottom =637
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
End
