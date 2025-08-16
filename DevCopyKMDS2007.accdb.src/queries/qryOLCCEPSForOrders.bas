Operation =1
Option =0
Having ="(((OLCCEPS.OrderNumber)=[prmOrderNumber]))"
Begin InputTables
    Name ="OLCCEPS"
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Expression ="OLCCEPS.OLCCItem"
    Alias ="SumOfShipped"
    Expression ="Sum(OLCCEPS.Shipped)"
    Alias ="FirstOfPrice"
    Expression ="First(OLCCEPS.Price)"
    Alias ="FirstOfCurPrice"
    Expression ="First(vwItemsWithCurrentPricing.CurPrice)"
    Alias ="FirstOfUnitsPerCase"
    Expression ="First(OLCCEPS.UnitsPerCase)"
    Expression ="OLCCEPS.OrderNumber"
    Expression ="OLCCEPS.OrderDate"
End
Begin Parameters
    Name ="prmOrderNumber"
    Flag =4
    Name ="prmOrderDate"
    Flag =8
End
Begin Joins
    LeftTable ="OLCCEPS"
    RightTable ="vwItemsWithCurrentPricing"
    Expression ="OLCCEPS.OLCCItem=vwItemsWithCurrentPricing.ID"
    Flag =1
End
Begin OrderBy
    Expression ="OLCCEPS.OLCCItem"
    Flag =0
End
Begin Groups
    Expression ="OLCCEPS.OLCCItem"
    GroupLevel =0
    Expression ="OLCCEPS.OrderNumber"
    GroupLevel =0
    Expression ="OLCCEPS.OrderDate"
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
        dbText "Name" ="OLCCEPS.OLCCItem"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCEPS.OrderDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCEPS.OrderNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfShipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfCurPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfUnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =19
    Top =83
    Right =1715
    Bottom =760
    Left =-1
    Top =-1
    Right =1658
    Bottom =236
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =172
        Top =23
        Right =399
        Bottom =280
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
    Begin
        Left =506
        Top =28
        Right =650
        Bottom =283
        Top =0
        Name ="OLCCEPS"
        Name =""
    End
End
