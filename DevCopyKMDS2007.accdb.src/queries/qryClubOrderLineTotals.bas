Operation =1
Option =0
Begin InputTables
    Name ="Sysfile"
    Name ="Club"
    Name ="OrdHead"
    Name ="OrderLine"
    Name ="Item"
    Name ="SalesDepartment"
End
Begin OutputColumns
    Expression ="OrderLine.OrdheadID"
    Alias ="LiquorQuantity"
    Expression ="Sum(IIf(SalesDepartment.ID=1,Orderline.QuantityShipped,0))"
    Alias ="LiquorExtPrice"
    Expression ="Sum(IIf(SalesDepartment.ID=1,Orderline.Price*[QuantityShipped],0))"
    Alias ="LiquorDiscountRate"
    Expression ="IIf(Club.ApplyDiscount,sysfile.ClubDiscRate,0)"
    Alias ="LiquorDiscount"
    Expression ="((([LiquorExtPrice]*[LiquorDiscountRate])\\1)/100)"
    Alias ="LiquorExtCost"
    Expression ="Sum(IIf(SalesDepartment.ID=1,Orderline.Cost*[QuantityShipped],0))"
    Expression ="Club.NonLiquor"
    Alias ="NonLiquorQuantity"
    Expression ="Sum(IIf(SalesDepartment.ID<>1,Orderline.QuantityShipped,0))"
    Alias ="NonLiquorExtPrice"
    Expression ="Sum(IIf(SalesDepartment.ID<>1,(Orderline.Price-OrderLine.BuyDown)*[QuantityShipp"
        "ed],0))"
    Alias ="NonLiquorDiscountRate"
    Expression ="IIf(Club.ApplyDiscount,sysfile.NonClubDiscRate,0)"
    Alias ="NonLiquorExtPriceDiscountable"
    Expression ="Sum(IIf(SalesDepartment.ID<>1,IIf(SalesDepartment.AllowClubDiscounts,(Orderline."
        "Price-OrderLine.BuyDown)*Orderline.QuantityShipped,0),0))"
    Alias ="NonLiquorDiscount"
    Expression ="((([NonLiquorExtPriceDiscountable]*[NonLiquorDiscountRate])\\1)/100)"
    Alias ="NonLiquorExtCost"
    Expression ="Sum(IIf(SalesDepartment.ID<>1,Orderline.Cost*[QuantityShipped],0))"
    Alias ="Quantity"
    Expression ="Sum(OrderLine.QuantityShipped)"
    Alias ="ExtPrice"
    Expression ="[LiquorExtPrice]+[NonLiquorExtPrice]"
    Alias ="ExtCost"
    Expression ="[LiquorExtCost]+[NonLiquorExtCost]"
    Alias ="ExtDiscount"
    Expression ="[LiquorDiscount]+[NonLiquorDiscount]"
    Alias ="ExtNetPrice"
    Expression ="[LiquorExtPrice]+[NonLiquorExtPrice]-[LiquorDiscount]-[NonLiquorDiscount]"
End
Begin Joins
    LeftTable ="Club"
    RightTable ="OrdHead"
    Expression ="Club.ID = OrdHead.ClubID"
    Flag =1
    LeftTable ="OrdHead"
    RightTable ="OrderLine"
    Expression ="OrdHead.ID = OrderLine.OrdheadID"
    Flag =1
    LeftTable ="OrderLine"
    RightTable ="Item"
    Expression ="OrderLine.ItemID = Item.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SalesDepartment"
    Expression ="Item.SalesDepartmentID = SalesDepartment.ID"
    Flag =1
End
Begin Groups
    Expression ="OrderLine.OrdheadID"
    GroupLevel =0
    Expression ="IIf(Club.ApplyDiscount,sysfile.ClubDiscRate,0)"
    GroupLevel =0
    Expression ="Club.NonLiquor"
    GroupLevel =0
    Expression ="IIf(Club.ApplyDiscount,sysfile.NonClubDiscRate,0)"
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
dbMemo "Filter" ="([qryClubOrderLineTotals].[OrdheadID]=8856)"
Begin
    Begin
        dbText "Name" ="LiquorExtCost"
        dbInteger "ColumnWidth" ="1695"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LiquorDiscountRate"
        dbInteger "ColumnWidth" ="1890"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LiquorDiscount"
        dbInteger "ColumnWidth" ="2490"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NonLiquorExtPrice"
        dbInteger "ColumnWidth" ="2745"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NonLiquorExtCost"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NonLiquorDiscountRate"
        dbInteger "ColumnWidth" ="2220"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NonLiquorExtPriceDiscountable"
        dbInteger "ColumnWidth" ="2895"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LiquorExtPrice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ExtPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NonLiquorDiscount"
        dbInteger "ColumnWidth" ="2190"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LiquorQuantity"
        dbInteger "ColumnWidth" ="1485"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.NonLiquor"
        dbInteger "ColumnWidth" ="1065"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NonLiquorQuantity"
        dbInteger "ColumnWidth" ="2550"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtDiscount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtNetPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderLine.OrdheadID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =9
    Top =214
    Right =1868
    Bottom =487
    Left =-1
    Top =-1
    Right =1827
    Bottom =173
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =970
        Top =10
        Right =1155
        Bottom =169
        Top =0
        Name ="Sysfile"
        Name =""
    End
    Begin
        Left =11
        Top =12
        Right =125
        Bottom =156
        Top =0
        Name ="Club"
        Name =""
    End
    Begin
        Left =187
        Top =30
        Right =283
        Bottom =144
        Top =0
        Name ="OrdHead"
        Name =""
    End
    Begin
        Left =330
        Top =15
        Right =489
        Bottom =174
        Top =0
        Name ="OrderLine"
        Name =""
    End
    Begin
        Left =516
        Top =18
        Right =653
        Bottom =162
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =721
        Top =6
        Right =934
        Bottom =180
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
End
