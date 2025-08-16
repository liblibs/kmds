Operation =1
Option =0
Where ="(((CStr([Item].[ID]))=[parItemID]))"
Begin InputTables
    Name ="Item"
    Name ="qryItemsSalePriceNow"
    Name ="qryBuyDownGoingOnNow"
    Name ="qryItemOnReserve"
End
Begin OutputColumns
    Alias ="Code"
    Expression ="CStr(Item.ID)"
    Alias ="Name"
    Expression ="CStr(Item.Description)"
    Alias ="Avail"
    Expression ="[Onhand]-[SoldTodayCr1]-[SoldTodayCr2]-[SoldTodayCr3]-[SoldTodayCr4]+[Transferre"
        "dToday]-[Reserved]"
    Alias ="Price"
    Expression ="CCur(Format(IIf(IsNull([SalePrice]),Item.Price,[SalePrice]),'#.00'))"
    Alias ="Buydown"
    Expression ="Format(IIf(IsNull([qryBuyDownGoingOnNow].[SpcPrice]),0,[qryBuyDownGoingOnNow].[S"
        "pcPrice]),'#.00')"
    Expression ="Item.SalesDepartmentID"
    Alias ="RegularPrice"
    Expression ="Item.Price"
    Alias ="Reserved"
    Expression ="IIf(IsNull([OnReserve]),0,[OnReserve])"
    Alias ="CurPrice"
    Expression ="[Price]-[Buydown]"
End
Begin Parameters
    Name ="parItemID"
    Flag =4
End
Begin Joins
    LeftTable ="Item"
    RightTable ="qryItemsSalePriceNow"
    Expression ="Item.ID = qryItemsSalePriceNow.ItemID"
    Flag =2
    LeftTable ="Item"
    RightTable ="qryBuyDownGoingOnNow"
    Expression ="Item.ID = qryBuyDownGoingOnNow.ItemsCode"
    Flag =2
    LeftTable ="Item"
    RightTable ="qryItemOnReserve"
    Expression ="Item.ID = qryItemOnReserve.ItemID"
    Flag =2
End
Begin OrderBy
    Expression ="Item.ID"
    Flag =0
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
        dbText "Name" ="RegularPrice"
        dbInteger "ColumnWidth" ="1275"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Avail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buydown"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Reserved"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurPrice"
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
    Bottom =274
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =44
        Top =49
        Right =140
        Bottom =163
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =205
        Top =-10
        Right =301
        Bottom =104
        Top =0
        Name ="qryItemsSalePriceNow"
        Name =""
    End
    Begin
        Left =200
        Top =129
        Right =296
        Bottom =243
        Top =0
        Name ="qryBuyDownGoingOnNow"
        Name =""
    End
    Begin
        Left =474
        Top =14
        Right =781
        Bottom =98
        Top =0
        Name ="qryItemOnReserve"
        Name =""
    End
End
