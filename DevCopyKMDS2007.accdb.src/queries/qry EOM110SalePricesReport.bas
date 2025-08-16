Operation =1
Option =0
Where ="(((Newprice.StartDate)>\" \") AND ((Newprice.EndDate)>\"000000\"))"
Begin InputTables
    Name ="Newprice"
    Name ="Item"
End
Begin OutputColumns
    Expression ="Newprice.Code"
    Expression ="Item.Description"
    Alias ="SalePriceStartDate"
    Expression ="CDate(Mid([startdate],3,2) & \"/\" & Right([startdate],2) & \"/\" & Left([startD"
        "ate],2))"
    Alias ="SalePriceEndDate"
    Expression ="CDate(Mid([Enddate],3,2) & \"/\" & Right([Enddate],2) & \"/\" & Left([EndDate],2"
        ") & \" 23:59:59\")"
    Alias ="Old"
    Expression ="IIf(IsNull(Item.Price),newprice.Price*0.01,Item.price)"
    Alias ="Sale"
    Expression ="[NewPrice]*0.01"
    Expression ="Item.Onhand"
    Alias ="OldValue"
    Expression ="[Onhand]*Item.Price"
    Alias ="NewValue"
    Expression ="[Onhand]*([NewPrice]*0.01)"
    Alias ="Markup"
    Expression ="[onhand]*(([NewPrice]*0.01)-Item.Price)"
End
Begin Joins
    LeftTable ="Newprice"
    RightTable ="Item"
    Expression ="Newprice.Code = Item.ID"
    Flag =1
End
Begin OrderBy
    Expression ="Newprice.Code"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="0"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="SalePriceStartDate"
        dbInteger "ColumnWidth" ="1845"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalePriceEndDate"
        dbInteger "ColumnWidth" ="2490"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Old"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sale"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OldValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Markup"
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
    Bottom =205
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =95
        Top =31
        Right =216
        Bottom =160
        Top =0
        Name ="Newprice"
        Name =""
    End
    Begin
        Left =334
        Top =27
        Right =430
        Bottom =141
        Top =0
        Name ="Item"
        Name =""
    End
End
