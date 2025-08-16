Operation =1
Option =0
Where ="(((CCur(Format(Newprice.NewPrice*0.01,\"Fixed\")))<>item.Price) And ((Newprice.S"
    "tartDate)>\" \") And ((Newprice.EndDate)=\"000000\"))"
Begin InputTables
    Name ="Newprice"
    Name ="Item"
End
Begin OutputColumns
    Expression ="Newprice.Code"
    Expression ="Item.Description"
    Alias ="Old"
    Expression ="Item.Price"
    Alias ="New"
    Expression ="CCur(Format(Newprice.NewPrice*0.01,\"Fixed\"))"
    Expression ="Item.Onhand"
    Alias ="OldValue"
    Expression ="[onhand]*Item.Price"
    Alias ="NewValue"
    Expression ="CCur(Format([onhand]*Newprice.NewPrice*0.01,\"Fixed\"))"
    Alias ="NewPriceStartDate"
    Expression ="CDate(Mid([startdate],3,2) & \"/\" & Right([startdate],2) & \"/\" & Left([startD"
        "ate],2))"
    Alias ="Markup"
    Expression ="IIf(((Newprice.NewPrice*0.01)-Item.Price)>0,((Newprice.NewPrice*0.01)-Item.Price"
        ")*[onhand],0)"
    Alias ="MarkDown"
    Expression ="IIf(((Newprice.NewPrice*0.01)-Item.Price)<=0,((Newprice.NewPrice*0.01)-Item.Pric"
        "e)*[onhand],0)"
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
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="NewPriceStartDate"
        dbInteger "ColumnWidth" ="2175"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Newprice.Code"
        dbInteger "ColumnWidth" ="645"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="New"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
        dbByte "DecimalPlaces" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbInteger "ColumnWidth" ="4080"
        dbBoolean "ColumnHidden" ="0"
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
    Begin
        dbText "Name" ="MarkDown"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Old"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Onhand"
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
    Bottom =188
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =33
        Top =25
        Right =129
        Bottom =154
        Top =0
        Name ="Newprice"
        Name =""
    End
    Begin
        Left =257
        Top =16
        Right =353
        Bottom =130
        Top =0
        Name ="Item"
        Name =""
    End
End
