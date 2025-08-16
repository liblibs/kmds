Operation =3
Name ="ItemSalePrice"
Option =0
Where ="(((Newprice.StartDate)>\" \") AND ((Newprice.EndDate)>\"000000\"))"
Begin InputTables
    Name ="Newprice"
    Name ="Item"
End
Begin OutputColumns
    Name ="ItemID"
    Expression ="Newprice.Code"
    Alias ="NewPriceStartDate"
    Name ="StartDate"
    Expression ="CDate(Mid([startdate],3,2) & \"/\" & Right([startdate],2) & \"/\" & Left([startD"
        "ate],2))"
    Alias ="NewPriceEndDate"
    Name ="EndDate"
    Expression ="CDate(Mid([Enddate],3,2) & \"/\" & Right([Enddate],2) & \"/\" & Left([EndDate],2"
        ") & \" 23:59:59\")"
    Alias ="Old"
    Name ="OldPrice"
    Expression ="IIf(IsNull(Item.Price),newprice.Price*0.01,Item.price)"
    Alias ="New"
    Name ="SalePrice"
    Expression ="[NewPrice]*0.01"
End
Begin Joins
    LeftTable ="Newprice"
    RightTable ="Item"
    Expression ="Newprice.Code = Item.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Newprice.StartDate"
        dbInteger "ColumnWidth" ="1395"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewPriceStartDate"
        dbInteger "ColumnWidth" ="2175"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewPriceEndDate"
        dbInteger "ColumnWidth" ="3075"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Old"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="New"
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
    ColumnsShown =651
    Begin
        Left =34
        Top =18
        Right =155
        Bottom =147
        Top =0
        Name ="Newprice"
        Name =""
    End
    Begin
        Left =240
        Top =11
        Right =336
        Bottom =125
        Top =0
        Name ="Item"
        Name =""
    End
End
