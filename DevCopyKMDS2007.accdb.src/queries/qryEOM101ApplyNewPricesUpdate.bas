Operation =4
Option =0
Where ="(((Newprice.StartDate)>\" \") AND ((Newprice.EndDate)=\"000000\"))"
Begin InputTables
    Name ="Newprice"
    Name ="Item"
End
Begin OutputColumns
    Name ="Item.Price"
    Expression ="CCur(Newprice.NewPrice*0.01)"
    Name ="Item.ChangeDate"
    Expression ="Now()"
End
Begin Joins
    LeftTable ="Newprice"
    RightTable ="Item"
    Expression ="Newprice.Code = Item.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="0"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbBoolean "FailOnError" ="0"
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
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbInteger "ColumnWidth" ="4080"
        dbBoolean "ColumnHidden" ="0"
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
    ColumnsShown =579
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
        Left =202
        Top =15
        Right =298
        Bottom =129
        Top =0
        Name ="Item"
        Name =""
    End
End
