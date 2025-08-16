Operation =3
Name ="MetrologicLookup"
Option =0
Where ="(((Len(upcitem.ID))=12) And ((Format(Mid(UPCItem.ID,2,10),\"!@@@@@@@@@@@@@@\")) "
    "Between \"01\" And \"98\") And ((vwItemsWithCurrentPricing.SalesDepartmentID) Be"
    "tween [prmSalesdepartmentIdBegin] And [prmSalesdepartmentIdEnd]) And ((vwItemsWi"
    "thCurrentPricing.CategoryDescr) Between [prmCategoryDescBegin] And [prmCategoryD"
    "escEnd]) And ((Len(UPCItem.id))<15))"
Begin InputTables
    Name ="UPCItem"
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Alias ="UPC"
    Name ="UPC"
    Expression ="Format(Mid([UPCItem].[ID],2,10),\"!@@@@@@@@@@@@@@\")"
    Alias ="Descr"
    Name ="Descr"
    Expression ="Format(Left([Description],9),\"!@@@@@@@@@\")"
    Alias ="Item"
    Name ="Item"
    Expression ="Format([ItemiD],\"!@@@@@@@@@@@\")"
    Alias ="PRI"
    Name ="PRI"
    Expression ="Format(Format([Price],\"Fixed\"),\"@@@@@@\")"
    Name ="curPrice"
    Expression ="vwItemsWithCurrentPricing.curPrice"
    Name ="SalesDepartmentID"
    Expression ="vwItemsWithCurrentPricing.SalesDepartmentID"
    Alias ="CatDesc"
    Name ="CatDesc"
    Expression ="vwItemsWithCurrentPricing.CategoryDescr"
    Alias ="Shelf"
    Name ="Shelf"
    Expression ="vwItemsWithCurrentPricing.onshelf"
    Name ="Description"
    Expression ="vwItemsWithCurrentPricing.Description"
    Name ="UnitsPerCase"
    Expression ="vwItemsWithCurrentPricing.UnitsPerCase"
    Alias ="UCase"
    Name ="UCase"
    Expression ="Format([UnitsPerCase],\"@@@\")"
End
Begin Parameters
    Name ="prmSalesdepartmentIdBegin"
    Flag =4
    Name ="prmSalesdepartmentIdEnd"
    Flag =4
    Name ="prmCategoryDescBegin"
    Flag =10
    Name ="prmCategoryDescEnd"
    Flag =10
End
Begin Joins
    LeftTable ="UPCItem"
    RightTable ="vwItemsWithCurrentPricing"
    Expression ="UPCItem.ItemID = vwItemsWithCurrentPricing.ID"
    Flag =1
End
Begin OrderBy
    Expression ="UPCItem.ID"
    Flag =0
    Expression ="Format(Mid([UPCItem].[ID],2,10),\"!@@@@@@@@@@@@@@\")"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="UPC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Descr"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PRI"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CatDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.curPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Shelf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCItem.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-141
    Top =97
    Right =1283
    Bottom =862
    Left =-1
    Top =-1
    Right =1392
    Bottom =184
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="UPCItem"
        Name =""
    End
    Begin
        Left =402
        Top =9
        Right =611
        Bottom =268
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
End
