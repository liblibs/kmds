Operation =2
Name ="MetrologicLookup"
Option =0
Where ="(((Format(UPCItem.ID,\"!@@@@@@@@@@@@@@\"))>\"0001\") And ((vwItemsWithCurrentPri"
    "cing.SalesDepartmentID) Between [prmSalesdepartmentIdBegin] And [prmSalesdepartm"
    "entIdEnd]) And ((vwItemsWithCurrentPricing.CategoryDescr) Between [prmCategoryDe"
    "scBegin] And [prmCategoryDescEnd]) And ((Len(UPCItem.id))<15))"
Begin InputTables
    Name ="UPCItem"
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Alias ="UPC"
    Expression ="Format([UPCItem].[ID],\"!@@@@@@@@@@@@@@\")"
    Alias ="Descr"
    Expression ="Format(Left([Description],9),\"!@@@@@@@@@\")"
    Alias ="Item"
    Expression ="Format([ItemiD],\"!@@@@@@@@@@@\")"
    Alias ="PRI"
    Expression ="Format(Format([Price],\"Fixed\"),\"@@@@@@\")"
    Expression ="vwItemsWithCurrentPricing.curPrice"
    Expression ="vwItemsWithCurrentPricing.SalesDepartmentID"
    Alias ="CatDesc"
    Expression ="vwItemsWithCurrentPricing.CategoryDescr"
    Alias ="Shelf"
    Expression ="vwItemsWithCurrentPricing.onshelf"
    Expression ="vwItemsWithCurrentPricing.Description"
    Expression ="vwItemsWithCurrentPricing.UnitsPerCase"
    Alias ="UCase"
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
    Expression ="Format([UPCItem].[ID],\"!@@@@@@@@@@@@@@\")"
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
End
Begin
    State =0
    Left =48
    Top =46
    Right =1472
    Bottom =811
    Left =-1
    Top =-1
    Right =1392
    Bottom =184
    Left =0
    Top =0
    ColumnsShown =539
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
