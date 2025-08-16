Operation =1
Option =0
Where ="(((Item.VendorID)=[parVendorID]) AND ((SubCategory.Description)<>\"OBSOLETE\"))"
Begin InputTables
    Name ="qryItemsSalePriceNow"
    Name ="Item"
    Name ="SubCategory"
End
Begin OutputColumns
    Expression ="Item.VendorID"
    Expression ="qryItemsSalePriceNow.ItemID"
    Expression ="Item.Description"
    Expression ="Item.Price"
    Expression ="qryItemsSalePriceNow.SalePrice"
    Expression ="Item.Cost"
    Expression ="SubCategory.Description"
End
Begin Parameters
    Name ="parVendorID"
    Flag =4
End
Begin Joins
    LeftTable ="qryItemsSalePriceNow"
    RightTable ="Item"
    Expression ="qryItemsSalePriceNow.ItemID = Item.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SubCategory"
    Expression ="Item.SubCategoryID = SubCategory.ID"
    Flag =1
End
Begin OrderBy
    Expression ="qryItemsSalePriceNow.ItemID"
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
        dbText "Name" ="Item.Description"
        dbInteger "ColumnWidth" ="2805"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.VendorID"
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
    Right =1193
    Bottom =296
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =351
        Top =61
        Right =507
        Bottom =175
        Top =0
        Name ="qryItemsSalePriceNow"
        Name =""
    End
    Begin
        Left =156
        Top =144
        Right =252
        Bottom =333
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =555
        Top =12
        Right =699
        Bottom =156
        Top =0
        Name ="SubCategory"
        Name =""
    End
End
