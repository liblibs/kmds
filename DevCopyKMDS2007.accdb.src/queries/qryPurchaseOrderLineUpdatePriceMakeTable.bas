Operation =2
Name ="PurchaseOrderLineUpdatePrice"
Option =0
Begin InputTables
    Name ="Item"
    Name ="Size"
    Name ="qryItemsSalePriceNow"
End
Begin OutputColumns
    Expression ="Item.ID"
    Expression ="Item.Description"
    Expression ="Item.Onhand"
    Expression ="Item.Cost"
    Alias ="Sprice"
    Expression ="IIf(IsNull([SalePrice]),[Price],[SalePrice])"
    Alias ="Star"
    Expression ="IIf(IsNull([SalePrice]),\" \",\"*\")"
    Alias ="ItemVendorID"
    Expression ="Item.VendorID"
    Expression ="Item.UnitsPerCase"
    Alias ="PkgSize"
    Expression ="Size.Description"
    Expression ="Item.SalesDepartmentID"
    Expression ="Size.Oz"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="Size"
    Expression ="Item.SizeID = Size.ID"
    Flag =2
    LeftTable ="Item"
    RightTable ="qryItemsSalePriceNow"
    Expression ="Item.ID = qryItemsSalePriceNow.ItemID"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sprice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Star"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemVendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PkgSize"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =38
    Top =63
    Right =858
    Bottom =611
    Left =-1
    Top =-1
    Right =788
    Bottom =203
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Size"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="qryItemsSalePriceNow"
        Name =""
    End
End
