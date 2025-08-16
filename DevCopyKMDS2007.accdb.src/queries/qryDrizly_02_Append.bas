Operation =3
Name ="Drizly"
Option =0
Where ="(((SubCategory.Description) Like \"*ONL\"))"
Begin InputTables
    Name ="SubCategory"
    Name ="qryVWItemsWithCurrentPricing"
    Name ="qryUPCItemMax"
    Name ="Item"
    Name ="Sysfile"
End
Begin OutputColumns
    Alias ="StoreLocationId"
    Name ="StoreLocationId"
    Expression ="\"OLCC \" & [sysfile].[ID]"
    Alias ="ProductID"
    Name ="ProductID"
    Expression ="qryVWItemsWithCurrentPricing.ID"
    Name ="Description"
    Expression ="qryVWItemsWithCurrentPricing.Description"
    Alias ="RetailPrice"
    Name ="RetailPrice"
    Expression ="qryVWItemsWithCurrentPricing.CurPrice"
    Alias ="QuantityOnHand"
    Name ="QuantityOnHand"
    Expression ="qryVWItemsWithCurrentPricing.OnShelf"
    Alias ="UPCCode"
    Name ="UPCCode"
    Expression ="qryUPCItemMax.UPCCodeMax"
    Alias ="Size"
    Name ="Size"
    Expression ="qryVWItemsWithCurrentPricing.PkgSize"
    Name ="OtherInfo"
    Expression ="SubCategory.Description"
    Name ="PackQty"
    Expression ="Item.QtyPack"
    Name ="PackRetailPrice"
    Expression ="Item.QtyPackPrice"
End
Begin Joins
    LeftTable ="qryVWItemsWithCurrentPricing"
    RightTable ="qryUPCItemMax"
    Expression ="qryVWItemsWithCurrentPricing.ID = qryUPCItemMax.ItemID"
    Flag =1
    LeftTable ="qryVWItemsWithCurrentPricing"
    RightTable ="Item"
    Expression ="qryVWItemsWithCurrentPricing.ID = Item.ID"
    Flag =1
    LeftTable ="SubCategory"
    RightTable ="Item"
    Expression ="SubCategory.ID = Item.SubCategoryID"
    Flag =1
End
Begin OrderBy
    Expression ="qryVWItemsWithCurrentPricing.Description"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="SubCategory.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Drizly.ProductID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ProductID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Drizly.RetailPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RetailPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVWItemsWithCurrentPricing.OnShelf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVWItemsWithCurrentPricing.Description"
        dbInteger "ColumnWidth" ="4485"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Drizly.QuantityOnHand"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuantityOnHand"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Drizly.UPCCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVWItemsWithCurrentPricing.PkgSize"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Size"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StoreLocation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StoreLocationId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.QtyPack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.QtyPackPrice"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1291
    Bottom =846
    Left =-1
    Top =-1
    Right =1259
    Bottom =437
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =230
        Bottom =256
        Top =0
        Name ="SubCategory"
        Name =""
    End
    Begin
        Left =550
        Top =51
        Right =694
        Bottom =573
        Top =0
        Name ="qryVWItemsWithCurrentPricing"
        Name =""
    End
    Begin
        Left =774
        Top =29
        Right =918
        Bottom =173
        Top =0
        Name ="qryUPCItemMax"
        Name =""
    End
    Begin
        Left =311
        Top =55
        Right =462
        Bottom =430
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =966
        Top =12
        Right =1110
        Bottom =156
        Top =0
        Name ="Sysfile"
        Name =""
    End
End
