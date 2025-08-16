Operation =1
Option =0
Where ="(((qryVWItemsWithCurrentPricing.OnShelf)>-1) AND ((Settings.KeyName)=\"DrizlySto"
    "reID\") AND ((SubCategory.Description) Like \"*ONL\"))"
Begin InputTables
    Name ="SubCategory"
    Name ="qryVWItemsWithCurrentPricing"
    Name ="qryUPCItemMax"
    Name ="Item"
    Name ="Sysfile"
    Name ="Settings"
End
Begin OutputColumns
    Alias ="StoreLocationId"
    Expression ="Settings.Value"
    Alias ="ProductID"
    Expression ="qryVWItemsWithCurrentPricing.ID"
    Expression ="qryVWItemsWithCurrentPricing.Description"
    Alias ="RetailPrice"
    Expression ="qryVWItemsWithCurrentPricing.CurPrice"
    Alias ="QuantityOnHand"
    Expression ="qryVWItemsWithCurrentPricing.OnShelf"
    Alias ="UPCCode"
    Expression ="qryUPCItemMax.UPCCodeMax"
    Alias ="Size"
    Expression ="qryVWItemsWithCurrentPricing.PkgSize"
    Alias ="OtherInfo"
    Expression ="Left([subcategory].[Description],Len([subcategory].[Description])-4)"
    Expression ="Item.QtyPack"
    Expression ="Item.QtyPackPrice"
    Expression ="Item.Deposit"
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
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="ProductID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RetailPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVWItemsWithCurrentPricing.Description"
        dbInteger "ColumnWidth" ="4485"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuantityOnHand"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Size"
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
    Begin
        dbText "Name" ="OtherInfo"
        dbInteger "ColumnWidth" ="2145"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Deposit"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1291
    Bottom =806
    Left =-1
    Top =-1
    Right =1259
    Bottom =335
    Left =0
    Top =0
    ColumnsShown =539
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
        Left =543
        Top =-18
        Right =687
        Bottom =504
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
    Begin
        Left =894
        Top =240
        Right =1038
        Bottom =384
        Top =0
        Name ="Settings"
        Name =""
    End
End
