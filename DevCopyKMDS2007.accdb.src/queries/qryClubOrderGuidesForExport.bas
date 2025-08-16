Operation =1
Option =0
Where ="(((ClubOrderGuides.ClubID)=[Forms]![Clubs]![ID]))"
Begin InputTables
    Name ="Sysfile"
    Name ="ClubOrderGuides"
    Name ="qryPurchaseOrderItemsWithTodaysPrice"
    Name ="Club"
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Alias ="Line"
    Expression ="ClubOrderGuides.LineID"
    Alias ="Qty"
    Expression ="0"
    Expression ="ClubOrderGuides.ItemID"
    Expression ="qryPurchaseOrderItemsWithTodaysPrice.Description"
    Expression ="ClubOrderGuides.Par"
    Alias ="Price"
    Expression ="qryPurchaseOrderItemsWithTodaysPrice.Sprice"
    Alias ="Liquor"
    Expression ="[Qty]*[Price]"
    Alias ="NonLiquor"
    Expression ="0"
    Alias ="AsOf"
    Expression ="\"'\" & CStr(Format(Now(),\"mmm d\"\", \"\"yyyy\"))"
    Alias ="Category"
    Expression ="vwItemsWithCurrentPricing.CategoryDescr"
    Expression ="qryPurchaseOrderItemsWithTodaysPrice.PkgSize"
    Alias ="PricePer125"
    Expression ="([Sprice]/[Oz])*1.25"
    Expression ="ClubOrderGuides.ClubID"
    Alias ="ClubName"
    Expression ="Club.Name"
    Alias ="StoreName"
    Expression ="Sysfile.Name"
End
Begin Joins
    LeftTable ="ClubOrderGuides"
    RightTable ="qryPurchaseOrderItemsWithTodaysPrice"
    Expression ="ClubOrderGuides.ItemID = qryPurchaseOrderItemsWithTodaysPrice.ID"
    Flag =1
    LeftTable ="ClubOrderGuides"
    RightTable ="Club"
    Expression ="ClubOrderGuides.ClubID = Club.ID"
    Flag =1
    LeftTable ="ClubOrderGuides"
    RightTable ="vwItemsWithCurrentPricing"
    Expression ="ClubOrderGuides.ItemID = vwItemsWithCurrentPricing.ID"
    Flag =1
End
Begin OrderBy
    Expression ="ClubOrderGuides.LineID"
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
        dbText "Name" ="ClubOrderGuides.ItemID"
        dbInteger "ColumnWidth" ="765"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="ClubOrderGuides.Par"
        dbInteger "ColumnWidth" ="720"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemsWithTodaysPrice.Description"
        dbInteger "ColumnWidth" ="2910"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="AsOf"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="Category"
        dbInteger "ColumnWidth" ="4575"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="PricePer125"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1920"
        dbInteger "ColumnOrder" ="12"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ClubName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Price"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="ClubOrderGuides.ClubID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemsWithTodaysPrice.PkgSize"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="StoreName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Qty"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Line"
        dbInteger "ColumnWidth" ="735"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Liquor"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NonLiquor"
        dbInteger "ColumnOrder" ="8"
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
    Bottom =21
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =21
        Top =191
        Right =165
        Bottom =335
        Top =0
        Name ="Sysfile"
        Name =""
    End
    Begin
        Left =134
        Top =44
        Right =266
        Bottom =158
        Top =0
        Name ="ClubOrderGuides"
        Name =""
    End
    Begin
        Left =375
        Top =149
        Right =769
        Bottom =263
        Top =0
        Name ="qryPurchaseOrderItemsWithTodaysPrice"
        Name =""
    End
    Begin
        Left =3
        Top =6
        Right =99
        Bottom =120
        Top =0
        Name ="Club"
        Name =""
    End
    Begin
        Left =397
        Top =-8
        Right =493
        Bottom =106
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
End
