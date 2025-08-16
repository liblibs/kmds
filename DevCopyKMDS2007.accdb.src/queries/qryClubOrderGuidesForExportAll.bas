Operation =1
Option =0
Where ="(((ClubOrderGuides.ClubID)=[parClubID]))"
Begin InputTables
    Name ="ClubOrderGuides"
    Name ="qryPurchaseOrderItemsWithTodaysPrice"
    Name ="Club"
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Expression ="ClubOrderGuides.ClubID"
    Alias ="ClubName"
    Expression ="Club.Name"
    Expression ="ClubOrderGuides.LineID"
    Expression ="ClubOrderGuides.ItemID"
    Expression ="qryPurchaseOrderItemsWithTodaysPrice.Description"
    Expression ="ClubOrderGuides.Par"
    Alias ="Price"
    Expression ="qryPurchaseOrderItemsWithTodaysPrice.Sprice"
    Alias ="OnSale"
    Expression ="qryPurchaseOrderItemsWithTodaysPrice.Star"
    Alias ="AsOf"
    Expression ="Format(Now(),\"mm/dd/yyyy\")"
    Alias ="Category"
    Expression ="vwItemsWithCurrentPricing.CategoryDescr"
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
        dbText "Name" ="ClubOrderGuides.LineID"
        dbInteger "ColumnWidth" ="735"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubOrderGuides.ItemID"
        dbInteger "ColumnWidth" ="765"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubOrderGuides.Par"
        dbInteger "ColumnWidth" ="465"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemsWithTodaysPrice.Description"
        dbInteger "ColumnWidth" ="2910"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AsOf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Category"
        dbInteger "ColumnWidth" ="4575"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OnSale"
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
    Bottom =230
    Left =0
    Top =0
    ColumnsShown =539
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
        Left =307
        Top =111
        Right =701
        Bottom =225
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
