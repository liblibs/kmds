Operation =1
Option =0
Where ="(((Item.RestockAmount)<>0) AND ((Item.ID) Is Not Null))"
Begin InputTables
    Name ="Item"
    Name ="Size"
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Expression ="vwItemsWithCurrentPricing.SalesDptDescr"
    Alias ="Category"
    Expression ="vwItemsWithCurrentPricing.CategoryDescr"
    Expression ="Item.ID"
    Alias ="Description"
    Expression ="vwItemsWithCurrentPricing.Description"
    Alias ="Size"
    Expression ="Size.Description"
    Expression ="Item.RestockAmount"
    Alias ="Dollars"
    Expression ="[RestockAmount]*[vwItemsWithCurrentPricing].[Price]"
    Expression ="vwItemsWithCurrentPricing.OnShelf"
    Expression ="Item.Location"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="Size"
    Expression ="Item.SizeID = Size.ID"
    Flag =1
    LeftTable ="vwItemsWithCurrentPricing"
    RightTable ="Item"
    Expression ="vwItemsWithCurrentPricing.ID = Item.ID"
    Flag =1
End
Begin OrderBy
    Expression ="vwItemsWithCurrentPricing.CategoryDescr"
    Flag =0
    Expression ="Size.Description"
    Flag =0
    Expression ="Item.ID"
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
        dbText "Name" ="Item.RestockAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dollars"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Category"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Size"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.OnShelf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.SalesDptDescr"
        dbInteger "ColumnWidth" ="2400"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Location"
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
    Bottom =240
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =72
        Top =30
        Right =227
        Bottom =258
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =330
        Top =133
        Right =474
        Bottom =277
        Top =0
        Name ="Size"
        Name =""
    End
    Begin
        Left =510
        Top =14
        Right =654
        Bottom =158
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
End
