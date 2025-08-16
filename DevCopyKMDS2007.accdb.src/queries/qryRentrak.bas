Operation =1
Option =0
Where ="(((Len([UPCItem].[ID]))>7))"
Begin InputTables
    Name ="UPCItem"
    Name ="Item"
    Name ="SubCategory"
    Name ="Size"
End
Begin OutputColumns
    Expression ="SubCategory.Description"
    Expression ="Item.ID"
    Expression ="Item.Description"
    Alias ="SizeDesc"
    Expression ="Size.Description"
    Expression ="Size.CaseQty"
    Expression ="Item.Price"
    Alias ="UPCCode"
    Expression ="Max(UPCItem.ID)"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="SubCategory"
    Expression ="Item.SubCategoryID = SubCategory.ID"
    Flag =1
    LeftTable ="UPCItem"
    RightTable ="Item"
    Expression ="UPCItem.ItemID = Item.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="Size"
    Expression ="Item.SizeID = Size.ID"
    Flag =1
End
Begin OrderBy
    Expression ="SubCategory.Description"
    Flag =0
    Expression ="Item.ID"
    Flag =0
End
Begin Groups
    Expression ="SubCategory.Description"
    GroupLevel =0
    Expression ="Item.ID"
    GroupLevel =0
    Expression ="Item.Description"
    GroupLevel =0
    Expression ="Size.Description"
    GroupLevel =0
    Expression ="Size.CaseQty"
    GroupLevel =0
    Expression ="Item.Price"
    GroupLevel =0
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
        dbText "Name" ="SubCategory.Description"
        dbInteger "ColumnWidth" ="2340"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Price"
        dbInteger "ColumnWidth" ="885"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCCode"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="UPCItem"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =284
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =392
        Top =78
        Right =488
        Bottom =192
        Top =0
        Name ="SubCategory"
        Name =""
    End
    Begin
        Left =526
        Top =6
        Right =622
        Bottom =120
        Top =0
        Name ="Size"
        Name =""
    End
End
