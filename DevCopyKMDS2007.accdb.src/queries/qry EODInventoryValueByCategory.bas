Operation =1
Option =0
Where ="(((SalesDepartment.IncludeInInventoryValueReports)<>0))"
Begin InputTables
    Name ="SalesGroup"
    Name ="SalesDepartment"
    Name ="Category"
    Name ="Item"
    Name ="qryItemsSalePriceNow"
    Name ="SubCategory"
End
Begin OutputColumns
    Alias ="Group"
    Expression ="SalesGroup.ID"
    Alias ="GroupDesc"
    Expression ="SalesGroup.Description"
    Alias ="Department"
    Expression ="SalesDepartment.ID"
    Alias ="DepartmentDesc"
    Expression ="SalesDepartment.Description"
    Alias ="CategoryID"
    Expression ="Category.ID"
    Alias ="CategoryDescription"
    Expression ="Category.Description"
    Alias ="SubCategoryDescription"
    Expression ="SubCategory.Description"
    Expression ="Item.ID"
    Expression ="Item.Description"
    Expression ="Item.Onhand"
    Expression ="Item.Price"
    Expression ="qryItemsSalePriceNow.SalePrice"
    Expression ="Item.Cost"
    Expression ="SalesDepartment.IncludeInInventoryValueReports"
    Expression ="SalesDepartment.TargetGP"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="qryItemsSalePriceNow"
    Expression ="Item.ID = qryItemsSalePriceNow.ItemID"
    Flag =2
    LeftTable ="Item"
    RightTable ="SubCategory"
    Expression ="Item.SubCategoryID = SubCategory.ID"
    Flag =2
    LeftTable ="Category"
    RightTable ="SubCategory"
    Expression ="Category.ID = SubCategory.CategoryID"
    Flag =3
    LeftTable ="SalesDepartment"
    RightTable ="Item"
    Expression ="SalesDepartment.ID = Item.SalesDepartmentID"
    Flag =1
    LeftTable ="SalesGroup"
    RightTable ="SalesDepartment"
    Expression ="SalesGroup.ID = SalesDepartment.SalesGroupID"
    Flag =1
End
Begin OrderBy
    Expression ="SalesGroup.ID"
    Flag =0
    Expression ="SalesDepartment.ID"
    Flag =0
    Expression ="Category.ID"
    Flag =0
    Expression ="SubCategory.Description"
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
        dbText "Name" ="Item.Onhand"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Price"
        dbInteger "ColumnWidth" ="840"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbInteger "ColumnWidth" ="840"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbInteger "ColumnWidth" ="3120"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DepartmentDesc"
        dbInteger "ColumnWidth" ="1710"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupDesc"
        dbInteger "ColumnWidth" ="1185"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Department"
        dbInteger "ColumnWidth" ="1260"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CategoryDescription"
        dbInteger "ColumnWidth" ="4170"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubCategoryDescription"
        dbInteger "ColumnWidth" ="4170"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CategoryID"
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
    Bottom =183
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =37
        Top =12
        Right =133
        Bottom =111
        Top =0
        Name ="SalesGroup"
        Name =""
    End
    Begin
        Left =171
        Top =20
        Right =267
        Bottom =119
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =907
        Top =70
        Right =1003
        Bottom =184
        Top =0
        Name ="Category"
        Name =""
    End
    Begin
        Left =316
        Top =23
        Right =481
        Bottom =167
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =550
        Top =23
        Right =646
        Bottom =122
        Top =0
        Name ="qryItemsSalePriceNow"
        Name =""
    End
    Begin
        Left =572
        Top =111
        Right =668
        Bottom =225
        Top =0
        Name ="SubCategory"
        Name =""
    End
End
