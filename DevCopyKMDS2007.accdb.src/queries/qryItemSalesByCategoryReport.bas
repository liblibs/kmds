Operation =1
Option =0
Where ="(((SalesDepartment.ID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBe"
    "gin And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd) And ((IIf(IsNull(Sub"
    "Category.ID),\"n/a\",Subcategory.Description)) Between Forms!PrintInventoryRepor"
    "t!cboCategoryIDBegin And Forms!PrintInventoryReport!cboCategoryIDEnd))"
Begin InputTables
    Name ="SalesDepartment"
    Name ="Item"
    Name ="qryItemsByDaysSupplySoldTotals"
    Name ="Size"
    Name ="SubCategory"
End
Begin OutputColumns
    Alias ="SalesDptID"
    Expression ="SalesDepartment.ID"
    Alias ="SalesDptDesc"
    Expression ="SalesDepartment.Description"
    Alias ="Category"
    Expression ="IIf(IsNull(SubCategory.ID),\"n/a\",Subcategory.Description)"
    Expression ="Item.ID"
    Expression ="Item.Description"
    Alias ="SizeDesc"
    Expression ="Size.Description"
    Expression ="Item.Onhand"
    Expression ="Item.Price"
    Expression ="qryItemsByDaysSupplySoldTotals.SumOfQuantity"
    Expression ="qryItemsByDaysSupplySoldTotals.SumOfPrice"
    Expression ="Item.SetupDate"
    Expression ="Item.DateLastSold"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="qryItemsByDaysSupplySoldTotals"
    Expression ="Item.ID = qryItemsByDaysSupplySoldTotals.Code"
    Flag =1
    LeftTable ="SalesDepartment"
    RightTable ="Item"
    Expression ="SalesDepartment.ID = Item.SalesDepartmentID"
    Flag =1
    LeftTable ="Item"
    RightTable ="Size"
    Expression ="Item.SizeID = Size.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SubCategory"
    Expression ="Item.SubCategoryID = SubCategory.ID"
    Flag =2
End
Begin OrderBy
    Expression ="SalesDepartment.ID"
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
        dbText "Name" ="Category"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDptID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDptDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeDesc"
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
    Bottom =217
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =208
        Bottom =195
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =242
        Top =3
        Right =388
        Bottom =192
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =446
        Top =-1
        Right =675
        Bottom =98
        Top =0
        Name ="qryItemsByDaysSupplySoldTotals"
        Name =""
    End
    Begin
        Left =443
        Top =106
        Right =539
        Bottom =220
        Top =0
        Name ="Size"
        Name =""
    End
    Begin
        Left =657
        Top =78
        Right =753
        Bottom =237
        Top =0
        Name ="SubCategory"
        Name =""
    End
End
