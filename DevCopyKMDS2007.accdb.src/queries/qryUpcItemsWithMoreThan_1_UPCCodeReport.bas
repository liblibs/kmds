Operation =1
Option =0
Where ="(((SalesDepartment.ID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBe"
    "gin And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd) And ((UPCItem.ID)<>C"
    "Str(item.id)))"
Begin InputTables
    Name ="qryUpcItemsWithMoreThan_1_UPCCode"
    Name ="UPCItem"
    Name ="Item"
    Name ="SalesDepartment"
    Name ="SubCategory"
End
Begin OutputColumns
    Alias ="SalesDpt"
    Expression ="SalesDepartment.ID"
    Alias ="SalesDptDesc"
    Expression ="SalesDepartment.Description"
    Alias ="Category"
    Expression ="SubCategory.Description"
    Expression ="qryUpcItemsWithMoreThan_1_UPCCode.ItemID"
    Expression ="Item.Description"
    Expression ="Item.Onhand"
    Expression ="Item.Price"
    Alias ="UPCCode"
    Expression ="UPCItem.ID"
    Expression ="UPCItem.SetupDate"
End
Begin Joins
    LeftTable ="qryUpcItemsWithMoreThan_1_UPCCode"
    RightTable ="UPCItem"
    Expression ="qryUpcItemsWithMoreThan_1_UPCCode.ItemID = UPCItem.ItemID"
    Flag =1
    LeftTable ="UPCItem"
    RightTable ="Item"
    Expression ="UPCItem.ItemID = Item.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SalesDepartment"
    Expression ="Item.SalesDepartmentID = SalesDepartment.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SubCategory"
    Expression ="Item.SubCategoryID = SubCategory.ID"
    Flag =2
End
Begin OrderBy
    Expression ="SalesDepartment.ID"
    Flag =0
    Expression ="SubCategory.Description"
    Flag =0
    Expression ="Item.Description"
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
        dbText "Name" ="SalesDpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDptDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Category"
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
    Bottom =200
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =90
        Top =0
        Name ="qryUpcItemsWithMoreThan_1_UPCCode"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="UPCItem"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =459
        Top =2
        Right =555
        Bottom =116
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =675
        Top =76
        Right =771
        Bottom =190
        Top =0
        Name ="SubCategory"
        Name =""
    End
End
