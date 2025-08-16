Operation =1
Option =0
Where ="(((Item.SalesDepartmentID) Between Forms!MetrologicOptimiser!cboSalesDepartmentI"
    "DBegin And Forms!MetrologicOptimiser!cboSalesdepartmentIDEnd) And ((SubCategory."
    "Description) Between Forms!MetrologicOptimiser!cboNotScannedCategoryIDBegin And "
    "Forms!MetrologicOptimiser!cboNotScannedCategoryIDEnd) And ((Item.Onhand)<>0) And"
    " ((qryMetrologicScanBatchTotalByItemID.SumOfQuantity) Is Null))"
Begin InputTables
    Name ="Item"
    Name ="SubCategory"
    Name ="qryMetrologicScanBatchTotalByItemID"
    Name ="SalesDepartment"
End
Begin OutputColumns
    Expression ="Item.SalesDepartmentID"
    Expression ="SalesDepartment.Description"
    Expression ="SubCategory.ID"
    Expression ="SubCategory.Description"
    Alias ="ItemID"
    Expression ="Item.ID"
    Alias ="ItemDescription"
    Expression ="Item.Description"
    Expression ="Item.Onhand"
    Expression ="qryMetrologicScanBatchTotalByItemID.SumOfQuantity"
    Expression ="Item.Price"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="SubCategory"
    Expression ="Item.SubCategoryID = SubCategory.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="qryMetrologicScanBatchTotalByItemID"
    Expression ="Item.ID = qryMetrologicScanBatchTotalByItemID.ItemID"
    Flag =2
    LeftTable ="Item"
    RightTable ="SalesDepartment"
    Expression ="Item.SalesDepartmentID = SalesDepartment.ID"
    Flag =1
End
Begin OrderBy
    Expression ="Item.SalesDepartmentID"
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
        dbText "Name" ="SubCategory.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubCategory.Description"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Onhand"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryMetrologicScanBatchTotalByItemID.SumOfQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDescription"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-299
    Top =85
    Right =1302
    Bottom =850
    Left =-1
    Top =-1
    Right =1569
    Bottom =329
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =274
        Top =12
        Right =418
        Bottom =298
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =27
        Top =13
        Right =171
        Bottom =157
        Top =0
        Name ="SubCategory"
        Name =""
    End
    Begin
        Left =548
        Top =117
        Right =794
        Bottom =333
        Top =0
        Name ="qryMetrologicScanBatchTotalByItemID"
        Name =""
    End
    Begin
        Left =842
        Top =12
        Right =986
        Bottom =156
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
End
