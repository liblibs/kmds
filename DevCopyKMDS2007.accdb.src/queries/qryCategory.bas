Operation =1
Option =0
Begin InputTables
    Name ="SubCategory"
    Name ="Category"
    Name ="SalesDepartment"
    Name ="qryItemsByCategoryCounts"
End
Begin OutputColumns
    Expression ="Category.SalesDepartmentID"
    Expression ="SalesDepartment.Description"
    Alias ="CatDescription"
    Expression ="Category.Description"
    Alias ="SubCatDesc"
    Expression ="SubCategory.Description"
    Expression ="SubCategory.TargetGP"
    Alias ="Items"
    Expression ="IIf(IsNull([qryItemsByCategoryCounts].[count]),0,[qryItemsByCategoryCounts].[cou"
        "nt])"
End
Begin Joins
    LeftTable ="SubCategory"
    RightTable ="Category"
    Expression ="SubCategory.CategoryID = Category.ID"
    Flag =1
    LeftTable ="Category"
    RightTable ="SalesDepartment"
    Expression ="Category.SalesDepartmentID = SalesDepartment.ID"
    Flag =1
    LeftTable ="SubCategory"
    RightTable ="qryItemsByCategoryCounts"
    Expression ="SubCategory.ID = qryItemsByCategoryCounts.SubCategoryID"
    Flag =2
End
Begin OrderBy
    Expression ="Category.SalesDepartmentID"
    Flag =0
    Expression ="Category.Description"
    Flag =0
    Expression ="SubCategory.Description"
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
        dbText "Name" ="SubCategory.TargetGP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Category.SalesDepartmentID"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CatDescription"
        dbInteger "ColumnWidth" ="2370"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubcatDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Items"
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
    Bottom =429
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =524
        Top =28
        Right =668
        Bottom =172
        Top =0
        Name ="SubCategory"
        Name =""
    End
    Begin
        Left =274
        Top =55
        Right =418
        Bottom =199
        Top =0
        Name ="Category"
        Name =""
    End
    Begin
        Left =36
        Top =63
        Right =180
        Bottom =207
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =716
        Top =12
        Right =860
        Bottom =156
        Top =0
        Name ="qryItemsByCategoryCounts"
        Name =""
    End
End
