dbMemo "SQL" ="SELECT SalesGroup.ID AS [Group], SalesGroup.Description AS GroupDesc, SalesDepar"
    "tment.ID AS Department, SalesDepartment.Description AS DepartmentDesc, Category."
    "ID AS CategoryID, Category.Description AS CategoryDescription, SubCategory.Descr"
    "iption AS SubCategoryDescription, Item.ID, Item.Description, Item.Onhand, Item.P"
    "rice, qryItemsSalePriceNow.SalePrice, Item.Cost, SalesDepartment.IncludeInInvent"
    "oryValueReports, SalesDepartment.TargetGP, IIf(Forms!PrintInventoryReport!chkPri"
    "ntZeroOnhands,-1,0) AS PrintZeroOnhands, Item.DateLastSold\015\012FROM SalesGrou"
    "p INNER JOIN (SalesDepartment INNER JOIN (Category RIGHT JOIN ((Item LEFT JOIN q"
    "ryItemsSalePriceNow ON Item.ID = qryItemsSalePriceNow.ItemID) LEFT JOIN SubCateg"
    "ory ON Item.SubCategoryID = SubCategory.ID) ON Category.ID = SubCategory.Categor"
    "yID) ON SalesDepartment.ID = Item.SalesDepartmentID) ON SalesGroup.ID = SalesDep"
    "artment.SalesGroupID\015\012WHERE (((SalesDepartment.ID) Between Forms!PrintInve"
    "ntoryReport!cboSalesDepartmentIDBegin And Forms!PrintInventoryReport!cboSalesDep"
    "artmentIDEnd) And ((SubCategory.Description) Between Forms!PrintInventoryReport!"
    "cboCategoryIDBegin And Forms!PrintInventoryReport!cboCategoryIDEnd) And ((Item.O"
    "nhand)<>0) And ((SalesDepartment.IncludeInInventoryValueReports)<>0) And ((Item."
    "DateLastSold)<Forms!PrintInventoryReport!txtDateRangeFrom) And ((IIf(Forms!Print"
    "InventoryReport!chkPrintZeroOnhands,-1,0))=0)) Or (((SalesDepartment.ID) Between"
    " Forms!PrintInventoryReport!cboSalesDepartmentIDBegin And Forms!PrintInventoryRe"
    "port!cboSalesDepartmentIDEnd) And ((SubCategory.Description) Between Forms!Print"
    "InventoryReport!cboCategoryIDBegin And Forms!PrintInventoryReport!cboCategoryIDE"
    "nd) And ((SalesDepartment.IncludeInInventoryValueReports)<>0) And ((Item.DateLas"
    "tSold)<Forms!PrintInventoryReport!txtDateRangeFrom) And ((IIf(Forms!PrintInvento"
    "ryReport!chkPrintZeroOnhands,-1,0))=-1))\015\012ORDER BY SalesGroup.ID, SalesDep"
    "artment.ID, Category.ID, SubCategory.Description, Item.DateLastSold;\015\012"
dbMemo "Connect" =""
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
        dbInteger "ColumnWidth" ="2280"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PrintZeroOnhands"
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
    Begin
        dbText "Name" ="Item.DateLastSold"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2145"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qryItemsSalePriceNow.SalePrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.IncludeInInventoryValueReports"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.TargetGP"
        dbLong "AggregateType" ="-1"
    End
End
