Operation =1
Option =0
Where ="(((SalesDepartment.ID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBe"
    "gin And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd) And ((Vendor.ID) Bet"
    "ween Forms!PrintInventoryReport!cboVendorIDBegin And Forms!PrintInventoryReport!"
    "cboVendorIDEnd))"
Begin InputTables
    Name ="SalesDepartment"
    Name ="Item"
    Name ="qryItemsByDaysSupplySoldTotals"
    Name ="Size"
    Name ="SubCategory"
    Name ="Vendor"
End
Begin OutputColumns
    Alias ="SalesDptID"
    Expression ="SalesDepartment.ID"
    Alias ="SalesDptDesc"
    Expression ="SalesDepartment.Description"
    Alias ="VendorID"
    Expression ="Vendor.ID"
    Alias ="VendorName"
    Expression ="Vendor.Name"
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
    Expression ="qryItemsByDaysSupplySoldTotals.SumOfCost"
    Alias ="Profit"
    Expression ="IIf([SumOfPrice]=0,0,([SumOfPrice]-[SumOfCost])/[SumOfPrice])"
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
    LeftTable ="Item"
    RightTable ="Vendor"
    Expression ="Item.VendorID = Vendor.ID"
    Flag =1
End
Begin OrderBy
    Expression ="SalesDepartment.ID"
    Flag =0
    Expression ="Vendor.ID"
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
        dbText "Name" ="Profit"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
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
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Onhand"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemsByDaysSupplySoldTotals.SumOfQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemsByDaysSupplySoldTotals.SumOfPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.DateLastSold"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemsByDaysSupplySoldTotals.SumOfCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorName"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1539
    Bottom =677
    Left =-1
    Top =-1
    Right =1507
    Bottom =219
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
        Bottom =144
        Top =0
        Name ="qryItemsByDaysSupplySoldTotals"
        Name =""
    End
    Begin
        Left =649
        Top =153
        Right =745
        Bottom =267
        Top =0
        Name ="Size"
        Name =""
    End
    Begin
        Left =522
        Top =179
        Right =618
        Bottom =338
        Top =0
        Name ="SubCategory"
        Name =""
    End
    Begin
        Left =793
        Top =12
        Right =937
        Bottom =156
        Top =0
        Name ="Vendor"
        Name =""
    End
End
