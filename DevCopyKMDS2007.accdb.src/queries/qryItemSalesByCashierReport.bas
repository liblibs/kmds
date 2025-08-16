Operation =1
Option =0
Where ="(((SalesDepartment.ID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBe"
    "gin And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd) And ((InvoLine.Setup"
    "Date) Between Forms!PrintInventoryReport!txtDateRangeFrom And Forms!PrintInvento"
    "ryReport!txtDateRangeTo))"
Begin InputTables
    Name ="SalesDepartment"
    Name ="Item"
    Name ="Size"
    Name ="SubCategory"
    Name ="InvoLine"
    Name ="Cashier"
End
Begin OutputColumns
    Expression ="InvoLine.CashierID"
    Alias ="SalesDptID"
    Expression ="SalesDepartment.ID"
    Expression ="InvoLine.SetupDate"
    Alias ="CashierName"
    Expression ="Cashier.Name"
    Alias ="SalesDptDesc"
    Expression ="SalesDepartment.Description"
    Alias ="Category"
    Expression ="IIf(IsNull(SubCategory.ID),\"n/a\",Subcategory.Description)"
    Expression ="Item.ID"
    Expression ="Item.Description"
    Alias ="SizeDesc"
    Expression ="Size.Description"
    Expression ="Item.Onhand"
    Expression ="InvoLine.Quantity"
    Expression ="InvoLine.Price"
    Alias ="Register"
    Expression ="InvoLine.CashSysID"
    Alias ="Invoice"
    Expression ="InvoLine.InvoHeadID"
End
Begin Joins
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
    RightTable ="InvoLine"
    Expression ="Item.ID = InvoLine.Code"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="Cashier"
    Expression ="InvoLine.CashierID = Cashier.ID"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLine.CashierID"
    Flag =0
    Expression ="SalesDepartment.ID"
    Flag =0
    Expression ="InvoLine.SetupDate"
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
    Begin
        dbText "Name" ="InvoLine.Quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashierName"
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
        dbText "Name" ="InvoLine.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.CashierID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.CashSysID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Register"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Invoice"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =114
    Top =178
    Right =1405
    Bottom =733
    Left =-1
    Top =-1
    Right =1259
    Bottom =238
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
        Left =443
        Top =106
        Right =539
        Bottom =220
        Top =0
        Name ="Size"
        Name =""
    End
    Begin
        Left =1019
        Top =-3
        Right =1115
        Bottom =156
        Top =0
        Name ="SubCategory"
        Name =""
    End
    Begin
        Left =715
        Top =0
        Right =856
        Bottom =180
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =1017
        Top =152
        Right =1161
        Bottom =296
        Top =0
        Name ="Cashier"
        Name =""
    End
End
