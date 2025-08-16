Operation =1
Option =0
Where ="(((InvoLine.TransType)=\"S\" Or (InvoLine.TransType)=\"C\" Or (InvoLine.TransTyp"
    "e)=\"D\") And ((InvoLine.SetupDate) Between Forms!PrintInventoryReport!txtDateRa"
    "ngeFrom And Forms!PrintInventoryReport!txtDateRangeTo))"
Begin InputTables
    Name ="InvoLine"
    Name ="SalesGroup"
    Name ="SalesDepartment"
    Name ="Item"
    Name ="SubCategory"
End
Begin OutputColumns
    Expression ="SalesGroup.GroupLevel"
    Alias ="Group"
    Expression ="SalesGroup.ID"
    Alias ="GroupDesc"
    Expression ="SalesGroup.Description"
    Alias ="Department"
    Expression ="SalesDepartment.ID"
    Alias ="DepartmentDesc"
    Expression ="SalesDepartment.Description"
    Alias ="SubCategoryDescr"
    Expression ="IIf(IsNull(SubCategory.Description),\"n/a\",SubCategory.Description)"
    Alias ="TotalQ"
    Expression ="Sum(InvoLine.Quantity)"
    Alias ="TotalP"
    Expression ="Sum(InvoLine.Price)"
    Alias ="Q1"
    Expression ="Sum(IIf([CashsysID]=1,[Quantity],0))"
    Alias ="D1"
    Expression ="Sum(IIf([CashsysID]=1,involine.Price,0))"
    Alias ="Q2"
    Expression ="Sum(IIf([CashsysID]=2,[Quantity],0))"
    Alias ="D2"
    Expression ="Sum(IIf([CashsysID]=2,involine.Price,0))"
    Alias ="Q3"
    Expression ="Sum(IIf([CashsysID]=3,[Quantity],0))"
    Alias ="D3"
    Expression ="Sum(IIf([CashsysID]=3,involine.Price,0))"
    Alias ="Q4"
    Expression ="Sum(IIf([CashsysID]=4,[Quantity],0))"
    Alias ="D4"
    Expression ="Sum(IIf([CashsysID]=4,involine.Price,0))"
End
Begin Joins
    LeftTable ="SalesDepartment"
    RightTable ="Item"
    Expression ="SalesDepartment.ID = Item.SalesDepartmentID"
    Flag =1
    LeftTable ="SalesGroup"
    RightTable ="SalesDepartment"
    Expression ="SalesGroup.ID = SalesDepartment.SalesGroupID"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SubCategory"
    Expression ="Item.SubCategoryID = SubCategory.ID"
    Flag =2
End
Begin OrderBy
    Expression ="SalesGroup.ID"
    Flag =0
    Expression ="SalesDepartment.ID"
    Flag =0
    Expression ="IIf(IsNull(SubCategory.Description),\"n/a\",SubCategory.Description)"
    Flag =0
End
Begin Groups
    Expression ="SalesGroup.GroupLevel"
    GroupLevel =0
    Expression ="SalesGroup.ID"
    GroupLevel =0
    Expression ="SalesGroup.Description"
    GroupLevel =0
    Expression ="SalesDepartment.ID"
    GroupLevel =0
    Expression ="SalesDepartment.Description"
    GroupLevel =0
    Expression ="IIf(IsNull(SubCategory.Description),\"n/a\",SubCategory.Description)"
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
        dbText "Name" ="DepartmentDesc"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubCategoryDescr"
        dbInteger "ColumnWidth" ="3165"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Department"
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
    Bottom =245
    Left =0
    Top =53
    ColumnsShown =543
    Begin
        Left =71
        Top =20
        Right =195
        Bottom =209
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =660
        Top =0
        Right =756
        Bottom =114
        Top =0
        Name ="SalesGroup"
        Name =""
    End
    Begin
        Left =487
        Top =2
        Right =583
        Bottom =116
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =333
        Top =-6
        Right =429
        Bottom =108
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =549
        Top =137
        Right =645
        Bottom =251
        Top =0
        Name ="SubCategory"
        Name =""
    End
End
