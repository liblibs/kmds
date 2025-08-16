Operation =1
Option =0
Where ="(((InvoLine.SetupDate) Between [Forms]![PrintInventoryReport]![txtDateRangeFrom]"
    " And [Forms]![PrintInventoryReport]![txtDateRangeTo]))"
Having ="(((SalesDepartment.ID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBe"
    "gin And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd) And ((InvoLine.Trans"
    "Type)=\"S\" Or (InvoLine.TransType)=\"C\" Or (InvoLine.TransType)=\"D\"))"
Begin InputTables
    Name ="InvoLine"
    Name ="SalesGroup"
    Name ="SalesDepartment"
    Name ="Item"
End
Begin OutputColumns
    Expression ="SalesGroup.GroupLevel"
    Alias ="Group"
    Expression ="SalesGroup.ID"
    Alias ="GroupDesc"
    Expression ="SalesGroup.Description"
    Alias ="Department"
    Expression ="SalesDepartment.ID"
    Expression ="InvoLine.TransType"
    Alias ="DepartmentDesc"
    Expression ="IIf(involine.Transtype='C',Trim(SalesDepartment.Description) & '-Club Sales',Sal"
        "esDepartment.Description)"
    Expression ="Item.Description"
    Expression ="Item.ID"
    Alias ="TotalQ"
    Expression ="Sum(InvoLine.Quantity)"
    Alias ="TotalP"
    Expression ="Sum(InvoLine.Price)"
    Alias ="TotalC"
    Expression ="Sum(InvoLine.Cost)"
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
End
Begin OrderBy
    Expression ="SalesGroup.ID"
    Flag =0
    Expression ="SalesDepartment.ID"
    Flag =0
    Expression ="IIf(involine.Transtype='C',Trim(SalesDepartment.Description) & '-Club Sales',Sal"
        "esDepartment.Description)"
    Flag =0
    Expression ="Item.Description"
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
    Expression ="InvoLine.TransType"
    GroupLevel =0
    Expression ="IIf(involine.Transtype='C',Trim(SalesDepartment.Description) & '-Club Sales',Sal"
        "esDepartment.Description)"
    GroupLevel =0
    Expression ="Item.Description"
    GroupLevel =0
    Expression ="Item.ID"
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
        dbInteger "ColumnWidth" ="5865"
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
        dbText "Name" ="TotalC"
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
    Top =16
    ColumnsShown =543
    Begin
        Left =129
        Top =29
        Right =253
        Bottom =218
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =660
        Top =37
        Right =756
        Bottom =151
        Top =0
        Name ="SalesGroup"
        Name =""
    End
    Begin
        Left =487
        Top =39
        Right =583
        Bottom =153
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =333
        Top =31
        Right =429
        Bottom =145
        Top =0
        Name ="Item"
        Name =""
    End
End
