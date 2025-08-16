Operation =1
Option =0
Where ="(((Item.SetupDate)>Format(Now()-1,\"mm/dd/yyyy\"\" 23\"\":\"\"59\"\":\"\"59\"\"\""
    ")))"
Begin InputTables
    Name ="SalesGroup"
    Name ="SalesDepartment"
    Name ="Item"
    Name ="UPCItem"
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
    Expression ="Item.ID"
    Expression ="Item.Description"
    Alias ="UPCCode"
    Expression ="UPCItem.ID"
    Expression ="Item.Price"
    Expression ="Item.SetupDate"
    Alias ="SetupSince"
    Expression ="Format(Now()-1,\"mm/dd/yyyy\")"
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
    LeftTable ="Item"
    RightTable ="UPCItem"
    Expression ="Item.ID = UPCItem.ItemID"
    Flag =2
End
Begin OrderBy
    Expression ="SalesGroup.ID"
    Flag =0
    Expression ="SalesDepartment.ID"
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
        dbText "Name" ="GroupDesc"
        dbInteger "ColumnWidth" ="1440"
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
        dbText "Name" ="SetupSince"
        dbInteger "ColumnWidth" ="2175"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DepartmentDesc"
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
    Bottom =195
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =679
        Top =10
        Right =775
        Bottom =124
        Top =0
        Name ="SalesGroup"
        Name =""
    End
    Begin
        Left =542
        Top =11
        Right =638
        Bottom =125
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =377
        Top =14
        Right =473
        Bottom =128
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =193
        Top =5
        Right =289
        Bottom =119
        Top =0
        Name ="UPCItem"
        Name =""
    End
End
