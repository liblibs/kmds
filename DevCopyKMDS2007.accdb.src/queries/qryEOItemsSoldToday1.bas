Operation =1
Option =0
Having ="(((InvoLine5.TransType)=\"S\" Or (InvoLine5.TransType)=\"D\"))"
Begin InputTables
    Name ="InvoLine5"
    Name ="SalesGroup"
    Name ="SalesDepartment"
    Name ="Item"
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
    Alias ="SoldToday"
    Expression ="Sum(InvoLine5.Quantity)"
    Alias ="DollarsSoldToday"
    Expression ="Sum(InvoLine5.Price)"
    Expression ="InvoLine5.TransType"
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
    LeftTable ="InvoLine5"
    RightTable ="Item"
    Expression ="InvoLine5.Code = Item.ID"
    Flag =1
End
Begin OrderBy
    Expression ="SalesGroup.ID"
    Flag =0
    Expression ="SalesDepartment.ID"
    Flag =0
    Expression ="Item.ID"
    Flag =0
End
Begin Groups
    Expression ="SalesGroup.ID"
    GroupLevel =0
    Expression ="SalesGroup.Description"
    GroupLevel =0
    Expression ="SalesDepartment.ID"
    GroupLevel =0
    Expression ="SalesDepartment.Description"
    GroupLevel =0
    Expression ="Item.ID"
    GroupLevel =0
    Expression ="Item.Description"
    GroupLevel =0
    Expression ="InvoLine5.TransType"
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
        dbText "Name" ="SoldToday"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DollarsSoldToday"
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
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.TransType"
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
    Bottom =179
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =41
        Top =37
        Right =176
        Bottom =196
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =665
        Top =18
        Right =761
        Bottom =132
        Top =0
        Name ="SalesGroup"
        Name =""
    End
    Begin
        Left =509
        Top =24
        Right =605
        Bottom =138
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =222
        Top =68
        Right =318
        Bottom =182
        Top =0
        Name ="Item"
        Name =""
    End
End
