Operation =1
Option =0
Where ="(((Item.Onhand)<0) AND ((SalesDepartment.NegativeOnhand)<>0))"
Begin InputTables
    Name ="SalesGroup"
    Name ="SalesDepartment"
    Name ="Item"
    Name ="qryUPCItemMax"
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
    Expression ="qryUPCItemMax.UPCCodeMax"
    Expression ="Item.Onhand"
    Expression ="Item.Price"
    Expression ="SalesDepartment.NegativeOnhand"
    Alias ="Ext"
    Expression ="[Onhand]*[Price]"
End
Begin Joins
    LeftTable ="SalesDepartment"
    RightTable ="Item"
    Expression ="SalesDepartment.ID=Item.SalesDepartmentID"
    Flag =1
    LeftTable ="SalesGroup"
    RightTable ="SalesDepartment"
    Expression ="SalesGroup.ID=SalesDepartment.SalesGroupID"
    Flag =1
    LeftTable ="Item"
    RightTable ="qryUPCItemMax"
    Expression ="Item.ID=qryUPCItemMax.ItemID"
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
    Begin
        dbText "Name" ="Ext"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1435
    Bottom =677
    Left =-1
    Top =-1
    Right =1403
    Bottom =178
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =648
        Top =8
        Right =744
        Bottom =122
        Top =0
        Name ="SalesGroup"
        Name =""
    End
    Begin
        Left =406
        Top =13
        Right =567
        Bottom =172
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =243
        Top =14
        Right =339
        Bottom =173
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =60
        Top =-3
        Right =156
        Bottom =96
        Top =0
        Name ="qryUPCItemMax"
        Name =""
    End
End
