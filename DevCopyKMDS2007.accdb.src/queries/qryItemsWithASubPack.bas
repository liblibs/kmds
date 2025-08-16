Operation =1
Option =0
Where ="(((Item.SalesDepartmentID) Between Forms!PrintInventoryReport!cboSalesDepartment"
    "IDBegin And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd) And ((Item.Units"
    "PerPack)>1))"
Begin InputTables
    Name ="Item"
    Name ="UPCItem"
    Name ="Item"
    Alias ="Item_1"
    Name ="SalesDepartment"
    Name ="SalesDepartment"
    Alias ="SalesDepartment_1"
End
Begin OutputColumns
    Alias ="CartonSD"
    Expression ="Item.SalesDepartmentID"
    Alias ="CartonSDDesc"
    Expression ="SalesDepartment.Description"
    Alias ="CartonID"
    Expression ="Item.ID"
    Alias ="CartonDesc"
    Expression ="Item.Description"
    Alias ="CartonSUBUpc"
    Expression ="Item.SubPack"
    Alias ="CartonPack"
    Expression ="Item.UnitsPerPack"
    Alias ="CartonPrice"
    Expression ="Item.Price"
    Alias ="PackSD"
    Expression ="Item_1.SalesDepartmentID"
    Alias ="PackSDDesc"
    Expression ="SalesDepartment_1.Description"
    Alias ="PackID"
    Expression ="Item_1.ID"
    Alias ="PackDesc"
    Expression ="Item_1.Description"
    Alias ="PackPrice"
    Expression ="Item_1.Price"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="UPCItem"
    Expression ="Item.SubPack = UPCItem.ID"
    Flag =2
    LeftTable ="UPCItem"
    RightTable ="Item_1"
    Expression ="UPCItem.ItemID = Item_1.ID"
    Flag =2
    LeftTable ="Item"
    RightTable ="SalesDepartment"
    Expression ="Item.SalesDepartmentID = SalesDepartment.ID"
    Flag =1
    LeftTable ="Item_1"
    RightTable ="SalesDepartment_1"
    Expression ="Item_1.SalesDepartmentID = SalesDepartment_1.ID"
    Flag =2
End
Begin OrderBy
    Expression ="Item.SalesDepartmentID"
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
        dbText "Name" ="CartonSDDesc"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CartonID"
        dbInteger "ColumnWidth" ="930"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CartonDesc"
        dbInteger "ColumnWidth" ="3180"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CartonSUBUpc"
        dbInteger "ColumnWidth" ="1500"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CartonPack"
        dbInteger "ColumnWidth" ="1200"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PackSD"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PackSDDesc"
        dbInteger "ColumnWidth" ="1680"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PackDesc"
        dbInteger "ColumnWidth" ="3435"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PackID"
        dbInteger "ColumnWidth" ="810"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CartonSD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CartonPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PackPrice"
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
    Bottom =247
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =169
        Bottom =240
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =467
        Top =47
        Right =563
        Bottom =161
        Top =0
        Name ="UPCItem"
        Name =""
    End
    Begin
        Left =652
        Top =13
        Right =748
        Bottom =187
        Top =0
        Name ="Item_1"
        Name =""
    End
    Begin
        Left =207
        Top =6
        Right =303
        Bottom =120
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =795
        Top =35
        Right =891
        Bottom =149
        Top =0
        Name ="SalesDepartment_1"
        Name =""
    End
End
