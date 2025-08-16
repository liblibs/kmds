Operation =1
Option =0
Where ="(((Item.Onhand)<>0) AND ((Item.Deposit)<>0))"
Begin InputTables
    Name ="SalesDepartment"
    Name ="Item"
    Name ="SubCategory"
End
Begin OutputColumns
    Expression ="Item.SalesDepartmentID"
    Expression ="SalesDepartment.Description"
    Expression ="Item.ID"
    Expression ="Item.Description"
    Expression ="Item.Onhand"
    Expression ="Item.Deposit"
    Alias ="DepositTotal"
    Expression ="[onhand]*[deposit]"
    Alias ="SubCategoryDescription"
    Expression ="SubCategory.Description"
End
Begin Joins
    LeftTable ="SalesDepartment"
    RightTable ="Item"
    Expression ="SalesDepartment.ID=Item.SalesDepartmentID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SubCategory"
    Expression ="Item.SubCategoryID=SubCategory.ID"
    Flag =1
End
Begin OrderBy
    Expression ="Item.SalesDepartmentID"
    Flag =0
    Expression ="Item.Description"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.Description"
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
        dbText "Name" ="Item.Deposit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DepositTotal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubCategoryDescription"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-403
    Top =69
    Right =1062
    Bottom =880
    Left =-1
    Top =-1
    Right =1433
    Bottom =195
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="SubCategory"
        Name =""
    End
End
