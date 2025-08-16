Operation =1
Option =0
Begin InputTables
    Name ="Item"
    Name ="qryUPCItemCountPerItemID"
    Name ="SalesDepartment"
    Name ="SubCategory"
End
Begin OutputColumns
    Expression ="SubCategory.Description"
    Expression ="qryUPCItemCountPerItemID.Count"
    Expression ="Item.ID"
    Expression ="Item.Description"
    Expression ="Item.SalesDepartmentID"
    Alias ="SDesc"
    Expression ="SalesDepartment.Description"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="qryUPCItemCountPerItemID"
    Expression ="Item.ID = qryUPCItemCountPerItemID.ItemID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SalesDepartment"
    Expression ="Item.SalesDepartmentID = SalesDepartment.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SubCategory"
    Expression ="Item.SubCategoryID = SubCategory.ID"
    Flag =1
End
Begin OrderBy
    Expression ="SubCategory.Description"
    Flag =0
    Expression ="qryUPCItemCountPerItemID.Count"
    Flag =1
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
        dbText "Name" ="SDesc"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =32
        Top =27
        Right =128
        Bottom =141
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =90
        Top =0
        Name ="qryUPCItemCountPerItemID"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =120
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =695
        Top =1
        Right =791
        Bottom =115
        Top =0
        Name ="SubCategory"
        Name =""
    End
End
