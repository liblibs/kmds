Operation =1
Option =0
Begin InputTables
    Name ="MetrologicScanbatch"
    Name ="UPCItem"
    Name ="Item"
    Name ="SalesDepartment"
    Name ="SubCategory"
End
Begin OutputColumns
    Expression ="MetrologicScanbatch.ID"
    Expression ="UPCItem.ItemID"
    Expression ="Item.Description"
    Expression ="MetrologicScanbatch.ScanCode"
    Expression ="MetrologicScanbatch.Quantity"
    Expression ="Item.Price"
    Expression ="MetrologicScanbatch.ScanDate"
    Expression ="MetrologicScanbatch.ScanTime"
    Alias ="SalesDptDescr"
    Expression ="SalesDepartment.Description"
    Alias ="CategoryDescr"
    Expression ="SubCategory.Description"
End
Begin Joins
    LeftTable ="MetrologicScanbatch"
    RightTable ="UPCItem"
    Expression ="MetrologicScanbatch.ScanCode=UPCItem.ID"
    Flag =1
    LeftTable ="UPCItem"
    RightTable ="Item"
    Expression ="UPCItem.ItemID=Item.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SalesDepartment"
    Expression ="Item.SalesDepartmentID=SalesDepartment.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SubCategory"
    Expression ="Item.SubCategoryID=SubCategory.ID"
    Flag =1
End
Begin OrderBy
    Expression ="MetrologicScanbatch.ScanDate"
    Flag =0
    Expression ="MetrologicScanbatch.ScanTime"
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
        dbText "Name" ="MetrologicScanbatch.Quantity"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MetrologicScanbatch.ScanCode"
        dbInteger "ColumnWidth" ="2130"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MetrologicScanbatch.ScanDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MetrologicScanbatch.ScanTime"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MetrologicScanbatch.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCItem.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbInteger "ColumnWidth" ="2445"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDptDescr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CategoryDescr"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =13
    Top =31
    Right =1149
    Bottom =779
    Left =-1
    Top =-1
    Right =1104
    Bottom =324
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =-6
        Right =192
        Bottom =223
        Top =0
        Name ="MetrologicScanbatch"
        Name =""
    End
    Begin
        Left =316
        Top =25
        Right =460
        Bottom =169
        Top =0
        Name ="UPCItem"
        Name =""
    End
    Begin
        Left =508
        Top =12
        Right =652
        Bottom =156
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =508
        Top =156
        Right =652
        Bottom =300
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =700
        Top =168
        Right =844
        Bottom =312
        Top =0
        Name ="SubCategory"
        Name =""
    End
End
