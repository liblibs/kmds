Operation =5
Option =0
Where ="(((Category.SalesDepartmentID)=1) AND ((SubCategory.ID)>0))"
Begin InputTables
    Name ="SubCategory"
    Name ="Category"
End
Begin OutputColumns
    Expression ="Category.SalesDepartmentID"
    Expression ="SubCategory.*"
    Expression ="SubCategory.ID"
End
Begin Joins
    LeftTable ="SubCategory"
    RightTable ="Category"
    Expression ="SubCategory.CategoryID = Category.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Category.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubCategory.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubCategory.CategoryID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubCategory.Description"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2295"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SubCategory.TargetGP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubCategory.Setup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubCategory.Change"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubCategory.upsize_ts"
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
    Bottom =261
    Left =0
    Top =0
    ColumnsShown =771
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="SubCategory"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =190
        Top =0
        Name ="Category"
        Name =""
    End
End
