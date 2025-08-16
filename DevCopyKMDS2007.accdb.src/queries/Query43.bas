Operation =4
Option =0
Begin InputTables
    Name ="ScioInitialOrder"
    Name ="Item"
    Name ="SubCategory"
End
Begin OutputColumns
    Name ="Item.SubCategoryID"
    Expression ="[SubCategory].[id]"
End
Begin Joins
    LeftTable ="ScioInitialOrder"
    RightTable ="Item"
    Expression ="ScioInitialOrder.Code = Item.ID"
    Flag =1
    LeftTable ="ScioInitialOrder"
    RightTable ="SubCategory"
    Expression ="ScioInitialOrder.Category = SubCategory.Description"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="ScioInitialOrder.Category"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="cat"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubCategory.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SubCategoryID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =3
    Top =206
    Right =1534
    Bottom =1039
    Left =-1
    Top =-1
    Right =1508
    Bottom =563
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="ScioInitialOrder"
        Name =""
    End
    Begin
        Left =332
        Top =189
        Right =476
        Bottom =516
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =547
        Top =45
        Right =691
        Bottom =189
        Top =0
        Name ="SubCategory"
        Name =""
    End
End
