Operation =4
Option =0
Begin InputTables
    Name ="Item"
    Name ="OLCC Categories"
    Name ="SubCategory"
End
Begin OutputColumns
    Name ="Item.SubCategoryID"
    Expression ="[SubCategory].[ID]"
End
Begin Joins
    LeftTable ="OLCC Categories"
    RightTable ="SubCategory"
    Expression ="[OLCC Categories].Category = SubCategory.Description"
    Flag =1
    LeftTable ="Item"
    RightTable ="OLCC Categories"
    Expression ="Item.ID = [OLCC Categories].Code"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SubCategoryID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubCategory.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubCategory.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Description"
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
    Bottom =295
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =210
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="OLCC Categories"
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
