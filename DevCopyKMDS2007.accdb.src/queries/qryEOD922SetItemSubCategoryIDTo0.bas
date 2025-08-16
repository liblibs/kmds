Operation =4
Option =0
Where ="(((SubCategory.ID) Is Null))"
Begin InputTables
    Name ="Item"
    Name ="SubCategory"
End
Begin OutputColumns
    Name ="Item.SubCategoryID"
    Expression ="0"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="SubCategory"
    Expression ="Item.SubCategoryID = SubCategory.ID"
    Flag =2
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
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1410
    Bottom =877
    Left =-1
    Top =-1
    Right =1378
    Bottom =520
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =198
        Bottom =289
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =297
        Top =0
        Name ="SubCategory"
        Name =""
    End
End
