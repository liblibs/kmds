Operation =4
Option =0
Where ="(((Item.PictureID)<>0) AND ((SalesPictures.ID) Is Null))"
Begin InputTables
    Name ="Item"
    Name ="SalesPictures"
End
Begin OutputColumns
    Name ="Item.PictureID"
    Expression ="0"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="SalesPictures"
    Expression ="Item.PictureID = SalesPictures.ID"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Item.PictureID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesPictures.ID"
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
    Bottom =237
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =217
        Bottom =235
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =346
        Top =16
        Right =588
        Bottom =227
        Top =0
        Name ="SalesPictures"
        Name =""
    End
End
