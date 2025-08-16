Operation =5
Option =0
Where ="(((SalesPicturesTemp.FileName) Is Null) AND ((SalesPictures.ID)>0))"
Begin InputTables
    Name ="SalesPicturesTemp"
    Name ="SalesPictures"
End
Begin OutputColumns
    Expression ="SalesPicturesTemp.FileName"
    Expression ="SalesPictures.*"
    Expression ="SalesPictures.ID"
End
Begin Joins
    LeftTable ="SalesPicturesTemp"
    RightTable ="SalesPictures"
    Expression ="SalesPicturesTemp.FileName = SalesPictures.FileName"
    Flag =3
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="SalesPicturesTemp.FileName"
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
    Bottom =224
    Left =0
    Top =0
    ColumnsShown =771
    Begin
        Left =436
        Top =24
        Right =580
        Bottom =168
        Top =0
        Name ="SalesPicturesTemp"
        Name =""
    End
    Begin
        Left =57
        Top =13
        Right =201
        Bottom =157
        Top =0
        Name ="SalesPictures"
        Name =""
    End
End
