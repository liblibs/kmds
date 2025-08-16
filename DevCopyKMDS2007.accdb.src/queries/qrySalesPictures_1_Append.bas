Operation =3
Name ="SalesPictures"
Option =0
Where ="(((SalesPictures.ID) Is Null))"
Begin InputTables
    Name ="SalesPictures"
    Name ="SalesPicturesTemp"
End
Begin OutputColumns
    Name ="FileName"
    Expression ="SalesPicturesTemp.FileName"
End
Begin Joins
    LeftTable ="SalesPictures"
    RightTable ="SalesPicturesTemp"
    Expression ="SalesPictures.FileName = SalesPicturesTemp.FileName"
    Flag =3
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
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
    Begin
        dbText "Name" ="SalesPictures.FileName"
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
    Bottom =186
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =487
        Top =24
        Right =631
        Bottom =168
        Top =0
        Name ="SalesPictures"
        Name =""
    End
    Begin
        Left =76
        Top =30
        Right =220
        Bottom =174
        Top =0
        Name ="SalesPicturesTemp"
        Name =""
    End
End
