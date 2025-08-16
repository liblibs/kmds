Operation =3
Name ="ClubOrderGuides"
Option =0
Where ="(((ClubOrderGuideTemp.ClubID)=[parClubID]))"
Begin InputTables
    Name ="ClubOrderGuideTemp"
    Name ="Item"
End
Begin OutputColumns
    Name ="ClubID"
    Expression ="ClubOrderGuideTemp.ClubID"
    Name ="LineID"
    Expression ="ClubOrderGuideTemp.LineID"
    Name ="ItemID"
    Expression ="ClubOrderGuideTemp.ItemID"
    Name ="Par"
    Expression ="ClubOrderGuideTemp.Par"
    Name ="QuotedPrice"
    Expression ="ClubOrderGuideTemp.QuotedPrice"
    Alias ="LI"
    Name ="LiquorItem"
    Expression ="IIf([SalesDepartmentID]=1,1,0)"
End
Begin Joins
    LeftTable ="ClubOrderGuideTemp"
    RightTable ="Item"
    Expression ="ClubOrderGuideTemp.ItemID = Item.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="LI"
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
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="ClubOrderGuideTemp"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
End
