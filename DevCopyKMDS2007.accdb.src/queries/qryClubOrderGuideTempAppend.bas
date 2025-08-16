Operation =3
Name ="ClubOrderGuideTemp"
Option =0
Where ="(((ClubOrderGuides.ClubID)=[parClubID]))"
Begin InputTables
    Name ="ClubOrderGuides"
    Name ="Item"
End
Begin OutputColumns
    Name ="LineId"
    Expression ="ClubOrderGuides.LineID"
    Name ="ItemID"
    Expression ="ClubOrderGuides.ItemID"
    Name ="Par"
    Expression ="ClubOrderGuides.Par"
    Name ="ClubID"
    Expression ="ClubOrderGuides.ClubID"
    Name ="QuotedPrice"
    Expression ="ClubOrderGuides.QuotedPrice"
    Alias ="li"
    Name ="LiquorItem"
    Expression ="IIf([SalesDepartmentID]=1,1,0)"
End
Begin Parameters
    Name ="parClubID"
    Flag =4
End
Begin Joins
    LeftTable ="ClubOrderGuides"
    RightTable ="Item"
    Expression ="ClubOrderGuides.ItemID = Item.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="li"
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
        Right =201
        Bottom =120
        Top =0
        Name ="ClubOrderGuides"
        Name =""
    End
    Begin
        Left =239
        Top =6
        Right =335
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
End
