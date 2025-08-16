Operation =1
Option =0
Where ="(((ClubOrderGuides.ClubID)=[parClubID]) AND ((IIf([SalesDepartmentID]=1,-1,0))=["
    "parLiquor]))"
Begin InputTables
    Name ="Item"
    Name ="qryItemsSalePriceNow"
    Name ="ClubOrderGuides"
End
Begin OutputColumns
    Expression ="ClubOrderGuides.ClubID"
    Expression ="ClubOrderGuides.OrderID"
    Expression ="ClubOrderGuides.LineID"
    Expression ="ClubOrderGuides.ItemID"
    Expression ="ClubOrderGuides.Par"
    Expression ="Item.Description"
    Expression ="Item.Price"
    Alias ="SPrice"
    Expression ="IIf(IsNull([SalePrice]),Item.Price,[SalePrice])"
    Alias ="SP"
    Expression ="IIf(IsNull([SalePrice]),0,-1)"
    Expression ="Item.Cost"
    Expression ="ClubOrderGuides.QuotedPrice"
    Alias ="Liquor"
    Expression ="IIf([SalesDepartmentID]=1,-1,0)"
End
Begin Parameters
    Name ="parClubID"
    Flag =4
    Name ="parLiquor"
    Flag =1
End
Begin Joins
    LeftTable ="Item"
    RightTable ="qryItemsSalePriceNow"
    Expression ="Item.ID = qryItemsSalePriceNow.ItemID"
    Flag =2
    LeftTable ="Item"
    RightTable ="ClubOrderGuides"
    Expression ="Item.ID = ClubOrderGuides.ItemID"
    Flag =1
End
Begin OrderBy
    Expression ="ClubOrderGuides.LineID"
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
        dbText "Name" ="SPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Liquor"
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
    ColumnsShown =539
    Begin
        Left =213
        Top =10
        Right =309
        Bottom =124
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =352
        Top =8
        Right =448
        Bottom =122
        Top =0
        Name ="qryItemsSalePriceNow"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =161
        Bottom =120
        Top =0
        Name ="ClubOrderGuides"
        Name =""
    End
End
