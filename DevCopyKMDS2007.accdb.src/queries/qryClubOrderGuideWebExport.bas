Operation =1
Option =0
Where ="(((ClubOrderGuides.ClubID)=[parClubID]))"
Begin InputTables
    Name ="ClubOrderGuides"
    Name ="Club"
End
Begin OutputColumns
    Expression ="ClubOrderGuides.ClubID"
    Expression ="ClubOrderGuides.OrderID"
    Expression ="ClubOrderGuides.LineID"
    Expression ="ClubOrderGuides.ItemID"
    Expression ="ClubOrderGuides.Description"
    Expression ="ClubOrderGuides.Par"
    Expression ="ClubOrderGuides.QuotedPrice"
    Expression ="Club.NonLiquor"
    Expression ="ClubOrderGuides.LiquorItem"
    Expression ="Club.Contact"
    Expression ="Club.MaintainParLevels"
End
Begin Parameters
    Name ="parClubID"
    Flag =4
End
Begin Joins
    LeftTable ="ClubOrderGuides"
    RightTable ="Club"
    Expression ="ClubOrderGuides.ClubID = Club.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="ClubOrderGuides.ClubID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubOrderGuides.OrderID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubOrderGuides.LineID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubOrderGuides.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubOrderGuides.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubOrderGuides.Par"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubOrderGuides.QuotedPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.NonLiquor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubOrderGuides.LiquorItem"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.Contact"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.MaintainParLevels"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-7
    Top =29
    Right =1446
    Bottom =782
    Left =-1
    Top =-1
    Right =1430
    Bottom =451
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="ClubOrderGuides"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Club"
        Name =""
    End
End
