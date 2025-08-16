Operation =4
Option =0
Where ="(((Item.AvgTurns)=999))"
Begin InputTables
    Name ="Item"
    Name ="KitItem"
End
Begin OutputColumns
    Name ="Item.AvgTurns"
    Expression ="0"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="KitItem"
    Expression ="Item.ID=KitItem.ItemID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Item.AvgTurns"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KitItem.KitID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-414
    Top =69
    Right =929
    Bottom =1537
    Left =-1
    Top =-1
    Right =1311
    Bottom =1184
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =272
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =242
        Top =12
        Right =384
        Bottom =458
        Top =0
        Name ="KitItem"
        Name =""
    End
End
