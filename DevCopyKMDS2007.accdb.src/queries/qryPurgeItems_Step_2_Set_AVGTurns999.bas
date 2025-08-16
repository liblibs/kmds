Operation =4
Option =0
Where ="(((InvoLine.Code) Is Null) AND ((Item.Onhand)=0) AND ((Item.SetupDate)<Now()-60)"
    " AND ((ClubOrderGuides.ClubID) Is Null))"
Begin InputTables
    Name ="Item"
    Name ="InvoLine"
    Name ="ClubOrderGuides"
End
Begin OutputColumns
    Name ="Item.AvgTurns"
    Expression ="999"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="InvoLine"
    Expression ="Item.ID = InvoLine.Code"
    Flag =2
    LeftTable ="Item"
    RightTable ="ClubOrderGuides"
    Expression ="Item.ID = ClubOrderGuides.ItemID"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Item.DateLastSold"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =142
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =336
        Top =-1
        Right =432
        Bottom =113
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =204
        Top =62
        Right =300
        Bottom =176
        Top =0
        Name ="ClubOrderGuides"
        Name =""
    End
End
