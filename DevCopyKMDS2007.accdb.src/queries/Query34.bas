Operation =4
Option =0
Where ="(((crescentUPC.lid)<>[crescentUPC].[Itemid] And (crescentUPC.lid)>0))"
Begin InputTables
    Name ="crescentUPC"
    Name ="UPCItem"
End
Begin OutputColumns
    Name ="UPCItem.ID"
    Expression ="UPCECheckdigit(Trim([crescentUPC].[ID]))"
End
Begin Joins
    LeftTable ="crescentUPC"
    RightTable ="UPCItem"
    Expression ="crescentUPC.ID = UPCItem.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="UPCItem.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCItem.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="crescentUPC.lid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="crescentUPC.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="crescentUPC.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="crescentUPC.UPCE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="upce"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =2
    Left =-4
    Top =-23
    Right =868
    Bottom =517
    Left =-1
    Top =-1
    Right =849
    Bottom =272
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =86
        Top =52
        Right =230
        Bottom =196
        Top =0
        Name ="crescentUPC"
        Name =""
    End
    Begin
        Left =474
        Top =68
        Right =618
        Bottom =212
        Top =0
        Name ="UPCItem"
        Name =""
    End
End
