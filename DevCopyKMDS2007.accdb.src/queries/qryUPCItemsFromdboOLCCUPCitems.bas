Operation =3
Name ="UPCItem"
Option =0
Where ="(((UPCItem.ID) Is Null))"
Begin InputTables
    Name ="dbo_OLCCScanCode"
    Name ="Item"
    Name ="UPCItem"
End
Begin OutputColumns
    Name ="ID"
    Expression ="dbo_OLCCScanCode.UPCCode"
    Name ="ItemID"
    Expression ="dbo_OLCCScanCode.Code"
End
Begin Joins
    LeftTable ="dbo_OLCCScanCode"
    RightTable ="Item"
    Expression ="dbo_OLCCScanCode.Code = Item.ID"
    Flag =1
    LeftTable ="dbo_OLCCScanCode"
    RightTable ="UPCItem"
    Expression ="dbo_OLCCScanCode.UPCCode = UPCItem.ID"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="dbo_OLCCScanCode.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_OLCCScanCode.UPCCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_OLCCScanCode.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCItem.ID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =61
    Top =38
    Right =1244
    Bottom =748
    Left =-1
    Top =-1
    Right =1160
    Bottom =440
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =197
        Bottom =371
        Top =0
        Name ="dbo_OLCCScanCode"
        Name =""
    End
    Begin
        Left =508
        Top =164
        Right =652
        Bottom =308
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =467
        Top =35
        Right =611
        Bottom =179
        Top =0
        Name ="UPCItem"
        Name =""
    End
End
