Operation =3
Name ="UPCItem"
Option =0
Where ="(((UPCItem.ID) Is Null))"
Begin InputTables
    Name ="UPCItem"
    Name ="ItemUPCID"
End
Begin OutputColumns
    Name ="ItemID"
    Expression ="ItemUPCID.ItemID"
    Name ="ID"
    Expression ="ItemUPCID.UPCID"
End
Begin Joins
    LeftTable ="ItemUPCID"
    RightTable ="UPCItem"
    Expression ="ItemUPCID.UPCID = UPCItem.ID"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCItem.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryUPCItemsAppendItemID.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryUPCItemsAppendItemID.UPCItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemUPCID.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemUPCID.UPCID"
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
    Bottom =212
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =311
        Top =45
        Right =455
        Bottom =189
        Top =0
        Name ="UPCItem"
        Name =""
    End
    Begin
        Left =56
        Top =25
        Right =200
        Bottom =169
        Top =0
        Name ="ItemUPCID"
        Name =""
    End
End
