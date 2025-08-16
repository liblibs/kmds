Operation =3
Name ="UPCItem"
Option =0
Where ="(((UPCItem.ItemID) Is Null))"
Begin InputTables
    Name ="UPCItem"
    Name ="qryMetrologicForOLCCShelfTags"
End
Begin OutputColumns
    Name ="ItemID"
    Expression ="qryMetrologicForOLCCShelfTags.ID"
    Name ="ID"
    Expression ="qryMetrologicForOLCCShelfTags.NewUPC"
End
Begin Joins
    LeftTable ="qryMetrologicForOLCCShelfTags"
    RightTable ="UPCItem"
    Expression ="qryMetrologicForOLCCShelfTags.NewUPC = UPCItem.ID"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="qryMetrologicForOLCCShelfTags.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryMetrologicForOLCCShelfTags.NewUPC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCItem.ItemID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-305
    Top =108
    Right =1166
    Bottom =873
    Left =-1
    Top =-1
    Right =1439
    Bottom =482
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =136
        Top =63
        Right =280
        Bottom =207
        Top =0
        Name ="UPCItem"
        Name =""
    End
    Begin
        Left =406
        Top =51
        Right =550
        Bottom =195
        Top =0
        Name ="qryMetrologicForOLCCShelfTags"
        Name =""
    End
End
