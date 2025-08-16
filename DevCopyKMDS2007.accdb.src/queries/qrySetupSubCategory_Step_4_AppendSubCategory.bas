Operation =3
Name ="SubCategory"
Option =0
Begin InputTables
    Name ="OLCCWebItems"
End
Begin OutputColumns
    Name ="Description"
    Expression ="OLCCWebItems.Category"
    Alias ="CatID"
    Name ="CategoryID"
    Expression ="1"
End
Begin Groups
    Expression ="OLCCWebItems.Category"
    GroupLevel =0
    Expression ="1"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="CatID"
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
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =655
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="OLCCWebItems"
        Name =""
    End
End
