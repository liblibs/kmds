Operation =1
Option =0
Where ="(((Len([ID]))>7) AND ((UPCItem.ID) Not Like \"99900*\"))"
Begin InputTables
    Name ="UPCItem"
End
Begin OutputColumns
    Expression ="UPCItem.ItemID"
    Alias ="UPCCodeMax"
    Expression ="Max(UPCItem.ID)"
    Alias ="Count"
    Expression ="Sum(1)"
End
Begin OrderBy
    Expression ="UPCItem.ItemID"
    Flag =0
    Expression ="Max(UPCItem.ID)"
    Flag =0
End
Begin Groups
    Expression ="UPCItem.ItemID"
    GroupLevel =0
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
        dbText "Name" ="UPCCodeMax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCItem.ItemID"
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
    Right =1193
    Bottom =108
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="UPCItem"
        Name =""
    End
End
