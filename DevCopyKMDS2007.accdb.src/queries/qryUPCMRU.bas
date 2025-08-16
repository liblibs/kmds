Operation =1
Option =0
Having ="(((Len([id]))=8 Or (Len([id]))=12))"
Begin InputTables
    Name ="UPCItem"
End
Begin OutputColumns
    Alias ="FirstOfChangeDate"
    Expression ="First(UPCItem.ChangeDate)"
    Expression ="UPCItem.ItemID"
    Expression ="UPCItem.ID"
    Alias ="lid"
    Expression ="Len([id])"
End
Begin OrderBy
    Expression ="First(UPCItem.ChangeDate)"
    Flag =1
End
Begin Groups
    Expression ="UPCItem.ItemID"
    GroupLevel =0
    Expression ="UPCItem.ID"
    GroupLevel =0
    Expression ="Len([id])"
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
        dbText "Name" ="UPCItem.ChangeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCItem.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCItem.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfChangeDate"
        dbInteger "ColumnWidth" ="2175"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =39
    Top =88
    Right =1180
    Bottom =829
    Left =-1
    Top =-1
    Right =1109
    Bottom =457
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="UPCItem"
        Name =""
    End
End
