Operation =1
Option =0
Where ="(((Len([ID]))>7))"
Begin InputTables
    Name ="UPCItem"
End
Begin OutputColumns
    Expression ="UPCItem.ItemID"
    Alias ="FirstOfChangeDate"
    Expression ="First(UPCItem.ChangeDate)"
    Alias ="FirstOfID"
    Expression ="First(UPCItem.ID)"
End
Begin OrderBy
    Expression ="UPCItem.ItemID"
    Flag =0
    Expression ="First(UPCItem.ChangeDate)"
    Flag =1
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
        dbText "Name" ="FirstOfID"
        dbInteger "ColumnWidth" ="2715"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfChangeDate"
        dbInteger "ColumnWidth" ="2625"
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
    Bottom =110
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
