Operation =1
Option =0
Where ="((([UPCItemGH].[ID])>\"19\"))"
Having ="(((Len([id]))=12))"
Begin InputTables
    Name ="UPCItemGH"
End
Begin OutputColumns
    Alias ="lid"
    Expression ="Len([id])"
    Alias ="Vendor"
    Expression ="Left([id],6)"
    Alias ="Vendor5"
    Expression ="Mid([id],2,5)"
End
Begin Groups
    Expression ="Len([id])"
    GroupLevel =0
    Expression ="Left([id],6)"
    GroupLevel =0
    Expression ="Mid([id],2,5)"
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
        dbText "Name" ="Vendor"
        dbInteger "ColumnWidth" ="1275"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor5"
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
    Bottom =151
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =267
        Bottom =122
        Top =0
        Name ="UPCItemGH"
        Name =""
    End
End
