Operation =1
Option =0
Where ="(((Len([ID]))=10) AND ((IsNumeric([ID]))=True))"
Begin InputTables
    Name ="UPCItem"
End
Begin OutputColumns
    Expression ="UPCItem.ID"
    Expression ="UPCItem.ItemID"
    Alias ="LID"
    Expression ="Len([ID])"
    Alias ="LUN"
    Expression ="IsNumeric([ID])"
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
        dbText "Name" ="UPCItem.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCItem.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LUN"
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
    Bottom =377
    Left =0
    Top =0
    ColumnsShown =539
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
