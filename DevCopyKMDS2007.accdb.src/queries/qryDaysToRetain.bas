Operation =1
Option =0
Where ="(((Settings.KeyName)=\"DaysToRetain\"))"
Begin InputTables
    Name ="Settings"
End
Begin OutputColumns
    Alias ="Days"
    Expression ="IIf((CLng([Value]) Between 500 And 1000),CLng([Value]),500)"
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
        dbText "Name" ="Days"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-421
    Top =128
    Right =1068
    Bottom =865
    Left =-1
    Top =-1
    Right =1466
    Bottom =387
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Settings"
        Name =""
    End
End
