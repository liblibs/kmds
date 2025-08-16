Operation =4
Option =0
Where ="(((NabcaOlccItem.OldCode) Like \"*B\"))"
Begin InputTables
    Name ="NabcaOlccItem"
End
Begin OutputColumns
    Name ="NabcaOlccItem.ItemID"
    Expression ="CLng(Left([oldcode],5) & \"2\")"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="NabcaOlccItem.OldCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CodeNumeric"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NabcaOlccItem.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="nc"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1042
    Bottom =846
    Left =-1
    Top =-1
    Right =1010
    Bottom =340
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =305
        Top =0
        Name ="NabcaOlccItem"
        Name =""
    End
End
