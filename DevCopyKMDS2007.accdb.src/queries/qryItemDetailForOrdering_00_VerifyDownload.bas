Operation =1
Option =0
Where ="(((ItemDetailForOrdering.F2) Like \"999*\"))"
Begin InputTables
    Name ="ItemDetailForOrdering"
End
Begin OutputColumns
    Expression ="ItemDetailForOrdering.F1"
    Expression ="ItemDetailForOrdering.F2"
    Expression ="ItemDetailForOrdering.F3"
    Expression ="ItemDetailForOrdering.F4"
    Expression ="ItemDetailForOrdering.F5"
    Expression ="ItemDetailForOrdering.F6"
    Expression ="ItemDetailForOrdering.F7"
    Expression ="ItemDetailForOrdering.F8"
    Expression ="ItemDetailForOrdering.F9"
    Expression ="ItemDetailForOrdering.F10"
    Expression ="ItemDetailForOrdering.F11"
    Expression ="ItemDetailForOrdering.F12"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="ItemDetailForOrdering.F2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F11"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-248
    Top =58
    Right =893
    Bottom =799
    Left =-1
    Top =-1
    Right =1109
    Bottom =389
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="ItemDetailForOrdering"
        Name =""
    End
End
