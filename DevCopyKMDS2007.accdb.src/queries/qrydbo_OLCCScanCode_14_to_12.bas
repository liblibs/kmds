Operation =4
Option =0
Where ="(((Len([UPCCODE]))=14))"
Begin InputTables
    Name ="dbo_OLCCScanCode"
End
Begin OutputColumns
    Name ="dbo_OLCCScanCode.UPCCodeKMDS"
    Expression ="CHECKDIGIT11(Mid([UPCCODE],3,11))"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="dbo_OLCCScanCode.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="upc12"
        dbInteger "ColumnWidth" ="4065"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_OLCCScanCode.UPCCodeKMDS"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-194
    Top =136
    Right =1268
    Bottom =947
    Left =-1
    Top =-1
    Right =1430
    Bottom =528
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =372
        Bottom =337
        Top =0
        Name ="dbo_OLCCScanCode"
        Name =""
    End
End
