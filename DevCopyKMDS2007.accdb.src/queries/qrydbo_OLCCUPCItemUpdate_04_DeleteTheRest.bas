Operation =5
Option =0
Where ="(((dbo_OLCCScanCode.UPCCodeKMDS) Is Null))"
Begin InputTables
    Name ="dbo_OLCCScanCode"
End
Begin OutputColumns
    Expression ="dbo_OLCCScanCode.UPCCodeKMDS"
    Expression ="dbo_OLCCScanCode.*"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="lid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="upc10"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCScanCode.Field3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Len([upccode])"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_OLCCScanCode.Field3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_OLCCScanCode.UPCCodeKMDS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCScanCode.UPCCode"
        dbInteger "ColumnWidth" ="3000"
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
    Bottom =175
    Left =0
    Top =0
    ColumnsShown =771
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="dbo_OLCCScanCode"
        Name =""
    End
End
