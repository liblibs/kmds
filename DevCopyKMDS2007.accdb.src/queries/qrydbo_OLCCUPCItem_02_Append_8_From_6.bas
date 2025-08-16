Operation =3
Name ="dbo_OLCCScanCode"
Option =0
Where ="((((Len([upccode])))=6))"
Begin InputTables
    Name ="dbo_OLCCScanCode"
End
Begin OutputColumns
    Name ="Code"
    Expression ="dbo_OLCCScanCode.Code"
    Alias ="Expr1"
    Name ="UPCCode"
    Expression ="upcecheckdigit(\"0\" & [upccode])"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="dbo_OLCCScanCode.UPCCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_OLCCScanCode.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbInteger "ColumnWidth" ="1575"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =27
    Top =0
    Right =1273
    Bottom =846
    Left =-1
    Top =-1
    Right =1223
    Bottom =576
    Left =0
    Top =0
    ColumnsShown =651
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
