Operation =3
Name ="OLASiItemDetailForOrderingReportDate"
Option =0
Having ="(((ItemDetailForOrdering.F3) Like \"Report Date*\"))"
Begin InputTables
    Name ="ItemDetailForOrdering"
End
Begin OutputColumns
    Name ="F3"
    Expression ="ItemDetailForOrdering.F3"
    Alias ="ReportDate"
    Name ="ReportDate"
    Expression ="First(Right([F3],10))"
End
Begin Groups
    Expression ="ItemDetailForOrdering.F3"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="OLASItemDetailForOrdering.ReportDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemDetailForOrdering.F3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ReportDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =77
    Top =162
    Right =1114
    Bottom =872
    Left =-1
    Top =-1
    Right =1005
    Bottom =426
    Left =0
    Top =0
    ColumnsShown =655
    Begin
        Left =79
        Top =20
        Right =223
        Bottom =164
        Top =0
        Name ="ItemDetailForOrdering"
        Name =""
    End
End
