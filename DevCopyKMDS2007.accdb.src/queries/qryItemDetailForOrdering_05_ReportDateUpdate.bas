Operation =4
Option =0
Begin InputTables
    Name ="OLASItemDetailForOrdering"
    Name ="OLASiItemDetailForOrderingReportDate"
End
Begin OutputColumns
    Name ="OLASItemDetailForOrdering.ReportDate"
    Expression ="[OLASiItemDetailForOrderingReportDate].[ReportDate]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="OLASItemDetailForOrdering.ReportDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemDetailForOrderingReportDate.ReportDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-96
    Top =159
    Right =1019
    Bottom =965
    Left =-1
    Top =-1
    Right =1083
    Bottom =522
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =314
        Top =0
        Name ="OLASItemDetailForOrdering"
        Name =""
    End
    Begin
        Left =491
        Top =124
        Right =635
        Bottom =268
        Top =0
        Name ="OLASiItemDetailForOrderingReportDate"
        Name =""
    End
End
