Operation =3
Name ="dbo_CRReports"
Option =0
Where ="(((dbo_CRReports.ID) Is Null))"
Begin InputTables
    Name ="CRReports"
    Name ="dbo_CRReports"
End
Begin OutputColumns
    Name ="ReportName"
    Expression ="CRReports.ReportName"
    Name ="ZCopies"
    Expression ="CRReports.ZCopies"
End
Begin Joins
    LeftTable ="CRReports"
    RightTable ="dbo_CRReports"
    Expression ="CRReports.ReportName = dbo_CRReports.ReportName"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
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
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="CRReports"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="dbo_CRReports"
        Name =""
    End
End
