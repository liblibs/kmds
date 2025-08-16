Operation =1
Option =0
Where ="(((InvoLine5.TransType)=\"T\"))"
Begin InputTables
    Name ="InvoLine5"
End
Begin OutputColumns
    Expression ="InvoLine5.InvoHeadID"
    Expression ="InvoLine5.CashSysID"
    Expression ="InvoLine5.SetupDate"
    Alias ="TenderCode"
    Expression ="First(InvoLine5.Code)"
End
Begin OrderBy
    Expression ="First(InvoLine5.Code)"
    Flag =0
End
Begin Groups
    Expression ="InvoLine5.InvoHeadID"
    GroupLevel =0
    Expression ="InvoLine5.CashSysID"
    GroupLevel =0
    Expression ="InvoLine5.SetupDate"
    GroupLevel =0
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
        dbText "Name" ="InvoLine5.SetupDate"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TenderCode"
        dbInteger "ColumnWidth" ="1635"
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
    Bottom =278
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =172
        Top =6
        Right =306
        Bottom =271
        Top =0
        Name ="InvoLine5"
        Name =""
    End
End
