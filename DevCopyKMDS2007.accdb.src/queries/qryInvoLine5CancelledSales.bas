Operation =1
Option =0
Where ="(((InvoLine5.Code)=[CancelTenderID]))"
Begin InputTables
    Name ="InvoLine5"
    Name ="Sysfile"
End
Begin OutputColumns
    Expression ="InvoLine5.CashierID"
    Expression ="InvoLine5.SetupDate"
    Expression ="InvoLine5.InvoHeadID"
    Expression ="InvoLine5.CashSysID"
    Expression ="InvoLine5.ID"
    Alias ="MaxID"
    Expression ="InvoLine5.Id\\2"
    Expression ="InvoLine5.Code"
    Expression ="InvoLine5.InputCode"
    Expression ="Sysfile.CancelTenderID"
End
Begin OrderBy
    Expression ="InvoLine5.CashierID"
    Flag =0
    Expression ="InvoLine5.SetupDate"
    Flag =0
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
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxID"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =293
        Bottom =120
        Top =0
        Name ="Sysfile"
        Name =""
    End
End
