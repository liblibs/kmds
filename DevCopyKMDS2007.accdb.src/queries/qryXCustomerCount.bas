Operation =1
Option =2
Where ="(((InvoLine.TransType)=\"S\"))"
Begin InputTables
    Name ="InvoLine"
End
Begin OutputColumns
    Expression ="InvoLine.CashSysID"
    Expression ="InvoLine.CashierID"
    Expression ="InvoLine.InvoHeadID"
    Alias ="count"
    Expression ="Sum(1)"
    Alias ="Sales"
    Expression ="Sum(InvoLine.Price)"
End
Begin Groups
    Expression ="InvoLine.CashSysID"
    GroupLevel =0
    Expression ="InvoLine.CashierID"
    GroupLevel =0
    Expression ="InvoLine.InvoHeadID"
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
        dbText "Name" ="count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sales"
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
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
