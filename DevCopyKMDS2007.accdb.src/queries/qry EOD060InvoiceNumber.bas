Operation =1
Option =0
Having ="(((InvoLine5.TransType)=\"D\"))"
Begin InputTables
    Name ="InvoLine5"
End
Begin OutputColumns
    Expression ="InvoLine5.InvoHeadID"
    Expression ="InvoLine5.CashSysID"
    Expression ="InvoLine5.TransType"
    Alias ="InvoiceNumber"
    Expression ="InvoLine5.InputCode"
End
Begin Groups
    Expression ="InvoLine5.InvoHeadID"
    GroupLevel =0
    Expression ="InvoLine5.CashSysID"
    GroupLevel =0
    Expression ="InvoLine5.TransType"
    GroupLevel =0
    Expression ="InvoLine5.InputCode"
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
        dbText "Name" ="InvoiceNumber"
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
        Name ="InvoLine5"
        Name =""
    End
End
