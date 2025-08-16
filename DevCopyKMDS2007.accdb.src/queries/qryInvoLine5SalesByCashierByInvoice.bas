Operation =1
Option =0
Begin InputTables
    Name ="InvoLine5"
    Name ="Sysfile"
End
Begin OutputColumns
    Expression ="InvoLine5.CashierID"
    Expression ="InvoLine5.CashSysID"
    Expression ="InvoLine5.InvoHeadID"
    Alias ="TransactionLines"
    Expression ="Sum(IIf([TransType]=\"S\",1,0))"
    Alias ="Sales"
    Expression ="Sum(IIf([TransType]=\"S\",[Price],0))"
    Alias ="Transactions"
    Expression ="Max(IIf([TransType]=\"S\",1,0))"
    Alias ="NoSales"
    Expression ="Max(IIf([TransType]=\"T\" And [CODE]=[NoSaleTenderID],1,0))"
    Alias ="CancelledSales"
    Expression ="Max(IIf([TransType]=\"T\" And [CODE]=[CancelTenderID],1,0))"
    Alias ="IDScanCount"
    Expression ="Sum(IIf([TransType]=\"T\" And [CODE]=[ScannedODLTenderID],1,0))"
End
Begin OrderBy
    Expression ="InvoLine5.CashSysID"
    Flag =0
    Expression ="InvoLine5.InvoHeadID"
    Flag =0
End
Begin Groups
    Expression ="InvoLine5.CashierID"
    GroupLevel =0
    Expression ="InvoLine5.CashSysID"
    GroupLevel =0
    Expression ="InvoLine5.InvoHeadID"
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
        dbText "Name" ="TransactionLines"
        dbInteger "ColumnWidth" ="1680"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sales"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transactions"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NoSales"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CancelledSales"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IDScanCount"
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
    Bottom =121
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =118
        Top =3
        Right =214
        Bottom =117
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =252
        Top =6
        Right =408
        Bottom =120
        Top =0
        Name ="Sysfile"
        Name =""
    End
End
