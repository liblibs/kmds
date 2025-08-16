Operation =3
Name ="InvoLine"
Option =0
Where ="(((GaribaldiSalesHist.Oct)<>0))"
Begin InputTables
    Name ="GaribaldiSalesHist"
End
Begin OutputColumns
    Alias ="InvoHeadID"
    Name ="InvoHeadID"
    Expression ="99"
    Name ="ID"
    Expression ="GaribaldiSalesHist.ID"
    Alias ="CashSysID"
    Name ="CashSysID"
    Expression ="99"
    Alias ="Code"
    Name ="Code"
    Expression ="GaribaldiSalesHist.itemid"
    Alias ="Quantity"
    Name ="Quantity"
    Expression ="GaribaldiSalesHist.Oct"
    Alias ="InputCode"
    Name ="InputCode"
    Expression ="\"Conversion\""
    Alias ="TransType"
    Name ="TransType"
    Expression ="\"S\""
    Alias ="SetupDate"
    Name ="SetupDate"
    Expression ="#10/2/2014#"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="InvoLine.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GaribaldiSalesHist.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.CashSysID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashSysID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.InputCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InputCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.TransType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SetupDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-305
    Top =159
    Right =1122
    Bottom =907
    Left =-1
    Top =-1
    Right =1395
    Bottom =448
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =330
        Top =0
        Name ="GaribaldiSalesHist"
        Name =""
    End
End
