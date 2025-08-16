Operation =1
Option =0
Having ="(((InvoLine.TransType)=\"P\") AND ((Format([SetupDate],\"yyyy-mm-dd\"))>Format(N"
    "ow()-30,\"yyyy-mm-dd\")))"
Begin InputTables
    Name ="InvoLine"
End
Begin OutputColumns
    Alias ="InvoiceHeadID"
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.TransType"
    Expression ="InvoLine.InputCode"
    Alias ="SumOfQuantity"
    Expression ="Sum(InvoLine.Quantity)"
    Alias ="SumOfPrice"
    Expression ="Sum(InvoLine.Price)"
    Alias ="Date"
    Expression ="Format([SetupDate],\"yyyy-mm-dd\")"
End
Begin OrderBy
    Expression ="Format([SetupDate],\"yyyy-mm-dd\")"
    Flag =1
End
Begin Groups
    Expression ="InvoLine.InvoHeadID"
    GroupLevel =0
    Expression ="InvoLine.TransType"
    GroupLevel =0
    Expression ="InvoLine.InputCode"
    GroupLevel =0
    Expression ="Format([SetupDate],\"yyyy-mm-dd\")"
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
        dbText "Name" ="SumOfPrice"
        dbInteger "ColumnWidth" ="1200"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
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
