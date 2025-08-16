Operation =1
Option =0
Where ="(((Format([InvoLine].[SetupDate],\"yyyymm\"))>=Format((Now()-365),\"yyyymm\") An"
    "d (Format([InvoLine].[SetupDate],\"yyyymm\"))<Format((Now()),\"yyyymm\")))"
Begin InputTables
    Name ="InvoLine"
End
Begin OutputColumns
    Expression ="InvoLine.Code"
    Expression ="InvoLine.TransType"
    Alias ="SumOfQuantity"
    Expression ="Sum(InvoLine.Quantity)"
    Alias ="SumOfPrice"
    Expression ="Sum(InvoLine.Price)"
    Alias ="YYYY"
    Expression ="Format(InvoLine.SetupDate,\"yyyy\")"
    Alias ="MM"
    Expression ="Format(InvoLine.SetupDate,\"mm\")"
End
Begin OrderBy
    Expression ="InvoLine.Code"
    Flag =0
    Expression ="Format(InvoLine.SetupDate,\"yyyy\")"
    Flag =0
    Expression ="Format(InvoLine.SetupDate,\"mm\")"
    Flag =0
End
Begin Groups
    Expression ="InvoLine.Code"
    GroupLevel =0
    Expression ="InvoLine.TransType"
    GroupLevel =0
    Expression ="Format(InvoLine.SetupDate,\"yyyy\")"
    GroupLevel =0
    Expression ="Format(InvoLine.SetupDate,\"mm\")"
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
        dbText "Name" ="SumOfQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="YYYY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.TransType"
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
    Bottom =194
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =138
        Top =34
        Right =268
        Bottom =223
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
