Operation =1
Option =0
Where ="(((InvoLine.TransType)=\"T\") AND ((InvoLine.SetupDate) Between #1/1/2013# And #"
    "12/31/2013#))"
Begin InputTables
    Name ="InvoLine"
End
Begin OutputColumns
    Alias ="BusinessDATE"
    Expression ="Format([SetupDate],\"yyyymmdd\")"
    Expression ="InvoLine.Code"
    Alias ="COUNT"
    Expression ="Sum(1)"
    Alias ="SumOfPrice"
    Expression ="Sum(InvoLine.Price)"
End
Begin OrderBy
    Expression ="Format([SetupDate],\"yyyymmdd\")"
    Flag =0
    Expression ="InvoLine.Code"
    Flag =0
End
Begin Groups
    Expression ="Format([SetupDate],\"yyyymmdd\")"
    GroupLevel =0
    Expression ="InvoLine.Code"
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
        dbText "Name" ="InvoLine.TransType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="COUNT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DATE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BusinessDATE"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =17
    Top =-1
    Right =1441
    Bottom =836
    Left =-1
    Top =-1
    Right =1392
    Bottom =554
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =441
        Bottom =440
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
