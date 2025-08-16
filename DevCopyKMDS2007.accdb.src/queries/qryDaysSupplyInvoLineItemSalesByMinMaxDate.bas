Operation =1
Option =0
Where ="(((InvoLine.TransType)=\"S\" Or (InvoLine.TransType)=\"C\"))"
Begin InputTables
    Name ="InvoLine"
End
Begin OutputColumns
    Expression ="InvoLine.Code"
    Alias ="SumOfQuantity"
    Expression ="Sum(InvoLine.Quantity)"
    Alias ="SumOfPrice"
    Expression ="Sum(InvoLine.Price)"
    Alias ="MinOfSetupDate"
    Expression ="Min(InvoLine.SetupDate)"
    Alias ="MaxOfSetupDate"
    Expression ="Max(InvoLine.SetupDate)"
End
Begin OrderBy
    Expression ="Min(InvoLine.SetupDate)"
    Flag =0
End
Begin Groups
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
        dbText "Name" ="SumOfPrice"
        dbInteger "ColumnWidth" ="1845"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfSetupDate"
        dbInteger "ColumnWidth" ="2190"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MinOfSetupDate"
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
        Left =205
        Top =2
        Right =301
        Bottom =116
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
