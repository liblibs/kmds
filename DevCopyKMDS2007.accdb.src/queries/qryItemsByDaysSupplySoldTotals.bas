Operation =1
Option =0
Where ="(((InvoLine.TransType)=\"S\" Or (InvoLine.TransType)=\"C\") And ((InvoLine.Setup"
    "Date) Between Forms!PrintInventoryReport!txtDateRangeFrom And Forms!PrintInvento"
    "ryReport!txtDateRangeTo))"
Begin InputTables
    Name ="InvoLine"
End
Begin OutputColumns
    Expression ="InvoLine.Code"
    Alias ="SumOfQuantity"
    Expression ="Sum(InvoLine.Quantity)"
    Alias ="SumOfPrice"
    Expression ="Sum(InvoLine.Price)"
    Alias ="SumOfCost"
    Expression ="Sum(InvoLine.Cost)"
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
        dbText "Name" ="SumOfQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfCost"
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
    Bottom =211
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =55
        Top =2
        Right =243
        Bottom =206
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
