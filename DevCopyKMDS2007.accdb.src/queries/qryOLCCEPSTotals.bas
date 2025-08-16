Operation =1
Option =0
Begin InputTables
    Name ="OLCCEPS"
End
Begin OutputColumns
    Expression ="OLCCEPS.OrderDate"
    Expression ="OLCCEPS.OrderNumber"
    Alias ="SumOfShipped"
    Expression ="Sum(OLCCEPS.Shipped)"
    Alias ="Extended"
    Expression ="Sum([Shipped]*[Price])"
End
Begin OrderBy
    Expression ="OLCCEPS.OrderDate"
    Flag =1
    Expression ="OLCCEPS.OrderNumber"
    Flag =1
End
Begin Groups
    Expression ="OLCCEPS.OrderDate"
    GroupLevel =0
    Expression ="OLCCEPS.OrderNumber"
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
        dbText "Name" ="OLCCEPS.OrderDate"
        dbInteger "ColumnWidth" ="1890"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCEPS.OrderNumber"
        dbInteger "ColumnWidth" ="1845"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Extended"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfShipped"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =69
    Top =62
    Right =1591
    Bottom =853
    Left =-1
    Top =-1
    Right =1490
    Bottom =491
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =372
        Bottom =301
        Top =0
        Name ="OLCCEPS"
        Name =""
    End
End
