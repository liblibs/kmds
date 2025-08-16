Operation =1
Option =0
Having ="(((InvoLine.TransType)=\"P\"))"
Begin InputTables
    Name ="InvoLine"
End
Begin OutputColumns
    Alias ="VendorNo"
    Expression ="InvoLine.SalesGroup"
    Alias ="PONumber"
    Expression ="InvoLine.SalesDept"
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.InputCode"
    Alias ="TotalQty"
    Expression ="Sum(InvoLine.Quantity)"
    Alias ="TotalPrice"
    Expression ="Sum(InvoLine.Price)"
    Alias ="PostDate"
    Expression ="InvoLine.SetupDate"
    Alias ="TotalCost"
    Expression ="Sum(InvoLine.Cost)"
    Expression ="InvoLine.CashSysID"
End
Begin OrderBy
    Expression ="InvoLine.SetupDate"
    Flag =1
End
Begin Groups
    Expression ="InvoLine.SalesGroup"
    GroupLevel =0
    Expression ="InvoLine.SalesDept"
    GroupLevel =0
    Expression ="InvoLine.InvoHeadID"
    GroupLevel =0
    Expression ="InvoLine.InputCode"
    GroupLevel =0
    Expression ="InvoLine.SetupDate"
    GroupLevel =0
    Expression ="InvoLine.CashSysID"
    GroupLevel =0
    Expression ="InvoLine.TransType"
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
        dbText "Name" ="PostDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalPrice"
        dbInteger "ColumnWidth" ="1200"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PONumber"
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
