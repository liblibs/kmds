Operation =1
Option =0
Where ="(((InvoLine.Price)<>0) AND ((InvoLine.TransType)=\"T\"))"
Having ="(((InvoLine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom And F"
    "orms!PrintInventoryReport!txtDateRangeTo) And ((First(InvoLine.Code))<100))"
Begin InputTables
    Name ="InvoLine"
End
Begin OutputColumns
    Expression ="InvoLine.CashSysID"
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.SetupDate"
    Alias ="TenderCode"
    Expression ="First(InvoLine.Code)"
End
Begin Groups
    Expression ="InvoLine.CashSysID"
    GroupLevel =0
    Expression ="InvoLine.InvoHeadID"
    GroupLevel =0
    Expression ="InvoLine.SetupDate"
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
        dbText "Name" ="TenderCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.CashSysID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.SetupDate"
        dbInteger "ColumnWidth" ="2310"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =112
    Top =940
    Right =1198
    Bottom =1897
    Left =-1
    Top =-1
    Right =1054
    Bottom =273
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =-2
        Top =9
        Right =234
        Bottom =254
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
