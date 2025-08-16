Operation =1
Option =0
Where ="(((InvoLine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom And F"
    "orms!PrintInventoryReport!txtDateRangeTo) And ((InvoLine.Code) Between Forms!Pri"
    "ntInventoryReport!cboItemIDBegin And Forms!PrintInventoryReport!cboItemIDEnd) An"
    "d ((InvoLine.TransType)=\"S\"))"
Begin InputTables
    Name ="InvoLine"
End
Begin OutputColumns
    Alias ="DateOfSale"
    Expression ="InvoLine.SetupDate"
    Expression ="InvoLine.CashSysID"
    Expression ="InvoLine.InvoHeadID"
End
Begin OrderBy
    Expression ="InvoLine.SetupDate"
    Flag =0
End
Begin Groups
    Expression ="InvoLine.SetupDate"
    GroupLevel =0
    Expression ="InvoLine.CashSysID"
    GroupLevel =0
    Expression ="InvoLine.InvoHeadID"
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
        dbText "Name" ="DateOfSale"
        dbInteger "ColumnWidth" ="2280"
        dbBoolean "ColumnHidden" ="0"
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
