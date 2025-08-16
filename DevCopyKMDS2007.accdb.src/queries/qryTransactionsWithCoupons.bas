Operation =1
Option =0
Begin InputTables
    Name ="Sysfile"
    Name ="InvoLine"
End
Begin OutputColumns
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.CashSysID"
    Expression ="InvoLine.SetupDate"
    Alias ="SumOfPrice"
    Expression ="Sum(InvoLine.Price)"
End
Begin Joins
    LeftTable ="Sysfile"
    RightTable ="InvoLine"
    Expression ="Sysfile.CouponItemID = InvoLine.Code"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLine.SetupDate"
    Flag =0
End
Begin Groups
    Expression ="InvoLine.InvoHeadID"
    GroupLevel =0
    Expression ="InvoLine.CashSysID"
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
        dbText "Name" ="InvoLine.SetupDate"
        dbInteger "ColumnWidth" ="2925"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPrice"
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
    Bottom =212
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =172
        Top =6
        Right =387
        Bottom =224
        Top =0
        Name ="Sysfile"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =210
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
