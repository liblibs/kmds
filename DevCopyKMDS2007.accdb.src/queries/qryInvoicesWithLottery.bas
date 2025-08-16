Operation =1
Option =0
Where ="(((Item.SalesDepartmentID)=7 Or (Item.SalesDepartmentID)=11))"
Begin InputTables
    Name ="InvoLine"
    Name ="Item"
End
Begin OutputColumns
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.CashSysID"
    Expression ="InvoLine.SetupDate"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
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
        dbInteger "ColumnWidth" ="2310"
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
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
End
