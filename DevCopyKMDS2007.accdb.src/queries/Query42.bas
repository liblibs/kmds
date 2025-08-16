Operation =1
Option =0
Where ="(((InvoLine.TransType)=\"S\"))"
Begin InputTables
    Name ="205hEAD"
    Name ="InvoLine"
End
Begin OutputColumns
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.ID"
    Expression ="InvoLine.CashSysID"
    Expression ="InvoLine.CashierID"
    Expression ="InvoLine.ShiftID"
    Expression ="InvoLine.Code"
    Expression ="InvoLine.SalesDept"
    Expression ="InvoLine.Quantity"
    Expression ="InvoLine.Price"
    Expression ="InvoLine.InputCode"
    Expression ="InvoLine.SalesGroup"
    Expression ="InvoLine.Drawer"
    Expression ="InvoLine.TransType"
    Expression ="InvoLine.UpdateOnhand"
    Expression ="InvoLine.Cost"
    Expression ="InvoLine.SetupDate"
End
Begin Joins
    LeftTable ="205hEAD"
    RightTable ="InvoLine"
    Expression ="[205hEAD].InvoHeadID = InvoLine.InvoHeadID"
    Flag =1
    LeftTable ="205hEAD"
    RightTable ="InvoLine"
    Expression ="[205hEAD].CashSysID = InvoLine.CashSysID"
    Flag =1
    LeftTable ="205hEAD"
    RightTable ="InvoLine"
    Expression ="[205hEAD].SetupDate = InvoLine.SetupDate"
    Flag =1
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
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.ID"
        dbInteger "ColumnWidth" ="0"
        dbBoolean "ColumnHidden" ="-1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.CashSysID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.CashierID"
        dbInteger "ColumnWidth" ="1005"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.ShiftID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.SalesDept"
        dbInteger "ColumnWidth" ="0"
        dbBoolean "ColumnHidden" ="-1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.InputCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.SalesGroup"
        dbInteger "ColumnWidth" ="0"
        dbBoolean "ColumnHidden" ="-1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Drawer"
        dbInteger "ColumnWidth" ="0"
        dbBoolean "ColumnHidden" ="-1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.TransType"
        dbInteger "ColumnWidth" ="0"
        dbBoolean "ColumnHidden" ="-1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.UpdateOnhand"
        dbInteger "ColumnWidth" ="0"
        dbBoolean "ColumnHidden" ="-1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Cost"
        dbInteger "ColumnWidth" ="0"
        dbBoolean "ColumnHidden" ="-1"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =2
    Left =-4
    Top =-23
    Right =997
    Bottom =565
    Left =-1
    Top =-1
    Right =952
    Bottom =272
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="205hEAD"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =273
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
