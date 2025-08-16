Operation =1
Option =0
Where ="(((InvoLine.TransType)=\"D\") AND ((InvoLine.Quantity)=-1))"
Begin InputTables
    Name ="InvoLine"
    Name ="Club"
End
Begin OutputColumns
    Expression ="InvoLine.TransType"
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.CashSysID"
    Alias ="ClubID"
    Expression ="InvoLine.SalesDept"
    Alias ="InvoiceNumber"
    Expression ="InvoLine.SalesGroup"
    Alias ="Order"
    Expression ="InvoLine.SalesGroup"
    Expression ="InvoLine.SetupDate"
    Expression ="Club.Name"
    Expression ="InvoLine.Quantity"
    Expression ="InvoLine.Price"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Club"
    Expression ="InvoLine.SalesDept = Club.ID"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLine.InvoHeadID"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
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
        dbText "Name" ="InvoLine.InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.CashSysID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.SalesDept"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.CashierID"
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
        dbText "Name" ="InvoLine.upsize_ts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="COUNT"
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
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Drawer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.UpdateOnhand"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1003"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.Name"
        dbInteger "ColumnWidth" ="2145"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Order"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-171
    Top =56
    Right =946
    Bottom =567
    Left =-1
    Top =-1
    Right =1085
    Bottom =218
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =47
        Top =-13
        Right =275
        Bottom =247
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =324
        Top =12
        Right =468
        Bottom =156
        Top =0
        Name ="Club"
        Name =""
    End
End
