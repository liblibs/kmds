Operation =1
Option =0
Where ="(((InvoLine.Quantity)<0) AND ((qryXCancelledSales.InvoHeadID) Is Null) AND ((Inv"
    "oLine.TransType)=\"S\") AND ((SalesDepartment.Discounts)=False))"
Begin InputTables
    Name ="Cashier"
    Name ="InvoLine"
    Name ="qryXCancelledSales"
    Name ="Item"
    Name ="SalesDepartment"
End
Begin OutputColumns
    Expression ="InvoLine.CashierID"
    Alias ="CashierName"
    Expression ="Cashier.Name"
    Expression ="InvoLine.SetupDate"
    Alias ="RegisterNumber"
    Expression ="InvoLine.CashSysID"
    Alias ="InvoiceNumber"
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.Code"
    Expression ="Item.Description"
    Expression ="InvoLine.Quantity"
    Alias ="Amount"
    Expression ="InvoLine.Price"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="qryXCancelledSales"
    Expression ="InvoLine.InvoHeadID = qryXCancelledSales.InvoHeadID"
    Flag =2
    LeftTable ="InvoLine"
    RightTable ="qryXCancelledSales"
    Expression ="InvoLine.CashSysID = qryXCancelledSales.CashSysID"
    Flag =2
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SalesDepartment"
    Expression ="Item.SalesDepartmentID = SalesDepartment.ID"
    Flag =1
    LeftTable ="Cashier"
    RightTable ="InvoLine"
    Expression ="Cashier.ID = InvoLine.CashierID"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLine.CashierID"
    Flag =0
    Expression ="InvoLine.SetupDate"
    Flag =0
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
        dbText "Name" ="InvoLine.Quantity"
        dbInteger "ColumnWidth" ="945"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Code"
        dbInteger "ColumnWidth" ="645"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.SetupDate"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceNumber"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RegisterNumber"
        dbInteger "ColumnWidth" ="1680"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Amount"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashierName"
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
    Bottom =140
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =3
        Top =17
        Right =99
        Bottom =101
        Top =0
        Name ="Cashier"
        Name =""
    End
    Begin
        Left =255
        Top =5
        Right =375
        Bottom =149
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =411
        Top =-8
        Right =717
        Bottom =76
        Top =0
        Name ="qryXCancelledSales"
        Name =""
    End
    Begin
        Left =409
        Top =80
        Right =505
        Bottom =164
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =612
        Top =74
        Right =708
        Bottom =158
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
End
