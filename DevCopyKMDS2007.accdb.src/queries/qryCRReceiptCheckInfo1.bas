Operation =1
Option =0
Where ="(((InvoLine.InvoHeadID)=[parInvoiceNo]) AND ((InvoLine.CashSysID)=[parRegisterNu"
    "mber]) AND ((Tender.GetCheckNumber)<>0)) OR (((InvoLine.InvoHeadID)=[parInvoiceN"
    "o]) AND ((InvoLine.CashSysID)=[parRegisterNumber]) AND ((Tender.GetCheckCashing)"
    "<>0))"
Begin InputTables
    Name ="InvoLine"
    Name ="Tender"
    Name ="CheckCashing"
End
Begin OutputColumns
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.CashSysID"
    Expression ="Tender.GetCheckNumber"
    Alias ="CheckNumber"
    Expression ="InvoLine.SalesGroup"
    Alias ="CheckCashingID"
    Expression ="InvoLine.SalesDept"
    Alias ="Customer"
    Expression ="[FirstName] & \" \" & [LastName]"
    Alias ="CheckAmount"
    Expression ="0-[Price]"
    Expression ="Tender.GetCheckCashing"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Tender"
    Expression ="InvoLine.Code = Tender.ID"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="CheckCashing"
    Expression ="InvoLine.SalesDept = CheckCashing.ID"
    Flag =2
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
        dbText "Name" ="Customer"
        dbInteger "ColumnWidth" ="2700"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CheckNumber"
        dbInteger "ColumnWidth" ="1830"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CheckCashingID"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CheckAmount"
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
    Bottom =205
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =259
        Top =30
        Right =355
        Bottom =144
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =84
        Top =89
        Right =180
        Bottom =203
        Top =0
        Name ="Tender"
        Name =""
    End
    Begin
        Left =454
        Top =4
        Right =550
        Bottom =118
        Top =0
        Name ="CheckCashing"
        Name =""
    End
End
