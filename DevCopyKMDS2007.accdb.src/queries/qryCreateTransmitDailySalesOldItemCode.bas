Operation =1
Option =0
Having ="(((Format([Code],\"000000\")) Between 1000 And 99999) AND ((Format(Sum([quantity"
    "]),\"00000;-0000\"))<>0) AND ((InvoLine5.TransType)=\"S\" Or (InvoLine5.TransTyp"
    "e)=\"C\"))"
Begin InputTables
    Name ="InvoLine5"
    Name ="qryCreateTransmitDailySalesInvoiceNo"
    Name ="qryCreateTransmitDailySalesTenderCode"
End
Begin OutputColumns
    Expression ="InvoLine5.CashSysID"
    Alias ="TransactionNumber"
    Expression ="Format(First(((InvoLine5.InvoHeadID*100)+[ID])),\"000000000\")"
    Alias ="CreateDate"
    Expression ="Format(Now(),'yymmdd')"
    Alias ="CreateTime"
    Expression ="Format(Now(),'hhnn')"
    Alias ="OLCCItemCode"
    Expression ="Format([Code],\"000000\")"
    Alias ="OLCCQuantity"
    Expression ="Format(Sum([quantity]),\"00000;-0000\")"
    Alias ="RegisterDate"
    Expression ="Format(InvoLine5.SetupDate,'yymmdd')"
    Alias ="RegisterTime"
    Expression ="Format(InvoLine5.SetupDate,'hhnnss')"
    Alias ="UPCCode"
    Expression ="Format(IIf(IsNumeric([InputCode]),[InputCode],[Code]),\"!@@@@@@@@@@@@@@\")"
    Alias ="SellingPrice"
    Expression ="Format(((InvoLine5.Price/[quantity])*100),\"0000000\")"
    Alias ="LIN"
    Expression ="qryCreateTransmitDailySalesInvoiceNo.LicenseeInvoiceNumber"
    Alias ="PN"
    Expression ="qryCreateTransmitDailySalesInvoiceNo.PremisesNumber"
    Expression ="qryCreateTransmitDailySalesInvoiceNo.Discount"
    Expression ="qryCreateTransmitDailySalesTenderCode.TenderCode"
End
Begin Joins
    LeftTable ="InvoLine5"
    RightTable ="qryCreateTransmitDailySalesInvoiceNo"
    Expression ="InvoLine5.SetupDate = qryCreateTransmitDailySalesInvoiceNo.SetupDate"
    Flag =2
    LeftTable ="InvoLine5"
    RightTable ="qryCreateTransmitDailySalesInvoiceNo"
    Expression ="InvoLine5.CashSysID = qryCreateTransmitDailySalesInvoiceNo.CashSysID"
    Flag =2
    LeftTable ="InvoLine5"
    RightTable ="qryCreateTransmitDailySalesInvoiceNo"
    Expression ="InvoLine5.InvoHeadID = qryCreateTransmitDailySalesInvoiceNo.InvoHeadID"
    Flag =2
    LeftTable ="InvoLine5"
    RightTable ="qryCreateTransmitDailySalesTenderCode"
    Expression ="InvoLine5.InvoHeadID = qryCreateTransmitDailySalesTenderCode.InvoHeadID"
    Flag =1
    LeftTable ="InvoLine5"
    RightTable ="qryCreateTransmitDailySalesTenderCode"
    Expression ="InvoLine5.CashSysID = qryCreateTransmitDailySalesTenderCode.CashSysID"
    Flag =1
    LeftTable ="InvoLine5"
    RightTable ="qryCreateTransmitDailySalesTenderCode"
    Expression ="InvoLine5.SetupDate = qryCreateTransmitDailySalesTenderCode.SetupDate"
    Flag =1
End
Begin OrderBy
    Expression ="Format(First(((InvoLine5.InvoHeadID*100)+[ID])),\"000000000\")"
    Flag =0
End
Begin Groups
    Expression ="InvoLine5.CashSysID"
    GroupLevel =0
    Expression ="Format(Now(),'yymmdd')"
    GroupLevel =0
    Expression ="Format(Now(),'hhnn')"
    GroupLevel =0
    Expression ="Format([Code],\"000000\")"
    GroupLevel =0
    Expression ="Format(InvoLine5.SetupDate,'yymmdd')"
    GroupLevel =0
    Expression ="Format(InvoLine5.SetupDate,'hhnnss')"
    GroupLevel =0
    Expression ="Format(IIf(IsNumeric([InputCode]),[InputCode],[Code]),\"!@@@@@@@@@@@@@@\")"
    GroupLevel =0
    Expression ="Format(((InvoLine5.Price/[quantity])*100),\"0000000\")"
    GroupLevel =0
    Expression ="qryCreateTransmitDailySalesInvoiceNo.LicenseeInvoiceNumber"
    GroupLevel =0
    Expression ="qryCreateTransmitDailySalesInvoiceNo.PremisesNumber"
    GroupLevel =0
    Expression ="qryCreateTransmitDailySalesInvoiceNo.Discount"
    GroupLevel =0
    Expression ="qryCreateTransmitDailySalesTenderCode.TenderCode"
    GroupLevel =0
    Expression ="InvoLine5.TransType"
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
        dbText "Name" ="SellingPrice"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransactionNumber"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCCode"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CreateDate"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CreateTime"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCItemCode"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCQuantity"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RegisterDate"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RegisterTime"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LIN"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PN"
        dbInteger "ColumnOrder" ="9"
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
    Right =1193
    Bottom =214
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =19
        Top =3
        Right =115
        Bottom =228
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =273
        Top =9
        Right =442
        Bottom =123
        Top =0
        Name ="qryCreateTransmitDailySalesInvoiceNo"
        Name =""
    End
    Begin
        Left =263
        Top =126
        Right =359
        Bottom =240
        Top =0
        Name ="qryCreateTransmitDailySalesTenderCode"
        Name =""
    End
End
