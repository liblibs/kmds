Operation =1
Option =0
Begin InputTables
    Name ="Sysfile"
    Name ="qryClubInvoiceHistoryByInvoice"
    Name ="InvoLine"
    Name ="Club"
    Name ="Item"
End
Begin OutputColumns
    Expression ="Club.NonLiquor"
    Expression ="qryClubInvoiceHistoryByInvoice.ClubID"
    Alias ="ClubName"
    Expression ="Club.Name"
    Alias ="YYYYMM"
    Expression ="Format(qryClubInvoiceHistoryByInvoice.SETUPDATE,\"yyyy-mm\")"
    Alias ="Qty"
    Expression ="Sum(IIf(InvoLine.TransType='C',[Quantity],0))"
    Alias ="Gross"
    Expression ="Sum(IIf(InvoLine.TransType='C',InvoLine.Price,0))"
    Alias ="Discount"
    Expression ="Sum(IIf(InvoLine.TransType='D',InvoLine.Price,0))"
    Alias ="GrossLiquor"
    Expression ="Sum(IIf(Item.SalesDepartmentID=1 And InvoLine.TransType='C',InvoLine.Price,0))"
    Alias ="LiquorDiscount"
    Expression ="Sum(IIf(InvoLine.TransType='D' And InvoLine.Code=Sysfile.ClubDiscntItemID,InvoLi"
        "ne.Price,0))"
    Alias ="GrossNonLiquor"
    Expression ="Sum(IIf(Item.SalesDepartmentID<>1 And InvoLine.TransType='C',InvoLine.Price,0))"
    Alias ="NonLiquorDiscount"
    Expression ="Sum(IIf(InvoLine.TransType='D' And InvoLine.Code=Sysfile.NonClubDiscntItemID,Inv"
        "oLine.Price,0))"
End
Begin Joins
    LeftTable ="qryClubInvoiceHistoryByInvoice"
    RightTable ="InvoLine"
    Expression ="qryClubInvoiceHistoryByInvoice.InvoHeadID = InvoLine.InvoHeadID"
    Flag =1
    LeftTable ="qryClubInvoiceHistoryByInvoice"
    RightTable ="InvoLine"
    Expression ="qryClubInvoiceHistoryByInvoice.CashSysID = InvoLine.CashSysID"
    Flag =1
    LeftTable ="qryClubInvoiceHistoryByInvoice"
    RightTable ="InvoLine"
    Expression ="qryClubInvoiceHistoryByInvoice.SetupDate = InvoLine.SetupDate"
    Flag =1
    LeftTable ="qryClubInvoiceHistoryByInvoice"
    RightTable ="Club"
    Expression ="qryClubInvoiceHistoryByInvoice.ClubID = Club.ID"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
End
Begin OrderBy
    Expression ="Club.Name"
    Flag =0
    Expression ="Format(qryClubInvoiceHistoryByInvoice.SETUPDATE,\"yyyy-mm\")"
    Flag =0
End
Begin Groups
    Expression ="Club.NonLiquor"
    GroupLevel =0
    Expression ="qryClubInvoiceHistoryByInvoice.ClubID"
    GroupLevel =0
    Expression ="Club.Name"
    GroupLevel =0
    Expression ="Format(qryClubInvoiceHistoryByInvoice.SETUPDATE,\"yyyy-mm\")"
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
        dbText "Name" ="Qty"
        dbInteger "ColumnWidth" ="915"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GrossNonLiquor"
        dbInteger "ColumnWidth" ="1560"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NonLiquorDiscount"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="YYYYMM"
        dbInteger "ColumnWidth" ="2400"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gross"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Discount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GrossLiquor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LiquorDiscount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.NonLiquor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubInvoiceHistoryByInvoice.ClubID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NonLiquor"
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
    Bottom =305
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =12
        Top =-5
        Right =150
        Bottom =289
        Top =0
        Name ="Sysfile"
        Name =""
    End
    Begin
        Left =283
        Top =12
        Right =517
        Bottom =156
        Top =0
        Name ="qryClubInvoiceHistoryByInvoice"
        Name =""
    End
    Begin
        Left =547
        Top =14
        Right =643
        Bottom =128
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =139
        Top =42
        Right =235
        Bottom =156
        Top =0
        Name ="Club"
        Name =""
    End
    Begin
        Left =695
        Top =6
        Right =936
        Bottom =195
        Top =0
        Name ="Item"
        Name =""
    End
End
