Operation =1
Option =0
Begin InputTables
    Name ="Sysfile"
    Name ="qry EOD060"
    Name ="InvoLine5"
    Name ="Club"
    Name ="Item"
    Name ="qry EOD060InvoiceNumber"
End
Begin OutputColumns
    Expression ="Club.NonLiquor"
    Expression ="[qry EOD060].ClubID"
    Alias ="ClubName"
    Expression ="Club.Name"
    Expression ="InvoLine5.InvoHeadID"
    Alias ="Qty"
    Expression ="Sum(IIf(InvoLine5.TransType='C',[Quantity],0))"
    Alias ="Gross"
    Expression ="Sum(IIf(InvoLine5.TransType='C',InvoLine5.Price,0))"
    Alias ="Discount"
    Expression ="Sum(IIf(InvoLine5.TransType='D',InvoLine5.Price,0))"
    Alias ="GrossLiquor"
    Expression ="Sum(IIf(Item.SalesDepartmentID=1 And InvoLine5.TransType='C',involine5.Price,0))"
    Alias ="LiquorDiscount"
    Expression ="Sum(IIf(InvoLine5.TransType='D' And Involine5.Code=Sysfile.ClubDiscntItemID,Invo"
        "Line5.Price,0))"
    Alias ="GrossNonLiquor"
    Expression ="Sum(IIf(Item.SalesDepartmentID<>1 And InvoLine5.TransType='C',involine5.Price,0)"
        ")"
    Alias ="NonLiquorDiscount"
    Expression ="Sum(IIf(InvoLine5.TransType='D' And Involine5.Code=Sysfile.NonClubDiscntItemID,I"
        "nvoLine5.Price,0))"
    Expression ="[qry EOD060InvoiceNumber].InvoiceNumber"
End
Begin Joins
    LeftTable ="qry EOD060"
    RightTable ="InvoLine5"
    Expression ="[qry EOD060].InvoHeadID = InvoLine5.InvoHeadID"
    Flag =1
    LeftTable ="qry EOD060"
    RightTable ="Club"
    Expression ="[qry EOD060].ClubID = Club.ID"
    Flag =1
    LeftTable ="InvoLine5"
    RightTable ="Item"
    Expression ="InvoLine5.Code = Item.ID"
    Flag =1
    LeftTable ="InvoLine5"
    RightTable ="qry EOD060InvoiceNumber"
    Expression ="InvoLine5.CashSysID = [qry EOD060InvoiceNumber].CashSysID"
    Flag =1
    LeftTable ="InvoLine5"
    RightTable ="qry EOD060InvoiceNumber"
    Expression ="InvoLine5.InvoHeadID = [qry EOD060InvoiceNumber].InvoHeadID"
    Flag =1
End
Begin Groups
    Expression ="Club.NonLiquor"
    GroupLevel =0
    Expression ="[qry EOD060].ClubID"
    GroupLevel =0
    Expression ="Club.Name"
    GroupLevel =0
    Expression ="InvoLine5.InvoHeadID"
    GroupLevel =0
    Expression ="[qry EOD060InvoiceNumber].InvoiceNumber"
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
        dbText "Name" ="InvoLine5.InvoHeadID"
        dbInteger "ColumnWidth" ="1080"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Discount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gross"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Qty"
        dbInteger "ColumnWidth" ="915"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="GrossNonLiquor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NonLiquorDiscount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubName"
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
    Bottom =321
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =12
        Top =-5
        Right =127
        Bottom =289
        Top =0
        Name ="Sysfile"
        Name =""
    End
    Begin
        Left =274
        Top =7
        Right =390
        Bottom =106
        Top =0
        Name ="qry EOD060"
        Name =""
    End
    Begin
        Left =489
        Top =13
        Right =735
        Bottom =202
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =128
        Top =15
        Right =224
        Bottom =129
        Top =0
        Name ="Club"
        Name =""
    End
    Begin
        Left =752
        Top =15
        Right =993
        Bottom =204
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =308
        Top =192
        Right =404
        Bottom =306
        Top =0
        Name ="qry EOD060InvoiceNumber"
        Name =""
    End
End
