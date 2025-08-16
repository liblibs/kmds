Operation =1
Option =0
Where ="(((InvoLine.TransType)<>\"T\"))"
Begin InputTables
    Name ="qryInvoLineByFirstPaymentType"
    Name ="InvoLine"
    Name ="Item"
    Name ="SalesDepartment"
End
Begin OutputColumns
    Expression ="Item.SalesDepartmentID"
    Expression ="SalesDepartment.Description"
    Alias ="SumOfPrice"
    Expression ="Sum(InvoLine.Price)"
    Alias ="Cash"
    Expression ="Sum(IIf([Tendercode]=1,[InvoLine].[Price],0))"
    Alias ="Checks"
    Expression ="Sum(IIf([Tendercode]=2,[InvoLine].[Price],0))"
    Alias ="CreditCard"
    Expression ="Sum(IIf([Tendercode]=3,[InvoLine].[Price],0))"
    Alias ="Other"
    Expression ="Sum(IIf([Tendercode]>3,[InvoLine].[Price],0))"
    Alias ="Count"
    Expression ="Sum(1)"
End
Begin Joins
    LeftTable ="qryInvoLineByFirstPaymentType"
    RightTable ="InvoLine"
    Expression ="qryInvoLineByFirstPaymentType.SetupDate=InvoLine.SetupDate"
    Flag =1
    LeftTable ="qryInvoLineByFirstPaymentType"
    RightTable ="InvoLine"
    Expression ="qryInvoLineByFirstPaymentType.InvoHeadID=InvoLine.InvoHeadID"
    Flag =1
    LeftTable ="qryInvoLineByFirstPaymentType"
    RightTable ="InvoLine"
    Expression ="qryInvoLineByFirstPaymentType.CashSysID=InvoLine.CashSysID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SalesDepartment"
    Expression ="Item.SalesDepartmentID=SalesDepartment.ID"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code=Item.ID"
    Flag =1
End
Begin OrderBy
    Expression ="Item.SalesDepartmentID"
    Flag =0
End
Begin Groups
    Expression ="Item.SalesDepartmentID"
    GroupLevel =0
    Expression ="SalesDepartment.Description"
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
        dbText "Name" ="Checks"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cash"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CreditCard"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Other"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Count"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1457
    Bottom =677
    Left =-1
    Top =-1
    Right =1425
    Bottom =277
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =13
        Top =52
        Right =157
        Bottom =302
        Top =0
        Name ="qryInvoLineByFirstPaymentType"
        Name =""
    End
    Begin
        Left =291
        Top =35
        Right =435
        Bottom =179
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =501
        Top =92
        Right =645
        Bottom =236
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =729
        Top =124
        Right =873
        Bottom =268
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
End
