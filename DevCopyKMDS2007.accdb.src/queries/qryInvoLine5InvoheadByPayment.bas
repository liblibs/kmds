Operation =1
Option =0
Where ="(((InvoLine5.TransType)<>\"T\"))"
Begin InputTables
    Name ="qryInvoLine5InvoheadByFirstPaymentType"
    Name ="InvoLine5"
    Name ="Item"
    Name ="SalesDepartment"
End
Begin OutputColumns
    Expression ="Item.SalesDepartmentID"
    Expression ="SalesDepartment.Description"
    Alias ="SumOfPrice"
    Expression ="Sum(InvoLine5.Price)"
    Alias ="Cash"
    Expression ="Sum(IIf([Tendercode]=1,[InvoLine5].[Price],0))"
    Alias ="Checks"
    Expression ="Sum(IIf([Tendercode]=2,[InvoLine5].[Price],0))"
    Alias ="CreditCard"
    Expression ="Sum(IIf([Tendercode]=3,[InvoLine5].[Price],0))"
    Alias ="Other"
    Expression ="Sum(IIf([Tendercode]>3,[InvoLine5].[Price],0))"
    Alias ="Count"
    Expression ="Sum(1)"
End
Begin Joins
    LeftTable ="qryInvoLine5InvoheadByFirstPaymentType"
    RightTable ="InvoLine5"
    Expression ="qryInvoLine5InvoheadByFirstPaymentType.InvoHeadID = InvoLine5.InvoHeadID"
    Flag =1
    LeftTable ="qryInvoLine5InvoheadByFirstPaymentType"
    RightTable ="InvoLine5"
    Expression ="qryInvoLine5InvoheadByFirstPaymentType.CashSysID = InvoLine5.CashSysID"
    Flag =1
    LeftTable ="qryInvoLine5InvoheadByFirstPaymentType"
    RightTable ="InvoLine5"
    Expression ="qryInvoLine5InvoheadByFirstPaymentType.SetupDate = InvoLine5.SetupDate"
    Flag =1
    LeftTable ="InvoLine5"
    RightTable ="Item"
    Expression ="InvoLine5.Code = Item.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SalesDepartment"
    Expression ="Item.SalesDepartmentID = SalesDepartment.ID"
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
    Right =1225
    Bottom =677
    Left =-1
    Top =-1
    Right =1209
    Bottom =327
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =312
        Top =0
        Name ="qryInvoLine5InvoheadByFirstPaymentType"
        Name =""
    End
    Begin
        Left =269
        Top =12
        Right =461
        Bottom =368
        Top =0
        Name ="InvoLine5"
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
