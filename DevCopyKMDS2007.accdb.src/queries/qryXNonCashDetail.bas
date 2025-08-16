Operation =1
Option =0
Where ="(((Tender.ID)<>1) AND ((0-[Price])<>0) AND ((Left([InputCode],4))<>\"Drop\" And "
    "(Left([InputCode],4))<>\"Chan\") AND ((Tender.NonCashDetailReport)=Yes) AND ((In"
    "voLine.TransType)=\"T\"))"
Begin InputTables
    Name ="InvoLine"
    Name ="Tender"
    Name ="CheckCashing"
    Name ="Cashier"
End
Begin OutputColumns
    Expression ="InvoLine.CashierID"
    Alias ="CashierName"
    Expression ="Cashier.Name"
    Alias ="TenderID"
    Expression ="Tender.ID"
    Alias ="MOP"
    Expression ="InvoLine.InputCode"
    Alias ="InvoiceNumber"
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.SetupDate"
    Alias ="Amount"
    Expression ="0-[Price]"
    Expression ="Tender.GetCheckCashing"
    Alias ="CheckCashing"
    Expression ="InvoLine.SalesDept"
    Expression ="Tender.GetCheckNumber"
    Alias ="CheckNumber"
    Expression ="InvoLine.SalesGroup"
    Alias ="CheckCashingName"
    Expression ="[FirstName] & \" \" & [LastName]"
    Alias ="InputType"
    Expression ="Left([InputCode],4)"
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
    LeftTable ="InvoLine"
    RightTable ="Cashier"
    Expression ="InvoLine.CashierID = Cashier.ID"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLine.CashierID"
    Flag =0
    Expression ="Tender.ID"
    Flag =0
    Expression ="InvoLine.InputCode"
    Flag =0
    Expression ="0-[Price]"
    Flag =0
    Expression ="Tender.NonCashDetailReport"
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
        dbText "Name" ="InvoLine.SetupDate"
        dbInteger "ColumnWidth" ="2175"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CheckCashing"
        dbInteger "ColumnWidth" ="1500"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CheckCashingName"
        dbInteger "ColumnWidth" ="2655"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TenderID"
        dbInteger "ColumnWidth" ="1020"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashierName"
        dbInteger "ColumnWidth" ="1425"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MOP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Amount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CheckNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InputType"
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
    Bottom =297
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =138
        Top =2
        Right =233
        Bottom =251
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =11
        Top =3
        Right =107
        Bottom =87
        Top =0
        Name ="Tender"
        Name =""
    End
    Begin
        Left =423
        Top =126
        Right =615
        Bottom =270
        Top =0
        Name ="CheckCashing"
        Name =""
    End
    Begin
        Left =421
        Top =2
        Right =517
        Bottom =86
        Top =0
        Name ="Cashier"
        Name =""
    End
End
