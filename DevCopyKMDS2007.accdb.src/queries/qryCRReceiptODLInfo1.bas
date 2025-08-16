Operation =1
Option =0
Where ="(((InvoHead.ID)=[parInvoiceNo]) AND ((InvoHead.CashSysID)=[parRegisterNumber]) A"
    "ND ((InvoLine.TransType)=\"X\"))"
Begin InputTables
    Name ="Cashier"
    Name ="InvoHead"
    Name ="InvoLine"
    Name ="Tender"
End
Begin OutputColumns
    Alias ="InvoiceNumber"
    Expression ="InvoHead.ID"
    Expression ="InvoHead.CashSysID"
    Alias ="CashierName"
    Expression ="Cashier.Name"
    Expression ="InvoHead.SetupDate"
    Expression ="InvoLine.TransType"
    Expression ="InvoLine.Code"
    Alias ="ODL"
    Expression ="InvoLine.SalesDept"
    Alias ="DOB"
    Expression ="InvoLine.InputCode"
End
Begin Joins
    LeftTable ="InvoHead"
    RightTable ="InvoLine"
    Expression ="InvoHead.ID = InvoLine.InvoHeadID"
    Flag =1
    LeftTable ="InvoHead"
    RightTable ="InvoLine"
    Expression ="InvoHead.CashSysID = InvoLine.CashSysID"
    Flag =1
    LeftTable ="Cashier"
    RightTable ="InvoHead"
    Expression ="Cashier.ID = InvoHead.CashierID"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="Tender"
    Expression ="InvoLine.Code = Tender.ID"
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
        dbText "Name" ="InvoiceNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashierName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ODL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DOB"
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
    Bottom =201
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =17
        Top =31
        Right =113
        Bottom =115
        Top =0
        Name ="Cashier"
        Name =""
    End
    Begin
        Left =132
        Top =14
        Right =261
        Bottom =203
        Top =0
        Name ="InvoHead"
        Name =""
    End
    Begin
        Left =291
        Top =3
        Right =402
        Bottom =207
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =442
        Top =0
        Right =538
        Bottom =84
        Top =0
        Name ="Tender"
        Name =""
    End
End
