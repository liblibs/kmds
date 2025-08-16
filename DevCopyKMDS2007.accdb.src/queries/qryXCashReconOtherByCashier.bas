Operation =1
Option =0
Where ="(((InvoLine.TransType)='T'))"
Having ="(((InvoLine.Code)=1) AND ((InvoLine.InputCode)<>\"Cash\" And (InvoLine.InputCode"
    ")<>\"Change\") AND ((Cashier.ID)=[parCashierNumber]))"
Begin InputTables
    Name ="InvoLine"
    Name ="InvoHead"
    Name ="Cashier"
    Name ="Tender"
End
Begin OutputColumns
    Expression ="InvoLine.Code"
    Alias ="Transact"
    Expression ="InvoLine.InputCode"
    Alias ="SumOfQuantity"
    Expression ="Sum(InvoLine.Quantity)"
    Alias ="SumOfPrice"
    Expression ="Sum(InvoLine.Price)"
    Expression ="Cashier.ID"
    Expression ="Cashier.Name"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="InvoHead"
    Expression ="InvoLine.InvoHeadID = InvoHead.ID"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="InvoHead"
    Expression ="InvoLine.CashSysID = InvoHead.CashSysID"
    Flag =1
    LeftTable ="InvoHead"
    RightTable ="Cashier"
    Expression ="InvoHead.CashierID = Cashier.ID"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="Tender"
    Expression ="InvoLine.Code = Tender.ID"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLine.Code"
    Flag =0
End
Begin Groups
    Expression ="InvoLine.Code"
    GroupLevel =0
    Expression ="InvoLine.InputCode"
    GroupLevel =0
    Expression ="Cashier.ID"
    GroupLevel =0
    Expression ="Cashier.Name"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Transact"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPrice"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =282
        Top =-4
        Right =378
        Bottom =110
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =143
        Top =13
        Right =239
        Bottom =127
        Top =0
        Name ="InvoHead"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =90
        Top =0
        Name ="Cashier"
        Name =""
    End
    Begin
        Left =433
        Top =-2
        Right =529
        Bottom =82
        Top =0
        Name ="Tender"
        Name =""
    End
End
