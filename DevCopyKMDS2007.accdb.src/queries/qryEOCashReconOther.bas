Operation =1
Option =0
Having ="(((InvoLine5.Code)=1) AND ((InvoLine5.InputCode)<>\"Cash\" And (InvoLine5.InputC"
    "ode)<>\"Change\") AND ((InvoLine5.TransType)='T'))"
Begin InputTables
    Name ="InvoHead5"
    Name ="Cashier"
    Name ="InvoLine5"
    Name ="Tender"
End
Begin OutputColumns
    Expression ="InvoHead5.CashierID"
    Alias ="CashierName"
    Expression ="Cashier.Name"
    Alias ="TenderID"
    Expression ="InvoLine5.Code"
    Alias ="Transact"
    Expression ="InvoLine5.InputCode"
    Alias ="CR"
    Expression ="InvoLine5.CashSysID"
    Alias ="Invoice"
    Expression ="InvoLine5.InvoHeadID"
    Alias ="Time"
    Expression ="Format(Invohead5.SetupDate,\"Short Time\")"
    Alias ="TranCount"
    Expression ="Sum((1))"
    Alias ="Amount"
    Expression ="Sum(0-[Price])"
End
Begin Joins
    LeftTable ="InvoHead5"
    RightTable ="Cashier"
    Expression ="InvoHead5.CashierID = Cashier.ID"
    Flag =1
    LeftTable ="InvoLine5"
    RightTable ="Tender"
    Expression ="InvoLine5.Code = Tender.ID"
    Flag =1
    LeftTable ="InvoHead5"
    RightTable ="InvoLine5"
    Expression ="InvoHead5.ID = InvoLine5.InvoHeadID"
    Flag =1
    LeftTable ="InvoHead5"
    RightTable ="InvoLine5"
    Expression ="InvoHead5.CashSysID = InvoLine5.CashSysID"
    Flag =1
End
Begin OrderBy
    Expression ="InvoHead5.CashierID"
    Flag =0
    Expression ="InvoLine5.InputCode"
    Flag =0
    Expression ="Sum(0-[Price])"
    Flag =0
End
Begin Groups
    Expression ="InvoHead5.CashierID"
    GroupLevel =0
    Expression ="Cashier.Name"
    GroupLevel =0
    Expression ="InvoLine5.Code"
    GroupLevel =0
    Expression ="InvoLine5.InputCode"
    GroupLevel =0
    Expression ="InvoLine5.CashSysID"
    GroupLevel =0
    Expression ="InvoLine5.InvoHeadID"
    GroupLevel =0
    Expression ="Format(Invohead5.SetupDate,\"Short Time\")"
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
        dbText "Name" ="Time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TranCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Amount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashierName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TenderID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transact"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CR"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Invoice"
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
    Bottom =180
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="InvoHead5"
        Name =""
    End
    Begin
        Left =160
        Top =91
        Right =256
        Bottom =205
        Top =0
        Name ="Cashier"
        Name =""
    End
    Begin
        Left =302
        Top =28
        Right =398
        Bottom =142
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =530
        Top =-17
        Right =683
        Bottom =142
        Top =0
        Name ="Tender"
        Name =""
    End
End
