Operation =1
Option =0
Having ="(((Left([InputCode],4))=\"Cash\" Or (Left([InputCode],4))=\"Drop\") AND ((InvoLi"
    "ne5.InputCode)<>\"Cash\" And (InvoLine5.InputCode)<>\"Change\") AND ((InvoLine5."
    "TransType)='T'))"
Begin InputTables
    Name ="Cashier"
    Name ="InvoLine5"
    Name ="Tender"
End
Begin OutputColumns
    Expression ="InvoLine5.CashierID"
    Alias ="CashierName"
    Expression ="Cashier.Name"
    Alias ="TenderID"
    Expression ="InvoLine5.Code"
    Expression ="Tender.Description"
    Alias ="CashORLoan"
    Expression ="Left([InputCode],4)"
    Alias ="Transact"
    Expression ="InvoLine5.InputCode"
    Alias ="CR"
    Expression ="InvoLine5.CashSysID"
    Alias ="Invoice"
    Expression ="InvoLine5.InvoHeadID"
    Alias ="Time"
    Expression ="Format(InvoLine5.SetupDate,\"Short Time\")"
    Alias ="TranCount"
    Expression ="Sum((1))"
    Alias ="Amount"
    Expression ="Sum(InvoLine5.Price)"
    Alias ="Loan"
    Expression ="Sum(IIf(Left([InputCode],4)=\"Cash\",0-[Price],0))"
    Alias ="Drop"
    Expression ="Sum(IIf(Left([InputCode],4)=\"Drop\",[Price],0))"
End
Begin Joins
    LeftTable ="InvoLine5"
    RightTable ="Tender"
    Expression ="InvoLine5.Code = Tender.ID"
    Flag =1
    LeftTable ="Cashier"
    RightTable ="InvoLine5"
    Expression ="Cashier.ID = InvoLine5.CashierID"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLine5.CashierID"
    Flag =0
    Expression ="InvoLine5.Code"
    Flag =0
    Expression ="Left([InputCode],4)"
    Flag =0
    Expression ="InvoLine5.InvoHeadID"
    Flag =0
    Expression ="Format(InvoLine5.SetupDate,\"Short Time\")"
    Flag =0
    Expression ="Sum(InvoLine5.Price)"
    Flag =0
End
Begin Groups
    Expression ="InvoLine5.CashierID"
    GroupLevel =0
    Expression ="Cashier.Name"
    GroupLevel =0
    Expression ="InvoLine5.Code"
    GroupLevel =0
    Expression ="Tender.Description"
    GroupLevel =0
    Expression ="Left([InputCode],4)"
    GroupLevel =0
    Expression ="InvoLine5.InputCode"
    GroupLevel =0
    Expression ="InvoLine5.CashSysID"
    GroupLevel =0
    Expression ="InvoLine5.InvoHeadID"
    GroupLevel =0
    Expression ="Format(InvoLine5.SetupDate,\"Short Time\")"
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
        dbText "Name" ="CashORLoan"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Loan"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Drop"
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
        Left =525
        Top =4
        Right =678
        Bottom =163
        Top =0
        Name ="Tender"
        Name =""
    End
End
