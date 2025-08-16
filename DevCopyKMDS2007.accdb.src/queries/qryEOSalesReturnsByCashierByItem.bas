Operation =1
Option =0
Where ="(((InvoLine5.TransType)=\"S\") And ((InvoLine5.Code)<>Sysfile.CouponItemID And ("
    "InvoLine5.Code)<>Sysfile.DiscountItemID And (InvoLine5.Code)<>Sysfile.LotteryPai"
    "dOutItemID))"
Having ="(((Sum(InvoLine5.Quantity))<0) AND ((InvoLine5.TransType)=\"S\"))"
Begin InputTables
    Name ="Sysfile"
    Name ="Cashier"
    Name ="InvoLine5"
    Name ="Item"
End
Begin OutputColumns
    Expression ="InvoLine5.CashierID"
    Alias ="CashierName"
    Expression ="Cashier.Name"
    Expression ="InvoLine5.TransType"
    Expression ="InvoLine5.SetupDate"
    Expression ="InvoLine5.CashSysID"
    Expression ="InvoLine5.InvoHeadID"
    Expression ="InvoLine5.Code"
    Expression ="Item.Description"
    Alias ="SumOfQuantity"
    Expression ="Sum(InvoLine5.Quantity)"
    Alias ="SumOfPrice"
    Expression ="Sum(InvoLine5.Price)"
End
Begin Joins
    LeftTable ="InvoLine5"
    RightTable ="Item"
    Expression ="InvoLine5.Code = Item.ID"
    Flag =1
    LeftTable ="Cashier"
    RightTable ="InvoLine5"
    Expression ="Cashier.ID = InvoLine5.CashierID"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLine5.CashierID"
    Flag =0
    Expression ="InvoLine5.SetupDate"
    Flag =0
End
Begin Groups
    Expression ="InvoLine5.CashierID"
    GroupLevel =0
    Expression ="Cashier.Name"
    GroupLevel =0
    Expression ="InvoLine5.TransType"
    GroupLevel =0
    Expression ="InvoLine5.SetupDate"
    GroupLevel =0
    Expression ="InvoLine5.CashSysID"
    GroupLevel =0
    Expression ="InvoLine5.InvoHeadID"
    GroupLevel =0
    Expression ="InvoLine5.Code"
    GroupLevel =0
    Expression ="Item.Description"
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
        dbText "Name" ="SumOfQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashierName"
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
    Bottom =504
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =535
        Top =41
        Right =732
        Bottom =395
        Top =0
        Name ="Sysfile"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="Cashier"
        Name =""
    End
    Begin
        Left =163
        Top =51
        Right =259
        Bottom =345
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =329
        Top =94
        Right =425
        Bottom =208
        Top =0
        Name ="Item"
        Name =""
    End
End
