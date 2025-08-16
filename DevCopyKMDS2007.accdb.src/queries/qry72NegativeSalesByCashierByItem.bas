Operation =1
Option =0
Where ="(((Involine.TransType)=\"S\") And ((Involine.Code)<>Sysfile.CouponItemID And (In"
    "voline.Code)<>Sysfile.DiscountItemID And (Involine.Code)<>Sysfile.LotteryPaidOut"
    "ItemID))"
Having ="(((Involine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom And F"
    "orms!PrintInventoryReport!txtDateRangeTo) And ((Sum(Involine.Quantity))<0) And ("
    "(Involine.TransType)=\"S\"))"
Begin InputTables
    Name ="Sysfile"
    Name ="Cashier"
    Name ="Involine"
    Name ="Item"
End
Begin OutputColumns
    Expression ="Involine.CashierID"
    Alias ="CashierName"
    Expression ="Cashier.Name"
    Expression ="Involine.TransType"
    Expression ="Involine.SetupDate"
    Expression ="Involine.CashSysID"
    Expression ="Involine.InvoHeadID"
    Expression ="Involine.Code"
    Expression ="Item.Description"
    Alias ="SumOfQuantity"
    Expression ="Sum(Involine.Quantity)"
    Alias ="SumOfPrice"
    Expression ="Sum(Involine.Price)"
End
Begin Joins
    LeftTable ="Involine"
    RightTable ="Item"
    Expression ="Involine.Code = Item.ID"
    Flag =1
    LeftTable ="Cashier"
    RightTable ="Involine"
    Expression ="Cashier.ID = Involine.CashierID"
    Flag =1
End
Begin OrderBy
    Expression ="Involine.CashierID"
    Flag =0
    Expression ="Involine.SetupDate"
    Flag =0
End
Begin Groups
    Expression ="Involine.CashierID"
    GroupLevel =0
    Expression ="Cashier.Name"
    GroupLevel =0
    Expression ="Involine.TransType"
    GroupLevel =0
    Expression ="Involine.SetupDate"
    GroupLevel =0
    Expression ="Involine.CashSysID"
    GroupLevel =0
    Expression ="Involine.InvoHeadID"
    GroupLevel =0
    Expression ="Involine.Code"
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
    Begin
        dbText "Name" ="InvoLine5.CashierID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Involine.CashierID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Involine.TransType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Involine.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Involine.CashSysID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Involine.InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Involine.Code"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =135
    Top =25
    Right =1360
    Bottom =702
    Left =-1
    Top =-1
    Right =1193
    Bottom =453
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =706
        Top =36
        Right =903
        Bottom =390
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
        Left =182
        Top =12
        Right =326
        Bottom =156
        Top =0
        Name ="Involine"
        Name =""
    End
    Begin
        Left =459
        Top =97
        Right =555
        Bottom =211
        Top =0
        Name ="Item"
        Name =""
    End
End
