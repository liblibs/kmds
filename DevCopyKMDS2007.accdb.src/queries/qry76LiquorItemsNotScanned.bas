Operation =1
Option =0
Where ="(((Involine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom And F"
    "orms!PrintInventoryReport!txtDateRangeTo) And ((CStr([CODE]))=Involine.InputCode"
    ") And ((Item.SalesDepartmentID)=1) And ((Involine.TransType)=\"S\"))"
Begin InputTables
    Name ="Involine"
    Name ="Cashier"
    Name ="Item"
End
Begin OutputColumns
    Expression ="Involine.CashierID"
    Alias ="CashierName"
    Expression ="Cashier.Name"
    Alias ="Date-Time"
    Expression ="Involine.SetupDate"
    Alias ="Invoice"
    Expression ="Involine.InvoHeadID"
    Alias ="Reigister"
    Expression ="Involine.CashSysID"
    Alias ="ItemID"
    Expression ="Involine.Code"
    Alias ="UPCCode"
    Expression ="Involine.InputCode"
    Expression ="Item.Description"
    Expression ="Involine.Quantity"
    Alias ="Amount"
    Expression ="Involine.Price"
End
Begin Joins
    LeftTable ="Involine"
    RightTable ="Cashier"
    Expression ="Involine.CashierID = Cashier.ID"
    Flag =2
    LeftTable ="Involine"
    RightTable ="Item"
    Expression ="Involine.Code = Item.ID"
    Flag =1
End
Begin OrderBy
    Expression ="Involine.CashierID"
    Flag =0
    Expression ="Involine.Code"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="InvoLine5.CashierID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.Quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date-Time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Invoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Reigister"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashierName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Amount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Involine.CashierID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Involine.Quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =78
    Top =136
    Right =1384
    Bottom =552
    Left =-1
    Top =-1
    Right =1274
    Bottom =236
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Involine"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Cashier"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =217
        Top =0
        Name ="Item"
        Name =""
    End
End
