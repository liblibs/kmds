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
    Expression ="Involine.CashierID"
    Alias ="CashierName"
    Expression ="Cashier.Name"
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
        dbText "Name" ="Involine.CashierID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Involine.Quantity"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =106
    Top =124
    Right =856
    Bottom =540
    Left =-1
    Top =-1
    Right =718
    Bottom =315
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
        Left =285
        Top =179
        Right =429
        Bottom =323
        Top =0
        Name ="Item"
        Name =""
    End
End
