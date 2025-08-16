Operation =1
Option =0
Where ="(((CStr([CODE]))=Involine5.InputCode) And ((Item.SalesDepartmentID)=1) And ((Inv"
    "oline5.TransType)=\"S\"))"
Begin InputTables
    Name ="Involine5"
    Name ="Cashier"
    Name ="Item"
End
Begin OutputColumns
    Alias ="Date-Time"
    Expression ="Involine5.SetupDate"
    Alias ="Invoice"
    Expression ="Involine5.InvoHeadID"
    Alias ="Reigister"
    Expression ="Involine5.CashSysID"
    Alias ="ItemID"
    Expression ="Involine5.Code"
    Alias ="UPCCode"
    Expression ="Involine5.InputCode"
    Expression ="Item.Description"
    Expression ="Involine5.CashierID"
    Alias ="CashierName"
    Expression ="Cashier.Name"
    Expression ="Involine5.Quantity"
    Alias ="Amount"
    Expression ="Involine5.Price"
End
Begin Joins
    LeftTable ="Involine5"
    RightTable ="Cashier"
    Expression ="Involine5.CashierID = Cashier.ID"
    Flag =2
    LeftTable ="Involine5"
    RightTable ="Item"
    Expression ="Involine5.Code = Item.ID"
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
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =2
    Left =-10
    Top =-33
    Right =1412
    Bottom =706
    Left =-1
    Top =-1
    Right =1390
    Bottom =287
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Involine5"
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
        Left =409
        Top =12
        Right =576
        Bottom =222
        Top =0
        Name ="Item"
        Name =""
    End
End
