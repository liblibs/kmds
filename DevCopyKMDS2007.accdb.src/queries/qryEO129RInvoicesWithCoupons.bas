Operation =1
Option =0
Begin InputTables
    Name ="Sysfile"
    Name ="InvoLine5"
    Name ="Item"
End
Begin OutputColumns
    Expression ="InvoLine5.InvoHeadID"
    Expression ="InvoLine5.CashSysID"
    Expression ="InvoLine5.SetupDate"
    Expression ="Sysfile.CouponItemID"
    Alias ="CouponDescription"
    Expression ="Item.Description"
    Alias ="SumOfPrice"
    Expression ="Sum(InvoLine5.Price)"
End
Begin Joins
    LeftTable ="Sysfile"
    RightTable ="InvoLine5"
    Expression ="Sysfile.CouponItemID = InvoLine5.Code"
    Flag =1
    LeftTable ="InvoLine5"
    RightTable ="Item"
    Expression ="InvoLine5.Code = Item.ID"
    Flag =1
End
Begin Groups
    Expression ="InvoLine5.InvoHeadID"
    GroupLevel =0
    Expression ="InvoLine5.CashSysID"
    GroupLevel =0
    Expression ="InvoLine5.SetupDate"
    GroupLevel =0
    Expression ="Sysfile.CouponItemID"
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
        dbText "Name" ="InvoLine5.InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.CashSysID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.CouponItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CouponDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPrice"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =5
    Top =72
    Right =1299
    Bottom =830
    Left =-1
    Top =-1
    Right =1262
    Bottom =406
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =205
        Bottom =441
        Top =0
        Name ="Sysfile"
        Name =""
    End
    Begin
        Left =474
        Top =1
        Right =815
        Bottom =409
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =1068
        Top =66
        Right =1204
        Bottom =287
        Top =0
        Name ="Item"
        Name =""
    End
End
