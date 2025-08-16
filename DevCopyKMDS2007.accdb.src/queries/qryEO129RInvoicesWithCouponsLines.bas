Operation =1
Option =0
Begin InputTables
    Name ="Cashier"
    Name ="InvoLine5"
    Name ="qryEO129RInvoicesWithCoupons"
    Name ="Item"
End
Begin OutputColumns
    Expression ="qryEO129RInvoicesWithCoupons.CouponItemID"
    Expression ="qryEO129RInvoicesWithCoupons.CouponDescription"
    Expression ="InvoLine5.InvoHeadID"
    Expression ="InvoLine5.CashSysID"
    Expression ="InvoLine5.ID"
    Expression ="InvoLine5.SetupDate"
    Expression ="InvoLine5.CashierID"
    Expression ="Cashier.Name"
    Expression ="InvoLine5.Code"
    Expression ="Item.Description"
    Expression ="InvoLine5.Quantity"
    Expression ="InvoLine5.Price"
    Expression ="InvoLine5.InputCode"
    Expression ="InvoLine5.TransType"
    Alias ="CouponTotal"
    Expression ="IIf([code]=[CouponItemID],[InvoLine5].[Price],0)"
    Alias ="IsCoupon"
    Expression ="IIf([code]=[CouponItemID],\"*\",\"\")"
End
Begin Joins
    LeftTable ="qryEO129RInvoicesWithCoupons"
    RightTable ="InvoLine5"
    Expression ="qryEO129RInvoicesWithCoupons.InvoHeadID = InvoLine5.InvoHeadID"
    Flag =1
    LeftTable ="qryEO129RInvoicesWithCoupons"
    RightTable ="InvoLine5"
    Expression ="qryEO129RInvoicesWithCoupons.CashSysID = InvoLine5.CashSysID"
    Flag =1
    LeftTable ="qryEO129RInvoicesWithCoupons"
    RightTable ="InvoLine5"
    Expression ="qryEO129RInvoicesWithCoupons.SetupDate = InvoLine5.SetupDate"
    Flag =1
    LeftTable ="InvoLine5"
    RightTable ="Item"
    Expression ="InvoLine5.Code = Item.ID"
    Flag =2
    LeftTable ="InvoLine5"
    RightTable ="Cashier"
    Expression ="InvoLine5.CashierID = Cashier.ID"
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
        dbText "Name" ="InvoLine5.Quantity"
        dbLong "AggregateType" ="-1"
    End
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
        dbText "Name" ="InvoLine5.CashierID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cashier.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.Code"
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
        dbText "Name" ="InvoLine5.InputCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.TransType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryEO129RInvoicesWithCoupons.InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryEO129RInvoicesWithCoupons.CouponItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryEO129RInvoicesWithCoupons.CouponDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryEO129RInvoicesWithCoupons.SumOfPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CouponTotal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IsCoupon"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-32
    Top =54
    Right =1621
    Bottom =812
    Left =-1
    Top =-1
    Right =1304
    Bottom =372
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =505
        Top =26
        Right =649
        Bottom =170
        Top =0
        Name ="Cashier"
        Name =""
    End
    Begin
        Left =289
        Top =40
        Right =433
        Bottom =377
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =44
        Top =17
        Right =208
        Bottom =288
        Top =0
        Name ="qryEO129RInvoicesWithCoupons"
        Name =""
    End
    Begin
        Left =506
        Top =178
        Right =650
        Bottom =394
        Top =0
        Name ="Item"
        Name =""
    End
End
