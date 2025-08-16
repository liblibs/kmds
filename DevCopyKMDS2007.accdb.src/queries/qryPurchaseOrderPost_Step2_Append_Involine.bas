Operation =3
Name ="InvoLine"
Option =0
Where ="(((PurchaseOrderLine.QuantityShipped)<>0) AND ((PurchaseOrderLine.PurchaseOrderI"
    "D)=[parPurchaseOrderID]))"
Begin InputTables
    Name ="PurchaseOrderLine"
    Name ="Item"
    Name ="Vendor"
End
Begin OutputColumns
    Alias ="InvoHeadID"
    Name ="InvoHeadID"
    Expression ="[parInvoiceNumber]"
    Alias ="CashSysID"
    Name ="CashSysID"
    Expression ="[parCashSysID]"
    Alias ="ID"
    Name ="ID"
    Expression ="PurchaseOrderLine.LineID"
    Alias ="CashierID"
    Name ="CashierID"
    Expression ="[parCashierID]"
    Alias ="ShiftID"
    Name ="ShiftID"
    Expression ="[parShiftID]"
    Alias ="Code"
    Name ="Code"
    Expression ="PurchaseOrderLine.ItemID"
    Alias ="SalesDept"
    Name ="SalesDept"
    Expression ="Item.SalesDepartmentID"
    Alias ="Quantity"
    Name ="Quantity"
    Expression ="PurchaseOrderLine.QuantityShipped"
    Alias ="XPrice"
    Name ="Price"
    Expression ="[QuantityShipped]*Item.Price"
    Alias ="InputCode"
    Name ="InputCode"
    Expression ="Left(CStr([PurchaseOrderID]) & \" \" & Vendor.Name,15)"
    Alias ="SalesGroup"
    Name ="SalesGroup"
    Expression ="Item.Onhand"
    Alias ="TransType"
    Name ="TransType"
    Expression ="\"P\""
    Alias ="SetupDate"
    Name ="SetupDate"
    Expression ="Now()"
End
Begin Parameters
    Name ="parInvoiceNumber"
    Flag =4
    Name ="parCashSysID"
    Flag =4
    Name ="parCashierID"
    Flag =4
    Name ="parShiftID"
    Flag =4
    Name ="parPurchaseOrderID"
    Flag =4
End
Begin Joins
    LeftTable ="PurchaseOrderLine"
    RightTable ="Item"
    Expression ="PurchaseOrderLine.ItemID = Item.ID"
    Flag =1
    LeftTable ="PurchaseOrderLine"
    RightTable ="Vendor"
    Expression ="PurchaseOrderLine.VendorID = Vendor.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashSysID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashierID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ShiftID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDept"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="XPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InputCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SetupDate"
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
    Bottom =329
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =241
        Top =14
        Right =425
        Bottom =278
        Top =0
        Name ="PurchaseOrderLine"
        Name =""
    End
    Begin
        Left =493
        Top =10
        Right =639
        Bottom =304
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =81
        Top =35
        Right =177
        Bottom =149
        Top =0
        Name ="Vendor"
        Name =""
    End
End
