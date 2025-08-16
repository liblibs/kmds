Operation =1
Option =0
Where ="(((InvoLine.InvoHeadID)=Forms!Vendors!VendorsPurchaseOrdersSub.Form!InvoHeadID) "
    "And ((InvoLine.CashSysID)=Forms!Vendors!VendorsPurchaseOrdersSub.Form!CashSysID)"
    " And ((InvoLine.SetupDate)=Forms!Vendors!VendorsPurchaseOrdersSub.Form!Date) And"
    " ((InvoLine.Quantity)<>0))"
Begin InputTables
    Name ="InvoLine"
    Name ="Item"
    Name ="Cashier"
    Name ="TranTypes"
    Name ="Vendor"
    Name ="Size"
    Name ="qryUPCItemMax"
End
Begin OutputColumns
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.CashSysID"
    Expression ="InvoLine.SetupDate"
    Alias ="POID"
    Expression ="Item.SalesDepartmentID"
    Alias ="VendorID"
    Expression ="InvoLine.SalesGroup"
    Alias ="VendorName"
    Expression ="Vendor.Name"
    Expression ="Vendor.Contact"
    Expression ="Vendor.Address1"
    Expression ="Vendor.Address2"
    Alias ="CitySTZ"
    Expression ="[City] & \", \" & [St] & \"  \" & [Zip]"
    Expression ="Vendor.Phone1"
    Expression ="Vendor.PoAtCost"
    Expression ="Vendor.PrintUpcOnPo"
    Alias ="LineID"
    Expression ="InvoLine.ID"
    Alias ="ItemID"
    Expression ="InvoLine.Code"
    Expression ="Item.Description"
    Alias ="QuantityShipped"
    Expression ="InvoLine.Quantity"
    Alias ="Cost"
    Expression ="InvoLine.Cost/InvoLine.quantity"
    Alias ="Price"
    Expression ="InvoLine.price/InvoLine.quantity"
    Alias ="SizeDesc"
    Expression ="Size.Description"
    Alias ="Cases"
    Expression ="[Quantity]\\[UnitsPerCase]"
    Expression ="Item.UnitsPerCase"
    Alias ="ExtCost"
    Expression ="InvoLine.Cost"
    Alias ="ExtPrice"
    Expression ="InvoLine.Price"
    Expression ="Item.StockNumber"
    Alias ="FAX"
    Expression ="Vendor.Phone2"
    Expression ="qryUPCItemMax.UPCCodeMax"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code=Item.ID"
    Flag =2
    LeftTable ="InvoLine"
    RightTable ="Cashier"
    Expression ="InvoLine.CashierID=Cashier.ID"
    Flag =2
    LeftTable ="InvoLine"
    RightTable ="TranTypes"
    Expression ="InvoLine.TransType=TranTypes.Code"
    Flag =2
    LeftTable ="InvoLine"
    RightTable ="Vendor"
    Expression ="InvoLine.SalesGroup=Vendor.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="Size"
    Expression ="Item.SizeID=Size.ID"
    Flag =2
    LeftTable ="Item"
    RightTable ="qryUPCItemMax"
    Expression ="Item.ID=qryUPCItemMax.ItemID"
    Flag =2
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
        dbText "Name" ="CitySTZ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cases"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="POID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LineID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuantityShipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FAX"
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
    Bottom =316
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =299
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =674
        Top =115
        Right =770
        Bottom =288
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =649
        Top =-11
        Right =745
        Bottom =103
        Top =0
        Name ="Cashier"
        Name =""
    End
    Begin
        Left =258
        Top =222
        Right =354
        Bottom =321
        Top =0
        Name ="TranTypes"
        Name =""
    End
    Begin
        Left =495
        Top =166
        Right =591
        Bottom =334
        Top =0
        Name ="Vendor"
        Name =""
    End
    Begin
        Left =845
        Top =99
        Right =941
        Bottom =213
        Top =0
        Name ="Size"
        Name =""
    End
    Begin
        Left =847
        Top =-7
        Right =943
        Bottom =92
        Top =0
        Name ="qryUPCItemMax"
        Name =""
    End
End
