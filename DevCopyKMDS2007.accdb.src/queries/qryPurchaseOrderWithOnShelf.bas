Operation =1
Option =0
Where ="(((PurchaseOrder.ID)=Forms!PurchaseOrderChoices!txtPurchaseOrderID) And ((Purcha"
    "seOrderLine.QuantityShipped)<>0))"
Begin InputTables
    Name ="Vendor"
    Name ="PurchaseOrder"
    Name ="PurchaseOrderLine"
    Name ="vwItemsWithCurrentPricing"
    Name ="qryUPCItemMax"
    Name ="Size"
End
Begin OutputColumns
    Expression ="PurchaseOrder.ID"
    Expression ="PurchaseOrder.VendorID"
    Expression ="PurchaseOrder.OrderCount"
    Expression ="PurchaseOrder.OrderDate"
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
    Expression ="PurchaseOrderLine.LineID"
    Expression ="PurchaseOrderLine.ItemID"
    Expression ="vwItemsWithCurrentPricing.Description"
    Expression ="PurchaseOrderLine.QuantityShipped"
    Expression ="PurchaseOrderLine.Cost"
    Expression ="PurchaseOrderLine.Price"
    Expression ="PurchaseOrderLine.SetupDate"
    Expression ="PurchaseOrderLine.Repack"
    Alias ="SizeDesc"
    Expression ="Size.Description"
    Alias ="Cases"
    Expression ="IIf([UnitsPerCase]>1,[QuantityShipped]\\[UnitsPerCase],0)"
    Alias ="Singles"
    Expression ="IIf([UnitsPerCase]>1,[QuantityShipped]-(([QuantityShipped]\\[UnitsPerCase])*[Uni"
        "tsPerCase]),[QuantityShipped])"
    Expression ="vwItemsWithCurrentPricing.UnitsPerCase"
    Alias ="ExtCost"
    Expression ="[QuantityShipped]*PurchaseOrderLine.Cost"
    Alias ="ExtPrice"
    Expression ="[QuantityShipped]*PurchaseOrderLine.Price"
    Expression ="vwItemsWithCurrentPricing.StockNumber"
    Expression ="qryUPCItemMax.UPCCodeMax"
    Alias ="FAX"
    Expression ="Vendor.Phone2"
    Expression ="Vendor.PoAtCost"
    Expression ="vwItemsWithCurrentPricing.OnShelf"
    Expression ="vwItemsWithCurrentPricing.DaysSupply"
    Alias ="DaysLeft"
    Expression ="IIf([AverageDailySales]>0 And [Onhand]>0,[Onhand]/[AverageDailySales],0)"
    Expression ="vwItemsWithCurrentPricing.AverageDailySales"
End
Begin Joins
    LeftTable ="PurchaseOrder"
    RightTable ="PurchaseOrderLine"
    Expression ="PurchaseOrder.ID=PurchaseOrderLine.PurchaseOrderID"
    Flag =1
    LeftTable ="Vendor"
    RightTable ="PurchaseOrderLine"
    Expression ="Vendor.ID=PurchaseOrderLine.VendorID"
    Flag =1
    LeftTable ="PurchaseOrderLine"
    RightTable ="vwItemsWithCurrentPricing"
    Expression ="PurchaseOrderLine.ItemID=vwItemsWithCurrentPricing.ID"
    Flag =1
    LeftTable ="vwItemsWithCurrentPricing"
    RightTable ="qryUPCItemMax"
    Expression ="vwItemsWithCurrentPricing.ID=qryUPCItemMax.ItemID"
    Flag =2
    LeftTable ="vwItemsWithCurrentPricing"
    RightTable ="Size"
    Expression ="vwItemsWithCurrentPricing.SizeID=Size.ID"
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
        dbText "Name" ="PurchaseOrder.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.OnShelf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrder.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrder.OrderCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrder.OrderDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.Contact"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.Address1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.Address2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.Phone1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1010"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.PrintUpcOnPo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.LineID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.QuantityShipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.Repack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryUPCItemMax.UPCCodeMax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.PoAtCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Singles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FAX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cases"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.StockNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.DaysSupply"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.AverageDailySales"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3180"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="DaysLeft"
        dbInteger "ColumnWidth" ="3180"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =2
    Left =-8
    Top =-30
    Right =1265
    Bottom =885
    Left =-1
    Top =-1
    Right =1241
    Bottom =433
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =331
        Top =0
        Right =455
        Bottom =114
        Top =0
        Name ="Vendor"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =161
        Bottom =150
        Top =0
        Name ="PurchaseOrder"
        Name =""
    End
    Begin
        Left =194
        Top =1
        Right =290
        Bottom =235
        Top =0
        Name ="PurchaseOrderLine"
        Name =""
    End
    Begin
        Left =494
        Top =95
        Right =688
        Bottom =423
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
    Begin
        Left =747
        Top =28
        Right =946
        Bottom =149
        Top =0
        Name ="qryUPCItemMax"
        Name =""
    End
    Begin
        Left =834
        Top =158
        Right =1000
        Bottom =310
        Top =0
        Name ="Size"
        Name =""
    End
End
