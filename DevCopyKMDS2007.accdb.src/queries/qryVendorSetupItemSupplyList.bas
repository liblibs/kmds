Operation =1
Option =0
Begin InputTables
    Name ="Vendor"
    Name ="qryVendorSetupItemSupplyListRAW"
End
Begin OutputColumns
    Expression ="qryVendorSetupItemSupplyListRAW.VendorID"
    Expression ="Vendor.Name"
    Expression ="Vendor.Contact"
    Expression ="Vendor.Address1"
    Expression ="Vendor.Address2"
    Expression ="Vendor.City"
    Expression ="Vendor.St"
    Expression ="Vendor.Zip"
    Expression ="Vendor.Phone1"
    Expression ="Vendor.Phone2"
    Expression ="Vendor.Phone3"
    Expression ="Vendor.OrderCount"
    Expression ="Vendor.PoAtCost"
    Expression ="Vendor.PrintUpcOnPo"
    Expression ="Vendor.POReportForm"
    Expression ="Vendor.VendorNotes"
    Expression ="qryVendorSetupItemSupplyListRAW.ID"
    Expression ="qryVendorSetupItemSupplyListRAW.StockNumber"
    Expression ="qryVendorSetupItemSupplyListRAW.Description"
    Expression ="qryVendorSetupItemSupplyListRAW.SizeD"
    Expression ="qryVendorSetupItemSupplyListRAW.Onhand"
    Expression ="qryVendorSetupItemSupplyListRAW.Cost"
    Expression ="qryVendorSetupItemSupplyListRAW.TodaysPrice"
    Expression ="qryVendorSetupItemSupplyListRAW.OnReserve"
    Expression ="qryVendorSetupItemSupplyListRAW.Avail"
    Expression ="qryVendorSetupItemSupplyListRAW.OrderAt"
    Expression ="qryVendorSetupItemSupplyListRAW.OrderTo"
    Alias ="SOQ"
    Expression ="IIf([Avail]<=[OrderAt],[OrderTo]-[Avail],0)"
End
Begin Joins
    LeftTable ="qryVendorSetupItemSupplyListRAW"
    RightTable ="Vendor"
    Expression ="qryVendorSetupItemSupplyListRAW.VendorID = Vendor.ID"
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
        dbText "Name" ="SOQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contact"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Address1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Address2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="St"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Zip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Phone1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Phone2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Phone3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PoAtCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PrintUpcOnPo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="POReportForm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorItemSupplyListRAW.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorSetupItemSupplyListRAW.OrderTo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorItemSupplyListRAW.OrderAt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorSetupItemSupplyListRAW.OrderAt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorItemSupplyListRAW.Avail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorSetupItemSupplyListRAW.Avail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorItemSupplyListRAW.OnReserve"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorSetupItemSupplyListRAW.OnReserve"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorItemSupplyListRAW.TodaysPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorSetupItemSupplyListRAW.TodaysPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorItemSupplyListRAW.Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorSetupItemSupplyListRAW.Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorItemSupplyListRAW.Onhand"
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
        dbText "Name" ="Vendor.City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.VendorNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorItemSupplyListRAW_1.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.St"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorItemSupplyListRAW.OrderTo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorSetupItemSupplyListRAW.Onhand"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorItemSupplyListRAW.SizeD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorSetupItemSupplyListRAW.SizeD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorItemSupplyListRAW.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorSetupItemSupplyListRAW.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorItemSupplyListRAW.StockNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorSetupItemSupplyListRAW.StockNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorItemSupplyListRAW.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorSetupItemSupplyListRAW.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.POReportForm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.Zip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.PoAtCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.PrintUpcOnPo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.Phone1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.Phone2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.Phone3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.OrderCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryVendorSetupItemSupplyListRAW.VendorID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =2
    Left =-8
    Top =-31
    Right =1410
    Bottom =704
    Left =-1
    Top =-1
    Right =1386
    Bottom =216
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =332
        Top =9
        Right =428
        Bottom =123
        Top =0
        Name ="Vendor"
        Name =""
    End
    Begin
        Left =668
        Top =12
        Right =812
        Bottom =156
        Top =0
        Name ="qryVendorSetupItemSupplyListRAW"
        Name =""
    End
End
