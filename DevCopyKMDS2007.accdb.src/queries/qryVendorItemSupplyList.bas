Operation =1
Option =0
Begin InputTables
    Name ="qryVendorItemSupplyListRAW"
    Name ="Vendor"
End
Begin OutputColumns
    Expression ="qryVendorItemSupplyListRAW.VendorID"
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
    Expression ="qryVendorItemSupplyListRAW.ID"
    Expression ="qryVendorItemSupplyListRAW.StockNumber"
    Expression ="qryVendorItemSupplyListRAW.Description"
    Expression ="qryVendorItemSupplyListRAW.SizeD"
    Expression ="qryVendorItemSupplyListRAW.Onhand"
    Expression ="qryVendorItemSupplyListRAW.Cost"
    Expression ="qryVendorItemSupplyListRAW.TodaysPrice"
    Expression ="qryVendorItemSupplyListRAW.OnReserve"
    Expression ="qryVendorItemSupplyListRAW.Avail"
    Expression ="qryVendorItemSupplyListRAW.OrderAt"
    Expression ="qryVendorItemSupplyListRAW.OrderTo"
    Alias ="SOQ"
    Expression ="IIf([Avail]<=[OrderAt],[OrderTo]-[Avail],0)"
End
Begin Joins
    LeftTable ="qryVendorItemSupplyListRAW"
    RightTable ="Vendor"
    Expression ="qryVendorItemSupplyListRAW.VendorID = Vendor.ID"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =77
        Top =3
        Right =173
        Bottom =117
        Top =0
        Name ="qryVendorItemSupplyListRAW"
        Name =""
    End
    Begin
        Left =332
        Top =9
        Right =428
        Bottom =123
        Top =0
        Name ="Vendor"
        Name =""
    End
End
