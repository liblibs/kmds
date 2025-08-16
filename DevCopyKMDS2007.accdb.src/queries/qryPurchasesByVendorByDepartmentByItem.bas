Operation =1
Option =0
Where ="(((InvoLine.SalesGroup) Between Forms!PrintInventoryReport!cboVendorIDBegin And "
    "Forms!PrintInventoryReport!cboVendorIDEnd) And ((SalesDepartment.ID) Between For"
    "ms!PrintInventoryReport!cboSalesDepartmentIDBegin And Forms!PrintInventoryReport"
    "!cboSalesDepartmentIDEnd) And ((PriceGroup.Description) Between Forms!PrintInven"
    "toryReport!cboPriceGroupIDBegin And Forms!PrintInventoryReport!cboPriceGroupIDEn"
    "d) And ((InvoLine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom"
    " And Forms!PrintInventoryReport!txtDateRangeTo) And ((InvoLine.TransType)=\"P\")"
    ")"
Begin InputTables
    Name ="InvoLine"
    Name ="SalesGroup"
    Name ="SalesDepartment"
    Name ="Item"
    Name ="Vendor"
    Name ="PriceGroup"
End
Begin OutputColumns
    Expression ="InvoLine.SalesGroup"
    Alias ="VendorName"
    Expression ="Vendor.Name"
    Alias ="Department"
    Expression ="SalesDepartment.ID"
    Alias ="DepartmentDesc"
    Expression ="IIf(involine.Transtype='C',\"Club Sales\",SalesDepartment.Description)"
    Expression ="Item.PriceGroupID"
    Alias ="PriceGroupDescription"
    Expression ="PriceGroup.Description"
    Expression ="Item.Description"
    Expression ="Item.ID"
    Expression ="InvoLine.SetupDate"
    Alias ="TotalQ"
    Expression ="InvoLine.Quantity"
    Alias ="TotalP"
    Expression ="InvoLine.Price"
    Expression ="InvoLine.TransType"
    Expression ="InvoLine.InputCode"
    Alias ="PONumber"
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.Cost"
End
Begin Joins
    LeftTable ="SalesDepartment"
    RightTable ="Item"
    Expression ="SalesDepartment.ID = Item.SalesDepartmentID"
    Flag =1
    LeftTable ="SalesGroup"
    RightTable ="SalesDepartment"
    Expression ="SalesGroup.ID = SalesDepartment.SalesGroupID"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="Vendor"
    Expression ="Item.VendorID = Vendor.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="PriceGroup"
    Expression ="Item.PriceGroupID = PriceGroup.ID"
    Flag =1
End
Begin OrderBy
    Expression ="SalesDepartment.ID"
    Flag =0
    Expression ="Item.PriceGroupID"
    Flag =0
    Expression ="Item.Description"
    Flag =0
    Expression ="InvoLine.SetupDate"
    Flag =0
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
        dbText "Name" ="DepartmentDesc"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Department"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PriceGroupDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PONumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.SalesGroup"
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
    Bottom =273
    Left =0
    Top =64
    ColumnsShown =539
    Begin
        Left =65
        Top =-20
        Right =189
        Bottom =256
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =750
        Top =-20
        Right =846
        Bottom =94
        Top =0
        Name ="SalesGroup"
        Name =""
    End
    Begin
        Left =548
        Top =-13
        Right =644
        Bottom =101
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =407
        Top =23
        Right =503
        Bottom =137
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =237
        Top =103
        Right =333
        Bottom =217
        Top =0
        Name ="Vendor"
        Name =""
    End
    Begin
        Left =660
        Top =137
        Right =756
        Bottom =251
        Top =0
        Name ="PriceGroup"
        Name =""
    End
End
