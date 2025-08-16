Operation =1
Option =0
Where ="(((vwItemsWithCurrentPricing.SalesDepartmentID) Between Forms!PrintInventoryRepo"
    "rt!cboSalesDepartmentIDBegin And Forms!PrintInventoryReport!cboSalesDepartmentID"
    "End) And ((vwItemsWithCurrentPricing.CategoryDescr) Between Forms!PrintInventory"
    "Report!cboCategoryIDBegin And Forms!PrintInventoryReport!cboCategoryIDEnd) And ("
    "(vwItemsWithCurrentPricing.Onhand)<>0) And ((IIf(Forms!PrintInventoryReport!chkP"
    "rintZeroOnhands,-1,0))=0)) Or (((vwItemsWithCurrentPricing.SalesDepartmentID) Be"
    "tween Forms!PrintInventoryReport!cboSalesDepartmentIDBegin And Forms!PrintInvent"
    "oryReport!cboSalesDepartmentIDEnd) And ((vwItemsWithCurrentPricing.CategoryDescr"
    ") Between Forms!PrintInventoryReport!cboCategoryIDBegin And Forms!PrintInventory"
    "Report!cboCategoryIDEnd) And ((IIf(Forms!PrintInventoryReport!chkPrintZeroOnhand"
    "s,-1,0))=-1))"
Begin InputTables
    Name ="vwItemsWithCurrentPricing"
    Name ="qryInvoLineYTransMostRecent"
End
Begin OutputColumns
    Alias ="MostRecentCount"
    Expression ="IIf(IsNull([MaxOfSetupDate]),0,[MaxOfSetupDate])"
    Expression ="vwItemsWithCurrentPricing.SalesDepartmentID"
    Expression ="vwItemsWithCurrentPricing.SalesDptDescr"
    Expression ="vwItemsWithCurrentPricing.SubCategoryID"
    Expression ="vwItemsWithCurrentPricing.CategoryDescr"
    Expression ="vwItemsWithCurrentPricing.ID"
    Expression ="vwItemsWithCurrentPricing.Description"
    Expression ="vwItemsWithCurrentPricing.Onhand"
    Expression ="vwItemsWithCurrentPricing.Price"
    Expression ="vwItemsWithCurrentPricing.Cost"
End
Begin Joins
    LeftTable ="vwItemsWithCurrentPricing"
    RightTable ="qryInvoLineYTransMostRecent"
    Expression ="vwItemsWithCurrentPricing.ID = qryInvoLineYTransMostRecent.Code"
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
        dbText "Name" ="MostRecentCount"
        dbInteger "ColumnWidth" ="3330"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.SalesDptDescr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.SubCategoryID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.CategoryDescr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Onhand"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Cost"
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
    Right =1193
    Bottom =179
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =99
        Top =21
        Right =317
        Bottom =354
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
    Begin
        Left =365
        Top =12
        Right =509
        Bottom =156
        Top =0
        Name ="qryInvoLineYTransMostRecent"
        Name =""
    End
End
