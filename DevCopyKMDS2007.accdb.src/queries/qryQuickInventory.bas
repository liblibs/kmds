Operation =1
Option =0
Where ="(((vwItemsWithCurrentPricing.SalesDepartmentID) Between Forms!PrintInventoryRepo"
    "rt!cboSalesDepartmentIDBegin And Forms!PrintInventoryReport!cboSalesDepartmentID"
    "End) And ((vwItemsWithCurrentPricing.Onhand)<>0) And ((IIf(Forms!PrintInventoryR"
    "eport!chkPrintZeroOnhands,-1,0))=0)) Or (((vwItemsWithCurrentPricing.SalesDepart"
    "mentID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin And Forms!P"
    "rintInventoryReport!cboSalesDepartmentIDEnd) And ((IIf(Forms!PrintInventoryRepor"
    "t!chkPrintZeroOnhands,-1,0))=-1))"
Begin InputTables
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Expression ="vwItemsWithCurrentPricing.SalesDepartmentID"
    Expression ="vwItemsWithCurrentPricing.ID"
    Expression ="vwItemsWithCurrentPricing.Onhand"
    Alias ="OnTheShelf"
    Expression ="IIf([OnShelf]<>[Onhand],[Onshelf],0)"
    Expression ="vwItemsWithCurrentPricing.Description"
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
        dbText "Name" ="vwItemsWithCurrentPricing.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OnTheShelf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Onhand"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.Description"
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
    Bottom =243
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
End
