Operation =1
Option =0
Where ="(((Item.Onhand)<>0) And ((CLng(IIf([AverageDailySales]>0,[Onhand]/[AverageDailyS"
    "ales],0))) Between CLng(Forms!PrintInventoryReport!txtNumericRangeFrom) And CLng"
    "(Forms!PrintInventoryReport!txtNumericRangeTo)) And ((IIf(Forms!PrintInventoryRe"
    "port!chkPrintZeroOnhands,-1,0))=0)) Or (((CLng(IIf([AverageDailySales]>0,[Onhand"
    "]/[AverageDailySales],0))) Between CLng(Forms!PrintInventoryReport!txtNumericRan"
    "geFrom) And CLng(Forms!PrintInventoryReport!txtNumericRangeTo)) And ((IIf(Forms!"
    "PrintInventoryReport!chkPrintZeroOnhands,-1,0))=-1))"
Begin InputTables
    Name ="SalesDepartment"
    Name ="Item"
    Name ="qryItemsByDaysSupplyAvgDailySales"
End
Begin OutputColumns
    Alias ="SaleDpt"
    Expression ="SalesDepartment.ID"
    Alias ="SaleDptDesc"
    Expression ="SalesDepartment.Description"
    Expression ="Item.ID"
    Expression ="Item.Description"
    Expression ="Item.Price"
    Expression ="Item.Onhand"
    Expression ="qryItemsByDaysSupplyAvgDailySales.AverageDailySales"
    Alias ="DaysSupply"
    Expression ="CLng(IIf([AverageDailySales]>0,[Onhand]/[AverageDailySales],0))"
    Alias ="RangeOfDaysFrom"
    Expression ="Forms!PrintInventoryReport!txtNumericRangeFrom"
    Alias ="RangeOfDaysTo"
    Expression ="Forms!PrintInventoryReport!txtNumericRangeTo"
    Expression ="Item.OrderAt"
    Expression ="Item.OrderTo"
    Expression ="qryItemsByDaysSupplyAvgDailySales.Sold"
    Expression ="qryItemsByDaysSupplyAvgDailySales.Days"
End
Begin Joins
    LeftTable ="SalesDepartment"
    RightTable ="Item"
    Expression ="SalesDepartment.ID = Item.SalesDepartmentID"
    Flag =1
    LeftTable ="Item"
    RightTable ="qryItemsByDaysSupplyAvgDailySales"
    Expression ="Item.ID = qryItemsByDaysSupplyAvgDailySales.ID"
    Flag =1
End
Begin OrderBy
    Expression ="SalesDepartment.ID"
    Flag =0
    Expression ="Item.ID"
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
        dbText "Name" ="Item.Description"
        dbInteger "ColumnWidth" ="2925"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbInteger "ColumnWidth" ="615"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Price"
        dbInteger "ColumnWidth" ="885"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemsByDaysSupplyAvgDailySales.AverageDailySales"
        dbInteger "ColumnWidth" ="1770"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DaysSupply"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RangeOfDaysFrom"
        dbInteger "ColumnWidth" ="1785"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SaleDpt"
        dbInteger "ColumnWidth" ="615"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SaleDptDesc"
        dbInteger "ColumnWidth" ="2730"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RangeOfDaysTo"
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
    Bottom =300
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =94
        Top =17
        Right =190
        Bottom =131
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =340
        Top =18
        Right =537
        Bottom =282
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =739
        Top =83
        Right =895
        Bottom =217
        Top =0
        Name ="qryItemsByDaysSupplyAvgDailySales"
        Name =""
    End
End
