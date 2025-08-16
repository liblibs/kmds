Operation =1
Option =0
Where ="(((SalesDepartment.ID) Between [Forms]![PrintInventoryReport]![cboSalesDepartmen"
    "tIDBegin] And [Forms]![PrintInventoryReport]![cboSalesDepartmentIDEnd]))"
Begin InputTables
    Name ="SalesDepartment"
    Name ="Item"
    Name ="qryItemsByDaysSupplySoldTotals"
End
Begin OutputColumns
    Expression ="Item.ID"
    Alias ="Sold"
    Expression ="IIf(IsNull([SumOfQuantity]),0,[SumOfQuantity])"
    Alias ="Days"
    Expression ="DateDiff(\"d\",Forms!PrintInventoryReport!txtDateRangeFrom,Forms!PrintInventoryR"
        "eport!txtDateRangeTo)+1"
    Alias ="AverageDailySales"
    Expression ="IIf([Days]=0,0,[Sold]/[Days])"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="qryItemsByDaysSupplySoldTotals"
    Expression ="Item.ID = qryItemsByDaysSupplySoldTotals.Code"
    Flag =2
    LeftTable ="SalesDepartment"
    RightTable ="Item"
    Expression ="SalesDepartment.ID = Item.SalesDepartmentID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Days"
        dbInteger "ColumnOrder" ="3"
        dbInteger "ColumnWidth" ="585"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sold"
        dbInteger "ColumnOrder" ="2"
        dbInteger "ColumnWidth" ="540"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbInteger "ColumnOrder" ="1"
        dbInteger "ColumnWidth" ="840"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AverageDailySales"
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
    Bottom =325
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =196
        Top =8
        Right =385
        Bottom =317
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =452
        Top =30
        Right =574
        Bottom =129
        Top =0
        Name ="qryItemsByDaysSupplySoldTotals"
        Name =""
    End
End
