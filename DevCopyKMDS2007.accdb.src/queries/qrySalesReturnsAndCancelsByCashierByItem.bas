Operation =1
Option =0
Where ="(((InvoLine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom And F"
    "orms!PrintInventoryReport!txtDateRangeTo) And ((InvoLine.Quantity)<0) And ((Invo"
    "Line.TransType)=\"S\"))"
Begin InputTables
    Name ="InvoLine"
    Name ="Cashier"
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Expression ="Cashier.Name"
    Expression ="vwItemsWithCurrentPricing.ID"
    Expression ="vwItemsWithCurrentPricing.Description"
    Expression ="vwItemsWithCurrentPricing.OnShelf"
    Expression ="InvoLine.SetupDate"
    Expression ="InvoLine.CashSysID"
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.InputCode"
    Expression ="InvoLine.Quantity"
    Expression ="InvoLine.Price"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Cashier"
    Expression ="InvoLine.CashierID = Cashier.ID"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="vwItemsWithCurrentPricing"
    Expression ="InvoLine.Code = vwItemsWithCurrentPricing.ID"
    Flag =1
End
Begin OrderBy
    Expression ="Cashier.Name"
    Flag =0
    Expression ="vwItemsWithCurrentPricing.ID"
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
        dbText "Name" ="vwItemsWithCurrentPricing.Description"
        dbInteger "ColumnWidth" ="3000"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.SetupDate"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =236
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =260
        Top =66
        Right =356
        Bottom =240
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="Cashier"
        Name =""
    End
    Begin
        Left =714
        Top =9
        Right =877
        Bottom =228
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
End
