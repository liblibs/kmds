Operation =1
Option =0
Where ="(((Item.ID) Between Forms!PrintInventoryReport!cboItemIDBegin And Forms!PrintInv"
    "entoryReport!cboItemIDEnd) And ((InvoLine.SetupDate) Between Forms!PrintInventor"
    "yReport!txtDateRangeFrom And Forms!PrintInventoryReport!txtDateRangeTo))"
Begin InputTables
    Name ="InvoLine"
    Name ="Item"
    Name ="TranTypes"
    Name ="Cashier"
End
Begin OutputColumns
    Expression ="Item.ID"
    Expression ="Item.Description"
    Expression ="InvoLine.SetupDate"
    Alias ="Trans"
    Expression ="TranTypes.Description"
    Expression ="InvoLine.Price"
    Expression ="InvoLine.InputCode"
    Alias ="Purchases"
    Expression ="IIf(\"Stock Receipts\"=[Trans],[Quantity],0)"
    Alias ="Counter"
    Expression ="IIf(\" Sales - Counter\"=[Trans],[Quantity],0)"
    Alias ="Clubs"
    Expression ="IIf(\" Sales - Clubs\"=[Trans],[Quantity],0)"
    Expression ="InvoLine.Quantity"
    Alias ="Transfers"
    Expression ="IIf(\"Transfers\"=[Trans],[Quantity],0)"
    Alias ="Adjustments"
    Expression ="IIf(\"Variance Adjustment\"=[Trans],[Quantity],0)"
    Alias ="Emp"
    Expression ="Cashier.Name"
    Expression ="InvoLine.TransType"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="TranTypes"
    Expression ="InvoLine.TransType = TranTypes.Code"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="Cashier"
    Expression ="InvoLine.CashierID = Cashier.ID"
    Flag =2
End
Begin OrderBy
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
        dbText "Name" ="InvoLine.SetupDate"
        dbInteger "ColumnWidth" ="2190"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Trans"
        dbInteger "ColumnWidth" ="1515"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbInteger "ColumnWidth" ="2190"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Price"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.InputCode"
        dbInteger "ColumnWidth" ="1845"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Purchases"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Counter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clubs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transfers"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Adjustments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Emp"
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
    Bottom =186
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =191
        Top =6
        Right =287
        Bottom =195
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =38
        Top =8
        Right =134
        Bottom =167
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =381
        Top =121
        Right =477
        Bottom =220
        Top =0
        Name ="TranTypes"
        Name =""
    End
    Begin
        Left =515
        Top =6
        Right =611
        Bottom =120
        Top =0
        Name ="Cashier"
        Name =""
    End
End
