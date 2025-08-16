Operation =1
Option =0
Where ="(((BuyDown.ID)=Forms!BuyDowns!ID) And ((BuyDownItem.QuantitySold)<>0))"
Begin InputTables
    Name ="BuyDown"
    Name ="BuyDownItem"
    Name ="Item"
End
Begin OutputColumns
    Expression ="BuyDown.ID"
    Expression ="BuyDown.RepsName"
    Expression ="BuyDown.ReferenceNo"
    Expression ="BuyDown.Club"
    Expression ="BuyDown.Order"
    Expression ="BuyDown.InvoiceNo"
    Expression ="BuyDown.BeginDate"
    Expression ="BuyDown.EndDate"
    Expression ="BuyDown.BuydownDollars"
    Expression ="BuyDown.DollarsToDate"
    Expression ="BuyDown.BuydownQty"
    Expression ="BuyDown.Quantity"
    Expression ="BuyDown.LastSoldDate"
    Expression ="BuyDown.Status"
    Expression ="BuyDown.PaidFlag"
    Expression ="BuyDown.ItemsID"
    Expression ="BuyDown.SetupDate"
    Expression ="BuyDown.ChangeDate"
    Expression ="BuyDownItem.ItemsCode"
    Expression ="Item.Description"
    Expression ="BuyDownItem.QuantitySold"
    Expression ="BuyDownItem.DiscountDollars"
End
Begin Joins
    LeftTable ="BuyDown"
    RightTable ="BuyDownItem"
    Expression ="BuyDown.ID = BuyDownItem.BuyDownID"
    Flag =1
    LeftTable ="BuyDownItem"
    RightTable ="Item"
    Expression ="BuyDownItem.ItemsCode = Item.ID"
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
        dbText "Name" ="BuyDown.ChangeDate"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyDown.Club"
        dbInteger "ColumnWidth" ="570"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyDown.Order"
        dbInteger "ColumnWidth" ="660"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyDown.InvoiceNo"
        dbInteger "ColumnWidth" ="1005"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyDown.ItemsID"
        dbInteger "ColumnWidth" ="840"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyDown.SetupDate"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =92
        Top =14
        Right =188
        Bottom =128
        Top =0
        Name ="BuyDown"
        Name =""
    End
    Begin
        Left =291
        Top =5
        Right =421
        Bottom =119
        Top =0
        Name ="BuyDownItem"
        Name =""
    End
    Begin
        Left =425
        Top =6
        Right =521
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
End
