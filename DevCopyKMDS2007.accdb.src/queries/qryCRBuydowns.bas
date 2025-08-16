Operation =1
Option =0
Where ="(((InvoLine.InvoHeadID)=[parInvoiceNo]) AND ((InvoLine.InputCode) Like \"Buydown"
    "%\"))"
Begin InputTables
    Name ="InvoLine"
    Name ="BuyDownItem"
End
Begin OutputColumns
    Alias ="BuyDownQty"
    Expression ="0-InvoLine.Quantity"
    Alias ="BuydownPrice"
    Expression ="0-[Price]"
    Alias ="BuyDownItemID"
    Expression ="InvoLine.SalesDept"
End
Begin Parameters
    Name ="parInvoiceNo"
    Flag =4
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="BuyDownItem"
    Expression ="InvoLine.SalesDept = BuyDownItem.ItemsCode"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLine.InvoHeadID"
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
        dbText "Name" ="BuyDownQty"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuydownPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyDownItemID"
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
    Bottom =177
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =148
        Bottom =195
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =240
        Top =9
        Right =375
        Bottom =183
        Top =0
        Name ="BuyDownItem"
        Name =""
    End
End
