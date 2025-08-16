Operation =1
Option =0
Where ="(((BuyDownItem.BuyDownID)=[parBuydownNumber]) And ((InvoLine.SetupDate) Between "
    "buydown.BeginDate And buydown.EndDate) And ((InvoLine.TransType)=\"S\" Or (InvoL"
    "ine.TransType)=\"C\"))"
Begin InputTables
    Name ="BuyDown"
    Name ="BuyDownItem"
    Name ="Item"
    Name ="InvoLine"
End
Begin OutputColumns
    Expression ="BuyDown.RepsName"
    Expression ="BuyDown.ReferenceNo"
    Expression ="Item.Description"
    Expression ="BuyDownItem.ItemsCode"
    Expression ="BuyDownItem.BuyDownID"
    Alias ="InvoiceDate"
    Expression ="InvoLine.SetupDate"
    Alias ="CashRegister"
    Expression ="InvoLine.CashSysID"
    Alias ="InvoiceNumber"
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.Quantity"
    Expression ="InvoLine.Price"
End
Begin Joins
    LeftTable ="BuyDownItem"
    RightTable ="Item"
    Expression ="BuyDownItem.ItemsCode = Item.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="InvoLine"
    Expression ="Item.ID = InvoLine.Code"
    Flag =1
    LeftTable ="BuyDown"
    RightTable ="BuyDownItem"
    Expression ="BuyDown.ID = BuyDownItem.BuyDownID"
    Flag =1
End
Begin OrderBy
    Expression ="Item.Description"
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
        dbText "Name" ="InvoiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashRegister"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceNumber"
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
        Left =30
        Top =-3
        Right =126
        Bottom =111
        Top =0
        Name ="BuyDown"
        Name =""
    End
    Begin
        Left =147
        Top =-5
        Right =243
        Bottom =109
        Top =0
        Name ="BuyDownItem"
        Name =""
    End
    Begin
        Left =342
        Top =5
        Right =438
        Bottom =119
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =537
        Top =-16
        Right =633
        Bottom =98
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
