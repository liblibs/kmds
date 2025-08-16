Operation =1
Option =0
Begin InputTables
    Name ="qryInvoLineReceiptsWithItemsSelected"
    Name ="InvoLine"
    Name ="Cashier"
    Name ="Item"
End
Begin OutputColumns
    Expression ="qryInvoLineReceiptsWithItemsSelected.DateOfSale"
    Expression ="qryInvoLineReceiptsWithItemsSelected.CashSysID"
    Expression ="qryInvoLineReceiptsWithItemsSelected.InvoHeadID"
    Expression ="InvoLine.ID"
    Alias ="Trans"
    Expression ="IIf([TransType]=\"T\",\"Tender\",\"Sale\")"
    Expression ="InvoLine.CashierID"
    Alias ="Cashier"
    Expression ="IIf(IsNull(Cashier.Name),\"Deleted\",Cashier.Name)"
    Expression ="InvoLine.Code"
    Expression ="Item.Description"
    Expression ="InvoLine.Quantity"
    Expression ="InvoLine.Price"
    Expression ="InvoLine.InputCode"
End
Begin Joins
    LeftTable ="qryInvoLineReceiptsWithItemsSelected"
    RightTable ="InvoLine"
    Expression ="qryInvoLineReceiptsWithItemsSelected.CashSysID = InvoLine.CashSysID"
    Flag =1
    LeftTable ="qryInvoLineReceiptsWithItemsSelected"
    RightTable ="InvoLine"
    Expression ="qryInvoLineReceiptsWithItemsSelected.InvoHeadID = InvoLine.InvoHeadID"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="Cashier"
    Expression ="InvoLine.CashierID = Cashier.ID"
    Flag =2
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =2
End
Begin OrderBy
    Expression ="qryInvoLineReceiptsWithItemsSelected.DateOfSale"
    Flag =0
    Expression ="InvoLine.ID"
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
        dbText "Name" ="Trans"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cashier"
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
    Bottom =292
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =12
        Top =12
        Right =278
        Bottom =96
        Top =0
        Name ="qryInvoLineReceiptsWithItemsSelected"
        Name =""
    End
    Begin
        Left =334
        Top =0
        Right =480
        Bottom =279
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =527
        Top =4
        Right =623
        Bottom =118
        Top =0
        Name ="Cashier"
        Name =""
    End
    Begin
        Left =531
        Top =124
        Right =627
        Bottom =238
        Top =0
        Name ="Item"
        Name =""
    End
End
