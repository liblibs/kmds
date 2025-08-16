Operation =1
Option =0
Where ="(((InvoLine.InvoHeadID)=[parInvoiceNo]) AND ((InvoLine.CashSysID)=[parRegisterNu"
    "mber]))"
Begin InputTables
    Name ="Cashier"
    Name ="InvoLine"
    Name ="Tender"
    Name ="Item"
End
Begin OutputColumns
    Alias ="InvoiceNumber"
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.CashSysID"
    Alias ="CashierName"
    Expression ="Cashier.Name"
    Expression ="InvoLine.ID"
    Expression ="InvoLine.TransType"
    Expression ="InvoLine.Code"
    Alias ="TDescription"
    Expression ="InvoLine.InputCode"
    Alias ="IDescription"
    Expression ="Item.Description"
    Expression ="InvoLine.Quantity"
    Expression ="InvoLine.Price"
    Expression ="InvoLine.InputCode"
    Expression ="InvoLine.SetupDate"
End
Begin Parameters
    Name ="parInvoiceNo"
    Flag =4
    Name ="parRegisterNumber"
    Flag =4
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Tender"
    Expression ="InvoLine.Code = Tender.ID"
    Flag =2
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =2
    LeftTable ="Cashier"
    RightTable ="InvoLine"
    Expression ="Cashier.ID = InvoLine.CashierID"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLine.InvoHeadID"
    Flag =0
    Expression ="InvoLine.ID"
    Flag =0
    Expression ="InvoLine.TransType"
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
        dbText "Name" ="InvoiceNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashierName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IDescription"
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
    Bottom =201
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =17
        Top =31
        Right =113
        Bottom =115
        Top =0
        Name ="Cashier"
        Name =""
    End
    Begin
        Left =291
        Top =3
        Right =402
        Bottom =207
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =442
        Top =0
        Right =538
        Bottom =84
        Top =0
        Name ="Tender"
        Name =""
    End
    Begin
        Left =440
        Top =114
        Right =536
        Bottom =198
        Top =0
        Name ="Item"
        Name =""
    End
End
