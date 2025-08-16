Operation =1
Option =0
Where ="(((InvoLine.TransType)='T'))"
Having ="(((Cashier.ID)=[parCashierID]))"
Begin InputTables
    Name ="Cashier"
    Name ="InvoLine"
    Name ="Tender"
End
Begin OutputColumns
    Expression ="InvoLine.Code"
    Expression ="Tender.Description"
    Alias ="SumOfQuantity"
    Expression ="Sum(InvoLine.Quantity)"
    Alias ="SumOfPrice"
    Expression ="Sum(InvoLine.Price)"
    Expression ="Cashier.ID"
    Expression ="Cashier.Name"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Tender"
    Expression ="InvoLine.Code = Tender.ID"
    Flag =1
    LeftTable ="Cashier"
    RightTable ="InvoLine"
    Expression ="Cashier.ID = InvoLine.CashierID"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLine.Code"
    Flag =0
    Expression ="Cashier.ID"
    Flag =0
End
Begin Groups
    Expression ="InvoLine.Code"
    GroupLevel =0
    Expression ="Tender.Description"
    GroupLevel =0
    Expression ="Cashier.ID"
    GroupLevel =0
    Expression ="Cashier.Name"
    GroupLevel =0
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
        dbText "Name" ="SumOfQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPrice"
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
    Bottom =194
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =17
        Top =128
        Right =113
        Bottom =212
        Top =0
        Name ="Cashier"
        Name =""
    End
    Begin
        Left =225
        Top =13
        Right =321
        Bottom =142
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =406
        Top =-8
        Right =502
        Bottom =76
        Top =0
        Name ="Tender"
        Name =""
    End
End
