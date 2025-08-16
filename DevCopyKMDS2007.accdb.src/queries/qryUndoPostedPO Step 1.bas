Operation =3
Name ="PurchaseOrderLine"
Option =0
Where ="(((InvoLine.TransType)=\"P\") AND ((InvoLine.InvoHeadID)=[InvoiceHeadID]))"
Begin InputTables
    Name ="InvoLine"
End
Begin OutputColumns
    Alias ="PO"
    Name ="PurchaseOrderID"
    Expression ="[NewPO]"
    Alias ="vid"
    Name ="VendorID"
    Expression ="[VendorNumber]"
    Name ="ItemID"
    Expression ="InvoLine.Code"
    Alias ="Q"
    Name ="QuantityOrdered"
    Expression ="0-[Quantity]"
    Alias ="q1"
    Name ="QuantityShipped"
    Expression ="0-[Quantity]"
    Name ="LineID"
    Expression ="InvoLine.ID"
End
Begin Parameters
    Name ="VendorNumber"
    Flag =4
    Name ="NewPO"
    Flag =4
    Name ="InvoiceHeadID"
    Flag =4
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="PO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="q1"
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
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =179
        Bottom =120
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
