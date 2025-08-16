Operation =1
Option =0
Begin InputTables
    Name ="InvoiceLineLD"
End
Begin OutputColumns
    Expression ="InvoiceLineLD.InvoHeadID"
    Expression ="InvoiceLineLD.SetupDate"
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
        dbText "Name" ="InvoiceLineLD.InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceLineLD.SetupDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =76
    Top =132
    Right =1570
    Bottom =877
    Left =-1
    Top =-1
    Right =1462
    Bottom =462
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =39
        Top =12
        Right =192
        Bottom =371
        Top =0
        Name ="InvoiceLineLD"
        Name =""
    End
End
